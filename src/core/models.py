from pydantic import BaseModel, Field, field_validator, computed_field
from datetime import date, datetime
from typing import Dict, ClassVar, Optional, Any
import pandas as pd
from src.core.config import GlobalConfig, DEFAULT_MONTHLY_FLEX_QUOTA, DEFAULT_ANNUAL_FLEX_QUOTA

class EmployeeBase(BaseModel):
    """
    员工基础信息模型。
    """
    index: int = Field(default=0, description="序号")
    name: str = Field(description="员工姓名")
    department: Optional[str] = Field(default=None, description="所属部门")
    join_date: Optional[date] = Field(default=None, description="入职日期")
    opening_balance: float = Field(default=0.0, description="期初年休假结余天数 (从一月份考勤表获取)")

    @field_validator('opening_balance', mode='before')
    @classmethod
    def parse_opening_balance(cls, v):
        """处理期初结余，支持负数"""
        if pd.isna(v) or v == "":
            return 0.0
        try:
            return float(v)
        except ValueError:
            return 0.0

    @field_validator('join_date', mode='before')
    @classmethod
    def parse_date(cls, v):
        if pd.isna(v) or v == "":
            return None
        if isinstance(v, (pd.Timestamp, datetime)):
            return v.date()
        if isinstance(v, str):
            try:
                # 尝试多种日期格式
                for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"]:
                    try:
                        return datetime.strptime(v, fmt).date()
                    except ValueError:
                        continue
            except ValueError:
                pass
        return None

    @computed_field
    def flexible_quota(self) -> float:
        """
        计算年度弹性假总额度。
        规则：
        1. 老员工（join_date < target_year）：固定 48 天。
        2. 新员工（join_date >= target_year）：(12 - 入职月份) * 4。
        """
        target_year = GlobalConfig.get_year()
        threshold_date = date(target_year, 1, 1)
        
        # 如果无入职日期，默认作为老员工处理 (兜底逻辑)
        if self.join_date is None:
            return DEFAULT_ANNUAL_FLEX_QUOTA
            
        if self.join_date < threshold_date:
            return DEFAULT_ANNUAL_FLEX_QUOTA
        
        # 新员工计算逻辑
        months_remaining = 12 - self.join_date.month
        return max(0.0, float(months_remaining * DEFAULT_MONTHLY_FLEX_QUOTA))

    @computed_field
    def calculated_opening_balance(self) -> float:
        """
        根据入职日期计算期初余额。
        业务规则：
        - 老员工：48天
        - 新员工：(12 - 入职月份) * 4
        
        修正：如果是1月份入职的新员工，或者是历史数据延续下来的员工，
        如果考勤表中有明确的"截止到上月底剩余未休"（即 opening_balance），
        且该值可能是负数（透支），我们应该如何处理？
        
        根据用户最新需求：
        "1月份数据处理规则：将'截止到上月底剩余未休'字段值作为上一年度的剩余年休假增量"
        这意味着，如果 opening_balance 有值（非0），它应该被纳入计算。
        
        如果员工是老员工（48天额度），加上这个增量（可能是负数）。
        如果员工是新员工，通常没有上一年度剩余，直接按规则计算额度。
        
        但之前的逻辑是：calculated_opening_balance 完全由入职日期决定（48 或 (12-m)*4）。
        这里我们需要区分“年度额度”和“期初余额”。
        
        年度额度 = 48 (老) 或 (12-m)*4 (新)
        期初余额 = 年度额度 + 上年结转(即 opening_balance)
        
        因此，我们修改这个字段的含义为“年度总可用期初”，或者拆分字段。
        为了保持兼容性，我们让 calculated_opening_balance 代表“年度基础额度 + 结转余额”。
        """
        base_quota = self.flexible_quota
        return base_quota + self.opening_balance

class MonthlyAttendance(BaseModel):
    """
    月度考勤记录模型。
    """
    month: int = Field(ge=1, le=12, description="月份 (1-12)")
    
    # 提取的原始字段
    balance_end_of_last_month: Optional[float] = Field(default=None, description="截止到上月底剩余未休")
    start_time: Optional[Any] = Field(default=None, description="出场时间")
    end_time: Optional[Any] = Field(default=None, description="返场时间")
    actual_leave_days: float = Field(ge=0.0, default=0.0, description="当月实际休假总天数")
    
    # 动态字段
    holiday_statuses: Dict[date, str] = Field(default_factory=dict, description="特定节假日的出勤状态")
    
    # 快照字段（仅供参考/验证）
    snapshot_annual_balance_inc_holidays: Optional[float] = Field(default=None, description="年剩余假期（含法定节假日）")
    snapshot_annual_balance_exc_future_holidays: Optional[float] = Field(default=None, description="年剩余假期（不含未来法定节假日）")

    @computed_field
    def public_holidays_taken(self) -> float:
        """计算当月已休的法定节假日天数。"""
        count = 0
        month_holidays = GlobalConfig.get_holidays().get(self.month, [])
        for h_date in month_holidays:
            status = self.holiday_statuses.get(h_date, '班')
            if status == '休':
                count += 1
        return float(count)

    @computed_field
    def deduction(self) -> float:
        """
        计算当月扣除的年假天数。
        规则：直接返回考勤表中的'休假天数'字段，不做任何扣减。
        """
        return self.actual_leave_days

    @computed_field
    def bonus_days(self) -> float:
        """
        计算特殊补假天数。
        业务规则：2月份春节期间全勤上班奖励2天。
        """
        if self.month == 2:
            month_holidays = GlobalConfig.get_holidays().get(2, [])
            if not month_holidays:
                return 0.0

            worked_count = 0
            for d in month_holidays:
                status = self.holiday_statuses.get(d, '班')
                if status != '休':
                    worked_count += 1
            
            # 全勤奖励
            if len(month_holidays) > 0 and worked_count == len(month_holidays):
                 return 2.0
        return 0.0

class EmployeeAnnualReport(BaseModel):
    employee: EmployeeBase = Field(description="员工基础信息")
    monthly_records: Dict[int, MonthlyAttendance] = Field(default_factory=dict)
    
    @computed_field
    def total_deduction(self) -> float:
        return sum(record.deduction for record in self.monthly_records.values())

    @computed_field
    def total_bonus(self) -> float:
        return sum(record.bonus_days for record in self.monthly_records.values())

    @computed_field
    def remaining_balance(self) -> float:
        """
        年度剩余年假余额。
        公式：(基础额度 + 结转增量) + 年度补假 - 年度累计扣除。
        """
        # 注意：calculated_opening_balance 现在包含了 (base_quota + opening_balance)
        base_balance = self.employee.calculated_opening_balance
        return base_balance + self.total_bonus - self.total_deduction

    @computed_field
    def notes(self) -> str:
        notes_list = []
        if self.total_bonus > 0:
            notes_list.append(f"获得补假{self.total_bonus}天")
        if self.remaining_balance < 0:
            notes_list.append(f"余额透支{abs(self.remaining_balance)}天")
        return "; ".join(notes_list)
