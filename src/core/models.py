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
    
    # 1月考勤表特有字段：上一年度剩余未休
    # 仅用于计算1月份的结转
    last_year_balance: float = Field(default=0.0, description="截止到上月底剩余未休 (从1月考勤表获取)")

    @field_validator('last_year_balance', mode='before')
    @classmethod
    def parse_balance(cls, v):
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

class MonthlyAttendance(BaseModel):
    """
    月度考勤记录模型。
    """
    month: int = Field(ge=1, le=12, description="月份 (1-12)")
    
    # 原始数据
    actual_leave_days: float = Field(ge=0.0, default=0.0, description="当月休假天数 (考勤表读取)")
    holiday_statuses: Dict[date, str] = Field(default_factory=dict, description="特定节假日的出勤状态")
    
    # 计算字段 (递归计算生成)
    opening_balance: float = Field(default=0.0, description="本月结转")
    
    @computed_field
    def holiday_leave_days(self) -> float:
        """
        计算当月节假日休假折算天数。
        规则：
        - "休" -> 1.0
        - "休0.5" / "班0.5" -> 0.5 (都有半天休息)
        - 其他 -> 0.0
        """
        count = 0.0
        month_holidays = GlobalConfig.get_holidays().get(self.month, [])
        for h_date in month_holidays:
            status = self.holiday_statuses.get(h_date, '')
            if not status: continue
            
            s = str(status).strip()
            if s == '休':
                count += 1.0
            elif s in ['休0.5', '班0.5']:
                count += 0.5
        return count

    @computed_field
    def holiday_count(self) -> int:
        """当月法定节假日总天数 (配置值)"""
        return len(GlobalConfig.get_holidays().get(self.month, []))

    @computed_field
    def bonus_days(self) -> float:
        """
        计算特殊补假天数。
        业务规则：2月份春节期间全勤上班奖励2天。
        全勤判定：必须所有法定假日状态均为'班'。
        """
        if self.month == 2:
            month_holidays = GlobalConfig.get_holidays().get(2, [])
            if not month_holidays:
                return 0.0

            worked_count = 0
            for d in month_holidays:
                status = self.holiday_statuses.get(d, '班')
                s = str(status).strip()
                # 只有严格的'班'才算全勤，'班0.5'或'休0.5'都不算
                if s == '班':
                    worked_count += 1
            
            # 全勤奖励
            if len(month_holidays) > 0 and worked_count == len(month_holidays):
                 return 2.0
        return 0.0

class EmployeeAnnualReport(BaseModel):
    employee: EmployeeBase = Field(description="员工基础信息")
    monthly_records: Dict[int, MonthlyAttendance] = Field(default_factory=dict)
    
    def calculate_monthly_balances(self):
        """
        执行月度结转递归计算。
        规则：
        1月结转 = 上年剩余 + 4
        2-12月结转 = 上月结转 - 上月休假 + 4 + 上月节假日数 - 上月节假日休假数 + 上月补假
        """
        # 确保按月份顺序处理
        sorted_months = sorted(self.monthly_records.keys())
        if not sorted_months: return

        # 1. 计算 1 月
        if 1 in self.monthly_records:
            rec_jan = self.monthly_records[1]
            # 1月结转 = 上年剩余 + 4
            rec_jan.opening_balance = self.employee.last_year_balance + 4.0
        else:
            # 如果没有1月记录（中途入职？），需特殊处理
            # 假设默认从存在的第一个月开始？或者补全1-12月？
            # 业务通常要求补全空月。在 reader 中应该已经补全了 MonthlyAttendance 对象。
            pass

        # 2. 递归计算后续月份
        for m in range(2, 13):
            if m not in self.monthly_records: continue
            
            prev_record = self.monthly_records.get(m - 1)
            curr_record = self.monthly_records[m]
            
            if prev_record:
                # 本月结转 = 上月结转 - 上月休假 + 4 + 上月节假日数 - 上月节假日休假数 + 上月补假
                # 注意：上月休假 = 考勤表休假天数 (actual_leave_days)
                
                balance = (prev_record.opening_balance 
                         - prev_record.actual_leave_days 
                         + 4.0 
                         + prev_record.holiday_count 
                         - prev_record.holiday_leave_days
                         + prev_record.bonus_days)
                
                curr_record.opening_balance = round(balance, 1) # 保留1位小数

    @computed_field
    def total_leave_taken(self) -> float:
        """
        合计已休 = 历月累计“休假天数” + 历月累计“节假日休假数”
        """
        total = 0.0
        for r in self.monthly_records.values():
            total += r.actual_leave_days + r.holiday_leave_days
        return round(total, 1)

    @computed_field
    def total_holiday_leave(self) -> float:
        """
        节假日未加班天数（即节假日休假数累计）
        """
        return round(sum(r.holiday_leave_days for r in self.monthly_records.values()), 1)

    @computed_field
    def total_bonus(self) -> float:
        """
        年度补假总天数
        """
        return round(sum(r.bonus_days for r in self.monthly_records.values()), 1)

    @computed_field
    def remaining_balance(self) -> float:
        """
        剩余未休 = 12月结转 - 12月休假 + 12月节假日数 - 12月节假日休假数 + 12月补假 (不再+4)
        """
        rec_dec = self.monthly_records.get(12)
        if not rec_dec: return 0.0
        
        val = (rec_dec.opening_balance 
             - rec_dec.actual_leave_days 
             + rec_dec.holiday_count 
             - rec_dec.holiday_leave_days
             + rec_dec.bonus_days)
        return round(val, 1)

    @computed_field
    def notes(self) -> str:
        """备注：透支提示等"""
        if self.remaining_balance < 0:
            return f"透支{abs(self.remaining_balance)}天"
        return ""
