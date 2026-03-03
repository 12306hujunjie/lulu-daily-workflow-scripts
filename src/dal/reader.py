from pydantic import BaseModel, Field
from typing import Dict, List, Optional
import pandas as pd
import re
from src.core.models import EmployeeBase, MonthlyAttendance, EmployeeAnnualReport
from src.core.config import GlobalConfig

from datetime import date

class ExcelReader(BaseModel):
    """
    Excel 数据读取器。
    负责从文件加载原始数据，并将其转换为 Core 层定义的 Pydantic 模型。
    """
    attendance_file_path: str = Field(description="月度考勤表路径")
    stats_file_path: Optional[str] = Field(default=None, description="历史年份统计表路径 (可选，用于获取入职日期)")
    
    # 内部缓存
    raw_attendance_dfs: Dict[str, pd.DataFrame] = Field(default_factory=dict)
    historical_employee_info: Dict[str, dict] = Field(default_factory=dict) # Name -> {department, join_date}
    # 记录每个月实际出现的节假日日期
    active_monthly_holidays: Dict[int, List[date]] = Field(default_factory=dict) 

    # 转换后的模型数据
    employees: Dict[str, EmployeeBase] = Field(default_factory=dict) # Name -> EmployeeBase
    attendance_records: Dict[str, Dict[int, MonthlyAttendance]] = Field(default_factory=dict) # Name -> Month -> Record

    class Config:
        arbitrary_types_allowed = True

    def load_files(self):
        """加载考勤表 Excel 文件并识别年份。同时尝试加载历史统计表。"""
        # 1. 加载考勤表
        try:
            print(f"正在加载考勤表: {self.attendance_file_path}")
            xl = pd.ExcelFile(self.attendance_file_path)
            
            # 尝试从第一个 Sheet 的标题中提取年份
            if xl.sheet_names:
                first_sheet = xl.sheet_names[0]
                df_peek = pd.read_excel(self.attendance_file_path, sheet_name=first_sheet, nrows=5, header=None)
                year = self._extract_year_from_sheet(df_peek)
                if year:
                    print(f"检测到考勤年份: {year}")
                    GlobalConfig.set_year(year)
                else:
                    print("警告: 无法从考勤表中识别年份，将使用默认配置。")

            # 加载所有月份数据
            for m in range(1, 13):
                sheet_name = f"{m}月"
                if sheet_name in xl.sheet_names:
                    df = pd.read_excel(self.attendance_file_path, sheet_name=sheet_name)
                    self.raw_attendance_dfs[sheet_name] = df
                    
        except Exception as e:
            raise ValueError(f"加载考勤表失败: {e}")

        # 2. 加载历史统计表 (如果提供了路径)
        if self.stats_file_path:
            try:
                print(f"正在加载历史统计表: {self.stats_file_path}")
                # 假设历史表结构与之前类似，Sheet名为 "员工年休假统计表 (带节假日)" 或类似的
                # 我们先读取所有Sheet，找包含 "统计表" 的
                xl_stats = pd.ExcelFile(self.stats_file_path)
                target_sheet = None
                for sheet in xl_stats.sheet_names:
                    if "统计表" in sheet:
                        target_sheet = sheet
                        break
                
                if target_sheet:
                    df_stats = pd.read_excel(self.stats_file_path, sheet_name=target_sheet, header=None)
                    self._parse_historical_info(df_stats)
                else:
                    print("警告: 历史统计表中未找到匹配的 Sheet。")
            except Exception as e:
                print(f"警告: 加载历史统计表失败: {e}")

    def _extract_year_from_sheet(self, df: pd.DataFrame) -> Optional[int]:
        """从 Sheet 的前几行尝试提取年份 (如 '2026年1月份考勤表')。"""
        for i in range(min(5, len(df))):
            for col in df.columns:
                val = str(df.iloc[i, col])
                match = re.search(r"(\d{4})年", val)
                if match:
                    return int(match.group(1))
        return None

    def _parse_historical_info(self, df: pd.DataFrame):
        """解析历史统计表，提取部门和入职日期。"""
        # 寻找表头行 (包含 "姓名", "部门", "入职日期")
        header_idx = -1
        for i in range(min(10, len(df))):
            row_vals = [str(v).strip() for v in df.iloc[i].values if pd.notna(v)]
            if "姓名" in row_vals and "入职日期" in row_vals:
                header_idx = i
                break
        
        if header_idx == -1:
            print("警告: 无法在历史统计表中定位表头。")
            return

        # 设置列名
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx+1:]
        
        count = 0
        for _, row in df.iterrows():
            name = row.get("姓名")
            if pd.isna(name): continue
            name = str(name).strip()
            
            dept = row.get("部门")
            join_date = row.get("入职日期")
            
            info = {}
            if pd.notna(dept): info["department"] = str(dept).strip()
            if pd.notna(join_date): info["join_date"] = join_date # 保持原始格式，让模型去解析
            
            self.historical_employee_info[name] = info
            count += 1
            
        print(f"从历史表中提取了 {count} 名员工的信息。")

    def parse_data(self):
        """
        遍历所有月份的考勤表，提取员工信息和考勤记录。
        """
        holidays_config = GlobalConfig.get_holidays()
        target_year = GlobalConfig.get_year()
        index_counter = 1
        
        # 预先填充所有月份的 active_monthly_holidays，即使考勤表不存在
        # 这样 ExcelReportGenerator 就能渲染所有月份的列头
        for m in range(1, 13):
            self.active_monthly_holidays[m] = holidays_config.get(m, [])

        if not self.raw_attendance_dfs:
            print("警告: 未加载任何考勤数据，但将生成空报表框架。")
            return

        sorted_months = sorted([int(k.replace("月", "")) for k in self.raw_attendance_dfs.keys()])
        
        for month in sorted_months:
            sheet_name = f"{month}月"
            df = self.raw_attendance_dfs[sheet_name]
            
            # 定位表头
            header_idx = -1
            for idx, row in df.iterrows():
                if '姓名' in row.values:
                    header_idx = idx
                    break
            
            if header_idx == -1:
                print(f"警告: {sheet_name} 未找到表头，跳过。")
                continue

            df.columns = df.iloc[header_idx]
            df = df.iloc[header_idx+1:]

            month_holidays = holidays_config.get(month, [])
            
            # 修正：直接使用所有法定节假日，不再通过考勤表列过滤
            # 这样可以保证报表中显示所有法定假日，即使考勤表中没有这些列（空缺月份）
            # 对于考勤表中存在的列，后续会正常提取状态；不存在的列，状态默认为空或班
            
            present_holidays = month_holidays # 直接使用全量假日
            self.active_monthly_holidays[month] = present_holidays

            for _, row in df.iterrows():
                name = row.get('姓名')
                if pd.isna(name): continue
                name = str(name).strip()
                
                # 1. 处理员工信息 (EmployeeBase)
                if name not in self.employees:
                    # 优先从历史信息中获取部门和入职日期
                    hist_info = self.historical_employee_info.get(name, {})
                    
                    # 部门：历史表 > 考勤表 "部门" > 考勤表 "区域" (注意用户说部门是部门，区域是区域，不要混淆)
                    # 考勤表中如果只有 "区域" 列，且没有 "部门" 列，怎么处理？
                    # 用户："部门是部门，区域是区域，这块不要录错了，部门和入职时间可以从{历史年份}年员工年休假统计表这个参考的历史总表中获取"
                    # 这意味着如果历史表有部门，就用历史表的。如果历史表没有，考勤表可能有 "部门" 列吗？
                    # 我们先看历史表。
                    dept = hist_info.get("department")
                    
                    # 入职日期：历史表 > 推断 (最早出现月份的1号)
                    join_date_raw = hist_info.get("join_date")
                    
                    # 如果没有历史入职日期，推断为本月1号 (作为最早出现月份)
                    if join_date_raw is None:
                        join_date_raw = f"{target_year}-{month:02d}-01"
                    
                    # 修正：期初余额的读取逻辑
                    # 只有在处理 1 月份考勤表时，才读取"截止到上月底剩余未休"作为年度增量
                    # 对于非1月份出现的员工（中途入职或1月无记录），该值默认为0
                    
                    opening_balance_val = 0.0
                    if month == 1:
                        col_balance = '截止到上月底剩余未休'
                        # 模糊匹配列名
                        matched_col = None
                        for col in row.index:
                            if col_balance in str(col):
                                matched_col = col
                                break
                                
                        if matched_col and pd.notna(row[matched_col]):
                            try:
                                opening_balance_val = float(row[matched_col])
                            except:
                                pass
                    
                    emp = EmployeeBase(
                        index=index_counter,
                        name=name,
                        department=dept,
                        join_date=join_date_raw,
                        opening_balance=opening_balance_val
                    )
                    
                    # 更新 opening_balance 为计算值 (确保报表显示正确)
                    # 注意：EmployeeBase 的 calculated_opening_balance 已经包含了 opening_balance_val
                    # 这里不需要再赋值回去，因为 computed_field 是只读的属性
                    # 但为了让 models.py 中的 remaining_balance 使用正确的值，我们需要确保 emp.opening_balance 保持原始读取值
                    # models.py 中的 calculated_opening_balance = flexible_quota + opening_balance
                    # 所以 emp.opening_balance 应该存储增量值 (opening_balance_val)
                    
                    self.employees[name] = emp
                    index_counter += 1
                
                # 2. 处理月度考勤 (MonthlyAttendance)
                actual_leave = 0.0
                if '休假天数' in row and pd.notna(row['休假天数']):
                    try: actual_leave = float(row['休假天数'])
                    except: pass
                
                # 提取额外字段 (仅供参考，不参与核心计算了)
                balance_end_last = None
                if '截止到上月底剩余未休' in row and pd.notna(row['截止到上月底剩余未休']):
                    try: balance_end_last = float(row['截止到上月底剩余未休'])
                    except: pass
                
                start_time = row.get('出场时间')
                end_time = row.get('返场时间')
                
                snapshot_inc = None
                if '年剩余假期（含法定节假日）' in row and pd.notna(row['年剩余假期（含法定节假日）']):
                    try: snapshot_inc = float(row['年剩余假期（含法定节假日）'])
                    except: pass
                
                snapshot_exc = None
                if '年剩余假期（不含未来法定节假日）' in row and pd.notna(row['年剩余假期（不含未来法定节假日）']):
                    try: snapshot_exc = float(row['年剩余假期（不含未来法定节假日）'])
                    except: pass

                # 提取节假日状态
                # present_holidays 现在包含了所有法定假日
                holiday_statuses = {}
                for h_date in present_holidays:
                    date_str = f"{h_date.month}.{h_date.day}"
                    pattern = r"(?:^|\D)" + re.escape(date_str) + r"(?:\D|$)"
                    
                    matched_col = None
                    for col in df.columns:
                        if pd.notna(col):
                            col_str = str(col)
                            if re.search(pattern, col_str):
                                matched_col = col
                                break
                    
                    if matched_col:
                        val = row[matched_col]
                        status = str(val).strip() if pd.notna(val) else '班'
                        holiday_statuses[h_date] = status
                    else:
                        # 如果考勤表中没有这一列，通常意味着这不是一个需要在该月特别标注的日子，
                        # 或者该月数据尚未生成。
                        # 为了统计表展示，我们将其标记为 '班' (默认工作状态) 或者 '' (空，表示未知/未发生)
                        # 用户之前的需求是：空缺月份显示为空值
                        # 如果考勤表存在（df不为空），但没有这列，说明可能是普通工作日或未记录
                        # 暂时设为 '班' 以保持兼容，或者 ''？
                        # 用户说 "未在考勤表中体现的月份，为什么没有相关的休假列？这些应该是固定的...这些空缺月份的休假数值应显示为空值"
                        # 这里是处理 "已存在的考勤表"。
                        # 如果考勤表里没这列，但确实是法定假，可能是还没发生，或者没记录。
                        # 既然是法定假，理论上应该有记录。如果没记录，可能是还没到那天？
                        # 我们设为 '班'，但在 Presentation 层，如果整个月份都没记录，会显示为空。
                        # 但如果月份存在，只是列不存在，那应该显示什么？
                        # 假设如果列不存在，就是正常上班？
                        holiday_statuses[h_date] = '班'

                record = MonthlyAttendance(
                    month=month,
                    actual_leave_days=actual_leave,
                    holiday_statuses=holiday_statuses,
                    balance_end_of_last_month=balance_end_last,
                    start_time=start_time,
                    end_time=end_time,
                    snapshot_annual_balance_inc_holidays=snapshot_inc,
                    snapshot_annual_balance_exc_future_holidays=snapshot_exc
                )
                
                if name not in self.attendance_records:
                    self.attendance_records[name] = {}
                self.attendance_records[name][month] = record
        
        print(f"成功处理 {len(self.employees)} 名员工的考勤数据。")

    def get_full_reports(self) -> List[EmployeeAnnualReport]:
        reports = []
        sorted_employees = sorted(self.employees.values(), key=lambda x: x.index)
        
        for emp in sorted_employees:
            monthly_data = self.attendance_records.get(emp.name, {})
            full_monthly_data = {}
            for m in range(1, 13):
                if m in monthly_data:
                    full_monthly_data[m] = monthly_data[m]
                else:
                    full_monthly_data[m] = MonthlyAttendance(month=m)
            
            report = EmployeeAnnualReport(
                employee=emp,
                monthly_records=full_monthly_data
            )
            reports.append(report)
        return reports
