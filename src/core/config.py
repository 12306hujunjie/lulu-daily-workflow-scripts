from typing import Dict, List
from datetime import date
from src.core.holiday_service import HolidayService

# --- 配置常量 ---
# 默认年份 (如果无法从文件自动检测)
DEFAULT_TARGET_YEAR = 2026

class GlobalConfig:
    _target_year: int = DEFAULT_TARGET_YEAR
    _holiday_config: Dict[int, List[date]] = {}

    @classmethod
    def set_year(cls, year: int):
        cls._target_year = year
        try:
            print(f"正在获取 {year} 年法定节假日配置...")
            cls._holiday_config = HolidayService.get_holidays_for_year(year)
        except Exception as e:
            print(f"警告: 获取节假日数据失败, 将使用空配置. Error: {e}")
            cls._holiday_config = {}

    @classmethod
    def get_year(cls) -> int:
        return cls._target_year

    @classmethod
    def get_holidays(cls) -> Dict[int, List[date]]:
        if not cls._holiday_config:
            # Lazy load if not set
            cls.set_year(cls._target_year)
        return cls._holiday_config

# 默认每月弹性假额度（天）
DEFAULT_MONTHLY_FLEX_QUOTA = 4.0

# 默认老员工年度弹性假额度（天）
DEFAULT_ANNUAL_FLEX_QUOTA = 48.0
