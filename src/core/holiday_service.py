from datetime import date
from typing import Dict, List

import holidays


class HolidayService:
    @staticmethod
    def get_holidays_for_year(year: int) -> Dict[int, List[date]]:
        cn_holidays = holidays.China(years=year)
        holidays_map: Dict[int, List[date]] = {}

        for d, name in cn_holidays.items():
            lower_name = str(name).lower()
            if "day off" in lower_name or "observed" in lower_name:
                continue
            holidays_map.setdefault(d.month, []).append(d)

        for month in holidays_map:
            holidays_map[month].sort()

        return holidays_map
