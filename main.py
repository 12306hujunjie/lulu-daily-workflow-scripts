import os
import sys
from src.dal.reader import ExcelReader
from src.presentation.excel_generator import ExcelReportGenerator
from src.core.config import GlobalConfig

# --- 全局配置 ---
# 输入文件路径
ATTENDANCE_FILE = "外围生物安全考勤表2026.xlsx"
# 历史统计表路径 (如果存在，程序会自动使用)
# 假设年份是 N，历史表应该是 N-1 年的
# 这里我们采用一个启发式规则：先尝试检测考勤表年份，然后构造历史表文件名
HISTORICAL_STATS_FILE_TEMPLATE = "{year}年员工年休假统计表-12月.xlsx"

def main():
    """
    主程序入口。
    编排整个数据处理流程：读取 -> 处理 -> 生成报表。
    """
    print("=== 开始执行年休假统计流程 (自动识别年份) ===")
    
    # 1. 预先读取考勤表以获取年份
    # 为了构建历史表路径，我们需要先知道年份。
    # 这里我们简单地先初始化 Reader，load 考勤表，拿到年份后，再决定是否 reload 历史表。
    
    reader = ExcelReader(
        attendance_file_path=ATTENDANCE_FILE
    )
    
    try:
        # 第一步：只加载考勤表，为了识别年份
        reader.load_files()
        
        target_year = GlobalConfig.get_year()
        print(f"当前统计目标年份: {target_year}")
        
        # 第二步：构建历史表路径并尝试加载
        prev_year = target_year - 1
        prev_year_short = str(prev_year)[-2:]
        
        # 关键词列表：只要文件名包含其中之一即可
        keywords = [
            f"{prev_year}年员工年休假统计表",
            f"{prev_year_short}年员工年休假统计表"
        ]
        
        historical_path = None
        print(f"正在寻找历史统计表，关键词: {keywords}")
        
        # 遍历当前目录下的所有文件进行模糊匹配
        for filename in os.listdir('.'):
            # 忽略非xlsx文件和临时文件
            if not filename.endswith('.xlsx') or filename.startswith('~$'):
                continue
                
            # 检查是否包含关键词
            for kw in keywords:
                if kw in filename:
                    historical_path = filename
                    break
            
            if historical_path:
                break
        
        if historical_path:
            print(f"发现历史统计表: {historical_path}")
            reader.stats_file_path = historical_path
            # 重新加载以包含历史信息
            reader.load_files()
        else:
            print(f"未找到 {prev_year} 年的历史统计表，将仅基于考勤表推断入职信息。")
        
        # 第三步：解析数据
        reader.parse_data()
        
    except Exception as e:
        print(f"数据加载或解析失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    # 3. 获取业务模型 (Core)
    reports = reader.get_full_reports()
    
    # 4. 生成报表 (Presentation)
    output_path = f"{target_year}年员工年休假统计表_Final.xlsx"
    generator = ExcelReportGenerator(output_path=output_path)
    try:
        # 将 Reader 中识别到的活跃节假日配置传递给生成器
        generator.generate(reports, reader.active_monthly_holidays)
        print("=== 流程执行成功 ===")
    except Exception as e:
        print(f"报表生成失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    sys.path.append(os.getcwd())
    main()
