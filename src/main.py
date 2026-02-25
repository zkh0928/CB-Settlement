# -*- coding: utf-8 -*-
"""
收入核算系统 - 主程序 (Phase 1 MVP)

流程:
1. 扫描CSV文件
2. 解析文件
3. 执行核算(过滤Transfer)
4. 多店铺聚合
5. 导出Excel报表
"""
import sys
from pathlib import Path
from typing import List

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.parser import AmazonCSVParser
from src.calculator import RevenueCalculator, MonthlyAggregator
from src.reporter import ExcelExporter
from src.models import StoreMonthlyResult


class RevenueAccountingApp:
    """收入核算系统应用"""
    
    def __init__(self):
        self.parser = AmazonCSVParser()
        self.calculator = RevenueCalculator()
        self.aggregator = MonthlyAggregator()
        self.exporter = ExcelExporter()
    
    def run(self, input_dir: str, output_file: str):
        """运行核算流程"""
        print("="*60)
        print("跨境电商与店铺收入核算系统 - Phase 1 MVP")
        print("="*60)
        
        input_path = Path(input_dir)
        if not input_path.exists():
            print(f"错误: 输入目录不存在 {input_dir}")
            return
        
        # 1. 扫描文件
        files = list(input_path.glob('**/*.csv'))
        print(f"扫描到 {len(files)} 个CSV文件")
        
        store_results: List[StoreMonthlyResult] = []
        parsed_count = 0
        
        # 2. 逐个处理
        for f in files:
            # 解析
            parse_result = self.parser.parse(str(f))
            if not parse_result.success:
                err_msg = parse_result.errors[0] if parse_result.errors else "未知错误"
                print(f"X 解析失败: {f.name} - {err_msg}")
                continue
            
            parsed_count += 1
            
            # 核算
            calc_result = self.calculator.calculate(
                transactions=parse_result.transactions,
                store_id=parse_result.store_id,
                store_name=parse_result.store_name,
                year_month=parse_result.year_month,
                currency=parse_result.currency
            )
            
            if not calc_result.verification_passed:
                print(f"! 校验警告: {f.name}")
                for note in calc_result.verification_notes:
                    print(f"  - {note}")
            
            # 聚合
            store_result = self.aggregator.aggregate_store(calc_result)
            store_results.append(store_result)
            
            # 简单进度日志
            print(f"√ 已处理: {store_result.store_name} ({store_result.year_month}) "
                  f"- 净结算: {store_result.platform_net_settlement} {store_result.currency}")
        
        print("-" * 60)
        print(f"处理完成: {parsed_count}/{len(files)} 个文件")
        
        # 3. 导出报表
        if store_results:
            report_out = self.exporter.export(store_results, output_file)
            if report_out.success:
                print(f"\n报表已生成: {report_out.file_path}")
                print(f"包含店铺数: {report_out.total_stores}")
            else:
                print(f"\n报表生成失败: {report_out.message}")
        else:
            print("\n无有效数据，未生成报表")
            
        print("="*60)


def main():
    """入口函数"""
    app = RevenueAccountingApp()
    
    # 配置路径
    input_dir = r'C:\Users\EDY\Desktop\CB-Settlement\data\部分店铺收入\亚马逊'
    output_file = r'C:\Users\EDY\Desktop\CB-Settlement\output\月度核算报表_Phase1.xlsx'
    
    app.run(input_dir, output_file)


if __name__ == '__main__':
    main()
