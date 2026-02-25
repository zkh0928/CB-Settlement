# -*- coding: utf-8 -*-
"""
Phase 1 多平台收入核算入口

支持平台:
- Amazon (CSV)
- Temu (Excel)
- SHEIN (Excel)
- 托管店铺 (Excel)
- 速卖通 (Excel)
"""
import sys
import warnings
from pathlib import Path
from decimal import Decimal
from collections import defaultdict
import pandas as pd

# 忽略 openpyxl 警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 添加源码路径
sys.path.insert(0, str(Path(__file__).parent))

from src.parser.multi_platform_scanner import MultiPlatformScanner, PlatformFile
from src.parser.amazon_parser import AmazonCSVParser
from src.parser.temu_parser import TemuParser
from src.parser.shein_parser import SheinParser
from src.parser.managed_store_parser import ManagedStoreParser
from src.parser.aliexpress_parser import AliExpressParser
from src.calculator.revenue_calculator import RevenueCalculator


def run_phase1_multiplatform():
    """运行 Phase 1 多平台核算"""
    
    print("=" * 70)
    print("跨境电商收入核算系统 - Phase 1 多平台版")
    print("=" * 70)
    
    # 1. 扫描文件
    scanner = MultiPlatformScanner([
        r'C:\Users\EDY\Desktop\CB-Settlement\data\部分店铺收入\亚马逊',
        r'C:\Users\EDY\Desktop\CB-Settlement\data\部分店铺收入\多平台',
        r'C:\Users\EDY\Desktop\CB-Settlement\data\部分店铺收入\速卖通',
    ])
    
    platform_files = scanner.scan()
    
    total_files = sum(len(files) for files in platform_files.values())
    print(f"\n扫描到 {total_files} 个文件")
    for platform, files in platform_files.items():
        if files:
            print(f"  - {platform}: {len(files)} 个")
    
    # 2. 解析并计算
    calculator = RevenueCalculator()
    results = []  # (platform, store, month, currency, net_settlement, transfer)
    errors = []
    
    parsers = {
        'amazon': AmazonCSVParser(),
        'temu': TemuParser(),
        'shein': SheinParser(),
        'managed_store': ManagedStoreParser(),
        'aliexpress': AliExpressParser(),
    }
    
    for platform, files in platform_files.items():
        parser = parsers.get(platform)
        if not parser:
            continue
        
        for pf in files:
            try:
                # 解析 - Amazon 返回 ParseResult，其他返回 (txns, meta) 元组
                if platform == 'amazon':
                    parse_result = parser.parse(pf.file_path)
                    if not parse_result.success:
                        errors.append((pf.file_path, '; '.join(parse_result.errors)))
                        continue
                    txns = parse_result.transactions
                    meta = {
                        'store_name': parse_result.store_name,
                        'site': parse_result.marketplace,
                        'currency': parse_result.currency,
                        'year_month': parse_result.year_month,
                    }
                else:
                    txns, meta = parser.parse(pf.file_path)
                
                if not txns:
                    continue
                
                # 计算 - 分离 Transfer
                included = [t for t in txns if not t.is_excluded_from_revenue()]
                excluded = [t for t in txns if t.is_excluded_from_revenue()]
                
                net_settlement = sum(t.total for t in included)
                transfer_amount = sum(t.total for t in excluded)
                
                store_name = meta.get('store_name', pf.store_name)
                currency = meta.get('currency', 'USD')
                # 解析器未解析出月份时（如日期列为空），用扫描器从文件夹得到的月份
                year_month = meta.get('year_month') or pf.year_month
                site = meta.get('site', '')
                
                results.append({
                    'platform': platform,
                    'store_name': store_name,
                    'site': site,
                    'year_month': year_month,
                    'currency': currency,
                    'total_records': len(txns),
                    'included_records': len(included),
                    'excluded_records': len(excluded),
                    'net_settlement': float(net_settlement),
                    'transfer_amount': float(transfer_amount),
                })
                
                
                print(f"✓ {platform:12s} | {store_name[:15]:15s} | {year_month:7s} | {net_settlement:>12,.2f} {currency}")
                
            except Exception as e:
                errors.append((pf.file_path, str(e)))
    
    # 3. 生成报表
    print(f"\n成功处理: {len(results)} 个文件")
    if errors:
        print(f"失败: {len(errors)} 个文件")
    
    if results:
        df = pd.DataFrame(results)
        
        # 汇总统计
        print("\n" + "=" * 70)
        print("各平台月度汇总")
        print("=" * 70)
        
        summary = df.groupby(['platform', 'currency']).agg({
            'net_settlement': 'sum',
            'total_records': 'sum',
        }).reset_index()
        
        for _, row in summary.iterrows():
            print(f"{row['platform']:15s} | {row['net_settlement']:>15,.2f} {row['currency']:3s} | {int(row['total_records']):>6d} 条")
        
        # 输出 Excel
        output_path = r'C:\Users\EDY\Desktop\CB-Settlement\output\月度核算报表_Phase1_多平台.xlsx'
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        # 详细数据
        df_output = df[['platform', 'store_name', 'site', 'year_month', 'currency', 
                        'total_records', 'included_records', 'net_settlement', 'transfer_amount']]
        df_output.columns = ['平台', '店铺', '站点', '月份', '币种', 
                             '交易数', '参与计算', '平台净结算', '提现金额']
        
        # 如果主输出文件被占用（例如已在 Excel 中打开），
        # 自动退回到带后缀的备份文件，避免整个流程报错中断。
        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                df_output.to_excel(writer, sheet_name='详细数据', index=False)
                summary.to_excel(writer, sheet_name='平台汇总', index=False)
            final_path = output_path
        except PermissionError:
            backup_path = r'C:\Users\EDY\Desktop\CB-Settlement\output\月度核算报表_Phase1_多平台_auto.xlsx'
            with pd.ExcelWriter(backup_path, engine='xlsxwriter') as writer:
                df_output.to_excel(writer, sheet_name='详细数据', index=False)
                summary.to_excel(writer, sheet_name='平台汇总', index=False)
            final_path = backup_path
            print(f"\n注意：原始报表文件被占用，已自动写入备份文件: {backup_path}")
        
        print(f"\n报表已生成: {final_path}")
    
    print("=" * 70)


if __name__ == '__main__':
    run_phase1_multiplatform()
