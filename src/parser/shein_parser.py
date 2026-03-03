# -*- coding: utf-8 -*-
"""
SHEIN 平台解析器

解析 已完成账单-账单商品维度-供货价-xxx.xlsx 文件
"""
import pandas as pd
from decimal import Decimal
from pathlib import Path
from typing import List, Tuple
from datetime import datetime
import re
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from src.models import Transaction, TransactionType


class SheinParser:
    """SHEIN 账单解析器"""
    
    def __init__(self):
        self.platform = 'shein'
    
    def parse(self, file_path: str) -> Tuple[List[Transaction], dict]:
        """解析 SHEIN 账单 Excel 文件"""
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 从文件名解析店铺名和站点
        store_name, site = self._extract_store_info(file_path.name)
        
        transactions = []
        all_months = set()
        
        try:
            # SHEIN 文件首行可能是汇总，需要跳过
            df = pd.read_excel(file_path, header=1)
        except Exception as e:
            return [], {'error': str(e)}
        
        if df.empty:
            return [], {'error': '空文件'}
        
        # 列名可能是文字，需要找到对应列
        # 常见列名映射
        col_map = {}
        for col in df.columns:
            col_str = str(col).lower()
            if '订单号' in str(col) or 'order' in col_str:
                col_map['order_id'] = col
            elif '应收金额' in str(col):
                col_map['amount'] = col
            elif '打款日期' in str(col) or '签收' in str(col):
                col_map['date'] = col
            elif '账单类型' in str(col):
                col_map['type'] = col
            elif '站点' in str(col):
                col_map['site'] = col
        
        # 如果找不到关键列，尝试用位置
        if 'amount' not in col_map:
            # 尝试找最后一个数值列
            for col in reversed(df.columns.tolist()):
                if df[col].dtype in ['float64', 'int64']:
                    col_map['amount'] = col
                    break
        
        if 'amount' not in col_map:
            return [], {'error': '找不到金额列'}
        
        for idx, row in df.iterrows():
            try:
                # 解析金额
                amount_val = row.get(col_map.get('amount'))
                if pd.isna(amount_val):
                    continue
                
                amount = Decimal(str(amount_val))
                
                # 解析时间
                date_time = None
                date_col = col_map.get('date')
                if date_col:
                    time_val = row.get(date_col)
                    if time_val and not pd.isna(time_val):
                        try:
                            date_time = pd.to_datetime(time_val)
                        except:
                            pass
                
                if date_time:
                    all_months.add(date_time.strftime('%Y-%m'))
                
                # 交易类型
                type_val = str(row.get(col_map.get('type', ''), 'ORDER')).strip()
                txn_type = TransactionType.REFUND if '退款' in type_val else TransactionType.ORDER
                
                txn = Transaction(
                    date_time=date_time,
                    type=txn_type,
                    type_raw=type_val,
                    order_id=str(row.get(col_map.get('order_id', ''), '')).strip(),
                    total=amount,
                    platform=self.platform,
                    store_id=store_name.lower().replace(' ', '_'),
                    store_name=store_name,
                    currency=self._site_to_currency(site),
                    source_file=str(file_path),
                    row_number=idx + 2,
                )
                transactions.append(txn)
                
            except Exception as e:
                continue
        
        meta = {
            'platform': self.platform,
            'store_name': store_name,
            'site': site,
            'currency': self._site_to_currency(site),
            'year_month': list(all_months)[0] if all_months else '',
            'total_records': len(transactions),
            'source_file': str(file_path),
        }
        
        return transactions, meta
    
    def _extract_store_info(self, filename: str) -> Tuple[str, str]:
        """从文件名提取店铺名和站点"""
        # 示例: 天基希音UK 已完成账单-账单商品维度-供货价-2025-08-05+02_55--360142954.xlsx
        
        # 提取站点
        site_match = re.search(r'(UK|DE|FR|IT|ES|US)', filename, re.IGNORECASE)
        site = site_match.group(1).upper() if site_match else 'GLOBAL'
        
        # 提取店铺名
        store_name = None
        for marker in ['已完成账单', '账单商品维度', '账单明细']:
            store_match = re.match(rf'^(.+?)\s*{marker}', filename)
            if store_match:
                store_name = store_match.group(1).strip()
                break

        if not store_name:
            store_name = filename.split('.')[0]
        
        return store_name, site
    
    def _site_to_currency(self, site: str) -> str:
        """站点到币种的映射"""
        currency_map = {
            'UK': 'GBP',
            'DE': 'EUR', 'FR': 'EUR', 'IT': 'EUR', 'ES': 'EUR',
            'US': 'USD',
        }
        return currency_map.get(site, 'USD')


# 测试
if __name__ == '__main__':
    parser = SheinParser()
    test_file = r'd:\app\收入核算系统\跨境电商数据\部分店铺收入\多平台\多平台收入-7月\天基希音UK 已完成账单-账单商品维度-供货价-2025-08-05+02_55--360142954.xlsx'
    
    txns, meta = parser.parse(test_file)
    print(f"店铺: {meta.get('store_name')}")
    print(f"站点: {meta.get('site')}")
    print(f"币种: {meta.get('currency')}")
    print(f"解析记录数: {len(txns)}")
    
    if txns:
        net = sum(t.total for t in txns)
        print(f"平台净结算: {net} {meta.get('currency')}")
