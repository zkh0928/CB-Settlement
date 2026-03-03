# -*- coding: utf-8 -*-
"""
托管店铺解析器

解析 收支明细_xxx.xlsx 文件
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


class ManagedStoreParser:
    """托管店铺 收支明细 解析器"""
    
    # 费用项到交易类型的映射
    FEE_TYPE_MAP = {
        '供货款': TransactionType.ORDER,
        '售后退款': TransactionType.REFUND,
        '履约服务费': TransactionType.SERVICE_FEE,
        '技术服务费': TransactionType.SERVICE_FEE,
        '提现': TransactionType.TRANSFER,  # 排除
    }
    
    def __init__(self):
        self.platform = 'managed_store'
    
    def parse(self, file_path: str) -> Tuple[List[Transaction], dict]:
        """
        解析托管店铺收支明细 Excel 文件
        """
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        # 从文件名解析店铺名
        store_name = self._extract_store_name(file_path.name)
        
        transactions = []
        all_months = set()
        
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            return [], {'error': str(e)}
        
        if df.empty:
            return [], {'error': '空文件'}
        
        # 只保留有效列 (排除重复的Unnamed列)
        valid_cols = [c for c in df.columns if not str(c).startswith('Unnamed')]
        
        for idx, row in df.iterrows():
            try:
                # 解析费用项
                fee_item = str(row.get('费用项', '')).strip()
                if not fee_item:
                    continue
                
                # 确定交易类型
                txn_type = self.FEE_TYPE_MAP.get(fee_item, TransactionType.OTHER)
                
                # 解析金额
                amount_val = row.get('金额(CNY)', 0)
                if pd.isna(amount_val):
                    continue
                
                amount = Decimal(str(amount_val))
                
                # 解析时间
                date_time = None
                time_val = row.get('结算时间')
                if time_val and not pd.isna(time_val):
                    try:
                        if isinstance(time_val, str):
                            # 格式: 2025/07/30 08:40:25
                            date_time = datetime.strptime(time_val, '%Y/%m/%d %H:%M:%S')
                        else:
                            date_time = pd.to_datetime(time_val)
                    except:
                        pass
                
                if date_time:
                    all_months.add(date_time.strftime('%Y-%m'))
                
                txn = Transaction(
                    date_time=date_time,
                    type=txn_type,
                    type_raw=fee_item,
                    order_id=str(row.get('订单号', '')).strip() if row.get('订单号') else '',
                    total=amount,
                    platform=self.platform,
                    store_id=store_name.lower().replace(' ', '_'),
                    store_name=store_name,
                    currency='CNY',
                    source_file=str(file_path),
                    row_number=idx + 2,
                )
                transactions.append(txn)
                
            except Exception as e:
                continue
        
        meta = {
            'platform': self.platform,
            'store_name': store_name,
            'site': 'GLOBAL',
            'currency': 'CNY',
            'year_month': list(all_months)[0] if all_months else '',
            'total_records': len(transactions),
            'source_file': str(file_path),
        }
        
        return transactions, meta
    
    def _extract_store_name(self, filename: str) -> str:
        """从文件名提取店铺名"""
        # 示例: 天基托管 收支明细_20250701-20250731.xlsx
        match = re.match(r'^(.+?)\s*收支明细', filename)
        if match:
            return match.group(1).strip()

        # 兼容: 天基托管 Sc594baac466445ae8f0cd0a0cabd744af.xlsx
        match = re.match(r'^(.+?)\s+Sc[0-9a-f]+', filename, re.IGNORECASE)
        if match:
            return match.group(1).strip()

        match = re.match(r'^(.+?托管)', filename)
        if match:
            return match.group(1).strip()

        return filename.split('.')[0]


# 测试
if __name__ == '__main__':
    parser = ManagedStoreParser()
    test_file = r'd:\app\收入核算系统\跨境电商数据\部分店铺收入\多平台\多平台收入-7月\天基托管 收支明细_20250701-20250731.xlsx'
    
    txns, meta = parser.parse(test_file)
    print(f"店铺: {meta.get('store_name')}")
    print(f"解析记录数: {len(txns)}")
    
    if txns:
        # 排除 Transfer 后的净结算
        included = [t for t in txns if not t.is_excluded_from_revenue()]
        excluded = [t for t in txns if t.is_excluded_from_revenue()]
        
        net = sum(t.total for t in included)
        print(f"参与计算: {len(included)} 条")
        print(f"排除: {len(excluded)} 条")
        print(f"平台净结算: {net} CNY")
        
        # 按费用项统计
        by_type = {}
        for t in included:
            by_type[t.type_raw] = by_type.get(t.type_raw, Decimal('0')) + t.total
        
        print("\n按费用项汇总:")
        for k, v in by_type.items():
            print(f"  {k}: {v}")
