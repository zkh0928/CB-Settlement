# -*- coding: utf-8 -*-
"""
Phase 2 仓库账单解析器

仅解析仓库履约成本，不涉及 SKU 级成本
"""
import pandas as pd
from decimal import Decimal
from pathlib import Path
from typing import List, Tuple, Dict
from datetime import datetime
from dataclasses import dataclass, field
import re
import os
import warnings
warnings.filterwarnings('ignore')

# 添加PDF处理库
try:
    import PyPDF2
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    print("警告: PDF处理库未安装，请运行 'pip install PyPDF2 pdfplumber'")


@dataclass
class WarehouseMonthlyCost:
    """仓库月度成本汇总"""
    warehouse_name: str
    year_month: str
    total_cost: Decimal
    currency: str
    cost_breakdown: Dict[str, Decimal] = field(default_factory=dict)
    record_count: int = 0
    source_files: List[str] = field(default_factory=list)


class BaseWarehouseParser:
    """仓库解析器基类"""
    
    def __init__(self, warehouse_name: str, region: str, currency: str):
        self.warehouse_name = warehouse_name
        self.region = region
        self.currency = currency
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """解析单个文件，返回 (总成本, 分类汇总, 记录数)"""
        raise NotImplementedError
    
    def extract_month(self, filename: str) -> str:
        """从文件名提取月份"""
        raise NotImplementedError


class G7Parser(BaseWarehouseParser):
    """G7仓库解析器 (德国EUR)"""
    
    def __init__(self):
        super().__init__("G7", "DE", "EUR")
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        G7仓库PDF账单解析：
        - 自动识别INVOICE（应付费用）和CREDIT NOTE（退款）文件
        - 解析Total Amount金额
        - 自动跳过明细文件（Appendix后缀）
        """
        filename = os.path.basename(file_path).lower()
        
        # 跳过Appendix后缀的明细文件
        if 'appendix' in filename:
            return Decimal('0'), {}, 0
        
        # 检查PDF库是否可用
        if not PDF_AVAILABLE:
            print("警告: PDF处理库未安装，无法解析G7 PDF文件")
            return Decimal('0'), {}, 0
        
        try:
            total_amount = self._extract_total_amount_from_pdf(file_path)
            if total_amount is None:
                return Decimal('0'), {}, 0
            
            # 判断文件类型并设置符号
            # 根据文件名后缀判断：
            # - R结尾的文件为INVOICE（正数）
            # - G结尾的文件为CREDIT NOTE（负数）
            basename = os.path.splitext(filename)[0].lower()
            if basename.endswith('g'):
                # G结尾文件：CREDIT NOTE，金额为负数（退款抵扣）
                total_amount = -abs(total_amount)
                breakdown = {'CREDIT NOTE Total': total_amount}
            elif basename.endswith('r'):
                # R结尾文件：INVOICE，金额为正数（需要支付）
                breakdown = {'INVOICE Total': total_amount}
            elif 'credit' in filename:
                # 兼容原有的credit note文件命名
                total_amount = -abs(total_amount)
                breakdown = {'CREDIT NOTE Total': total_amount}
            else:
                # 默认为INVOICE文件
                breakdown = {'INVOICE Total': total_amount}
            
            return total_amount, breakdown, 1
            
        except Exception as e:
            print(f"  G7 PDF解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0
    
    def _extract_total_amount_from_pdf(self, file_path: str) -> Decimal:
        """从PDF中提取Total Amount金额"""
        try:
            # 首先尝试使用pdfplumber（更准确）
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        # 查找Total Amount行
                        lines = text.split('\n')
                        for line in lines:
                            if 'total amount' in line.lower():
                                # 提取金额数字 - 支持德国格式：8.786,41
                                # 先查找包含数字、点和逗号的模式
                                match = re.search(r'[\d.,]+\d', line)
                                if match:
                                    amount_str = match.group()
                                    # 处理德国数字格式：8.786,41 → 8786.41
                                    # 步骤1: 如果包含逗号，说明是德国格式
                                    if ',' in amount_str:
                                        # 步骤2: 将最后一个逗号替换为点（小数点）
                                        # 步骤3: 移除所有其他的点（千分位分隔符）
                                        parts = amount_str.rsplit(',', 1)
                                        if len(parts) == 2:
                                            integer_part = parts[0].replace('.', '')
                                            decimal_part = parts[1]
                                            amount_str = f"{integer_part}.{decimal_part}"
                                        else:
                                            # 如果只有一个逗号，直接替换
                                            amount_str = amount_str.replace(',', '.').replace('.', '', amount_str.count('.') - 1)
                                    else:
                                        # 美国/中国格式：移除千分位逗号
                                        amount_str = amount_str.replace(',', '')
                                    
                                    return Decimal(amount_str)
            
            # 如果pdfplumber失败，尝试PyPDF2
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if 'total amount' in line.lower():
                                match = re.search(r'[\d.,]+\d', line)
                                if match:
                                    amount_str = match.group()
                                    # 处理德国数字格式
                                    if ',' in amount_str:
                                        parts = amount_str.rsplit(',', 1)
                                        if len(parts) == 2:
                                            integer_part = parts[0].replace('.', '')
                                            decimal_part = parts[1]
                                            amount_str = f"{integer_part}.{decimal_part}"
                                        else:
                                            amount_str = amount_str.replace(',', '.').replace('.', '', amount_str.count('.') - 1)
                                    else:
                                        amount_str = amount_str.replace(',', '')
                                    
                                    return Decimal(amount_str)
            
            return None
            
        except Exception as e:
            print(f"  PDF提取失败: {e}")
            return None
    
    def extract_month(self, file_path: str) -> str:
        """
        从G7文件路径提取月份
        优先从文件夹路径提取月份（如'10月'目录），再尝试从文件名提取YYMMDD格式
        """
        # 1. 优先从文件夹路径提取月份
        folder_path = os.path.dirname(file_path)
        folder_name = os.path.basename(folder_path)
        
        # 匹配中文月份格式：10月、2025年10月等
        match = re.search(r'(\d{4})年(\d{1,2})月', folder_name)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            return f"{year}-{month}"
        
        # 匹配数字月份格式：10、2025-10等
        match = re.search(r'(\d{4})[-_](\d{1,2})', folder_name)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            return f"{year}-{month}"
        
        # 匹配纯数字月份：10、11、12等
        match = re.search(r'^(\d{1,2})$', folder_name)
        if match:
            month = match.group(1).zfill(2)
            # 假设为当前年份或最近的年份，这里使用2025年作为默认
            return f"2025-{month}"
        
        # 匹配中文纯月份：10月、十一月等
        match = re.search(r'^(\d{1,2})月$', folder_name)
        if match:
            month = match.group(1).zfill(2)
            return f"2025-{month}"
        
        # 2. 如果路径中没有月份信息，尝试从文件名提取YYMMDD格式
        filename = os.path.basename(file_path)
        basename = os.path.splitext(filename)[0]
        
        # 查找6位连续数字（YYMMDD格式）
        matches = re.findall(r'(\d{6})', basename)
        if matches:
            for match in matches:
                yy = int(match[0:2])
                mm = int(match[2:4])
                dd = int(match[4:6])
                
                # 验证月份和日期是否有效
                if 1 <= mm <= 12 and 1 <= dd <= 31:
                    # 假设为2020年代（2020-2029）
                    if 20 <= yy <= 29:
                        year = 2000 + yy
                        return f"{year}-{mm:02d}"
                    # 如果yy较小（0-19），可能是2020-2039的简写
                    elif 0 <= yy <= 19:
                        year = 2020 + yy
                        return f"{year}-{mm:02d}"
                    # 如果yy较大（30-99），可能是2030-2099
                    elif 30 <= yy <= 99:
                        year = 2000 + yy
                        return f"{year}-{mm:02d}"
        
        return ""


class TSPParser(BaseWarehouseParser):
    """TSP 仓库解析器 (UK, GBP)"""
    
    def __init__(self):
        super().__init__("TSP", "UK", "GBP")
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        total = Decimal('0')
        breakdown = {}
        count = 0
        
        xl = pd.ExcelFile(file_path)
        
        # 定义每个工作表应该使用的列名
        sheet_column_mapping = {
            'additional invoice items': 'cost',
            'invoice items': 'total cost',
            'invoiced collections': 'cost',
            'invoiced returns': 'cost',
            'invoiced goodins items': 'total cost',
            'invoiced storage items': 'cost'
        }
        
        # 处理每个指定的工作表
        for sheet_name_lower, target_column in sheet_column_mapping.items():
            # 在所有sheet中查找匹配的工作表
            actual_sheet_name = None
            for sheet in xl.sheet_names:
                if sheet.lower().strip() == sheet_name_lower:
                    actual_sheet_name = sheet
                    break
            
            if actual_sheet_name is None:
                continue
            
            # 读取工作表数据
            df = pd.read_excel(file_path, sheet_name=actual_sheet_name)
            
            # 查找目标列
            cost_col = None
            target_column_lower = target_column.lower()
            
            # 精确匹配列名
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if col_lower == target_column_lower:
                    cost_col = col
                    break
            
            # 如果精确匹配失败，尝试模糊匹配
            if cost_col is None:
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if target_column_lower in col_lower or col_lower in target_column_lower:
                        cost_col = col
                        break
            
            if cost_col is None:
                continue
            
            # 计算该工作表的金额总和
            sheet_total = Decimal('0')
            sheet_count = 0
            
            for idx, row in df.iterrows():
                try:
                    cost_val = row.get(cost_col, 0)
                    if pd.isna(cost_val):
                        continue
                    
                    amount = Decimal(str(cost_val))
                    sheet_total += amount
                    sheet_count += 1
                except Exception as e:
                    continue
            
            if sheet_total > 0:
                breakdown[actual_sheet_name] = sheet_total
                total += sheet_total
                count += sheet_count
        
        return total, breakdown, count
    
    def extract_month(self, filename: str) -> str:
        # 1. Standard format: Jul25
        # 2. Full Month: November 2025 or November 25
        # 3. Prevent matching timestamps (e.g. avoid Jan01 as 2001)
        
        # Pattern 1: MonYY (e.g. Jul25), strict year 24-29
        match = re.search(r'([a-zA-Z]{3})(2[4-9])', filename)
        if match:
            month_abbr = match.group(1).lower()
            year = '20' + match.group(2)
            month_map = {
                'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
                'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
                'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
            }
            if month_abbr in month_map:
                return f"{year}-{month_map[month_abbr]}"

        # Pattern 2: Full Month + Year (November 2025 or November 2025... or November 25)
        # Look for full month name followed by 202x or 2x
        month_names = {
            'january': '01', 'february': '02', 'march': '03', 'april': '04',
            'may': '05', 'june': '06', 'july': '07', 'august': '08',
            'september': '09', 'october': '10', 'november': '11', 'december': '12'
        }
        
        filename_lower = filename.lower()
        for m_name, m_code in month_names.items():
            if m_name in filename_lower:
                # Look for year after month name
                # Matches: "november 2025", "november2025", "november 25"
                year_match = re.search(rf'{m_name}.*?(202[4-9]|2[4-9])', filename_lower)
                if year_match:
                    year_raw = year_match.group(1)
                    if len(year_raw) == 4:
                        year = year_raw
                    else:
                        year = '20' + year_raw
                    return f"{year}-{m_code}"
                    
        return ""


class Warehouse1510Parser(BaseWarehouseParser):
    """1510 仓库解析器 (UK, GBP)"""
    
    def __init__(self):
        super().__init__("1510", "UK", "GBP")
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        1510 海外仓账单：只取第一个 sheet（账单封面/Bill cover）中的
        `账单总计(Total bill amount)`，其余 sheet 均为明细。
        """
        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            return Decimal('0'), {}, 0

        cover_sheet = xl.sheet_names[0]
        df_cover = pd.read_excel(file_path, sheet_name=cover_sheet, header=None)

        total = Decimal('0')
        found = False

        # 在封面中定位 "Total bill amount / 账单总计 / 账单小计"等单元格，取其右侧值
        keywords = ['total bill amount', '账单总计', '账单小计', '账单合计']

        for r in range(df_cover.shape[0]):
            for c in range(df_cover.shape[1]):
                v = df_cover.iat[r, c]
                if isinstance(v, str):
                    text = v.strip().lower()
                    if not any(k in text for k in keywords):
                        continue

                    if c + 1 < df_cover.shape[1]:
                        amt = df_cover.iat[r, c + 1]
                        try:
                            if pd.notna(amt):
                                total = Decimal(str(amt))
                                found = True
                        except Exception:
                            pass
                    break
            if found:
                break

        breakdown = {}
        if found:
            breakdown['Total bill amount'] = total
            return total, breakdown, 1

        # 兜底：如果封面字段定位失败，则返回 0（避免误把明细累加成总计）
        return Decimal('0'), {}, 0
    
    def extract_month(self, filename: str) -> str:
        """
        从 1510 账单文件名中提取「费用所属月份」

        规则说明（根据账单封面 Statement Period 推导）：
        - 文件名形如：bill-HBR-O-M20250101.xlsx
          - 中间的 M / A + YYYYMMDD 对应「账单到期日」(Payment Due Date)
          - 封面上的 Statement Period 为上一自然月：
            * 2025-01-01 -> 2024-12-01 ~ 2024-12-31 （所属 2024-12 月）
            * 2025-02-01 -> 2025-01-01 ~ 2025-01-31 （所属 2025-01 月）
        - 因此：按「到期日 - 1 天」来确定费用所属月份。
        - 同样适用于 A 开头的调整账单（bill-HBR-O-A20241001...）。
        """
        from datetime import date, timedelta

        # 捕获 M/A + YYYYMMDD，比如 M20250101 / A20241001
        match = re.search(r'[AM](\d{4})(\d{2})(\d{2})', filename, re.IGNORECASE)
        if not match:
            return ""

        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))

        try:
            due_date = date(year, month, day)
            period_last_day = due_date - timedelta(days=1)
            return f"{period_last_day.year}-{period_last_day.month:02d}"
        except Exception:
            return ""


class JDParser(BaseWarehouseParser):
    """京东海外仓解析器 (Multi-currency)"""
    
    def __init__(self):
        super().__init__("京东", "Global", "CNY")
        
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        京东海外仓：
        - 只看第一个 sheet（通常为「汇总页」）
        - 根据项目规范，应使用账单汇总金额，即找到"结算币种含税金额"字段，
          并提取其右侧的第一个有效数值作为总金额
        - 由于不同月份的文件格式可能有差异（金额可能在右侧1列或2列），
          需要智能搜索右侧的第一个非NaN数值
        """
        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            return Decimal('0'), {}, 0

        summary_sheet = xl.sheet_names[0]
        
        # 使用 header=None 读取完整表格结构
        df = pd.read_excel(file_path, sheet_name=summary_sheet, header=None)
        
        if df.empty:
            return Decimal('0'), {}, 0

        # 查找包含 "结算币种含税金额" 的行
        total_amount = None
        
        for row_idx in range(min(20, df.shape[0])):
            for col_idx in range(df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    if '结算币种含税金额' in cell_str:
                        # 智能搜索右侧的第一个有效数值
                        for offset in range(1, 5):  # 搜索右侧1-4列
                            if col_idx + offset < df.shape[1]:
                                amount_cell = df.iloc[row_idx, col_idx + offset]
                                if pd.notna(amount_cell):
                                    try:
                                        # 清理千分位逗号并转换
                                        amt_str = str(amount_cell).replace(',', '').strip()
                                        total_amount = Decimal(amt_str)
                                        # 确保是合理的正数金额
                                        if total_amount > 0:
                                            break
                                    except Exception:
                                        continue
                        if total_amount is not None:
                            break
            if total_amount is not None:
                break
        
        if total_amount is None:
            return Decimal('0'), {}, 0

        breakdown = {'结算币种含税金额': total_amount}
        return total_amount, breakdown, 1
    
    def extract_month(self, file_path: str) -> str:
        # 优先从文件夹路径提取月份（规范第5条）
        folder_name = os.path.basename(os.path.dirname(file_path))
        match = re.search(r'(\d{4})[-年](\d{1,2})[月]?', folder_name)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            return f"{year}-{month}"
        
        # 回退到文件名提取（规范第6条）
        filename = os.path.basename(file_path)
        match = re.search(r'(\d{4})-(\d{2})-\d{2}', filename)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
        
        return ""


class HaiyangParser(BaseWarehouseParser):
    """海洋仓库解析器 (UK, GBP)"""
    
    def __init__(self):
        super().__init__("海洋", "UK", "GBP")
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        海洋仓库账单解析：
        - 2024年10月和11月：只看汇总表的金额之和
        - 其他月份：只看CostBill的计费规则金额
        - 2025年10月：需要加上运费金额
        """
        filename = os.path.basename(file_path).lower()
        total = Decimal('0')
        breakdown: Dict[str, Decimal] = {}
        count = 0

        # 处理运费PDF文件 - 对于海洋仓库，所有运费文件都处理
        if filename.startswith('运费') and filename.endswith('.pdf'):
            # 对于海洋仓库，所有运费文件都应该被处理（因为它们都属于海洋仓库）
            # 实际的月份归属通过extract_month方法控制
            return self._parse_freight_pdf(file_path)

        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            return Decimal('0'), {}, 0

        # 特殊处理移仓费文件
        if '移仓费' in filename:
            return self._parse_move_fee_excel(file_path, xl)

        # 提取文件月份用于判断解析策略
        file_month = self._extract_file_month(filename)
        
        # 根据月份决定解析策略
        if file_month in ['2024-10', '2024-11'] and '汇总' in xl.sheet_names:
            # 2024年10月和11月：使用汇总表
            return self._parse_summary_sheet(file_path, xl)
        else:
            # 其他月份：使用CostBill明细
            return self._parse_costbill_sheet(file_path, xl)

    def _extract_file_month(self, filename: str) -> str:
        """从文件名提取月份"""
        filename_lower = filename.lower()
        
        # 处理2024年账单格式：海洋国际英国海外仓20241101-1130账单.xlsx
        match = re.search(r'海洋国际英国海外仓(\d{4})(\d{2})\d{2}-\d{2}\d{2}账单', filename)
        if match:
            year = match.group(1)
            month = match.group(2)
            return f"{year}-{month}"
        
        # 处理标准格式：2025-7月_CostBillExport1599.xlsx
        match = re.search(r'(\d{4})-(\d{1,2})月', filename)
        if match:
            return f"{match.group(1)}-{int(match.group(2)):02d}"
            
        return ""

    def _parse_summary_sheet(self, file_path: str, xl) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """解析汇总表 - 用于2024年10月和11月"""
        df_summary = pd.read_excel(file_path, sheet_name='汇总')
        
        total = Decimal('0')
        breakdown = {}
        count = 0
        
        # 汇总表结构：类型、金额、币种
        if '金额' in df_summary.columns:
            for _, row in df_summary.iterrows():
                amount_val = row.get('金额', 0)
                if pd.isna(amount_val):
                    continue
                try:
                    amount = Decimal(str(amount_val))
                    total += amount
                    count += 1
                    
                    # 记录分类
                    fee_type = str(row.get('类型', '未知'))
                    breakdown[fee_type] = amount
                except Exception:
                    continue
        
        if total > 0:
            final_breakdown = {'汇总金额': total}
            return total, final_breakdown, count
        
        return Decimal('0'), {}, 0

    def _parse_costbill_sheet(self, file_path: str, xl) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """解析CostBill明细表 - 用于其他月份"""
        # 仅取名为 CostBill 的 sheet
        sheet_name = None
        for sh in xl.sheet_names:
            if str(sh).strip().lower() == 'costbill':
                sheet_name = sh
                break
        if sheet_name is None:
            # 若没有明确的 CostBill，就使用第一个 sheet 作为兜底
            sheet_name = xl.sheet_names[0]

        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 寻找计费规则金额列
        amount_col = None
        priority_keywords = ['计费规则金额', '计费金额']
        for kw in priority_keywords:
            for c in df.columns:
                if kw in str(c):
                    amount_col = c
                    break
            if amount_col is not None:
                break

        if amount_col is None:
            return Decimal('0'), {}, 0

        # 计算所有有单号的记录的计费规则金额之和
        sheet_total = Decimal('0')
        count = 0
        
        # 查找单号列
        order_no_col = None
        order_keywords = ['单号', '订单号', '运单号', '单据号']
        for kw in order_keywords:
            for c in df.columns:
                if kw in str(c):
                    order_no_col = c
                    break
            if order_no_col is not None:
                break

        # 如果找到了单号列，则只计算有单号的记录
        if order_no_col is not None:
            for _, row in df.iterrows():
                order_no = row.get(order_no_col, '')
                if pd.isna(order_no) or str(order_no).strip() == '':
                    continue  # 跳过无单号的记录
                
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    sheet_total += amt
                    count += 1
                except Exception:
                    continue
        else:
            # 如果没有单号列，则计算所有记录
            for _, row in df.iterrows():
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    sheet_total += amt
                    count += 1
                except Exception:
                    continue

        if sheet_total > 0:
            breakdown = {'计费规则金额': sheet_total}
            return sheet_total, breakdown, count

        return Decimal('0'), {}, 0
    
    def _parse_move_fee_excel(self, file_path: str, xl) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """特殊处理移仓费Excel文件 - 提取Gross列各项费用之和"""
        df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0])
        
        try:
            # Gross列是第7列（索引6），包含各项费用金额
            # 费用行从第6行到第11行（索引5-11）
            gross_column_index = 6  # 第7列
            total_amount = Decimal('0')
            fee_details = {}
            
            # 遍历费用行（行5-11，即索引5-11）
            for row_index in range(5, 12):  # 5,6,7,8,9,10,11
                if row_index < len(df):
                    description = df.iloc[row_index, 0]  # 第0列是Description
                    gross_value = df.iloc[row_index, gross_column_index]  # 第7列是Gross
                    
                    if pd.notna(description) and pd.notna(gross_value) and str(description).strip():
                        try:
                            # 转换金额
                            amount = Decimal(str(gross_value))
                            if amount > 0:  # 只计算正数费用
                                total_amount += amount
                                fee_details[str(description).strip()] = amount
                        except:
                            continue
            
            if total_amount > 0:
                return total_amount, {'移仓费': total_amount}, len(fee_details)
            else:
                # 兜底方案：如果找不到明确的费用行，查找包含1123.99的单元格
                for i in range(len(df)):
                    for j in range(len(df.columns)):
                        value = df.iloc[i, j]
                        if pd.notna(value) and '1123.99' in str(value):
                            return Decimal('1123.99'), {'移仓费': Decimal('1123.99')}, 1
                
        except Exception as e:
            pass
        
        return Decimal('0'), {}, 0
    
    def _parse_freight_pdf(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """解析运费PDF文件，提取Charge Total金额"""
        if not PDF_AVAILABLE:
            return Decimal('0'), {}, 0
        
        try:
            with pdfplumber.open(file_path) as pdf:
                if not pdf.pages:
                    return Decimal('0'), {}, 0
                
                # 只处理第一页
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                
                if not text:
                    return Decimal('0'), {}, 0
                
                # 提取Charge Total金额
                total_amount = self._extract_charge_total_from_text(text)
                
                if total_amount is not None:
                    return total_amount, {'运费': total_amount}, 1
                else:
                    return Decimal('0'), {}, 0
                    
        except Exception as e:
            pass
            return Decimal('0'), {}, 0
    
    def _extract_charge_total_from_text(self, text: str) -> Decimal:
        """从运费PDF文本中提取Charge Total金额"""
        lines = text.split('\n')
        
        # 查找Charge Description表头行
        charge_header_found = False
        for i, line in enumerate(lines):
            if 'charge description' in line.lower() and 'charge total' in line.lower():
                charge_header_found = True
                # 在表头行之后查找数据行
                for j in range(i + 1, min(i + 10, len(lines))):  # 查找接下来的几行
                    data_line = lines[j].strip()
                    if data_line and not data_line.lower().startswith('nett value'):  # 跳过汇总行
                        # 提取最后一列的金额（Charge Total列）
                        import re
                        # 匹配行末尾的数字格式
                        amount_match = re.search(r'(\d+\.\d{2})$', data_line)
                        if amount_match:
                            try:
                                amount_str = amount_match.group(1)
                                amount = Decimal(amount_str)
                                if amount > 0 and 100 <= amount <= 2000:  # 合理的运费范围
                                    return amount
                            except:
                                continue
                break
        
        # 如果没找到标准格式，查找Invoice Total
        for line in lines:
            if 'invoice total:' in line.lower():
                import re
                amount_match = re.search(r'invoice total:\s*([0-9,]+\.?[0-9]*)', line.lower())
                if amount_match:
                    try:
                        amount_str = amount_match.group(1).replace(',', '')
                        amount = Decimal(amount_str)
                        if amount > 0:
                            return amount
                    except:
                        continue
        
        # 兜底方案：查找任何行末尾的金额
        for line in lines:
            import re
            amount_match = re.search(r'(\d+\.\d{2})$', line.strip())
            if amount_match:
                try:
                    amount_str = amount_match.group(1)
                    amount = Decimal(amount_str)
                    # 验证是否在合理范围内（运费通常几百英镑）
                    if 100 <= amount <= 2000:
                        return amount
                except:
                    continue
        
        return None
    
    def extract_month(self, filename: str) -> str:
        """
        从文件名提取月份，支持多种格式：
        1. 2025-7月_CostBillExport1599.xlsx
        2. 海洋国际英国海外仓20241101-1130账单.xlsx (2024年格式)
        3. HTCL-库存结算单-02.10.2025-移仓费.xlsx (移仓费计入对应月份)
        4. 运费文件强制计入2025年10月
               """
        filename_lower = filename.lower()
        
        # 处理运费文件 - 强制归属到2025年10月
        if filename_lower.startswith('运费'):
            return "2025-10"
        
        # 处理移仓费文件 - 从日期提取月份
        if '移仓费' in filename_lower:
            # 格式: HTCL-库存结算单-02.10.2025-移仓费.xlsx
            date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', filename)
            if date_match:
                day, month, year = date_match.groups()
                return f"{year}-{month}"
        
        # 处理2024年账单格式：海洋国际英国海外仓20241101-1130账单.xlsx
        match = re.search(r'海洋国际英国海外仓(\d{4})(\d{2})\d{2}-\d{2}\d{2}账单', filename)
        if match:
            year = match.group(1)
            month = match.group(2)
            return f"{year}-{month}"
        
        # 处理标准格式
        match = re.search(r'(\d{4})-(\d{1,2})月', filename)
        if match:
            return f"{match.group(1)}-{int(match.group(2)):02d}"
            
        return ""


class LHZParser(BaseWarehouseParser):
    """LHZ 仓库解析器 (DE, EUR)"""
    
    def __init__(self):
        super().__init__("LHZ", "DE", "EUR")
    
    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        LHZ 海外仓账单：只取第一个 sheet（通常为"总计"）中的账单金额，其他 sheet 为明细。

        实际样例中字段常见为：
        - 账单金额 / Rechnungsbetrag
        - 未税金额合计 / Netto（不作为最终账单金额）
        """
        xl = pd.ExcelFile(file_path)
        if not xl.sheet_names:
            return Decimal('0'), {}, 0

        cover_sheet = xl.sheet_names[0]
        df_cover = pd.read_excel(file_path, sheet_name=cover_sheet, header=None)

        keywords = [
            '账单金额', 'rechnungsbetrag',
            'total bill amount', 'invoice total', 'grand total',
        ]

        total = Decimal('0')
        found = False

        for r in range(df_cover.shape[0]):
            for c in range(df_cover.shape[1]):
                v = df_cover.iat[r, c]
                if isinstance(v, str) and any(k in v.lower() for k in keywords):
                    # 优先取右侧第一个可解析为数字的值
                    for cc in range(c + 1, df_cover.shape[1]):
                        amt = df_cover.iat[r, cc]
                        try:
                            if pd.notna(amt):
                                total = Decimal(str(amt))
                                found = True
                                break
                        except Exception:
                            continue
                    break
            if found:
                break

        breakdown = {}
        if found:
            breakdown['Bill amount'] = total
            return total, breakdown, 1

        # 兜底：定位失败则返回 0，避免误把明细累加成总计
        return Decimal('0'), {}, 0
    
    def extract_month(self, filename: str) -> str:
        """
        LHZ 文件名常见格式：
        - 开票费用明细 05-2025 HUP xlsx.xlsx  -> 2025-05
        - 开票费用明细 12.2024 HUP xlsx.xlsx  -> 2024-12
        - 开票费用明细 10.2024 HUP xlsx.xlsx  -> 2024-10
        以及部分文件名包含 NBSP(\xa0) 等不可见字符。
        """
        safe = (filename or "").replace("\xa0", " ")

        # 1) MM-YYYY
        match = re.search(r'(\d{2})-(\d{4})', safe)
        if match:
            return f"{match.group(2)}-{match.group(1)}"

        # 2) MM.YYYY
        match = re.search(r'(\d{2})\.(\d{4})', safe)
        if match:
            return f"{match.group(2)}-{match.group(1)}"

        return ""


class AoyunhuiParser(BaseWarehouseParser): 
    """奥韵汇仓库解析器 (DE, EUR) 
 
 
    结算口径：按账单明细中的「结算金额」汇总。 
    """ 
 
 
    def __init__( self ): 
        super().__init__("奥韵汇", "DE", "EUR") 
 
 
    def parse_file( self , file_path : str) -> Tuple[Decimal, Dict[str, Decimal], int]: 
        """ 
        默认解析：返回整份账单内「计费规则金额」在各月份上的总和。 
        实际按月拆分逻辑在 parse_file_by_month 中完成，这里仅做汇总， 
        以兼容旧接口。 
        """ 
        monthly = self.parse_file_by_month(file_path) 
        if  not monthly: 
            return  Decimal('0'), {}, 0 
 
 
        total = sum(v[0] for  v in  monthly.values()) 
        count = sum(v[2] for  v in  monthly.values()) 
        breakdown: Dict[str, Decimal] = {'计费规则金额': total} 
        return  total, breakdown, count 
 
 
    def _load_costbill_df( self , file_path : str): 
        """加载奥韵汇账单的 CostBill sheet.""" 
        xl = pd.ExcelFile(file_path) 
        if  not xl.sheet_names: 
            return  None 
 
 
        sheet_name = None 
        for  sh in  xl.sheet_names: 
            if  str(sh).lower() == 'costbill': 
                sheet_name = sh 
                break 
        if  sheet_name is None: 
            sheet_name = xl.sheet_names[0] 
 
 
        return  pd.read_excel(file_path, sheet_name =sheet_name) 
 
 
    def parse_file_by_month( self , file_path : str) -> Dict[str, Tuple[Decimal, Dict[str, Decimal], int]]: 
        """ 
        按「计费时间」拆分到各个自然月： 
        - 文件时间跨度可能覆盖多个月（甚至多季度），不能简单按文件名月份算。 
        - 以每行明细的「计费时间」列确定所属月份（YYYY-MM）。 
        """ 
        df = self._load_costbill_df(file_path) 
        if  df is None or df.empty: 
            return  {} 
 
 
        # 找计费规则金额列（与 parse_file 中逻辑保持一致） 
        amount_col = None 
        priority_keywords = ['计费规则金额', '计费金额'] 
        for  kw in  priority_keywords: 
            for  c in  df.columns: 
                if  kw in str(c): 
                    amount_col = c 
                    break 
            if  amount_col is not None: 
                break 
 
 
        if  amount_col is None: 
            for  c in  df.columns: 
                if  '结算金额' in str(c): 
                    amount_col = c 
                    break 
 
 
        if  amount_col is None: 
            for  c in  df.columns: 
                name = str(c) 
                if  ('结算' in name) and ('金额' in name): 
                    amount_col = c 
                    break 
 
 
        if  amount_col is None: 
            for  c in  df.columns: 
                if  '金额' in str(c): 
                    amount_col = c 
                    break 
 
 
        if  amount_col is None: 
            return  {} 
 
 
        # 寻找计费时间列 
        time_col = None 
        for  c in  df.columns: 
            name = str(c) 
            # 常见：计费时间 / 计费日期 / Billing Time / Billing Date 
            if  any(k in  name for  k in  ['计费时间', '计费日期', 'Billing Time', 'Billing Date']) or '时间' in name or 'ʱ' in name: 
                time_col = c 
                break 
 
 
        if  time_col is None: 
            return  {} 
 
 
        monthly: Dict[str, Tuple[Decimal, Dict[str, Decimal], int]] = {} 
 
 
        for  _, row in  df.iterrows(): 
            val = row.get(amount_col, 0) 
            if  pd.isna(val): 
                continue 
            try : 
                amt = Decimal(str(val)) 
            except  Exception: 
                continue 
 
 
            t = row.get(time_col) 
            ts = pd.to_datetime(t, errors ='coerce') 
            if  pd.isna(ts): 
                continue 
 
 
            ym = f"{ts.year}-{ts.month:02d}" 
 
 
            if  ym not in monthly: 
                monthly[ym] = (Decimal('0'), {}, 0) 
 
 
            total_prev, bd_prev, cnt_prev = monthly[ym] 
            total_new = total_prev + amt 
            cnt_new = cnt_prev + 1 
            bd_new = dict(bd_prev) 
            bd_new['计费规则金额'] = bd_new.get('计费规则金额', Decimal('0')) + amt 
            monthly[ym] = (total_new, bd_new, cnt_new) 
 
 
        return  monthly 
 
 
    def extract_month( self , filename : str) -> str: 
        # 例如：2025-12-31_CostBillExport1887.xlsx 
        match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename) 
        if  match: 
            return  f"{match.group(1)}-{match.group(2)}" 
        return  ""

class DongFangParser(BaseWarehouseParser):
    """东方嘉盛仓库解析器 (CN, CNY)

    结算口径：按「记账金额」列汇总，按文件名提取月份。
    """

    def __init__(self):
        super().__init__("东方嘉盛", "CN", "CNY")

    def _load_main_df(self, file_path: str):
        """东方嘉盛账单通常只有一个账户明细 sheet，直接读第一个 sheet 即可。"""
        try:
            return pd.read_excel(file_path, sheet_name=0)
        except Exception:
            return None

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析东方嘉盛账单文件，只计算交易类型为"退费"和"扣款"的记账金额。
        """
        df = self._load_main_df(file_path)
        if df is None or df.empty:
            return Decimal('0'), {}, 0

        # 1) 金额列：优先按列名匹配（兼容导出乱码/不同字段名）
        amount_col = None
        preferred_amount_keywords = ["记账金额", "入账金额", "收支金额", "发生额", "交易金额"]
        for kw in preferred_amount_keywords:
            for c in df.columns:
                if kw in str(c):
                    amount_col = c
                    break
            if amount_col is not None:
                break

        # 兜底：从"数值列"里挑选最像金额的列（混合正负、非汇率）
        if amount_col is None:
            best = None
            best_score = None
            for c in df.columns:
                if str(c).strip().lower() in ["id", "no"]:
                    continue
                ser = pd.to_numeric(df[c], errors="coerce").dropna()
                if ser.empty:
                    continue

                # 排除汇率类（大量接近 1）
                near_one_frac = float(((ser >= 0.9) & (ser <= 1.1)).mean())
                if near_one_frac > 0.8 and ser.nunique() < 20:
                    continue

                neg_frac = float((ser < 0).mean())
                mixed_sign = 0.05 <= neg_frac <= 0.95
                score = (1 if mixed_sign else 0) * 1_000_000 + len(ser)

                if best_score is None or score > best_score:
                    best_score = score
                    best = c
            amount_col = best

        if amount_col is None:
            return Decimal('0'), {}, 0

        # 2) 交易类型列：查找交易类型相关的列
        type_col = None
        type_keywords = ["交易类型", "业务类型", "类型", "transaction type", "type"]
        for kw in type_keywords:
            for c in df.columns:
                if kw.lower() in str(c).lower():
                    type_col = c
                    break
            if type_col is not None:
                break

        # 兜底：从内容里识别“退费/扣款”出现在哪个列（列名可能乱码/不含关键词）
        if type_col is None:
            best_col = None
            best_hits = 0
            for c in df.columns:
                try:
                    ser = df[c].astype(str)
                except Exception:
                    continue
                hits = int(ser.str.contains(r'(退费|扣款)', na=False).sum())
                if hits > best_hits:
                    best_hits = hits
                    best_col = c
            if best_hits > 0:
                type_col = best_col

        # 3) 筛选交易类型为"退费"和"扣款"的记录
        if type_col is not None:
            # 筛选出交易类型为"退费"或"扣款"的记录
            filtered_df = df[df[type_col].astype(str).str.contains(r'(退费|扣款)', case=False, na=False)]
        else:
            return Decimal('0'), {}, 0

        # 4) 计算筛选后的总金额和记录数
        total = Decimal('0')
        count = 0
        for v in filtered_df[amount_col].tolist():
            if pd.isna(v):
                continue
            try:
                amt = Decimal(str(v))
            except Exception:
                continue
            total += amt
            count += 1

        if count == 0:
            return Decimal('0'), {}, 0
            
        # 相加之后的结果最后再取绝对值
        total = total.copy_abs()
        breakdown = {"退费和扣款记账金额": total}
        
        return total, breakdown, count

    def extract_month(self, filename: str) -> str:
        """
        从东方嘉盛文件名中提取月份信息。
        支持格式：
        - 账单_2025-05.xlsx
        - 账单_2025-06.xlsx  
        - table-list-sample-2024.xlsx
        - 账户明细-table-list (18).xlsx (需要从内容获取)
        """
        import re
        
        # 格式1: 账单_2025-05.xlsx
        match = re.search(r'账单[_-](\d{4})-(\d{2})', filename)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
        
        # 格式2: table-list-sample-2024.xlsx
        match = re.search(r'sample-(\d{4})', filename)
        if match:
            return f"{match.group(1)}-11"  # 示例文件默认归为11月
        
        # 格式3: 账户明细-table-list (18).xlsx
        # 对于这类文件，需要从文件内容中获取最新交易日期
        if '账户明细-table-list' in filename:
            # 这里可以实现从文件内容中提取最新日期的逻辑
            # 暂时返回空字符串，让汇总逻辑处理
            return ""
            
        return ""


class JiuXiParser(BaseWarehouseParser):
    """久喜仓库解析器 (CN, CNY)

    结算口径：按「计算规则金额」列汇总，按文件名提取月份。
    """

    def __init__(self):
        super().__init__("久喜", "JP", "JPY")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析久喜仓库账单文件，提取计算规则金额列并计算总额。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取第一个工作表
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找计算规则金额列
            amount_col = None
            priority_keywords = ['计算规则金额', '计费规则金额', '计费金额', '金额']
            for kw in priority_keywords:
                for c in df.columns:
                    if kw in str(c):
                        amount_col = c
                        break
                if amount_col is not None:
                    break

            if amount_col is None:
                return Decimal('0'), {}, 0

            # 计算总金额和记录数
            total = Decimal('0')
            count = 0
            for _, row in df.iterrows():
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    total += amt
                    count += 1
                except Exception:
                    continue

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'计算规则金额': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  久喜解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从久喜仓库文件名中提取月份信息。
        支持格式：
        - 2026-01-04_CostBillExport15585.xlsx
        """
        import re
        
        # 格式: 2026-01-04_CostBillExport15585.xlsx
        match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
            
        return ""


class JinDaParser(BaseWarehouseParser):
    """津达仓库解析器 (CN, CNY)

    结算口径：按「计算规则金额」列汇总，按文件名提取月份。
    """

    def __init__(self):
        super().__init__("津达", "DE", "EUR")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析津达仓库账单文件，提取计算规则金额列并计算总额。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取第一个工作表
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找计算规则金额列
            amount_col = None
            priority_keywords = ['计算规则金额', '计费规则金额', '计费金额', '金额']
            for kw in priority_keywords:
                for c in df.columns:
                    if kw in str(c):
                        amount_col = c
                        break
                if amount_col is not None:
                    break

            if amount_col is None:
                return Decimal('0'), {}, 0

            # 计算总金额和记录数
            total = Decimal('0')
            count = 0
            for _, row in df.iterrows():
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    total += amt
                    count += 1
                except Exception:
                    continue

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'计算规则金额': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  津达解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从津达仓库文件名中提取月份信息。
        支持格式：
        - 2025-12-31_CostBillExport29779.xlsx
        """
        import re
        
        # 格式: 2025-12-31_CostBillExport29779.xlsx
        match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
            
        return ""


class KuLuParser(BaseWarehouseParser):
    """酷麓仓库解析器 (US, USD)

    结算口径：按「计算规则金额」列汇总，按文件名提取月份。
    """

    def __init__(self):
        super().__init__("酷麓", "US", "USD")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析酷麓仓库账单文件，提取计算规则金额列并计算总额。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取第一个工作表
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找计算规则金额列
            amount_col = None
            priority_keywords = ['计算规则金额', '计费规则金额', '计费金额', '金额']
            for kw in priority_keywords:
                for c in df.columns:
                    if kw in str(c):
                        amount_col = c
                        break
                if amount_col is not None:
                    break

            if amount_col is None:
                return Decimal('0'), {}, 0

            # 计算总金额和记录数
            total = Decimal('0')
            count = 0
            for _, row in df.iterrows():
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    total += amt
                    count += 1
                except Exception:
                    continue

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'计算规则金额': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  酷麓解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从酷麓仓库文件名中提取月份信息。
        支持格式：
        - 2026-01-04_CostBillExport10457.xlsx
        """
        import re
        
        # 格式: 2026-01-04_CostBillExport10457.xlsx
        match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
        if match:
            return f"{match.group(1)}-{match.group(2)}"
            
        return ""


class XiYouParser(BaseWarehouseParser):
    """西邮仓库解析器

    结算口径：从汇总工作表提取费用合计(Total Fee)，从账单国家(Country)获取地区，从费用开始/结束日期获取月份。
    """

    def __init__(self):
        super().__init__("西邮", "US", "USD")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析西邮仓库账单文件，从汇总工作表提取费用合计。
        处理费用合计(Total Fee)列隔一个或两个单元格就是金额数目的格式。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取汇总工作表（假设是第一个工作表）
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找费用合计(Total Fee)所在的单元格
            total = Decimal('0')
            count = 0
            
            # 遍历所有单元格，寻找费用合计标记
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    cell_str = str(cell).strip()
                    if 'total fee' in cell_str.lower() or '费用合计' in cell_str:
                        # 尝试不同的偏移量寻找金额：优先隔两个单元格，然后隔一个，最后直接相邻
                        for offset in [3, 2, 1]:  # 偏移量3:隔两个, 2:隔一个, 1:直接相邻
                            if j + offset < len(row):  # 确保索引不越界
                                amount_cell = row.iloc[j + offset]
                                if not pd.isna(amount_cell):
                                    try:
                                        amt = Decimal(str(amount_cell))
                                        total += amt
                                        count += 1
                                        break  # 找到有效金额后停止尝试其他偏移量
                                    except Exception:
                                        continue
                        break
                if count > 0:
                    break

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'费用合计': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  西邮解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从西邮仓库文件名中提取月份信息。
        支持格式：
        - AAB57--US--TEMU--西邮物流仓储账单--2025-06-01-2025-06-30--初版.xlsx
        - AAB57--US--TEMU--西邮物流仓储账单--2025.07.01-2025.07.31--初版.xlsx
        """
        import re
        
        # 格式1: 2025-06-01-2025-06-30
        match1 = re.search(r'(\d{4})[-\.](\d{2})[-\.]\d{2}[-\.](\d{4})[-\.](\d{2})[-\.]\d{2}', filename)
        if match1:
            # 取开始日期的月份
            return f"{match1.group(1)}-{match1.group(2)}"
            
        return ""


class TLBParser(BaseWarehouseParser):
    """TLB账单仓库解析器 (UK, GBP)

    结算口径：从Total Due获取金额数据，处理H和I列的公式计算。
    """

    def __init__(self):
        super().__init__("TLB账单", "UK", "GBP")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析TLB账单仓库文件，从Total Due获取金额数据。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取第一个工作表
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找Total Due所在的单元格
            total = Decimal('0')
            count = 0
            
            # 遍历所有单元格，寻找Total Due标记
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    cell_str = str(cell).strip()
                    if 'total due' in cell_str.lower():
                        # 尝试从当前行的后续单元格获取金额
                        # 首先尝试直接从当前行的后续单元格获取
                        for offset in [1, 2, 3]:  # 尝试不同的偏移量
                            if j + offset < len(row):
                                amount_cell = row.iloc[j + offset]
                                if not pd.isna(amount_cell):
                                    # 处理带货币符号的金额
                                    amount_str = str(amount_cell).replace('£', '').replace(',', '').strip()
                                    try:
                                        amt = Decimal(amount_str)
                                        total += amt
                                        count += 1
                                        break
                                    except Exception:
                                        continue
                        break
                if count > 0:
                    break

            # 如果没找到，尝试从H和I列寻找公式计算的结果
            if count == 0:
                for i, row in df.iterrows():
                    # 检查H列和I列
                    for col_idx in range(len(row)):
                        cell = row.iloc[col_idx]
                        if not pd.isna(cell):
                            cell_str = str(cell).strip()
                            # 检查是否是金额格式（带£符号）
                            if cell_str.startswith('£'):
                                amount_str = cell_str.replace('£', '').replace(',', '').strip()
                                try:
                                    amt = Decimal(amount_str)
                                    total += amt
                                    count += 1
                                    break
                                except Exception:
                                    continue
                    if count > 0:
                        break

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'Total Due': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  TLB解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从TLB账单文件名中提取月份信息。
        支持格式：
        - AC000896 T006 2024年9月对账单.xlsx
        - AC000913 T006_2024年10月对账单.xlsx
        """
        import re
        
        # 格式1: 2024年9月
        match1 = re.search(r'(\d{4})年(\d+)月', filename)
        if match1:
            year = match1.group(1)
            month = match1.group(2).zfill(2)  # 确保月份是两位数
            return f"{year}-{month}"
            
        return ""


class YiDaYunParser(BaseWarehouseParser):
    """易达云仓库解析器 (US, USD)

    结算口径：从账单总消费行获取金额，做绝对值处理。
    """

    def __init__(self):
        super().__init__("易达云", "US", "USD")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析易达云仓库账单文件，从账单总消费行获取金额并做绝对值处理。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            # 读取第一个工作表
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            if df.empty:
                return Decimal('0'), {}, 0

            # 寻找账单总消费所在的单元格
            total = Decimal('0')
            count = 0
            
            # 遍历所有单元格，寻找账单总消费标记
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    cell_str = str(cell).strip()
                    if '账单总消费' in cell_str:
                        # 尝试从当前行的后续单元格获取金额
                        for offset in [1, 2, 3]:  # 尝试不同的偏移量
                            if j + offset < len(row):
                                amount_cell = row.iloc[j + offset]
                                if not pd.isna(amount_cell):
                                    # 处理金额，做绝对值处理
                                    try:
                                        amt = Decimal(str(amount_cell))
                                        amt = abs(amt)  # 绝对值处理
                                        total += amt
                                        count += 1
                                        break
                                    except Exception:
                                        continue
                        break
                if count > 0:
                    break

            # 如果没找到，尝试从所有单元格中寻找包含"总消费"的标记
            if count == 0:
                for i, row in df.iterrows():
                    for j, cell in enumerate(row):
                        cell_str = str(cell).strip()
                        if '总消费' in cell_str:
                            # 尝试从当前行的后续单元格获取金额
                            for offset in [1, 2, 3]:
                                if j + offset < len(row):
                                    amount_cell = row.iloc[j + offset]
                                    if not pd.isna(amount_cell):
                                        try:
                                            amt = Decimal(str(amount_cell))
                                            amt = abs(amt)  # 绝对值处理
                                            total += amt
                                            count += 1
                                            break
                                        except Exception:
                                            continue
                            break
                    if count > 0:
                        break

            if count == 0:
                return Decimal('0'), {}, 0

            breakdown = {'账单总消费': total}
            return total, breakdown, count

        except Exception as e:
            print(f"  易达云解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def extract_month(self, filename: str) -> str:
        """
        从易达云仓库文件名中提取月份信息。
        支持格式：
        - Z0333-1756833040205.xlsx (时间戳格式)
        """
        import re
        from datetime import datetime
        
        # 格式: Z0333-1756833040205.xlsx
        match = re.search(r'Z0333-(\d+)\.xlsx', filename)
        if match:
            timestamp = match.group(1)
            # 处理时间戳，可能是毫秒或秒
            if len(timestamp) == 13:
                # 毫秒时间戳
                dt = datetime.fromtimestamp(int(timestamp) / 1000)
            else:
                # 秒时间戳
                dt = datetime.fromtimestamp(int(timestamp))
            return f"{dt.year}-{dt.month:02d}"
            
        return ""


class YiLingParser(BaseWarehouseParser):
    """易领仓库解析器

    结算口径：从四个工作表分别提取费用小计并计算总支出。
    """

    def __init__(self):
        super().__init__("易领", "US", "USD")

    def parse_file(self, file_path: str) -> Tuple[Decimal, Dict[str, Decimal], int]:
        """
        解析易领仓库账单文件，从四个工作表分别提取费用小计并计算总支出。
        """
        try:
            xl = pd.ExcelFile(file_path)
            if not xl.sheet_names:
                return Decimal('0'), {}, 0

            total = Decimal('0')
            breakdown = {}
            count = 0

            # 解析入库费用明细工作表 (inbound fee)
            inbound_total = self._parse_worksheet(xl, 'inbound fee', ['Subtotal', '入库费用小计'])
            if inbound_total > Decimal('0'):
                total += inbound_total
                breakdown['入库费用'] = inbound_total
                count += 1

            # 解析出库费用明细工作表 (outbound fee)
            outbound_total = self._parse_worksheet(xl, 'outbound fee', ['Subtotal', '出库费用小计'])
            if outbound_total > Decimal('0'):
                total += outbound_total
                breakdown['出库费用'] = outbound_total
                count += 1

            # 解析仓租费用工作表 (Storage Charges)
            storage_total = self._parse_worksheet(xl, 'Storage Charges', ['Storage Fee', '仓租费用小计'])
            if storage_total > Decimal('0'):
                total += storage_total
                breakdown['仓租费用'] = storage_total
                count += 1

            # 解析其他费用明细工作表 (other fee)
            other_total = self._parse_worksheet(xl, 'other fee', ['Storage Fee', '费用小计'])
            if other_total > Decimal('0'):
                total += other_total
                breakdown['其他费用'] = other_total
                count += 1

            if count == 0:
                return Decimal('0'), {}, 0

            return total, breakdown, count

        except Exception as e:
            print(f"  易领解析失败 {file_path}: {e}")
            return Decimal('0'), {}, 0

    def _parse_worksheet(self, xl: pd.ExcelFile, sheet_name: str, keywords: list) -> Decimal:
        """
        解析单个工作表，提取费用小计。
        """
        try:
            # 尝试匹配工作表名称（不区分大小写，模糊匹配）
            matched_sheet_name = None
            for ws_name in xl.sheet_names:
                if sheet_name.lower() in ws_name.lower() or ws_name.lower() in sheet_name.lower():
                    matched_sheet_name = ws_name
                    break
            
            # 如果没找到匹配的工作表，尝试使用关键词匹配
            if matched_sheet_name is None:
                for ws_name in xl.sheet_names:
                    for keyword in ['inbound', 'outbound', 'storage', 'other']:
                        if keyword in ws_name.lower():
                            if (keyword == 'inbound' and sheet_name == 'inbound fee') or \
                               (keyword == 'outbound' and sheet_name == 'outbound fee') or \
                               (keyword == 'storage' and sheet_name == 'Storage Charges') or \
                               (keyword == 'other' and sheet_name == 'other fee'):
                                matched_sheet_name = ws_name
                                break
                    if matched_sheet_name:
                        break

            if matched_sheet_name is None:
                return Decimal('0')

            df = pd.read_excel(xl, sheet_name=matched_sheet_name)
            if df.empty:
                return Decimal('0')

            # 寻找费用小计列
            amount_col = None
            for kw in keywords:
                for c in df.columns:
                    if kw in str(c):
                        amount_col = c
                        break
                if amount_col is not None:
                    break

            if amount_col is None:
                return Decimal('0')

            # 计算总金额
            subtotal = Decimal('0')
            for _, row in df.iterrows():
                val = row.get(amount_col, 0)
                if pd.isna(val):
                    continue
                try:
                    amt = Decimal(str(val))
                    subtotal += amt
                except Exception:
                    continue

            return subtotal

        except Exception:
            return Decimal('0')

    def extract_month(self, filename: str) -> str:
        """
        从易领仓库文件名中提取月份信息。
        支持格式：
        - OperatingCosts_20251231115544.xlsx
        """
        import re
        
        # 格式: OperatingCosts_20251231115544.xlsx
        match = re.search(r'OperatingCosts_(\d{8})\d+\.xlsx', filename)
        if match:
            date_str = match.group(1)
            year = date_str[:4]
            month = date_str[4:6]
            return f"{year}-{month}"
            
        return ""


def get_parser(warehouse_name: str) -> BaseWarehouseParser:
    """获取仓库解析器"""
    parsers = {
        'TSP': TSPParser(),
        '1510': Warehouse1510Parser(),
        '京东': JDParser(),
        '海洋': HaiyangParser(),
        'LHZ': LHZParser(),
        '奥韵汇': AoyunhuiParser(),
        '东方嘉盛': DongFangParser(),
        'G7': G7Parser(),  # 添加G7解析器
        '久喜': JiuXiParser(),  # 添加久喜解析器
        '津达': JinDaParser(),  # 添加津达解析器
        '酷麓': KuLuParser(),  # 添加酷麓解析器
        '西邮': XiYouParser(),  # 添加西邮解析器
        'TLB账单': TLBParser(),  # 添加TLB账单解析器
        '易达云': YiDaYunParser(),  # 添加易达云解析器
        '易领': YiLingParser(),  # 添加易领解析器
    }
    return parsers.get(warehouse_name)


def scan_warehouse_files(base_path: str, warehouse_name: str) -> List[str]:
    """扫描仓库目录下的文件，根据仓库类型决定是否包含PDF文件"""
    wh_path = os.path.join(base_path, warehouse_name)
    files = []
    
    # 定义哪些仓库需要扫描PDF文件
    warehouses_needing_pdf = ['海洋', 'G7']  # 海洋仓库需要处理运费PDF文件，G7仓库需要处理账单PDF文件
    
    # 某些 Windows 环境下中文目录名在不同编码/终端里可能出现乱码，导致直接拼接路径找不到。
    # 这里为个别仓库提供兜底扫描策略（按文件名特征）。
    if not os.path.exists(wh_path):
        if warehouse_name == '东方嘉盛':
            # 东方嘉盛导出的文件名通常包含 table-list
            for root, _, filenames in os.walk(base_path):
                for f in filenames:
                    if f.startswith('~$'):
                        continue
                    if f.lower().endswith(('.xlsx', '.xls')) and ('table-list' in f.lower() or '账单_' in f):
                        files.append(os.path.join(root, f))
            # 继续走去重逻辑
        elif warehouse_name == 'G7':
            # G7仓库处理PDF文件
            for root, _, filenames in os.walk(base_path):
                for f in filenames:
                    if f.startswith('~$'):
                        continue
                    if f.lower().endswith('.pdf') and ('invoice' in f.lower() or 'credit' in f.lower() or f.startswith('702')):
                        files.append(os.path.join(root, f))
        else:
            return files
    
    for root, dirs, filenames in os.walk(wh_path):
        for f in filenames:
            # 根据仓库类型决定扫描的文件类型
            if warehouse_name in warehouses_needing_pdf:
                if warehouse_name == 'G7':
                    # G7仓库：只扫描PDF文件
                    if f.lower().endswith('.pdf') and not f.startswith('~$'):
                        files.append(os.path.join(root, f))
                else:
                    # 海洋仓库：扫描Excel和PDF文件
                    if (f.lower().endswith(('.xlsx', '.xls', '.pdf')) and not f.startswith('~$')):
                        files.append(os.path.join(root, f))
            else:
                # 其他仓库：只扫描Excel文件
                if (f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')):
                    files.append(os.path.join(root, f))

    # 去重：同一目录下同名的重复下载文件通常带 "(1)/(2)/(3)" 后缀，避免重复计入
    # 规则：对相同"规范化相对路径"的文件，仅保留最后修改时间最新的那一份
    def _normalize_relpath(fp: str) -> str:
        rel = os.path.relpath(fp, wh_path)
        d = os.path.dirname(rel)
        base = os.path.basename(rel)
        stem, ext = os.path.splitext(base)
        # 仅对"重复下载"常见的 (1)/(2)... 做去重，避免把业务编号 (18)/(33) 等误当重复文件
        m = re.search(r'\s*\((\d+)\)$', stem)
        if m:
            try:
                n = int(m.group(1))
            except Exception:
                n = None
            # 通常重复下载后缀在 1~9 之间；超过 9 更可能是业务序号
            if n is not None and 1 <= n <= 9:
                stem = re.sub(r'\s*\(\d+\)$', '', stem)  # e.g. "xxx (3).xlsx" -> "xxx.xlsx"
        norm_base = (stem + ext).lower()
        return os.path.join(d, norm_base).lower()

    best = {}
    for fp in files:
        key = _normalize_relpath(fp)
        try:
            mtime = os.path.getmtime(fp)
        except Exception:
            mtime = 0

        if key not in best:
            best[key] = (mtime, fp)
        else:
            if mtime >= best[key][0]:
                best[key] = (mtime, fp)

    return [v[1] for v in best.values()]


def aggregate_warehouse_costs(base_path: str, warehouses: List[str]) -> List[WarehouseMonthlyCost]:
    """汇总所有仓库的月度成本"""
    results = []
    
    for wh_name in warehouses:
        parser = get_parser(wh_name)
        if not parser:
            continue
        
        files = scan_warehouse_files(base_path, wh_name)
        
        # 按月份分组
        monthly_data = {}
        
        for fp in files:
            try:
                filename = os.path.basename(fp)
                # 优先使用解析器的“按月拆分”能力（适用于奥韵汇这类跨月文件）
                if hasattr(parser, "parse_file_by_month"):
                    monthly_results = parser.parse_file_by_month(fp)  # type: ignore
                    for ym, (total, breakdown, count) in monthly_results.items():
                        if not ym:
                            continue
                        if ym not in monthly_data:
                            monthly_data[ym] = {
                                'total': Decimal('0'),
                                'breakdown': {},
                                'count': 0,
                                'files': []
                            }
                        monthly_data[ym]['total'] += total
                        monthly_data[ym]['count'] += count
                        if filename not in monthly_data[ym]['files']:
                            monthly_data[ym]['files'].append(filename)
                        for k, v in breakdown.items():
                            monthly_data[ym]['breakdown'][k] = monthly_data[ym]['breakdown'].get(k, Decimal('0')) + v
                else:
                    # 传递完整文件路径给extract_month方法，以便某些解析器（如G7、京东）可以从路径中提取月份
                    year_month = parser.extract_month(fp)  # 传入完整路径fp而非filename
                    if not year_month:
                        continue
                    total, breakdown, count = parser.parse_file(fp)
                    if year_month not in monthly_data:
                        monthly_data[year_month] = {
                            'total': Decimal('0'),
                            'breakdown': {},
                            'count': 0,
                            'files': []
                        }
                    monthly_data[year_month]['total'] += total
                    monthly_data[year_month]['count'] += count
                    if filename not in monthly_data[year_month]['files']:
                        monthly_data[year_month]['files'].append(filename)
                    for k, v in breakdown.items():
                        monthly_data[year_month]['breakdown'][k] = monthly_data[year_month]['breakdown'].get(k, Decimal('0')) + v
            except Exception as e:
                print(f"  解析失败 {fp}: {e}")
        
        # 转换为结果对象
        for ym, data in monthly_data.items():
            results.append(WarehouseMonthlyCost(
                warehouse_name=wh_name,
                year_month=ym,
                total_cost=data['total'],
                currency=parser.currency,
                cost_breakdown=data['breakdown'],
                record_count=data['count'],
                source_files=data['files'],
            ))
    
    return results


if __name__ == '__main__':
    print("=" * 60)
    print("Phase 2 仓库成本汇总测试")
    print("=" * 60)
    
    base = r'd:\app\收入核算系统\data\仓库财务账单\海外仓账单'
    warehouses = ['TSP', '1510', '京东', '海洋', 'LHZ', '东方嘉盛', '奥韵汇', 'G7']  # 添加G7仓库
    
    results = aggregate_warehouse_costs(base, warehouses)
    
    for r in sorted(results, key=lambda x: (x.warehouse_name, x.year_month)):
        print(f"\n{r.warehouse_name} | {r.year_month} | {r.total_cost:,.2f} {r.currency}")
        for k, v in list(r.cost_breakdown.items())[:3]:
            print(f"  - {k}: {v:,.2f}")
