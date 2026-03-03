# -*- coding: utf-8 -*-
"""
Phase 2: 平台收入 × 仓库履约成本

阶段边界：
- 平台收入汇总（Phase 1）
- 仓库履约成本汇总（本阶段）
- 不做 SKU 级成本
- 不做订单级匹配
"""
from decimal import Decimal
from pathlib import Path
import sys
import warnings

import pandas as pd

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from src.interfaces import FixedExchangeRate
from src.parser.shanghai_huopan_parser import aggregate_shanghai_huopan_costs
from src.parser.warehouse_parser import aggregate_warehouse_costs

warnings.filterwarnings("ignore")


WAREHOUSE_REGION_MAP = {
    "TSP": "UK",
    "1510": "UK",
    "京东": "Global",
    "海洋": "UK",
    "LHZ": "DE",
    "奥韵汇": "DE",
    "东方嘉盛": "CN",
    "G7": "DE",
    "久喜": "DE",
    "津达": "DE",
    "酷鹿": "US",
    "西邮": "US",
    "TLB账单": "UK",
    "易达云": "US",
    "易领": "US",
    "AUS_FDM": "AU",
    "澳洲FDM": "AU",
    "澳洲ADM": "AU",
    "sphere freight": "AU",
    "Sphere Freight": "AU",
    "中转仓": "AU",
    "上海货盘": "CN",
}


def get_warehouse_region(warehouse_name: str) -> str:
    """Resolve warehouse region with simple alias fallback."""
    name = (warehouse_name or "").strip().split("|", 1)[0]
    if name in WAREHOUSE_REGION_MAP:
        return WAREHOUSE_REGION_MAP[name]

    upper_name = name.upper()
    if upper_name in WAREHOUSE_REGION_MAP:
        return WAREHOUSE_REGION_MAP[upper_name]

    return "-"


def split_warehouse_and_vendor(warehouse_name: str):
    """拆分仓库名与物流公司，格式：上海货盘|GAO。"""
    raw = (warehouse_name or "").strip()
    if "|" not in raw:
        return raw, ""
    base, vendor = raw.split("|", 1)
    return base.strip(), vendor.strip()


def build_shanghai_monthly_fee_summary(detail_rows):
    """按月份+物流公司汇总上海货盘费用科目。"""
    if not detail_rows:
        return pd.DataFrame()

    df = pd.DataFrame(detail_rows)
    month_col = "月份"
    company_col = "物流公司"
    subject_col = "费用科目"
    amount_col = "Amount"
    status_col = "状态"
    country_col = "国家"
    parse_mode_col = "解析方式"
    currency_col = "币种"
    fx_rate_col = "汇率"

    required_cols = {month_col, subject_col, amount_col, status_col}
    if not required_cols.issubset(df.columns):
        return pd.DataFrame()

    if company_col not in df.columns:
        df[company_col] = "-"

    base_fee_cols = [
        "拆柜费",
        "港口短驳拖车费",
        "尾程派送费",
        "头程海运费",
        "卸货费/入库卸柜费",
        "仓储费",
        "仓库操作费",
        "贴标费",
    ]
    dit_fee_cols = [
        "关税",
        "清关费",
        "ISF申报费",
        "尾程操作费",
        "FDA申报费",
        "LACEY申报费",
        "尾程综合费(未分项)",
    ]
    fee_cols = base_fee_cols + dit_fee_cols

    df = df[df[status_col] == "parsed"].copy()
    if df.empty:
        return pd.DataFrame()

    if parse_mode_col not in df.columns:
        df[parse_mode_col] = ""
    if country_col not in df.columns:
        df[country_col] = ""
    if currency_col not in df.columns:
        df[currency_col] = "USD"
    if fx_rate_col not in df.columns:
        df[fx_rate_col] = None

    df[month_col] = df[month_col].astype(str).str[:7]
    df = df[df[month_col].str.match(r"^\d{4}-\d{2}$", na=False)]
    df = df[df[subject_col].isin(fee_cols)]
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce")
    df = df.dropna(subset=[amount_col])
    if df.empty:
        return pd.DataFrame()

    # 头程海运费目录按「物流公司+国家」维度聚合；为兼容现有表头，不新增列，编码到物流公司值。
    is_headsea = df[parse_mode_col].astype(str).str.contains("headsea", case=False, na=False)
    country_series = df[country_col].fillna("").astype(str).str.strip()
    country_series = country_series.where(country_series != "", "UNKNOWN")
    df.loc[is_headsea, company_col] = (
        df.loc[is_headsea, company_col].fillna("").astype(str).str.strip().replace("", "-")
        + "|"
        + country_series.loc[is_headsea]
    )

    # 当前汇总表只有单金额列，按固定汇率统一折算为 CNY，避免多币种直接相加。
    rate_provider = FixedExchangeRate()

    def _parse_rate_value(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        try:
            dv = Decimal(str(v).strip())
        except Exception:
            return None
        if dv <= 0:
            return None
        return dv

    def _to_cny_amount(row) -> float:
        amt = float(row[amount_col])
        curr = str(row.get(currency_col, "") or "").upper().strip()
        if not curr or curr == "CNY":
            return amt
        try:
            invoice_rate = _parse_rate_value(row.get(fx_rate_col))
            if invoice_rate is not None:
                return float(Decimal(str(amt)) * invoice_rate)
            rate = rate_provider.get_rate(curr, "CNY")
            return float(Decimal(str(amt)) * rate)
        except Exception:
            return amt

    df[amount_col] = df.apply(_to_cny_amount, axis=1)

    pivot = (
        df.pivot_table(
            index=[month_col, company_col],
            columns=subject_col,
            values=amount_col,
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    for col in fee_cols:
        if col not in pivot.columns:
            pivot[col] = 0.0

    ordered_cols = [month_col, company_col] + fee_cols
    pivot = pivot[ordered_cols]
    pivot["合计"] = pivot[fee_cols].sum(axis=1)
    return pivot.sort_values([month_col, company_col]).reset_index(drop=True)



def run_phase2():
    """Phase 2 主入口"""
    print("=" * 70)
    print(" Phase 2: 平台收入 × 仓库履约成本")
    print(" 限制: 不含 SKU 级成本，不做订单级匹配")
    print("=" * 70)

    # 路径配置
    warehouse_data_path = Path(r"C:\Users\EDY\Desktop\CB-Settlement\data\仓库财务账单\海外仓账单")
    au_warehouse_data_path = Path(r"C:\Users\EDY\Desktop\CB-Settlement\data\仓库财务账单\澳洲")
    shanghai_huopan_path = Path(r"C:\Users\EDY\Desktop\CB-Settlement\data\仓库财务账单\上海货盘费用账单")
    output_path = Path(r"C:\Users\EDY\Desktop\CB-Settlement\output")

    # === 1. 平台收入汇总 (沿用 Phase 1 结果) ===
    print("\n[1] 加载平台收入数据...")

    possible_reports = [
        output_path / "月度核算报表_Phase1_多平台.xlsx",
        output_path / "月度核算报表_Phase1.xlsx",
        output_path / "多平台核算报表.xlsx",
    ]

    phase1_report = None
    for rp in possible_reports:
        if rp.exists():
            phase1_report = rp
            break

    platform_revenue = []
    if phase1_report:
        try:
            df = pd.read_excel(phase1_report)
            print(f"  从 Phase 1 报表加载 {len(df)} 条记录")
            platform_revenue = df.to_dict("records")
        except Exception as e:
            print(f"  加载失败: {e}")
    else:
        print(f"  Phase 1 报表不存在: {phase1_report}")
        print("  请先运行 run_phase1_multiplatform.py")

    # === 2. 仓库履约成本汇总 ===
    print("\n[2] 汇总仓库履约成本...")

    warehouses = [
        "TSP",
        "1510",
        "京东",
        "海洋",
        "LHZ",
        "奥韵汇",
        "东方嘉盛",
        "G7",
        "久喜",
        "津达",
        "酷鹿",
        "西邮",
        "TLB账单",
        "易达云",
        "易领",
    ]
    warehouse_costs = aggregate_warehouse_costs(str(warehouse_data_path), warehouses)
    warehouse_costs.extend(aggregate_warehouse_costs(str(au_warehouse_data_path), ["AUS_FDM", "sphere freight"]))

    # 新增：上海货盘费用账单
    shanghai_costs, shanghai_detail_rows = aggregate_shanghai_huopan_costs(str(shanghai_huopan_path))
    warehouse_costs.extend(shanghai_costs)

    print(f"  共解析 {len(warehouse_costs)} 条仓库月度记录")
    print(f"  上海货盘月度记录: {len(shanghai_costs)} 条")
    print(f"  上海货盘明细记录: {len(shanghai_detail_rows)} 条")

    # 按仓库统计
    wh_summary = {}
    for c in warehouse_costs:
        wh_base, _ = split_warehouse_and_vendor(c.warehouse_name)
        wh_key = wh_base if wh_base else c.warehouse_name
        if wh_key not in wh_summary:
            wh_summary[wh_key] = Decimal("0")
        wh_summary[wh_key] += c.total_cost

    for wh, total in wh_summary.items():
        print(f"    {wh}: {total:,.2f}")

    # === 3. 生成 Phase 2 报表 ===
    print("\n[3] 生成 Phase 2 报表...")
    output_file = output_path / "月度核算报表_Phase2.xlsx"

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # Sheet 1: 平台收入汇总
            if platform_revenue:
                df_platform = pd.DataFrame(platform_revenue)
                df_platform.to_excel(writer, sheet_name="平台收入汇总", index=False)
                print(f"  - 平台收入汇总: {len(df_platform)} 行")

            # Sheet 2: 仓库成本汇总
            warehouse_rollup = {}
            for c in warehouse_costs:
                warehouse_base, _ = split_warehouse_and_vendor(c.warehouse_name)
                wh_name = warehouse_base if warehouse_base else c.warehouse_name
                key = (c.year_month, wh_name, c.currency)
                if key not in warehouse_rollup:
                    warehouse_rollup[key] = {
                        "total": Decimal("0"),
                        "record_count": 0,
                        "files": set(),
                    }
                warehouse_rollup[key]["total"] += c.total_cost
                warehouse_rollup[key]["record_count"] += c.record_count
                warehouse_rollup[key]["files"].update(c.source_files)

            warehouse_rows = []
            for key in sorted(warehouse_rollup.keys()):
                year_month, wh_name, curr = key
                agg = warehouse_rollup[key]
                warehouse_rows.append(
                    {
                        "月份": year_month,
                        "仓库": wh_name,
                        "区域": get_warehouse_region(wh_name),
                        "履约成本合计": float(agg["total"]),
                        "币种": curr,
                        "记录数": agg["record_count"],
                        "文件数": len(agg["files"]),
                    }
                )

            df_warehouse = pd.DataFrame(warehouse_rows)
            df_warehouse.to_excel(writer, sheet_name="仓库成本汇总", index=False)
            print(f"  - 仓库成本汇总: {len(df_warehouse)} 行")

            # Sheet 3: 综合损益概览（不做汇率折算，按原币种核算）
            monthly_summary = {}

            def _norm_currency(curr):
                c = str(curr or "").upper().strip()
                return c if c and c not in {"NAN", "NONE", "-"} else "UNKNOWN"

            def _safe_decimal(value):
                try:
                    return Decimal(str(value))
                except Exception:
                    return Decimal("0")

            # 平台收入按月+币种汇总
            if platform_revenue:
                for row in platform_revenue:
                    month_val = str(row.get("月份", ""))[:7] if "月份" in row else None
                    if not month_val:
                        continue

                    currency_val = _norm_currency(row.get("币种", ""))
                    key = (month_val, currency_val)
                    if key not in monthly_summary:
                        monthly_summary[key] = {"收入": Decimal("0"), "成本": Decimal("0")}

                    revenue_val = row.get("平台净结算", 0)
                    if revenue_val and not pd.isna(revenue_val):
                        monthly_summary[key]["收入"] += _safe_decimal(revenue_val)

            # 仓库成本按月+币种汇总
            for c in warehouse_costs:
                currency_val = _norm_currency(c.currency)
                key = (c.year_month, currency_val)
                if key not in monthly_summary:
                    monthly_summary[key] = {"收入": Decimal("0"), "成本": Decimal("0")}
                monthly_summary[key]["成本"] += _safe_decimal(c.total_cost)

            summary_rows = []
            for (month, currency_val) in sorted(monthly_summary.keys()):
                data = monthly_summary[(month, currency_val)]
                revenue = data["收入"]
                cost = data["成本"]
                profit = revenue - cost

                if revenue == 0 and cost > 0:
                    remark = "⚠️ 该月该币种无平台收入数据"
                elif cost == 0 and revenue > 0:
                    remark = "⚠️ 该月该币种无仓库成本数据"
                elif revenue < 0:
                    remark = "⚠️ 该月该币种平台收入为负(退款/调整)"
                elif cost > 0 and revenue > 0 and cost > revenue * 10:
                    remark = "⚠️ 该币种成本远大于收入,数据可能不完整"
                else:
                    remark = "不含SKU采购成本（原币种）"

                summary_rows.append(
                    {
                        "月份": month,
                        "币种": currency_val,
                        "平台总收入": float(revenue),
                        "仓库总成本": float(cost),
                        "毛利(不含商品成本)": float(profit),
                        "备注": remark,
                    }
                )

            df_summary = pd.DataFrame(
                summary_rows,
                columns=["月份", "币种", "平台总收入", "仓库总成本", "毛利(不含商品成本)", "备注"],
            )
            df_summary.to_excel(writer, sheet_name="综合损益概览", index=False)
            print(f"  - 综合损益概览: {len(df_summary)} 行")

            # Sheet 4: 上海货盘仓库成本汇总
            shanghai_rows = []
            for c in sorted(shanghai_costs, key=lambda x: (x.year_month, x.warehouse_name)):
                warehouse_base, logistics_vendor = split_warehouse_and_vendor(c.warehouse_name)
                wh_name = warehouse_base if warehouse_base else "上海货盘"
                region = "UK" if (logistics_vendor or "").strip().upper() == "SARR" else get_warehouse_region(wh_name)
                shanghai_rows.append(
                    {
                        "月份": c.year_month,
                        "仓库": wh_name,
                        "物流公司": logistics_vendor if logistics_vendor else "-",
                        "区域": region,
                        "履约成本合计": float(c.total_cost),
                        "币种": c.currency,
                        "记录数": c.record_count,
                        "文件数": len(c.source_files),
                    }
                )

            if shanghai_rows:
                df_shanghai_cost = pd.DataFrame(shanghai_rows)
                df_shanghai_cost.to_excel(writer, sheet_name="上海货盘仓库成本汇总", index=False)
                print(f"  - 上海货盘仓库成本汇总: {len(df_shanghai_cost)} 行")

            # Sheet 5: 上海货盘费用类型月汇总表
            if shanghai_detail_rows:
                df_sh_monthly = build_shanghai_monthly_fee_summary(shanghai_detail_rows)
                if not df_sh_monthly.empty:
                    df_sh_monthly.to_excel(writer, sheet_name="上海货盘费用类型月汇总表", index=False)
                    print(f"  - 上海货盘费用类型月汇总表: {len(df_sh_monthly)} 行")

            # Sheet 6: 限制说明
            limitations = [
                {"项目": "数据范围", "说明": "仅含仓库履约成本，不含SKU商品成本"},
                {"项目": "匹配能力", "说明": "无「订单→SKU→成本」链路"},
                {"项目": "订单号", "说明": "仓库订单号 ≠ 平台 order_id"},
                {
                    "项目": "上海货盘账单",
                    "说明": "已按拆柜费/港口短驳拖车费/卸货费/仓储费/仓库操作费/贴标费输出月度汇总，明细仅在解析结果中保留",
                },
                {"项目": "综合损益概览币种", "说明": "按月份+币种核算，不做汇率折算"},
                {"项目": "Phase 3", "说明": "SKU级成本、商品毛利需补充订单明细数据"},
            ]
            df_limits = pd.DataFrame(limitations)
            df_limits.to_excel(writer, sheet_name="限制说明", index=False)

        print(f"\n报表已生成: {output_file}")

    except PermissionError:
        print(f"\n❌ 错误: 无法写入文件 {output_file}")
        print("💡 原因: 文件可能已被打开。请关闭 Excel 文件后重试。")
    except Exception as e:
        print(f"\n❌ 生成报表时出错: {e}")

    print("=" * 70)
    return output_file


if __name__ == "__main__":
    run_phase2()
