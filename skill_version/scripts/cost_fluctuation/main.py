#!/usr/bin/env python3
"""
月度完工产品成本波动分析
=========================
读取各月工单明细，按 完工产品材料费+直接人工+制造费用 跨月汇总，
选取 Top N 产品，计算各月完工单价波动。

用法:
  python -m scripts.cost_fluctuation.main --output-dir ../output --top 20

输出:
  output/cost_fluctuation.json  — Top产品年汇总 + 月度明细 + 波动摘要
  output/cost_fluctuation.csv   — 月度明细表 (每行=产品×月份)
"""

import argparse
import json
import sys
import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import numpy as np


# ============================================================================
# 工具函数
# ============================================================================

def parse_month_key(dirname: str):
    """2025Y01M -> (2025, 1)"""
    dirname = os.path.basename(dirname.rstrip('/\\'))
    try:
        parts = dirname.split('Y')
        if len(parts) != 2:
            return None
        year = int(parts[0])
        month = int(parts[1].replace('M', ''))
        return (year, month)
    except (ValueError, IndexError):
        return None


def find_month_dirs(output_dir: str) -> list[tuple[int, int, Path]]:
    """扫描 output/ 下所有月份目录"""
    root = Path(output_dir)
    months = []
    for entry in root.iterdir():
        if entry.is_dir() and not entry.name.startswith('_'):
            key = parse_month_key(entry.name)
            if key:
                months.append((key[0], key[1], entry))
    return sorted(months)


def load_order_detail(month_dir: Path) -> pd.DataFrame:
    """读取工单明细表，按产品编码汇总当月完工成本"""
    csv_path = month_dir / '工单明细.csv'
    if not csv_path.exists():
        return None
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    # 按产品编码汇总
    agg = df.groupby('产品编码').agg(
        完工数量=('完工数量', 'sum'),
        完工材料费=('完工产品材料费', 'sum'),
        完工人工费=('完工产品直接人工', 'sum'),
        完工制费=('完工产品制造费用', 'sum'),
    ).reset_index()
    agg['完工成本合计'] = agg['完工材料费'] + agg['完工人工费'] + agg['完工制费']
    return agg


# ============================================================================
# 核心分析
# ============================================================================

def _get_dimension_hint(p_rows) -> str:
    """从单价变化推断归因"""
    if len(p_rows) < 2:
        return ""
    first = p_rows.iloc[0]
    last = p_rows.iloc[-1]
    qty_change = abs(float(last['完工数量']) - float(first['完工数量']))
    amt_change = abs(float(last['完工成本合计']) - float(first['完工成本合计']))
    if amt_change < 0.01:
        return "基本无变化"
    if qty_change > 0 and amt_change / qty_change > 100:
        return "单价变动主导"
    else:
        return "数量+单价综合变动"


def analyze(output_dir: str, top_n: int = 20):
    """
    1. 用工单明细跨月汇总（完工产品材料费+人工+制费），取 Top N
    2. 对 Top N 产品逐月计算完工单价
    3. 只有1个月有生产的产品不计入波动排名
    """
    month_dirs = find_month_dirs(output_dir)
    if not month_dirs:
        return {"status": "no_data", "message": "未找到月份数据"}

    # ---- 逐月读取工单明细 ----
    monthly = {}  # {(y,m): DataFrame (产品编码汇总)}
    for y, m, mdir in month_dirs:
        df = load_order_detail(mdir)
        if df is not None and not df.empty:
            monthly[(y, m)] = df

    if not monthly:
        return {"status": "no_order_data", "message": "未找到工单明细数据"}

    # ---- Step 1: 跨月汇总 → Top N ----
    # 用 完工材料费+人工+制费 作为产品完工成本
    product_year_total = {}  # {产品编码: 全年完工成本合计}
    for (y, m), df in monthly.items():
        for _, row in df.iterrows():
            prod = str(row['产品编码'])
            cost = float(row['完工成本合计'])
            product_year_total[prod] = product_year_total.get(prod, 0) + cost

    total_year_cost = sum(product_year_total.values())
    ranked = sorted(product_year_total.items(), key=lambda x: x[1], reverse=True)
    top_products = ranked[:top_n]

    top_summary = []
    for rank, (prod, amt) in enumerate(top_products, 1):
        top_summary.append({
            '序号': rank,
            '产品编码': prod,
            '完工金额': round(amt, 2),
            '占当年完工金额比': f"{amt/total_year_cost:.4%}" if total_year_cost > 0 else "N/A",
        })

    # ---- Step 2: Top N 月度明细 ----
    top_set = {p for p, _ in top_products}
    month_keys = sorted(monthly.keys())
    detail_rows = []

    for prod in sorted(top_set):
        prev_unit_price = None
        for y, m in month_keys:
            df = monthly[(y, m)]
            row = df[df['产品编码'] == prod]
            month_label = f"{y}Y{m:02d}M"

            if row.empty:
                # 该月无生产，单价为0
                detail_rows.append({
                    '产品编码': prod,
                    '月份': month_label,
                    '完工成本合计': 0,
                    '完工数量': 0,
                    '完工单价': 0,
                    '单价较上期波动': "0.00%",
                })
                prev_unit_price = 0
                continue

            r = row.iloc[0]
            cost_total = float(r['完工成本合计'])
            finished_qty = float(r['完工数量'])
            unit_price = cost_total / finished_qty if finished_qty > 0 else 0

            # 单价波动 (只在上期单价>0时计算)
            if prev_unit_price is not None and prev_unit_price > 0:
                price_change = (unit_price - prev_unit_price) / prev_unit_price
            else:
                price_change = 0

            detail_rows.append({
                '产品编码': prod,
                '月份': month_label,
                '完工成本合计': round(cost_total, 2),
                '完工数量': round(finished_qty, 2),
                '完工单价': round(unit_price, 4),
                '单价较上期波动': f"{price_change:.4%}",
            })
            prev_unit_price = unit_price

    # ---- 波动统计 ----
    df_detail = pd.DataFrame(detail_rows)
    if df_detail.empty:
        return {"status": "no_detail", "message": "Top产品无月度明细"}

    # 提取波动数值用于统计
    df_detail['_波动值'] = df_detail['单价较上期波动'].apply(
        lambda x: float(x.replace('%', '')) / 100 if isinstance(x, str) and x != 'N/A' else 0
    )

    # 波动摘要：统计每个产品在各月的生产情况
    fluctuation_records = []
    for prod in sorted(top_set):
        p_rows = df_detail[df_detail['产品编码'] == prod].sort_values('月份')

        # 只统计有生产的月份（完工数量 > 0）
        active_rows = p_rows[p_rows['完工数量'] > 0]
        if len(active_rows) < 2:
            # 只有1个月有生产，不计入波动排名
            continue

        unit_prices = active_rows['完工单价'].values
        price_changes = []
        prev_up = None
        for _, ar in active_rows.iterrows():
            up = float(ar['完工单价'])
            if prev_up is not None and prev_up > 0:
                price_changes.append((up - prev_up) / prev_up)
            prev_up = up

        if len(price_changes) > 0:
            abs_changes = np.abs(price_changes)
            max_abs_idx = int(np.argmax(abs_changes))
            max_change = float(abs_changes[max_abs_idx])
            max_change_signed = float(price_changes[max_abs_idx])
            # 找到对应的月份
            active_months = active_rows['月份'].values
            max_change_month = active_months[1 + max_abs_idx] if 1 + max_abs_idx < len(active_months) else active_months[-1]
        else:
            max_change = 0
            max_change_signed = 0
            max_change_month = ''

        fluctuation_records.append({
            '产品编码': prod,
            '年度总完工金额': round(float(p_rows['完工成本合计'].sum()), 2),
            '月均单价': round(float(unit_prices.mean()), 4),
            '最高单价': round(float(unit_prices.max()), 4),
            '最低单价': round(float(unit_prices.min()), 4),
            '最大月波动': f"{max_change:.2%}",
            '_max_change_signed': max_change_signed,
            '最大波动月份': max_change_month,
            '生产月份数': len(active_rows),
        })

    fluctuation_records.sort(key=lambda x: abs(float(x['最大月波动'].replace('%',''))/100), reverse=True)

    month_labels = [f"{y}Y{m:02d}M" for y, m in month_keys]

    # ---- 波动分布 ----
    abs_changes_vals = []
    for fr in fluctuation_records:
        val = abs(float(fr['最大月波动'].replace('%', '')) / 100)
        abs_changes_vals.append(val)
    abs_changes_vals = np.array(abs_changes_vals) if abs_changes_vals else np.array([])

    fluctuation_distribution = {
        "大幅波动(>5%)": int(np.sum(abs_changes_vals > 0.05)),
        "小幅波动(1%-5%)": int(np.sum((abs_changes_vals > 0.01) & (abs_changes_vals <= 0.05))),
        "基本持平(<1%)": int(np.sum(abs_changes_vals <= 0.01)),
    }

    # ---- 月度对统计 ----
    month_pair_stats = {}
    for i in range(1, len(month_keys)):
        prev_key = f"{month_keys[i-1][0]}Y{month_keys[i-1][1]:02d}M"
        curr_key = f"{month_keys[i][0]}Y{month_keys[i][1]:02d}M"
        pair_rows = df_detail[(df_detail['月份'] == prev_key) | (df_detail['月份'] == curr_key)]
        month_pair_stats[f"{prev_key}->{curr_key}"] = {
            '产品数': len(pair_rows['产品编码'].unique()),
            '上涨产品数': int(len(pair_rows[pair_rows['_波动值'] > 0.001]['产品编码'].unique())),
            '下跌产品数': int(len(pair_rows[pair_rows['_波动值'] < -0.001]['产品编码'].unique())),
            '平均完工金额': round(float(pair_rows['完工成本合计'].mean()), 2),
        }

    # ---- Top5 波动 (含归因) ----
    top5_list = []
    for f in fluctuation_records[:5]:
        prod = f['产品编码']
        p_rows = df_detail[df_detail['产品编码'] == prod].sort_values('月份')
        # 只取有生产的月份
        active_rows = p_rows[p_rows['完工数量'] > 0]
        if len(active_rows) >= 2:
            top5_list.append({
                '产品编码': prod,
                '起始月份': month_labels[0],
                '终止月份': month_labels[-1],
                '前期成本': round(float(active_rows.iloc[0]['完工成本合计']), 2),
                '本期成本': round(float(active_rows.iloc[-1]['完工成本合计']), 2),
                '变动金额': round(float(active_rows.iloc[-1]['完工成本合计'] - active_rows.iloc[0]['完工成本合计']), 2),
                '变动百分比': f["最大月波动"],
                '变动方向': '上涨' if f.get('_max_change_signed', 0) > 0.001 else ('下跌' if f.get('_max_change_signed', 0) < -0.001 else '持平'),
                '主要变动维度': _get_dimension_hint(active_rows),
                '生产月份数': f['生产月份数'],
            })

    summary = {
        "status": "completed",
        "analysis_timestamp": datetime.now().isoformat(),
        "months_analyzed": len(month_keys),
        "month_range": f"{month_labels[0]} → {month_labels[-1]}" if month_labels else "",
        "total_products": len(product_year_total),
        "total_year_cost": round(total_year_cost, 2),
        "top_n": top_n,
        "top_products_summary": top_summary,
        "fluctuation_ranking": fluctuation_records,
        "fluctuation_distribution": fluctuation_distribution,
        "month_pair_statistics": month_pair_stats,
        "top5_fluctuations": top5_list,
    }

    return {
        "summary": summary,
        "detail": detail_rows,
        "dataframe": df_detail.drop(columns=['_波动值']),
    }


# ============================================================================
# 主入口
# ============================================================================

def main():
    parser = argparse.ArgumentParser(description="月度完工产品成本波动分析")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="pipeline 输出目录")
    parser.add_argument("--top", type=int, default=20,
                        help="分析前 N 大产品 (默认 20)")
    args = parser.parse_args()

    output_dir = args.output_dir
    if not Path(output_dir).is_dir():
        print(f"[ERR] 输出目录不存在: {output_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"[SCAN] 成本波动分析: Top {args.top} 产品")
    result = analyze(output_dir, top_n=args.top)

    status = result.get("summary", {}).get("status", "")
    if status != "completed":
        msg = result.get("message", result.get("summary", {}).get("message", status))
        print(f"[WARN] {msg}")
        sys.exit(0)

    summary = result["summary"]

    # ---- 写出 ----
    root = Path(output_dir)
    json_path = root / "cost_fluctuation.json"
    csv_path = root / "cost_fluctuation.csv"

    # 清理内部字段
    for item in summary.get('fluctuation_ranking', []):
        item.pop('_max_change_signed', None)

    # 包裹 summary + detail，兼容 HTML 模板的 .summary 访问
    wrapper = {"summary": summary, "detail": result.get("detail", [])}
    json_path.write_text(json.dumps(wrapper, ensure_ascii=False, indent=2), encoding='utf-8')
    result["dataframe"].to_csv(csv_path, index=False, encoding='utf-8-sig')

    # ---- 打印摘要 ----
    print()
    print("=" * 60)
    print("  月度完工产品成本波动分析")
    print("=" * 60)
    print(f"月份范围: {summary['month_range']}")
    print(f"产品总数: {summary['total_products']}, 全年完工总成本: Y{summary['total_year_cost']:,.2f}")
    print()
    print(f"--- Top {args.top} 产品 ---")
    for t in summary['top_products_summary']:
        print(f"  {t['序号']:2d}. {t['产品编码']}: Y{t['完工金额']:,.2f} ({t['占当年完工金额比']})")
    print()
    print("--- 单价波动排名 (最大月波动) ---")
    for f in summary.get('fluctuation_ranking', [])[:10]:
        print(f"  {f['产品编码']}: 月均Y{f['月均单价']:,.2f}, "
              f"最高Y{f['最高单价']:,.2f}, 最低Y{f['最低单价']:,.2f}, "
              f"最大波动 {f['最大月波动']} ({f['最大波动月份']})")
    print()
    print(f"[OK] 结果: {json_path.resolve()}")
    print(f"[OK] 结果: {csv_path.resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
