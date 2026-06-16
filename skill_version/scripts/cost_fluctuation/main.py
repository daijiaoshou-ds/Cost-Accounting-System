#!/usr/bin/env python3
"""
月度完工产品成本波动分析
=========================
读取各月收发存汇总，计算 Top N 产品的月度完工单价波动。

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


def load_receipt_storage(month_dir: Path) -> pd.DataFrame:
    """读取收发存汇总表"""
    csv_path = month_dir / '收发存汇总.csv'
    if not csv_path.exists():
        return None
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    # 物料编码可能是数字，pandas 会读成 float，需要还原为整数字符串
    df['物料编码'] = df['物料编码'].apply(_clean_code)
    return df


def _clean_code(val):
    """将 float 物料编码还原为纯数字字符串 (去掉 .0)"""
    try:
        if pd.isna(val):
            return ''
        f = float(val)
        if f == int(f):
            return str(int(f))
        return str(val).strip()
    except (ValueError, TypeError):
        return str(val).strip()


# ============================================================================
# 核心分析
# ============================================================================

def _get_dimension_hint(p_rows) -> str:
    """从单价变化推断归因（简化版：根据完工金额和数量变化判断是量变还是价变）"""
    if len(p_rows) < 2:
        return ""
    first = p_rows.iloc[0]
    last = p_rows.iloc[-1]
    qty_change = abs(float(last['完工入库数量']) - float(first['完工入库数量']))
    amt_change = abs(float(last['完工金额']) - float(first['完工金额']))
    if amt_change < 0.01:
        return "基本无变化"
    if qty_change > 0 and amt_change / qty_change > 100:
        return "单价变动主导"
    else:
        return "数量+单价综合变动"


def analyze(output_dir: str, top_n: int = 20):
    """
    1. 汇总全年各产品完工金额，取 Top N
    2. 对 Top N 产品逐月展开：完工金额、入库数量、完工单价、单价波动
    """
    month_dirs = find_month_dirs(output_dir)
    if not month_dirs:
        return {"status": "no_data", "message": "未找到月份数据"}

    # ---- 逐月读取 ----
    monthly = {}  # {(y,m): DataFrame (收发存汇总)}
    for y, m, mdir in month_dirs:
        df = load_receipt_storage(mdir)
        if df is not None and not df.empty:
            monthly[(y, m)] = df

    if not monthly:
        return {"status": "no_receipt_data", "message": "未找到收发存汇总数据"}

    # ---- Step 1: 全年汇总 → Top N ----
    product_year_total = {}  # {产品编码: 全年完工金额}
    for (y, m), df in monthly.items():
        for _, row in df.iterrows():
            mat = str(row['物料编码'])
            amt = float(row.get('本期收入金额', 0) or 0)
            product_year_total[mat] = product_year_total.get(mat, 0) + amt

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
            row = df[df['物料编码'] == prod]
            if row.empty:
                continue

            r = row.iloc[0]
            finished_amt = float(r.get('本期收入金额', 0) or 0)
            finished_qty = float(r.get('本期完工数量', 0) or 0)
            unit_price = finished_amt / finished_qty if finished_qty > 0 else 0

            # 单价波动
            if prev_unit_price is not None and prev_unit_price > 0:
                price_change = (unit_price - prev_unit_price) / prev_unit_price
            else:
                price_change = 0

            month_label = f"{y}Y{m:02d}M"
            detail_rows.append({
                '产品编码': prod,
                '月份': month_label,
                '完工金额': round(finished_amt, 2),
                '完工入库数量': round(finished_qty, 2),
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

    # 波动摘要
    fluctuation_records = []
    for prod in sorted(top_set):
        p_rows = df_detail[df_detail['产品编码'] == prod].sort_values('月份')
        if len(p_rows) < 2:
            continue
        unit_prices = p_rows['完工单价'].values
        price_changes = p_rows['_波动值'].values[1:]  # skip first month
        if len(price_changes) > 0:
            abs_changes = np.abs(price_changes)
            max_abs_idx = int(np.argmax(abs_changes))
            max_change = float(abs_changes[max_abs_idx])
            max_change_signed = float(price_changes[max_abs_idx])
            max_change_month = p_rows.iloc[1 + max_abs_idx]['月份']
        else:
            max_change = 0
            max_change_signed = 0
            max_change_month = ''

        fluctuation_records.append({
            '产品编码': prod,
            '年度总完工金额': round(float(p_rows['完工金额'].sum()), 2),
            '月均单价': round(float(unit_prices.mean()), 4),
            '最高单价': round(float(unit_prices.max()), 4),
            '最低单价': round(float(unit_prices.min()), 4),
            '最大月波动': f"{max_change:.2%}",
            '_max_change_signed': max_change_signed,
            '最大波动月份': max_change_month,
        })

    fluctuation_records.sort(key=lambda x: abs(float(x['最大月波动'].replace('%',''))/100), reverse=True)

    month_labels = [f"{y}Y{m:02d}M" for y, m in month_keys]

    # ---- 波动分布 (HTML 模板兼容) ----
    abs_changes = df_detail.groupby('产品编码')['_波动值'].apply(
        lambda x: float(np.max(np.abs(x.values)))
    ).values
    fluctuation_distribution = {
        "大幅波动(>5%)": int(np.sum(abs_changes > 0.05)),
        "小幅波动(1%-5%)": int(np.sum((abs_changes > 0.01) & (abs_changes <= 0.05))),
        "基本持平(<1%)": int(np.sum(abs_changes <= 0.01)),
    }

    # ---- 月度对统计 (HTML 模板兼容) ----
    month_pair_stats = {}
    for i in range(1, len(month_keys)):
        prev_key = f"{month_keys[i-1][0]}Y{month_keys[i-1][1]:02d}M"
        curr_key = f"{month_keys[i][0]}Y{month_keys[i][1]:02d}M"
        pair_rows = df_detail[(df_detail['月份'] == prev_key) | (df_detail['月份'] == curr_key)]
        # stats by comparing consecutive months
        month_pair_stats[f"{prev_key}->{curr_key}"] = {
            '产品数': len(pair_rows['产品编码'].unique()),
            '上涨产品数': int(len(pair_rows[pair_rows['_波动值'] > 0.001]['产品编码'].unique())),
            '下跌产品数': int(len(pair_rows[pair_rows['_波动值'] < -0.001]['产品编码'].unique())),
            '平均完工金额': round(float(pair_rows['完工金额'].mean()), 2),
        }

    # ---- Top5 波动 (HTML 模板兼容: 含归因) ----
    top5_list = []
    for f in fluctuation_records[:5]:
        prod = f['产品编码']
        # 找最大波动月份的前一月对比
        p_rows = df_detail[df_detail['产品编码'] == prod].sort_values('月份')
        top5_list.append({
            '产品编码': prod,
            '起始月份': month_labels[0],
            '终止月份': month_labels[-1],
            '前期成本': round(float(p_rows.iloc[0]['完工金额']), 2) if len(p_rows) > 0 else 0,
            '本期成本': round(float(p_rows.iloc[-1]['完工金额']), 2) if len(p_rows) > 0 else 0,
            '变动金额': round(float(p_rows.iloc[-1]['完工金额'] - p_rows.iloc[0]['完工金额']), 2) if len(p_rows) > 0 else 0,
            '变动百分比': f["最大月波动"],
            '变动方向': '上涨' if f.get('_max_change_signed', 0) > 0.001 else ('下跌' if f.get('_max_change_signed', 0) < -0.001 else '持平'),
            '主要变动维度': _get_dimension_hint(p_rows) if len(p_rows) > 0 else '',
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
    df = result["dataframe"]

    # ---- 写出 ----
    root = Path(output_dir)
    json_path = root / "cost_fluctuation.json"
    csv_path = root / "cost_fluctuation.csv"

    # 清理内部字段
    for item in summary.get('fluctuation_ranking', []):
        item.pop('_max_change_signed', None)

    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')

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
