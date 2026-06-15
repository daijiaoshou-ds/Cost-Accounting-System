#!/usr/bin/env python3
"""
销售毛利率分析脚本
====================
读取 pipeline 输出的销售成本明细（含销售金额），按批次/产品/月份计算毛利率。

用法:
  python scripts/margin_analysis.py --output-dir ./output

输出:
  output/margin_analysis.json  -- 毛利率摘要 (总毛利/产品线/异常批次)
  output/margin_analysis.csv   -- 毛利率明细 (每行=批次×月份)
"""

import argparse
import json
import sys
import os
from pathlib import Path
from collections import defaultdict
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


def load_sales_cost(month_dir: Path) -> pd.DataFrame:
    """读取销售成本明细表"""
    csv_path = month_dir / '销售成本明细.csv'
    if not csv_path.exists():
        return None
    return pd.read_csv(csv_path, encoding='utf-8-sig')


# ============================================================================
# 核心分析
# ============================================================================

def analyze_margin(output_dir: str) -> dict:
    """
    按销售批次计算毛利率，汇总产品线/月份维度。
    返回: dict with summary + detailed records
    """
    month_dirs = find_month_dirs(output_dir)
    if not month_dirs:
        return {
            "status": "no_data",
            "message": "未找到任何月份数据",
        }

    # ---- 逐月加载销售成本 ----
    all_records = []
    monthly_stats = {}
    per_product_stats = defaultdict(lambda: {
        '销售金额合计': 0, '销售成本合计': 0, '毛利合计': 0, '批次数量': 0
    })

    for year, month, mdir in month_dirs:
        df = load_sales_cost(mdir)
        if df is None or df.empty:
            continue

        month_label = f"{year}Y{month:02d}M"
        month_revenue = 0
        month_cost = 0
        month_margin = 0
        batch_count = 0

        for _, row in df.iterrows():
            product = str(row.get('物料编码', ''))
            batch = str(row.get('销售批次号', ''))
            qty = float(row.get('销售数量', 0) or 0)
            revenue = float(row.get('销售金额', 0) or 0)
            cost_total = float(row.get('销售成本_合计', 0) or 0)
            margin = revenue - cost_total
            margin_pct = margin / revenue if revenue > 0 else None

            all_records.append({
                '月份': month_label,
                '产品编码': product,
                '销售批次号': batch,
                '销售数量': round(qty, 2),
                '销售金额': round(revenue, 2),
                '销售成本': round(cost_total, 2),
                '毛利': round(margin, 2),
                '毛利率': f"{margin_pct:.2%}" if margin_pct is not None else "N/A",
            })

            month_revenue += revenue
            month_cost += cost_total
            month_margin += margin
            batch_count += 1

            per_product_stats[product]['销售金额合计'] += revenue
            per_product_stats[product]['销售成本合计'] += cost_total
            per_product_stats[product]['毛利合计'] += margin
            per_product_stats[product]['批次数量'] += 1

        if batch_count > 0:
            monthly_stats[month_label] = {
                '批次数量': batch_count,
                '销售金额': round(month_revenue, 2),
                '销售成本': round(month_cost, 2),
                '毛利': round(month_margin, 2),
                '毛利率': f"{month_margin/month_revenue:.2%}" if month_revenue > 0 else "N/A",
            }

    if not all_records:
        return {
            "status": "no_sales_data",
            "message": "未找到销售成本明细数据（请确认上传了销售数据表）",
        }

    df_all = pd.DataFrame(all_records)

    # ---- 异常批次检测 ----
    # 1. 负毛利批次
    df_all['_margin'] = df_all['毛利']
    negative_margin = df_all[df_all['_margin'] < 0]
    # 2. 毛利率异常低 (<5%)
    df_all['_margin_pct'] = df_all.apply(
        lambda r: r['_margin'] / r['销售金额'] if r['销售金额'] > 0 else None, axis=1
    )
    low_margin = df_all[(df_all['_margin_pct'].notna()) & (df_all['_margin_pct'] < 0.05)]

    # ---- 产品线汇总 ----
    product_summary = []
    for prod, stats in per_product_stats.items():
        r = stats['销售金额合计']
        c = stats['销售成本合计']
        m = stats['毛利合计']
        product_summary.append({
            '产品编码': prod,
            '批次数量': stats['批次数量'],
            '销售金额合计': round(r, 2),
            '销售成本合计': round(c, 2),
            '毛利合计': round(m, 2),
            '毛利率': f"{m/r:.2%}" if r > 0 else "N/A",
        })

    product_summary.sort(key=lambda x: x['毛利合计'], reverse=True)

    # ---- 总览指标 ----
    total_revenue = sum(s['销售金额'] for s in monthly_stats.values())
    total_cost = sum(s['销售成本'] for s in monthly_stats.values())
    total_margin = total_revenue - total_cost
    total_batches = sum(s['批次数量'] for s in monthly_stats.values())

    summary = {
        "status": "completed",
        "analysis_timestamp": datetime.now().isoformat(),
        "months_analyzed": len(monthly_stats),
        "total_batches": total_batches,
        "total_revenue": round(total_revenue, 2),
        "total_cost": round(total_cost, 2),
        "total_margin": round(total_margin, 2),
        "overall_margin_rate": f"{total_margin/total_revenue:.2%}" if total_revenue > 0 else "N/A",
        "monthly_statistics": monthly_stats,
        "product_line_summary": product_summary,
        "anomalies": {
            "negative_margin_batches": len(negative_margin),
            "low_margin_batches": len(low_margin) - len(negative_margin),
            "negative_margin_details": negative_margin[[
                '月份', '产品编码', '销售批次号', '销售金额', '销售成本', '毛利', '毛利率'
            ]].to_dict(orient='records') if len(negative_margin) > 0 else [],
            "low_margin_details": low_margin[[
                '月份', '产品编码', '销售批次号', '销售金额', '销售成本', '毛利', '毛利率'
            ]].to_dict(orient='records')[:20] if len(low_margin) > 0 else [],
        },
    }

    return {
        "summary": summary,
        "records": all_records,
        "dataframe": df_all.drop(columns=['_margin', '_margin_pct'], errors='ignore'),
    }


# ============================================================================
# 主入口
# ============================================================================

def main():
    parser = argparse.ArgumentParser(description="销售毛利率分析")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="pipeline 输出目录 (默认 ./output)")
    args = parser.parse_args()

    output_dir = args.output_dir
    if not Path(output_dir).is_dir():
        print(f"[ERR] 输出目录不存在: {output_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"[SCAN] 毛利率分析: {output_dir}")
    result = analyze_margin(output_dir)

    if result.get("summary", {}).get("status") != "completed":
        msg = result.get('message', str(result)[:200])
        print(f"[WARN]  {msg}")
        sys.exit(0)

    summary = result["summary"]
    df = result["dataframe"]

    # ---- 写出 JSON ----
    json_path = Path(output_dir) / "margin_analysis.json"
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')

    # ---- 写出 CSV ----
    csv_path = Path(output_dir) / "margin_analysis.csv"
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')

    # ---- 打印摘要 ----
    print()
    print("=" * 60)
    print("  销售毛利率分析报告")
    print("=" * 60)
    print(f"月份跨度: {summary['months_analyzed']} 个月")
    print(f"总批次: {summary['total_batches']}")
    print(f"总销售收入: Y{summary['total_revenue']:,.2f}")
    print(f"总销售成本: Y{summary['total_cost']:,.2f}")
    print(f"总毛利: Y{summary['total_margin']:,.2f}")
    print(f"整体毛利率: {summary['overall_margin_rate']}")
    print()

    print("【月度概览】")
    for month, stats in summary['monthly_statistics'].items():
        print(f"  {month}: {stats['批次数量']}批次, "
              f"收入Y{stats['销售金额']:,.2f}, "
              f"毛利率 {stats['毛利率']}")

    print()
    print("【产品线毛利率排名】")
    for p in summary['product_line_summary'][:10]:
        print(f"  {p['产品编码']}: {p['批次数量']}批次, "
              f"毛利Y{p['毛利合计']:,.2f}, "
              f"毛利率 {p['毛利率']}")

    anomalies = summary['anomalies']
    if anomalies['negative_margin_batches'] > 0:
        print()
        print(f"[WARN] 负毛利批次: {anomalies['negative_margin_batches']} 个")
        for item in anomalies['negative_margin_details'][:5]:
            print(f"  {item['月份']} {item['产品编码']} {item['销售批次号']}: "
                  f"收入Y{item['销售金额']:,.2f}, 成本Y{item['销售成本']:,.2f}, "
                  f"毛利率 {item['毛利率']}")

    if anomalies['low_margin_batches'] > 0:
        print(f"[WARN] 低毛利批次(<5%): {anomalies['low_margin_batches']} 个")

    print()
    print(f"[OK] 结果已写出: {json_path}")
    print(f"[OK] 结果已写出: {csv_path}")
    print("=" * 60)

    sys.exit(0)


if __name__ == "__main__":
    main()
