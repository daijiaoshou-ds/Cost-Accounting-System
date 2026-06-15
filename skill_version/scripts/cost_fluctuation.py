#!/usr/bin/env python3
"""
月度成本波动分析脚本
=====================
读取 pipeline 输出的超级成本还原数据，计算每个产品在连续月份之间的成本波动。

用法:
  python scripts/cost_fluctuation.py --output-dir ./output

输出:
  output/cost_fluctuation.json  -- 波动摘要 (总波动/前5异常/月度趋势)
  output/cost_fluctuation.csv   -- 波动明细 (每行=产品×月份对)
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
    """扫描 output/ 下所有月份目录，返回 [(year, month, path)] 排序列表"""
    root = Path(output_dir)
    months = []
    for entry in root.iterdir():
        if entry.is_dir() and not entry.name.startswith('_'):
            key = parse_month_key(entry.name)
            if key:
                months.append((key[0], key[1], entry))
    return sorted(months)


# ============================================================================
# 核心分析
# ============================================================================

def load_super_restoration(month_dir: Path) -> pd.DataFrame:
    """读取超级还原完工成本表"""
    csv_path = month_dir / '超级还原_完工成本.csv'
    if not csv_path.exists():
        return None
    return pd.read_csv(csv_path, encoding='utf-8-sig')


def analyze_fluctuation(output_dir: str) -> dict:
    """
    分析每个产品在连续月份间的成本波动。
    返回: dict with summary + detailed records
    """
    month_dirs = find_month_dirs(output_dir)
    if len(month_dirs) < 2:
        return {
            "status": "insufficient_data",
            "message": f"需要至少2个月数据，当前只有 {len(month_dirs)} 个月",
            "months_found": [f"{y}Y{m:02d}M" for y, m, _ in month_dirs],
        }

    # ---- 逐月加载超级还原数据 ----
    monthly_data = {}  # {(y,m): DataFrame}
    for year, month, mdir in month_dirs:
        df = load_super_restoration(mdir)
        if df is not None and not df.empty:
            monthly_data[(year, month)] = df

    if len(monthly_data) < 2:
        return {
            "status": "insufficient_super_data",
            "message": "超级成本还原数据不足需2个月",
            "months_with_data": [f"{y}Y{m:02d}M" for y, m in monthly_data],
        }

    # ---- 按月+产品汇总成本 ----
    # 超级还原表结构: 成本序号, 产品编码, 成本维度, 金额, 占产品总成本比例, 占该维度总金额比例
    month_product_cost = {}  # {(y,m): {product: (total_cost, dims: {dim: cost})}}
    for (y, m), df in monthly_data.items():
        prod_costs = {}
        for prod, g in df.groupby('产品编码'):
            total = float(g['金额'].sum())
            dims = {}
            for _, row in g.iterrows():
                dim = str(row['成本维度'])
                dims[dim] = float(row['金额'])
            prod_costs[str(prod)] = (total, dims)
        month_product_cost[(y, m)] = prod_costs

    # ---- 计算逐月波动 ----
    month_keys = sorted(month_product_cost.keys())
    records = []

    for i in range(1, len(month_keys)):
        prev_key = month_keys[i - 1]
        curr_key = month_keys[i]
        prev_data = month_product_cost[prev_key]
        curr_data = month_product_cost[curr_key]

        all_products = set(prev_data.keys()) | set(curr_data.keys())

        for prod in sorted(all_products):
            prev_total = prev_data.get(prod, (0, {}))[0]
            curr_total = curr_data.get(prod, (0, {}))[0]

            if prev_total > 0:
                change_abs = curr_total - prev_total
                change_pct = change_abs / prev_total
            elif curr_total > 0:
                change_abs = curr_total
                change_pct = float('inf')
            else:
                continue

            # 波动归因：找出变动最大的维度
            prev_dims = prev_data.get(prod, (0, {}))[1]
            curr_dims = curr_data.get(prod, (0, {}))[1]
            all_dims = set(prev_dims.keys()) | set(curr_dims.keys())

            dim_changes = {}
            for dim in all_dims:
                d_prev = prev_dims.get(dim, 0)
                d_curr = curr_dims.get(dim, 0)
                dim_changes[dim] = round(d_curr - d_prev, 2)

            # 排序取前3大变动维度
            top_dims = sorted(dim_changes.items(), key=lambda x: abs(x[1]), reverse=True)[:3]
            top_dim_str = "; ".join(f"{d}({'+' if v >=0 else ''}{v:.0f})" for d, v in top_dims)

            records.append({
                '起始月份': f"{prev_key[0]}Y{prev_key[1]:02d}M",
                '终止月份': f"{curr_key[0]}Y{curr_key[1]:02d}M",
                '产品编码': prod,
                '前期成本': round(prev_total, 2),
                '本期成本': round(curr_total, 2),
                '变动金额': round(change_abs, 2),
                '变动百分比': f"{change_pct:.2%}",
                '变动方向': '上涨' if change_abs > 0 else ('下跌' if change_abs < 0 else '持平'),
                '主要变动维度': top_dim_str,
            })

    # ---- 构建摘要 ----
    df_all = pd.DataFrame(records)
    if df_all.empty:
        return {"status": "no_fluctuation", "message": "所有产品成本无变化"}

    # 找出最大波动项
    df_all['_变动金额_abs'] = df_all['变动金额'].abs()
    top5 = df_all.nlargest(5, '_变动金额_abs')

    # 按月份对统计
    month_pair_stats = {}
    for pair_key, g in df_all.groupby(['起始月份', '终止月份']):
        key_str = f"{pair_key[0]}→{pair_key[1]}"
        month_pair_stats[key_str] = {
            '产品数': len(g),
            '上涨产品数': int((g['变动方向'] == '上涨').sum()),
            '下跌产品数': int((g['变动方向'] == '下跌').sum()),
            '平均变动金额': round(float(g['变动金额'].mean()), 2),
            '最大上涨': round(float(g['变动金额'].max()), 2),
            '最大下跌': round(float(g['变动金额'].min()), 2),
        }

    summary = {
        "status": "completed",
        "analysis_timestamp": datetime.now().isoformat(),
        "months_analyzed": len(month_keys),
        "products_tracked": int(df_all['产品编码'].nunique()),
        "total_records": len(records),
        "month_pair_statistics": month_pair_stats,
        "top5_fluctuations": top5[[
            '产品编码', '起始月份', '终止月份', '前期成本', '本期成本',
            '变动金额', '变动百分比', '变动方向', '主要变动维度'
        ]].to_dict(orient='records'),
        "fluctuation_distribution": {
            "上涨(>5%)": int((df_all['_变动金额_abs'] / df_all['前期成本'] > 0.05).sum()),
            "微涨(1%-5%)": int(((df_all['_变动金额_abs'] / df_all['前期成本'] > 0.01) & (df_all['_变动金额_abs'] / df_all['前期成本'] <= 0.05)).sum()),
            "基本持平(<1%)": int((df_all['_变动金额_abs'] / df_all['前期成本'] <= 0.01).sum()),
        },
    }

    return {
        "summary": summary,
        "records": records,
        "dataframe": df_all.drop(columns=['_变动金额_abs']),
    }


# ============================================================================
# 主入口
# ============================================================================

def main():
    parser = argparse.ArgumentParser(description="月度成本波动分析")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="pipeline 输出目录 (默认 ./output)")
    args = parser.parse_args()

    output_dir = args.output_dir
    if not Path(output_dir).is_dir():
        print(f"[ERR] 输出目录不存在: {output_dir}", file=sys.stderr)
        sys.exit(1)

    print(f"[SCAN] 分析成本波动: {output_dir}")
    result = analyze_fluctuation(output_dir)

    if result.get("status") == "insufficient_data":
        print(f"[WARN]  {result['message']}")
        json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
        sys.exit(0)

    summary = result["summary"]
    records = result["records"]
    df = result["dataframe"]

    # ---- 写出 JSON ----
    json_path = Path(output_dir) / "cost_fluctuation.json"
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')

    # ---- 写出 CSV ----
    csv_path = Path(output_dir) / "cost_fluctuation.csv"
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')

    # ---- 打印摘要 ----
    print()
    print("=" * 60)
    print("  月度成本波动分析报告")
    print("=" * 60)
    print(f"月份跨度: {summary['months_analyzed']} 个月")
    print(f"产品追踪: {summary['products_tracked']} 个")
    print()
    print("【Top5 波动最大的产品】")
    for item in summary['top5_fluctuations']:
        direction = item['变动方向']
        print(f"  {item['产品编码']}: {item['起始月份']}→{item['终止月份']} "
              f"{direction} Y{item['变动金额']:,.2f} ({item['变动百分比']})")
        print(f"    归因: {item['主要变动维度']}")
    print()
    for pair, stats in summary.get('month_pair_statistics', {}).items():
        print(f"  {pair}: {stats['产品数']}产品, "
              f"↑{stats['上涨产品数']} ↓{stats['下跌产品数']}, "
              f"均值Y{stats['平均变动金额']:,.2f}")
    print()
    print(f"[OK] 结果已写出: {json_path}")
    print(f"[OK] 结果已写出: {csv_path}")
    print("=" * 60)

    sys.exit(0)


if __name__ == "__main__":
    main()
