#!/usr/bin/env python3
"""
成本分析 HTML 报告生成器
=========================
读取 pipeline 分析输出的 JSON，注入 assets/report_template.html 模板，
生成独立的 HTML 报告。

用法:
  python scripts/generate_report.py --output-dir ../output

输出:
  output/cost_report.html  (浏览器直接打开，动态渲染图表)
"""

import argparse
import json
import sys
from pathlib import Path
from datetime import datetime


def load_json(path: Path) -> dict:
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def main():
    parser = argparse.ArgumentParser(description="生成成本分析 HTML 报告")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="pipeline 输出目录")
    args = parser.parse_args()

    root = Path(args.output_dir)

    # ---- 读取数据 ----
    index_path = root / "_index.json"
    fluct_path = root / "cost_fluctuation.json"
    margin_path = root / "margin_analysis.json"

    missing = []
    for p, name in [(index_path, "_index.json"), (fluct_path, "cost_fluctuation.json"),
                     (margin_path, "margin_analysis.json")]:
        if not p.exists():
            missing.append(name)
    if missing:
        print(f"[ERR] 缺少数据文件: {missing}")
        print("请先运行 pipeline_cli.py, cost_fluctuation.py, margin_analysis.py")
        sys.exit(1)

    data = {
        "index": load_json(index_path),
        "fluctuation": load_json(fluct_path),
        "margin": load_json(margin_path),
    }

    # ---- 读取模板 ----
    script_dir = Path(__file__).resolve().parent
    template_path = script_dir.parent / "assets" / "report_template.html"
    if not template_path.exists():
        print(f"[ERR] 模板不存在: {template_path}")
        sys.exit(1)

    template = template_path.read_text(encoding='utf-8')

    # ---- 注入数据 ----
    data_json = json.dumps(data, ensure_ascii=False, indent=2)
    html = template.replace("__REPORT_DATA_PLACEHOLDER__", data_json)

    # ---- 写出 ----
    report_path = (root / "cost_report.html").resolve()
    report_path.write_text(html, encoding='utf-8')
    print(f"[OK] 报告已生成: {report_path}")
    print(f"     浏览器打开即可查看")


if __name__ == "__main__":
    main()
