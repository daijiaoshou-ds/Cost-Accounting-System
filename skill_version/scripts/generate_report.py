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


def generate_report(output_dir: str) -> str:
    """
    从 output 目录的 JSON 数据生成 HTML 报告。

    Returns:
        生成的 HTML 文件路径；若数据不完整则返回 None
    """
    root = Path(output_dir)

    index_path = root / "_index.json"
    fluct_path = root / "cost_fluctuation.json"
    margin_path = root / "margin_analysis.json"

    missing = []
    for p, name in [(index_path, "_index.json"), (fluct_path, "cost_fluctuation.json"),
                     (margin_path, "margin_analysis.json")]:
        if not p.exists():
            missing.append(name)
    if missing:
        print(f"[WARN] 缺少数据文件: {missing}，跳过 HTML 报告生成")
        return None

    data = {
        "index": load_json(index_path),
        "fluctuation": load_json(fluct_path),
        "margin": load_json(margin_path),
    }

    script_dir = Path(__file__).resolve().parent
    template_path = script_dir.parent / "assets" / "report_template.html"
    if not template_path.exists():
        print(f"[ERR] 模板不存在: {template_path}")
        return None

    template = template_path.read_text(encoding='utf-8')
    data_json = json.dumps(data, ensure_ascii=False, indent=2)
    html = template.replace("__REPORT_DATA_PLACEHOLDER__", data_json)

    report_path = (root / "cost_report.html").resolve()
    report_path.write_text(html, encoding='utf-8')
    print(f"[OK] 报告已生成: {report_path}")
    return str(report_path)


def main():
    # 强制 UTF-8 输出，避免 Windows 终端中文乱码
    import sys, io
    if sys.platform == "win32":
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="生成成本分析 HTML 报告")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="pipeline 输出目录")
    args = parser.parse_args()

    report_path = generate_report(args.output_dir)
    if report_path is None:
        sys.exit(1)
    print(f"     浏览器打开即可查看")


if __name__ == "__main__":
    main()
