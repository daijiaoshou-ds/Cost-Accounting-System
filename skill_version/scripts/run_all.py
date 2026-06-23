#!/usr/bin/env python3
"""
统一编排脚本
=============
AI 只需调用这一个命令，完成：环境检查 → 核算 → 波动分析 → 毛利率分析 → HTML 报告

用法:
  python scripts/run_all.py --config config.json --output-dir output
  python scripts/run_all.py --config config.json --skip-env   # 跳过环境检查
"""

import argparse
import json
import sys
import time
import os
import webbrowser
from pathlib import Path
from datetime import datetime

# 确保 skill_version 及子模块目录在 path 中
_skill_dir = Path(__file__).resolve().parent.parent
_cost_calc_dir = _skill_dir / "scripts" / "cost_calculation"
for d in [str(_skill_dir), str(_cost_calc_dir)]:
    if d not in sys.path:
        sys.path.insert(0, d)


def step_result(status, **kwargs):
    return {"status": status, **kwargs}


def main():
    # 强制 UTF-8 输出，避免 Windows 终端中文乱码
    import io
    if sys.platform == "win32":
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        except Exception:
            pass

    parser = argparse.ArgumentParser(description="矩阵成本核算 — 统一编排")
    parser.add_argument("--config", type=str, default="test_config.json",
                        help="JSON 配置文件路径")
    parser.add_argument("--output-dir", type=str, default="./output",
                        help="输出目录")
    parser.add_argument("--skip-env", action="store_true",
                        help="跳过环境检查")
    parser.add_argument("--top", type=int, default=20,
                        help="成本波动 Top N 产品数")
    parser.add_argument("--force", action="store_true", default=False,
                        help="强制计算（超领时自动归一化而非报错）")
    args = parser.parse_args()

    config_path = Path(args.config)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    run_summary = {
        "status": "running",
        "steps": {},
        "total_time_seconds": 0,
    }
    t_total = time.time()

    # ---- Step 0: 环境检查 ----
    if not args.skip_env:
        t0 = time.time()
        try:
            from scripts.check_env import run_check
            env = run_check()
            status = "ok" if env["all_ok"] else "warning"
            run_summary["steps"]["env_check"] = step_result(
                status, all_ok=env["all_ok"], missing=env.get("missing_packages", [])
            )
            if not env["all_ok"]:
                print(f"[WARN] 环境不完整: {env['missing_packages']}")
        except Exception as e:
            run_summary["steps"]["env_check"] = step_result("error", error=str(e))
        run_summary["steps"]["env_check"]["time"] = round(time.time() - t0, 3)

    # ---- Step 1: 核心核算 ----
    t0 = time.time()
    try:
        from scripts.cost_calculation.pipeline_cli import run_from_config
        ctx, index_path = run_from_config(
            str(config_path.resolve()), str(output_dir.resolve()),
            force_calculate_override=args.force if args.force else None
        )
        if ctx.log.has_errors or index_path is None:
            run_summary["steps"]["core_calculation"] = step_result(
                "error", errors=ctx.log.errors
            )
        else:
            run_summary["steps"]["core_calculation"] = step_result(
                "ok", index=str(index_path), months=len(ctx.all_months)
            )
    except Exception as e:
        run_summary["steps"]["core_calculation"] = step_result("error", error=str(e))
    run_summary["steps"]["core_calculation"]["time"] = round(time.time() - t0, 3)

    # ---- Step 2: 成本波动 ----
    t0 = time.time()
    try:
        from scripts.cost_fluctuation.main import analyze as analyze_fluctuation
        fluct_result = analyze_fluctuation(str(output_dir.resolve()), top_n=args.top)
        status = fluct_result.get("summary", {}).get("status", "error")
        if status == "completed":
            summary = fluct_result["summary"]
            # 清理内部字段
            for item in summary.get('fluctuation_ranking', []):
                item.pop('_max_change_signed', None)
            fluct_json_path = output_dir / "cost_fluctuation.json"
            fluct_csv_path = output_dir / "cost_fluctuation.csv"
            # 包裹 summary + detail，兼容 HTML 模板的 .summary 访问
            wrapper = {"summary": summary, "detail": fluct_result.get("detail", [])}
            fluct_json_path.write_text(json.dumps(wrapper, ensure_ascii=False, indent=2), encoding='utf-8')
            fluct_result["dataframe"].to_csv(fluct_csv_path, index=False, encoding='utf-8-sig')
        else:
            fluct_json_path = output_dir / "cost_fluctuation.json"
        run_summary["steps"]["cost_fluctuation"] = step_result(
            "ok" if status == "completed" else "warning",
            products=fluct_result.get("summary", {}).get("total_products", 0),
            json=str(fluct_json_path.resolve()) if fluct_json_path.exists() else None,
        )
    except Exception as e:
        run_summary["steps"]["cost_fluctuation"] = step_result("error", error=str(e))
    run_summary["steps"]["cost_fluctuation"]["time"] = round(time.time() - t0, 3)

    # ---- Step 3: 毛利率分析 ----
    t0 = time.time()
    try:
        from scripts.margin_analysis.main import analyze_margin
        margin_result = analyze_margin(str(output_dir.resolve()))
        status = margin_result.get("summary", {}).get("status", "error")
        if status == "completed":
            margin_json_path = output_dir / "margin_analysis.json"
            margin_csv_path = output_dir / "margin_analysis.csv"
            # 包裹 summary + records_count，兼容 HTML 模板的 .summary 访问
            wrapper = {"summary": margin_result["summary"], "records_count": len(margin_result.get("records", []))}
            margin_json_path.write_text(
                json.dumps(wrapper, ensure_ascii=False, indent=2),
                encoding='utf-8')
            margin_result["dataframe"].to_csv(margin_csv_path, index=False, encoding='utf-8-sig')
        else:
            margin_json_path = output_dir / "margin_analysis.json"
        run_summary["steps"]["margin_analysis"] = step_result(
            "ok" if status == "completed" else "warning",
            batches=margin_result.get("summary", {}).get("total_batches", 0),
            json=str(margin_json_path.resolve()) if margin_json_path.exists() else None,
        )
    except Exception as e:
        run_summary["steps"]["margin_analysis"] = step_result("error", error=str(e))
    run_summary["steps"]["margin_analysis"]["time"] = round(time.time() - t0, 3)

    # ---- Step 4: HTML 报告 ----
    t0 = time.time()
    try:
        from scripts.generate_report import generate_report as gen_html
        html_path = gen_html(str(output_dir.resolve()))
        if html_path:
            run_summary["steps"]["html_report"] = step_result("ok", html=html_path)
            # 自动在默认浏览器打开报告
            try:
                webbrowser.open(f"file:///{html_path}")
            except Exception:
                pass
        else:
            run_summary["steps"]["html_report"] = step_result(
                "skipped", reason="missing analysis JSON files"
            )
    except Exception as e:
        run_summary["steps"]["html_report"] = step_result("error", error=str(e))
    run_summary["steps"]["html_report"]["time"] = round(time.time() - t0, 3)

    # ---- 汇总 ----
    total_time = time.time() - t_total
    run_summary["total_time_seconds"] = round(total_time, 3)

    # 状态判断: error 优先 → warning → ok
    if any(s.get("status") == "error" for s in run_summary["steps"].values()):
        run_summary["status"] = "failed"
    elif all(s.get("status") == "ok" for s in run_summary["steps"].values()):
        run_summary["status"] = "completed"
    else:
        run_summary["status"] = "completed_with_warnings"

    summary_path = output_dir / "_run_summary.json"
    summary_path.write_text(json.dumps(run_summary, ensure_ascii=False, indent=2), encoding='utf-8')

    print()
    print(f"=== Run Complete: {run_summary['status']} ({total_time:.1f}s) ===")
    for step_name, step in run_summary["steps"].items():
        icon = "[OK]" if step["status"] == "ok" else "[WARN]" if step["status"] == "warning" else "[ERR]"
        print(f"  {icon} {step_name}: {step['status']} ({step.get('time', 0):.1f}s)")
    print(f"  Summary: {summary_path.resolve()}")


if __name__ == "__main__":
    main()
