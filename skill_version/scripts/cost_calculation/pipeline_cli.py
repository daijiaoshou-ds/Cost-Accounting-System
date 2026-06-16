#!/usr/bin/env python3
"""
矩阵成本核算引擎 -- CLI 编排脚本
=================================
取代 Streamlit 前端，提供命令行接口供 AI 和开发者使用。

用法:
  # Config 模式 (AI 自动化)
  python scripts/pipeline_cli.py --config config.json

  # 文件参数模式 (开发者快速测试)
  python scripts/pipeline_cli.py --purchase 采购.xlsx --io 投入产出.xlsx \
      --initial 期初.xlsx --labor 工费.xlsx \
      --finished 入库.xlsx --sales 销售.xlsx \
      --step-method --super-restore

输出:
  output/
  ├── _index.json          # 全部月份汇总
  ├── _pipeline_log.txt    # SOP 流水线日志
  └── 2025Y01M/            # 每月独立目录
      ├── summary.json
      ├── 收发存汇总.csv
      └── ...
"""

import argparse
import json
import sys
import os
import time
import traceback
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd

# 确保本目录在 path 中（兼容从任意目录导入）
_self_dir = str(Path(__file__).resolve().parent)
if _self_dir not in sys.path:
    sys.path.insert(0, _self_dir)

from logic import CostCalculator, to_excel, load_and_aggregate, TABLE_SCHEMA
from pipeline import CostPipeline
from field_utils import (
    detect_file_type, smart_match,
    create_edge_table, create_path_table,
)


# ============================================================================
# 常量
# ============================================================================

MONTH_DIR_FMT = "{y}Y{m:02d}M"  # 如 2025Y01M


# ============================================================================
# 参数解析
# ============================================================================

def build_parser():
    p = argparse.ArgumentParser(
        description="矩阵成本核算引擎 -- CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # --- 输入模式 ---
    inp = p.add_argument_group("输入 (二选一)")
    inp.add_argument("--config", type=str, default=None,
                     help="JSON 配置文件路径 (含 files / mapping_overrides / options)")
    inp.add_argument("--purchase", type=str, help="采购入库明细 Excel")
    inp.add_argument("--io", type=str, help="投入产出明细 Excel")
    inp.add_argument("--initial", type=str, help="期初明细 Excel (可选)")
    inp.add_argument("--labor", type=str, help="人工制费成本 Excel (可选)")
    inp.add_argument("--finished", type=str, help="产品入库明细 Excel (可选)")
    inp.add_argument("--sales", type=str, help="销售数据 Excel (可选)")

    # --- 计算选项 ---
    opt = p.add_argument_group("计算选项")
    opt.add_argument("--step-method", action="store_true", default=False,
                     help="启用逐步结转法")
    opt.add_argument("--super-restore", action="store_true", default=False,
                     help="启用超级成本还原")

    # --- 输出 ---
    out = p.add_argument_group("输出")
    out.add_argument("--output-dir", type=str, default="./output",
                     help="结果输出目录 (默认 ./output)")
    out.add_argument("--quiet", action="store_true", default=False,
                     help="静默模式 (只输出 JSON 摘要)")
    out.add_argument("--dry-run", action="store_true", default=False,
                     help="仅预览字段匹配结果，不执行核算")

    return p


# ============================================================================
# 配置加载
# ============================================================================

def load_config_from_json(config_path: str) -> dict:
    """从 JSON 文件读取配置"""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def build_config_from_args(args) -> dict:
    """从 CLI 参数构造配置字典"""
    files = {}
    for key in TABLE_SCHEMA:
        path = getattr(args, key, None)
        if path:
            files[key] = path
    return {
        "files": files,
        "mapping_overrides": {},
        "options": {
            "calculate_step_method": args.step_method,
            "calculate_super_restoration": args.super_restore,
        },
        "output_dir": args.output_dir,
    }


# ============================================================================
# 文件处理
# ============================================================================

def resolve_file(file_ref):
    """
    将文件引用统一为 BytesIO。
    支持: 文件路径 (str) 或已经是 file-like 对象。
    """
    if isinstance(file_ref, str):
        path = Path(file_ref)
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在: {path}")
        with open(path, 'rb') as fh:
            return BytesIO(fh.read())
    # 已经是 file-like
    return file_ref


def build_mappings(config: dict, file_dict: dict) -> dict:
    """
    为每张表构建字段映射。
    1. 先用 smart_match 自动识别
    2. 再用 config["mapping_overrides"] 覆盖
    返回: {table: {原列名: 标准名}}
    """
    overrides = config.get("mapping_overrides", {})
    mappings = {}

    for table, f in file_dict.items():
        if table not in config.get("files", {}):
            continue
        # 读取预览 (只需前几行)
        f.seek(0)
        df_preview = pd.read_excel(f, nrows=5)

        # 自动匹配: auto_map = {标准名: 原列名}
        auto_map = smart_match(df_preview.columns, table)

        # 应用覆盖
        table_overrides = overrides.get(table, {})
        # table_overrides 格式: {标准名: 原列名}
        final_map = {}
        schema = TABLE_SCHEMA[table]
        # 必要字段 + 可选字段 都要参与映射
        schema_cols = list(schema['required']) + list(schema.get('optional', {}).keys())

        for std_col in schema_cols:
            # 优先用覆盖
            if std_col in table_overrides:
                final_map[table_overrides[std_col]] = std_col
            elif std_col in auto_map:
                final_map[auto_map[std_col]] = std_col

        mappings[table] = final_map
        # 注意：不要在这里关闭 f，file_dict 中的 BytesIO 稍后会传给 pipeline

    return mappings


def prepare_file_dict(config: dict) -> dict:
    """将文件路径转换为 BytesIO 对象"""
    file_dict = {}
    for table, file_ref in config.get("files", {}).items():
        file_dict[table] = resolve_file(file_ref)
    return file_dict


# ============================================================================
# 结果输出
# ============================================================================

def write_results(ctx, output_dir: str):
    """
    将 PipelineContext 中的结果写入 output_dir。
    结构:
      output_dir/
        _index.json
        _pipeline_log.txt
        2025Y01M/
          summary.json
          收发存汇总.csv
          工单明细.csv
          ...
    """
    root = Path(output_dir)
    root.mkdir(parents=True, exist_ok=True)

    # ---- 流水线日志 (txt 给人看, json 给 AI 读) ----
    (root / "_pipeline_log.txt").write_text(ctx.log.summary(), encoding='utf-8')
    log_dict = ctx.log.to_dict()
    log_dict["metadata"] = {"pipeline_version": "3.5", "status": "completed" if not ctx.log.has_errors else "error"}
    log_dict["validation"] = getattr(ctx, 'validation', {})
    (root / "_pipeline_log.json").write_text(
        json.dumps(log_dict, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )

    # ---- 月份索引 ----
    index_data = _build_index(ctx)
    (root / "_index.json").write_text(
        json.dumps(index_data, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )

    # ---- 逐月目录 ----
    for year, month in ctx.all_months:
        _write_month_dir(ctx, year, month, root)


def _build_index(ctx) -> dict:
    """构建顶层索引 JSON"""
    month_keys = [MONTH_DIR_FMT.format(y=y, m=m) for y, m in ctx.all_months]
    month_summaries = {}

    for (y, m), calc in ctx.monthly_calc.items():
        key = MONTH_DIR_FMT.format(y=y, m=m)
        result = ctx.monthly_results.get((y, m), {})
        sc = result.get('收发存', pd.DataFrame())
        n = len(calc.all_nodes) if calc.all_nodes else 0
        month_summaries[key] = {
            "matrix_dimension": n,
            "material_count": len(calc.material_nodes) if calc.material_nodes else 0,
            "order_count": len(calc.order_nodes) if calc.order_nodes else 0,
            "ending_total_cost": round(float(sc['期末金额'].sum()), 2) if not sc.empty else 0,
            "computation_time_seconds": round(ctx.log.stage_times.get(f'S4_{y}Y{m:02d}M_核心矩阵运算', 0), 3),
            "has_step_result": '逐步结转_工单明细' in result and not result['逐步结转_工单明细'].empty,
            "has_super_result": '超级还原_完工成本' in result,
            "has_sales_result": '销售成本明细' in result,
        }

    return {
        "run_timestamp": datetime.now().isoformat(),
        "pipeline_version": "3.5",
        "total_months": len(ctx.all_months),
        "months": month_keys,
        "files_used": list(ctx.mapping_dict.keys()),
        "options": ctx.options,
        "status": "completed" if not ctx.log.has_errors else "completed_with_errors",
        "error_count": len(ctx.log.errors),
        "warning_count": len(ctx.log.warnings),
        "total_time_seconds": round(ctx.log.stage_times.get('总耗时', 0), 3),
        "month_summaries": month_summaries,
    }


def _write_month_dir(ctx, year, month, root):
    """写入一个月的结果目录"""
    mdir = root / MONTH_DIR_FMT.format(y=year, m=month)
    mdir.mkdir(parents=True, exist_ok=True)

    result = ctx.monthly_results.get((year, month), {})
    calc = ctx.monthly_calc.get((year, month))

    # ---- CSV 表格 ----
    _write_csv(mdir, result, '收发存汇总', result.get('收发存'))
    _write_csv(mdir, result, '工单明细', result.get('工单明细'))
    _write_csv(mdir, result, '工单产品材料明细', result.get('工单产品材料明细'))
    _write_csv(mdir, result, '成本明细', result.get('成本明细'))

    # 扩展结果
    _write_csv(mdir, result, '逐步结转_工单明细', result.get('逐步结转_工单明细'))
    _write_csv(mdir, result, '逐步结转_工单产品材料明细', result.get('逐步结转_工单产品材料明细'))
    _write_csv(mdir, result, '逐步结转_成本明细', result.get('逐步结转_成本明细'))
    _write_csv(mdir, result, '超级还原_完工成本', result.get('超级还原_完工成本'))
    _write_csv(mdir, result, '超级还原_TopN汇总', result.get('超级还原_TopN汇总'))
    _write_csv(mdir, result, '超级还原_维度定义', result.get('超级还原_维度定义'))
    _write_csv(mdir, result, '超级还原_验证差异', result.get('超级还原_验证差异'))
    _write_csv(mdir, result, '超级还原_销售成本', result.get('超级还原_销售成本'))
    _write_csv(mdir, result, '销售成本明细', result.get('销售成本明细'))

    # 边表 & 路径表
    if calc:
        _write_csv(mdir, result, '边表', create_edge_table(calc))
        _write_csv(mdir, result, '路径表', create_path_table(calc))

    # ---- summary.json ----
    sc = result.get('收发存', pd.DataFrame())
    summary = {
        "year": year,
        "month": month,
        "matrix_dimension": len(calc.all_nodes) if calc and calc.all_nodes else 0,
        "material_count": len(calc.material_nodes) if calc and calc.material_nodes else 0,
        "order_count": len(calc.order_nodes) if calc and calc.order_nodes else 0,
        "ending_total_cost": round(float(sc['期末金额'].sum()), 2) if not sc.empty else 0,
        "max_receipt_amount": round(float(sc['本期收入金额'].max()), 2) if not sc.empty else 0,
        "available_files": [k for k, v in result.items() if isinstance(v, pd.DataFrame) and not v.empty],
    }
    (mdir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )


def _write_csv(mdir, result_dict, name, df):
    """写入 CSV，跳过空 DataFrame"""
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        return
    csv_path = mdir / f"{name}.csv"
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')


# ============================================================================
# 文本摘要
# ============================================================================

def print_summary(ctx, output_dir: str = "./output"):
    """打印人类可读的执行摘要"""
    print()
    print("=" * 60)
    print("  矩阵成本核算引擎 -- 运行报告")
    print("=" * 60)
    print()
    print(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    status = "[OK] 完成" if not ctx.log.has_errors else "[ERR] 有错误"
    print(f"状态: {status} ({len(ctx.log.errors)} 错误, {len(ctx.log.warnings)} 警告)")
    print(f"总耗时: {ctx.log.stage_times.get('总耗时', 0):.2f}s")
    print()

    if ctx.all_months:
        first = ctx.all_months[0]
        last = ctx.all_months[-1]
        print(f"月份范围: {MONTH_DIR_FMT.format(y=first[0], m=first[1])} → "
              f"{MONTH_DIR_FMT.format(y=last[0], m=last[1])} ({len(ctx.all_months)}个月)")
        print()
        print("月份摘要:")
        for y, m in ctx.all_months:
            key = MONTH_DIR_FMT.format(y=y, m=m)
            calc = ctx.monthly_calc.get((y, m))
            result = ctx.monthly_results.get((y, m), {})
            sc = result.get('收发存', pd.DataFrame())
            n = len(calc.all_nodes) if calc and calc.all_nodes else 0
            mat_n = len(calc.material_nodes) if calc and calc.material_nodes else 0
            cost = round(float(sc['期末金额'].sum()), 2) if not sc.empty else 0
            print(f"  {key}: {n}x{n} matrix, {mat_n} materials, ending cost {cost:,.2f}")

    print()
    print(f"输出目录: {Path(output_dir).resolve()}")
    print()
    if ctx.log.has_warnings:
        print(f"⚠ 有 {len(ctx.log.warnings)} 条警告，详见 output/_pipeline_log.txt")
    print("=" * 60)


# ============================================================================
# 可导入的入口函数 (供 run_all.py 调用)
# ============================================================================

def run_from_config(config_path: str, output_dir: str):
    """
    从配置文件执行完整核算流水线。
    返回: (PipelineContext, index_json_path)
    """
    config = load_config_from_json(config_path)
    files = config.get("files", {})
    options = config.get("options", {})

    if not files.get("purchase") or not files.get("io"):
        raise ValueError("至少需要 purchase 和 io 文件")

    file_dict = {}
    for table, file_ref in files.items():
        file_dict[table] = resolve_file(file_ref)

    mapping_dict = build_mappings(config, file_dict)

    pipeline = CostPipeline()
    ctx = pipeline.run(
        file_dict, mapping_dict,
        calculate_step_method=options.get('calculate_step_method', False),
        calculate_super_restoration=options.get('calculate_super_restoration', True),
        stop_on_error=True,
    )

    if ctx.log.has_errors or not ctx.all_months:
        return ctx, None

    write_results(ctx, output_dir)
    index_path = Path(output_dir) / "_index.json"
    return ctx, str(index_path.resolve())


# ============================================================================
# 主入口 (CLI)
# ============================================================================

def main():
    # 强制 UTF-8 输出，避免 Windows 终端中文乱码
    import sys, io
    if sys.platform == "win32":
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
        except Exception:
            pass

    parser = build_parser()
    args = parser.parse_args()

    # ----- 解析配置 -----
    if args.config:
        config = load_config_from_json(args.config)
        # 合并 CLI 覆盖
        if args.output_dir != "./output":
            config["output_dir"] = args.output_dir
        config.setdefault("options", {})
        config["options"].setdefault("calculate_step_method", args.step_method)
        config["options"].setdefault("calculate_super_restoration", args.super_restore)
    else:
        config = build_config_from_args(args)

    files = config.get("files", {})
    options = config.get("options", {})
    output_dir = config.get("output_dir", "./output")

    if not files.get("purchase") or not files.get("io"):
        print("[ERR] 错误: 至少需要 --purchase 和 --io 文件", file=sys.stderr)
        sys.exit(1)

    # ----- Dry-run: 仅预览字段匹配，不执行核算 -----
    if args.dry_run:
        from field_utils import smart_match
        preview = {}
        for table, file_ref in files.items():
            f = resolve_file(file_ref)
            f.seek(0)
            df_cols = pd.read_excel(f, nrows=5).columns.tolist()
            auto = smart_match(df_cols, table)
            ov = config.get("mapping_overrides", {}).get(table, {})
            schema = TABLE_SCHEMA[table]['required']
            matched, missing = {}, []
            for std_col in schema:
                if std_col in ov:
                    matched[std_col] = ov[std_col]
                elif std_col in auto:
                    matched[std_col] = auto[std_col]
                else:
                    missing.append(std_col)
            preview[table] = {"excel_columns": df_cols, "matched": matched, "missing": missing}
        print(json.dumps(preview, ensure_ascii=False, indent=2))
        sys.exit(0)

    # ----- 准备文件 & 映射 -----
    if not args.quiet:
        print("[FILE] 读取文件 & 自动匹配字段...")

    try:
        file_dict = prepare_file_dict(config)
        for table, file_ref in files.items():
            if not args.quiet:
                print(f"  {table}: {Path(str(file_ref)).name}")

        mapping_dict = build_mappings(config, file_dict)

        if not args.quiet:
            print()
            print("[MAP] 字段映射结果:")
            for table, mp in mapping_dict.items():
                if mp:
                    # 显示 {标准名 ← 原列名}
                    items = ", ".join(f"{std}←{orig}" for orig, std in mp.items())
                    print(f"  {table}: {items}")
                else:
                    print(f"  {table}: (空)")
            print()
    except FileNotFoundError as e:
        print(f"[ERR] {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"[ERR] 文件准备失败: {e}", file=sys.stderr)
        if not args.quiet:
            traceback.print_exc()
        sys.exit(1)

    # ----- 执行流水线 -----
    if not args.quiet:
        print("[RUN] 执行 SOP 流水线...")
        print(f"   S1 上传及配置字段 → S2 ETL数据清洗 → S3 矩阵校验")
        print(f"   → S4 核心矩阵运算 → S5 生成结果")
        print()

    pipeline = CostPipeline()
    t0 = time.time()

    try:
        ctx = pipeline.run(
            file_dict, mapping_dict,
            calculate_step_method=options.get('calculate_step_method', False),
            calculate_super_restoration=options.get('calculate_super_restoration', False),
            log_dir=os.path.join(output_dir, 'log'),
            stop_on_error=True,
        )
    except Exception as e:
        print(f"[ERR] 流水线执行失败: {e}", file=sys.stderr)
        if not args.quiet:
            traceback.print_exc()
        sys.exit(1)

    elapsed = time.time() - t0

    # ----- 检查结果 -----
    if ctx.log.has_errors:
        print(f"[ERR] 流水线发现 {len(ctx.log.errors)} 个错误:")
        for e in ctx.log.errors:
            print(f"   - {e}")
        sys.exit(1)

    if not ctx.all_months:
        print("[ERR] 未生成任何月份结果，请检查输入数据", file=sys.stderr)
        sys.exit(1)

    # ----- 写出结果 -----
    if not args.quiet:
        print(f"[WRITE] 写出結果到 {output_dir}/ ...")

    try:
        write_results(ctx, output_dir)
    except Exception as e:
        print(f"[ERR] 结果写出失败: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

    # ----- 打印摘要 -----
    if not args.quiet:
        print_summary(ctx, output_dir)
    else:
        # 静默模式输出 JSON
        index_path = Path(output_dir) / "_index.json"
        if index_path.exists():
            print(index_path.read_text(encoding='utf-8'))

    print(f"[OK] 完成 ({elapsed:.1f}s) -- {len(ctx.all_months)} 个月已处理")
    sys.exit(0)


if __name__ == "__main__":
    main()
