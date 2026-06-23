"""
矩阵成本核算引擎 3.5 — SOP 流水线编排层
===========================================
将成本核算过程拆分为 5 个标准阶段，每阶段有明确的输入/输出/校验，
贯穿全程的 PipelineLog 收集所有警告、错误和性能指标。

SOP Stages:
  S1  上传及配置字段     → 校验字段映射完整性
  S2  ETL 数据清洗       → load_and_aggregate → 按月拆分原子数据
  S3  矩阵校验与日志     → W列和 / DAG无环 / D矩阵范围 / 输出日志
  S4  核心矩阵运算       → 构建 W/D/F → spsolve → WIP
  S5  生成结果           → 收发存/明细 + 可选(逐步结转/超级还原/销售)
"""

import time
import os
import numpy as np
import pandas as pd
from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Optional, Any

from logic import (
    CostCalculator, to_excel, load_and_aggregate, TABLE_SCHEMA
)


# ============================================================================
# PipelineLog — 贯穿流水线的日志收集器
# ============================================================================
@dataclass
class PipelineLog:
    """收集各阶段的警告、错误、性能指标，最后可导出为文本日志"""

    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    metrics: Dict[str, Any] = field(default_factory=dict)
    stage_times: Dict[str, float] = field(default_factory=dict)

    # ---- 记录方法 ----
    def warn(self, msg: str):
        self.warnings.append(msg)

    def error(self, msg: str):
        self.errors.append(msg)

    def metric(self, key: str, value: Any):
        self.metrics[key] = value

    def time_stage(self, stage_name: str, elapsed: float):
        self.stage_times[stage_name] = elapsed

    # ---- 判断 ----
    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0

    # ---- 导出 ----
    def summary(self) -> str:
        """生成人类可读的日志摘要"""
        lines = []
        lines.append("=" * 60)
        lines.append("  矩阵成本核算引擎 — Pipeline 运行日志")
        lines.append("=" * 60)
        lines.append("")

        # 耗时统计
        if self.stage_times:
            lines.append("【各阶段耗时】")
            for name, t in self.stage_times.items():
                lines.append(f"  {name}: {t:.3f}s")
            total = sum(self.stage_times.values())
            lines.append(f"  ─────────────────")
            lines.append(f"  总耗时: {total:.3f}s")
            lines.append("")

        # 关键指标
        if self.metrics:
            lines.append("【关键指标】")
            for k, v in self.metrics.items():
                lines.append(f"  {k}: {v}")
            lines.append("")

        # 警告
        if self.warnings:
            lines.append(f"【警告 ({len(self.warnings)}条)】")
            for w in self.warnings:
                lines.append(f"  ⚠ {w}")
            lines.append("")

        # 错误
        if self.errors:
            lines.append(f"【错误 ({len(self.errors)}条)】")
            for e in self.errors:
                lines.append(f"  ❌ {e}")
            lines.append("")

        if not self.warnings and not self.errors:
            lines.append("✅ 无警告，无错误 — 流水线运行正常")

        lines.append("=" * 60)
        return "\n".join(lines)

    def write_to_file(self, filepath: str):
        """将日志写入文件"""
        os.makedirs(os.path.dirname(filepath) or '.', exist_ok=True)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(self.summary())

    def to_dict(self) -> dict:
        """输出结构化 JSON，供 AI 直接读取解读"""
        return {
            "stage_times": dict(self.stage_times),
            "metrics": dict(self.metrics),
            "warnings": list(self.warnings),
            "errors": list(self.errors),
        }


# ============================================================================
# PipelineContext — 贯穿流水线的上下文
# ============================================================================
@dataclass
class PipelineContext:
    """携带所有中间状态，在各阶段之间传递"""

    # 输入
    file_dict: Dict[str, Any] = field(default_factory=dict)
    mapping_dict: Dict[str, Dict[str, str]] = field(default_factory=dict)
    options: Dict[str, bool] = field(default_factory=dict)

    # S2 产出
    monthly_data: Dict[Tuple[int, int], Dict[str, pd.DataFrame]] = field(default_factory=dict)
    all_months: List[Tuple[int, int]] = field(default_factory=list)

    # S3 产出
    validation_passed: bool = False
    validation: dict = field(default_factory=dict)  # 结构化校验结果

    # S4 产出
    monthly_calc: Dict[Tuple[int, int], CostCalculator] = field(default_factory=dict)
    monthly_X: Dict[Tuple[int, int], np.ndarray] = field(default_factory=dict)

    # S5 产出
    monthly_results: Dict[Tuple[int, int], dict] = field(default_factory=dict)

    # 日志
    log: PipelineLog = field(default_factory=PipelineLog)


# ============================================================================
# CostPipeline — SOP 流水线主类
# ============================================================================
class CostPipeline:
    """
    矩阵成本核算 SOP 流水线

    用法:
        pipeline = CostPipeline()
        ctx = pipeline.run(file_dict, mapping_dict,
                           calculate_step_method=True,
                           force_calculate=True)
        # 访问结果: ctx.monthly_results[(2026, 1)]
        # 查看日志: ctx.log.summary()
    """

    # ================================================================
    # S1: 上传及配置字段
    # ================================================================
    def stage1_upload_config(self, ctx: PipelineContext) -> PipelineContext:
        """
        S1 — 上传及配置字段
        ├── 校验每张表的字段映射完整性
        ├── 检查必须表 (purchase + io) 是否已上传
        └── 记录缺失字段
        """
        t0 = time.time()
        ctx.log.metric('stage1_表数量', len(ctx.file_dict))

        required_tables = ['purchase', 'io']
        for table in required_tables:
            if table not in ctx.file_dict or ctx.file_dict[table] is None:
                ctx.log.error(f"缺少必须表: {table}")
            else:
                ctx.log.metric(f'stage1_{table}', '已上传')

        for table in ctx.file_dict:
            schema = TABLE_SCHEMA[table]
            required_fields = schema['required']
            user_map = ctx.mapping_dict.get(table, {})

            # 用户映射后应覆盖所有必要字段
            mapped_standards = set(user_map.values())
            missing = [f for f in required_fields if f not in mapped_standards]

            if missing:
                ctx.log.warn(f"表 '{table}' 未映射字段: {missing}")
            else:
                ctx.log.metric(f'stage1_{table}_字段数', len(user_map))

        ctx.log.time_stage('S1_上传及配置字段', time.time() - t0)
        return ctx

    # ================================================================
    # S2: ETL 数据清洗
    # ================================================================
    def stage2_etl_clean(self, ctx: PipelineContext) -> PipelineContext:
        """
        S2 — ETL 数据清洗
        ├── Polars/pandas 读取 Excel
        ├── 字段重命名 (用户字段→标准名→内部名)
        ├── 字符串去空格、数值 coerce
        ├── GroupBy 原子化聚合
        └── 按 (年度, 月份) 拆分
        """
        t0 = time.time()

        # 重置文件指针
        for k, f_item in ctx.file_dict.items():
            if hasattr(f_item, 'seek'):
                try:
                    f_item.seek(0)
                except Exception:
                    pass

        try:
            ctx.monthly_data = load_and_aggregate(ctx.file_dict, ctx.mapping_dict)
        except ValueError as e:
            ctx.log.error(f"ETL 失败: {e}")
            ctx.log.time_stage('S2_ETL数据清洗', time.time() - t0)
            return ctx
        except Exception as e:
            ctx.log.error(f"ETL 未知错误: {e}")
            ctx.log.time_stage('S2_ETL数据清洗', time.time() - t0)
            return ctx

        if not ctx.monthly_data:
            ctx.log.error("ETL 后无任何月份数据，请检查年度/月份字段")
            ctx.log.time_stage('S2_ETL数据清洗', time.time() - t0)
            return ctx

        ctx.all_months = sorted(ctx.monthly_data.keys())

        # 统计
        total_rows = 0
        for key, tables in ctx.monthly_data.items():
            for t, df in tables.items():
                total_rows += len(df)
        ctx.log.metric('stage2_月份数', len(ctx.all_months))
        ctx.log.metric('stage2_月份范围', f"{ctx.all_months[0]} → {ctx.all_months[-1]}")
        ctx.log.metric('stage2_聚合后总行数', total_rows)

        ctx.log.time_stage('S2_ETL数据清洗', time.time() - t0)
        return ctx

    # ================================================================
    # S3: 矩阵校验与日志
    # ================================================================
    def stage3_matrix_validate(self, ctx: PipelineContext) -> PipelineContext:
        """
        S3 — 矩阵校验与日志排查
        ├── 取第一个月份构建 W/D 矩阵做抽样校验
        ├── 保险1: W 列和 ≤ 1（防超领）
        ├── 保险2: DAG 无环（防自循环）
        ├── 保险3: D 矩阵范围 [0, 1]
        └── 校验结果写入 ctx.validation（最终输出到 _pipeline_log.json）
        """
        t0 = time.time()

        if not ctx.all_months:
            ctx.log.error("无可校验数据（无月份）")
            ctx.log.time_stage('S3_矩阵校验与日志', time.time() - t0)
            return ctx

        # 抽样第一个月份做校验
        sample_key = ctx.all_months[0]
        sample_data = ctx.monthly_data[sample_key]

        calc = CostCalculator()
        force_calc = ctx.options.get('force_calculate', True)
        try:
            calc.load_data(sample_data)
            calc.calculate(force_calculate=force_calc)  # 触发三道保险
        except ValueError as e:
            error_msg = str(e)
            hint = ""
            if '超领' in error_msg or '强制计算' in error_msg:
                hint = ("检测到物料超领（领用 > 可供发出），成本计算无法继续。"
                        "请使用 --force 参数重试，系统将自动归一化超领列。"
                        "CLI: python pipeline_cli.py --config config.json --force\n"
                        "也可在 config.json 中设置 options.force_calculate = true")
            ctx.log.error(f"矩阵校验失败: {error_msg}")
            if hint:
                ctx.log.warn(hint)
            # 失败时也填充结构化 validation，让 AI 可以解读
            n_nodes = len(calc.all_nodes) if hasattr(calc, 'all_nodes') and calc.all_nodes else 0
            ctx.validation = {
                "sample_month": f"{sample_key[0]}Y{sample_key[1]:02d}M",
                "matrix_dimension": n_nodes,
                "stage": "S3_matrix_build",
                "passed": False,
                "error_type": "ValueError",
                "error_message": error_msg,
                "suggested_action": (
                    "检测到物料超领。建议使用 --force 参数重新运行："
                    "python pipeline_cli.py --config config.json --force"
                ) if ('超领' in error_msg or '强制计算' in error_msg) else "检查输入数据",
                "W_col_sum_check": {"passed": False, "detail": "矩阵构建阶段失败，参见 error_message"},
                "self_loop_check": {"passed": False, "detail": "矩阵构建阶段失败，参见 error_message"},
                "D_range_check": {"passed": False, "detail": "矩阵构建阶段失败，参见 error_message"},
            }
            ctx.log.time_stage('S3_矩阵校验与日志', time.time() - t0)
            return ctx
        except Exception as e:
            error_msg = str(e)
            ctx.log.error(f"矩阵校验未知错误: {error_msg}")
            ctx.validation = {
                "sample_month": f"{sample_key[0]}Y{sample_key[1]:02d}M",
                "stage": "S3_matrix_build",
                "passed": False,
                "error_type": type(e).__name__,
                "error_message": error_msg,
            }
            ctx.log.time_stage('S3_矩阵校验与日志', time.time() - t0)
            return ctx

        # 后续校验
        W = calc.W_matrix
        D_mat = calc.D_matrix
        n = len(calc.all_nodes)

        # 保险1: 列和校验
        col_sums = np.array(W.sum(axis=0)).flatten()
        max_col = float(np.max(col_sums))
        over_one = np.where(col_sums > 1 + 1e-6)[0]
        ctx.log.metric('stage3_矩阵维度', n)
        ctx.log.metric('stage3_W最大列和', f"{max_col:.4f}")

        if len(over_one) > 0:
            bad_nodes = [calc.all_nodes[i] for i in over_one]
            ctx.log.warn(f"W 矩阵列和超过 1 的节点({len(bad_nodes)}个): {bad_nodes[:10]}")
            ctx.log.metric('stage3_超领节点数', len(bad_nodes))
        else:
            ctx.log.metric('stage3_列和校验', '通过 (全部 <=1)')

        # 保险2: DAG 无环
        diag = W.diagonal()
        if np.any(diag > 1e-6):
            bad_idx = np.where(diag > 1e-6)[0]
            bad_nodes = [calc.all_nodes[i] for i in bad_idx]
            ctx.log.error(f"W 矩阵存在自环: {bad_nodes}")
        else:
            ctx.log.metric('stage3_无环校验', '通过')

        # 保险3: D 范围
        if np.any(D_mat < 0) or np.any(D_mat > 1):
            ctx.log.error("D 矩阵存在非法值 (不在 [0,1] 范围内)")
        else:
            ctx.log.metric('stage3_D矩阵范围', '通过')

        # 统计信息
        nnz = W.nnz
        sparsity = nnz / (n * n) * 100 if n > 0 else 0
        ctx.log.metric('stage3_非零边数', nnz)
        ctx.log.metric('stage3_稀疏度', f"{sparsity:.4f}%")

        # ---- 结构化 validation 块 (供 AI 读取) ----
        over_one_nodes = [calc.all_nodes[i] for i in over_one] if len(over_one) > 0 else []
        # 收集强制计算归一化记录（AI 可据此告知用户哪些物料被处理了）
        force_log = getattr(calc, '_force_normalization_log', [])
        ctx.validation = {
            "sample_month": f"{sample_key[0]}Y{sample_key[1]:02d}M",
            "matrix_dimension": n,
            "W_max_col_sum": round(max_col, 6),
            "W_col_sum_check": {
                "passed": len(over_one) == 0,
                "threshold": 1.0,
                "over_one_nodes": over_one_nodes[:20],
                "over_one_count": len(over_one),
            },
            "self_loop_check": {
                "passed": not np.any(diag > 1e-6),
                "self_loop_nodes": [calc.all_nodes[i] for i in np.where(diag > 1e-6)[0]] if np.any(diag > 1e-6) else [],
            },
            "D_range_check": {
                "passed": not (np.any(D_mat < 0) or np.any(D_mat > 1)),
                "min": round(float(np.min(D_mat)), 6),
                "max": round(float(np.max(D_mat)), 6),
            },
            "nnz": int(nnz),
            "sparsity_percent": round(sparsity, 4),
            "force_normalized": bool(force_log),
            "force_normalized_count": len(force_log),
            "force_normalized_nodes": [e['node'] for e in force_log[:20]],
        }

        ctx.validation_passed = not ctx.log.has_errors
        ctx.log.time_stage('S3_矩阵校验与日志', time.time() - t0)
        return ctx

    # ================================================================
    # S4: 核心矩阵运算
    # ================================================================
    def stage4_core_compute(self, ctx: PipelineContext) -> PipelineContext:
        """
        S4 — 核心矩阵运算
        ├── 逐月构建 CostCalculator
        ├── 构建节点 → W矩阵(稀疏CSR) → D阀门 → F外部投入
        ├── spsolve: X_mat = (I-WD)^{-1} F_mat
        ├── spsolve: X_loh = (I-W)^{-1} F_loh
        ├── WIP = W × (1-D) × X_mat
        └── 记录每月的矩阵维度 & 求解耗时
        """
        t0 = time.time()

        if not ctx.all_months:
            ctx.log.error("无月份数据，无法执行矩阵运算")
            ctx.log.time_stage('S4_核心矩阵运算', time.time() - t0)
            return ctx

        total_n = 0
        total_nnz = 0

        for year, month in ctx.all_months:
            month_start = time.time()
            data = ctx.monthly_data[(year, month)]

            calc = CostCalculator()
            calc.load_data(data)

            try:
                calc.calculate(
                    calculate_step_method=ctx.options.get('calculate_step_method', False),
                    force_calculate=ctx.options.get('force_calculate', False),
                )
            except Exception as e:
                ctx.log.error(f"{year}年{month}月 矩阵求解失败: {e}")
                continue

            # 超级成本还原：将 X 拆解到原始成本维度（期初/采购/工费）
            try:
                sr = calc.super_restore(
                    top_products=None,  # 还原全部产品
                    force_calculate=ctx.options.get('force_calculate', False),
                )
                calc.result.update(sr)
                ctx.log.metric(f'stage4_{year}Y{month:02d}M_超级还原维度',
                               calc.performance_log.get('超级还原维度数', 0))
            except Exception as e:
                ctx.log.warn(f"{year}年{month}月 超级还原失败: {e}")

            ctx.monthly_calc[(year, month)] = calc
            ctx.monthly_X[(year, month)] = calc.X_total

            n = len(calc.all_nodes)
            nnz = calc.W_matrix.nnz
            total_n = max(total_n, n)
            total_nnz += nnz

            elapsed = time.time() - month_start
            stage_key = f'S4_{year}Y{month:02d}M_核心矩阵运算'
            ctx.log.time_stage(stage_key, elapsed)
            ctx.log.metric(f'stage4_{year}Y{month:02d}M_维度', n)
            ctx.log.metric(f'stage4_{year}Y{month:02d}M_非零边', nnz)
            ctx.log.metric(f'stage4_{year}Y{month:02d}M_耗时', f'{elapsed:.3f}s')

        ctx.log.metric('stage4_最大矩阵维度', total_n)
        ctx.log.metric('stage4_总非零边数', total_nnz)
        ctx.log.metric('stage4_完成月份数', len(ctx.monthly_calc))

        ctx.log.time_stage('S4_核心矩阵运算', time.time() - t0)
        return ctx

    # ================================================================
    # S5: 生成结果
    # ================================================================
    def stage5_generate_results(self, ctx: PipelineContext) -> PipelineContext:
        """
        S5 — 生成结果
        ├── 收发存汇总表 (物料维度)
        ├── 工单投入产出明细 (汇总表 + 三维明细)
        ├── 成本明细 (节点维度)
        ├── [可选] 逐步结转法
        ├── [可选] 超级成本还原
        ├── [可选] 销售成本
        └── Edge Table / Path Table
        """
        t0 = time.time()

        for year, month in ctx.all_months:
            calc = ctx.monthly_calc.get((year, month))
            if calc is None or calc.result is None:
                ctx.log.warn(f"{year}年{month}月 无计算结果，跳过")
                continue

            ctx.monthly_results[(year, month)] = calc.result

        ctx.log.metric('stage5_结果月份数', len(ctx.monthly_results))

        # 统计关键指标（以第一个月为例）
        if ctx.monthly_results:
            sample_key = ctx.all_months[0]
            sample_res = ctx.monthly_results[sample_key]
            sc = sample_res.get('收发存', pd.DataFrame())
            if not sc.empty:
                ctx.log.metric('stage5_物料数(示例)', len(sc))
                ctx.log.metric('stage5_期末总成本(示例)', f"Y{sc['期末金额'].sum():,.2f}")

        ctx.log.time_stage('S5_生成结果', time.time() - t0)
        return ctx

    # ================================================================
    # run: 一键执行完整 SOP 流水线
    # ================================================================
    def run(self,
            file_dict: Dict[str, Any],
            mapping_dict: Dict[str, Dict[str, str]],
            calculate_step_method: bool = False,
            force_calculate: bool = True,
            stop_on_error: bool = True) -> PipelineContext:
        """
        执行完整的 5 阶段 SOP 流水线。

        Parameters:
        -----------
        file_dict :  {表名: file_like_object}
        mapping_dict : {表名: {原始列名: 标准名}}
        calculate_step_method : 是否计算逐步结转法
        force_calculate : 是否强制计算（超领时自动归一化），默认开启
        stop_on_error : S3 校验不通过时是否终止
        stop_on_error : S3 校验不通过时是否终止

        Returns:
        --------
        PipelineContext — 包含所有中间状态、结果和日志
        """
        total_start = time.time()

        ctx = PipelineContext(
            file_dict=file_dict,
            mapping_dict=mapping_dict,
            options={
                'calculate_step_method': calculate_step_method,
                'force_calculate': force_calculate,
            }
        )

        # ==================== S1 ====================
        ctx = self.stage1_upload_config(ctx)
        if ctx.log.has_errors and stop_on_error:
            ctx.log.time_stage('总耗时', time.time() - total_start)
            return ctx

        # ==================== S2 ====================
        ctx = self.stage2_etl_clean(ctx)
        if ctx.log.has_errors and stop_on_error:
            ctx.log.time_stage('总耗时', time.time() - total_start)
            return ctx

        # ==================== S3 ====================
        ctx = self.stage3_matrix_validate(ctx)
        if ctx.log.has_errors and stop_on_error:
            ctx.log.time_stage('总耗时', time.time() - total_start)
            return ctx

        # ==================== S4 ====================
        ctx = self.stage4_core_compute(ctx)
        if ctx.log.has_errors and stop_on_error:
            ctx.log.time_stage('总耗时', time.time() - total_start)
            return ctx

        # ==================== S5 ====================
        ctx = self.stage5_generate_results(ctx)

        ctx.log.time_stage('总耗时', time.time() - total_start)
        return ctx


# ============================================================================
# 便捷函数
# ============================================================================
def run_pipeline(file_dict, mapping_dict, **kwargs) -> PipelineContext:
    """便捷函数：一键执行 SOP 流水线"""
    pipeline = CostPipeline()
    return pipeline.run(file_dict, mapping_dict, **kwargs)
