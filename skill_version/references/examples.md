# 使用示例

## 示例 1: 基础成本核算

用户："帮我算一下 2025 年 1 月的成本"

AI 操作：
1. 要求用户提供采购入库明细和投入产出明细（最少两张表）
2. Phase 1：创建 config.json → 跑 dry-run → 确认字段映射
3. Phase 2：`python scripts/run_all.py --config config.json --output-dir ./output`
4. 读 `output/_run_summary.json` 确认跑通
5. 读 `output/_index.json` → 转述摘要给用户

## 示例 2: 全年数据

用户："我上传了 2025 年 1-12 月的数据，帮我逐月核算"

AI 操作：
1. 确认 6 张表都包含了 1-12 月的数据（年度/月份字段）
2. 创建 config.json，无需特殊选项
3. 运行 `run_all.py`
4. 读 `output/_index.json` → 12 个月全部有结果
5. 读 `output/cost_fluctuation.json` → 给用户看全年波动趋势

## 示例 3: 用户追问成本波动

用户："2003000161 这个产品的成本为什么波动这么大？"

AI 操作：
1. 打开 `output/cost_fluctuation.json`，在 `fluctuation_ranking` 里找这个产品
2. 打开 `references/analysis_guide.md` → 路径 A
3. 打开 `output/cost_fluctuation.csv`，看各月完工数量和完工单价
4. 打开波动最大月份对应的 `output/{month}/工单明细.csv`，筛选该产品，看完工材料费/人工/制费
5. 回复用户：哪两个月、单价从多少变到多少、什么原因

## 示例 4: 用户追问毛利异常

用户："为什么有些批次毛利率是负的？"

AI 操作：
1. 打开 `output/margin_analysis.json` → 看 `anomalies`
2. 打开 `references/analysis_guide.md` → 路径 B
3. 如果 `total_revenue` 为 0：告诉用户销售数据里没填销售金额
4. 如果有异常批次：打开 `output/{month}/销售成本明细.csv`，看该批次的成本构成
5. 回复用户：多少个异常批次、什么原因

## 示例 5: 字段映射覆盖

用户："我的 Excel 列名比较特殊，'加工单号'而非'工单号'"

AI 操作：
1. 在 config.json 中添加 `mapping_overrides`：
```json
{
  "mapping_overrides": {
    "io": {
      "工单号": "加工单号",
      "材料编码": "原材料编码"
    }
  }
}
```
2. 重新跑 `--dry-run` 确认 missing 清空
3. 运行 `run_all.py`

## 示例输出: _index.json

```json
{
  "run_timestamp": "2026-06-21T15:30:00",
  "pipeline_version": "3.5",
  "total_months": 1,
  "months": ["2025Y01M"],
  "files_used": ["purchase", "io"],
  "options": {
    "calculate_step_method": false,
    "force_calculate": true
  },
  "status": "completed",
  "error_count": 0,
  "warning_count": 0,
  "total_time_seconds": 2.451,
  "month_summaries": {
    "2025Y01M": {
      "matrix_dimension": 85,
      "material_count": 55,
      "order_count": 30,
      "ending_total_cost": 1234567.89,
      "computation_time_seconds": 0.5,
      "has_step_result": false,
      "has_super_result": false,
      "has_sales_result": false
    }
  }
}
```

## 示例输出: _pipeline_log.json (校验通过)

```json
{
  "validation": {
    "sample_month": "2025Y01M",
    "matrix_dimension": 85,
    "W_col_sum_check": { "passed": true, "over_one_count": 0 },
    "self_loop_check": { "passed": true },
    "D_range_check": { "passed": true },
    "sparsity_percent": 0.03,
    "force_normalized": false,
    "force_normalized_count": 0
  },
  "errors": [],
  "warnings": []
}
```
