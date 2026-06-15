# 使用示例

## 示例 1: 基础成本核算

用户："帮我算一下 2025 年 1 月的成本"

AI 操作：
1. 要求用户上传采购入库明细和投入产出明细（最少两张表）
2. 创建 config.json
3. 运行 `python scripts/pipeline_cli.py --config config.json`
4. 读取 `output/_index.json` → 获取摘要
5. 回复用户："1月共有 55 个物料，期末总成本 ¥1,234,567.89"

## 示例 2: 全年数据

用户："我上传了 2025 年 1-12 月的数据，帮我逐月核算"

AI 操作：
1. 确认 6 张表都包含了 1-12 月的数据（年度/月份字段）
2. 创建 config.json，无需特殊选项
3. 运行流水线
4. 读取 `output/_index.json` → 12 个月全部有结果
5. 可选：逐月对比期末总成本趋势

## 示例 3: 带超级成本还原

用户："算 2025 年 1 月的成本，同时做超级成本还原，我想知道每个原始成本来源去了哪些产品"

AI 操作：
1. 确认所有文件齐全
2. config.json 中设置 `"calculate_super_restoration": true`
3. 运行流水线
4. 结果中会多出：`超级还原_完工成本.csv`、`超级还原_TopN汇总.csv`
5. 读取 TopN 汇总：每个产品的前 5 大成本来源

## 示例 4: 销售成本追溯

用户："我有销售数据，帮我算每个批次的销售成本"

AI 操作：
1. 要求上传 sales 表（含出库单号）
2. 运行流水线（不需要额外选项，销售成本会自动计算）
3. 结果中会有 `销售成本明细.csv`
4. 公式：C = B × S × X（批次 × 销售比率 × 成本）

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
2. 运行流水线，系统会优先使用覆盖的映射

## 示例 6: 开发者 CLI 快速测试

```bash
cd skill_version

# 创建 test_config.json:
cat > test_config.json << 'EOF'
{
  "files": {
    "purchase": "../test_data/采购入库.xlsx",
    "io": "../test_data/投入产出.xlsx",
    "initial": "../test_data/期初.xlsx",
    "labor": "../test_data/工费.xlsx"
  },
  "options": {
    "calculate_step_method": false,
    "calculate_super_restoration": false
  },
  "output_dir": "./output"
}
EOF

# 运行
python scripts/pipeline_cli.py --config test_config.json

# 验证
python -c "
import json
idx = json.load(open('output/_index.json'))
print(f'月份: {idx[\"months\"]}')
print(f'状态: {idx[\"status\"]}')
print(f'摘要: {json.dumps(idx[\"month_summaries\"], indent=2, ensure_ascii=False)}')
"
```

## 示例输出: _index.json

```json
{
  "run_timestamp": "2026-06-15T15:30:00",
  "pipeline_version": "3.5",
  "total_months": 1,
  "months": ["2025Y01M"],
  "files_used": ["purchase", "io"],
  "options": {
    "calculate_step_method": false,
    "calculate_super_restoration": false
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
      "has_step_result": false,
      "has_super_result": false,
      "has_sales_result": false
    }
  }
}
```

## 示例输出: summary.json

```json
{
  "year": 2025,
  "month": 1,
  "matrix_dimension": 85,
  "material_count": 55,
  "order_count": 30,
  "ending_total_cost": 1234567.89,
  "max_receipt_amount": 50000.00,
  "available_files": [
    "收发存汇总", "工单明细", "工单产品材料明细", "成本明细"
  ]
}
```
