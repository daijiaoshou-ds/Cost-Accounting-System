# 校验日志——怎么看、怎么跟用户解释

你在 Phase 3 读 `output/_pipeline_log.json` 的时候，照着这个来。

## JSON 长这样

```json
{
  "validation": {
    "sample_month": "2026Y01M",
    "matrix_dimension": 3239,
    "W_col_sum_check": { "passed": true, "threshold": 1.0, "over_one_nodes": [], "over_one_count": 0 },
    "self_loop_check": { "passed": true, "self_loop_nodes": [] },
    "D_range_check": { "passed": true, "min": 0.0, "max": 1.0 },
    "nnz": 3198,
    "sparsity_percent": 0.0305
  },
  "warnings": [],
  "errors": []
}
```

## ⚠️ 铁律：先看 errors，再看 validation

> 如果 `errors` 里有东西，但 `validation` 是空的（或者所有校验项都标 `passed: false` 还带了 `error_message`），说明矩阵构建阶段就直接挂了。**这时候别看 W_col_sum_check 那些了，它们都是被连坐的——根因在 `errors` 和 `validation.error_message` 里。**

## 怎么看每一项

### W 列和校验 (`W_col_sum_check`)

检查有没有物料被领用的总量超过了可供发出的量（>100%）。

- **passed=true** → 没问题。
- **passed=false** → `over_one_nodes` 里列出了哪些物料超领了。
- **跟用户说**："物料 XX 被领用的总量超过了 100%。可能的原因：领用数量填错了，或者 BOM 层级不对导致重复计算，或者期初+采购数量偏小。建议检查一下 XX 的领用记录。"

### 自环校验 (`self_loop_check`)

检查有没有工单领用了自己产出的物料（自己吃自己，矩阵没法求逆）。

- **passed=true** → 没问题。
- **passed=false** → `self_loop_nodes` 里列出了哪些节点自环了。
- **跟用户说**："工单 XX 的材料编码里填了它自己的产品编码，形成了一个循环引用。检查一下 BOM 数据，看看这个工单的材料编码是不是误填成了产品编码。"

### D 矩阵范围校验 (`D_range_check`)

检查完工率是不是在 0~100% 之间。完工率 = 完工数量 / (完工数量 + 在产数量)。

- **passed=true** → 正常。
- **passed=false** → 有工单的完工率越界了。
- **跟用户说**："有工单的完工率不在 0~100% 范围内。检查一下完工数量和在产数量有没有填负值或者异常值。"

### 矩阵维度和稀疏度

- `matrix_dimension` = 物料节点数 + 工单节点数。几千到几万都正常。
- `nnz` = W 矩阵里非零边的数量。
- `sparsity_percent` = 稀疏度。越小越好，通常 < 1%。

告诉用户的时候可以提一句："共 XX 个节点，稀疏度 XX%，求解很快。"

## 实际处理步骤

1. 打开 `_pipeline_log.json`，先扫一眼 `errors` 数组。
2. 如果 `errors` 有东西：
   - 同时 `validation` 是空的 → 直接告诉用户："矩阵构建阶段就出错了：`errors[0]`。"
   - 同时 `validation` 里所有 passed 都是 false → 看 `error_message`，根因在那。
3. 如果 `errors` 是空的：
   - 再看 `validation` 里各项的 `passed`。
   - 全部 true → 恭喜用户。
   - 某几项 false → 按上面的话术解释具体是哪里的问题。
4. 常见的根因消息：
   - `ValueError: 列和超过 1` → 超领，检查 BOM 领用数量。
   - `ValueError: 存在自环` → 工单领了自己，检查 BOM 材料编码。
   - 如果 validation 里每项都写着 `"detail": "矩阵构建阶段失败"` → 别纠结这些字段了，直接看 `error_message`。
