---
name: cost-calculation
description: >
  矩阵成本核算引擎（Matrix Cost Accounting Engine）。
  基于投入产出 Leontief 模型 X = (I-WD)⁻¹F 的企业成本核算系统。
  接收 6 张标准 Excel 表，自动字段匹配，逐月执行平行结转法+超级成本还原，
  输出收发存汇总、工单投入产出明细、成本明细、成本波动分析、毛利率分析。
  Trigger: 成本核算, 算成本, 计算成本, 成本计算, 成本还原, 成本结转,
  收发存, 工单成本, 成本分析, 成本波动, 毛利率, 帮我算成本, 核算成本,
  cost calculation, cost accounting, cost analysis
---

# 矩阵成本核算引擎

基于稀疏矩阵算法的企业成本核算 Skill。用户上传 6 张 Excel 表，AI 执行流水线计算，输出结构化结果供成本分析。

**核心公式**：`X = (I - WD)⁻¹ × F`，其中 W 为物料↔工单二部图流转矩阵，D 为完工率阀门，F 为外部投入。

**默认计算策略**：平行结转法 + 超级成本还原（不做逐步结转法）。

---

## 1. AI 交互协议

当用户说出触发词（如"帮我算成本"）时，按以下流程执行：

### Step 1: 收集文件
要求用户提供以下 **6 张必须的 Excel 表**（缺一不可）：

| 表名 | 必要字段 | 说明 |
|------|---------|------|
| **采购入库明细** | 年度、月份、存货编码、采购数量、采购金额 | 本期采购数据 |
| **投入产出明细** | 年度、月份、工单号、产品编码、材料编码、领用数量、完工数量、在产数量 | BOM + 工单消耗关系 |
| **期初明细** | 年度、月份、存货编码、数量、直接材料、直接人工、制造费用 | 期初结存 |
| **人工制费成本** | 年度、月份、工单号、直接人工、制造费用 | 间接费用分摊 |
| **产品入库明细** | 年度、月份、存货编码、入库数量 | 完工入库口径 |
| **销售数据** | 年度、月份、存货编码、出库单号、销售数量、**销售金额** | 销售成本追溯 + 毛利率计算 |

> ⚠️ 6 张表**全部必须提供**。缺少任何一张表或任何一个必要字段，计算都会失败。

用户可上传多个期间的数据（如 2025年1-12月），系统自动按月份拆分核算。

### Step 2: 创建配置文件
将所有文件路径写入 JSON 配置。**默认选项不需要询问用户：**
- 使用平行结转法（不做逐步结转法）
- 启用超级成本还原（AI 成本分析必需）

```json
{
  "files": {
    "purchase": "/path/to/采购入库明细.xlsx",
    "io": "/path/to/投入产出明细.xlsx",
    "initial": "/path/to/期初明细.xlsx",
    "labor": "/path/to/人工制费成本.xlsx",
    "finished": "/path/to/产品入库明细.xlsx",
    "sales": "/path/to/销售数据.xlsx"
  },
  "options": {
    "calculate_step_method": false,
    "calculate_super_restoration": true
  },
  "output_dir": "./output"
}
```

如果用户列名不标准，添加 `mapping_overrides` 手动指定：
```json
{
  "mapping_overrides": {
    "io": { "工单号": "加工单号", "产品编码": "产出物料" },
    "sales": { "销售金额": "销售收入" }
  }
}
```

### Step 3: 执行核算流水线
```bash
cd skill_version
python scripts/pipeline_cli.py --config config.json --output-dir ./output
```

5 个 SOP 阶段自动执行：S1 字段校验 → S2 ETL清洗 → S3 矩阵校验 → S4 矩阵求解 → S5 生成结果。

### Step 4: 执行分析流水线
核算完成后，AI 需要做两个分析。**这两个分析有对应的脚本，AI 不可手动拼凑数据计算。**

```bash
# 分析1: 月度成本波动
python scripts/cost_fluctuation.py --output-dir ./output

# 分析2: 毛利率分析
python scripts/margin_analysis.py --output-dir ./output
```

### Step 5: 读取分析结果并呈现
先读分析脚本输出的 JSON 获取总览，再按需读取 CSV 明细：

```
output/
├── _index.json                    ← 核算总览
├── _pipeline_log.txt              ← 有问题时读
├── 2025Y01M/ ...                  ← 每月核算结果
├── cost_fluctuation.json          ← ★ 成本波动分析结果
├── cost_fluctuation.csv           ← 波动明细表
├── margin_analysis.json           ← ★ 毛利率分析结果
└── margin_analysis.csv            ← 毛利率明细表
```

### Step 6: 向用户呈现分析结论

**成本波动分析**应覆盖：
- 每个产品月环比成本变动（金额 & 百分比）
- 波动最大的 5 个产品及原因（从超级成本还原追踪到具体维度：期初/采购/工费）
- 异常飙升或骤降的产品预警

**毛利率分析**应覆盖：
- 每个销售批次的毛利率（毛利 / 销售金额）
- 各产品线的平均毛利率
- 毛利率异常的批次（如负毛利、毛利率骤降）及可能原因

---

## 2. 分析脚本速查

| 脚本 | 功能 | 输出 |
|------|------|------|
| `scripts/pipeline_cli.py` | 执行5阶段核算 | `output/{month}/` 每月CSV |
| `scripts/cost_fluctuation.py` | 月度成本波动分析 | `output/cost_fluctuation.{json,csv}` |
| `scripts/margin_analysis.py` | 毛利率分析 | `output/margin_analysis.{json,csv}` |

---

## 3. 输出文件速查

| 文件 | 何时读 |
|------|--------|
| `output/_index.json` | 每次必读，总览 |
| `output/_pipeline_log.txt` | 报错/警告时读 |
| `output/cost_fluctuation.json` | 成本波动分析时读 |
| `output/margin_analysis.json` | 毛利率分析时读 |
| `output/{月份}/summary.json` | 需要单月摘要时读 |
| `output/{月份}/收发存汇总.csv` | 物料级收发存 |
| `output/{月份}/超级还原_完工成本.csv` | 成本波动归因追踪 |
| `output/{月份}/销售成本明细.csv` | 毛利率分析的数据来源 |

**原则：按需读取，不要一次全读。**

---

## 4. 核心理论

详见 `references/methodology.md`。要点：

- **W 矩阵**：物料↔工单二部图。CSR 稀疏存储。
- **D 矩阵**：完工率对角阵。工单 D = 完工/(完工+在产)，物料 D=1.0。
- **F 矩阵**：N×3（料/工/费）。来源=期初+采购+工费。
- **双路径求解**：材料路径 X_mat=(I-WD)⁻¹F_mat；工费路径 X_loh=(I-W)⁻¹F_loh。
- **超级成本还原**：每个原始成本来源独立追踪，精确到维度级别。
- **三大保险**：列和≤1 / 无自环 / D∈[0,1]。

---

## 5. 开发者测试

```bash
cd skill_version
python scripts/pipeline_cli.py --config test_config.json
# 检查 output/_pipeline_log.txt 无错误
# 检查 output/_index.json
# 运行分析
python scripts/cost_fluctuation.py --output-dir ./output
python scripts/margin_analysis.py --output-dir ./output
```

---

## 6. 故障排查

| 症状 | 原因 | 解决 |
|------|------|------|
| "缺少必要字段: X" | 字段匹配失败 | config 中指定 `mapping_overrides` |
| "W矩阵列和超过1" | 有物料超额领用 | 检查领用数量合理性 |
| "W矩阵存在自环" | 工单领用了自己产出的物料 | 检查 BOM 循环引用 |
| "D矩阵存在非法值" | 完工率不在 [0,1] | 检查完工/在产数量 |
| 分析脚本报 "No super restoration data found" | 未开启超级成本还原 | config 中设置 `calculate_super_restoration: true` |
| 毛利率为空 | 缺少销售金额字段 | 检查 sales 表是否包含销售金额列 |
