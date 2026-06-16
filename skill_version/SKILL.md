---
name: cost-calculation
description: >
  矩阵成本核算引擎。基于 Leontief 模型 X=(I-WD)^(-1)F 的企业成本核算 Skill。
  程序负责算和输出结构化结果，AI 负责调用程序、解读结果、与人对话。
  Trigger: 成本核算, 算成本, 计算成本, 成本计算, 成本还原, 成本分析, 成本波动,
  毛利率, 帮我算成本, 核算成本, cost calculation, cost accounting
---

# 矩阵成本核算引擎

> 你的角色：调用命令、读 JSON/CSV、把结果翻译给人听。**不要自己算，让程序算。**

---

## 你的操作流程

```
Phase 0  跑 check_env.py --json → 环境有问题？你帮用户修
Phase 1  跑 --dry-run 预览字段 → 帮用户配 mapping → 生成 config.json
Phase 2  跑 run_all.py → 等结果 → 读 _run_summary.json 看状态
Phase 3  读 _pipeline_log.json → 校验过了？恭喜用户 / 没过？解释为什么
Phase 4  读 _index.json + 分析 JSON → 给用户看总览
Phase 5  用户追问 → 读 analysis_guide.md → 按路径深入 CSV → 回答
```

---

## Phase 0: 环境就绪

你的职责不是"检查环境"，而是**让环境能用**。

### 先诊断

```bash
python scripts/check_env.py --json
```

你会看到类似这样的输出：

```json
{
  "python": {"version": "3.11.0", "ok": true},
  "all_ok": false,
  "missing_packages": ["scipy"]
}
```

### 再修复

`all_ok=false` 时，**你自己动手修，不要推给用户**：

| 看到什么 | 你该做什么 |
|---------|-----------|
| Python < 3.10 | 帮用户装 Python 3.10+。Windows: `winget install Python.Python.3.13` |
| 缺 pip 包 | 直接跑 `pip install -r requirements.txt` |
| 完全空白环境 | 装 Python → 建 venv → pip install，一条龙 |
| 缺 fastexcel | 不碍事，openpyxl 能兜底。可以顺手装上 |

### 最后验证

修完再跑一次 `check_env.py --json`，`all_ok=true` 了就往下走。

---

## Phase 1: 搞定字段映射

用户的 Excel 列名千奇百怪——"实收数量"其实是"采购数量"，"加工单号"其实是"工单号"。你得在跑核算之前把这些对应关系搞清楚。

### 1.1 先看看能自动匹配多少

```bash
python scripts/cost_calculation/pipeline_cli.py --config config.json --dry-run
```

输出长这样：

```json
{
  "purchase": {
    "excel_columns": ["年度","月份","存货编码","实收数量","金额"],
    "matched": {"年度":"年度","采购数量":"实收数量","采购金额":"金额"},
    "missing": []
  },
  "io": {
    "excel_columns": ["年度","月份","加工单号","成品编码","原料编码","领用数量"],
    "matched": {"年度":"年度","月份":"月份","完工数量":"完工数量"},
    "missing": ["工单号", "材料编码"]
  }
}
```

- `missing: []` → 完美，不用管
- `missing: ["工单号", "材料编码"]` → 你得帮用户把这两个字段对上

### 1.2 有 missing？查手册

读 `references/field_reference.md`，里面有每张表的每个标准字段的所有常见别名。比如"工单号"的别名包括"加工单号"、"生产订单"、"MO"——你看用户的 Excel 列名里哪个像，就映射哪个。

### 1.3 生成 config.json

把你确认好的映射关系写成 `config.json`：

```json
{
  "files": {
    "purchase": "/path/to/采购入库.xlsx",
    "io": "/path/to/投入产出.xlsx",
    "initial": "/path/to/期初.xlsx",
    "labor": "/path/to/工费.xlsx",
    "finished": "/path/to/入库.xlsx",
    "sales": "/path/to/销售.xlsx"
  },
  "mapping_overrides": {
    "io": { "工单号": "加工单号", "材料编码": "原料编码" }
  },
  "options": {
    "calculate_step_method": false,
    "calculate_super_restoration": true
  },
  "output_dir": "./output"
}
```

> 默认用平行结转法 + 超级成本还原，别问用户要不要改。

---

## Phase 2: 一键跑起来

```bash
python scripts/run_all.py --config config.json --output-dir ./output
```

这个命令内部干了这些事（你不需要操心）：

```
Step 0: 环境检查
Step 1: 核心核算
         S1 字段校验 → S2 数据清洗 → S3 W/D矩阵校验 ← ⚠️ 这里挂了后面全停
         → S4 矩阵求解 → S5 出 CSV
Step 2: 成本波动分析 (Step 1 成了才跑)
Step 3: 毛利率分析   (Step 1 成了才跑)
Step 4: HTML 报告    (Step 2+3 成了才跑)
```

### 跑完第一件事：读 `output/_run_summary.json`

看 `status` 字段：

| status | 什么意思 | 你该做什么 |
|--------|---------|-----------|
| `completed` | 全绿 | 进 Phase 3 确认校验细节 |
| `failed` | 有步骤挂了 | 看 `core_calculation` 是不是 error → 是就跳到 Phase 3 读错误 |
| `completed_with_warnings` | 跑完了但有非致命问题 | 进 Phase 3 看 warnings |

---

## Phase 3: 解读校验结果

读 `output/_pipeline_log.json`。打开 `references/log_interpretation.md` 对着看。

### 先看 `errors` 数组！

> 如果 `errors` 里面有东西、但 `validation` 是空的（或者全部 `passed: false`），说明矩阵构建阶段就炸了。这时候别看 validation 了，直接读 `errors` 和 `error_message`。

### 校验全通过

告诉用户：
- 矩阵多大（`matrix_dimension` 个节点）
- 稀疏度多少（`sparsity_percent`%）
- "矩阵校验通过，无超领、自环或完工率异常"

### 哪项没通过

| 哪项挂了 | 怎么跟用户说 |
|---------|------------|
| `W_col_sum_check` | "物料 XX 被领用的总量超过了可供发出的量。检查一下 BOM 里的领用数量有没有填错。" |
| `self_loop_check` | "工单 XX 领用了自己产出的物料，循环引用了。看看 BOM 里材料编码是不是填成了产品编码。" |
| `D_range_check` | "有工单的完工率不在 0~100% 之间。检查完工数量和在产数量有没有负值或填错。" |

---

## Phase 4: 看结果

### 先读这三份（必读）

| 文件 | 里面有啥 |
|------|---------|
| `output/_index.json` | 算了几个月、矩阵多大、多少物料、期末总成本多少 |
| `output/cost_fluctuation.json` | 成本最高的前 N 个产品、各月单价、谁波动大谁稳定 |
| `output/margin_analysis.json` | 总毛利、每月/产品线毛利率、哪些批次有问题 |

### 让用户打开 HTML 报告

告诉用户："在浏览器打开 `output/cost_report.html`，里面有完整的图表。"

### 用户问你细节了？再深入 CSV

**不要一口气全读**——用户问什么你读什么：

| 用户问什么 | 你读哪个文件 |
|-----------|------------|
| "这个产品成本为什么波动？" | `output/{month}/超级还原_完工成本.csv` |
| "这个批次毛利为什么低？" | `output/{month}/收发存汇总.csv` |
| "BOM 结构是什么样的？" | `output/{month}/工单产品材料明细.csv` |
| "每个节点的成本多少？" | `output/{month}/成本明细.csv` |

---

## Phase 5: 回答用户的追问

打开 `references/analysis_guide.md`，里面有三条标准路径：

- **用户问成本波动** → 路径 A：读波动 JSON + 超级还原 CSV，追维度变化
- **用户问毛利率异常** → 路径 B：读毛利率 JSON + 销售成本 CSV，查收入/成本构成
- **用户问整体趋势** → 路径 C：跨月对比波动排名 + 毛利率变化

按路径一步步查数据，最后用中文回复用户：波动多少、什么原因、有什么建议。

---

## 出问题了？

| 症状 | 你的行动 |
|------|---------|
| `all_ok=false` | 你帮用户装环境，别只是报告 |
| 字段对不上 | `--dry-run` → 查 `field_reference.md` → 加 `mapping_overrides` |
| `core_calculation` error | 读 `_pipeline_log.json` → 先看 `errors` → 再看 `validation` |
| `status=failed` | 逐步骤看谁挂了，优先修 core_calculation |
| 结果全 0 | 问用户数据是不是完整 |
| 毛利率全 N/A | 正常——sales 表没填销售金额，告诉用户就行 |
| 看到 Polars dtype warning | 无视它。polars 读 Excel 时的正常提示，不影响结果 |

---

## 输出文件速查

| 文件 | 什么时候读 | 怎么读 |
|------|----------|--------|
| `output/_run_summary.json` | Phase 2 跑完 | 必读 |
| `output/_pipeline_log.json` | 有报错/警告 | 必读 |
| `output/_index.json` | Phase 4 | 必读 |
| `output/cost_fluctuation.json` | 用户问波动 | 读摘要 |
| `output/cost_fluctuation.csv` | 用户要明细 | 按需读 |
| `output/margin_analysis.json` | 用户问毛利 | 读摘要 |
| `output/margin_analysis.csv` | 用户要明细 | 按需读 |
| `output/cost_report.html` | Phase 4 | 让用户浏览器打开 |
| `output/{month}/*.csv` | Phase 5 深入追踪 | 按需读 |

## 参考文档速查

| 文档 | 什么时候翻 |
|------|----------|
| `references/field_reference.md` | Phase 1：字段对不上时查别名 |
| `references/log_interpretation.md` | Phase 3：解读校验日志 |
| `references/analysis_guide.md` | Phase 5：用户追问时找分析路径 |
| `references/methodology.md` | 用户问"这个模型是什么原理"时 |
| `references/troubleshooting.md` | 上面故障排查表解决不了时 |
| `references/examples.md` | 不确定怎么做时参考例子 |

---

## 项目文件地图——每个文件是干嘛的

### 你会直接调用的

| 文件 | 什么时候用 |
|------|----------|
| `scripts/check_env.py` | Phase 0：`python scripts/check_env.py --json` 诊断环境 |
| `scripts/cost_calculation/pipeline_cli.py` | Phase 1：`--dry-run` 预览字段 / Phase 2：被 run_all 内部调用 |
| `scripts/run_all.py` | Phase 2：`python scripts/run_all.py --config config.json --output-dir ./output` |
| `scripts/generate_report.py` | 单独重生成 HTML 时：`python scripts/generate_report.py --output-dir ./output` |

### 你需要读的参考文档

| 文件 | 什么时候读 |
|------|----------|
| `references/field_reference.md` | Phase 1：`--dry-run` 发现字段对不上时 |
| `references/log_interpretation.md` | Phase 3：解读 `_pipeline_log.json` 时 |
| `references/analysis_guide.md` | Phase 5：用户追问成本波动/毛利率时 |
| `references/methodology.md` | 用户问矩阵模型原理时 |
| `references/troubleshooting.md` | 故障排查表搞不定时 |
| `references/examples.md` | 不确定操作流程时 |
| `docs/design_overview.md` | 想了解整体架构设计时（选读） |

### 你不需要关心的（程序内部文件）

这些文件由 `run_all.py` 或 `pipeline_cli.py` 内部调用，**你别直接调它们**：

| 文件 | 被谁调用 | 干什么的 |
|------|---------|---------|
| `scripts/cost_calculation/pipeline.py` | pipeline_cli.py | 5 阶段 SOP 流水线 |
| `scripts/cost_calculation/logic.py` | pipeline.py | 矩阵运算 (W/D/F/X) |
| `scripts/cost_calculation/field_utils.py` | pipeline_cli.py | 字段自动匹配 |
| `scripts/cost_fluctuation/main.py` | run_all.py | 成本波动分析 |
| `scripts/margin_analysis/main.py` | run_all.py | 毛利率分析 |

### 基础设施（你不用管）

| 文件 | 干嘛的 |
|------|--------|
| `requirements.txt` | pip 依赖清单，Phase 0 用它装包 |
| `assets/report_template.html` | HTML 报告模板，`generate_report.py` 注入数据用 |
| `scripts/__init__.py` | Python 包标记 |
| `test.bat` / `test_config.json` | 开发者测试用，你不需要碰 |
| `.gitignore` | 排除 output/ 和 __pycache__ |
