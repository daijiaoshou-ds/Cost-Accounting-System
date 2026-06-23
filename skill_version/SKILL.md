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
Phase 0  检查 config.json 是否已配置 → 没有就帮用户创建
Phase 1  跑 check_env.py --json → 环境有问题？你帮用户修
Phase 2  跑 --dry-run 预览字段 → 帮用户配 mapping → 更新 config.json
Phase 3  跑 run_all.py → 等结果 → 读 _run_summary.json 看状态
Phase 4  读 _pipeline_log.json → 校验过了？恭喜用户 / 没过？解释为什么
Phase 5  读 _index.json + 分析 JSON → 给用户看总览
Phase 6  用户追问 → 读 analysis_guide.md → 按路径深入 CSV → 回答
```

---

## Phase 0: 配置就绪 — config.json 有没有？

### 先检查

读一下项目根目录的 `config.json`：

```bash
python -c "import json; c=json.load(open('config.json','r',encoding='utf-8')); print('files:', list(c.get('files',{}).keys())); print('output_dir:', c.get('output_dir','./output'))"
```

### 没有 config.json？

告诉用户："成本核算需要 6 张 Excel 表，请把文件路径发给我。"

**采购入库明细** 、 **投入产出明细** + **期初明细**、**人工制费成本**、**产品入库明细**、**销售数据**

**这几张表都必须要给**

拿到路径后创建 `config.json`：

```json
{
  "files": {
    "purchase": "用户给的采购入库路径",
    "io": "用户给的投入产出路径",
    "initial": "用户给的期初路径（可选）",
    "labor": "用户给的工费路径（可选）",
    "finished": "用户给的入库路径（可选）",
    "sales": "用户给的销售路径（可选）"
  },
  "mapping_overrides": {},
  "options": { "calculate_step_method": false },
  "output_dir": "./output"
}
```

### 有 config.json 但文件路径不对？

`files` 里的路径打不开 → 告诉用户哪个文件路径不对，让他重新给。

---

## Phase 1: 环境就绪

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
| 缺 pip 包 | 直接跑 `pip install -r requirements.txt`。国内网络慢就用清华镜像：`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple` |
| 完全空白环境 | 装 Python → 建 venv → pip install（国内用 `-i https://pypi.tuna.tsinghua.edu.cn/simple`），一条龙 |
| 缺 fastexcel | 不碍事，openpyxl 能兜底。可以顺手装上 |

### 最后验证

修完再跑一次 `check_env.py --json`，`all_ok=true` 了就往下走。

---

## Phase 2: 搞定字段映射

用户的 Excel 列名千奇百怪——"实收数量"其实是"采购数量"，"加工单号"其实是"工单号"。你得在跑核算之前把这些对应关系搞清楚。

### 2.1 跑 dry-run 看自动匹配了多少

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

### 2.2 有 missing？查手册

读 `references/field_reference.md`，里面有每张表的每个标准字段的所有常见别名。比如"工单号"的别名包括"加工单号"、"生产订单"、"MO"——你看用户的 Excel 列名里哪个像，就映射哪个。

### 2.3 更新 config.json

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
    "calculate_step_method": false
  },
  "output_dir": "./output"
}
```

> 默认用平行结转法（`calculate_step_method=false`）。不用问用户要不要改。

---

## Phase 3: 一键跑起来

```bash
python scripts/run_all.py --config config.json --output-dir ./output
```

这个命令内部干了这些事（你不需要操心）：

```
Step 0: 环境检查
Step 1: 核心核算
         S1 字段校验 → S2 数据清洗 → S3 W/D矩阵校验（强制计算默认开启，超领自动归一化）
         → S4 矩阵求解 → S5 出 CSV
Step 2: 成本波动分析 (Step 1 成了才跑)
Step 3: 毛利率分析   (Step 1 成了才跑)
Step 4: HTML 报告    (Step 2+3 成了才跑)
```

### 跑完第一件事：读 `output/_run_summary.json`

看 `status` 字段：

| status | 什么意思 | 你该做什么 |
|--------|---------|-----------|
| `completed` | 全绿 | 进 Phase 4 确认校验细节 |
| `failed` | 有步骤挂了 | 看 `core_calculation` 是不是 error → 是就跳到 Phase 4 读错误 |
| `completed_with_warnings` | 跑完了但有非致命问题 | 进 Phase 4 看 warnings |

---

## Phase 4: 解读校验结果

读 `output/_pipeline_log.json`。

- `validation` 块 → 校验结果，通过/未通过一目了然
- `errors` 数组 → 有东西就是出错了
- `warnings` 数组 → 非致命问题
- `stage_times` → 各阶段耗时

### 先看 `errors` 数组！

> 如果 `errors` 里面有东西、但 `validation` 是空的（或者全部 `passed: false`），说明矩阵构建阶段就炸了。这时候别看 validation 了，直接读 `errors` 和 `error_message`。

### 校验全通过

把 `_pipeline_log.json` 里的 `validation` 块和 `stage_times` 转述给用户。**不用你自己去数字——JSON 里程序已经写好了。**
如果 `validation.W_col_sum_check.over_one_nodes` 非空，说明有物料被归一化了，提醒用户即可，不影响结果准确性。

### 哪项没通过

把 `_pipeline_log.json` 里 `validation` 对应字段和 `errors` 数组贴给用户看，程序已经写清楚了哪个节点、什么问题。你只需要一句话概括严重程度和排查方向：

| 哪项挂了 | 你一句话概括 |
|---------|------------|
| `W_col_sum_check` | 有物料领用超过可供发出，系统已自动归一化。日志里记录了被归一化的节点。 |
| `self_loop_check` | 存在工单领用自己产出的物料（循环引用），日志里列出了自环节点。 |
| `D_range_check` | 有工单的完工率不在 0~100% 之间，日志里标了异常值。 |

---

## Phase 5: 看结果

程序已经把分析结果写好了，你只需要读 JSON 的 `summary` 块，转述给用户。**不要自己重新算。**

### 先读这三份（必读）

| 文件 | 里面有啥 | 你怎么做 |
|------|---------|---------|
| `output/_index.json` | 算了几个月、矩阵多大、多少物料、期末总成本多少 | 读 `month_summaries`，报给用户 |
| `output/cost_fluctuation.json` | 成本最高的前 N 个产品、各月单价、谁波动大谁稳定 | 读 `summary.top_products_summary` 和 `fluctuation_ranking` 前5 |
| `output/margin_analysis.json` | 总毛利、每月/产品线毛利率、哪些批次有问题 | 读 `summary` 里的 `total_*` 和 `anomalies` |

### 让用户打开 HTML 报告

告诉用户："在浏览器打开 `output/cost_report.html`，里面有完整的图表。"

### 用户问你细节了？再深入 CSV

**不要一口气全读**——用户问什么你读什么：

| 用户问什么 | 你读哪个文件 |
|-----------|------------|
| "这个产品成本为什么波动？" | `output/{month}/工单明细.csv` |
| "这个批次毛利为什么低？" | `output/{month}/收发存汇总.csv` |
| "BOM 结构是什么样的？" | `output/{month}/工单产品材料明细.csv` |
| "每个节点的成本多少？" | `output/{month}/成本明细.csv` |

---

## Phase 6: 回答用户的追问

打开 `references/analysis_guide.md`，里面有三条标准路径：

- **用户问成本波动** → 路径 A：读波动 JSON + 工单明细 CSV，追维度变化
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
| `output/_run_summary.json` | Phase 3 跑完 | 必读 |
| `output/_pipeline_log.json` | Phase 4 / 有报错/警告 | 必读 |
| `output/_index.json` | Phase 5 | 必读 |
| `output/cost_fluctuation.json` | 用户问波动 | 读摘要 |
| `output/cost_fluctuation.csv` | 用户要明细 | 按需读 |
| `output/margin_analysis.json` | 用户问毛利 | 读摘要 |
| `output/margin_analysis.csv` | 用户要明细 | 按需读 |
| `output/cost_report.html` | Phase 5 | 让用户浏览器打开 |
| `output/{month}/*.csv` | Phase 6 深入追踪 | 按需读 |

## 参考文档速查

| 文档 | 什么时候翻 |
|------|----------|
| `references/field_reference.md` | Phase 2：字段对不上时查别名 |
| `references/log_interpretation.md` | Phase 4：解读校验日志 |
| `references/analysis_guide.md` | Phase 6：用户追问时找分析路径 |
| `references/methodology.md` | 用户问"这个模型是什么原理"时 |
| `references/troubleshooting.md` | 上面故障排查表解决不了时 |
| `references/examples.md` | 不确定怎么做时参考例子 |

---

## 项目文件地图——每个文件是干嘛的

### 你会直接调用的

| 文件 | 什么时候用 |
|------|----------|
| `scripts/check_env.py` | Phase 1：`python scripts/check_env.py --json` 诊断环境 |
| `scripts/cost_calculation/pipeline_cli.py` | Phase 2：`--dry-run` 预览字段 / Phase 3：被 run_all 内部调用 |
| `scripts/run_all.py` | Phase 3：`python scripts/run_all.py --config config.json --output-dir ./output` |
| `scripts/generate_report.py` | 单独重生成 HTML 时：`python scripts/generate_report.py --output-dir ./output` |

### 你需要读的参考文档

| 文件 | 什么时候读 |
|------|----------|
| `references/field_reference.md` | Phase 2：`--dry-run` 发现字段对不上时 |
| `references/log_interpretation.md` | Phase 4：解读 `_pipeline_log.json` 时 |
| `references/analysis_guide.md` | Phase 6：用户追问成本波动/毛利率时 |
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
