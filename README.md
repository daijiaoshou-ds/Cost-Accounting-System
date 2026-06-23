# 🧮 矩阵成本核算引擎 v4.0

基于 **Leontief 投入产出模型** 的企业成本核算系统，通过稀疏矩阵算法精确计算产品成本（料、工、费），支持成本还原、销售成本追溯、成本波动分析与毛利率分析。

> **核心公式**：`X = (I - WD)⁻¹ × F`
>
> `W` — 物料-工单流转稀疏矩阵 | `D` — 完工率阀门矩阵 | `F` — 外部投入向量（期初/采购/人工/制费）

---

## 🚀 两套版本

| 版本 | 入口 | 适用场景 |
|------|------|---------|
| **CLI + AI Skill** | `skill_version/` | 批量核算、自动化、AI 对话驱动 |
| **Streamlit Web** | `streamlit_version/` | 交互式界面、可视化探索、手动上传 |

### 方式一：AI Skill（推荐）

直接对 AI 说"帮我算成本"，AI 会自动：

```bash
# AI 会自动执行以下命令，你只需要提供 Excel 文件路径
python scripts/run_all.py --config config.json --output-dir ./output
```

详见 [skill_version/SKILL.md](skill_version/SKILL.md)

### 方式二：Streamlit 界面

```bash
cd streamlit_version
pip install -r requirements.txt
streamlit run app.py
```

浏览器打开 `http://localhost:8501`，上传 Excel 即可。

---

## 📦 安装依赖

### 国内用户（清华镜像，速度快）

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 国际用户

```bash
pip install -r requirements.txt
```

### 依赖清单

```
pandas >= 2.0.0      # 数据处理
numpy >= 1.24.0       # 数值计算
scipy >= 1.10.0       # 稀疏矩阵求解 (spsolve)
polars >= 1.0.0       # 高性能数据引擎
openpyxl >= 3.1.0     # Excel 读写
fastexcel >= 0.12.0   # Excel 加速读取（可选）
```

---

## 📋 数据要求

系统需要以下 Excel 数据文件：

| 文件类型 | 必需字段 | 说明 |
|---------|---------|------|
| **采购入库明细** | 物料编码、采购数量、采购金额 | 本期采购数据 |
| **投入产出明细** | 工单号、产品编码、完工数量、材料编码、领用数量 | 生产消耗关系 |
| **期初结存** | 物料编码、期初金额、期初数量 | 物料期初库存（可选） |
| **工单人工制费** | 工单号、人工、制费 | 间接费用分摊（可选） |
| **产成品入库明细** | 产品编码、入库数量 | 收发存完工口径（可选） |
| **销售数据** | 物料编码、销售数量、销售批次号、销售金额 | 销售成本追溯（可选） |

---

## 🏗️ 技术架构

```
┌──────────────────────────────────────────────────┐
│                   输入层                           │
│   Excel 上传 / config.json 配置 / AI 对话触发       │
└──────────────┬───────────────────────────────────┘
               ▼
┌──────────────────────────────────────────────────┐
│               SOP 流水线 (5 阶段)                   │
│                                                    │
│  S1 字段校验  →  S2 数据清洗  →  S3 矩阵校验        │
│  (自动匹配)     (聚合分组)      (自环/列和/完工率)    │
│                                                    │
│  S4 矩阵求解  →  S5 输出结果                        │
│  (LU 分解)      (CSV + JSON)                       │
└──────────────┬───────────────────────────────────┘
               ▼
┌──────────────────────────────────────────────────┐
│                 分析 & 报告                         │
│                                                    │
│  成本波动分析  +  毛利率分析  +  HTML 可视化报告     │
└──────────────────────────────────────────────────┘
```

### 核心算法

1. **独立节点架构**：物料#车间 + 物料#仓库 单一 W 矩阵，统一求解
2. **稀疏矩阵求解**：全程 CSR 格式，`spsolve` LU 分解 + 正则化 `1e-10·I`
3. **先来后到消耗**：按工单号排序分配，自环跳过，超领自动归一化
4. **平行结转法**：完工率控制材料成本完工/在产分配，工费全额进完工
5. **销售成本追溯**：`C = B × S × X`，批次级成本拆分到销售明细

---

## 📁 项目结构

```
.
├── skill_version/               # CLI + AI Skill 版本
│   ├── SKILL.md                 # AI Skill 定义（AI 读这个）
│   ├── config.json              # 配置文件（AI 检查点）
│   ├── requirements.txt         # Python 依赖
│   ├── test.bat                 # 一键测试脚本
│   ├── scripts/
│   │   ├── run_all.py           # 统一编排入口
│   │   ├── check_env.py         # 环境诊断
│   │   ├── generate_report.py   # HTML 报告生成
│   │   ├── cost_calculation/    # 核心核算引擎
│   │   ├── cost_fluctuation/    # 成本波动分析
│   │   └── margin_analysis/     # 毛利率分析
│   ├── references/              # AI 参考文档
│   └── assets/                  # 报告模板
│
├── streamlit_version/           # Streamlit Web 版本
│   ├── app.py                   # 主界面入口
│   ├── logic.py                 # 核心计算引擎
│   └── docs/                    # 架构文档
│
├── doc/                         # 设计文档
│   ├── theory.md                # 算法原理
│   ├── theory4.0.md             # v4.0 矩阵设计
│   └── PLAN_4.0.md              # 版本规划
│
└── requirements.txt             # 全局依赖
```

---

## 🛠️ 技术栈

- **Python 3.10+**
- **Pandas** — 数据处理与聚合
- **NumPy / SciPy** — 稀疏矩阵运算（`scipy.sparse`, `scipy.sparse.linalg.spsolve`）
- **Polars** — 高性能数据引擎
- **OpenPyXL / FastExcel** — Excel 读写
- **Streamlit** — Web 交互界面
- **Plotly** — 交互式可视化

---

## 🔧 常见问题

| 问题 | 解决 |
|------|------|
| pip 安装慢 | 用清华镜像：`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple` |
| Python 版本太低 | 装 Python 3.10+：`winget install Python.Python.3.13` |
| 结果全是 0 | 检查数据是否完整，销售金额有没有填 |
| 毛利率 N/A | 正常——销售表没填销售金额时纯成本显示 |
| 字段对不上 | 用 `mapping_overrides` 手动映射（AI 会自动处理） |

---

<p align="center">Made with ❤️ for cost accounting</p>
