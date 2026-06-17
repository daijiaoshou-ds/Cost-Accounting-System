import streamlit as st
import pandas as pd
import numpy as np
from logic import CostCalculator, to_excel, load_and_aggregate, TABLE_SCHEMA, write_debug_log
import time
import re
from io import BytesIO
import zipfile

# 页面配置
st.set_page_config(
    page_title="矩阵成本核算引擎 3.5",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==================== 全局样式 ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1f77b4;
    }
    .upload-section {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        margin-bottom: 1rem;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    div[data-testid="stSelectbox"] {
        margin-bottom: 0.5rem;
    }
    div[data-testid="stSelectbox"] label {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #333 !important;
        margin-bottom: 0.2rem !important;
    }
    div[data-testid="column"] {
        padding: 0 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 必要字段声明 ====================
TABLE_REQUIRED_FIELDS = {
    'initial': ['年度', '月份', '存货编码', '数量', '直接材料', '直接人工', '制造费用', '库存类型'],
    'purchase': ['年度', '月份', '存货编码', '采购数量', '采购金额'],
    'labor': ['年度', '月份', '工单号', '直接人工', '制造费用'],
    'io': ['年度', '月份', '工单号', '产品编码', '材料编码', '领用数量', '完工数量', '在产数量'],
    'finished': ['年度', '月份', '存货编码', '入库数量'],
    'sales': ['年度', '月份', '存货编码', '出库单号', '销售数量'],
}

TABLE_NAMES = {
    'initial': '期初明细',
    'purchase': '采购入库明细',
    'labor': '人工制费成本',
    'io': '投入产出明细',
    'finished': '产品入库明细',
    'sales': '销售数据',
}

TABLE_REQUIRED = {
    'initial': False,
    'purchase': True,
    'labor': False,
    'io': True,
    'finished': False,
    'sales': False,
}


# ==================== 工具函数 ====================
def detect_file_type(filename):
    """根据文件名检测文件类型"""
    name_lower = filename.lower()
    if any(kw in name_lower for kw in ['产成品入库', '产品入库', '成品入库', '完工入库', 'finished']):
        return 'finished'
    elif any(kw in name_lower for kw in ['采购入库', '采购', 'purchase']):
        return 'purchase'
    elif any(kw in name_lower for kw in ['投入', '产出', 'io', '工单']):
        return 'io'
    elif any(kw in name_lower for kw in ['期初', 'initial', '期初结存']):
        return 'initial'
    elif any(kw in name_lower for kw in ['销售', '出库', '发货', 'sales', 'sell']):
        return 'sales'
    elif any(kw in name_lower for kw in ['人工', '制费', '费用', 'labor', 'cost']):
        return 'labor'
    return None


def smart_match(columns, file_type):
    """智能字段匹配 — 3.5版本：支持年度/月份及新字段名"""
    cols = list(columns)
    result = {}

    patterns = {
        'initial': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8), (r'^月度$', 7)],
            '存货编码': [(r'存货.*编码$', 10), (r'物料.*编码$', 10), (r'料号$', 9), (r'品号$', 8),
                       (r'物料代码', 7), (r'材料编码', 6), (r'存货代码', 7), (r'编码$', 5)],
            '数量': [(r'期初.*数量$', 10), (r'期初.*库存$', 9), (r'结存数量', 8), (r'^数量$', 5)],
            '直接材料': [(r'直接材料$', 10), (r'期初.*材料$', 9), (r'^材料$', 7),
                       (r'期初材料金额', 8), (r'材料成本$', 6)],
            '直接人工': [(r'直接人工$', 10), (r'^人工$', 9), (r'人工费$', 9),
                       (r'期初.*人工$', 8), (r'人工金额', 7)],
            '制造费用': [(r'制造费用$', 10), (r'^制费$', 9), (r'间接费用', 7),
                       (r'期初.*制费$', 8), (r'制造成本', 6)],
            '库存类型': [(r'库存类型$', 10), (r'库别$', 9), (r'库存类别', 8),
                       (r'仓库类型$', 8), (r'存储位置', 7), (r'库位', 6)],
        },
        'purchase': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8)],
            '存货编码': [(r'存货.*编码$', 10), (r'物料.*编码$', 10), (r'料号$', 9), (r'品号$', 8),
                       (r'物料代码', 7), (r'存货代码', 7), (r'编码$', 5)],
            '采购数量': [(r'采购.*数量$', 10), (r'入库.*数量$', 9), (r'实收数量', 8), (r'采购数', 8), (r'^数量$', 3)],
            '采购金额': [(r'采购.*金额$', 10), (r'采购.*成本$', 9), (r'实收金额', 8),
                       (r'采购.*价值', 8), (r'^金额$', 3)],
        },
        'labor': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8)],
            '工单号': [(r'工单号?$', 10), (r'加工单号$', 10), (r'生产订单$', 9), (r'订单号$', 8),
                     (r'工单', 7), (r'^MO', 6)],
            '直接人工': [(r'直接人工$', 10), (r'^人工$', 9), (r'人工费$', 9),
                       (r'人工金额', 7), (r'直接工资', 8)],
            '制造费用': [(r'制造费用$', 10), (r'^制费$', 9), (r'间接费用', 7),
                       (r'制造成本', 6), (r'费用金额', 5)],
        },
        'io': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8)],
            '工单号': [(r'工单号?$', 10), (r'加工单号$', 10), (r'生产订单$', 9), (r'订单号$', 8),
                     (r'工单', 7), (r'^MO', 6)],
            '产品编码': [(r'产品.*编码$', 10), (r'产成品', 9), (r'成品编码', 8), (r'产品代码', 7), (r'产出编码', 6)],
            '材料编码': [(r'材料.*编码$', 10), (r'物料.*编码', 9), (r'领料编码', 8), (r'材料代码', 7), (r'原料编码', 7)],
            '领用数量': [(r'领用.*数量$', 10), (r'领料.*数量$', 9), (r'消耗.*数量$', 8),
                       (r'用量$', 7), (r'领用数', 6), (r'实发数量', 8)],
            '完工数量': [(r'完工.*数量$', 10), (r'产出.*数量$', 9), (r'产量$', 8), (r'完工数', 7), (r'合格数量', 7)],
            '在产数量': [(r'在产.*数量$', 10), (r'在产品.*数量$', 9), (r'在制品.*数量$', 8),
                       (r'在产数', 7), (r'在制品', 6), (r'未完工数量', 7)],
        },
        'finished': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8)],
            '存货编码': [(r'存货.*编码$', 10), (r'产品.*编码$', 10), (r'物料.*编码$', 10),
                       (r'产成品.*编码$', 9), (r'成品编码$', 8), (r'料号$', 9),
                       (r'品号$', 8), (r'产品代码', 6), (r'编码$', 5)],
            '入库数量': [(r'入库.*数量$', 10), (r'完工.*数量$', 9), (r'入库数', 8), (r'^数量$', 5), (r'入库量', 7)],
        },
        'sales': {
            '年度': [(r'^年度$', 10), (r'^年$', 9), (r'year', 8)],
            '月份': [(r'^月份$', 10), (r'^月$', 9), (r'^期间$', 8), (r'month', 8)],
            '存货编码': [(r'存货.*编码$', 10), (r'物料.*编码$', 10), (r'产品.*编码$', 10),
                       (r'料号$', 9), (r'品号$', 8), (r'存货代码', 7), (r'编码$', 5)],
            '出库单号': [(r'出库单号$', 10), (r'出库单$', 9), (r'发货单号$', 9),
                       (r'批次号?$', 8), (r'订单号$', 7), (r'签收单号$', 8),
                       (r'单号$', 6), (r'批次$', 7)],
            '销售数量': [(r'销售.*数量$', 10), (r'出库.*数量$', 9), (r'发货.*数量$', 8),
                       (r'出库数$', 7), (r'销量$', 8), (r'^数量$', 5)],
        },
    }

    patterns_for_type = patterns.get(file_type, {})

    for standard_col, regex_list in patterns_for_type.items():
        best_match = None
        best_score = 0

        for col in cols:
            col_str = str(col).strip()
            for pattern, score in regex_list:
                if re.search(pattern, col_str, re.IGNORECASE):
                    if score > best_score:
                        best_score = score
                        best_match = col
                    break

        if best_match:
            result[standard_col] = best_match
            cols.remove(best_match)

    return result


def render_field_mapping(file_key, required_fields, auto_map, df_columns):
    """渲染字段映射界面 — 垂直布局"""
    final_map = {}

    for std_col in required_fields:
        matched_col = auto_map.get(std_col)
        col1, col2 = st.columns([1, 2])
        with col1:
            st.markdown(f"**{std_col}**")
        with col2:
            options = ['(跳过)'] + list(df_columns)
            default_idx = options.index(matched_col) if matched_col in options else 0
            selected = st.selectbox(
                f"map_{file_key}_{std_col}",
                options,
                index=default_idx,
                label_visibility="collapsed",
                key=f"map_{file_key}_{std_col}"
            )
            if selected != '(跳过)':
                final_map[selected] = std_col
        st.markdown("<hr style='margin: 0.5rem 0; opacity: 0.3;'>", unsafe_allow_html=True)

    return final_map


def format_time(seconds):
    return f"{seconds:.3f}s"


def render_time_metrics(perf, total_time):
    st.markdown("### ⚡ 计算性能")
    n = perf.get('矩阵维度', 0)
    if n > 0:
        st.caption(f"矩阵维度: {n}×{n}")

    display_keys = ['数据清洗', '计算数量', '矩阵求解']
    stages = [(k, perf[k]) for k in display_keys if k in perf]

    if not stages:
        return

    cols = st.columns(len(stages) + 1)
    for i, (name, value) in enumerate(stages):
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size: 0.9rem; color: #666;">{name}</div>
                <div style="font-size: 1.6rem; font-weight: bold; color: #1f77b4;">{format_time(value)}</div>
            </div>
            """, unsafe_allow_html=True)

    with cols[-1]:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #28a745;">
            <div style="font-size: 0.9rem; color: #666;">总耗时</div>
            <div style="font-size: 1.6rem; font-weight: bold; color: #28a745;">{format_time(total_time)}</div>
        </div>
        """, unsafe_allow_html=True)


def create_edge_table(calc):
    """生成边表：相邻流转关系"""
    try:
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            return None
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes
        rows = []
        edge_id = 0
        W_coo = W.tocoo()
        for i, j, w in zip(W_coo.row, W_coo.col, W_coo.data):
            if w > 0.001:
                edge_id += 1
                source = all_nodes[j]
                target = all_nodes[i]
                weight = float(w)
                is_material_to_order = (source in material_nodes) and (target in order_nodes)
                is_order_to_material = (source in order_nodes) and (target in material_nodes)
                consume_ratio = f"{weight:.2%}" if is_material_to_order else "—"
                output_ratio = f"{weight:.2%}" if is_order_to_material else "—"
                rows.append({
                    '边ID': f"E{edge_id:03d}",
                    '起点': source,
                    '起点类型': '工单' if source in order_nodes else '物料',
                    '终点': target,
                    '终点类型': '工单' if target in order_nodes else '物料',
                    '消耗比例': consume_ratio,
                    '产出比例': output_ratio,
                })
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    except Exception:
        return None


def create_path_table(calc):
    """生成路径表：完整的根到叶路径"""
    try:
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            return None
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes

        out_edges = {node: [] for node in all_nodes}
        in_edges = {node: [] for node in all_nodes}
        edge_weight = {}

        W_coo = W.tocoo()
        for i, j, w in zip(W_coo.row, W_coo.col, W_coo.data):
            if w > 0.001:
                source = all_nodes[j]
                target = all_nodes[i]
                weight = float(w)
                out_edges[source].append(target)
                in_edges[target].append(source)
                edge_weight[(source, target)] = weight

        roots = [node for node in material_nodes if not in_edges[node]]
        if not roots:
            roots = list(material_nodes)
        leaves = [node for node in material_nodes if not out_edges[node]]

        all_paths = []

        def dfs(current, path, weights, visited):
            if current in visited:
                return
            new_path = path + [current]
            new_visited = visited | {current}
            if current in leaves and len(new_path) >= 2:
                all_paths.append((new_path, weights))
                return
            for next_node in out_edges[current]:
                w = edge_weight.get((current, next_node), 1.0)
                dfs(next_node, new_path, weights + [w], new_visited)

        for root in roots:
            dfs(root, [], [], set())

        max_layers = max(len(path) for path, _ in all_paths) if all_paths else 0

        rows = []
        for idx, (path, weights) in enumerate(all_paths, 1):
            row = {'路径ID': f"P{idx:03d}"}
            for i in range(1, max_layers + 1):
                row[f'第{i}层'] = path[i - 1] if i <= len(path) else ''
            final_product = None
            for node in reversed(path):
                if node in material_nodes:
                    final_product = node
                    break
            row['最终成品'] = final_product if final_product else path[-1]
            consume_ratios = []
            for i in range(len(path) - 1):
                src, dst = path[i], path[i + 1]
                w = edge_weight.get((src, dst), 1.0)
                consume_ratios.append(f"{w:.0%}")
            row['消耗关系'] = ' × '.join(consume_ratios) if consume_ratios else '—'
            rows.append(row)

        if rows:
            df = pd.DataFrame(rows)
            layer_cols = [f'第{i}层' for i in range(1, max_layers + 1)]
            return df[['路径ID'] + layer_cols + ['最终成品', '消耗关系']]
        return pd.DataFrame()
    except Exception:
        return None


# ==================== Session State ====================
def init_session_state():
    defaults = {
        'current_page': 'home',
        'monthly_results': {},
        'monthly_perf': {},
        'monthly_total_time': {},
        'monthly_calc': {},
        'all_months': [],
        'file_dict': {},
        'mapping_dict': {},
        'uploaded_dfs': {},
        'calculation_done': False,
        'restore_result': None,
        'restore_calc': None,
        'restore_total_time': None,
        'step_cost_file': None,
        'step_cost_mapping': None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session_state()


# ==================== 首页 ====================
def render_home():
    st.markdown('<div class="main-header">🧮 矩阵成本核算引擎 3.5</div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align: center; color: #666; margin-bottom: 3rem;">
        基于线性代数的企业成本核算系统 — 支持全期间数据上传，逐月核算
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    padding: 2rem; border-radius: 15px; color: white; text-align: center;">
            <div style="font-size: 3rem; margin-bottom: 1rem;">📊</div>
            <div style="font-size: 1.5rem; font-weight: bold; margin-bottom: 0.5rem;">多期间成本核算</div>
            <div style="font-size: 0.9rem; opacity: 0.9; margin-bottom: 1.5rem;">
                上传全期间数据（含年度/月份），逐月自动核算各产品的料工费成本
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("进入成本核算", key="btn_cost", use_container_width=True):
            st.session_state.current_page = 'cost'
            st.rerun()

    with col2:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                    padding: 2rem; border-radius: 15px; color: white; text-align: center;">
            <div style="font-size: 3rem; margin-bottom: 1rem;">🔄</div>
            <div style="font-size: 1.5rem; font-weight: bold; margin-bottom: 0.5rem;">成本还原</div>
            <div style="font-size: 0.9rem; opacity: 0.9; margin-bottom: 1.5rem;">
                基于逐步结转法数据，还原各产品的真实料工费结构
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("进入成本还原", key="btn_restore", use_container_width=True):
            st.session_state.current_page = 'restore'
            st.rerun()

    with st.expander("📖 技术原理"):
        st.markdown("""
        ### 核心公式

        系统基于投入产出分析（Leontief 模型）：

        $$X = (I - WD)^{-1} \\times F$$

        - **W**: 物料与工单之间的消耗/产出流转矩阵
        - **D**: 完工率阀门矩阵（材料按完工率分配，工费全额归完工）
        - **F**: 外部投入矩阵（期初、采购、人工、制费）
        - **X**: 各节点的总成本（料、工、费三列独立）

        ### 3.5 新特性

        - ✅ 全期间数据上传（年度+月份字段）
        - ✅ Polars 加速 Excel 读取
        - ✅ 严格字段校验 + 自动 GroupBy 原子化
        - ✅ 逐月核算 + 批量导出
        """)


def _handle_manual_mapping(key, required_fields, df_preview):
    """手动配置字段映射"""
    st.warning("⚠️ 未识别到匹配字段，请手动配置")
    cols_iter = st.columns(3)
    manual_map = {}
    for i, field in enumerate(required_fields):
        with cols_iter[i % 3]:
            selected = st.selectbox(
                field, ['(跳过)'] + list(df_preview.columns),
                key=f"cost_manual_{key}_{field}"
            )
            if selected != '(跳过)':
                manual_map[selected] = field
    if manual_map:
        st.session_state.mapping_dict[key] = manual_map
        st.session_state.file_dict[key] = df_preview
        st.session_state.uploaded_dfs[key] = df_preview


# ==================== 成本核算页面 ====================
def render_cost_accounting():
    st.markdown("### 📊 多期间成本核算")

    if st.button("← 返回首页", key="back_home_cost"):
        st.session_state.current_page = 'home'
        st.rerun()

    st.divider()

    # ==================== 步骤1: 上传数据 ====================
    st.markdown("#### 📁 上传数据文件")
    st.caption("每个表必须包含 **年度** 和 **月份** 字段，系统按月份自动拆分并逐月核算")

    # ==================== 批量上传（可选）====================
    with st.expander("📦 批量上传（可选，支持ZIP或多选Excel）", expanded=False):
        st.info("支持上传 ZIP 压缩包或同时选择多个 Excel 文件，系统根据文件名自动识别文件类型")

        batch_files = st.file_uploader(
            "上传文件（可多选或ZIP）",
            type=['xlsx', 'xls', 'zip'],
            accept_multiple_files=True,
            key="batch_upload"
        )

        auto_detected = {}  # {file_type: (file_obj, type_name, filename)}

        if batch_files:
            for f_batch in batch_files:
                try:
                    if f_batch.name.endswith('.zip'):
                        try:
                            with zipfile.ZipFile(f_batch) as z:
                                for name in z.namelist():
                                    if name.endswith(('.xlsx', '.xls')):
                                        file_type, type_name = detect_file_type(name), TABLE_NAMES.get(detect_file_type(name), '')
                                        if file_type:
                                            excel_data = BytesIO(z.read(name))
                                            excel_data.name = name
                                            auto_detected[file_type] = (excel_data, type_name, name)
                                            st.success(f"📦 ZIP 中找到: {type_name} - {name}")
                        except Exception as e_zip:
                            st.error(f"无法读取 ZIP 文件 {f_batch.name}: {e_zip}")
                    else:
                        file_type = detect_file_type(f_batch.name)
                        type_name = TABLE_NAMES.get(file_type, '')
                        if file_type:
                            auto_detected[file_type] = (f_batch, type_name, f_batch.name)
                            st.success(f"✓ 识别为 {type_name}: {f_batch.name}")
                        else:
                            st.warning(f"⚠️ 无法识别文件类型: {f_batch.name}")
                except Exception as e_batch:
                    st.error(f"处理文件 {f_batch.name} 时出错: {e_batch}")

    st.divider()

    # ==================== 单个文件上传 ====================
    st.markdown("#### 📄 逐个上传（精确控制）")

    file_configs = [
        ('purchase', '🔴 采购入库明细（必须）'),
        ('io', '🔴 投入产出明细（必须）'),
        ('initial', '⚪ 期初明细（可选）'),
        ('labor', '⚪ 人工制费成本（可选）'),
        ('finished', '⚪ 产品入库明细（可选）'),
        ('sales', '⚪ 销售数据（可选）'),
    ]

    for key, title in file_configs:
        required_fields = TABLE_REQUIRED_FIELDS[key]

        # 检查是否有批量上传的自动识别结果
        auto_file = None
        auto_name = None
        if key in auto_detected:
            auto_file, _, auto_name = auto_detected[key]
            if hasattr(auto_file, 'seek'):
                auto_file.seek(0)

        with st.expander(title, expanded=(TABLE_REQUIRED[key] and not auto_file)):
            if auto_file:
                st.success(f"✓ 已从批量上传识别: {auto_name}")
                # 从批量上传结果中读取
                try:
                    auto_file.seek(0)
                    df_preview = pd.read_excel(auto_file)
                    st.caption(f"✓ 读取成功: {len(df_preview)} 行 × {len(df_preview.columns)} 列")

                    with st.expander("📋 数据预览", expanded=False):
                        st.dataframe(df_preview.head(5), use_container_width=True, height=200)

                    auto_map = smart_match(df_preview.columns, key)
                    if auto_map:
                        matched_count = len(auto_map)
                        missing_count = len(required_fields) - matched_count
                        if missing_count > 0:
                            st.warning(f"自动匹配 {matched_count}/{len(required_fields)} 个字段，缺 {missing_count} 个，请确认或调整")
                        else:
                            st.success(f"✓ 自动匹配全部 {matched_count} 个字段")
                        final_map = render_field_mapping(f"cost_{key}", required_fields, auto_map, df_preview.columns)
                        if final_map:
                            st.session_state.mapping_dict[key] = final_map
                            st.session_state.file_dict[key] = auto_file
                            st.session_state.uploaded_dfs[key] = df_preview
                    else:
                        _handle_manual_mapping(key, required_fields, df_preview)
                except Exception as e:
                    st.error(f"❌ 读取失败: {e}")
            else:
                f_new = st.file_uploader(
                    f"上传 {title} Excel 文件",
                    type=['xlsx', 'xls'],
                    key=f"cost_{key}"
                )

                if f_new:
                    try:
                        f_new.seek(0)
                        df_preview = pd.read_excel(f_new)
                        st.caption(f"✓ 读取成功: {len(df_preview)} 行 × {len(df_preview.columns)} 列")

                        with st.expander("📋 数据预览", expanded=False):
                            st.dataframe(df_preview.head(5), use_container_width=True, height=200)

                        auto_map = smart_match(df_preview.columns, key)

                        if auto_map:
                            matched_count = len(auto_map)
                            missing_count = len(required_fields) - matched_count
                            if missing_count > 0:
                                st.warning(f"自动匹配 {matched_count}/{len(required_fields)} 个字段，缺 {missing_count} 个，请确认或调整")
                            else:
                                st.success(f"✓ 自动匹配全部 {matched_count} 个字段")
                            final_map = render_field_mapping(f"cost_{key}", required_fields, auto_map, df_preview.columns)
                            if final_map:
                                st.session_state.mapping_dict[key] = final_map
                                st.session_state.file_dict[key] = f_new
                                st.session_state.uploaded_dfs[key] = df_preview
                            else:
                                st.warning("⚠️ 至少需要映射一个字段")
                        else:
                            _handle_manual_mapping(key, required_fields, df_preview)

                    except Exception as e:
                        st.error(f"❌ 读取失败: {e}")
                elif TABLE_REQUIRED[key] and not auto_file:
                    st.info(f"⬆️ 请上传 {title} 文件")

    # ==================== 步骤2: 计算选项与执行 ====================
    st.divider()
    st.markdown("#### 🚀 执行计算")

    ready = 'purchase' in st.session_state.file_dict and 'io' in st.session_state.file_dict

    if not ready:
        st.warning("⚠️ 请至少上传【采购入库明细】和【投入产出明细】文件")
        return

    col_opt1, col_opt2, col_opt3, col_opt4 = st.columns([1, 1, 1, 1])
    with col_opt1:
        calculate_step_method = st.checkbox("📊 逐步结转法", value=False)
    with col_opt2:
        calculate_super_restoration = st.checkbox("🔬 超级成本还原", value=False)
    with col_opt3:
        force_calculate = st.checkbox("⚠️ 强制计算", value=False,
            help="超领物料自动归一化至1.0，单价乘以矫正系数")
    with col_opt4:
        debug_mode = st.checkbox("🔍 调试日志", value=False,
            help="输出debug_logs/log_{year}Y{month}M.txt")

    if st.button("🚀 执行成本核算", type="primary", use_container_width=True):
        with st.spinner("正在读取并聚合数据..."):
            try:
                file_dict = st.session_state.file_dict
                mapping_dict = st.session_state.mapping_dict

                # 重置文件指针
                for k, f_item in file_dict.items():
                    if hasattr(f_item, 'seek'):
                        f_item.seek(0)

                monthly_data = load_and_aggregate(file_dict, mapping_dict)

                if not monthly_data:
                    st.error("❌ 未找到任何月份数据，请检查文件中的年度/月份字段")
                    return

                all_months = sorted(monthly_data.keys())
                st.session_state.all_months = all_months
                st.session_state.monthly_results = {}
                st.session_state.monthly_perf = {}
                st.session_state.monthly_total_time = {}
                st.session_state.monthly_calc = {}

                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, (year, month) in enumerate(all_months):
                    status_text.text(f"正在计算 {year}年{month}月... ({i+1}/{len(all_months)})")

                    start = time.time()
                    calc = CostCalculator()
                    calc.load_data(monthly_data[(year, month)])
                    result = calc.calculate(
                        calculate_step_method=calculate_step_method,
                        calculate_super_restoration=calculate_super_restoration,
                        force_calculate=force_calculate
                    )
                    total_time = time.time() - start

                    st.session_state.monthly_results[(year, month)] = result
                    st.session_state.monthly_perf[(year, month)] = calc.get_performance()
                    st.session_state.monthly_total_time[(year, month)] = total_time
                    st.session_state.monthly_calc[(year, month)] = calc

                    progress_bar.progress((i + 1) / len(all_months))

                # 调试日志输出
                if debug_mode:
                    import os
                    log_dir = "debug_logs"
                    os.makedirs(log_dir, exist_ok=True)
                    for (year, month), calc in st.session_state.monthly_calc.items():
                        log_path = os.path.join(log_dir, f"log_{year}Y{month:02d}M.txt")
                        try:
                            write_debug_log(calc, log_path)
                        except Exception as log_e:
                            st.warning(f"调试日志写入失败 ({year}年{month}月): {log_e}")
                    st.caption(f"调试日志已输出到 debug_logs/")

                status_text.text(f"✅ 全部 {len(all_months)} 个月计算完成！")
                st.session_state.calculation_done = True
                st.success(f"✅ 共完成 {len(all_months)} 个月的核算")

            except Exception as e:
                st.error(f"❌ 计算失败: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # ==================== 步骤3: 显示结果 ====================
    if st.session_state.calculation_done and st.session_state.all_months:
        st.divider()
        st.markdown("#### 📊 核算结果")

        all_months = st.session_state.all_months

        # 月份选择器
        month_labels = [f"{y}年{m}月" for y, m in all_months]
        selected_label = st.selectbox("选择查看月份", month_labels, key="month_selector")
        selected_idx = month_labels.index(selected_label)
        year, month = all_months[selected_idx]

        result = st.session_state.monthly_results[(year, month)]
        perf = st.session_state.monthly_perf[(year, month)]
        total_time = st.session_state.monthly_total_time[(year, month)]
        calc = st.session_state.monthly_calc[(year, month)]

        # 性能指标
        render_time_metrics(perf, total_time)

        # 动态结果标签页
        has_step_result = '逐步结转_工单明细' in result and not result['逐步结转_工单明细'].empty
        has_super_result = '超级还原_完工成本' in result
        has_sales_result = '销售成本明细' in result

        tab_labels = [
            "📋 收发存汇总",
            "📊 工单投入产出明细",
            "📈 成本明细",
            "⛓️ 边表与路径表",
        ]
        if has_step_result:
            tab_labels.append("🔄 逐步结转法")
        if has_super_result:
            tab_labels.append("🔬 超级成本还原")
        if has_sales_result:
            tab_labels.append("🚚 销售成本")

        tab_vars = st.tabs(tab_labels)
        ti = 0

        # Tab 1: 收发存汇总
        with tab_vars[ti]:
            c1, c2, c3 = st.columns(3)
            with c1:
                total_val = result['收发存']['期末金额'].sum()
                st.metric("期末总成本", f"¥{total_val:,.2f}")
            with c2:
                max_val = result['收发存']['本期收入金额'].max()
                st.metric("最大收入项", f"¥{max_val:,.2f}")
            with c3:
                count_val = len(result['收发存'])
                st.metric("物料数量", count_val)
            st.dataframe(result['收发存'].sort_values('期末金额', ascending=False),
                         use_container_width=True, height=400)
        ti += 1

        # Tab 2: 工单投入产出明细
        with tab_vars[ti]:
            st1, st2 = st.tabs(["📊 投入产出汇总表", "🔍 投入产出明细表"])
            with st1:
                st.markdown("##### 投入产出汇总表（按工单+产品聚合）")
                st.dataframe(result['工单明细'].sort_values('工单号'),
                             use_container_width=True, height=500)
            with st2:
                st.markdown("##### 投入产出明细表（按工单+产品+材料展开）")
                if '工单产品材料明细' in result and not result['工单产品材料明细'].empty:
                    st.dataframe(result['工单产品材料明细'].sort_values(['工单号', '产品编码', '材料编码']),
                                 use_container_width=True, height=500)
                else:
                    st.info("暂无明细数据")
        ti += 1

        # Tab 3: 成本明细
        with tab_vars[ti]:
            st.dataframe(result['成本明细'].sort_values('总成本', ascending=False),
                         use_container_width=True, height=500)
        ti += 1

        # Tab 4: 边表与路径表
        with tab_vars[ti]:
            st.markdown("##### 边表：相邻流转关系")
            edge_df = create_edge_table(calc)
            if edge_df is not None and not edge_df.empty:
                st.dataframe(edge_df, use_container_width=True, height=350)
            else:
                st.info("暂无边数据")

            st.divider()
            st.markdown("##### 路径表：根到叶的完整路径")
            path_df = create_path_table(calc)
            if path_df is not None and not path_df.empty:
                st.dataframe(path_df, use_container_width=True, height=350)
            else:
                st.info("暂无路径数据")
        ti += 1

        # Tab: 逐步结转法
        if has_step_result:
            with tab_vars[ti]:
                st.markdown("##### 逐步结转法计算结果")
                st.info("**人工** = (I+W×D)×F_工  |  **制费** = (I+W×D)×F_制费  |  **材料** = X总成本 - 人工 - 制费")
                s1, s2, s3 = st.tabs(["📊 投入产出汇总表", "🔍 投入产出明细表", "📈 成本明细"])
                with s1:
                    st.dataframe(result['逐步结转_工单明细'].sort_values('工单号'),
                                 use_container_width=True, height=500)
                with s2:
                    if '逐步结转_工单产品材料明细' in result and not result['逐步结转_工单产品材料明细'].empty:
                        st.dataframe(result['逐步结转_工单产品材料明细'].sort_values(['工单号', '产品编码', '材料编码']),
                                     use_container_width=True, height=500)
                    else:
                        st.info("暂无明细数据")
                with s3:
                    st.dataframe(result['逐步结转_成本明细'].sort_values('总成本', ascending=False),
                                 use_container_width=True, height=500)
            ti += 1

        # Tab: 超级成本还原
        if has_super_result:
            with tab_vars[ti]:
                st.markdown("##### 🔬 超级成本还原")
                n_dims = perf.get('超级还原维度数', 0)
                st.caption(f"成本维度总数: **{n_dims}** 个")

                has_sales_super = '超级还原_销售成本' in result
                sub_labels = ["📋 完工成本", "🏆 TopN汇总", "📖 维度定义", "✅ 验证差异"]
                if has_sales_super:
                    sub_labels.append("🚚 销售成本")
                sub_vars = st.tabs(sub_labels)
                si = 0

                with sub_vars[si]:
                    st.dataframe(result.get('超级还原_完工成本', pd.DataFrame()), use_container_width=True, height=500)
                si += 1
                with sub_vars[si]:
                    st.dataframe(result.get('超级还原_TopN汇总', pd.DataFrame()), use_container_width=True, height=500)
                si += 1
                with sub_vars[si]:
                    st.dataframe(result.get('超级还原_维度定义', pd.DataFrame()), use_container_width=True, height=400)
                si += 1
                with sub_vars[si]:
                    st.dataframe(result.get('超级还原_验证差异', pd.DataFrame()), use_container_width=True, height=300)
                si += 1
                if has_sales_super:
                    with sub_vars[si]:
                        st.dataframe(result.get('超级还原_销售成本', pd.DataFrame()), use_container_width=True, height=500)
            ti += 1

        # Tab: 销售成本
        if has_sales_result:
            with tab_vars[ti]:
                st.markdown("##### 🚚 销售成本")
                st.info("公式：C = B × S × X")
                sales_df = result['销售成本明细']
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("销售总成本", f"¥{sales_df['销售成本_合计'].sum():,.2f}")
                with c2:
                    st.metric("销售总数量", f"{sales_df['销售数量'].sum():,.2f}")
                with c3:
                    st.metric("批次数量", sales_df['销售批次号'].nunique())
                st.dataframe(sales_df, use_container_width=True, height=500)
            ti += 1

        # ==================== 导出按钮 ====================
        st.divider()
        st.markdown("#### 📥 导出结果")

        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            # 当前月份下载
            export_sheets = _build_export_sheets(year, month, result, calc, has_step_result, has_super_result, has_sales_result)
            excel = to_excel(export_sheets)
            st.download_button(
                f"📥 下载 {year}年{month}月 结果",
                excel,
                f"成本核算结果{year}年{month}月.xlsx",
                use_container_width=True
            )

        with col_dl2:
            # 全部月份 ZIP 下载
            if len(all_months) > 1:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for y, m in all_months:
                        r = st.session_state.monthly_results[(y, m)]
                        c = st.session_state.monthly_calc[(y, m)]
                        hs = '逐步结转_工单明细' in r and not r['逐步结转_工单明细'].empty
                        hp = '超级还原_完工成本' in r
                        hl = '销售成本明细' in r
                        sheets = _build_export_sheets(y, m, r, c, hs, hp, hl)
                        single_excel = to_excel(sheets)
                        zf.writestr(f"成本核算结果{y}年{m}月.xlsx", single_excel.getvalue())
                zip_buffer.seek(0)
                st.download_button(
                    f"📦 下载全部 {len(all_months)} 个月结果 (ZIP)",
                    zip_buffer,
                    "成本核算结果_全部月份.zip",
                    use_container_width=True
                )
            else:
                st.caption("仅单月数据，无需批量下载")


def _build_export_sheets(year, month, result, calc, has_step_result, has_super_result, has_sales_result):
    """构建导出 sheets 字典"""
    prefix = f'{year}年{month}月-'
    sheets = {
        prefix + '收发存汇总': result['收发存'],
        prefix + '投入产出汇总表': result['工单明细'],
        prefix + '投入产出明细表': result.get('工单产品材料明细', pd.DataFrame()),
        prefix + '成本明细': result['成本明细'],
    }

    if has_step_result:
        sheets[prefix + '逐步结转_汇总表'] = result['逐步结转_工单明细']
        sheets[prefix + '逐步结转_明细表'] = result.get('逐步结转_工单产品材料明细', pd.DataFrame())
        sheets[prefix + '逐步结转_成本明细'] = result['逐步结转_成本明细']

    if has_super_result:
        sheets[prefix + '超级还原_完工成本'] = result.get('超级还原_完工成本', pd.DataFrame())
        if '超级还原_销售成本' in result:
            sheets[prefix + '超级还原_销售成本'] = result['超级还原_销售成本']

    if has_sales_result:
        sheets[prefix + '销售成本明细'] = result['销售成本明细']

    edge_df = create_edge_table(calc)
    path_df = create_path_table(calc)
    if edge_df is not None and not edge_df.empty:
        sheets[prefix + '边表'] = edge_df
    if path_df is not None and not path_df.empty:
        sheets[prefix + '路径表'] = path_df

    return sheets


# ==================== 成本还原页面 ====================
def render_cost_restoration():
    st.markdown("### 🔄 成本还原")

    if st.button("← 返回首页", key="back_home_restore"):
        st.session_state.current_page = 'home'
        st.rerun()

    st.divider()
    st.info("成本还原需要已完成的成本核算结果（复用W流转矩阵），基于逐步结转报表还原真实料工费结构。")

    # 上传逐步结转成本表
    st.markdown("#### 📄 上传逐步结转成本表")
    with st.expander("🔴 逐步结转成本表", expanded=True):
        f = st.file_uploader("上传逐步结转成本表", type=['xlsx', 'xls'], key="restore_step")
        if f:
            try:
                df = pd.read_excel(f)
                st.caption(f"✓ 读取成功: {len(df)} 行 × {len(df.columns)} 列")
                with st.expander("📋 数据预览", expanded=False):
                    st.dataframe(df.head(5), use_container_width=True, height=200)

                patterns = {
                    '物料编码': [(r'物料.*编码$', 10), (r'产品.*编码$', 10), (r'料号$', 9), (r'品号$', 8), (r'编码$', 5)],
                    '料': [(r'材料$', 10), (r'直接材料$', 9), (r'原材料$', 8), (r'料$', 5)],
                    '工': [(r'人工$', 10), (r'直接人工$', 9), (r'工$', 5)],
                    '费': [(r'制费$', 10), (r'制造费用$', 9), (r'费$', 5)],
                }
                auto_map = {}
                for std_col, regex_list in patterns.items():
                    for col in df.columns:
                        col_str = str(col).strip()
                        for pat, score in regex_list:
                            if re.search(pat, col_str, re.IGNORECASE):
                                auto_map[std_col] = col
                                break
                        if std_col in auto_map:
                            break

                final_map = {}
                cols_iter = st.columns(4)
                for i, std_col in enumerate(['物料编码', '料', '工', '费']):
                    with cols_iter[i]:
                        options = ['(跳过)'] + list(df.columns)
                        matched = auto_map.get(std_col)
                        default_idx = options.index(matched) if matched in options else 0
                        selected = st.selectbox(std_col, options, index=default_idx, key=f"restore_map_{std_col}")
                        if selected != '(跳过)':
                            final_map[selected] = std_col

                if len(final_map) >= 4:
                    st.session_state.step_cost_file = df
                    st.session_state.step_cost_mapping = final_map
                    st.success("✓ 字段匹配完成")
            except Exception as e:
                st.error(f"❌ 读取失败: {e}")

    # 上传基础数据（可选）
    st.divider()
    st.markdown("#### 📄 上传基础数据（可选，用于重建W矩阵）")
    with st.expander("基础数据", expanded=False):
        st.info("如不重新上传，将复用最近一次成本核算的W流转矩阵")

        uploaded = {}
        mappings = {}
        for key, title in [('purchase', '采购入库'), ('io', '投入产出明细'), ('initial', '期初结存'), ('labor', '工单人工制费')]:
            f2 = st.file_uploader(title, type=['xlsx', 'xls'], key=f"restore_base_{key}")
            if f2:
                df2 = pd.read_excel(f2)
                st.caption(f"✓ {len(df2)}行")
                req_fields = TABLE_REQUIRED_FIELDS.get(key, [])
                auto = smart_match(df2.columns, key)
                selected_map = {}
                for field in req_fields:
                    matched = auto.get(field)
                    if matched:
                        selected_map[matched] = field
                if selected_map:
                    mappings[key] = selected_map
                    uploaded[key] = f2

    st.divider()

    ready_restore = st.session_state.step_cost_file is not None
    if not ready_restore:
        st.warning("⚠️ 请上传逐步结转成本表")
        return

    if st.button("🔄 执行成本还原", type="primary", use_container_width=True):
        with st.spinner("正在计算..."):
            try:
                start = time.time()

                has_cost = bool(st.session_state.monthly_calc)

                if has_cost and not uploaded:
                    first_month = st.session_state.all_months[0]
                    calc = st.session_state.monthly_calc[first_month]
                    if calc.W_matrix is None:
                        calc.calculate()
                elif uploaded:
                    calc = CostCalculator()
                    calc.load_data({
                        'initial': uploaded.get('initial', pd.DataFrame()),
                        'purchase': uploaded.get('purchase', pd.DataFrame()),
                        'io': uploaded.get('io', pd.DataFrame()),
                        'labor': uploaded.get('labor', pd.DataFrame()),
                        'finished': pd.DataFrame(),
                        'sales': pd.DataFrame(),
                    })
                    calc.calculate()
                else:
                    st.error("❌ 没有可用的成本核算数据，请先执行成本核算或手动上传基础数据")
                    return

                restore_result = calc.calculate_cost_restoration(
                    st.session_state.step_cost_file,
                    st.session_state.step_cost_mapping
                )

                total_time = time.time() - start
                st.session_state.restore_result = restore_result
                st.session_state.restore_calc = calc
                st.session_state.restore_total_time = total_time
                st.success(f"✅ 成本还原完成！用时 {total_time:.3f}s")

            except Exception as e:
                st.error(f"❌ 计算失败: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    # 显示还原结果
    if st.session_state.restore_result:
        st.divider()
        result = st.session_state.restore_result
        calc = st.session_state.restore_calc
        step_df = st.session_state.step_cost_file
        step_map = st.session_state.step_cost_mapping
        reverse_map = {v: k for k, v in step_map.items()}
        col_mat = reverse_map.get('料')
        col_lab = reverse_map.get('工')
        col_oh = reverse_map.get('费')
        col_code = reverse_map.get('物料编码')

        comparison = []
        for _, row in step_df.iterrows():
            mat = str(row[col_code])
            if mat in calc.node_index:
                idx = calc.node_index[mat]
                step_mat_val = float(row[col_mat] or 0)
                step_lab_val = float(row[col_lab] or 0)
                step_oh_val = float(row[col_oh] or 0)
                step_total = step_mat_val + step_lab_val + step_oh_val
                rest_mat, rest_lab, rest_oh = result['restored_costs'][idx]
                rest_total = rest_mat + rest_lab + rest_oh

                comparison.append({
                    '物料编码': mat,
                    '逐步结转-料': round(step_mat_val, 4),
                    '逐步结转-工': round(step_lab_val, 4),
                    '逐步结转-费': round(step_oh_val, 4),
                    '逐步结转-合计': round(step_total, 4),
                    '还原后-真料': round(float(rest_mat), 4),
                    '还原后-真工': round(float(rest_lab), 4),
                    '还原后-真费': round(float(rest_oh), 4),
                    '还原后-合计': round(float(rest_total), 4),
                    '料差异': round(float(rest_mat) - step_mat_val, 4),
                    '工差异': round(float(rest_lab) - step_lab_val, 4),
                    '费差异': round(float(rest_oh) - step_oh_val, 4),
                })

        comp_df = pd.DataFrame(comparison)
        st.dataframe(comp_df, use_container_width=True, height=400)

        excel = to_excel({'成本还原对比': comp_df, '还原后明细': result['detail_df']})
        st.download_button("📥 下载还原结果", excel, "成本还原结果.xlsx", use_container_width=True)


# ==================== 主程序 ====================
def main():
    page = st.session_state.current_page
    if page == 'home':
        render_home()
    elif page == 'cost':
        render_cost_accounting()
    elif page == 'restore':
        render_cost_restoration()


if __name__ == "__main__":
    main()
