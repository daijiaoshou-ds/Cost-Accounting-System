import streamlit as st
import pandas as pd
import numpy as np
from logic import CostCalculator, to_excel
import time
import re
from io import BytesIO
import zipfile
import os

# 页面配置
st.set_page_config(
    page_title="矩阵成本核算引擎",
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
    .nav-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        cursor: pointer;
        transition: transform 0.3s;
        text-align: center;
    }
    .nav-card:hover {
        transform: translateY(-5px);
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
    .field-mapping-box {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 3px solid #28a745;
        margin: 0.5rem 0;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    /* 自定义下拉框样式 */
    div[data-testid="stSelectbox"] {
        margin-bottom: 0.5rem;
    }
    div[data-testid="stSelectbox"] label {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #333 !important;
        margin-bottom: 0.2rem !important;
    }
    /* 调整列间距 */
    div[data-testid="column"] {
        padding: 0 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 工具函数 ====================
def smart_match(columns, file_type):
    """智能字段匹配"""
    cols = list(columns)
    result = {}
    
    patterns = {
        '期初': {
            '物料编码': [
                (r'物料.*编码$', 10), (r'料号$', 9), (r'品号$', 8), 
                (r'物料代码', 7), (r'材料编码', 6), (r'存货编码', 6), (r'编码$', 5)
            ],
            '期初金额': [
                (r'期初.*金额$', 10), (r'期初.*成本$', 9), (r'期初金额', 8),
                (r'期初价值', 7), (r'金额$', 3)
            ],
            '期初数量': [
                (r'期初.*数量$', 10), (r'期初.*库存$', 9), (r'期初数量', 8),
                (r'数量$', 2)
            ]
        },
        '采购': {
            '物料编码': [
                (r'物料.*编码$', 10), (r'存货编码$', 9), (r'料号$', 9), (r'品号$', 8),
                (r'物料代码', 7), (r'编码$', 5)
            ],
            '采购数量': [
                (r'采购.*数量$', 10), (r'入库.*数量$', 9), (r'实收数量', 8),
                (r'采购数', 8), (r'数量$', 3)
            ],
            '采购金额': [
                (r'采购.*金额$', 10), (r'采购.*成本$', 9), (r'金额$', 3)
            ]
        },
        '投入产出': {
            '工单号': [
                (r'工单号?$', 10), (r'生产订单$', 9), (r'订单号$', 8),
                (r'工单', 7), (r'^MO', 6)
            ],
            '产品编码': [
                (r'产品.*编码$', 10), (r'产成品', 9), (r'成品编码', 8),
                (r'产品代码', 7), (r'产出编码', 6)
            ],
            '产品完工数量': [
                (r'完工.*数量$', 10), (r'产出.*数量$', 9), (r'产量$', 8),
                (r'完工数', 7), (r'数量', 2)
            ],
            '在产品数量': [
                (r'在产.*数量$', 10), (r'在产品.*数量$', 9), (r'在制品.*数量$', 8),
                (r'在产数', 7), (r'在制品', 6)
            ],
            '材料编码': [
                (r'材料.*编码$', 10), (r'物料.*编码', 9), (r'领料编码', 8),
                (r'材料代码', 7)
            ],
            '材料领用数量': [
                (r'领用.*数量$', 10), (r'领料.*数量$', 9), (r'消耗.*数量$', 8),
                (r'用量$', 7), (r'领用数', 6)
            ]
        },
        '入库': {
            '产品编码': [
                (r'产品.*编码$', 10), (r'产成品.*编码$', 9), (r'成品编码$', 8),
                (r'物料编码$', 7), (r'产品代码', 6), (r'编码$', 5)
            ],
            '入库数量': [
                (r'入库.*数量$', 10), (r'完工.*数量$', 9), (r'入库数', 8),
                (r'数量$', 5), (r'入库量', 7)
            ]
        },
        '工单费用': {
            '工单号': [
                (r'工单号?$', 10), (r'生产订单$', 9), (r'订单号$', 8), (r'工单', 7)
            ],
            '人工': [
                (r'人工$', 10), (r'人工费$', 9), (r'直接人工', 8), (r'人工金额', 7)
            ],
            '制费': [
                (r'制费$', 10), (r'制造费用$', 9), (r'制造费', 8), 
                (r'间接费用', 7), (r'制造成本', 6)
            ]
        },
        '逐步结转': {
            '物料编码': [
                (r'物料.*编码$', 10), (r'产品.*编码$', 10), (r'料号$', 9), 
                (r'品号$', 8), (r'编码$', 5)
            ],
            '料': [
                (r'材料$', 10), (r'直接材料$', 9), (r'原材料$', 8), (r'料$', 5)
            ],
            '工': [
                (r'人工$', 10), (r'直接人工$', 9), (r'工$', 5)
            ],
            '费': [
                (r'制费$', 10), (r'制造费用$', 9), (r'费$', 5)
            ]
        }
    }
    
    patterns_for_type = patterns.get(file_type, {})
    
    for standard_col, regex_list in patterns_for_type.items():
        best_match = None
        best_score = -1
        
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

def render_field_mapping(file_key, auto_map, df_columns, ptype):
    """渲染字段映射界面 - 垂直布局"""
    final_map = {}
    
    if not auto_map:
        return final_map
    
    # 垂直布局，每行一个字段
    for std_col, matched_col in auto_map.items():
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.markdown(f"**{std_col}**")
        
        with col2:
            options = ['(跳过)'] + list(df_columns)
            default_idx = options.index(matched_col) if matched_col in options else 0
            selected = st.selectbox(
                f"映射_{file_key}_{std_col}",
                options,
                index=default_idx,
                label_visibility="collapsed",
                key=f"map_{file_key}_{std_col}"
            )
            if selected != '(跳过)':
                final_map[selected] = std_col
        
        # 添加分隔线
        st.markdown("<hr style='margin: 0.5rem 0; opacity: 0.3;'>", unsafe_allow_html=True)
    
    return final_map

def detect_file_type(filename):
    """根据文件名检测文件类型"""
    name_lower = filename.lower()
    # 优先级：先匹配更具体的，再匹配通用的
    if any(kw in name_lower for kw in ['产成品入库', '产品入库', '成品入库', '完工入库', 'finished']):
        return 'finished', '产成品入库明细'
    elif any(kw in name_lower for kw in ['采购入库', '采购', 'purchase']):
        return 'purchase', '采购入库'
    elif any(kw in name_lower for kw in ['投入', '产出', 'io', '工单', 'mo']):
        return 'io', '投入产出明细'
    elif any(kw in name_lower for kw in ['期初', 'initial', '期初结存']):
        return 'initial', '期初结存'
    elif any(kw in name_lower for kw in ['人工', '制费', '费用', 'labor', 'cost']):
        return 'labor', '工单人工制费'
    return None, None

def format_time(seconds):
    """格式化时间为秒，保留3位小数"""
    return f"{seconds:.3f}s"

def render_time_metrics(perf, total_time):
    """渲染美观的时间指标 - 只显示关键指标"""
    st.markdown("### ⚡ 计算性能")
    
    # 显示矩阵维度
    n = perf.get('矩阵维度', 0)
    if n > 0:
        st.caption(f"矩阵维度: {n}×{n}")
    
    # 只显示指定的指标
    display_keys = ['数据清洗', '计算数量', '矩阵求解']
    stages = []
    for key in display_keys:
        if key in perf:
            stages.append((key, perf[key]))
    
    if not stages:
        return
    
    # 使用列布局显示时间卡片
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
        method = perf.get('求解方法', '')
        method_text = f"<div style='font-size: 0.8rem; color: #888;'>{method}</div>" if method else ""
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #28a745;">
            <div style="font-size: 0.9rem; color: #666;">总耗时</div>
            <div style="font-size: 1.6rem; font-weight: bold; color: #28a745;">{format_time(total_time)}</div>
            {method_text}
        </div>
        """, unsafe_allow_html=True)

def create_network_graph(calc):
    """创建材料穿透网络图"""
    try:
        from pyvis.network import Network
        import tempfile
        import os
        
        # 获取计算结果中的必要数据
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            st.warning("未找到流转矩阵数据")
            return None
        
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes
        
        n = len(all_nodes)
        
        # 创建网络图
        net = Network(height="600px", width="100%", directed=True, 
                      bgcolor="#ffffff", font_color="#333333")
        
        # 添加节点
        for i, node in enumerate(all_nodes):
            if node in order_nodes:
                # 工单节点 - 绿色
                net.add_node(node, label=node, color="#4CAF50", 
                           size=25, title=f"工单: {node}", shape="box")
            elif node in material_nodes:
                # 物料节点 - 蓝色
                net.add_node(node, label=node, color="#2196F3", 
                           size=20, title=f"物料: {node}", shape="dot")
        
        # 添加边（W 矩阵中的非零元素）
        for i in range(n):
            for j in range(n):
                if W[i, j] > 0.001:  # 只显示显著的流转
                    source = all_nodes[j]
                    target = all_nodes[i]
                    weight = W[i, j]
                    
                    # 边的粗细与权重成正比
                    width = max(1, weight * 10)
                    
                    # 箭头方向: source -> target
                    net.add_edge(source, target, width=width, 
                               title=f"系数: {weight:.4f}",
                               arrows="to", color="#999999")
        
        # 物理布局参数
        net.set_options("""
        {
          "physics": {
            "forceAtlas2Based": {
              "gravitationalConstant": -50,
              "centralGravity": 0.01,
              "springLength": 100,
              "springConstant": 0.08
            },
            "maxVelocity": 50,
            "solver": "forceAtlas2Based",
            "timestep": 0.35,
            "stabilization": {"iterations": 150}
          },
          "interaction": {
            "hover": true,
            "tooltipDelay": 200
          }
        }
        """)
        
        # 保存并读取HTML
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', 
                                         delete=False, encoding='utf-8') as f:
            temp_path = f.name
        
        net.save_graph(temp_path)
        
        with open(temp_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        os.unlink(temp_path)
        
        return html_content
        
    except ImportError:
        st.error("请安装 pyvis: pip install pyvis")
        return None
    except Exception as e:
        st.error(f"生成网络图失败: {e}")
        return None

# ==================== Session State 初始化 ====================
def init_session_state():
    """初始化所有 session state 变量"""
    defaults = {
        'current_page': 'home',
        'cost_result': None,
        'cost_calc': None,
        'cost_data': {},  # 保存成本核算的数据供成本还原复用
        'restore_result': None,
        'uploaded_files': {},
        'field_mappings': {},
        'step_cost_file': None,
        'step_cost_mapping': None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# ==================== 首页 ====================
def render_home():
    """渲染首页导航"""
    st.markdown('<div class="main-header">🧮 矩阵成本核算引擎</div>', 
                unsafe_allow_html=True)
    
    st.markdown("""
    <div style="text-align: center; color: #666; margin-bottom: 3rem;">
        基于线性代数的企业成本核算系统，使用逆矩阵算法精确计算产品成本
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    padding: 2rem; border-radius: 15px; color: white; text-align: center;">
            <div style="font-size: 3rem; margin-bottom: 1rem;">📊</div>
            <div style="font-size: 1.5rem; font-weight: bold; margin-bottom: 0.5rem;">
                成本核算
            </div>
            <div style="font-size: 0.9rem; opacity: 0.9; margin-bottom: 1.5rem;">
                基于投入产出数据，使用矩阵算法计算各产品的料工费成本
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
            <div style="font-size: 1.5rem; font-weight: bold; margin-bottom: 0.5rem;">
                成本还原
            </div>
            <div style="font-size: 0.9rem; opacity: 0.9; margin-bottom: 1.5rem;">
                基于逐步结转法数据，还原各产品的真实料工费结构
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("进入成本还原", key="btn_restore", use_container_width=True):
            st.session_state.current_page = 'restore'
            st.rerun()
    
    # 技术说明
    with st.expander("📖 技术原理"):
        st.markdown("""
        ### 核心公式
        
        系统基于投入产出分析（Leontief 模型）：
        
        $$X = (I - W)^{-1} \\times F$$
        
        其中：
        - **W**: 物料与工单之间的消耗/产出流转矩阵
        - **F**: 外部投入矩阵（期初、采购、人工、制费）
        - **X**: 各节点的总成本（料、工、费三列独立）
        
        ### 成本还原原理
        
        逐步结转法下，下级的工费被归入上级的"料"中。成本还原通过相同的流转矩阵 W，
        但将 F 矩阵替换为逐步结转法的料工费，从而计算出真实的料工费结构。
        """)

# ==================== 成本核算页面 ====================
def render_cost_accounting():
    """渲染成本核算页面"""
    st.markdown("### 📊 成本核算")
    
    # 返回按钮
    if st.button("← 返回首页", key="back_home_cost"):
        st.session_state.current_page = 'home'
        st.rerun()
    
    st.divider()
    
    # ==================== 批量上传区域 ====================
    st.markdown("#### 📁 批量上传（可选）")
    
    with st.expander("一次性上传多个文件，系统自动识别", expanded=False):
        st.info("支持 ZIP 压缩包或同时选择多个 Excel 文件，系统会根据文件名自动识别文件类型")
        
        batch_files = st.file_uploader(
            "上传文件（可多选）",
            type=['xlsx', 'xls', 'zip'],
            accept_multiple_files=True,
            key="batch_upload"
        )
        
        auto_detected = {}
        
        if batch_files:
            for f in batch_files:
                if f.name.endswith('.zip'):
                    # 处理 ZIP 文件
                    try:
                        with zipfile.ZipFile(f) as z:
                            for name in z.namelist():
                                if name.endswith(('.xlsx', '.xls')):
                                    file_type, type_name = detect_file_type(name)
                                    if file_type:
                                        with z.open(name) as excel_file:
                                            auto_detected[file_type] = (excel_file, type_name, name)
                                        st.success(f"📦 ZIP 中找到: {type_name} - {name}")
                    except Exception as e:
                        st.error(f"无法读取 ZIP 文件: {e}")
                else:
                    # 处理单个 Excel 文件
                    file_type, type_name = detect_file_type(f.name)
                    if file_type:
                        auto_detected[file_type] = (f, type_name, f.name)
                        st.success(f"✓ 识别为 {type_name}: {f.name}")
                    else:
                        st.warning(f"⚠️ 无法识别文件类型: {f.name}")
    
    st.divider()
    
    # ==================== 单个文件上传区域 ====================
    st.markdown("#### 📄 逐个上传（精确控制）")
    
    file_configs = [
        ('purchase', '🔴 采购入库', '采购', True),
        ('io', '🔴 投入产出明细', '投入产出', True),
        ('finished', '🔴 产成品入库明细', '入库', True),  # 新增：产成品入库
        ('initial', '⚪ 期初结存', '期初', False),
        ('labor', '⚪ 工单人工制费', '工单费用', False),
    ]
    
    uploaded = {}
    mappings = {}
    
    for key, title, ptype, is_required in file_configs:
        # 检查是否有批量上传的自动识别结果
        auto_file = None
        auto_name = None
        if key in auto_detected:
            auto_file, _, auto_name = auto_detected[key]
            auto_file.seek(0)  # 重置文件指针
        
        with st.expander(title, expanded=(is_required and not auto_file)):
            # 文件上传
            if auto_file:
                st.success(f"✓ 已从批量上传识别: {auto_name}")
                # 重新读取文件内容
                f = auto_detected[key][0]
                f.seek(0)
            else:
                help_text = f"{'【必须】' if is_required else '【可选】'}上传 Excel 文件"
                f = st.file_uploader(help_text, type=['xlsx', 'xls'], key=f"cost_{key}")
            
            if f:
                try:
                    f.seek(0)
                    df = pd.read_excel(f)
                    st.caption(f"✓ 读取成功: {len(df)} 行 × {len(df.columns)} 列")
                    
                    # 显示数据预览（折叠）
                    with st.expander("📋 数据预览", expanded=False):
                        st.dataframe(df.head(5), use_container_width=True, height=200)
                    
                    # 智能匹配字段
                    auto_map = smart_match(df.columns, ptype)
                    
                    if auto_map:
                        st.success(f"✓ 自动匹配 {len(auto_map)} 个字段，请确认或调整：")
                        
                        # 使用新的垂直布局字段映射
                        final_map = render_field_mapping(f"cost_{key}", auto_map, df.columns, ptype)
                        
                        if final_map:
                            mappings[key] = final_map
                            uploaded[key] = f
                        else:
                            st.warning("⚠️ 至少需要映射一个字段")
                    else:
                        st.warning("⚠️ 未自动识别字段，请手动选择")
                        manual_map = {}
                        cols = st.columns(2)
                        col_idx = 0
                        
                        # 根据类型确定需要的字段
                        required_fields = {
                            '采购': ['物料编码', '采购数量', '采购金额'],
                            '投入产出': ['工单号', '产品编码', '产品完工数量', '材料编码', '材料领用数量'],
                            '期初': ['物料编码', '期初金额'],
                            '工单费用': ['工单号', '人工', '制费']
                        }.get(ptype, [])
                        # 注意：在产品数量是可选字段，不是必须的
                        
                        for field in required_fields:
                            with cols[col_idx % 2]:
                                selected = st.selectbox(
                                    f"{field}",
                                    ['(跳过)'] + list(df.columns),
                                    key=f"cost_manual_{key}_{field}"
                                )
                                if selected != '(跳过)':
                                    manual_map[selected] = field
                            col_idx += 1
                        
                        if manual_map:
                            mappings[key] = manual_map
                            uploaded[key] = f
                            
                except Exception as e:
                    st.error(f"❌ 读取失败: {e}")
            elif is_required:
                st.info(f"⬆️ 请上传文件，或使用上方批量上传功能")
    
    # ==================== 步骤2: 执行计算 ====================
    st.divider()
    st.markdown("#### 🚀 执行计算")
    
    ready = 'purchase' in uploaded and 'io' in uploaded
    
    if not ready:
        st.warning("⚠️ 请至少上传【采购入库】和【投入产出明细】文件")
        return
    
    if st.button("🚀 执行成本核算", type="primary", use_container_width=True):
        with st.spinner("正在计算..."):
            try:
                start = time.time()
                
                calc = CostCalculator()
                calc.load_data(
                    uploaded.get('initial'),
                    uploaded['purchase'],
                    uploaded['io'],
                    uploaded.get('labor'),
                    mappings.get('initial', {}),
                    mappings.get('purchase', {}),
                    mappings.get('io', {}),
                    mappings.get('labor', {})
                )
                
                # 如果有产成品入库明细，传入calculate
                fin_file = uploaded.get('finished')
                fin_map = mappings.get('finished')
                if fin_file and fin_map:
                    fin_file.seek(0)
                    finished_df = pd.read_excel(fin_file)
                    result = calc.calculate(finished_df, fin_map)
                else:
                    result = calc.calculate()
                total_time = time.time() - start
                
                # 保存结果到 session state - 包含数据和计算器实例
                st.session_state.cost_result = result
                st.session_state.cost_calc = calc
                st.session_state.cost_total_time = total_time
                st.session_state.cost_perf = calc.get_performance()
                
                # 保存数据供成本还原复用
                st.session_state.cost_data = {
                    'uploaded': uploaded.copy(),
                    'mappings': mappings.copy()
                }
                
                st.success(f"✅ 计算完成！")
                
            except Exception as e:
                st.error(f"❌ 计算失败: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
    
    # ==================== 步骤3: 显示结果 ====================
    if st.session_state.cost_result:
        st.divider()
        st.markdown("#### 📊 计算结果")
        
        result = st.session_state.cost_result
        perf = st.session_state.cost_perf
        total_time = st.session_state.cost_total_time
        
        # 性能指标
        render_time_metrics(perf, total_time)
        
        # 结果标签页
        tab1, tab2, tab3 = st.tabs(["📋 收发存汇总", "📊 工单投入产出明细", "📈 成本明细"])
        
        with tab1:
            # 统计卡片
            c1, c2, c3 = st.columns(3)
            with c1:
                total = result['收发存']['期末金额'].sum()
                st.metric("期末总成本", f"¥{total:,.2f}")
            with c2:
                max_val = result['收发存']['本期收入金额'].max()
                st.metric("最大收入项", f"¥{max_val:,.2f}")
            with c3:
                count = len(result['收发存'])
                st.metric("物料数量", count)
            
            # 数据表
            st.dataframe(result['收发存'].sort_values('期末金额', ascending=False), 
                        use_container_width=True, height=400)
        
        with tab2:
            st.markdown("##### 工单投入产出明细（在产与完工拆分）")
            st.dataframe(result['工单明细'].sort_values('工单号'),
                        use_container_width=True, height=500)
        
        with tab3:
            st.dataframe(result['成本明细'].sort_values('总成本', ascending=False),
                        use_container_width=True, height=500)
        
        with tab3:
            st.markdown("##### 材料流转网络图")
            st.caption("蓝色圆点 = 物料节点，绿色方框 = 工单节点，连线粗细 = 流转系数")
            
            calc = st.session_state.cost_calc
            if calc:
                graph_html = create_network_graph(calc)
                if graph_html:
                    st.components.v1.html(graph_html, height=620, scrolling=False)
                else:
                    st.info("网络图生成失败，请检查 pyvis 是否安装")
        
        # 下载按钮
        st.divider()
        excel = to_excel({
            '收发存汇总': result['收发存'],
            '工单投入产出明细': result['工单明细'],
            '成本明细': result['成本明细']
        })
        st.download_button("📥 下载Excel结果", excel, "成本核算结果.xlsx", 
                          use_container_width=True)

# ==================== 成本还原页面 ====================
def render_cost_restoration():
    """渲染成本还原页面"""
    st.markdown("### 🔄 成本还原")
    
    # 返回按钮
    if st.button("← 返回首页", key="back_home_restore"):
        st.session_state.current_page = 'home'
        st.rerun()
    
    st.divider()
    
    # 说明
    st.info("""
    **成本还原说明**：成本还原需要先进行成本核算以获取物料流转矩阵，
    然后基于逐步结转法的料工费数据，计算出真实的料工费结构。
    """)
    
    # 检查是否有成本核算的数据可以复用
    has_cost_data = st.session_state.cost_data and 'uploaded' in st.session_state.cost_data
    
    if has_cost_data:
        st.success("✅ 检测到已完成的成本核算数据，可以直接用于成本还原")
        use_existing = st.checkbox("使用成本核算的基础数据", value=True, 
                                   help="勾选后将自动复用成本核算中上传的基础数据")
    else:
        use_existing = False
        st.warning("⚠️ 未检测到成本核算数据，请先完成成本核算或手动上传基础数据")
    
    # ==================== 步骤1: 基础数据 ====================
    st.markdown("#### 📄 上传基础数据（与成本核算相同）")
    
    if use_existing:
        with st.expander("✅ 已复用成本核算的基础数据", expanded=False):
            st.write("以下数据已自动加载：")
            for key in ['purchase', 'io', 'initial', 'labor']:
                if key in st.session_state.cost_data['uploaded']:
                    st.write(f"- ✅ {key}")
        
        # 直接使用已保存的数据
        uploaded = st.session_state.cost_data['uploaded']
        mappings = st.session_state.cost_data['mappings']
    else:
        # 手动上传
        file_configs = [
            ('purchase', '🔴 采购入库', '采购', True),
            ('io', '🔴 投入产出明细', '投入产出', True),
            ('initial', '⚪ 期初结存', '期初', False),
            ('labor', '⚪ 工单人工制费', '工单费用', False),
        ]
        
        uploaded = {}
        mappings = {}
        
        for key, title, ptype, is_required in file_configs:
            with st.expander(title, expanded=is_required):
                help_text = f"{'【必须】' if is_required else '【可选】'}上传 Excel 文件"
                f = st.file_uploader(help_text, type=['xlsx', 'xls'], key=f"restore_{key}")
                
                if f:
                    try:
                        df = pd.read_excel(f)
                        st.caption(f"✓ 读取成功: {len(df)} 行 × {len(df.columns)} 列")
                        
                        # 显示数据预览（折叠）
                        with st.expander("📋 数据预览", expanded=False):
                            st.dataframe(df.head(5), use_container_width=True, height=200)
                        
                        auto_map = smart_match(df.columns, ptype)
                        
                        if auto_map:
                            st.success(f"✓ 自动匹配 {len(auto_map)} 个字段，请确认或调整：")
                            
                            final_map = render_field_mapping(f"restore_{key}", auto_map, df.columns, ptype)
                            
                            if final_map:
                                mappings[key] = final_map
                                uploaded[key] = f
                        else:
                            st.warning("⚠️ 未自动识别字段，请手动选择")
                            
                    except Exception as e:
                        st.error(f"❌ 读取失败: {e}")
                elif is_required:
                    st.info(f"⬆️ 请上传 {title} 文件")
    
    # ==================== 步骤2: 逐步结转成本表 ====================
    st.divider()
    st.markdown("#### 📄 上传逐步结转成本表")
    
    with st.expander("🔴 逐步结转成本表", expanded=True):
        st.info("上传逐步结转法下核算出的各物料料工费数据")
        
        f = st.file_uploader("【必须】上传逐步结转成本表", type=['xlsx', 'xls'], 
                            key="restore_step_cost")
        
        if f:
            try:
                df = pd.read_excel(f)
                st.caption(f"✓ 读取成功: {len(df)} 行 × {len(df.columns)} 列")
                
                # 显示预览（折叠）
                with st.expander("📋 数据预览", expanded=False):
                    st.dataframe(df.head(5), use_container_width=True, height=200)
                
                # 字段匹配
                auto_map = smart_match(df.columns, '逐步结转')
                
                if auto_map:
                    st.success(f"✓ 自动匹配 {len(auto_map)} 个字段，请确认或调整：")
                    
                    final_map = {}
                    cols = st.columns(2)
                    col_idx = 0
                    
                    for std_col in ['物料编码', '料', '工', '费']:
                        matched = auto_map.get(std_col)
                        with cols[col_idx % 2]:
                            options = ['(跳过)'] + list(df.columns)
                            default_idx = options.index(matched) if matched else 0
                            selected = st.selectbox(
                                f"{std_col}",
                                options,
                                index=default_idx,
                                key=f"restore_step_map_{std_col}"
                            )
                            if selected != '(跳过)':
                                final_map[selected] = std_col
                        col_idx += 1
                    
                    if len(final_map) >= 4:
                        st.session_state.step_cost_file = df
                        st.session_state.step_cost_mapping = final_map
                    else:
                        st.warning("⚠️ 需要完整匹配物料编码、料、工、费四个字段")
                else:
                    st.warning("⚠️ 未自动识别，请手动选择")
                    
            except Exception as e:
                st.error(f"❌ 读取失败: {e}")
        else:
            st.info("⬆️ 请上传逐步结转成本表")
    
    # ==================== 步骤3: 执行还原 ====================
    st.divider()
    st.markdown("#### 🚀 执行成本还原")
    
    ready = ('purchase' in uploaded and 'io' in uploaded and 
             st.session_state.step_cost_file is not None)
    
    if not ready:
        st.warning("⚠️ 请完成上述所有数据上传")
        return
    
    if st.button("🔄 执行成本还原", type="primary", use_container_width=True):
        with st.spinner("正在计算..."):
            try:
                start = time.time()
                
                # 检查是否可以直接复用成本核算的计算器
                if use_existing and st.session_state.cost_calc:
                    calc = st.session_state.cost_calc
                    # 检查计算器是否有 W 矩阵
                    if calc.W_matrix is None:
                        # 需要重新计算以获取 W 矩阵
                        calc.calculate()
                else:
                    # 重新加载数据并计算
                    calc = CostCalculator()
                    calc.load_data(
                        uploaded.get('initial'),
                        uploaded['purchase'],
                        uploaded['io'],
                        uploaded.get('labor'),
                        mappings.get('initial', {}),
                        mappings.get('purchase', {}),
                        mappings.get('io', {}),
                        mappings.get('labor', {})
                    )
                    calc.calculate()
                
                # 执行成本还原
                step_df = st.session_state.step_cost_file
                step_map = st.session_state.step_cost_mapping
                
                restore_result = calc.calculate_cost_restoration(step_df, step_map)
                
                total_time = time.time() - start
                
                # 保存结果
                st.session_state.restore_result = restore_result
                st.session_state.restore_calc = calc
                st.session_state.restore_total_time = total_time
                
                st.success(f"✅ 成本还原完成！")
                
            except Exception as e:
                st.error(f"❌ 计算失败: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
    
    # ==================== 步骤4: 显示还原结果 ====================
    if st.session_state.restore_result:
        st.divider()
        st.markdown("#### 📊 还原结果")
        
        result = st.session_state.restore_result
        calc = st.session_state.restore_calc
        total_time = st.session_state.restore_total_time
        
        # 显示时间
        st.metric("计算用时", f"{total_time:.3f}s")
        
        # 对比表
        st.markdown("##### 成本还原对比表")
        
        # 构建对比数据
        step_df = st.session_state.step_cost_file
        step_map = st.session_state.step_cost_mapping
        
        comparison = []
        for _, row in step_df.iterrows():
            mat = str(row[step_map['物料编码']])
            if mat in calc.node_index:
                idx = calc.node_index[mat]
                step_mat = row[step_map['料']]
                step_lab = row[step_map['工']]
                step_oh = row[step_map['费']]
                step_total = step_mat + step_lab + step_oh
                
                rest_mat, rest_lab, rest_oh = result['restored_costs'][idx]
                rest_total = rest_mat + rest_lab + rest_oh
                
                comparison.append({
                    '物料编码': mat,
                    '逐步结转-料': round(step_mat, 4),
                    '逐步结转-工': round(step_lab, 4),
                    '逐步结转-费': round(step_oh, 4),
                    '逐步结转-合计': round(step_total, 4),
                    '还原后-真料': round(rest_mat, 4),
                    '还原后-真工': round(rest_lab, 4),
                    '还原后-真费': round(rest_oh, 4),
                    '还原后-合计': round(rest_total, 4),
                    '料差异': round(rest_mat - step_mat, 4),
                    '工差异': round(rest_lab - step_lab, 4),
                    '费差异': round(rest_oh - step_oh, 4),
                })
        
        comp_df = pd.DataFrame(comparison)
        st.dataframe(comp_df, use_container_width=True, height=400)
        
        # 可视化对比
        st.markdown("##### 成本结构对比")
        
        # 选择物料进行对比
        selected_mats = st.multiselect(
            "选择要对比的物料",
            options=comp_df['物料编码'].tolist(),
            default=comp_df['物料编码'].tolist()[:3]
        )
        
        if selected_mats:
            import plotly.graph_objects as go
            
            fig = go.Figure()
            
            for mat in selected_mats:
                row = comp_df[comp_df['物料编码'] == mat].iloc[0]
                
                fig.add_trace(go.Bar(
                    name=f"{mat} - 逐步结转",
                    x=['料', '工', '费'],
                    y=[row['逐步结转-料'], row['逐步结转-工'], row['逐步结转-费']],
                    marker_color='lightblue'
                ))
                
                fig.add_trace(go.Bar(
                    name=f"{mat} - 还原后",
                    x=['料', '工', '费'],
                    y=[row['还原后-真料'], row['还原后-真工'], row['还原后-真费']],
                    marker_color='darkblue'
                ))
            
            fig.update_layout(
                barmode='group',
                title='逐步结转法 vs 成本还原对比',
                yaxis_title='金额',
                height=500
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        # 下载按钮
        st.divider()
        excel = to_excel({
            '成本还原对比': comp_df,
            '还原后明细': result['detail_df']
        })
        st.download_button("📥 下载还原结果", excel, "成本还原结果.xlsx",
                          use_container_width=True)

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
