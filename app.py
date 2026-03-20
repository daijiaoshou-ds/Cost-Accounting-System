import streamlit as st
import pandas as pd
from logic import CostCalculator, to_excel
import time

st.set_page_config(page_title="矩阵成本核算引擎", page_icon="🧮", layout="wide")

st.title("🧮 矩阵成本核算引擎")

def smart_match(columns, file_type):
    """
    更智能的字段匹配算法
    返回：{标准名: 实际列名}
    """
    cols = list(columns)
    result = {}
    
    # 定义权重（越具体的权重越高）
    patterns = {
        '期初': {
            '物料编码': [
                (r'物料.*编码$', 10), (r'料号$', 9), (r'品号$', 8), 
                (r'物料代码', 7), (r'材料编码', 6), (r'编码$', 5)
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
                (r'物料.*编码$', 10), (r'料号$', 9), (r'品号$', 8),
                (r'物料代码', 7), (r'编码$', 5)
            ],
            '采购数量': [
                (r'采购.*数量$', 10), (r'入库.*数量$', 9), (r'采购数', 8),
                (r'数量$', 3)
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
            '材料编码': [
                (r'材料.*编码$', 10), (r'物料.*编码', 9), (r'领料编码', 8),
                (r'材料代码', 7)
            ],
            '材料领用数量': [
                (r'领用.*数量$', 10), (r'领料.*数量$', 9), (r'消耗.*数量$', 8),
                (r'用量$', 7), (r'领用数', 6)
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
        }
    }
    
    import re
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
            cols.remove(best_match)  # 已匹配的列不再参与其他匹配
    
    return result

# 初始化session state
if 'result' not in st.session_state:
    st.session_state.result = None
if 'perf' not in st.session_state:
    st.session_state.perf = {}

# 侧边栏上传
with st.sidebar:
    st.header("📤 数据上传")
    
    uploaded = {}
    mappings = {}
    
    # 必须上传的文件
    required_files = ['purchase', 'io']
    # 可选上传的文件
    optional_files = ['initial', 'labor']
    
    file_configs = [
        ('purchase', '采购入库', '采购', True),  # True = 必须
        ('io', '投入产出明细', '投入产出', True),
        ('initial', '期初结存', '期初', False),  # False = 可选
        ('labor', '工单人工制费', '工单费用', False)
    ]
    
    for key, title, ptype, is_required in file_configs:
        with st.expander(f"{'*' if is_required else ' '}{title}", expanded=(key=='purchase')):
            help_text = f"{'【必须上传】' if is_required else '【可选上传】'}上传{title}"
            f = st.file_uploader(help_text, type=['xlsx', 'xls'], key=key)
            if f:
                try:
                    df = pd.read_excel(f)
                    st.caption(f"✓ {len(df)} 行 × {len(df.columns)} 列")
                    
                    # 智能匹配
                    auto_map = smart_match(df.columns, ptype)
                    
                    if auto_map:
                        st.success(f"自动匹配 {len(auto_map)} 个字段")
                        # 显示匹配结果（可调整）
                        with st.container():
                            final_map = {}
                            for std_col, matched_col in auto_map.items():
                                # 让用户确认或修改
                                options = [matched_col] + [c for c in df.columns if c != matched_col]
                                selected = st.selectbox(
                                    f"{std_col}", 
                                    options, 
                                    key=f"{key}_{std_col}"
                                )
                                final_map[selected] = std_col
                            mappings[key] = final_map
                    else:
                        st.warning("未自动识别，请手动选择")
                        # 手动选择逻辑...
                    
                    uploaded[key] = f
                except Exception as e:
                    st.error(f"读取失败: {e}")
            elif is_required:
                st.info(f"⚠️ 请上传 {title} 文件")

# 主界面
# 只需要采购入库和投入产出明细
ready = 'purchase' in uploaded and 'io' in uploaded

# 添加计算模式选择
calc_mode = st.radio(
    "计算模式",
    ["成本核算", "材料穿透追溯"],
    index=0,
    help="成本核算：计算料工费成本；材料穿透追溯：追踪原材料到最终成品的传递路径"
)

if st.button(f"🚀 执行{calc_mode}", type="primary", disabled=not ready, use_container_width=True):
    if ready:
        with st.spinner("计算中..."):
            try:
                calc = CostCalculator()
                
                # 记录总时间
                start = time.time()
                
                # 加载数据（期初和人工制费为可选）
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
                
                if calc_mode == "成本核算":
                    result = calc.calculate()
                else:
                    result = calc.calculate_material_trace()
                
                total_time = time.time() - start
                
                st.session_state.result = result
                st.session_state.perf = calc.get_performance()
                st.session_state.total_time = total_time
                st.session_state.calc_mode = calc_mode
                
                st.success(f"✅ 计算完成！总用时: {total_time:.3f}秒")
                
            except Exception as e:
                st.error(f"错误: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

# 显示性能指标
if st.session_state.perf:
    st.subheader("⚡ 性能监控")
    
    perf = st.session_state.perf
    n_nodes = len(st.session_state.result['nodes']) if st.session_state.result else 0
    
    cols = st.columns(5)  # 改成5列，加上矩阵维度
    
    with cols[0]:
        st.metric("数据清洗", f"{perf.get('数据清洗', 0):.3f}s")
    with cols[1]:
        st.metric("矩阵构建", f"{perf.get('构建矩阵', 0):.3f}s")
    with cols[2]:
        st.metric("矩阵维度", f"{n_nodes}×{n_nodes}")  # 新增
    with cols[3]:
        st.metric("矩阵求解", f"{perf.get('矩阵求解', 0):.3f}s")
    with cols[4]:
        st.metric("总用时", f"{st.session_state.get('total_time', 0):.3f}s")
    
    # 详细性能分析（折叠）
    with st.expander("详细性能分析"):
        for k, v in perf.items():
            st.text(f"{k}: {v:.4f}s")

# 显示结果
if st.session_state.result:
    result = st.session_state.result
    calc_mode = st.session_state.get('calc_mode', '成本核算')
    
    st.subheader(f"📊 {calc_mode}结果")
    
    if calc_mode == "成本核算":
        # 指标卡
        c1, c2, c3 = st.columns(3)
        with c1:
            total = result['收发存']['总成本'].sum()
            st.metric("总成本", f"¥{total:,.2f}")
        with c2:
            max_val = result['收发存']['总成本'].max()
            st.metric("最大单项", f"¥{max_val:,.2f}")
        with c3:
            count = len(result['收发存'])
            st.metric("物料数量", count)
        
        # 标签页
        tab1, tab2 = st.tabs(["📋 收发存汇总", "📈 成本明细"])
        
        with tab1:
            st.dataframe(result['收发存'].sort_values('总成本', ascending=False), 
                        use_container_width=True, height=400)
            
            # 下载
            excel = to_excel({
                '收发存汇总': result['收发存'],
                '成本明细': result['明细']
            })
            st.download_button("📥 下载Excel", excel, "成本核算结果.xlsx")
        
        with tab2:
            st.dataframe(result['明细'].sort_values('总成本', ascending=False),
                        use_container_width=True, height=400)
    
    else:  # 材料穿透追溯模式
        # 指标卡
        c1, c2, c3 = st.columns(3)
        with c1:
            trace_count = len(result['材料传递路径'])
            st.metric("传递路径数", trace_count)
        with c2:
            raw_materials = result['材料传递路径']['原材料编码'].nunique()
            st.metric("原材料种类", raw_materials)
        with c3:
            end_products = result['材料传递路径']['最终成品编码'].nunique()
            st.metric("最终成品种类", end_products)
        
        # 标签页
        tab1, tab2 = st.tabs(["📋 材料传递路径", "📈 路径明细"])
        
        with tab1:
            st.dataframe(result['材料传递路径'].sort_values(['原材料编码', '层级']), 
                        use_container_width=True, height=400)
            
            # 下载
            excel = to_excel({
                '材料传递路径': result['材料传递路径'],
                '路径明细': result['路径明细'] if '路径明细' in result else result['材料传递路径']
            })
            st.download_button("📥 下载Excel", excel, "材料穿透追溯结果.xlsx")
        
        with tab2:
            if '路径明细' in result:
                st.dataframe(result['路径明细'],
                            use_container_width=True, height=400)
            else:
                st.info("暂无详细路径数据")