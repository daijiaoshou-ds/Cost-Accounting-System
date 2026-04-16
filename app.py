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

def create_sankey_graph(calc, result=None):
    """
    材料成本流向可视化 - 双视图
    返回两个函数：聚合视图函数、穿透视图函数
    """
    import plotly.graph_objects as go
    from collections import deque, defaultdict
    import json
    
    try:
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            st.warning("未找到流转矩阵数据")
            return None, None
        
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes
        n = len(all_nodes)
        
        # 颜色配置
        TYPE_FILL = {
            '物料_L0': '#B5D4F4',
            '工单':    '#FAC775',
            '物料_L2': '#9FE1CB',
            '物料_L4': '#CECBF6',
            '物料_L6': '#F5C4B3',
            '物料':    '#D3D1C7',
        }
        TYPE_STROKE = {
            '物料_L0': '#378ADD',
            '工单':    '#BA7517',
            '物料_L2': '#1D9E75',
            '物料_L4': '#7F77DD',
            '物料_L6': '#D85A30',
            '物料':    '#888780',
        }
        
        # 成本数据
        cost_data = {}
        if result and '成本明细' in result:
            for _, row in result['成本明细'].iterrows():
                nd = str(row.get('节点', ''))
                cost_data[nd] = {
                    'total': float(row.get('总成本', 0) or 0),
                    'liao':  float(row.get('材料成本', 0) or 0),
                    'gong':  float(row.get('人工成本', 0) or 0),
                    'fei':   float(row.get('制费成本', 0) or 0),
                    'qty':   float(row.get('数量', 0) or 0),
                }
        
        # 构建图结构
        def build_edges():
            edges = []
            out_map = defaultdict(list)
            in_map = defaultdict(list)
            edge_weight = {}
            for i in range(n):
                for j in range(n):
                    if W[i, j] > 0.001:
                        src, tgt, w = all_nodes[j], all_nodes[i], float(W[i, j])
                        edges.append((src, tgt, w))
                        out_map[src].append((tgt, w))
                        in_map[tgt].append((src, w))
                        edge_weight[(src, tgt)] = w
            return edges, out_map, in_map, edge_weight
        
        edges, out_map, in_map, edge_weight = build_edges()
        
        # BFS确定层级
        node_level = {}
        visited = set()
        queue = deque()
        
        for nd in material_nodes:
            if not in_map[nd]:
                node_level[nd] = 0
                queue.append(nd)
                visited.add(nd)
        
        if not queue and material_nodes:
            first = list(material_nodes)[0]
            node_level[first] = 0
            queue.append(first)
            visited.add(first)
        
        while queue:
            cur = queue.popleft()
            for tgt, _ in out_map[cur]:
                if tgt not in visited:
                    node_level[tgt] = node_level[cur] + 1
                    visited.add(tgt)
                    queue.append(tgt)
        
        for nd in all_nodes:
            if nd not in node_level:
                node_level[nd] = 99
        
        # 聚合视图 - 简洁桑基图
        def _sankey_aggregate():
            flow = defaultdict(float)
            for src, tgt, w in edges:
                src_lvl = node_level.get(src, 0)
                tgt_lvl = node_level.get(tgt, 0)
                src_cost = cost_data.get(src, {}).get('total', 0)
                val = w * src_cost if src_cost else w * 1000
                
                src_type = f"物料_L{src_lvl}" if src in material_nodes else '工单'
                tgt_type = f"物料_L{tgt_lvl}" if tgt in material_nodes else '工单'
                
                if src_type != tgt_type:
                    flow[(src_type, tgt_type)] += val
            
            if not flow:
                return None
            
            all_types = set()
            for (s, t), _ in flow.items():
                all_types.add(s)
                all_types.add(t)
            
            def sort_key(t):
                if t == '工单':
                    return (1, 0)
                if t.startswith('物料_L'):
                    return (0, int(t.split('L')[1]))
                return (2, 0)
            
            node_list = sorted(all_types, key=sort_key)
            type_idx = {t: i for i, t in enumerate(node_list)}
            
            type_cost = defaultdict(float)
            for nd in all_nodes:
                lvl = node_level.get(nd, 0)
                tp = f"物料_L{lvl}" if nd in material_nodes else '工单'
                type_cost[tp] += cost_data.get(nd, {}).get('total', 0)
            
            node_colors = [TYPE_FILL.get(t, '#D3D1C7') for t in node_list]
            node_custom = [f"类型：{t}<br>累计成本：¥{type_cost[t]:,.0f}" for t in node_list]
            
            sources, targets, values, lcolors = [], [], [], []
            for (st, tt), val in flow.items():
                if st in type_idx and tt in type_idx and val > 1:
                    sources.append(type_idx[st])
                    targets.append(type_idx[tt])
                    values.append(val)
                    stroke = TYPE_STROKE.get(st, '#888780')
                    r, g, b = int(stroke[1:3], 16), int(stroke[3:5], 16), int(stroke[5:7], 16)
                    lcolors.append(f'rgba({r},{g},{b},0.4)')
            
            fig = go.Figure(go.Sankey(
                arrangement='fixed',
                orientation='h',
                node=dict(
                    pad=20, thickness=24,
                    line=dict(color='rgba(0,0,0,0.08)', width=1),
                    label=node_list, color=node_colors,
                    customdata=node_custom,
                    hovertemplate='<b>%{label}</b><br>%{customdata}<extra></extra>',
                ),
                link=dict(
                    source=sources, target=targets, value=values,
                    color=lcolors,
                    hovertemplate='%{value:,.0f}<extra></extra>',
                ),
            ))
            fig.update_layout(
                font=dict(family='Microsoft YaHei,sans-serif', size=12, color='#333'),
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(l=30, r=30, t=20, b=20), height=350,
            )
            return fig
        
        # 穿透视图
        def _upstream_subgraph(target):
            sub_nodes = {target}
            queue = deque([target])
            while queue:
                cur = queue.popleft()
                for src, _ in in_map[cur]:
                    if src not in sub_nodes:
                        sub_nodes.add(src)
                        queue.append(src)
            return sub_nodes
        
        def _network_single(target):
            sub_nodes = _upstream_subgraph(target)
            sub_edges = [(s, t, w) for s, t, w in edges if s in sub_nodes and t in sub_nodes]
            
            level_in_sub = {}
            has_outgoing = {s for s, _, _ in sub_edges}
            roots = [nd for nd in sub_nodes if nd not in has_outgoing]
            if not roots:
                roots = [target]
            
            visited = set()
            queue = deque()
            for nd in roots:
                level_in_sub[nd] = 0
                visited.add(nd)
                queue.append(nd)
            while queue:
                cur = queue.popleft()
                for tgt, _ in out_map[cur]:
                    if tgt in sub_nodes and tgt not in visited:
                        level_in_sub[tgt] = level_in_sub[cur] + 1
                        visited.add(tgt)
                        queue.append(tgt)
            for nd in sub_nodes:
                if nd not in level_in_sub:
                    level_in_sub[nd] = 0
            
            level_nodes = defaultdict(list)
            for nd in sorted(sub_nodes):
                level_nodes[level_in_sub[nd]].append(nd)
            
            # 横向布局：层级从左到右排列
            NW, NH = 120, 44
            H_GAP, V_GAP = 80, 36
            MX, MY = 40, 40
            n_levels = max(level_nodes) + 1
            max_col = max(len(v) for v in level_nodes.values())
            
            # 计算SVG尺寸：宽度基于层级，高度基于最大列数
            SVG_W = MX*2 + n_levels*NW + (n_levels-1)*H_GAP
            SVG_H = MY*2 + max_col*NH + (max_col-1)*V_GAP
            
            node_pos = {}
            for lv, nodes in level_nodes.items():
                # x坐标基于层级（从左到右）
                x = MX + lv * (NW + H_GAP)
                # y坐标垂直居中排列
                col_h = len(nodes)*NH + (len(nodes)-1)*V_GAP
                start_y = (SVG_H - col_h) // 2
                for idx, nd in enumerate(nodes):
                    node_pos[nd] = (x, start_y + idx*(NH+V_GAP))
            
            def nj(nd):
                c = cost_data.get(nd, {})
                lvl = node_level.get(nd, 0)
                tp = '工单' if nd in order_nodes else f'物料_L{lvl}'
                d = {'name': nd, 'type': tp, 'total': c.get('total', 0),
                     'liao': c.get('liao', 0), 'gong': c.get('gong', 0),
                     'fei': c.get('fei', 0), 'qty': c.get('qty', 0)}
                return json.dumps(d, ensure_ascii=False).replace('"', '&quot;')
            
            parts = [f'<svg id="sg" width="{SVG_W}" height="{SVG_H}" style="display:block;background:transparent;" xmlns="http://www.w3.org/2000/svg">']
            parts.append('<defs>')
            parts.append('<marker id="ar" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">')
            parts.append('<path d="M2 1L8 5L2 9" fill="none" stroke="#aaa" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>')
            parts.append('</marker>')
            parts.append('<marker id="arh" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">')
            parts.append('<path d="M2 1L8 5L2 9" fill="none" stroke="#D85A30" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>')
            parts.append('</marker>')
            parts.append('</defs>')
            
            for src, tgt, w in sub_edges:
                if src not in node_pos or tgt not in node_pos:
                    continue
                x1, y1 = node_pos[src]
                x2, y2 = node_pos[tgt]
                sx, sy = x1+NW, y1+NH//2
                ex, ey = x2, y2+NH//2
                cx = (sx+ex)//2
                lw = max(1.0, min(3.5, 1.0+w*4))
                parts.append(f'<path class="edge" data-src="{src}" data-tgt="{tgt}" d="M{sx},{sy} C{cx},{sy} {cx},{ey} {ex},{ey}" fill="none" stroke="#C8C5BC" stroke-width="{lw:.1f}" marker-end="url(#ar)" opacity="0.75"/>')
            
            for nd, (x, y) in node_pos.items():
                lvl = node_level.get(nd, 0)
                tp = '工单' if nd in order_nodes else f'物料_L{lvl}'
                fill = TYPE_FILL.get(tp, '#D3D1C7')
                stroke = TYPE_STROKE.get(tp, '#888780')
                c = cost_data.get(nd, {})
                cstr = f"¥{c['total']:,.0f}" if c.get('total') else ''
                disp = nd[:10]+'...' if len(nd)>10 else nd
                is_tgt = 'stroke-width="2.5"' if nd == target else 'stroke-width="1"'
                ty = y + NH//2 - (8 if cstr else 0)
                
                parts.append(f'<g class="nd" data-info="{nj(nd)}" style="cursor:pointer;" onmouseenter="hiN(this)" onmouseleave="uhN(this)">')
                parts.append(f'<rect x="{x}" y="{y}" width="{NW}" height="{NH}" rx="8" fill="{fill}" stroke="{stroke}" {is_tgt}/>')
                parts.append(f'<text x="{x+NW//2}" y="{ty}" text-anchor="middle" dominant-baseline="central" font-size="13" font-weight="500" fill="#333">{disp}</text>')
                if cstr:
                    parts.append(f'<text x="{x+NW//2}" y="{ty+17}" text-anchor="middle" dominant-baseline="central" font-size="11" fill="{stroke}">{cstr}</text>')
                parts.append('</g>')
            parts.append('</svg>')
            
            html = '''<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{margin:0;font-family:"Microsoft YaHei",sans-serif;background:transparent}
#wrap{padding:0}#scroll{overflow:auto;border:1px solid #e8e8e8;border-radius:10px;background:#f8f9fa;padding:16px;box-shadow:inset 0 1px 3px rgba(0,0,0,0.05)}
.nd rect{transition:all .2s ease}.nd:hover rect{opacity:0.9;filter:brightness(0.97)}
.edge{transition:all .2s ease}
#tip{position:fixed;display:none;pointer-events:none;background:#fff;border:1px solid #e8e8e8;border-radius:10px;padding:12px 16px;box-shadow:0 4px 20px rgba(0,0,0,.12);font-size:12px;color:#333;min-width:180px;z-index:999}
.tt{font-weight:600;font-size:14px;margin-bottom:8px;color:#1a1a1a}
.tr{display:flex;justify-content:space-between;gap:20px;margin:4px 0}
.tk{color:#888;font-size:12px}.br{display:flex;align-items:center;gap:8px;margin:4px 0}
.bg{flex:1;height:6px;background:#f0f0f0;border-radius:3px;overflow:hidden}.bf{height:100%;border-radius:3px;transition:width 0.3s ease}
</style></head><body>
<div id="wrap"><div id="scroll">''' + '\n'.join(parts) + '''</div></div>
<div id="tip"></div>
<script>
const tip=document.getElementById('tip');
const f=n=>n?'¥'+n.toLocaleString('zh-CN',{maximumFractionDigits:0}):'—';
const p=(a,t)=>t?(a/t*100).toFixed(1)+'%':'—';
const br=(lb,v,t,col)=>{const pc=t?Math.round(v/t*100):0;return`<div class="br"><span style="width:18px;font-size:11px;color:#aaa">${lb}</span><div class="bg"><div class="bf" style="width:${pc}%;background:${col}"></div></div><span style="font-size:11px;color:#666;min-width:80px;text-align:right">${f(v)} (${p(v,t)})</span></div>`;};
function hiN(el){
const d=JSON.parse(el.dataset.info),nm=d.name;
document.querySelectorAll('.edge').forEach(e=>{const on=e.dataset.src===nm||e.dataset.tgt===nm;e.setAttribute('stroke',on?'#D85A30':'#C8C5BC');e.setAttribute('opacity',on?'1':'0.1');e.setAttribute('marker-end',on?'url(#arh)':'url(#ar)');});
const t=d.total;
let h=`<div class="tt">${d.name}<span style="font-size:11px;font-weight:400;color:#999;margin-left:5px">${d.type}</span></div>`;
if(t){h+=`<div class="tr"><span class="tk">总成本</span><span>${f(t)}</span></div>`;h+=br('料',d.liao,t,'#1D9E75')+br('工',d.gong,t,'#378ADD')+br('费',d.fei,t,'#D85A30');if(d.qty)h+=`<div class="tr" style="margin-top:5px"><span class="tk">数量</span><span>${d.qty.toLocaleString()} 件</span></div><div class="tr"><span class="tk">单价</span><span>¥${(t/d.qty).toFixed(2)}/件</span></div>`;}
else h+='<div style="color:#bbb;font-size:11px">暂无成本数据</div>';
tip.innerHTML=h;tip.style.display='block';
}
function uhN(){
document.querySelectorAll('.edge').forEach(e=>{e.setAttribute('stroke','#C8C5BC');e.setAttribute('opacity','0.75');e.setAttribute('marker-end','url(#ar)');});
tip.style.display='none';
}
document.addEventListener('mousemove',e=>{if(tip.style.display==='block'){tip.style.left=(e.clientX+14)+'px';tip.style.top=(e.clientY-10)+'px';}});
</script></body></html>'''
            return html, SVG_H + 20, len(sub_nodes), sum(1 for nd in sub_nodes if nd in order_nodes)
        
        return _sankey_aggregate, _network_single
        
    except Exception as e:
        import traceback
        st.error(f"生成流向图失败: {e}")
        st.code(traceback.format_exc())
        return None, None


def create_edge_table(calc):
    """生成边表：相邻流转关系
    
    | 边ID | 起点 | 终点 | 消耗比例 | 产出比例 |
    """
    try:
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            return None
        
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes
        n = len(all_nodes)
        
        rows = []
        edge_id = 0
        
        for i in range(n):
            for j in range(n):
                if W[i, j] > 0.001:
                    edge_id += 1
                    source = all_nodes[j]
                    target = all_nodes[i]
                    weight = float(W[i, j])
                    
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
        
    except Exception as e:
        import traceback
        st.error(f"生成边表失败: {e}")
        st.code(traceback.format_exc())
        return None


def create_path_table(calc):
    """生成路径表：完整的根到叶路径，带消耗关系
    
    列顺序：路径ID | 第1层 | 第2层 | ... | 第N层 | 最终成品 | 消耗关系
    """
    try:
        if not hasattr(calc, 'W_matrix') or calc.W_matrix is None:
            return None
        
        W = calc.W_matrix
        all_nodes = calc.all_nodes
        material_nodes = calc.material_nodes
        order_nodes = calc.order_nodes
        n = len(all_nodes)
        
        # 构建图结构，带权重
        out_edges = {node: [] for node in all_nodes}
        in_edges = {node: [] for node in all_nodes}
        edge_weight = {}
        
        for i in range(n):
            for j in range(n):
                if W[i, j] > 0.001:
                    source = all_nodes[j]
                    target = all_nodes[i]
                    weight = float(W[i, j])
                    out_edges[source].append(target)
                    in_edges[target].append(source)
                    edge_weight[(source, target)] = weight
        
        # 找出根节点（初始物料：没有入边）
        roots = [node for node in material_nodes if not in_edges[node]]
        
        if not roots:
            roots = list(material_nodes)
        
        # 找出叶节点（最终成品：没有出边的物料）
        leaves = [node for node in material_nodes if not out_edges[node]]
        
        # DFS找所有路径，同时记录权重
        all_paths = []
        
        def dfs(current, path, weights, visited):
            if current in visited:
                return
            
            new_path = path + [current]
            new_visited = visited | {current}
            
            # 如果是叶节点，记录路径
            if current in leaves and len(new_path) >= 2:
                all_paths.append((new_path, weights))
                return
            
            # 继续DFS
            for next_node in out_edges[current]:
                w = edge_weight.get((current, next_node), 1.0)
                dfs(next_node, new_path, weights + [w], new_visited)
        
        for root in roots:
            dfs(root, [], [], set())
        
        # 计算最大层数
        max_layers = max(len(path) for path, _ in all_paths) if all_paths else 0
        
        # 构建路径表
        rows = []
        for idx, (path, weights) in enumerate(all_paths, 1):
            row = {'路径ID': f"P{idx:03d}"}
            
            # 填充所有层级列
            for i in range(1, max_layers + 1):
                if i <= len(path):
                    row[f'第{i}层'] = path[i - 1]
                else:
                    row[f'第{i}层'] = ''
            
            # 最终成品
            final_product = None
            for node in reversed(path):
                if node in material_nodes:
                    final_product = node
                    break
            row['最终成品'] = final_product if final_product else path[-1]
            
            # 消耗关系
            consume_ratios = []
            for i in range(len(path) - 1):
                src, dst = path[i], path[i + 1]
                w = edge_weight.get((src, dst), 1.0)
                consume_ratios.append(f"{w:.0%}")
            
            if consume_ratios:
                row['消耗关系'] = ' × '.join(consume_ratios)
            else:
                row['消耗关系'] = '—'
            
            rows.append(row)
        
        # 创建DataFrame并确保列顺序
        if rows:
            df = pd.DataFrame(rows)
            layer_cols = [f'第{i}层' for i in range(1, max_layers + 1)]
            col_order = ['路径ID'] + layer_cols + ['最终成品', '消耗关系']
            df = df[col_order]
            return df
        
        return pd.DataFrame()
        
    except Exception as e:
        import traceback
        st.error(f"生成路径表失败: {e}")
        st.code(traceback.format_exc())
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
                try:
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
                            st.error(f"无法读取 ZIP 文件 {f.name}: {e}")
                    else:
                        # 处理单个 Excel 文件
                        file_type, type_name = detect_file_type(f.name)
                        if file_type:
                            auto_detected[file_type] = (f, type_name, f.name)
                            st.success(f"✓ 识别为 {type_name}: {f.name}")
                        else:
                            st.warning(f"⚠️ 无法识别文件类型: {f.name}")
                except Exception as e:
                    st.error(f"处理文件 {f.name} 时出错: {e}")
    
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
    
    # ==================== 步骤2: 计算选项与执行 ====================
    st.divider()
    st.markdown("#### 🚀 执行计算")
    
    ready = 'purchase' in uploaded and 'io' in uploaded
    
    if not ready:
        st.warning("⚠️ 请至少上传【采购入库】和【投入产出明细】文件")
        return
    
    # 计算选项
    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        calculate_step_method = st.checkbox(
            "📊 计算逐步结转法", 
            value=False,
            help="逐步结转法下，人工制费按(I+W×D)×F计算，材料=总成本-人工-制费"
        )
    with col_opt2:
        st.caption("💡 默认使用平行结转法（成本还原）。勾选后将额外计算逐步结转法结果")
    
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
                    result = calc.calculate(finished_df, fin_map, calculate_step_method=calculate_step_method)
                else:
                    result = calc.calculate(calculate_step_method=calculate_step_method)
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
        
        # 结果标签页 - 根据是否有逐步结转法结果动态调整
        has_step_result = '逐步结转_工单明细' in result and '逐步结转_成本明细' in result
        
        if has_step_result:
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📋 收发存汇总", 
                "📊 工单投入产出明细", 
                "📈 成本明细",
                "🕸️ 材料流向图",
                "⛓️ 边表与路径表",
                "🔄 逐步结转法"
            ])
        else:
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "📋 收发存汇总", 
                "📊 工单投入产出明细", 
                "📈 成本明细",
                "🕸️ 材料流向图",
                "⛓️ 边表与路径表"
            ])
        
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
            # 工单明细分两个子Tab：汇总表和明细表
            sub_tab1, sub_tab2 = st.tabs(["📊 投入产出汇总表", "🔍 投入产出明细表"])
            
            with sub_tab1:
                st.markdown("##### 投入产出汇总表（按工单+产品聚合）")
                st.dataframe(result['工单明细'].sort_values('工单号'),
                            use_container_width=True, height=500)
            
            with sub_tab2:
                st.markdown("##### 投入产出明细表（按工单+产品+材料展开）")
                st.caption("追踪每个材料的成本流转：材料 → 工单 → 产品 → 完工/在产")
                if '工单产品材料明细' in result and not result['工单产品材料明细'].empty:
                    st.dataframe(result['工单产品材料明细'].sort_values(['工单号', '产品编码', '材料编码']),
                                use_container_width=True, height=500)
                else:
                    st.info("暂无工单-产品-材料明细数据")
        
        with tab3:
            st.dataframe(result['成本明细'].sort_values('总成本', ascending=False),
                        use_container_width=True, height=500)
        
        with tab4:
            st.markdown("##### 材料成本流向可视化")
            st.caption("📍 聚合视图看整体结构，穿透视图看单品链路")
            
            calc = st.session_state.cost_calc
            if calc:
                agg_func, single_func = create_sankey_graph(calc, result)
                
                if agg_func and single_func:
                    sub_tab1, sub_tab2 = st.tabs(["📊 聚合视图（按层级）", "🔍 穿透视图（单品）"])
                    
                    with sub_tab1:
                        st.caption("将所有节点按层级聚合，看整体成本结构")
                        fig = agg_func()
                        if fig:
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("暂无聚合数据")
                    
                    with sub_tab2:
                        st.caption("选择一个节点，查看它的上游供应链路")
                        candidates = sorted(calc.all_nodes)
                        target = st.selectbox("选择要穿透的节点", candidates, help="选择任意节点查看其上游链路")
                        if target:
                            html, h, node_cnt, order_cnt = single_func(target)
                            st.caption(f"上游链路共 **{node_cnt}** 个节点，**{order_cnt}** 个工单")
                            st.components.v1.html(html, height=h, scrolling=True)
                else:
                    st.info("流向图生成失败")
        
        with tab5:
            st.markdown("##### 边表：相邻流转关系")
            st.caption("起点 → 终点的直接流转关系，已包含在Excel导出中")
            
            calc = st.session_state.cost_calc
            if calc:
                edge_df = create_edge_table(calc)
                if edge_df is not None and not edge_df.empty:
                    st.dataframe(edge_df, use_container_width=True, height=350)
                else:
                    st.info("暂无边数据")
            
            st.divider()
            
            st.markdown("##### 路径表：根到叶的完整路径")
            st.caption("从初始物料到最终成品的完整链条，已包含在Excel导出中")
            
            if calc:
                path_df = create_path_table(calc)
                if path_df is not None and not path_df.empty:
                    st.dataframe(path_df, use_container_width=True, height=350)
                else:
                    st.info("暂无路径数据")
        
        # Tab 6: 逐步结转法结果（如果有）
        if has_step_result:
            with tab6:
                st.markdown("##### 逐步结转法计算结果")
                st.info("""
                **逐步结转法说明**：
                - **人工** = (I + W×D) × F_工，即本期人工加上上一步转入的人工
                - **制费** = (I + W×D) × F_制费，即本期制费加上上一步转入的制费  
                - **材料** = X总成本 - 人工 - 制费，剩余部分为综合材料成本
                """)
                
                sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📊 投入产出汇总表", "🔍 投入产出明细表", "📈 成本明细"])
                
                with sub_tab1:
                    st.markdown("##### 逐步结转法 - 投入产出汇总表")
                    st.dataframe(result['逐步结转_工单明细'].sort_values('工单号'),
                                use_container_width=True, height=500)
                
                with sub_tab2:
                    st.markdown("##### 逐步结转法 - 投入产出明细表")
                    if '逐步结转_工单产品材料明细' in result and not result['逐步结转_工单产品材料明细'].empty:
                        st.dataframe(result['逐步结转_工单产品材料明细'].sort_values(['工单号', '产品编码', '材料编码']),
                                    use_container_width=True, height=500)
                    else:
                        st.info("暂无逐步结转法明细数据")
                
                with sub_tab3:
                    st.markdown("##### 逐步结转法 - 成本明细")
                    st.dataframe(result['逐步结转_成本明细'].sort_values('总成本', ascending=False),
                                use_container_width=True, height=500)
        
        # 下载按钮
        st.divider()
        
        # 准备导出数据
        export_sheets = {
            '收发存汇总': result['收发存'],
            '投入产出汇总表': result['工单明细'],
            '投入产出明细表': result['工单产品材料明细'] if '工单产品材料明细' in result else pd.DataFrame(),
            '成本明细': result['成本明细']
        }
        
        # 添加逐步结转法结果（如果有）
        if has_step_result:
            export_sheets['逐步结转_投入产出汇总表'] = result['逐步结转_工单明细']
            export_sheets['逐步结转_投入产出明细表'] = result.get('逐步结转_工单产品材料明细', pd.DataFrame())
            export_sheets['逐步结转_成本明细'] = result['逐步结转_成本明细']
        
        # 添加边表和路径表
        calc = st.session_state.cost_calc
        if calc:
            edge_df = create_edge_table(calc)
            path_df = create_path_table(calc)
            if edge_df is not None and not edge_df.empty:
                export_sheets['边表_相邻流转关系'] = edge_df
            if path_df is not None and not path_df.empty:
                export_sheets['路径表_根到叶路径'] = path_df
        
        excel = to_excel(export_sheets)
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
        
        # 反转映射用于读取数据
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
                step_mat = row[col_mat]
                step_lab = row[col_lab]
                step_oh = row[col_oh]
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
