"""
字段匹配与数据表生成工具
========================
从 app.py 中提取的纯 Python 逻辑函数，零 Streamlit 依赖。
提供给 pipeline_cli.py 和未来的其他 CLI 工具使用。
"""

import re
import pandas as pd
import numpy as np


# ============================================================================
# 文件类型检测
# ============================================================================

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


# ============================================================================
# 智能字段匹配
# ============================================================================

def smart_match(columns, file_type):
    """智能字段匹配 — 正则权重匹配，返回 {标准字段: 原始列名}"""
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
            '入库数量': [(r'入库.*数量$', 10), (r'完工.*数量$', 9), (r'入库数', 8),
                       (r'实收数量', 8), (r'^数量$', 5), (r'入库量', 7)],
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
            '销售金额': [(r'销售.*金额$', 10), (r'出库.*金额$', 9), (r'销售.*收入$', 9),
                       (r'销售.*价值', 8), (r'销售收入', 8), (r'出库金额', 8),
                       (r'^金额$', 3), (r'含税金额', 6)],
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


# ============================================================================
# 边表 & 路径表
# ============================================================================

def create_edge_table(calc):
    """生成边表：W 矩阵中相邻流转关系"""
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
                is_mat2order = (source in material_nodes) and (target in order_nodes)
                is_order2mat = (source in order_nodes) and (target in material_nodes)
                rows.append({
                    '边ID': f"E{edge_id:03d}",
                    '起点': source,
                    '起点类型': '工单' if source in order_nodes else '物料',
                    '终点': target,
                    '终点类型': '工单' if target in order_nodes else '物料',
                    '消耗比例': f"{weight:.2%}" if is_mat2order else "—",
                    '产出比例': f"{weight:.2%}" if is_order2mat else "—",
                })
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    except Exception:
        return None


def create_path_table(calc):
    """生成路径表：根到叶的完整流转路径"""
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
