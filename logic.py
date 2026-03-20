import pandas as pd
import numpy as np
from io import BytesIO
import time
import warnings
warnings.filterwarnings('ignore')

class CostCalculator:
    def __init__(self):
        self.initial_df = None
        self.purchase_df = None
        self.io_df = None
        self.labor_df = None
        self.result = None
        self.performance_log = {}
        
    def load_data(self, initial_file, purchase_file, io_file, labor_file,
                  initial_map, purchase_map, io_map, labor_map):
        """加载并清洗数据（期初和人工制费为可选）"""
        start_time = time.time()
        
        # 期初（可选）- 确保文件对象有效
        has_initial = (initial_file is not None and 
                      hasattr(initial_file, 'read') and 
                      initial_map and 
                      len(initial_map) > 0)
        if has_initial:
            df_init = pd.read_excel(initial_file)
            df_init = df_init.rename(columns=initial_map)
            df_init['物料编码'] = df_init['物料编码'].astype(str).str.strip()
            df_init['期初金额'] = pd.to_numeric(df_init['期初金额'], errors='coerce').fillna(0)
            if '期初数量' in df_init.columns:
                df_init['期初数量'] = pd.to_numeric(df_init['期初数量'], errors='coerce').fillna(0)
            else:
                df_init['期初数量'] = 0
            self.initial_df = df_init.groupby('物料编码').agg({
                '期初金额': 'sum',
                '期初数量': 'sum'
            }).reset_index()
        else:
            # 创建空的期初数据框
            self.initial_df = pd.DataFrame(columns=['物料编码', '期初金额', '期初数量'])
        
        # 采购（必须）
        df_pur = pd.read_excel(purchase_file)
        df_pur = df_pur.rename(columns=purchase_map)
        df_pur['物料编码'] = df_pur['物料编码'].astype(str).str.strip()
        df_pur['采购数量'] = pd.to_numeric(df_pur['采购数量'], errors='coerce').fillna(0)
        df_pur['采购金额'] = pd.to_numeric(df_pur['采购金额'], errors='coerce').fillna(0)
        self.purchase_df = df_pur.groupby('物料编码').agg({
            '采购数量': 'sum',
            '采购金额': 'sum'
        }).reset_index()
        
        # 投入产出（必须）
        df_io = pd.read_excel(io_file)
        df_io = df_io.rename(columns=io_map)
        df_io['工单号'] = df_io['工单号'].astype(str).str.strip()
        df_io['产品编码'] = df_io['产品编码'].astype(str).str.strip()
        df_io['材料编码'] = df_io['材料编码'].astype(str).str.strip()
        df_io['产品完工数量'] = pd.to_numeric(df_io['产品完工数量'], errors='coerce')
        df_io['材料领用数量'] = pd.to_numeric(df_io['材料领用数量'], errors='coerce').fillna(0)
        
        df_io_clean = df_io.groupby(['工单号', '产品编码', '材料编码']).agg({
            '材料领用数量': 'sum',
            '产品完工数量': 'first'
        }).reset_index()
        self.io_df = df_io_clean
        
        # 工单费用（可选）- 确保文件对象有效
        has_labor = (labor_file is not None and 
                    hasattr(labor_file, 'read') and 
                    labor_map and 
                    len(labor_map) > 0)
        if has_labor:
            df_lab = pd.read_excel(labor_file)
            df_lab = df_lab.rename(columns=labor_map)
            df_lab['工单号'] = df_lab['工单号'].astype(str).str.strip()
            df_lab['人工'] = pd.to_numeric(df_lab['人工'], errors='coerce').fillna(0)
            df_lab['制费'] = pd.to_numeric(df_lab['制费'], errors='coerce').fillna(0)
            self.labor_df = df_lab.groupby('工单号')[['人工', '制费']].sum().reset_index()
        else:
            # 创建空的人工制费数据框
            self.labor_df = pd.DataFrame(columns=['工单号', '人工', '制费'])
        
        self.performance_log['数据清洗'] = time.time() - start_time
        return True
    
    def calculate(self):
        """执行核心成本计算"""
        total_start = time.time()
        
        # Step 1: 构建节点
        t0 = time.time()
        products = set(self.io_df['产品编码'].unique())
        materials = set(self.io_df['材料编码'].unique())
        init_materials = set(self.initial_df['物料编码'].unique())
        pur_materials = set(self.purchase_df['物料编码'].unique())
        
        all_materials = products | materials | init_materials | pur_materials
        all_orders = set(self.io_df['工单号'].unique())
        
        material_nodes = sorted(list(all_materials))
        order_nodes = sorted(list(all_orders))
        all_nodes = material_nodes + order_nodes
        n = len(all_nodes)
        
        node_index = {node: i for i, node in enumerate(all_nodes)}
        self.performance_log['构建节点'] = time.time() - t0
        
        # Step 2: 计算可供发出数量
        t0 = time.time()
        available_qty = {}
        available_amount = {}
        
        for mat in material_nodes:
            available_qty[mat] = {'期初': 0, '采购': 0, '生产': 0, '合计': 0}
            available_amount[mat] = 0
        
        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码'])
            if mat in available_qty:
                available_qty[mat]['期初'] = row['期初数量']
                available_amount[mat] = row['期初金额']
        
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in available_qty:
                available_qty[mat]['采购'] = row['采购数量']
        
        production_summary = self.io_df.groupby('产品编码')['产品完工数量'].sum().reset_index()
        for _, row in production_summary.iterrows():
            mat = str(row['产品编码'])
            if mat in available_qty:
                available_qty[mat]['生产'] = row['产品完工数量']
        
        for mat in available_qty:
            available_qty[mat]['合计'] = sum(available_qty[mat].values())
        
        self.performance_log['计算数量'] = time.time() - t0
        
        # Step 3: 构建W矩阵
        t0 = time.time()
        W = np.zeros((n, n))
        
        # 工单→物料（产出）
        order_products = self.io_df.groupby('工单号')['产品编码'].apply(list).reset_index()
        for _, row in order_products.iterrows():
            order = str(row['工单号'])
            if order not in node_index:
                continue
            
            prods = row['产品编码']
            total = 0
            qty_map = {}
            for p in prods:
                p = str(p)
                q = self.io_df[(self.io_df['工单号']==order) & (self.io_df['产品编码']==p)]['产品完工数量'].iloc[0]
                qty_map[p] = q
                total += q
            
            for p, q in qty_map.items():
                if p in node_index:
                    ratio = q/total if total > 0 else 0
                    W[node_index[p], node_index[order]] = ratio
        
        # 物料→工单（消耗）- 关键修改：比例>1时强制设为1
        for _, row in self.io_df.iterrows():
            order = str(row['工单号'])
            material = str(row['材料编码'])
            issue = row['材料领用数量']
            
            if material not in node_index or order not in node_index:
                continue
            
            avail = available_qty.get(material, {}).get('合计', 0)
            if avail > 0:
                ratio = issue / avail
                # 关键修复：如果比例>1，强制设为1
                if ratio > 1:
                    ratio = 1.0
            else:
                ratio = 0
            
            W[node_index[order], node_index[material]] = ratio
        
        self.performance_log['构建矩阵'] = time.time() - t0
        
        # Step 4: F矩阵
        t0 = time.time()
        F = np.zeros((n, 3))
        
        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码'])
            if mat in node_index:
                F[node_index[mat], 0] += row['期初金额']
        
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in node_index:
                F[node_index[mat], 0] += row['采购金额']
        
        for _, row in self.labor_df.iterrows():
            order = str(row['工单号'])
            if order in node_index:
                F[node_index[order], 1] = row['人工']
                F[node_index[order], 2] = row['制费']
        
        self.performance_log['构建F矩阵'] = time.time() - t0
        
        # Step 5: 求解
        t0 = time.time()
        I = np.eye(n)
        try:
            X = np.linalg.solve(I - W, F)
        except np.linalg.LinAlgError as e:
            raise ValueError(f"矩阵求解失败：{e}")
        
        self.performance_log['矩阵求解'] = time.time() - t0
        self.performance_log['总计算时间'] = time.time() - total_start
        
        # Step 6: 结果整理
        issue_summary = self.io_df.groupby('材料编码')['材料领用数量'].sum().to_dict()
        
        result_list = []
        for mat in material_nodes:
            idx = node_index[mat]
            cost_mat, cost_lab, cost_oh = X[idx]
            total = cost_mat + cost_lab + cost_oh
            
            q = available_qty[mat]
            init_qty = q['期初']
            init_amt = available_amount[mat]
            receipt_qty = q['采购'] + q['生产']
            receipt_amt = total - init_amt
            
            issue_qty = issue_summary.get(mat, 0)
            total_qty = q['合计']
            unit_price = total / total_qty if total_qty > 0 else 0
            issue_amt = issue_qty * unit_price
            
            result_list.append({
                '物料编码': mat,
                '期初数量': init_qty,
                '期初金额': init_amt,
                '本期收入数量': receipt_qty,
                '收入金额': receipt_amt,
                '总成本': total,
                '本月发出数量': issue_qty,
                '发出单价': unit_price,
                '发出金额': issue_amt,
                '期末数量': total_qty - issue_qty,
                '期末金额': total - issue_amt,
                '料': cost_mat,
                '工': cost_lab,
                '费': cost_oh
            })
        
        detail_list = []
        for i, node in enumerate(all_nodes):
            detail_list.append({
                '节点': node,
                '类型': '物料' if node in material_nodes else '工单',
                '总成本': X[i].sum(),
                '料': X[i, 0],
                '工': X[i, 1],
                '费': X[i, 2]
            })
        
        self.result = {
            '收发存': pd.DataFrame(result_list),
            '明细': pd.DataFrame(detail_list),
            'nodes': all_nodes
        }
        
        return self.result
    
    def calculate_material_trace(self):
        """计算材料传递路径（原材料 → 最终成品）"""
        total_start = time.time()
        
        # Step 1: 构建节点
        t0 = time.time()
        products = set(self.io_df['产品编码'].unique())
        materials = set(self.io_df['材料编码'].unique())
        
        all_materials = products | materials
        all_orders = set(self.io_df['工单号'].unique())
        
        material_nodes = sorted(list(all_materials))
        order_nodes = sorted(list(all_orders))
        all_nodes = material_nodes + order_nodes
        n = len(all_nodes)
        
        node_index = {node: i for i, node in enumerate(all_nodes)}
        index_node = {i: node for i, node in enumerate(all_nodes)}
        self.performance_log['构建节点'] = time.time() - t0
        
        # Step 2: 计算可供发出数量
        t0 = time.time()
        available_qty = {}
        
        for mat in material_nodes:
            available_qty[mat] = {'采购': 0, '生产': 0, '合计': 0}
        
        # 采购数量
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in available_qty:
                available_qty[mat]['采购'] = row['采购数量']
        
        # 生产入库数量
        production_summary = self.io_df.groupby('产品编码')['产品完工数量'].first().reset_index()
        for _, row in production_summary.iterrows():
            mat = str(row['产品编码'])
            if mat in available_qty:
                available_qty[mat]['生产'] = row['产品完工数量']
        
        for mat in available_qty:
            available_qty[mat]['合计'] = available_qty[mat]['采购'] + available_qty[mat]['生产']
        
        self.performance_log['计算数量'] = time.time() - t0
        
        # Step 3: 构建W矩阵
        t0 = time.time()
        W = np.zeros((n, n))
        
        # 工单→物料（产出）- 只要工单只产出一种物料，产出关系就是100%
        order_products = self.io_df.groupby('工单号')['产品编码'].apply(list).reset_index()
        for _, row in order_products.iterrows():
            order = str(row['工单号'])
            if order not in node_index:
                continue
            
            prods = row['产品编码']
            # 该工单所有产品的完工数量
            qty_map = {}
            for p in prods:
                p = str(p)
                q = self.io_df[(self.io_df['工单号']==order) & (self.io_df['产品编码']==p)]['产品完工数量'].iloc[0]
                qty_map[p] = q
            
            # 产出比例 = 该产品数量 / 该工单该产品数量（实际上相同，所以是1.0）
            for p, q in qty_map.items():
                if p in node_index:
                    # 只要工单产出这个产品，比例就是1.0
                    W[node_index[p], node_index[order]] = 1.0
        
        # 物料→工单（消耗）
        zero_availability_warnings = []
        for _, row in self.io_df.iterrows():
            order = str(row['工单号'])
            material = str(row['材料编码'])
            issue = row['材料领用数量']
            
            if material not in node_index or order not in node_index:
                continue
            
            avail = available_qty.get(material, {}).get('合计', 0)
            if avail > 0:
                ratio = issue / avail
                if ratio > 1:
                    ratio = 1.0
            elif issue > 0:
                # 可供发出数量为0但材料被领用，这是一个数据问题
                # 强制设置为1，表示这个工单消耗了该材料的所有"虚拟"库存
                ratio = 1.0
                if material not in zero_availability_warnings:
                    zero_availability_warnings.append(material)
            else:
                ratio = 0
            
            W[node_index[order], node_index[material]] = ratio
        
        # 记录数据问题
        if zero_availability_warnings:
            self.performance_log['数据警告'] = f"以下物料无可用库存但被领用: {', '.join(zero_availability_warnings[:5])}"
        
        self.performance_log['构建矩阵'] = time.time() - t0
        
        # Step 4: 计算逆矩阵 (I-W)⁻¹
        t0 = time.time()
        I = np.eye(n)
        A = I - W
        
        # 检查矩阵是否接近奇异
        try:
            cond = np.linalg.cond(A)
        except:
            cond = float('inf')
        
        # 尝试多种方法求逆
        inv_matrix = None
        method_used = "直接求逆"
        
        # 方法1: 直接求逆
        try:
            if cond < 1e12:  # 条件数不太大时尝试直接求逆
                inv_matrix = np.linalg.inv(A)
        except np.linalg.LinAlgError:
            pass
        
        # 方法2: 添加正则化（如果直接求逆失败或条件数太大）
        if inv_matrix is None:
            try:
                # 添加小的正则化项，使矩阵可逆
                epsilon = 1e-10
                A_reg = A + epsilon * np.eye(n)
                inv_matrix = np.linalg.inv(A_reg)
                method_used = f"正则化求逆(ε={epsilon})"
            except np.linalg.LinAlgError:
                pass
        
        # 方法3: 使用伪逆（如果上述方法都失败）
        if inv_matrix is None:
            try:
                inv_matrix = np.linalg.pinv(A)
                method_used = "伪逆(Moore-Penrose)"
            except Exception as e:
                raise ValueError(f"矩阵求逆失败，已尝试所有方法：{e}")
        
        self.performance_log['矩阵求逆'] = time.time() - t0
        self.performance_log['矩阵条件数'] = cond
        self.performance_log['求逆方法'] = method_used
        
        # Step 5: 识别原材料和最终成品
        # 原材料：没有被作为产品生产的材料（只有采购，没有生产入库）
        # 最终成品：没有被领用的产品
        
        # 被作为产品生产的材料
        produced_materials = set(self.io_df['产品编码'].unique())
        # 被领用的材料
        consumed_materials = set(self.io_df['材料编码'].unique())
        
        # 原材料 = 被领用的材料 - 被生产出来的材料（纯采购材料）
        raw_materials = consumed_materials - produced_materials
        # 最终成品 = 被生产出来的材料 - 被领用的材料
        end_products = produced_materials - consumed_materials
        
        self.performance_log['识别节点'] = time.time() - t0
        
        # Step 6: 计算传递路径和数量关系
        t0 = time.time()
        trace_list = []
        
        # 用于追踪路径的辅助函数
        def find_path_dfs(start_idx, end_idx, visited=None):
            """使用DFS查找从start到end的所有路径"""
            if visited is None:
                visited = set()
            
            if start_idx == end_idx:
                return [[]]
            
            if start_idx in visited:
                return []
            
            visited.add(start_idx)
            paths = []
            
            # 从start可以到达的下一个节点
            for next_idx in range(n):
                if W[next_idx, start_idx] > 0 and next_idx not in visited:
                    sub_paths = find_path_dfs(next_idx, end_idx, visited.copy())
                    for sub_path in sub_paths:
                        paths.append([next_idx] + sub_path)
            
            return paths
        
        def build_path_string(path_indices):
            """将路径索引转换为可读的字符串"""
            if not path_indices:
                return "直接领用"
            
            path_nodes = [index_node[idx] for idx in path_indices]
            # 路径：工单 → 产品 → 工单 → 产品...
            return " → ".join(path_nodes)
        
        def calculate_level(path_indices):
            """计算传递层级（经过的半成品层数）"""
            if not path_indices:
                return 1
            # 统计经过的产品节点数
            product_count = sum(1 for idx in path_indices if index_node[idx] in produced_materials)
            return product_count + 1
        
        # 对每个原材料-最终成品组合
        for raw_mat in raw_materials:
            if raw_mat not in node_index:
                continue
            raw_idx = node_index[raw_mat]
            
            for end_prod in end_products:
                if end_prod not in node_index:
                    continue
                end_idx = node_index[end_prod]
                
                # 获取传递系数
                coefficient = inv_matrix[end_idx, raw_idx]
                
                if coefficient > 1e-10:  # 只保留有效的传递
                    # 查找所有路径
                    paths = find_path_dfs(raw_idx, end_idx)
                    
                    for path in paths:
                        path_str = build_path_string(path)
                        level = calculate_level(path)
                        
                        # 计算单位成品耗用原材料数量
                        # 使用逆矩阵元素作为系数
                        unit_consumption = coefficient
                        
                        trace_list.append({
                            '原材料编码': raw_mat,
                            '最终成品编码': end_prod,
                            '传递路径': path_str,
                            '层级': level,
                            '单位成品耗用原材料数量': unit_consumption
                        })
        
        self.performance_log['计算路径'] = time.time() - t0
        self.performance_log['总计算时间'] = time.time() - total_start
        
        # 创建结果DataFrame
        trace_df = pd.DataFrame(trace_list)
        
        # 如果没有找到路径，返回空DataFrame
        if trace_df.empty:
            trace_df = pd.DataFrame(columns=['原材料编码', '最终成品编码', '传递路径', '层级', '单位成品耗用原材料数量'])
        
        # 生成路径明细（用于调试或详细查看）
        detail_list = []
        for i, node in enumerate(all_nodes):
            node_type = '物料'
            if node in order_nodes:
                node_type = '工单'
            elif node in raw_materials:
                node_type = '原材料'
            elif node in end_products:
                node_type = '最终成品'
            elif node in produced_materials:
                node_type = '半成品'
            
            detail_list.append({
                '节点': node,
                '类型': node_type,
                '节点索引': i
            })
        
        detail_df = pd.DataFrame(detail_list)
        
        self.result = {
            '材料传递路径': trace_df,
            '路径明细': detail_df,
            'nodes': all_nodes,
            'raw_materials': list(raw_materials),
            'end_products': list(end_products),
            'W_matrix': W,
            'inv_matrix': inv_matrix
        }
        
        return self.result

    def get_performance(self):
        return self.performance_log

def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)
    return output