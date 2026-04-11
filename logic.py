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
        
        # 保存中间计算结果
        self.W_matrix = None
        self.all_nodes = None
        self.material_nodes = None
        self.order_nodes = None
        self.node_index = None
        self.X_total = None  # 总成本矩阵
        
    def load_data(self, initial_file, purchase_file, io_file, labor_file,
                  initial_map, purchase_map, io_map, labor_map):
        """加载并清洗数据"""
        start_time = time.time()
        
        # 期初（可选）
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
        df_io['产品完工数量'] = pd.to_numeric(df_io['产品完工数量'], errors='coerce').fillna(0)
        df_io['材料领用数量'] = pd.to_numeric(df_io['材料领用数量'], errors='coerce').fillna(0)
        
        if '在产品数量' in df_io.columns:
            df_io['在产品数量'] = pd.to_numeric(df_io['在产品数量'], errors='coerce').fillna(0)
        else:
            df_io['在产品数量'] = 0
        
        df_io_clean = df_io.groupby(['工单号', '产品编码', '材料编码']).agg({
            '材料领用数量': 'sum',
            '产品完工数量': 'sum',
            '在产品数量': 'sum'
        }).reset_index()
        
        self.io_df = df_io_clean
        
        # 工单费用（可选）
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
            self.labor_df = pd.DataFrame(columns=['工单号', '人工', '制费'])
        
        self.performance_log['数据清洗'] = time.time() - start_time
        return True
    
    def calculate(self, finished_df=None, finished_map=None):
        """执行核心成本计算（修正版）
        
        Parameters:
        -----------
        finished_df : DataFrame, optional
            产成品入库明细表，包含产品编码和入库数量
        finished_map : dict, optional
            字段映射 {原列名: 标准名}
        """
        from scipy import sparse
        from scipy.sparse.linalg import spsolve, splu
        
        total_start = time.time()
        
        # Step 1: 构建节点
        t0 = time.time()
        products = set(self.io_df['产品编码'].unique())
        materials = set(self.io_df['材料编码'].unique())
        init_materials = set(self.initial_df['物料编码'].unique())
        pur_materials = set(self.purchase_df['物料编码'].unique())
        
        all_materials = products | materials | init_materials | pur_materials
        all_orders = set(self.io_df['工单号'].unique())
        
        self.material_nodes = sorted(list(all_materials))
        self.order_nodes = sorted(list(all_orders))
        self.all_nodes = self.material_nodes + self.order_nodes
        n = len(self.all_nodes)
        
        self.node_index = {node: i for i, node in enumerate(self.all_nodes)}
        self.performance_log['构建节点'] = time.time() - t0
        
        # Step 2: 计算基础数量
        t0 = time.time()
        
        # 期初数据
        init_qty = {}
        init_amt = {}
        for mat in self.material_nodes:
            init_qty[mat] = 0
            init_amt[mat] = 0
        
        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码'])
            if mat in init_qty:
                init_qty[mat] = row['期初数量']
                init_amt[mat] = row['期初金额']
        
        # 采购数据
        pur_qty = {}
        for mat in self.material_nodes:
            pur_qty[mat] = 0
        
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in pur_qty:
                pur_qty[mat] = row['采购数量']
        
        # 生产数据（完工+在产）
        # 【重要区分】
        # 1. finished_qty_for_w: 用于W矩阵计算（必须来自投入产出明细）
        # 2. finished_qty_for_report: 用于收发存报表（来自入库明细，如果提供）
        
        finished_qty_for_w = {}  # 用于W矩阵
        finished_qty_for_report = {}  # 用于报表
        wip_qty = {}
        
        # 初始化
        for mat in self.material_nodes:
            finished_qty_for_w[mat] = 0
            finished_qty_for_report[mat] = 0
            wip_qty[mat] = 0
        
        # 1. W矩阵的完工数量：始终从投入产出获取
        prod_summary = self.io_df.groupby('产品编码')['产品完工数量'].sum().reset_index()
        for _, row in prod_summary.iterrows():
            mat = str(row['产品编码'])
            if mat in finished_qty_for_w:
                finished_qty_for_w[mat] = row['产品完工数量']
        
        # 2. 报表的完工数量：优先从入库明细获取，否则用投入产出
        if finished_df is not None and finished_map:
            df_fin = finished_df.rename(columns=finished_map)
            df_fin['产品编码'] = df_fin['产品编码'].astype(str).str.strip()
            df_fin['入库数量'] = pd.to_numeric(df_fin['入库数量'], errors='coerce').fillna(0)
            fin_summary = df_fin.groupby('产品编码')['入库数量'].sum().to_dict()
            for mat in self.material_nodes:
                finished_qty_for_report[mat] = fin_summary.get(mat, 0)
        else:
            # 没有入库明细，用投入产出的
            for mat in self.material_nodes:
                finished_qty_for_report[mat] = finished_qty_for_w[mat]
        
        # 在产品数量从投入产出获取
        wip_summary = self.io_df.groupby('产品编码')['在产品数量'].sum().reset_index()
        for _, row in wip_summary.iterrows():
            mat = str(row['产品编码'])
            if mat in wip_qty:
                wip_qty[mat] = row['在产品数量']
        
        self.performance_log['计算数量'] = time.time() - t0
        
        # Step 3: 构建W矩阵
        t0 = time.time()
        row_indices = []
        col_indices = []
        data = []
        
        # 可供发出数量 = 期初 + 采购 + 完工（W矩阵用投入产出的完工数量）
        available_qty = {}
        for mat in self.material_nodes:
            available_qty[mat] = init_qty[mat] + pur_qty[mat] + finished_qty_for_w[mat]
        
        # 按工单+产品聚合完工数量（避免重复）
        order_prod_finished = self.io_df.groupby(['工单号', '产品编码'])['产品完工数量'].sum().reset_index()

        # 计算每个工单的总完工数量
        order_total_finished = order_prod_finished.groupby('工单号')['产品完工数量'].sum()

        # 构建产出关系（按完工数量比例分配）
        for _, row in order_prod_finished.iterrows():
            order = str(row['工单号'])
            prod = str(row['产品编码'])
            finished = row['产品完工数量']
            
            if order not in self.node_index or prod not in self.node_index:
                continue
            
            # 产出比例 = 该产品完工数量 / 工单总完工数量
            total = order_total_finished.get(order, 0)
            if total > 0:
                ratio = finished / total
            else:
                # 如果总完工为0，平均分配（或根据业务逻辑处理）
                n_products = len(order_prod_finished[order_prod_finished['工单号']==order])
                ratio = 1.0 / n_products if n_products > 0 else 0
            
            row_indices.append(self.node_index[prod])
            col_indices.append(self.node_index[order])
            data.append(ratio)  # 分摊比例，不是1.0！
        
        # 物料→工单（消耗）
        for _, row in self.io_df.iterrows():
            order = str(row['工单号'])
            material = str(row['材料编码'])
            issue = row['材料领用数量']
            
            if material not in self.node_index or order not in self.node_index:
                continue
            
            avail = available_qty.get(material, 0)
            if avail > 0:
                ratio = issue / avail
                if ratio > 1:
                    ratio = 1.0
            else:
                ratio = 0
            
            row_indices.append(self.node_index[order])
            col_indices.append(self.node_index[material])
            data.append(ratio)
        
        W_sparse = sparse.csr_matrix((data, (row_indices, col_indices)), shape=(n, n))
        W_dense = W_sparse.toarray()
        self.W_matrix = W_dense
        
        self.performance_log['构建矩阵'] = time.time() - t0
        
        # Step 4: 构建阀门矩阵D
        D_mat = np.ones(n)
        D_loh = np.ones(n)
        
        order_production = self.io_df.groupby('工单号').agg({
            '产品完工数量': 'sum',
            '在产品数量': 'sum'
        }).reset_index()
        
        for _, row in order_production.iterrows():
            order = str(row['工单号'])
            if order not in self.node_index:
                continue
            
            finished = row['产品完工数量']
            wip = row['在产品数量']
            total = finished + wip
            
            if total > 0:
                D_mat[self.node_index[order]] = finished / total
            else:
                D_mat[self.node_index[order]] = 1.0
        
        for mat in self.material_nodes:
            D_mat[self.node_index[mat]] = 1.0
            D_loh[self.node_index[mat]] = 1.0
        
        # Step 5: 构建F矩阵（外部投入）
        t0 = time.time()
        F = np.zeros((n, 3))
        
        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码'])
            if mat in self.node_index:
                F[self.node_index[mat], 0] += row['期初金额']
        
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in self.node_index:
                F[self.node_index[mat], 0] += row['采购金额']
        
        for _, row in self.labor_df.iterrows():
            order = str(row['工单号'])
            if order in self.node_index:
                F[self.node_index[order], 1] = row['人工']
                F[self.node_index[order], 2] = row['制费']
        
        self.performance_log['构建F矩阵'] = time.time() - t0
        
        # Step 6: 求解总成本 X = (I-WD)^(-1) * F
        t0 = time.time()
        I = np.eye(n)
        
        # 材料路径：I - W * D_mat
        W_mat = W_dense * D_mat
        A_mat = I - W_mat
        
        # 工费路径：I - W（D_loh=I）
        A_loh = I - W_dense
        
        try:
            X_mat = np.linalg.solve(A_mat, F[:, 0:1])
            X_loh = np.linalg.solve(A_loh, F[:, 1:3])
        except np.linalg.LinAlgError:
            X_mat = np.linalg.pinv(A_mat) @ F[:, 0:1]
            X_loh = np.linalg.pinv(A_loh) @ F[:, 1:3]
        
        X = np.hstack([X_mat, X_loh])
        X = np.round(X, 4)
        self.X_total = X
        
        self.performance_log['矩阵求解'] = time.time() - t0
        self.performance_log['矩阵维度'] = n
 
        # Step 7: 计算WIP成本（公式：W × (1-D) × X_mat）
        # 注意：只用材料列 X_mat，工费不计算WIP
        t0 = time.time()
        I_minus_D = 1 - D_mat
        wip_at_order_mat = X_mat[:, 0] * I_minus_D  # (1-D) × X_mat
        
        # WIP_full = W × wip_at_order_mat，但结果是n×1
        WIP_mat_col = W_dense.dot(wip_at_order_mat)
        
        # 构建完整的WIP矩阵（3列：料、工、费，工费为0）
        WIP_full = np.zeros((n, 3))
        WIP_full[:, 0] = WIP_mat_col  # 只有材料有WIP
        WIP_full = np.round(WIP_full, 4)
        self.performance_log['在产品计算'] = time.time() - t0
        
        # Step 8: 结果整理
        # 8.1 收发存汇总表
        issue_summary = self.io_df.groupby('材料编码')['材料领用数量'].sum().to_dict()
        
        result_list = []
        for mat in self.material_nodes:
            idx = self.node_index[mat]
            
            # 总成本（从X矩阵直接取）
            total_mat = X[idx, 0]
            total_lab = X[idx, 1] 
            total_oh = X[idx, 2]
            total_cost = total_mat + total_lab + total_oh
            
            # WIP（在产品）
            wip_mat = WIP_full[idx, 0]
            wip_total = wip_mat
            
            # 完工产品成本
            finished_cost = total_cost - wip_total
            
            # 基础数量
            init_q = init_qty[mat]
            init_a = init_amt[mat]
            pur_q = pur_qty[mat]
            fin_q = finished_qty_for_report[mat]  # 报表用入库明细的完工数量
            
            # 收入数量 = 采购 + 完工
            receipt_qty = pur_q + fin_q
            
            # 收入金额 = 完工产品成本 - 期初金额（剔除以前期间的成本）
            receipt_amt = finished_cost - init_a
            
            issue_q = issue_summary.get(mat, 0)
            
            # 总数量 = 期初 + 收入
            total_q = init_q + receipt_qty
            
            unit_price = total_cost / total_q if total_q > 0 else 0
            issue_amt = issue_q * unit_price
            
            result_list.append({
                '物料编码': mat,
                '期初数量': round(init_q, 4),
                '期初金额': round(init_a, 4),
                '本期采购数量': round(pur_q, 4),
                '本期完工数量': round(fin_q, 4),
                '本期收入数量': round(receipt_qty, 4),
                '本期收入金额': round(receipt_amt, 4),
                '本月发出数量': round(issue_q, 4),
                '发出单价': round(unit_price, 4),
                '发出金额': round(issue_amt, 4),
                '期末数量': round(total_q - issue_q, 4),
                '期末金额': round(total_cost - issue_amt, 4),
            })
        
        # 8.2 工单投入产出明细（从X矩阵直接取工单成本）
        # 先按工单+产品聚合，获取每个工单的实际完工/在产（避免重复行）

        order_prod_summary = self.io_df.groupby(['工单号', '产品编码']).agg({
            '产品完工数量': 'sum',
            '在产品数量': 'sum'
        }).reset_index()

        order_detail_list = []

        for _, row in order_prod_summary.iterrows():
            order = str(row['工单号'])
            product = str(row['产品编码'])
            
            if order not in self.node_index or product not in self.node_index:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self.node_index[product]
            
            # 获取 W 矩阵产出比例（该工单对产品产出的贡献比例）
            w_ratio = W_dense[prod_idx, order_idx]
            
            # 完工率（来自 D 矩阵，只用于成本分配！）
            finished_ratio = D_mat[order_idx]
            
            # 实际数量（直接取原始数据，不乘任何比例！）
            finished_q = row['产品完工数量']  # ✅ 实际完工数量
            wip_q = row['在产品数量']        # ✅ 实际在产数量
            total_q = finished_q + wip_q
            
            # 该工单的总成本（从 X 矩阵取）
            order_total_mat = X[order_idx, 0]  # 材料
            order_total_lab = X[order_idx, 1]  # 人工
            order_total_oh = X[order_idx, 2]   # 制费
            
            # 按 W 比例和 D 比例分配成本
            # 材料成本（按完工率 D 分配）
            finished_mat = w_ratio * order_total_mat * finished_ratio      # 完工材料
            wip_mat = w_ratio * order_total_mat * (1 - finished_ratio)     # 在产材料

            # 人工/制费（全额给完工，不乘 D！）
            finished_lab = w_ratio * order_total_lab    # ✅ 全额人工
            finished_oh = w_ratio * order_total_oh      # ✅ 全额制费
            wip_lab = 0                                 # 在产无人工
            wip_oh = 0                                  # 在产无制费
            
            order_detail_list.append({
                '工单号': order,
                '产品编码': product,
                'W矩阵贡献比例': round(w_ratio, 4),  # 显示该工单贡献了多少
                '完工率(D矩阵)': round(finished_ratio, 4),  # 显示该工单完工比例
                '完工数量': round(finished_q, 2),      # ✅ 实际数量，不乘比例
                '在产品数量': round(wip_q, 2),         # ✅ 实际数量
                '在产品材料费': round(wip_mat, 2),
                '完工产品材料费': round(finished_mat, 2),
                '完工产品直接人工': round(finished_lab, 2),
                '完工产品制造费用': round(finished_oh, 2),
                '工单总成本': round(wip_mat + finished_mat + finished_lab + finished_oh, 2),  # 调试用
            })
        
        # 8.3 成本明细
        detail_list = []
        for i, node in enumerate(self.all_nodes):
            is_order = node in self.order_nodes
            detail_list.append({
                '节点': node,
                '类型': '工单' if is_order else '物料',
                '总成本': round(X[i].sum(), 4),
                '料': round(X[i, 0], 4),
                '工': round(X[i, 1], 4),
                '费': round(X[i, 2], 4),
            })
        
        self.performance_log['总计算时间'] = time.time() - total_start
        
        self.result = {
            '收发存': pd.DataFrame(result_list),
            '工单明细': pd.DataFrame(order_detail_list),
            '成本明细': pd.DataFrame(detail_list),
            'nodes': self.all_nodes
        }

        with open('debug_cost.txt', 'w', encoding='utf-8') as f:
            f.write("=== 成本调试报告 ===\n")
            f.write(f"F矩阵总和: {np.sum(F):.2f}\n")
            f.write(f"X矩阵总和: {np.sum(X):.2f}\n")
            f.write(f"放大倍数: {np.sum(X)/np.sum(F):.2f}\n")
            f.write("\n前10个节点成本:\n")
            for i in range(min(10, len(self.all_nodes))):
                f.write(f"{self.all_nodes[i]}: F={F[i]}, X={X[i]}\n")

        with open('debug_w_matrix.txt', 'w', encoding='utf-8') as f:
            f.write("=== W 矩阵列和检查 ===\n")
            
            col_sums = np.sum(W_dense, axis=0)
            max_col = np.max(col_sums)
            max_idx = np.argmax(col_sums)
            
            f.write(f"最大列和: {max_col:.4f} (应 ≤ 1.0)\n")
            f.write(f"问题节点: {self.all_nodes[max_idx]}\n\n")
            
            # 列出所有列和 > 1 的节点
            f.write("列和大于1的节点:\n")
            for i in range(len(self.all_nodes)):
                if col_sums[i] > 1.01:  # 允许1%误差
                    f.write(f"{self.all_nodes[i]}: {col_sums[i]:.4f}\n")
                    # 显示这一列的非零元素（谁领用了它）
                    non_zero = np.where(W_dense[:, i] > 0.001)[0]
                    f.write(f"  被领用情况:\n")
                    for j in non_zero:
                        f.write(f"    {self.all_nodes[j]}: {W_dense[j, i]:.4f}\n")
                    f.write("\n")
            
            # 检查是否有自环（对角线元素）
            f.write("\n=== 对角线检查 ===\n")
            diag = np.diag(W_dense)
            if np.any(diag > 0.001):
                f.write("警告：存在自环！\n")
                for i in range(len(self.all_nodes)):
                    if diag[i] > 0.001:
                        f.write(f"{self.all_nodes[i]}: {diag[i]:.4f}\n")

        return self.result
    
    def calculate_cost_restoration(self, step_cost_df, step_map):
        """成本还原计算"""
        from scipy import sparse
        from scipy.sparse.linalg import spsolve
        
        if self.W_matrix is None:
            raise ValueError("请先执行成本核算以构建流转矩阵 W")
        
        n = len(self.all_nodes)
        W_dense = self.W_matrix
        
        # 构建逐步结转法的 F 矩阵
        F_step = np.zeros((n, 3))
        
        for _, row in step_cost_df.iterrows():
            mat = str(row[step_map['物料编码']]).strip()
            if mat in self.node_index:
                idx = self.node_index[mat]
                material_cost = pd.to_numeric(row[step_map['料']], errors='coerce') or 0
                labor_cost = pd.to_numeric(row[step_map['工']], errors='coerce') or 0
                overhead_cost = pd.to_numeric(row[step_map['费']], errors='coerce') or 0
                
                F_step[idx, 0] = material_cost
                F_step[idx, 1] = labor_cost
                F_step[idx, 2] = overhead_cost
        
        # 计算还原成本
        I = np.eye(n)
        A = I - W_dense
        
        try:
            X_restored = np.linalg.solve(A, F_step)
        except np.linalg.LinAlgError:
            X_restored = np.linalg.pinv(A) @ F_step
        
        X_restored = np.round(X_restored, 4)
        
        restored_costs = []
        detail_list = []
        
        for i, node in enumerate(self.all_nodes):
            cost_mat, cost_lab, cost_oh = X_restored[i]
            total = cost_mat + cost_lab + cost_oh
            
            restored_costs.append({
                '节点': node,
                '真料': cost_mat,
                '真工': cost_lab,
                '真费': cost_oh,
                '合计': total
            })
            
            detail_list.append({
                '节点': node,
                '类型': '物料' if node in self.material_nodes else '工单',
                '真料': cost_mat,
                '真工': cost_lab,
                '真费': cost_oh,
                '合计': total
            })
        
        return {
            'restored_costs': X_restored,
            'restored_df': pd.DataFrame(restored_costs),
            'detail_df': pd.DataFrame(detail_list),
            'F_step': F_step
        }
    
    def get_performance(self):
        return self.performance_log

def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)
    return output
