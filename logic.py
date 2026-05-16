import pandas as pd
import numpy as np
from io import BytesIO
import time
import warnings
warnings.filterwarnings('ignore')

# ========== 稀疏矩阵支持 ==========
from scipy import sparse
from scipy.sparse.linalg import spsolve

class CostCalculator:
    def __init__(self):
        self.initial_df = None
        self.purchase_df = None
        self.io_df = None
        self.labor_df = None
        self.sales_df = None  # 销售数据
        self.result = None
        self.performance_log = {}
        
        # 保存中间计算结果
        self.W_matrix = None
        self.F_matrix = None  # 原始外部投入矩阵
        self.D_matrix = None  # 阀门矩阵
        self.all_nodes = None
        self.material_nodes = None
        self.order_nodes = None
        self.node_index = None
        self.X_total = None  # 总成本矩阵
        self.available_qty = None  # 可供发出数量（W矩阵用，材料完工口径）
        self.sales_available_qty = None  # 可供发出数量（销售用，产品完工口径）
        
    def load_data(self, initial_file, purchase_file, io_file, labor_file, sales_file=None,
                  initial_map=None, purchase_map=None, io_map=None, labor_map=None, sales_map=None):
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
        
        # 销售数据（可选）
        has_sales = (sales_file is not None and 
                    hasattr(sales_file, 'read') and 
                    sales_map and 
                    len(sales_map) > 0)
        if has_sales:
            df_sales = pd.read_excel(sales_file)
            df_sales = df_sales.rename(columns=sales_map)
            df_sales['物料编码'] = df_sales['物料编码'].astype(str).str.strip()
            df_sales['销售数量'] = pd.to_numeric(df_sales['销售数量'], errors='coerce').fillna(0)
            df_sales['销售批次号'] = df_sales['销售批次号'].astype(str).str.strip()
            self.sales_df = df_sales.groupby(['物料编码', '销售批次号'])['销售数量'].sum().reset_index()
        else:
            self.sales_df = pd.DataFrame(columns=['物料编码', '销售批次号', '销售数量'])
        
        self.performance_log['数据清洗'] = time.time() - start_time
        return True
    
    def calculate(self, finished_df=None, finished_map=None, calculate_step_method=False, calculate_super_restoration=False):
        """执行核心成本计算（修正版）
        
        Parameters:
        -----------
        finished_df : DataFrame, optional
            产成品入库明细表，包含产品编码和入库数量
        finished_map : dict, optional
            字段映射 {原列名: 标准名}
        calculate_step_method : bool, optional
            是否计算逐步结转法（平行结转法的补充），默认为False
        calculate_super_restoration : bool, optional
            是否计算超级成本还原，默认为False
        """
        # 稀疏矩阵模块已在文件顶部导入
        
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
        
        W = sparse.csr_matrix((data, (row_indices, col_indices)), shape=(n, n))
        self.W_matrix = W
        
        # ========== 三道保险（Release必备）==========
        # 保险1：列和校验（防超领导致成本发散）
        col_sums = np.array(W.sum(axis=0)).flatten()
        if np.any(col_sums > 1 + 1e-6):
            bad_nodes = [self.all_nodes[i] for i in np.where(col_sums > 1)[0]]
            raise ValueError(f"以下物料存在超领（领用>可供发出），成本将扭曲：{bad_nodes}")
        
        # 保险2：DAG无环校验（防自循环导致矩阵不可逆）
        diag = W.diagonal()
        if np.any(diag > 1e-6):
            bad_nodes = [self.all_nodes[i] for i in np.where(diag > 1e-6)[0]]
            raise ValueError(f"W矩阵存在自环（如工单领用自己产出的物料），节点：{bad_nodes}")
        
        self.performance_log['构建矩阵'] = time.time() - t0
        
        # 保存可供发出数量
        self.available_qty = available_qty  # W矩阵用：期初 + 采购 + 材料完工数量
        
        # 销售用可供发出数量：期初 + 采购 + 产品完工数量（报表口径）
        sales_available_qty = {}
        for mat in self.material_nodes:
            sales_available_qty[mat] = init_qty[mat] + pur_qty[mat] + finished_qty_for_report[mat]
        self.sales_available_qty = sales_available_qty
        
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
        
        # 保险3：D矩阵范围校验
        if np.any(D_mat < 0) or np.any(D_mat > 1):
            raise ValueError("D矩阵（完工率/转出率）存在非法值，必须在[0,1]之间")
        
        # 保存D矩阵供成本还原使用
        self.D_matrix = D_mat
        
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
        
        self.F_matrix = F  # 保存原始F矩阵供成本还原使用
        self.performance_log['构建F矩阵'] = time.time() - t0
        
        # Step 6: 求解总成本 X = (I-WD)^(-1) * F
        t0 = time.time()
        I = sparse.eye(n)
        
        # D矩阵列缩放：W_mat[:,j] = W[:,j] * D_mat[j]
        W_mat_data = W.data.copy() * D_mat[W.indices]
        W_mat = sparse.csr_matrix((W_mat_data, W.indices.copy(), W.indptr.copy()), shape=(n, n))
        A_mat = I - W_mat
        A_loh = I - W
        
        try:
            # spsolve返回1D数组，需reshape
            X_mat = spsolve(A_mat, F[:, 0]).reshape(-1, 1)
            X_loh = spsolve(A_loh, F[:, 1:3])
        except Exception:
            # Fallback：万一稀疏求解崩了，回退到稠密
            X_mat = np.linalg.solve(A_mat.toarray(), F[:, 0:1])
            X_loh = np.linalg.solve(A_loh.toarray(), F[:, 1:3])
        
        X = np.hstack([X_mat, X_loh])
        X = np.round(X, 4)
        self.X_total = X
        
        self.performance_log['矩阵求解'] = time.time() - t0
        self.performance_log['矩阵维度'] = n
        
        # ==================== 销售成本计算（如果有销售数据）====================
        sales_cost_df = pd.DataFrame()
        if not self.sales_df.empty:
            t0 = time.time()
            
            # 按物料汇总销售数量
            mat_sales = self.sales_df.groupby('物料编码')['销售数量'].sum().to_dict()
            batch_sales = self.sales_df  # 已按(物料, 批次)聚合
            
            # 构建 Sale 矩阵（N×N 对角阵）
            S = np.zeros((n, n))
            for mat, sales_qty in mat_sales.items():
                if mat in self.node_index:
                    idx = self.node_index[mat]
                    avail = self.sales_available_qty.get(mat, 0)
                    if avail > 0:
                        S[idx, idx] = sales_qty / avail
            
            # 构建 Batch 矩阵（M×N）
            batch_list = sorted(batch_sales['销售批次号'].unique())
            m = len(batch_list)
            batch_index = {b: i for i, b in enumerate(batch_list)}
            
            B = np.zeros((m, n))
            for _, row in batch_sales.iterrows():
                mat = str(row['物料编码'])
                batch = str(row['销售批次号'])
                qty = row['销售数量']
                if mat in self.node_index and batch in batch_index:
                    mat_idx = self.node_index[mat]
                    batch_idx = batch_index[batch]
                    total_sales = mat_sales.get(mat, 0)
                    if total_sales > 0:
                        B[batch_idx, mat_idx] = qty / total_sales
            
            # 计算 C = B × S × X
            SX = S.dot(X)  # N×3
            C = B.dot(SX)  # M×3
            
            # 构建销售成本结果
            sales_rows = []
            for _, row in batch_sales.iterrows():
                mat = str(row['物料编码'])
                batch = str(row['销售批次号'])
                qty = row['销售数量']
                if mat not in self.node_index:
                    continue
                mat_idx = self.node_index[mat]
                batch_idx = batch_index.get(batch)
                if batch_idx is None:
                    continue
                
                avail = self.sales_available_qty.get(mat, 0)
                sales_ratio = S[mat_idx, mat_idx]
                batch_ratio = B[batch_idx, mat_idx]
                
                cost_mat = batch_ratio * SX[mat_idx, 0]
                cost_lab = batch_ratio * SX[mat_idx, 1]
                cost_oh = batch_ratio * SX[mat_idx, 2]
                cost_total = cost_mat + cost_lab + cost_oh
                
                sales_rows.append({
                    '销售批次号': batch,
                    '物料编码': mat,
                    '销售数量': round(qty, 2),
                    '可供发出数量': round(avail, 2),
                    '销售占比': f"{sales_ratio:.2%}",
                    '批次占物料销售比': f"{batch_ratio:.2%}",
                    '销售成本_料': round(cost_mat, 2),
                    '销售成本_工': round(cost_lab, 2),
                    '销售成本_费': round(cost_oh, 2),
                    '销售成本_合计': round(cost_total, 2),
                })
            
            sales_cost_df = pd.DataFrame(sales_rows)
            if not sales_cost_df.empty:
                sales_cost_df = sales_cost_df.sort_values(['销售批次号', '物料编码'])
            
            self.performance_log['销售成本计算'] = time.time() - t0
        
        # ==================== 超级成本还原（可选）====================
        super_result = {}
        if calculate_super_restoration:
            t0 = time.time()
            
            # 一次遍历构建维度列表和 super_F 矩阵
            super_dims = []
            mat_dim_indices = []
            loh_dim_indices = []
            dim_info = []
            dim_entries = []  # (dim_idx, node_idx, value)
            
            # 材料维度（期初）
            for _, row in self.initial_df.iterrows():
                mat = str(row['物料编码']).strip()
                amt = row['期初金额']
                if mat in self.node_index and amt > 0:
                    idx = len(super_dims)
                    super_dims.append(f"{mat}_期初")
                    mat_dim_indices.append(idx)
                    dim_entries.append((idx, self.node_index[mat], amt))
                    dim_info.append({
                        '维度名称': f"{mat}_期初", '维度类型': '期初',
                        '来源节点': mat, '原始金额': amt,
                        '说明': f'物料{mat}的期初金额'
                    })
            
            # 材料维度（采购）
            for _, row in self.purchase_df.iterrows():
                mat = str(row['物料编码']).strip()
                amt = row['采购金额']
                if mat in self.node_index and amt > 0:
                    idx = len(super_dims)
                    super_dims.append(f"{mat}_采购")
                    mat_dim_indices.append(idx)
                    dim_entries.append((idx, self.node_index[mat], amt))
                    dim_info.append({
                        '维度名称': f"{mat}_采购", '维度类型': '采购',
                        '来源节点': mat, '原始金额': amt,
                        '说明': f'物料{mat}的采购金额'
                    })
            
            # 工费维度（一次遍历 labor_df）
            for _, row in self.labor_df.iterrows():
                order = str(row['工单号']).strip()
                lab = row['人工']
                oh = row['制费']
                if order in self.node_index:
                    if lab > 0:
                        idx = len(super_dims)
                        super_dims.append(f"{order}_人工")
                        loh_dim_indices.append(idx)
                        dim_entries.append((idx, self.node_index[order], lab))
                        dim_info.append({
                            '维度名称': f"{order}_人工", '维度类型': '人工',
                            '来源节点': order, '原始金额': lab,
                            '说明': f'工单{order}的直接人工'
                        })
                    if oh > 0:
                        idx = len(super_dims)
                        super_dims.append(f"{order}_制费")
                        loh_dim_indices.append(idx)
                        dim_entries.append((idx, self.node_index[order], oh))
                        dim_info.append({
                            '维度名称': f"{order}_制费", '维度类型': '制费',
                            '来源节点': order, '原始金额': oh,
                            '说明': f'工单{order}的制造费用'
                        })
            
            n_dims = len(super_dims)
            
            if n_dims > 0:
                # 构建 super_F 矩阵
                super_F = np.zeros((n, n_dims))
                for d_idx, node_idx, val in dim_entries:
                    super_F[node_idx, d_idx] = val
                
                # 批量求解
                super_X = np.zeros((n, n_dims))
                if len(mat_dim_indices) > 0:
                    try:
                        super_X[:, mat_dim_indices] = spsolve(A_mat, super_F[:, mat_dim_indices])
                    except Exception:
                        super_X[:, mat_dim_indices] = np.linalg.solve(A_mat.toarray(), super_F[:, mat_dim_indices])
                
                if len(loh_dim_indices) > 0:
                    try:
                        super_X[:, loh_dim_indices] = spsolve(A_loh, super_F[:, loh_dim_indices])
                    except Exception:
                        super_X[:, loh_dim_indices] = np.linalg.solve(A_loh.toarray(), super_F[:, loh_dim_indices])
                
                super_X = np.round(super_X, 4)
                
                # 向量化生成结果
                mat_indices = [self.node_index[mat] for mat in self.material_nodes]
                mat_total_cost = X[mat_indices].sum(axis=1)
                mat_super_X = super_X[mat_indices, :]
                
                # 【关键修复】dim_total 只包含物料节点（最终产品），排除工单节点
                dim_total = mat_super_X.sum(axis=0)
                
                # 长格式明细
                mask = np.abs(mat_super_X) >= 0.01
                rows, cols = np.where(mask)
                
                long_format_rows = []
                # 预计算每个(产品, 维度)的维度占比，供销售成本表复用
                product_dim_ratios = {}  # {(mat, dim_name): dim_ratio}
                
                for i in range(len(rows)):
                    r, c = rows[i], cols[i]
                    mat = self.material_nodes[r]
                    dim_name = super_dims[c]
                    amt = float(mat_super_X[r, c])
                    total_cost = float(mat_total_cost[r])
                    dim_total_val = float(dim_total[c])
                    
                    product_ratio = amt / total_cost if total_cost > 0 else 0
                    dim_ratio = amt / dim_total_val if dim_total_val > 0 else 0
                    product_dim_ratios[(mat, dim_name)] = dim_ratio
                    
                    long_format_rows.append({
                        '产品编码': mat,
                        '成本维度': dim_name,
                        '金额': round(amt, 2),
                        '占产品总成本比例': f"{product_ratio:.2%}",
                        '占该维度总金额比例': f"{dim_ratio:.2%}",
                    })
                
                super_detail_long = pd.DataFrame(long_format_rows)
                if not super_detail_long.empty:
                    super_detail_long = super_detail_long.sort_values(['产品编码', '金额'], ascending=[True, False])
                    # 添加成本序号（按产品分组，从1开始）
                    super_detail_long.insert(1, '成本序号', super_detail_long.groupby('产品编码').cumcount() + 1)
                
                # TopN 汇总：与长格式相同格式，但每个产品只保留前5大维度
                if not super_detail_long.empty:
                    super_topn_summary = super_detail_long.groupby('产品编码').head(5).reset_index(drop=True)
                else:
                    super_topn_summary = pd.DataFrame()
                
                # 维度定义表
                super_dim_definition = pd.DataFrame(dim_info)
                
                # 验证
                mat_super_sums = mat_super_X.sum(axis=1)
                max_diff = float(np.max(np.abs(mat_super_sums - mat_total_cost)))
                bad_mask = np.abs(mat_super_sums - mat_total_cost) > 1
                bad_idx = np.where(bad_mask)[0]
                
                verification_rows = []
                for r in bad_idx:
                    verification_rows.append({
                        '产品编码': self.material_nodes[r],
                        '超级还原合计': round(float(mat_super_sums[r]), 2),
                        '标准总成本': round(float(mat_total_cost[r]), 2),
                        '差异': round(float(np.abs(mat_super_sums[r] - mat_total_cost[r])), 2)
                    })
                
                super_verification = pd.DataFrame(verification_rows) if verification_rows else pd.DataFrame()
                
                # ========== 销售成本的超级还原（如果有销售数据）==========
                sales_super_long = pd.DataFrame()
                if not self.sales_df.empty and 'S' in dir() and 'B' in dir():
                    # 利用已计算的 S、B、super_X、super_dims
                    # 计算每个 (批次, 物料, 维度) 的销售成本
                    sales_super_rows = []
                    
                    # 预计算所有 (批次, 物料, 维度) 的金额
                    batch_dim_data = []  # [(batch, mat, dim_name, amt), ...]
                    
                    for _, row in batch_sales.iterrows():
                        mat = str(row['物料编码'])
                        batch = str(row['销售批次号'])
                        if mat not in self.node_index or batch not in batch_index:
                            continue
                        mat_idx = self.node_index[mat]
                        batch_idx = batch_index[batch]
                        batch_ratio = B[batch_idx, mat_idx]
                        sales_ratio = S[mat_idx, mat_idx]
                        
                        for d_idx, dim_name in enumerate(super_dims):
                            amt = batch_ratio * sales_ratio * super_X[mat_idx, d_idx]
                            if abs(amt) >= 0.01:
                                batch_dim_data.append((batch, mat, dim_name, amt))
                    
                    # 计算批次总成本
                    batch_totals = {}
                    for batch, mat, dim_name, amt in batch_dim_data:
                        if batch not in batch_totals:
                            batch_totals[batch] = 0
                        batch_totals[batch] += amt
                    
                    # 构建结果——维度占比直接复用完工成本表（二者应一致）
                    for batch, mat, dim_name, amt in batch_dim_data:
                        batch_total = batch_totals.get(batch, 1)
                        batch_ratio_pct = amt / batch_total if batch_total > 0 else 0
                        # 【关键修复】复用完工成本表的维度占比
                        dim_ratio_pct = product_dim_ratios.get((mat, dim_name), 0)
                        
                        sales_super_rows.append({
                            '销售批次号': batch,
                            '产品编码': mat,
                            '成本维度': dim_name,
                            '金额': round(amt, 2),
                            '占批次总成本比例': f"{batch_ratio_pct:.2%}",
                            '占该维度总金额比例': f"{dim_ratio_pct:.2%}",
                        })
                    
                    if sales_super_rows:
                        sales_super_long = pd.DataFrame(sales_super_rows)
                        sales_super_long = sales_super_long.sort_values(['销售批次号', '金额'], ascending=[True, False])
                        # 添加成本序号（按销售批次号+产品编码分组，从1开始）
                        sales_super_long.insert(2, '成本序号', sales_super_long.groupby(['销售批次号', '产品编码']).cumcount() + 1)
                
                super_result = {
                    '超级还原_完工成本': super_detail_long,
                    '超级还原_TopN汇总': super_topn_summary,
                    '超级还原_维度定义': super_dim_definition,
                    '超级还原_验证差异': super_verification if not super_verification.empty else pd.DataFrame({'说明': ['所有产品勾稽验证通过，最大差异: {:.2f}'.format(max_diff)]}),
                }
                
                # 如果有销售成本的超级还原，加入结果
                if not sales_super_long.empty:
                    super_result['超级还原_销售成本'] = sales_super_long
                
                self.performance_log['超级成本还原'] = time.time() - t0
                self.performance_log['超级还原维度数'] = n_dims
            else:
                super_result = {
                    '超级还原_完工成本': pd.DataFrame(),
                    '超级还原_TopN汇总': pd.DataFrame(),
                    '超级还原_维度定义': pd.DataFrame({'说明': ['未找到任何成本维度数据']}),
                    '超级还原_验证差异': pd.DataFrame(),
                }
        
        # Step 7: 计算WIP成本（公式：W × (1-D) × X_mat）
        # 注意：只用材料列 X_mat，工费不计算WIP
        t0 = time.time()
        I_minus_D = 1 - D_mat
        wip_at_order_mat = X_mat[:, 0] * I_minus_D  # (1-D) × X_mat
        
        # WIP_full = W × wip_at_order_mat，但结果是n×1
        WIP_mat_col = W.dot(wip_at_order_mat)
        
        # 构建完整的WIP矩阵（3列：料、工、费，工费为0）
        WIP_full = np.zeros((n, 3))
        WIP_full[:, 0] = WIP_mat_col  # 只有材料有WIP
        WIP_full = np.round(WIP_full, 4)
        self.performance_log['在产品计算'] = time.time() - t0
        
        # Step 8: 结果整理
        # 8.1 收发存汇总表
        issue_summary = self.io_df.groupby('材料编码')['材料领用数量'].sum().to_dict()
        
        # 按物料汇总销售数量（如果有销售数据）
        sales_summary = {}
        if not self.sales_df.empty:
            sales_summary = self.sales_df.groupby('物料编码')['销售数量'].sum().to_dict()
        
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
            
            # 发出数量 = 生产领用 + 销售出库
            issue_q = issue_summary.get(mat, 0) + sales_summary.get(mat, 0)
            
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
                '本月领用数量': round(issue_summary.get(mat, 0), 4),
                '本月销售数量': round(sales_summary.get(mat, 0), 4),
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
            w_ratio = W[prod_idx, order_idx]
            
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
        
        # 8.2b 工单-产品-材料三维明细（新增）
        # 【核心修正】材料成本分配逻辑：
        # 1. 工单-产品-汇总表的成本是正确的（基于W矩阵和完工率D）
        # 2. 明细表应该按"领用数量权重"把汇总表的成本拆分到各材料行
        # 即：每行材料成本 = 工单-产品总材料成本 × (该行领用数量 / 该工单-产品总领用数量)
        order_prod_mat_detail_list = []
        
        # 获取工单-产品的完工/在产数量
        order_prod_qty = self.io_df.groupby(['工单号', '产品编码']).agg({
            '产品完工数量': 'sum',
            '在产品数量': 'sum'
        }).reset_index()
        order_prod_qty_dict = {}
        for _, row in order_prod_qty.iterrows():
            key = (str(row['工单号']), str(row['产品编码']))
            order_prod_qty_dict[key] = (row['产品完工数量'], row['在产品数量'])
        
        # 第一步：计算每个工单-产品的总领用数量（用于材料成本分配）
        order_prod_total_issue = {}
        for _, row in self.io_df.iterrows():
            order = str(row['工单号'])
            product = str(row['产品编码'])
            issue_qty = row['材料领用数量']
            
            key = (order, product)
            if key not in order_prod_total_issue:
                order_prod_total_issue[key] = 0
            order_prod_total_issue[key] += issue_qty
        
        # 第二步：先计算工单-产品-汇总表的数据（与汇总表一致）
        order_prod_summary_data = {}
        for _, row in order_prod_summary.iterrows():
            order = str(row['工单号'])
            product = str(row['产品编码'])
            
            if order not in self.node_index or product not in self.node_index:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self.node_index[product]
            w_ratio = W[prod_idx, order_idx]
            finished_ratio = D_mat[order_idx]
            
            order_total_mat = X[order_idx, 0]
            order_total_lab = X[order_idx, 1]
            order_total_oh = X[order_idx, 2]
            
            # 与汇总表一致的计算
            finished_mat = w_ratio * order_total_mat * finished_ratio
            wip_mat = w_ratio * order_total_mat * (1 - finished_ratio)
            finished_lab = w_ratio * order_total_lab
            finished_oh = w_ratio * order_total_oh
            
            key = (order, product)
            order_prod_summary_data[key] = {
                'finished_mat': finished_mat,
                'wip_mat': wip_mat,
                'finished_lab': finished_lab,
                'finished_oh': finished_oh,
                'w_ratio': w_ratio,
                'finished_ratio': finished_ratio
            }
        
        # 第三步：遍历生成明细表，按领用数量权重拆分
        for _, row in self.io_df.iterrows():
            order = str(row['工单号'])
            product = str(row['产品编码'])
            material = str(row['材料编码'])
            issue_qty = row['材料领用数量']
            
            if order not in self.node_index or product not in self.node_index or material not in self.node_index:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self.node_index[product]
            mat_idx = self.node_index[material]
            
            # W矩阵比例（用于展示）
            w_mat_to_order = W[order_idx, mat_idx]
            w_order_to_prod = W[prod_idx, order_idx]
            finished_ratio = D_mat[order_idx]
            
            # 获取该工单-产品的汇总数据
            key = (order, product)
            summary = order_prod_summary_data.get(key, {
                'finished_mat': 0, 'wip_mat': 0, 'finished_lab': 0, 'finished_oh': 0,
                'w_ratio': 0, 'finished_ratio': 0
            })
            
            # 获取该工单-产品的总领用数量
            total_issue = order_prod_total_issue.get(key, 1)
            
            # 按领用数量权重拆分成本
            if total_issue > 0:
                weight = issue_qty / total_issue
                finished_mat_cost = summary['finished_mat'] * weight
                wip_mat_cost = summary['wip_mat'] * weight
                finished_lab_cost = summary['finished_lab'] * weight
                finished_oh_cost = summary['finished_oh'] * weight
            else:
                # 如果没有领用数量（只有工费），则材料成本为0，但工费需要分配
                # 按行数平均分配工费
                finished_mat_cost = 0
                wip_mat_cost = 0
                # 获取该工单-产品的行数
                row_count = len([r for r in self.io_df.iterrows() 
                                 if str(r[1]['工单号']) == order and str(r[1]['产品编码']) == product])
                if row_count > 0:
                    finished_lab_cost = summary['finished_lab'] / row_count
                    finished_oh_cost = summary['finished_oh'] / row_count
                else:
                    finished_lab_cost = summary['finished_lab']
                    finished_oh_cost = summary['finished_oh']
            
            # 获取该行的完工/在产数量（直接取原始数据）
            finished_q = row['产品完工数量']
            wip_q = row['在产品数量']
            
            order_prod_mat_detail_list.append({
                '工单号': order,
                '产品编码': product,
                '材料编码': material,
                '材料领用数量': round(issue_qty, 2),
                'W_材料到工单': round(w_mat_to_order, 4),
                'W_工单到产品': round(w_order_to_prod, 4),
                '完工率': round(finished_ratio, 4),
                '完工数量': round(finished_q, 2),
                '在产品数量': round(wip_q, 2),
                '在产品材料费': round(wip_mat_cost, 2),
                '完工产品材料费': round(finished_mat_cost, 2),
                '完工产品直接人工': round(finished_lab_cost, 2),
                '完工产品制造费用': round(finished_oh_cost, 2),
                '成本小计': round(wip_mat_cost + finished_mat_cost + finished_lab_cost + finished_oh_cost, 2),
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
        
        # ==================== 逐步结转法计算（可选）====================
        step_result = {}
        if calculate_step_method:
            t0 = time.time()
            
            # S_人工 = (I + W × D) × F_工
            # S_制费 = (I + W × D) × F_制费
            I_plus_WD = I + W_mat  # I + W × D_mat
            
            S_lab = I_plus_WD.dot(F[:, 1])  # 人工列
            S_oh = I_plus_WD.dot(F[:, 2])   # 制费列
            
            # S_材料 = X_combine总 - S_人工 - S_制费
            X_combine_total = X.sum(axis=1)  # 每行的料+工+费总和
            S_mat = X_combine_total - S_lab - S_oh
            
            # 构建 S 矩阵 (n × 3)
            S = np.column_stack([S_mat, S_lab, S_oh])
            S = np.round(S, 4)
            
            # ==================== DEBUG: 打印矩阵数据 ====================
            debug_log_path = r'D:\桌面\公众号\成本核算系统\log\debug_step_method.txt'
            with open(debug_log_path, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("逐步结转法调试报告\n")
                f.write("=" * 80 + "\n\n")
                
                # 1. 矩阵维度信息
                f.write(f"【矩阵维度】\n")
                f.write(f"  n (节点总数) = {n}\n")
                f.write(f"  物料节点数 = {len(self.material_nodes)}\n")
                f.write(f"  工单节点数 = {len(self.order_nodes)}\n\n")
                
                # 2. F 矩阵汇总
                f.write(f"【F矩阵 - 外部投入】\n")
                f.write(f"  F_材料总和 = {F[:, 0].sum():.4f}\n")
                f.write(f"  F_人工总和 = {F[:, 1].sum():.4f}\n")
                f.write(f"  F_制费总和 = {F[:, 2].sum():.4f}\n")
                f.write(f"  F_总计 = {F.sum():.4f}\n\n")
                
                # 3. X 矩阵汇总 (平行结转法结果)
                f.write(f"【X矩阵 - 平行结转法结果】\n")
                f.write(f"  X_材料总和 = {X[:, 0].sum():.4f}\n")
                f.write(f"  X_人工总和 = {X[:, 1].sum():.4f}\n")
                f.write(f"  X_制费总和 = {X[:, 2].sum():.4f}\n")
                f.write(f"  X_总计 = {X.sum():.4f}\n\n")
                
                # 4. S 矩阵汇总 (逐步结转法结果)
                f.write(f"【S矩阵 - 逐步结转法结果】\n")
                f.write(f"  S_材料总和 = {S[:, 0].sum():.4f}\n")
                f.write(f"  S_人工总和 = {S[:, 1].sum():.4f}\n")
                f.write(f"  S_制费总和 = {S[:, 2].sum():.4f}\n")
                f.write(f"  S_总计 = {S.sum():.4f}\n\n")
                
                # 5. 关键节点对比 (前10个工单和前10个物料)
                f.write("=" * 80 + "\n")
                f.write("【关键节点对比 - 前10个工单】\n")
                f.write("-" * 80 + "\n")
                f.write(f"{'节点':<15} {'类型':<8} {'F_工':<12} {'F_费':<12} {'X_工':<12} {'X_费':<12} {'S_工':<12} {'S_费':<12}\n")
                f.write("-" * 80 + "\n")
                
                order_count = 0
                for node in self.all_nodes:
                    if node in self.order_nodes:
                        idx = self.node_index[node]
                        f.write(f"{node:<15} {'工单':<8} {F[idx, 1]:<12.4f} {F[idx, 2]:<12.4f} {X[idx, 1]:<12.4f} {X[idx, 2]:<12.4f} {S[idx, 1]:<12.4f} {S[idx, 2]:<12.4f}\n")
                        order_count += 1
                        if order_count >= 10:
                            break
                
                f.write("\n")
                f.write("【关键节点对比 - 前10个物料】\n")
                f.write("-" * 80 + "\n")
                f.write(f"{'节点':<15} {'类型':<8} {'X_总':<12} {'X_料':<12} {'X_工':<12} {'X_费':<12} {'S_料':<12} {'S_工':<12} {'S_费':<12}\n")
                f.write("-" * 80 + "\n")
                
                mat_count = 0
                for node in self.all_nodes:
                    if node in self.material_nodes:
                        idx = self.node_index[node]
                        x_total = X[idx].sum()
                        s_total = S[idx].sum()
                        f.write(f"{node:<15} {'物料':<8} {x_total:<12.4f} {X[idx, 0]:<12.4f} {X[idx, 1]:<12.4f} {X[idx, 2]:<12.4f} {S[idx, 0]:<12.4f} {S[idx, 1]:<12.4f} {S[idx, 2]:<12.4f}\n")
                        mat_count += 1
                        if mat_count >= 10:
                            break
                
                # 6. I + W×D 矩阵的统计
                f.write("\n")
                f.write("=" * 80 + "\n")
                f.write("【I + W×D 矩阵统计】\n")
                row_sums = np.array(I_plus_WD.sum(axis=1)).flatten()
                f.write(f"  I_plus_WD 行和均值 = {row_sums.mean():.4f}\n")
                f.write(f"  I_plus_WD 行和最大 = {row_sums.max():.4f}\n")
                f.write(f"  I_plus_WD 行和最小 = {row_sums.min():.4f}\n")
                f.write(f"\n  对角线元素均值 (应为1.0) = {I_plus_WD.diagonal().mean():.4f}\n")
                
                # 7. 找几个工单的上下游关系
                f.write("\n")
                f.write("=" * 80 + "\n")
                f.write("【典型工单上下游关系】\n")
                sample_orders = self.order_nodes[:5] if len(self.order_nodes) >= 5 else self.order_nodes
                for order in sample_orders:
                    o_idx = self.node_index[order]
                    f.write(f"\n工单: {order}\n")
                    f.write(f"  本节点: F_工={F[o_idx, 1]:.4f}, F_费={F[o_idx, 2]:.4f}\n")
                    f.write(f"  本节点: X_工={X[o_idx, 1]:.4f}, X_费={X[o_idx, 2]:.4f}\n")
                    f.write(f"  本节点: S_工={S[o_idx, 1]:.4f}, S_费={S[o_idx, 2]:.4f}\n")
                    
                    # 找出流入这个工单的材料
                    incoming_mats = []
                    for mat in self.material_nodes:
                        m_idx = self.node_index[mat]
                        if W[o_idx, m_idx] > 0.001:  # 材料→工单
                            incoming_mats.append((mat, W[o_idx, m_idx], X[m_idx, 1], X[m_idx, 2], S[m_idx, 1], S[m_idx, 2]))
                    
                    if incoming_mats:
                        f.write(f"  上游材料 ({len(incoming_mats)}个):\n")
                        for mat, w, x_lab, x_oh, s_lab, s_oh in incoming_mats[:5]:
                            f.write(f"    {mat}: W={w:.4f}, X_工={x_lab:.4f}, X_费={x_oh:.4f}, S_工={s_lab:.4f}, S_费={s_oh:.4f}\n")
                    
                    # 找出这个工单产出的产品
                    outgoing_prods = []
                    for prod in self.material_nodes:
                        p_idx = self.node_index[prod]
                        if W[p_idx, o_idx] > 0.001:  # 工单→产品
                            outgoing_prods.append((prod, W[p_idx, o_idx], X[p_idx, 1], X[p_idx, 2], S[p_idx, 1], S[p_idx, 2]))
                    
                    if outgoing_prods:
                        f.write(f"  下游产品 ({len(outgoing_prods)}个):\n")
                        for prod, w, x_lab, x_oh, s_lab, s_oh in outgoing_prods[:5]:
                            f.write(f"    {prod}: W={w:.4f}, X_工={x_lab:.4f}, X_费={x_oh:.4f}, S_工={s_lab:.4f}, S_费={s_oh:.4f}\n")
            
            # 逐步结转法下的 WIP = W × (I-D) × S_材料
            wip_step_at_order = S_mat * I_minus_D  # (1-D) × S_材料
            WIP_step_col = W.dot(wip_step_at_order)
            WIP_step = np.zeros((n, 3))
            WIP_step[:, 0] = WIP_step_col
            WIP_step = np.round(WIP_step, 4)
            
            # 构建逐步结转法的工单明细
            step_order_detail_list = []
            for _, row in order_prod_summary.iterrows():
                order = str(row['工单号'])
                product = str(row['产品编码'])
                
                if order not in self.node_index or product not in self.node_index:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self.node_index[product]
                
                w_ratio = W[prod_idx, order_idx]
                finished_ratio = D_mat[order_idx]
                
                finished_q = row['产品完工数量']
                wip_q = row['在产品数量']
                
                # 逐步结转法下的工单成本
                order_S_mat = S[order_idx, 0]  # 材料
                order_S_lab = S[order_idx, 1]  # 人工
                order_S_oh = S[order_idx, 2]   # 制费
                
                # 分配成本（与X矩阵相同逻辑）
                finished_S_mat = w_ratio * order_S_mat * finished_ratio
                wip_S_mat = w_ratio * order_S_mat * (1 - finished_ratio)
                finished_S_lab = w_ratio * order_S_lab
                finished_S_oh = w_ratio * order_S_oh
                
                step_order_detail_list.append({
                    '工单号': order,
                    '产品编码': product,
                    'W矩阵贡献比例': round(w_ratio, 4),
                    '完工率(D矩阵)': round(finished_ratio, 4),
                    '完工数量': round(finished_q, 2),
                    '在产品数量': round(wip_q, 2),
                    '在产品材料费': round(wip_S_mat, 2),
                    '完工产品材料费': round(finished_S_mat, 2),
                    '完工产品直接人工': round(finished_S_lab, 2),
                    '完工产品制造费用': round(finished_S_oh, 2),
                    '工单总成本': round(wip_S_mat + finished_S_mat + finished_S_lab + finished_S_oh, 2),
                })
            
            # 构建逐步结转法的工单-产品-材料三维明细（与平行结转法逻辑相同，只是用S矩阵）
            step_order_prod_mat_detail_list = []
            
            # 第一步：计算每个工单-产品的总领用数量（用于材料成本分配）
            step_order_prod_total_issue = {}
            for _, row in self.io_df.iterrows():
                order = str(row['工单号'])
                product = str(row['产品编码'])
                issue_qty = row['材料领用数量']
                
                key = (order, product)
                if key not in step_order_prod_total_issue:
                    step_order_prod_total_issue[key] = 0
                step_order_prod_total_issue[key] += issue_qty
            
            # 第二步：先计算工单-产品-汇总表的数据（逐步结转法，与汇总表一致）
            step_order_prod_summary_data = {}
            for _, row in order_prod_summary.iterrows():
                order = str(row['工单号'])
                product = str(row['产品编码'])
                
                if order not in self.node_index or product not in self.node_index:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self.node_index[product]
                w_ratio = W[prod_idx, order_idx]
                finished_ratio = D_mat[order_idx]
                
                # 逐步结转法使用S矩阵
                order_S_mat = S[order_idx, 0]
                order_S_lab = S[order_idx, 1]
                order_S_oh = S[order_idx, 2]
                
                # 与汇总表一致的计算
                finished_mat = w_ratio * order_S_mat * finished_ratio
                wip_mat = w_ratio * order_S_mat * (1 - finished_ratio)
                finished_lab = w_ratio * order_S_lab
                finished_oh = w_ratio * order_S_oh
                
                key = (order, product)
                step_order_prod_summary_data[key] = {
                    'finished_mat': finished_mat,
                    'wip_mat': wip_mat,
                    'finished_lab': finished_lab,
                    'finished_oh': finished_oh
                }
            
            # 第三步：遍历生成明细表，按领用数量权重拆分
            for _, row in self.io_df.iterrows():
                order = str(row['工单号'])
                product = str(row['产品编码'])
                material = str(row['材料编码'])
                issue_qty = row['材料领用数量']
                
                if order not in self.node_index or product not in self.node_index or material not in self.node_index:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self.node_index[product]
                mat_idx = self.node_index[material]
                
                # W矩阵比例（用于展示）
                w_mat_to_order = W[order_idx, mat_idx]
                w_order_to_prod = W[prod_idx, order_idx]
                finished_ratio = D_mat[order_idx]
                
                # 获取该工单-产品的汇总数据
                key = (order, product)
                summary = step_order_prod_summary_data.get(key, {
                    'finished_mat': 0, 'wip_mat': 0, 'finished_lab': 0, 'finished_oh': 0
                })
                
                # 获取该工单-产品的总领用数量
                total_issue = step_order_prod_total_issue.get(key, 1)
                
                # 按领用数量权重拆分成本
                if total_issue > 0:
                    weight = issue_qty / total_issue
                    finished_mat_cost = summary['finished_mat'] * weight
                    wip_mat_cost = summary['wip_mat'] * weight
                    finished_lab_cost = summary['finished_lab'] * weight
                    finished_oh_cost = summary['finished_oh'] * weight
                else:
                    # 如果没有领用数量（只有工费），则材料成本为0，但工费需要分配
                    # 按行数平均分配工费
                    finished_mat_cost = 0
                    wip_mat_cost = 0
                    # 获取该工单-产品的行数
                    row_count = len([r for r in self.io_df.iterrows() 
                                     if str(r[1]['工单号']) == order and str(r[1]['产品编码']) == product])
                    if row_count > 0:
                        finished_lab_cost = summary['finished_lab'] / row_count
                        finished_oh_cost = summary['finished_oh'] / row_count
                    else:
                        finished_lab_cost = summary['finished_lab']
                        finished_oh_cost = summary['finished_oh']
                
                # 获取该行的完工/在产数量（直接取原始数据）
                finished_q = row['产品完工数量']
                wip_q = row['在产品数量']
                
                step_order_prod_mat_detail_list.append({
                    '工单号': order,
                    '产品编码': product,
                    '材料编码': material,
                    '材料领用数量': round(issue_qty, 2),
                    'W_材料到工单': round(w_mat_to_order, 4),
                    'W_工单到产品': round(w_order_to_prod, 4),
                    '完工率': round(finished_ratio, 4),
                    '完工数量': round(finished_q, 2),
                    '在产品数量': round(wip_q, 2),
                    '在产品材料费': round(wip_mat_cost, 2),
                    '完工产品材料费': round(finished_mat_cost, 2),
                    '完工产品直接人工': round(finished_lab_cost, 2),
                    '完工产品制造费用': round(finished_oh_cost, 2),
                    '成本小计': round(wip_mat_cost + finished_mat_cost + finished_lab_cost + finished_oh_cost, 2),
                })
            
            # 构建逐步结转法的成本明细
            step_detail_list = []
            for i, node in enumerate(self.all_nodes):
                is_order = node in self.order_nodes
                step_detail_list.append({
                    '节点': node,
                    '类型': '工单' if is_order else '物料',
                    '总成本': round(S[i].sum(), 4),
                    '料': round(S[i, 0], 4),
                    '工': round(S[i, 1], 4),
                    '费': round(S[i, 2], 4),
                })
            
            step_result = {
                '逐步结转_工单明细': pd.DataFrame(step_order_detail_list),
                '逐步结转_工单产品材料明细': pd.DataFrame(step_order_prod_mat_detail_list),
                '逐步结转_成本明细': pd.DataFrame(step_detail_list),
            }
            
            self.performance_log['逐步结转法计算'] = time.time() - t0
        
        self.performance_log['总计算时间'] = time.time() - total_start
        
        self.result = {
            '收发存': pd.DataFrame(result_list),
            '工单明细': pd.DataFrame(order_detail_list),
            '工单产品材料明细': pd.DataFrame(order_prod_mat_detail_list),  # 新增：三维明细
            '成本明细': pd.DataFrame(detail_list),
            'nodes': self.all_nodes
        }
        
        # 合并逐步结转法结果（如果有）
        if step_result:
            self.result.update(step_result)
        
        # 合并超级成本还原结果（如果有）
        if super_result:
            self.result.update(super_result)
        
        # 合并销售成本结果（如果有）
        if not sales_cost_df.empty:
            self.result['销售成本明细'] = sales_cost_df

        with open(r'D:\桌面\公众号\成本核算系统\log\debug_cost.txt', 'w', encoding='utf-8') as f:
            f.write("=== 成本调试报告 ===\n")
            f.write(f"F矩阵总和: {np.sum(F):.2f}\n")
            f.write(f"X矩阵总和: {np.sum(X):.2f}\n")
            f.write(f"放大倍数: {np.sum(X)/np.sum(F):.2f}\n")
            f.write("\n前10个节点成本:\n")
            for i in range(min(10, len(self.all_nodes))):
                f.write(f"{self.all_nodes[i]}: F={F[i]}, X={X[i]}\n")

        with open(r'D:\桌面\公众号\成本核算系统\log\debug_w_matrix.txt', 'w', encoding='utf-8') as f:
            f.write("=== W 矩阵列和检查 ===\n")
            
            col_sums = np.array(W.sum(axis=0)).flatten()
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
                    col = W[:, i].toarray().flatten()
                    non_zero = np.where(col > 0.001)[0]
                    f.write(f"  被领用情况:\n")
                    for j in non_zero:
                        f.write(f"    {self.all_nodes[j]}: {col[j]:.4f}\n")
                    f.write("\n")
            
            # 检查是否有自环（对角线元素）
            f.write("\n=== 对角线检查 ===\n")
            diag = W.diagonal()
            if np.any(diag > 0.001):
                f.write("警告：存在自环！\n")
                for i in range(len(self.all_nodes)):
                    if diag[i] > 0.001:
                        f.write(f"{self.all_nodes[i]}: {diag[i]:.4f}\n")

        return self.result
    
    def calculate_sales_cost(self, sales_df, sales_map):
        """计算销售成本
        
        公式: C = B × S × X
        
        Parameters:
        -----------
        sales_df : DataFrame
            销售数据，包含物料编码、销售数量、销售批次号
        sales_map : dict
            字段映射 {原列名: 标准名}
        
        Returns:
        --------
        pd.DataFrame
            销售成本明细表（每行一个批次+物料组合）
        """
        if self.X_total is None:
            raise ValueError("请先执行成本核算以获取X矩阵")
        
        if self.available_qty is None:
            raise ValueError("可供发出数量未计算，请重新执行成本核算")
        
        # 读取并清洗销售数据
        df = sales_df.rename(columns=sales_map)
        df['物料编码'] = df['物料编码'].astype(str).str.strip()
        df['销售数量'] = pd.to_numeric(df['销售数量'], errors='coerce').fillna(0)
        df['销售批次号'] = df['销售批次号'].astype(str).str.strip()
        
        # 按物料汇总销售数量
        mat_sales = df.groupby('物料编码')['销售数量'].sum().to_dict()
        
        # 按(物料, 批次)汇总销售数量
        batch_sales = df.groupby(['物料编码', '销售批次号'])['销售数量'].sum().reset_index()
        
        n = len(self.all_nodes)
        
        # ========== 构建 Sale 矩阵（N×N 对角阵）==========
        S = np.zeros((n, n))
        for mat, sales_qty in mat_sales.items():
            if mat in self.node_index:
                idx = self.node_index[mat]
                avail = self.available_qty.get(mat, 0)
                if avail > 0:
                    S[idx, idx] = sales_qty / avail
                else:
                    S[idx, idx] = 0
        
        # ========== 构建 Batch 矩阵（M×N）==========
        batch_list = sorted(batch_sales['销售批次号'].unique())
        m = len(batch_list)
        batch_index = {b: i for i, b in enumerate(batch_list)}
        
        B = np.zeros((m, n))
        for _, row in batch_sales.iterrows():
            mat = str(row['物料编码'])
            batch = str(row['销售批次号'])
            qty = row['销售数量']
            
            if mat in self.node_index and batch in batch_index:
                mat_idx = self.node_index[mat]
                batch_idx = batch_index[batch]
                total_sales = mat_sales.get(mat, 0)
                if total_sales > 0:
                    B[batch_idx, mat_idx] = qty / total_sales
        
        # ========== 计算 C = B × S × X ==========
        # S × X: N×3
        SX = S.dot(self.X_total)
        
        # B × SX: M×3
        C = B.dot(SX)
        
        # ========== 构建结果 DataFrame ==========
        result_list = []
        for _, row in batch_sales.iterrows():
            mat = str(row['物料编码'])
            batch = str(row['销售批次号'])
            qty = row['销售数量']
            
            if mat not in self.node_index:
                continue
            
            mat_idx = self.node_index[mat]
            batch_idx = batch_index.get(batch)
            if batch_idx is None:
                continue
            
            avail = self.available_qty.get(mat, 0)
            total_sales = mat_sales.get(mat, 0)
            sales_ratio = S[mat_idx, mat_idx]  # 销售占比 = 总销售 / 可供发出
            batch_ratio = B[batch_idx, mat_idx]  # 批次占比 = 批次销售 / 总销售
            
            # 从 C 矩阵取该批次的成本（如果同一批次有多个物料，需要按物料拆分）
            # 实际上 C[batch_idx] 是该批次所有物料的成本之和
            # 所以按 batch_ratio / 该批次在所有物料上的占比 来拆分
            # 更简单：直接用 batch_ratio * SX[mat_idx]
            cost_mat = batch_ratio * SX[mat_idx, 0]
            cost_lab = batch_ratio * SX[mat_idx, 1]
            cost_oh = batch_ratio * SX[mat_idx, 2]
            cost_total = cost_mat + cost_lab + cost_oh
            
            result_list.append({
                '销售批次号': batch,
                '物料编码': mat,
                '销售数量': round(qty, 2),
                '可供发出数量': round(avail, 2),
                '销售占比': f"{sales_ratio:.2%}",
                '批次占物料销售比': f"{batch_ratio:.2%}",
                '销售成本_料': round(cost_mat, 2),
                '销售成本_工': round(cost_lab, 2),
                '销售成本_费': round(cost_oh, 2),
                '销售成本_合计': round(cost_total, 2),
            })
        
        # 按批次号排序
        result_df = pd.DataFrame(result_list)
        if not result_df.empty:
            result_df = result_df.sort_values(['销售批次号', '物料编码'])
        
        return result_df
    
    def calculate_cost_restoration(self, step_cost_df, step_map):
        """成本还原计算
        
        公式：
        - X_人工 = (I - W × D)^(-1) × F_人工  (F_人工来自原始F矩阵，只有工单有值)
        - X_制费 = (I - W × D)^(-1) × F_制费  (F_制费来自原始F矩阵，只有工单有值)
        - X_材料 = S_总 - X_人工 - X_制费
        """
        if self.W_matrix is None:
            raise ValueError("请先执行成本核算以构建流转矩阵 W")
        
        if self.D_matrix is None:
            raise ValueError("D矩阵未找到，请重新执行成本核算")
        
        if self.F_matrix is None:
            raise ValueError("F矩阵未找到，请重新执行成本核算")
        
        n = len(self.all_nodes)
        W = self.W_matrix  # 现在是稀疏csr
        D_mat = self.D_matrix
        F_orig = self.F_matrix  # 原始F矩阵（只有工单有人工制费）
        
        # Step 1: 读取逐步结转法的 S_材料、S_人工、S_制费
        # 注意：step_map 是 {原列名: 标准名}，需要反转使用
        reverse_map = {v: k for k, v in step_map.items()}
        
        S_mat = np.zeros(n)   # 逐步结转的材料（综合材料成本）
        S_lab = np.zeros(n)   # 逐步结转的人工（已包含上游转入）
        S_oh = np.zeros(n)    # 逐步结转的制费（已包含上游转入）
        
        col_mat = reverse_map.get('料')
        col_lab = reverse_map.get('工')
        col_oh = reverse_map.get('费')
        col_code = reverse_map.get('物料编码')
        
        if not all([col_mat, col_lab, col_oh, col_code]):
            raise ValueError(f"字段映射不完整，需要：料、工、费、物料编码，当前映射: {step_map}")
        
        # 调试：检查输入数据的前几行
        debug_rows = []
        match_count = 0
        mismatch_rows = []
        
        for _, row in step_cost_df.iterrows():
            mat = str(row[col_code]).strip()
            if mat in self.node_index:
                idx = self.node_index[mat]
                val_mat = pd.to_numeric(row[col_mat], errors='coerce') or 0
                val_lab = pd.to_numeric(row[col_lab], errors='coerce') or 0
                val_oh = pd.to_numeric(row[col_oh], errors='coerce') or 0
                
                S_mat[idx] = val_mat
                S_lab[idx] = val_lab
                S_oh[idx] = val_oh
                match_count += 1
                
                if len(debug_rows) < 10:
                    debug_rows.append({
                        '物料': mat, '料': val_mat, '工': val_lab, '费': val_oh,
                        'idx': idx, '类型': '工单' if mat in self.order_nodes else '物料'
                    })
        
        # S_total = S_mat + S_lab + S_oh（来自逐步结转报表）
        S_total = S_mat + S_lab + S_oh
        
        # Step 2: 使用原始 F 矩阵的人工制费（只有工单有值）
        F_lab = F_orig[:, 1].copy()  # 原始人工（只有工单有值）
        F_oh = F_orig[:, 2].copy()   # 原始制费（只有工单有值）
        
        # Step 3: 构建 A 矩阵
        # 注意：成本还原的逻辑应该与成本核算一致
        # - 材料用 D_mat（材料按完工率分配）
        # - 工费用 D_loh = I（工费全额给完工产品，不给在产品）
        
        # 稀疏构建 A_mat 和 A_loh
        W_mat_data = W.data.copy() * D_mat[W.indices]
        W_mat = sparse.csr_matrix((W_mat_data, W.indices.copy(), W.indptr.copy()), shape=(n, n))
        A_mat = sparse.eye(n) - W_mat
        A_loh = sparse.eye(n) - W
        
        # Step 4: 求解 X_人工 和 X_制费（使用原始F矩阵的人工制费）
        # 注意：人工和制费应该用 A_loh（工费路径），不是 A_mat！
        try:
            X_true_lab = spsolve(A_loh, F_lab)
            X_true_oh = spsolve(A_loh, F_oh)
        except Exception:
            X_true_lab = np.linalg.solve(A_loh.toarray(), F_lab)
            X_true_oh = np.linalg.solve(A_loh.toarray(), F_oh)
        
        # Step 4: X_材料 = S_总 - X_人工 - X_制费
        X_true_mat = S_total - X_true_lab - X_true_oh
        
        # Step 5: 组装结果
        X_restored = np.column_stack([X_true_mat, X_true_lab, X_true_oh])
        X_restored = np.round(X_restored, 4)

        # Step 6: 验证勾稽（还原前后总成本必须相等）
        restored_total = np.sum(X_restored, axis=1)
        original_total = S_total
        diff = np.abs(restored_total - original_total)
        max_diff = np.max(diff) if len(diff) > 0 else 0       
        
        if max_diff > 1:  # 允许1元误差
            print(f"警告：成本还原不勾稽！最大差异: {max_diff:.2f}")
    
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
            'F_step': np.column_stack([S_mat, F_lab, F_oh]),  # 记录原始输入（S_mat是逐步结转材料，F_lab/F_oh是原始人工制费）
            'max_diff': max_diff
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
