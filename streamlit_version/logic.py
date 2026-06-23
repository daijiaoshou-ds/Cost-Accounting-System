import pandas as pd
import numpy as np
from io import BytesIO
import time
import warnings
warnings.filterwarnings('ignore')
import polars as pl

# ========== 稀疏矩阵支持 ==========
from scipy import sparse
from scipy.sparse.linalg import spsolve


def _read_excel(file):
    """统一 Excel 读取入口，优先 Polars+fastexcel，失败回退 pandas"""
    try:
        # 先用 polars 读
        df_pl = pl.read_excel(file)
        return df_pl.to_pandas()
    except Exception:
        try:
            file.seek(0)
            return pd.read_excel(file)
        except Exception:
            file.seek(0)
            return pd.read_excel(file, engine='openpyxl')


TABLE_SCHEMA = {
    'initial': {
        'required': ['年度', '月份', '存货编码', '数量', '直接材料', '直接人工', '制造费用'],
        'rename': {
            '存货编码': '物料编码', '数量': '期初数量',
            '直接材料': '期初材料', '直接人工': '期初人工', '制造费用': '期初制费',
            '库存类型': '库存类型',
        },
        'groupby': ['年度', '月份', '物料编码', '库存类型'],
        'agg': {'期初数量': 'sum', '期初材料': 'sum', '期初人工': 'sum', '期初制费': 'sum'},
        'optional': {'库存类型': '仓库'},
    },
    'purchase': {
        'required': ['年度', '月份', '存货编码', '采购数量', '采购金额'],
        'rename': {'存货编码': '物料编码'},
        'groupby': ['年度', '月份', '物料编码'],
        'agg': {'采购数量': 'sum', '采购金额': 'sum'},
    },
    'labor': {
        'required': ['年度', '月份', '工单号', '直接人工', '制造费用'],
        'rename': {'直接人工': '人工', '制造费用': '制费'},
        'groupby': ['年度', '月份', '工单号'],
        'agg': {'人工': 'sum', '制费': 'sum'},
    },
    'io': {
        'required': ['年度', '月份', '工单号', '产品编码', '材料编码', '领用数量', '完工数量', '在产数量'],
        'rename': {'领用数量': '材料领用数量', '完工数量': '产品完工数量', '在产数量': '在产品数量'},
        'groupby': ['年度', '月份', '工单号', '产品编码', '材料编码'],
        'agg': {'材料领用数量': 'sum', '产品完工数量': 'sum', '在产品数量': 'sum'},
    },
    'finished': {
        'required': ['年度', '月份', '存货编码', '入库数量'],
        'rename': {'存货编码': '产品编码'},
        'groupby': ['年度', '月份', '产品编码'],
        'agg': {'入库数量': 'sum'},
    },
    'sales': {
        'required': ['年度', '月份', '存货编码', '出库单号', '销售数量'],
        'rename': {'存货编码': '物料编码', '出库单号': '销售批次号'},
        'groupby': ['年度', '月份', '物料编码', '销售批次号'],
        'agg': {'销售数量': 'sum', '销售金额': 'sum'},
        'optional': {'销售金额': 0},
    },
}


def load_and_aggregate(file_dict, mapping_dict):
    """
    读取全部上传文件，按必要字段聚合后按月份分组。
    
    Parameters:
    -----------
    file_dict : dict
        {表名: file_like_object}，表名包括 'initial', 'purchase', 'labor', 'io', 'finished', 'sales'
    mapping_dict : dict
        {表名: {原始列名: 标准名}}
    
    Returns:
    --------
    dict
        {(year, month): {'initial': df, 'purchase': df, 'labor': df, 'io': df, 'finished': df, 'sales': df}}
    """
    result = {}
    for table, schema in TABLE_SCHEMA.items():
        required = schema['required']
        rename = schema['rename']
        groupby = schema['groupby']
        agg = schema['agg']
        
        optional = schema.get('optional', {})

        file = file_dict.get(table)
        if file is None:
            internal_cols = [rename.get(c, c) for c in required if c not in ('年度', '月份')]
            internal_cols = list(dict.fromkeys(internal_cols))
            # 包含 optional 字段
            for opt_field in optional.keys():
                opt_internal = rename.get(opt_field, opt_field)
                if opt_internal not in internal_cols:
                    internal_cols.append(opt_internal)
            df = pd.DataFrame(columns=internal_cols)
        else:
            df = _read_excel(file)
            col_map = mapping_dict.get(table, {})
            df = df.rename(columns=col_map)

            missing = [c for c in required if c not in df.columns]
            if missing:
                raise ValueError(f"表 {table} 缺少必要字段: {missing}")

            df = df.rename(columns=rename)
            # 构建 internal_cols：required + optional 中实际存在的列
            internal_cols = [rename.get(c, c) for c in required]
            internal_cols = list(dict.fromkeys(internal_cols))
            for opt_field in optional.keys():
                opt_internal = rename.get(opt_field, opt_field)
                if opt_internal in df.columns and opt_internal not in internal_cols:
                    internal_cols.append(opt_internal)
            df = df[internal_cols].copy()
            # 缺失的 optional 字段用默认值填充
            for opt_field, default_val in optional.items():
                opt_internal = rename.get(opt_field, opt_field)
                if opt_internal not in df.columns:
                    df[opt_internal] = default_val
        
        # 字符串字段去空格
        str_cols = [c for c in df.columns if '编码' in c or '类型' in c or c in ('工单号', '销售批次号')]
        for c in str_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()
        
        # 年度、月份转整数
        for c in ('年度', '月份'):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
        
        # 数值字段清洗
        numeric_cols = [c for c in df.columns if c not in str_cols and c not in ('年度', '月份')]
        for c in numeric_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        if not df.empty:
            df = df.groupby(groupby, as_index=False).agg(agg)
        
        if '年度' in df.columns and '月份' in df.columns:
            for (year, month), g in df.groupby(['年度', '月份']):
                key = (int(year), int(month))
                result.setdefault(key, {})[table] = (
                    g.drop(columns=['年度', '月份']).reset_index(drop=True)
                )
        else:
            result.setdefault((0, 0), {})[table] = df.reset_index(drop=True)
    
    # 确保每个月份都包含所有表（含 optional 字段）
    all_tables = list(TABLE_SCHEMA.keys())
    all_keys = list(result.keys())

    def _build_internal_cols(schema):
        internal_cols = [schema['rename'].get(c, c) for c in schema['required'] if c not in ('年度', '月份')]
        internal_cols = list(dict.fromkeys(internal_cols))
        for opt_field in schema.get('optional', {}).keys():
            opt_internal = schema['rename'].get(opt_field, opt_field)
            if opt_internal not in internal_cols:
                internal_cols.append(opt_internal)
        return internal_cols

    for key in all_keys:
        for table in all_tables:
            if table not in result[key]:
                schema = TABLE_SCHEMA[table]
                result[key][table] = pd.DataFrame(columns=_build_internal_cols(schema))

    # 如果没有找到任何月份数据（所有表都无年度/月份字段），返回默认键
    if not result:
        result[(0, 0)] = {}
        for table in all_tables:
            schema = TABLE_SCHEMA[table]
            result[(0, 0)][table] = pd.DataFrame(columns=_build_internal_cols(schema))

    return result


class CostCalculator:
    def __init__(self):
        self.initial_df = None
        self.purchase_df = None
        self.io_df = None
        self.labor_df = None
        self.sales_df = None  # 销售数据
        self.finished_df = None  # 完工入库数据
        self.result = None
        self.performance_log = {}
        
        # 保存中间计算结果
        self.W_matrix = None           # (兼容旧接口) 仓库W矩阵
        self.W_workshop = None         # 车间W矩阵
        self.W_warehouse = None        # 仓库W矩阵
        self.F_matrix = None           # (兼容旧接口) 原始外部投入矩阵
        self.F_workshop = None         # 车间F矩阵
        self.F_warehouse = None        # 仓库F矩阵
        self.D_matrix = None           # 阀门矩阵
        self.all_nodes = None
        self.material_nodes = None
        self.order_nodes = None
        self.node_index = None
        self.X_total = None            # 总成本矩阵
        self.available_qty = None      # 可供发出数量（W矩阵用，材料完工口径）
        self.sales_available_qty = None  # 可供发出数量（销售用，产品完工口径）

        # 双矩阵调试状态
        self._self_loop_skipped = []       # 自环跳过记录
        self._force_normalized = False     # 是否触发强制计算归一化
        self._force_normalization_log = [] # 归一化详情 [{node, col_sum, correction_factor}]
        self._workshop_consumed = {}       # {material: total_from_workshop}
        self._warehouse_consumed = {}      # {material: total_from_warehouse}
        
    def load_data(self, data_dict):
        """
        加载单月份已聚合的数据。
        
        data_dict: {
            'initial': df,
            'purchase': df,
            'io': df,
            'labor': df,
            'finished': df,
            'sales': df
        }
        """
        start_time = time.time()
        
        # 期初（可选），按 物料编码+库存类型 分组保留仓库/车间区分
        df_init = data_dict.get('initial', pd.DataFrame())
        if df_init.empty:
            self.initial_df = pd.DataFrame(columns=['物料编码', '库存类型', '期初数量', '期初材料', '期初人工', '期初制费'])
        else:
            df_init = df_init.copy()
            df_init['物料编码'] = df_init['物料编码'].astype(str).str.strip()
            # 确保 库存类型 列存在
            if '库存类型' not in df_init.columns:
                df_init['库存类型'] = '仓库'
            else:
                df_init['库存类型'] = df_init['库存类型'].astype(str).str.strip()
            for col in ['期初数量', '期初材料', '期初人工', '期初制费']:
                if col in df_init.columns:
                    df_init[col] = pd.to_numeric(df_init[col], errors='coerce').fillna(0)
                else:
                    df_init[col] = 0
            self.initial_df = df_init.groupby(['物料编码', '库存类型']).agg({
                '期初数量': 'sum',
                '期初材料': 'sum',
                '期初人工': 'sum',
                '期初制费': 'sum'
            }).reset_index()
        
        # 采购（必须）
        df_pur = data_dict.get('purchase', pd.DataFrame())
        if df_pur.empty:
            self.purchase_df = pd.DataFrame(columns=['物料编码', '采购数量', '采购金额'])
        else:
            df_pur = df_pur.copy()
            df_pur['物料编码'] = df_pur['物料编码'].astype(str).str.strip()
            df_pur['采购数量'] = pd.to_numeric(df_pur['采购数量'], errors='coerce').fillna(0)
            df_pur['采购金额'] = pd.to_numeric(df_pur['采购金额'], errors='coerce').fillna(0)
            self.purchase_df = df_pur.groupby('物料编码').agg({
                '采购数量': 'sum',
                '采购金额': 'sum'
            }).reset_index()
        
        # 投入产出（必须）
        df_io = data_dict.get('io', pd.DataFrame())
        if df_io.empty:
            self.io_df = pd.DataFrame(columns=['工单号', '产品编码', '材料编码', '材料领用数量', '产品完工数量', '在产品数量'])
        else:
            df_io = df_io.copy()
            df_io['工单号'] = df_io['工单号'].astype(str).str.strip()
            df_io['产品编码'] = df_io['产品编码'].astype(str).str.strip()
            df_io['材料编码'] = df_io['材料编码'].astype(str).str.strip()
            df_io['产品完工数量'] = pd.to_numeric(df_io['产品完工数量'], errors='coerce').fillna(0)
            df_io['材料领用数量'] = pd.to_numeric(df_io['材料领用数量'], errors='coerce').fillna(0)
            if '在产品数量' in df_io.columns:
                df_io['在产品数量'] = pd.to_numeric(df_io['在产品数量'], errors='coerce').fillna(0)
            else:
                df_io['在产品数量'] = 0
            self.io_df = df_io.groupby(['工单号', '产品编码', '材料编码']).agg({
                '材料领用数量': 'sum',
                '产品完工数量': 'sum',
                '在产品数量': 'sum'
            }).reset_index()
        
        # 工单费用（可选）
        df_lab = data_dict.get('labor', pd.DataFrame())
        if df_lab.empty:
            self.labor_df = pd.DataFrame(columns=['工单号', '人工', '制费'])
        else:
            df_lab = df_lab.copy()
            df_lab['工单号'] = df_lab['工单号'].astype(str).str.strip()
            df_lab['人工'] = pd.to_numeric(df_lab['人工'], errors='coerce').fillna(0)
            df_lab['制费'] = pd.to_numeric(df_lab['制费'], errors='coerce').fillna(0)
            self.labor_df = df_lab.groupby('工单号')[['人工', '制费']].sum().reset_index()
        
        # 销售数据（可选）
        df_sales = data_dict.get('sales', pd.DataFrame())
        if df_sales.empty:
            self.sales_df = pd.DataFrame(columns=['物料编码', '销售批次号', '销售数量', '销售金额'])
        else:
            df_sales = df_sales.copy()
            df_sales['物料编码'] = df_sales['物料编码'].astype(str).str.strip()
            df_sales['销售数量'] = pd.to_numeric(df_sales['销售数量'], errors='coerce').fillna(0)
            df_sales['销售批次号'] = df_sales['销售批次号'].astype(str).str.strip()
            if '销售金额' in df_sales.columns:
                df_sales['销售金额'] = pd.to_numeric(df_sales['销售金额'], errors='coerce').fillna(0)
            else:
                df_sales['销售金额'] = 0
            self.sales_df = df_sales.groupby(['物料编码', '销售批次号']).agg(
                销售数量=('销售数量', 'sum'),
                销售金额=('销售金额', 'sum'),
            ).reset_index()
        
        # 完工入库（可选）
        df_fin = data_dict.get('finished', pd.DataFrame())
        if df_fin.empty:
            self.finished_df = pd.DataFrame(columns=['产品编码', '入库数量'])
        else:
            df_fin = df_fin.copy()
            df_fin['产品编码'] = df_fin['产品编码'].astype(str).str.strip()
            df_fin['入库数量'] = pd.to_numeric(df_fin['入库数量'], errors='coerce').fillna(0)
            self.finished_df = df_fin.groupby('产品编码')['入库数量'].sum().reset_index()
        
        self.performance_log['数据清洗'] = time.time() - start_time
        return True
    
    def calculate(self, finished_df=None, finished_map=None,
                  calculate_step_method=False, force_calculate=False):
        """执行核心成本计算（独立节点单系统）

        Parameters
        ----------
        finished_df : DataFrame, optional
            产成品入库明细表
        finished_map : dict, optional
            字段映射
        calculate_step_method : bool, optional
            是否计算逐步结转法
        force_calculate : bool, optional
            是否强制计算（超领时自动归一化）
        """
        # 稀疏矩阵模块已在文件顶部导入

        # 若未传入 finished_df，优先使用 load_data 中保存的 finished 数据
        if finished_df is None and self.finished_df is not None:
            finished_df = self.finished_df

        total_start = time.time()

        # 重置调试状态
        self._self_loop_skipped = []
        self._force_normalized = False
        self._force_normalization_log = []
        self._workshop_consumed = {}
        self._warehouse_consumed = {}

        # Step 1: 构建节点（独立节点：物料#车间 + 物料#仓库 + 工单）
        t0 = time.time()
        WS = '#车间'
        WH = '#仓库'

        products = set(self.io_df['产品编码'].unique())
        materials = set(self.io_df['材料编码'].unique())
        init_materials = set(self.initial_df['物料编码'].unique())
        pur_materials = set(self.purchase_df['物料编码'].unique())
        all_materials = sorted(products | materials | init_materials | pur_materials)
        all_orders = sorted(set(self.io_df['工单号'].unique()))

        self.material_nodes = all_materials  # 原始物料编码列表
        self.order_nodes = all_orders

        # 独立节点：每个物料拆成车间+仓库两个节点
        ws_nodes = [m + WS for m in all_materials]
        wh_nodes = [m + WH for m in all_materials]
        self.all_nodes = ws_nodes + wh_nodes + all_orders
        n = len(self.all_nodes)
        self.node_index = {node: i for i, node in enumerate(self.all_nodes)}

        # 辅助映射
        n_mat = len(all_materials)
        self._mat_ws_idx = {m: i for i, m in enumerate(all_materials)}           # → ws node index
        self._mat_wh_idx = {m: i + n_mat for i, m in enumerate(all_materials)}   # → wh node index
        self.performance_log['构建节点'] = time.time() - t0

        # ============================================================
        # Step 2: 计算基础数量（车间/仓库分离）
        # ============================================================
        t0 = time.time()

        # --- 2a. 采购数据 ---
        pur_qty = {}
        pur_amt = {}
        for mat in self.material_nodes:
            pur_qty[mat] = 0
            pur_amt[mat] = 0
        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码'])
            if mat in pur_qty:
                pur_qty[mat] += row['采购数量']
                pur_amt[mat] += row['采购金额']

        # --- 2b. 期初数据：按库存类型分离（含工费）---
        workshop_init_qty = {}
        workshop_init_mat = {}
        workshop_init_lab = {}
        workshop_init_oh = {}
        warehouse_init_qty = {}
        warehouse_init_mat = {}
        warehouse_init_lab = {}
        warehouse_init_oh = {}
        for mat in self.material_nodes:
            workshop_init_qty[mat] = 0;  workshop_init_mat[mat] = 0
            workshop_init_lab[mat] = 0;  workshop_init_oh[mat] = 0
            warehouse_init_qty[mat] = 0; warehouse_init_mat[mat] = 0
            warehouse_init_lab[mat] = 0; warehouse_init_oh[mat] = 0

        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码'])
            inv_type = str(row.get('库存类型', '仓库')).strip()
            if mat in workshop_init_qty:
                if inv_type == '车间':
                    workshop_init_qty[mat] += row['期初数量']
                    workshop_init_mat[mat] += row.get('期初材料', 0)
                    workshop_init_lab[mat] += row.get('期初人工', 0)
                    workshop_init_oh[mat] += row.get('期初制费', 0)
                else:
                    warehouse_init_qty[mat] += row['期初数量']
                    warehouse_init_mat[mat] += row.get('期初材料', 0)
                    warehouse_init_lab[mat] += row.get('期初人工', 0)
                    warehouse_init_oh[mat] += row.get('期初制费', 0)

        # 兼容旧口径（供报表展示）
        init_qty = {}
        init_amt_mat = {}
        for mat in self.material_nodes:
            init_qty[mat] = warehouse_init_qty[mat] + workshop_init_qty[mat]
            init_amt_mat[mat] = warehouse_init_mat[mat] + workshop_init_mat[mat]

        # --- 2c. 生产数据（完工+在产） ---
        finished_qty_for_w = {}
        finished_qty_for_report = {}
        wip_qty = {}

        for mat in self.material_nodes:
            finished_qty_for_w[mat] = 0
            finished_qty_for_report[mat] = 0
            wip_qty[mat] = 0

        # W矩阵的完工/在产数量：一次 groupby 取两个字段
        io_prod_summary = self.io_df.groupby('产品编码').agg({
            '产品完工数量': 'sum',
            '在产品数量': 'sum'
        }).reset_index()
        for _, row in io_prod_summary.iterrows():
            mat = str(row['产品编码'])
            if mat in finished_qty_for_w:
                finished_qty_for_w[mat] = row['产品完工数量']
            if mat in wip_qty:
                wip_qty[mat] = row['在产品数量']

        # 报表的完工数量：优先从入库明细获取，否则用投入产出
        if finished_df is not None and not finished_df.empty:
            df_fin = finished_df.copy()
            if finished_map:
                df_fin = df_fin.rename(columns=finished_map)
            df_fin['产品编码'] = df_fin['产品编码'].astype(str).str.strip()
            df_fin['入库数量'] = pd.to_numeric(df_fin['入库数量'], errors='coerce').fillna(0)
            fin_summary = df_fin.groupby('产品编码')['入库数量'].sum().to_dict()
            for mat in self.material_nodes:
                finished_qty_for_report[mat] = fin_summary.get(mat, 0)
        else:
            for mat in self.material_nodes:
                finished_qty_for_report[mat] = finished_qty_for_w[mat]

        # --- 2d. 可供发出数量（双口径） ---
        # 车间可供发出 = 车间期初数量（只有期初WIP，不含采购/完工）
        workshop_avail = {}
        for mat in self.material_nodes:
            workshop_avail[mat] = workshop_init_qty[mat]

        # 仓库可供发出 = 仓库期初 + 采购 + 完工
        warehouse_avail = {}
        for mat in self.material_nodes:
            warehouse_avail[mat] = warehouse_init_qty[mat] + pur_qty[mat] + finished_qty_for_w[mat]

        # 销售用可供（纯仓库口径：不含车间期初）
        sales_available_qty = {}
        for mat in self.material_nodes:
            sales_available_qty[mat] = warehouse_init_qty[mat] + pur_qty[mat] + finished_qty_for_report[mat]

        self.performance_log['计算数量'] = time.time() - t0
        
        # ============================================================
        # Step 3: 构建 W 矩阵（独立节点单系统）
        #   产出边: order → 物料#仓库（完工品入仓库）
        #   消耗边: 物料#车间 → order（先来后到优先）, 物料#仓库 → order（车间不够才走仓库）
        # ============================================================
        t0 = time.time()

        # --- 3a. 产出边：order → 物料#仓库 ---
        order_prod_finished = self.io_df.groupby(['工单号', '产品编码'])['产品完工数量'].sum().reset_index()
        order_total_finished = order_prod_finished.groupby('工单号')['产品完工数量'].sum()

        out_rows, out_cols, out_data = [], [], []
        for _, row in order_prod_finished.iterrows():
            order = str(row['工单号'])
            prod = str(row['产品编码'])
            finished = row['产品完工数量']
            if order not in self.node_index or prod not in self._mat_wh_idx:
                continue
            total = order_total_finished.get(order, 0)
            if total > 0:
                ratio = finished / total
            else:
                n_products = len(order_prod_finished[order_prod_finished['工单号'] == order])
                ratio = 1.0 / n_products if n_products > 0 else 0
            out_rows.append(self._mat_wh_idx[prod])   # → 仓库节点
            out_cols.append(self.node_index[order])
            out_data.append(ratio)

        # --- 3b. 消耗边：先来后到（预分组）---
        ws_rows, ws_cols, ws_data = [], [], []   # 物料#车间 → order
        wh_rows, wh_cols, wh_data = [], [], []   # 物料#仓库 → order
        remaining_ws = workshop_init_qty.copy()
        self._workshop_consumed = {mat: 0.0 for mat in self.material_nodes}
        self._warehouse_consumed = {mat: 0.0 for mat in self.material_nodes}

        io_by_mat = dict(list(self.io_df.groupby('材料编码')))

        for mat, mat_io in io_by_mat.items():
            if mat not in self._mat_ws_idx:
                continue

            ws_avail = workshop_avail.get(mat, 0)
            wh_avail = warehouse_avail.get(mat, 0)
            total_ws = 0.0
            total_wh = 0.0
            rem = remaining_ws.get(mat, 0)
            ws_idx = self._mat_ws_idx[mat]
            wh_idx = self._mat_wh_idx[mat]

            for _, row in mat_io.iterrows():
                order = str(row['工单号'])
                prod_in_row = str(row['产品编码'])
                issue = row['材料领用数量']

                if order not in self.node_index:
                    continue

                # 2-hop 自环：同一物料被工单领用且产出 → 跳过消费边
                if mat == prod_in_row:
                    self._self_loop_skipped.append({
                        'order': order, 'material': mat, 'product': prod_in_row,
                        'type': '2hop_self_loop', 'issue_qty': float(issue),
                    })
                    continue

                # 先车间后仓库
                from_ws = min(issue, rem)
                rem -= from_ws
                from_wh = issue - from_ws
                total_ws += from_ws
                total_wh += from_wh

                if ws_avail > 0 and from_ws > 0:
                    ws_rows.append(self.node_index[order])
                    ws_cols.append(ws_idx)
                    ws_data.append(from_ws / ws_avail)

                if wh_avail > 0 and from_wh > 0:
                    wh_rows.append(self.node_index[order])
                    wh_cols.append(wh_idx)
                    wh_data.append(from_wh / wh_avail)

            remaining_ws[mat] = rem
            self._workshop_consumed[mat] = total_ws
            self._warehouse_consumed[mat] = total_wh

        # --- 3c. 组装单一 W 矩阵 ---
        all_rows = out_rows + ws_rows + wh_rows
        all_cols = out_cols + ws_cols + wh_cols
        all_data = out_data + ws_data + wh_data
        W = sparse.csr_matrix((all_data, (all_rows, all_cols)), shape=(n, n))
        self.W_matrix = W
        self.W_workshop = None  # 不再需要独立车间W
        self.W_warehouse = None

        # --- 3d. 校验 ---
        col_sums = np.array(W.sum(axis=0)).flatten()

        # 车间物料节点列和 ≤ 1（自然保证）
        for mat in self.material_nodes:
            ws_idx = self._mat_ws_idx[mat]
            if col_sums[ws_idx] > 1 + 1e-6:
                cs = float(col_sums[ws_idx])
                scale = 1.0 / cs
                W_csc = W.tocsc()
                start, end = W_csc.indptr[ws_idx], W_csc.indptr[ws_idx + 1]
                W_csc.data[start:end] *= scale
                W = W_csc.tocsr()

        # 仓库物料节点列和 > 1 → 超领
        wh_over = []
        for mat in self.material_nodes:
            wh_idx = self._mat_wh_idx[mat]
            if col_sums[wh_idx] > 1 + 1e-6:
                wh_over.append((wh_idx, mat, float(col_sums[wh_idx])))

        if wh_over:
            if not force_calculate:
                bad_nodes = [mat for _, mat, _ in wh_over]
                raise ValueError(
                    f"以下物料存在超领（领用>可供发出），成本将扭曲：{bad_nodes}\n"
                    f"可勾选「强制计算」自动处理超领"
                )
            else:
                self._force_normalized = True
                W_csc = W.tocsc()
                for wh_idx, mat, cs in wh_over:
                    scale = 1.0 / cs
                    start, end = W_csc.indptr[wh_idx], W_csc.indptr[wh_idx + 1]
                    W_csc.data[start:end] *= scale
                    self._force_normalization_log.append({
                        'node': mat, 'col_sum': cs, 'correction_factor': cs,
                    })
                W = W_csc.tocsr()

        # 保险2：对角线自环校验
        diag = W.diagonal()
        if np.any(np.abs(diag) > 1e-6):
            bad_nodes = [self.all_nodes[i] for i in np.where(np.abs(diag) > 1e-6)[0]]
            raise ValueError(f"W矩阵存在自环节点：{bad_nodes}")

        self.performance_log['构建矩阵'] = time.time() - t0
        self.W_matrix = W

        # 保存可供发出数量
        self.available_qty = {}
        for mat in self.material_nodes:
            self.available_qty[mat] = workshop_avail[mat] + warehouse_avail[mat]
        self.sales_available_qty = sales_available_qty

        # ============================================================
        # Step 4: D 矩阵
        # ============================================================
        D_mat = np.ones(n)

        order_production = self.io_df.groupby('工单号').agg({
            '产品完工数量': 'sum', '在产品数量': 'sum'
        }).reset_index()

        for _, row in order_production.iterrows():
            order = str(row['工单号'])
            if order not in self.node_index:
                continue
            finished = row['产品完工数量']
            wip = row['在产品数量']
            total = finished + wip
            D_mat[self.node_index[order]] = finished / total if total > 0 else 1.0

        # 物料节点 D=1（车间和仓库节点都是）
        # 已初始化为全1，无需额外设置

        if np.any(D_mat < 0) or np.any(D_mat > 1):
            raise ValueError("D矩阵存在非法值")

        self.D_matrix = D_mat

        # ============================================================
        # Step 5: F 矩阵（单一，含期初工费）
        # ============================================================
        t0 = time.time()
        F = np.zeros((n, 3))

        for mat in self.material_nodes:
            ws_idx = self._mat_ws_idx[mat]
            wh_idx = self._mat_wh_idx[mat]
            # 车间节点：期初料+工+费
            F[ws_idx, 0] = workshop_init_mat[mat]
            F[ws_idx, 1] = workshop_init_lab[mat]
            F[ws_idx, 2] = workshop_init_oh[mat]
            # 仓库节点：期初料+工+费 + 采购
            F[wh_idx, 0] = warehouse_init_mat[mat] + warehouse_init_lab[mat] + warehouse_init_oh[mat] + pur_amt[mat]

        # 工单节点：人工、制费
        for _, row in self.labor_df.iterrows():
            order = str(row['工单号'])
            if order in self.node_index:
                idx = self.node_index[order]
                F[idx, 1] += row['人工']
                F[idx, 2] += row['制费']

        self.F_matrix = F
        self.F_workshop = None
        self.F_warehouse = None
        self.performance_log['构建F矩阵'] = time.time() - t0

        # ============================================================
        # Step 6: 求解 X = (I - W×D)^(-1) × F[:,0] + (I - W)^(-1) × F[:,1:3]
        # ============================================================
        t0 = time.time()
        I_sp = sparse.eye(n)

        W_mat_data = W.data.copy() * D_mat[W.indices]
        W_mat = sparse.csr_matrix((W_mat_data, W.indices.copy(), W.indptr.copy()), shape=(n, n))
        A_mat = I_sp - W_mat
        A_loh = I_sp - W

        # 微量正则化（1e-10 不影响精度，但确保 LU 分解稳定，供超级还原复用）
        A_mat = A_mat + sparse.eye(n) * 1e-10
        A_loh = A_loh + sparse.eye(n) * 1e-10
        if force_calculate and self._force_normalized:
            # 已有正则化，无需再加
            pass

        # splu 分解并保存 LU 因子（超级还原复用）
        from scipy.sparse.linalg import splu
        self._lu_mat = None
        self._lu_loh = None
        try:
            self._lu_mat = splu(A_mat.tocsc())
            self._lu_loh = splu(A_loh.tocsc())
            X_mat = self._lu_mat.solve(F[:, 0]).reshape(-1, 1)
            X_loh = self._lu_loh.solve(F[:, 1:3])
        except Exception:
            # 兜底：稠密求解 + 重试 splu（加更强正则化）
            X_mat = np.linalg.solve(A_mat.toarray(), F[:, 0:1])
            X_loh = np.linalg.solve(A_loh.toarray(), F[:, 1:3])
            try:
                A_mat2 = A_mat + sparse.eye(n) * 1e-8
                A_loh2 = A_loh + sparse.eye(n) * 1e-8
                self._lu_mat = splu(A_mat2.tocsc())
                self._lu_loh = splu(A_loh2.tocsc())
            except Exception:
                pass
        except Exception:
            X_mat = np.linalg.solve(A_mat.toarray(), F[:, 0:1])
            X_loh = np.linalg.solve(A_loh.toarray(), F[:, 1:3])

        X = np.hstack([X_mat, X_loh])
        X = np.round(X, 4)
        self.X_total = X

        # 为兼容性构建 X_workshop / X_warehouse（从独立节点合并）
        X_ws = np.zeros((n, 3))
        X_wh = np.zeros((n, 3))
        for mat in self.material_nodes:
            X_ws[self._mat_ws_idx[mat]] = X[self._mat_ws_idx[mat]]
            X_wh[self._mat_wh_idx[mat]] = X[self._mat_wh_idx[mat]]
        self.X_workshop = np.round(X_ws, 4)
        self.X_warehouse = np.round(X_wh, 4)

        self.performance_log['矩阵求解'] = time.time() - t0
        self.performance_log['矩阵维度'] = n
        
        # ==================== 销售成本计算（如果有销售数据）====================
        sales_cost_df = pd.DataFrame()
        if not self.sales_df.empty:
            t0 = time.time()

            # 按物料汇总销售数量和销售金额
            mat_sales = self.sales_df.groupby('物料编码')['销售数量'].sum().to_dict()
            mat_revenue = {}
            if '销售金额' in self.sales_df.columns:
                mat_revenue = self.sales_df.groupby('物料编码')['销售金额'].sum().to_dict()
            batch_sales = self.sales_df  # 已按(物料, 批次)聚合

            # 构建 Sale 对角阵（稀疏，基于仓库节点）
            S_diag = np.zeros(n)
            for mat, sales_qty in mat_sales.items():
                if mat in self._mat_wh_idx:
                    wh_idx = self._mat_wh_idx[mat]
                    avail = self.sales_available_qty.get(mat, 0)
                    if avail > 0:
                        S_diag[wh_idx] = sales_qty / avail
            S_sp = sparse.diags(S_diag, format='csr')

            # 构建 Batch 矩阵（稀疏 CSR，基于仓库节点）
            batch_list = sorted(batch_sales['销售批次号'].unique())
            m = len(batch_list)
            batch_index = {b: i for i, b in enumerate(batch_list)}

            b_rows, b_cols, b_data = [], [], []
            batch_sales_info = {}
            for _, row in batch_sales.iterrows():
                mat = str(row['物料编码'])
                batch = str(row['销售批次号'])
                qty = row['销售数量']
                revenue = float(row.get('销售金额', 0) or 0)
                if mat not in self._mat_wh_idx or batch not in batch_index:
                    continue
                wh_idx = self._mat_wh_idx[mat]
                batch_idx = batch_index[batch]
                total_sales = mat_sales.get(mat, 0)
                if total_sales > 0:
                    ratio = qty / total_sales
                    b_rows.append(batch_idx)
                    b_cols.append(wh_idx)
                    b_data.append(ratio)
                batch_sales_info[(batch, mat)] = {
                    'qty': qty,
                    'revenue': revenue,
                    'avail': self.sales_available_qty.get(mat, 0),
                    'sales_ratio': S_diag[wh_idx],
                    'batch_ratio': ratio if total_sales > 0 else 0,
                }

            B_sp = sparse.csr_matrix((b_data, (b_rows, b_cols)), shape=(m, n))

            # C = B × S × X  (稀疏加速)
            # SX(i,:) = S_diag[i] * X[i,:]  (对角阵乘等价于逐行缩放)
            SX_val = S_diag[:, np.newaxis] * X  # N×3, 广播乘法
            C = B_sp.dot(SX_val)  # sparse × dense → dense M×3

            # 构建销售成本结果（使用预存信息避免二次遍历）
            sales_rows = []
            for (batch, mat), info in batch_sales_info.items():
                mat_idx = self._mat_wh_idx[mat]
                batch_idx = batch_index[batch]
                cost_mat = info['batch_ratio'] * SX_val[mat_idx, 0]
                cost_lab = info['batch_ratio'] * SX_val[mat_idx, 1]
                cost_oh = info['batch_ratio'] * SX_val[mat_idx, 2]
                sales_rows.append({
                    '销售批次号': batch,
                    '物料编码': mat,
                    '销售数量': round(info['qty'], 2),
                    '销售金额': round(info['revenue'], 2),
                    '可供发出数量': round(info['avail'], 2),
                    '销售占比': f"{info['sales_ratio']:.2%}",
                    '批次占物料销售比': f"{info['batch_ratio']:.2%}",
                    '销售成本_料': round(cost_mat, 2),
                    '销售成本_工': round(cost_lab, 2),
                    '销售成本_费': round(cost_oh, 2),
                    '销售成本_合计': round(cost_mat + cost_lab + cost_oh, 2),
                })

            sales_cost_df = pd.DataFrame(sales_rows)
            if not sales_cost_df.empty:
                sales_cost_df = sales_cost_df.sort_values(['销售批次号', '物料编码'])

            self.performance_log['销售成本计算'] = time.time() - t0
        
        # 超级成本还原已拆分为独立方法 super_restore()，此处留空
        super_result = {
            '超级还原_完工成本': pd.DataFrame(),
            '超级还原_TopN汇总': pd.DataFrame(),
            '超级还原_维度定义': pd.DataFrame({'说明': ['请调用 super_restore() 方法']}),
            '超级还原_验证差异': pd.DataFrame(),
        }

        # Step 7: 计算WIP成本（单矩阵）
            
        # Step 7: 计算WIP成本（单矩阵）
        # 公式：WIP = W × (X[:,0] × (1-D))
        t0 = time.time()
        I_minus_D = 1 - D_mat
        wip_at_order = X[:, 0] * I_minus_D
        WIP_col = W.dot(wip_at_order)
        WIP_full = np.zeros((n, 3))
        WIP_full[:, 0] = WIP_col
        WIP_full = np.round(WIP_full, 4)
        self.performance_log['在产品计算'] = time.time() - t0

        # Step 8: 结果整理
        # 8.0 完工入库金额（从工单聚合，含料工费）
        # X[wh_idx] 已通过单矩阵自动包含车间传导成本，这里汇总每个产品的工单转出额
        finished_cost_per_mat = {}
        for mat in self.material_nodes:
            finished_cost_per_mat[mat] = 0.0

        for _, row in order_prod_finished.iterrows():
            order = str(row['工单号'])
            prod = str(row['产品编码'])
            finished = row['产品完工数量']
            if order not in self.node_index or prod not in self._mat_wh_idx:
                continue

            order_idx = self.node_index[order]
            prod_wh_idx = self._mat_wh_idx[prod]
            total = order_total_finished.get(order, 0)

            if total > 0:
                w_ratio = finished / total
            else:
                n_products = len(order_prod_finished[order_prod_finished['工单号'] == order])
                w_ratio = 1.0 / n_products if n_products > 0 else 0

            finished_ratio = D_mat[order_idx]
            order_mat = X[order_idx, 0]
            order_lab = X[order_idx, 1]
            order_oh = X[order_idx, 2]

            finished_cost_per_mat[prod] += (
                w_ratio * order_mat * finished_ratio +
                w_ratio * order_lab +
                w_ratio * order_oh
            )

        # 8.1 收发存汇总表（纯会计逻辑：收入=采购金额+完工入库金额）
        # 按物料汇总销售数量（如果有销售数据）
        sales_summary = {}
        if not self.sales_df.empty:
            sales_summary = self.sales_df.groupby('物料编码')['销售数量'].sum().to_dict()

        result_list = []
        for mat in self.material_nodes:
            # 期初（仓库口径，不含车间）
            init_q = warehouse_init_qty[mat]
            init_a = warehouse_init_mat[mat]

            # 采购
            pur_q = pur_qty[mat]
            pur_a = pur_amt[mat]

            # 完工入库
            fin_q = finished_qty_for_report[mat]
            fin_a = finished_cost_per_mat.get(mat, 0)

            # 收入 = 采购 + 完工入库
            receipt_qty = pur_q + fin_q
            receipt_amt = pur_a + fin_a

            # 发出 = 仓库被领用 + 销售出库
            issue_q = self._warehouse_consumed.get(mat, 0) + sales_summary.get(mat, 0)

            # 月末一次加权平均（归一化后的X已是最终成本口径，不需额外矫正）
            total_cost = init_a + receipt_amt
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
                '本月领用数量': round(self._warehouse_consumed.get(mat, 0), 4),
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
            
            if order not in self.node_index or product not in self._mat_wh_idx:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self._mat_wh_idx[product]
            
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
            
            if order not in self.node_index or product not in self._mat_wh_idx:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self._mat_wh_idx[product]
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
            
            if order not in self.node_index or product not in self._mat_wh_idx or material not in self._mat_ws_idx:
                continue
            
            order_idx = self.node_index[order]
            prod_idx = self._mat_wh_idx[product]

            # 消耗比例：合并车间+仓库两个节点
            w_mat_to_order = (
                W[order_idx, self._mat_ws_idx[material]] +
                W[order_idx, self._mat_wh_idx[material]]
            )
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
            
            # S_人工 = (I + W_h × D) × F_工
            # S_制费 = (I + W_h × D) × F_制费
            W_step_data = W.data.copy() * D_mat[W.indices]
            W_step_mat = sparse.csr_matrix(
                (W_step_data, W.indices.copy(), W.indptr.copy()), shape=(n, n))
            I_plus_WD = I_sp + W_step_mat  # I + W_h × D_mat
            F_combined = self.F_matrix

            S_lab = I_plus_WD.dot(F_combined[:, 1])  # 人工列
            S_oh = I_plus_WD.dot(F_combined[:, 2])   # 制费列

            # S_材料 = X_combine总 - S_人工 - S_制费
            X_combine_total = X.sum(axis=1)  # 每行的料+工+费总和
            S_mat = X_combine_total - S_lab - S_oh

            # 构建 S 矩阵 (n × 3)
            S = np.column_stack([S_mat, S_lab, S_oh])
            S = np.round(S, 4)

            # 逐步结转法下的 WIP = W_h × (I-D) × S_材料
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
                
                if order not in self.node_index or product not in self._mat_wh_idx:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self._mat_wh_idx[product]
                
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
                
                if order not in self.node_index or product not in self._mat_wh_idx:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self._mat_wh_idx[product]
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
                
                if order not in self.node_index or product not in self._mat_wh_idx or material not in self._mat_ws_idx:
                    continue
                
                order_idx = self.node_index[order]
                prod_idx = self._mat_wh_idx[product]

                # 消耗比例：合并车间+仓库两个节点
                w_mat_to_order = (
                    W[order_idx, self._mat_ws_idx[material]] +
                    W[order_idx, self._mat_wh_idx[material]]
                )
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
        # 销售金额可能不存在（旧格式），安全兜底
        if '销售金额' in df.columns:
            df['销售金额'] = pd.to_numeric(df['销售金额'], errors='coerce').fillna(0)
        else:
            df['销售金额'] = 0

        # 按物料汇总销售数量
        mat_sales = df.groupby('物料编码')['销售数量'].sum().to_dict()

        # 按(物料, 批次)汇总销售数量和销售金额
        batch_sales = df.groupby(['物料编码', '销售批次号']).agg(
            销售数量=('销售数量', 'sum'),
            销售金额=('销售金额', 'sum'),
        ).reset_index()
        
        n = len(self.all_nodes)
        
        # ========== 构建 Sale 矩阵（N×N 对角阵）==========
        S = np.zeros((n, n))
        for mat, sales_qty in mat_sales.items():
            if mat in self.node_index:
                idx = self._mat_wh_idx[mat]
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
                mat_idx = self._mat_wh_idx[mat]
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
            revenue = float(row.get('销售金额', 0) or 0)

            if mat not in self.node_index:
                continue

            mat_idx = self._mat_wh_idx[mat]
            batch_idx = batch_index.get(batch)
            if batch_idx is None:
                continue

            avail = self.available_qty.get(mat, 0)
            total_sales = mat_sales.get(mat, 0)
            sales_ratio = S[mat_idx, mat_idx]  # 销售占比 = 总销售 / 可供发出
            batch_ratio = B[batch_idx, mat_idx]  # 批次占比 = 批次销售 / 总销售

            cost_mat = batch_ratio * SX[mat_idx, 0]
            cost_lab = batch_ratio * SX[mat_idx, 1]
            cost_oh = batch_ratio * SX[mat_idx, 2]
            cost_total = cost_mat + cost_lab + cost_oh

            result_list.append({
                '销售批次号': batch,
                '物料编码': mat,
                '销售数量': round(qty, 2),
                '销售金额': round(revenue, 2),
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
                idx = self._mat_wh_idx[mat]
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
    
    def super_restore(self, top_products=None, force_calculate=False):
        """超级成本还原：将 X 拆解到原始成本维度。

        需在 calculate() 之后调用。

        Parameters
        ----------
        top_products : list, optional
            限定只还原的产品编码列表。None = 全部产品。
        force_calculate : bool
            与 calculate() 保持一致。
        """
        if self._lu_mat is None or self._lu_loh is None:
            return {
                '超级还原_完工成本': pd.DataFrame(),
                '超级还原_TopN汇总': pd.DataFrame(),
                '超级还原_维度定义': pd.DataFrame({'说明': ['LU 分解不可用，请先用 force_calculate=True 运行']}),
                '超级还原_验证差异': pd.DataFrame(),
            }

        t0 = time.time()
        n = len(self.all_nodes)

        # 收集维度
        super_dims = []
        dim_info = []
        mat_src_nodes = []
        mat_values = []
        loh_src_nodes = []
        loh_values = []

        for _, row in self.initial_df.iterrows():
            mat = str(row['物料编码']).strip()
            inv_type = str(row.get('库存类型', '仓库')).strip()
            amt = row.get('期初材料', 0) + row.get('期初人工', 0) + row.get('期初制费', 0)
            if mat in self._mat_wh_idx and amt > 0:
                node_idx = self._mat_ws_idx[mat] if inv_type == '车间' else self._mat_wh_idx[mat]
                suffix = '#车间' if inv_type == '车间' else '#仓库'
                super_dims.append(f"{mat}{suffix}_期初")
                mat_src_nodes.append(node_idx)
                mat_values.append(amt)
                dim_info.append({
                    '维度名称': f"{mat}{suffix}_期初", '维度类型': '期初',
                    '来源节点': mat, '原始金额': amt,
                    '说明': f'物料{mat}({inv_type})的期初金额'
                })

        for _, row in self.purchase_df.iterrows():
            mat = str(row['物料编码']).strip()
            amt = row['采购金额']
            if mat in self._mat_wh_idx and amt > 0:
                super_dims.append(f"{mat}#仓库_采购")
                mat_src_nodes.append(self._mat_wh_idx[mat])
                mat_values.append(amt)
                dim_info.append({
                    '维度名称': f"{mat}#仓库_采购", '维度类型': '采购',
                    '来源节点': mat, '原始金额': amt,
                    '说明': f'物料{mat}的采购金额'
                })

        for _, row in self.labor_df.iterrows():
            order = str(row['工单号']).strip()
            amt = row['人工'] + row['制费']
            if order in self.node_index and amt > 0:
                super_dims.append(f"{order}_人工制费")
                loh_src_nodes.append(self.node_index[order])
                loh_values.append(amt)
                dim_info.append({
                    '维度名称': f"{order}_人工制费", '维度类型': '工费',
                    '来源节点': order, '原始金额': amt,
                    '说明': f'工单{order}的直接人工+制造费用'
                })

        n_dims = len(super_dims)
        if n_dims == 0:
            self.performance_log['超级成本还原'] = 0
            self.performance_log['超级还原维度数'] = 0
            return {
                '超级还原_完工成本': pd.DataFrame(),
                '超级还原_TopN汇总': pd.DataFrame(),
                '超级还原_维度定义': pd.DataFrame({'说明': ['未找到任何成本维度数据']}),
                '超级还原_验证差异': pd.DataFrame(),
            }

        # 确定要还原的产品范围
        if top_products is not None:
            product_set = set(top_products)
            mat_indices = np.array([
                self._mat_wh_idx[m] for m in self.material_nodes if m in product_set
            ])
            mat_names = [m for m in self.material_nodes if m in product_set]
        else:
            mat_indices = np.array([self._mat_wh_idx[m] for m in self.material_nodes])
            mat_names = list(self.material_nodes)

        n_mat = len(mat_indices)
        X = self.X_total
        mat_super_X = np.zeros((n_mat, n_dims))
        BATCH = 1000

        def _solve_group(lu, src_nodes, values, dim_start):
            if len(src_nodes) == 0:
                return
            src_arr = np.array(src_nodes)
            val_arr = np.array(values)
            unique_src, inverse = np.unique(src_arr, return_inverse=True)
            n_uniq = len(unique_src)

            for start in range(0, n_uniq, BATCH):
                end = min(start + BATCH, n_uniq)
                batch_src = unique_src[start:end]
                k = end - start
                F_batch = sparse.csc_matrix(
                    (np.ones(k), (batch_src, np.arange(k))), shape=(n, k))
                Y_batch = lu.solve(F_batch.toarray())
                Z_batch = Y_batch[mat_indices, :]

                batch_mask = (inverse >= start) & (inverse < end)
                batch_dims = np.where(batch_mask)[0]
                batch_rel = inverse[batch_dims] - start
                mat_super_X[:, dim_start + batch_dims] = (
                    Z_batch[:, batch_rel] * val_arr[batch_dims])

        _solve_group(self._lu_mat, mat_src_nodes, mat_values, 0)
        _solve_group(self._lu_loh, loh_src_nodes, loh_values, len(mat_src_nodes))

        mat_super_X = np.round(mat_super_X, 4)

        # 结果整理
        mat_total_cost = X[mat_indices].sum(axis=1)
        dim_total = mat_super_X.sum(axis=0)

        mask = np.abs(mat_super_X) >= 0.01
        rows, cols = np.where(mask)

        long_format_rows = []
        product_dim_ratios = {}
        for i in range(len(rows)):
            r, c = rows[i], cols[i]
            mat = mat_names[r]
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
            super_detail_long.insert(1, '成本序号', super_detail_long.groupby('产品编码').cumcount() + 1)

        super_topn_summary = (
            super_detail_long.groupby('产品编码').head(5).reset_index(drop=True)
            if not super_detail_long.empty else pd.DataFrame()
        )

        super_dim_definition = pd.DataFrame(dim_info)

        mat_super_sums = mat_super_X.sum(axis=1)
        max_diff = float(np.max(np.abs(mat_super_sums - mat_total_cost)))
        bad_mask = np.abs(mat_super_sums - mat_total_cost) > 1
        bad_idx = np.where(bad_mask)[0]
        verification_rows = []
        for r in bad_idx:
            verification_rows.append({
                '产品编码': mat_names[r],
                '超级还原合计': round(float(mat_super_sums[r]), 2),
                '标准总成本': round(float(mat_total_cost[r]), 2),
                '差异': round(float(np.abs(mat_super_sums[r] - mat_total_cost[r])), 2)
            })
        super_verification = pd.DataFrame(verification_rows) if verification_rows else pd.DataFrame()

        super_result = {
            '超级还原_完工成本': super_detail_long,
            '超级还原_TopN汇总': super_topn_summary,
            '超级还原_维度定义': super_dim_definition,
            '超级还原_验证差异': (
                super_verification if not super_verification.empty
                else pd.DataFrame({'说明': [f'所有产品勾稽验证通过，最大差异: {max_diff:.2f}']})
            ),
        }

        elapsed = time.time() - t0
        self.performance_log['超级成本还原'] = round(elapsed, 3)
        self.performance_log['超级还原维度数'] = n_dims
        return super_result

    def get_performance(self):
        return self.performance_log


def calculate_period(data_dict, finished_df=None, finished_map=None,
                     calculate_step_method=False, force_calculate=False):
    """对单月份数据执行成本核算"""
    calc = CostCalculator()
    calc.load_data(data_dict)
    if finished_df is None and 'finished' in data_dict:
        finished_df = data_dict['finished']
    result = calc.calculate(
        finished_df=finished_df,
        finished_map=finished_map,
        calculate_step_method=calculate_step_method,
        force_calculate=force_calculate
    )
    return result, calc


def write_debug_log(calc, filepath):
    """输出双矩阵调试日志到文件。

    Parameters
    ----------
    calc : CostCalculator
        已执行 calculate() 的实例
    filepath : str
        日志文件路径
    """
    with open(filepath, 'w', encoding='utf-8') as f:
        sep = "=" * 60
        f.write(f"{sep}\n")
        f.write("  成本核算调试日志（双矩阵版）\n")
        f.write(f"{sep}\n\n")

        # ---- S1: 节点统计 ----
        mn = len(calc.material_nodes) if calc.material_nodes else 0
        on_val = len(calc.order_nodes) if calc.order_nodes else 0
        f.write(f"【S1】节点统计\n")
        f.write(f"  物料节点: {mn}\n")
        f.write(f"  工单节点: {on_val}\n")
        f.write(f"  总节点: {mn + on_val}\n\n")

        # ---- S2: 可供发出数量 ----
        f.write(f"【S2】可供发出数量\n")
        if calc.available_qty:
            qtys = np.array(list(calc.available_qty.values()))
            f.write(f"  总计: min={float(qtys.min()):.2f}, max={float(qtys.max()):.2f}, "
                    f"mean={float(qtys.mean()):.2f}, zero={int((qtys == 0).sum())}\n")
            top5 = sorted(calc.available_qty.items(), key=lambda x: x[1], reverse=True)[:5]
            f.write(f"  Top5:\n")
            for mat, q in top5:
                f.write(f"    {mat}: {q:.2f}\n")
        else:
            f.write("  无数据\n")
        # 车间/仓库消耗摘要
        if calc._workshop_consumed:
            ws_vals = list(calc._workshop_consumed.values())
            wh_vals = list(calc._warehouse_consumed.values())
            ws_nonzero = sum(1 for v in ws_vals if v > 0)
            wh_nonzero = sum(1 for v in wh_vals if v > 0)
            f.write(f"  车间消耗: {ws_nonzero}物料非零, 总计{sum(ws_vals):.2f}\n")
            f.write(f"  仓库消耗: {wh_nonzero}物料非零, 总计{sum(wh_vals):.2f}\n")
        f.write("\n")

        # ---- S3: W矩阵统计 ----
        f.write(f"【S3】W矩阵\n")
        # 单矩阵模式：直接输出 W 统计
        if calc.W_matrix is not None:
            w_mat = calc.W_matrix
            nnz = w_mat.nnz
            n_total = (mn + on_val)
            sparsity = nnz / n_total**2 * 100 if n_total > 0 else 0
            col_sums = np.array(w_mat.sum(axis=0)).flatten()
            over_one = int((col_sums > 1 + 1e-6).sum())
            diag = w_mat.diagonal()
            loop_idx = np.where(np.abs(diag) > 1e-6)[0]
            f.write(f"  [W] nnz={nnz}, 稀疏度={sparsity:.4f}%, 维度={n_total}\n")
            f.write(f"    列和: min={float(col_sums.min()):.6f}, max={float(col_sums.max()):.6f}, "
                    f"mean={float(col_sums.mean()):.6f}, >1={over_one}\n")
            if len(loop_idx) > 0:
                nodes = [calc.all_nodes[i] for i in loop_idx] if calc.all_nodes else []
                f.write(f"    [WARN] 自环节点: {nodes}\n")
            else:
                f.write(f"    自环节点: 0\n")
        else:
            f.write(f"  [W] 无数据\n")
            if w_mat is not None:
                nnz = w_mat.nnz
                sparsity = nnz / (mn + on_val) ** 2 * 100 if (mn + on_val) > 0 else 0
                col_sums = np.array(w_mat.sum(axis=0)).flatten()
                over_one = int((col_sums > 1 + 1e-6).sum())
                diag = w_mat.diagonal()
                loop_idx = np.where(np.abs(diag) > 1e-6)[0]
                f.write(f"  [{w_name}] nnz={nnz}, 稀疏度={sparsity:.4f}%\n")
                f.write(f"    列和: min={float(col_sums.min()):.6f}, max={float(col_sums.max()):.6f}, "
                        f"mean={float(col_sums.mean()):.6f}, >1={over_one}\n")
                if len(loop_idx) > 0:
                    nodes = [calc.all_nodes[i] for i in loop_idx] if calc.all_nodes else []
                    f.write(f"    [WARN] 自环节点: {nodes}\n")
                else:
                    f.write(f"    自环节点: 0\n")
            else:
                f.write(f"  [{w_name}] 无数据\n")
        f.write("\n")

        # ---- S4: D矩阵 ----
        f.write(f"【S4】D矩阵\n")
        if calc.D_matrix is not None and calc.material_nodes and calc.node_index:
            D = calc.D_matrix
            mat_idx_list = [calc._mat_wh_idx[m] for m in calc.material_nodes if m in calc._mat_wh_idx]
            d1_count = sum(1 for i in mat_idx_list if abs(D[i] - 1.0) < 1e-10)
            f.write(f"  物料D=1: {d1_count}/{len(mat_idx_list)}\n")
            ord_idx_list = [calc.node_index[o] for o in calc.order_nodes if o in calc.node_index]
            if ord_idx_list:
                d_vals = [D[i] for i in ord_idx_list]
                f.write(f"  工单D: min={min(d_vals):.4f}, max={max(d_vals):.4f}, "
                        f"<1的数量={sum(1 for d in d_vals if d < 1.0)}\n")
        else:
            f.write("  无数据\n")
        f.write("\n")

        # ---- S5: F矩阵 ----
        f.write(f"【S5】F矩阵\n")
        for f_mat, f_name in [(calc.F_workshop, '车间'), (calc.F_warehouse, '仓库')]:
            if f_mat is not None:
                for ci, cn in enumerate(['材料', '人工', '制费']):
                    col = f_mat[:, ci]
                    f.write(f"  [{f_name}] {cn}: sum={float(col.sum()):.2f}, "
                            f"non-zero={int(np.count_nonzero(col))}\n")
            else:
                f.write(f"  [{f_name}] 无数据\n")
        f.write("\n")

        # ---- S6: X矩阵 ----
        f.write(f"【S6】X矩阵求解结果\n")
        if calc.X_total is not None:
            X = calc.X_total
            for ci, cn in enumerate(['材料', '人工', '制费']):
                col = X[:, ci]
                f.write(f"  {cn}: min={float(col.min()):.4f}, max={float(col.max()):.4f}, "
                        f"sum={float(col.sum()):.2f}, NaN={int(np.isnan(col).sum())}, "
                        f"Inf={int(np.isinf(col).sum())}, 全零行={int((col == 0).sum())}\n")
            # 车间/仓库分别统计
            if hasattr(calc, 'X_workshop') and calc.X_workshop is not None:
                Xw = calc.X_workshop
                f.write(f"  其中车间X: 料sum={float(Xw[:,0].sum()):.2f}, "
                        f"工sum={float(Xw[:,1].sum()):.2f}, 费sum={float(Xw[:,2].sum()):.2f}\n")
            if hasattr(calc, 'X_warehouse') and calc.X_warehouse is not None:
                Xh = calc.X_warehouse
                f.write(f"  其中仓库X: 料sum={float(Xh[:,0].sum()):.2f}\n")
        else:
            f.write("  无数据\n")
        f.write("\n")

        # ---- S7: 收发存摘要 ----
        f.write(f"【S7】收发存汇总\n")
        if calc.result and '收发存' in calc.result and not calc.result['收发存'].empty:
            df = calc.result['收发存']
            for col in ['期初数量', '本期收入数量', '本月发出数量', '期末数量', '期末金额']:
                if col in df.columns:
                    f.write(f"  {col}: min={float(df[col].min()):.2f}, "
                            f"max={float(df[col].max()):.2f}, sum={float(df[col].sum()):.2f}\n")
        else:
            f.write("  无数据\n")
        f.write("\n")

        # ---- 自环跳过记录 ----
        if calc._self_loop_skipped:
            f.write(f"【自环跳过】{len(calc._self_loop_skipped)}条\n")
            for entry in calc._self_loop_skipped[:10]:
                f.write(f"  工单={entry['order']}, 物料={entry['material']}, "
                        f"产品={entry['product']}, 领用={entry['issue_qty']}\n")
            if len(calc._self_loop_skipped) > 10:
                f.write(f"  ... 共{len(calc._self_loop_skipped)}条\n")
            f.write("\n")

        # ---- 强制计算记录 ----
        if calc._force_normalized and calc._force_normalization_log:
            f.write(f"【强制计算】{len(calc._force_normalization_log)}个物料被归一化\n")
            for entry in calc._force_normalization_log:
                f.write(f"  物料={entry['node']}, 原始列和={entry['col_sum']:.6f}, "
                        f"矫正系数={entry['correction_factor']:.6f}\n")
            f.write("\n")

        # ---- 矩阵维度 ----
        dim = calc.performance_log.get('矩阵维度', 'N/A') if calc.performance_log else 'N/A'
        f.write(f"矩阵维度: {dim}\n")
        f.write(f"{sep}\n")


def to_excel(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)
    return output
