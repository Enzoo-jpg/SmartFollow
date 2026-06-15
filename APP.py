import streamlit as st
import pandas as pd
import io

# 设置网页标题和图标
st.set_page_config(page_title="强生随访核验筛选系统", page_icon="📋", layout="centered")

st.title("📋 强生随访核验筛选系统")

# =================【🔥 严格按照您要求的表头字眼及顺序，一字不差】=================
TARGET_EXCEL_HEADERS = [
    '随访任务ID',
    '门店编码',
    '门店名称',
    '会员卡号或患者ID',
    '商品ID编码',
    '商品名',
    '化学名',
    '商品规格',
    '剂型',
    '随访时间',
    '随访日志'
]

# =================【🔥 核心：全量商品四维主数据库（已同步最新化学名与规格）】=================
PRODUCT_MASTER_LIST = [
    {"id_code": "2110529", "name": "兆珂", "chemical": "达雷妥尤单抗注射液", "spec": "400mg/20ml/瓶", "form": "针剂"},
    {"id_code": "2120346", "name": "兆珂速", "chemical": "达雷妥尤单抗注射液(皮下注射)", "spec": "1800mg(15ml)/1瓶/盒", "form": "针剂"},
    {"id_code": "1119657-1", "name": "特诺雅", "chemical": "古塞奇尤单抗注射液", "spec": "100mg/1ml/支(预充笔式注射器)", "form": "针剂"},
    {"id_code": "2110530", "name": "兆珂", "chemical": "达雷妥尤单抗注射液", "spec": "100mg/5ml/瓶/盒", "form": "针剂"},
    {"id_code": "11220278", "name": "安森珂", "chemical": "阿帕他胺片", "spec": "60mg*120片/瓶/盒", "form": "片剂"},
    {"id_code": "2123222S", "name": "优拓比", "chemical": "司来帕格片", "spec": "0.2mg*60片/盒", "form": "片剂"},
    {"id_code": "2123221S", "name": "优拓比", "chemical": "司来帕格片", "spec": "0.8mg*60片/盒", "form": "片剂"},
    {"id_code": "2110623", "name": "特诺雅达", "chemical": "古塞奇尤单抗注射液(静脉输注)", "spec": "200mg/20mL/瓶x1瓶/盒", "form": "针剂"},
    {"id_code": "2123224", "name": "优拓比", "chemical": "司来帕格片", "spec": "0.6mg*60片/盒", "form": "片剂"},
    {"id_code": "2123367", "name": "傲朴舒", "chemical": "马昔腾坦片", "spec": "10mg*30片/盒", "form": "片剂"},
    {"id_code": "1111245-1", "name": "类克", "chemical": "注射用英夫利西单抗", "spec": "100mg/瓶/盒", "form": "针剂"},
    {"id_code": "2123326", "name": "亿珂", "chemical": "伊布替尼胶囊", "spec": "140mg*90粒/盒", "form": "片剂"}
]

# =================【这里是你的药房名字映射配置区】=================
PHARMACY_MAPPING = {
    "DTC国药控股四川专业药房连锁有限公司金牛区一环路西三段药房": "国药控股四川专业药房连锁有限公司金牛区一环路西三段药房",
    "DTC国药控股四川专业药房连锁有限公司达州药房": "国药控股四川专业药房连锁有限公司达州药房", 
    "DTC国药控股德阳有限公司泰山路关爱大药房": "国药控股德阳有限公司泰山路关爱大药房",
    "国药康禾成都医药有限公司高新区和盛东街分公司": "国药康禾成都医药有限公司高新区和盛东街分公司",
    "DTC国药控股广元医药有限公司关爱大药房": "国药控股广元医药有限公司关爱大药房",
    "国药控股达州有限公司北外药房": "国药控股达州有限公司北外药房",
    "DTC国药控股四川医药股份有限公司眉山药房": "国药控股四川医药股份有限公司眉山药房",
    "国药控股四川医药股份有限公司新都区民生巷药房": "国药控股四川医药股份有限公司新都区民生巷药房",
    "DTC国药控股四川医药股份有限公司西昌便民药房": "国药控股四川医药股份有限公司西昌便民药房",
    "国药控股四川专业药房连锁有限公司南充店": "国药控股四川专业药房连锁有限公司南充店",
    "国药控股四川专业药房连锁有限公司遂宁药房": "国药控股四川专业药房连锁有限公司遂宁药房",
    "DTC四川环晟大药房有限公司": "四川环晟大药房有限公司",
    "国药控股四川医药股份有限公司攀枝花益康街药房": "国药控股四川医药股份有限公司攀枝花益康街药房",
    "DTC国药控股四川医药股份有限公司南充药房": "国药控股四川医药股份有限公司南充药房",
    "DTC国药控股内江有限公司第一大药房": "国药控股内江有限公司第一大药房",
    "DTC国药控股四川专业药房连锁有限公司雅安药房": "国药控股四川专业药房连锁有限公司雅安药房",
    "国药控股内江有限公司第二大药房": "国药控股内江有限公司第二大药房",
    "四川环晟大药房有限公司蜀南大道店": "四川环晟大药房有限公司蜀南大道店",
    "国药控股(乐山)川药医药有限公司滨河路店": "国药控股（乐山）川药医药有限公司乐山高新区店",
    "国药控股昊阳绵阳药业有限公司江油匡山路大药房": "国药控股昊阳绵阳药业有限公司江油匡山路大药房",
    "DTC国药控股四川专业药房连锁有限公司攀枝花药房": "国药控股四川专业药房连锁有限公司攀枝花药房",
    "DTC国药控股广安有限公司广安药房": "国药控股广安有限公司广安药房",
    "DTC国药控股四川医药股份有限公司泸州药房": "国药控股四川医药股份有限公司泸州药房",
    "DTC国药控股四川专业药房连锁有限公司资阳药房": "国药控股四川专业药房连锁有限公司资阳药房",
}

# =================【这里是你的门店编码映射配置区】=================
CODE_MAPPING = {
    "DM0900444": "9025007",   
    "DM0974343": "9025005",  
    "DM0560022": "9002001",  
    "DM0971153": "9026001",  
    "DM0766016": "9014001",  
    "DM1020863": "9033001",  
    "DM0766017": "9019001",  
    "DM0619615": "9008001",  
    "DM0716360": "9001001",  
    "D353383": "9025009",  
    "DM1062799": "9034001",  
    "DM0606864": "000",  
    "DM0710188": "9017001",  
    "DM0693426": "9012001",  
    "DM0690466": "9009001",  
    "DM1086439": "9025014",  
    "DM0680333": "9025007",  
    "DM1121464": "006",  
    "DM0802780": "9027001",  
    "DM1074241": "9030001",  
    "DM0606859": "9025012",  
    "DM0900441": "9011001",  
    "DM0641557": "9007001",  
    "DM1012587": "9025006",  
}
# =============================================================

st.markdown("""
### 💡 使用前须知：
1. **历史销售底表**：系统默认读取您上传的 Excel 文件的**第一个工作表**。表头必须包含：*销售时间、商品名称、门店名称、门店code、会员卡号、规格*。
2. **已完成随访记录表（可选）**：上传后系统会自动抓取其中的 *患者oneId、药品名称、门店、门店编码* 列进行差额补访比对。
3. **保留原始规格现状**：应用户要求，导出的最终随访表中的 `商品规格` 将**忠实保留和还原您销售底表中的字眼不变**（底层已加入智能容错识别，如自动关联片剂/粒剂、中英文括号等，确保不影响编码及化学名的匹配精度）。
""", unsafe_allow_html=True)

st.divider()

# =================【标准模板下载区】=================
st.markdown("### 📥 官方标准模板下载")
with st.expander("👉 如果忘记表头或怕格式有误，请点此展开下载标准模板"):
    df_base_tpl = pd.DataFrame(columns=['销售时间', '商品名称', '门店名称', '门店code', '会员卡号', '规格'])
    df_base_tpl.loc[0] = ['2026-05-01', '特诺雅', '国药控股四川专业药房连锁有限公司金牛区一环路西三段药房', '9025007', 'HZ001', '100mg']
    base_out = io.BytesIO()
    with pd.ExcelWriter(base_out, engine='openpyxl') as writer:
        df_base_tpl.to_excel(writer, index=False, sheet_name='Sheet1')
    base_tpl_bytes = base_out.getvalue()

    df_hist_tpl = pd.DataFrame(columns=['患者oneId', '药品名称', '门店', '门店编码'])
    df_hist_tpl.loc[0] = ['HZ001', '特诺雅-(古塞奇尤单抗注射液)', '国药控股四川专业药房连锁有限公司金牛区一环路西三段药房', '9025007']
    hist_out = io.BytesIO()
    with pd.ExcelWriter(hist_out, engine='openpyxl') as writer:
        df_hist_tpl.to_excel(writer, index=False, sheet_name='已完成随访记录')
    hist_tpl_bytes = hist_out.getvalue()

    st.download_button(label="📥 下载《1. 销售底表模板》", data=base_tpl_bytes, file_name="1_销售底表新标准模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button(label="📥 下载《2. 已完成随访模板》", data=hist_tpl_bytes, file_name="2_已完成随访记录表模板.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()

# 1. 自动生成月份列表
month_range = pd.date_range(start="2025-10-01", end="2028-12-01", freq="MS")
month_options = [dt.strftime("%Y年%m月") for dt in month_range]

selected_month_display = st.selectbox(
    "1. 请选择需要随访的月份：", 
    options=month_options, 
    index=month_options.index("2026年05月") if "2026年05月" in month_options else 0
)
target_month_str = pd.to_datetime(selected_month_display, format="%Y年%m月").strftime("%Y-%m")

# 2. 用户上传文件
uploaded_file = st.file_uploader("2. 上传历史销售底表 (.xlsx)", type=["xlsx"])

# 3. 用户可选上传已随访表格
uploaded_history = st.file_uploader("3. （可选）上传本月已完成的随访记录表 (.xlsx) —— 用于差额补访比对", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, dtype=str)
        df.columns = df.columns.str.strip()
        
        # 智能模糊匹配新旧表头列名
        time_col = next((col for col in df.columns if "时间" in col), None)
        name_col = next((col for col in df.columns if "商品" in col), None)
        pharmacy_col = next((col for col in df.columns if "门店名称" in col or "药房" in col), None)
        code_col = next((col for col in df.columns if "code" in col.lower() or "编码" in col), None) 
        
        id_col = next((col for col in df.columns if "卡号" in col or "会员" in col or "id" in col.lower()), None)
        if not id_col:
            id_col = next((col for col in df.columns if "患者" in col and "姓名" not in col), None)
            
        spec_col = next((col for col in df.columns if "规格" in col), None)
        
        if not all([time_col, name_col, pharmacy_col, code_col, id_col, spec_col]): 
            st.error("❌ 错误：底表中缺少必要列！请确保包含：销售时间、商品名称、门店名称、门店code、会员卡号、规格。")
        else:
            result_df = pd.DataFrame()
            followup_patients = pd.DataFrame()
            has_history = False
            hist_warning_flag = False
            
            with st.spinner("📊 正在读取数据并进行智能随访计算..."):
                # 基础映射与清洗
                df[pharmacy_col] = df[pharmacy_col].astype(str).str.strip().replace(PHARMACY_MAPPING)
                df[code_col] = df[code_col].astype(str).str.strip().replace(CODE_MAPPING)
                df[id_col] = df[id_col].astype(str).str.strip().str.lstrip(',')
                
                df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
                df = df.dropna(subset=[time_col]).sort_values(by=time_col, ascending=True)
                df['购买月份'] = df[time_col].dt.strftime('%Y-%m')
                
                # 4字段拼接组合成患者唯一KEY
                df['患者唯一KEY'] = (
                    df[name_col].astype(str).str.strip() + "_" + 
                    df[pharmacy_col].astype(str).str.strip() + "_" + 
                    df[code_col].astype(str).str.strip() + "_" + 
                    df[id_col].astype(str).str.strip()
                )
                
                df_first = df.drop_duplicates(subset=['患者唯一KEY'], keep='first').copy()
                df_first['首次购买月份'] = df_first[time_col].dt.strftime('%Y-%m')
                
                # 筛选老患者
                followup_patients = df_first[df_first['首次购买月份'] < target_month_str]
                
                if not followup_patients.empty:
                    result_df = pd.DataFrame({
                        '商品名称': followup_patients[name_col],
                        '门店名称': followup_patients[pharmacy_col],
                        '门店code': followup_patients[code_col],  
                        '会员卡号': followup_patients[id_col],    
                        '规格': followup_patients[spec_col]      
                    })
                    
                    if uploaded_history is not None:
                        try:
                            df_history = pd.read_excel(uploaded_history, dtype=str)
                            df_history.columns = df_history.columns.str.strip()
                            
                            req_hist_cols = ['患者oneId', '药品名称', '门店', '门店编码']
                            if all(col in df_history.columns for col in req_hist_cols):
                                df_history['门店'] = df_history['门店'].astype(str).str.strip().replace(PHARMACY_MAPPING)
                                df_history['门店编码'] = df_history['门店编码'].astype(str).str.strip().replace(CODE_MAPPING)
                                df_history['干净药品名'] = df_history['药品名称'].astype(str).str.split('-').str[0].str.strip()
                                df_history['患者oneId'] = df_history['患者oneId'].astype(str).str.strip().str.lstrip(',')
                                
                                df_history['4D_KEY'] = (
                                    df_history['干净药品名'] + "_" + 
                                    df_history['门店'] + "_" + 
                                    df_history['门店编码'] + "_" + 
                                    df_history['患者oneId']
                                )
                                
                                result_df['4D_KEY'] = (
                                    result_df['商品名称'].astype(str).str.strip() + "_" + 
                                    result_df['门店名称'].astype(str).str.strip() + "_" + 
                                    result_df['门店code'].astype(str).str.strip() + "_" + 
                                    result_df['会员卡号'].astype(str).str.strip()
                                )
                                
                                result_df = result_df[~result_df['4D_KEY'].isin(df_history['4D_KEY'])]
                                result_df = result_df.drop(columns=['4D_KEY'])
                                has_history = True
                            else:
                                hist_warning_flag = True
                        except Exception:
                            hist_warning_flag = True

            # ------------------ 计算结束，开始渲染结果 ------------------
            if followup_patients.empty:
                st.info(f"✨ 计算完成：{selected_month_display} 没有任何需要随访的老患者。")
            else:
                if hist_warning_flag:
                    st.warning("⚠️ 提示：上传的已随访大表中缺少必要的列（须包含：患者oneId、药品名称、门店、门店编码），已跳过差额比对。")
                
                if result_df.empty:
                    st.success(f"✨ 智能比对完成：{selected_month_display} 所有老患者的随访任务已全部达成，没有漏访！")
                else:
                    st.toast("🎉 随访名单计算完成！", icon="✅")
                    st.divider()
                    
                    st.markdown("### 🔍 随访数据多维快速筛选")
                    col1, col2 = st.columns(2)
                    with col1:
                        unique_products = sorted(result_df['商品名称'].unique().tolist())
                        selected_products = st.multiselect("📦 筛选商品品种 (不选默认全选)：", options=unique_products, default=[])
                    with col2:
                        unique_pharmacies = sorted(result_df['门店名称'].unique().tolist())
                        selected_pharmacies = st.multiselect("🏪 筛选药房门店 (不选默认全选)：", options=unique_pharmacies, default=[])
                    
                    display_df = result_df.copy()
                    if selected_products:
                        display_df = display_df[display_df['商品名称'].isin(selected_products)]
                    if selected_pharmacies:
                        display_df = display_df[display_df['门店名称'].isin(selected_pharmacies)]
                    
                    metric_label = f"🚨 {selected_month_display} 需【补随访】差额" if has_history else f"🎉 {selected_month_display} 需随访老患者总数"
                    if selected_products or selected_pharmacies:
                        metric_label += " (已应用筛选条件)"
                    st.metric(label=metric_label, value=f"{len(display_df)} 条任务")
                    
                    # 创建空的标准框架表格（严格遵循 TARGET_EXCEL_HEADERS 顺序）
                    export_final_df = pd.DataFrame(columns=TARGET_EXCEL_HEADERS)
                    
                    # 基础对齐映射关系（这里恢复对商品规格的直接映射映射）
                    MAPPING_DICTIONARY = {
                        '门店编码': '门店code',
                        '门店名称': '门店名称',
                        '会员卡号或患者ID': '会员卡号',
                        '商品名': '商品名称',
                        '商品规格': '规格'
                    }
                    
                    # 🔥【核心：后台双重容错对齐算法逻辑 —— 仅识别，不改写规格文字】
                    id_codes, chemicals, forms = [], [], []
                    for _, row in display_df.iterrows():
                        p_name = str(row['商品名称']).strip()
                        # 后台匹配时自动规范化，兼容“粒/片”以及中英文括号差异
                        p_spec = str(row['规格']).strip().lower().replace(" ", "").replace("粒", "片").replace("（", "(").replace("）", ")")
                        
                        # 第一步：按商品名称筛选
                        matches = [item for item in PRODUCT_MASTER_LIST if item['name'] == p_name]
                        
                        if not matches:
                            id_codes.append(""); chemicals.append(""); forms.append("")
                        elif len(matches) == 1:
                            id_codes.append(matches[0]['id_code'])
                            chemicals.append(matches[0]['chemical'])
                            forms.append(matches[0]['form'])
                        else:
                            # 第二步：多规格情况，通过比对剂量强度关键字符来判定
                            matched_item = None
                            for m in matches:
                                m_spec = m['spec'].lower().replace(" ", "").replace("粒", "片").replace("（", "(").replace("）", ")")
                                if m_spec in p_spec or p_spec in m_spec:
                                    matched_item = m
                                    break
                            
                            if not matched_item:
                                for m in matches:
                                    m_strength = m['spec'].split('/')[0].split('*')[0].lower().strip().replace("（", "(").replace("）", ")")
                                    if m_strength in p_spec:
                                        matched_item = m
                                        break
                                        
                            if not matched_item:
                                matched_item = matches[0]
                                
                            id_codes.append(matched_item['id_code'])
                            chemicals.append(matched_item['chemical'])
                            forms.append(matched_item['form'])
                    
                    # 严格依序注入最终要输出的 DataFrame
                    for col_name in TARGET_EXCEL_HEADERS:
                        if col_name in MAPPING_DICTIONARY:
                            export_final_df[col_name] = display_df[MAPPING_DICTIONARY[col_name]]
                        elif col_name == '商品ID编码':
                            export_final_df['商品ID编码'] = id_codes
                        elif col_name == '化学名':
                            export_final_df['化学名'] = chemicals
                        elif col_name == '剂型':
                            export_final_df['剂型'] = forms
                        else:
                            export_final_df[col_name] = "" 
                    
                    st.markdown("### 📄 标准随访表预览 (前100条 - 商品规格已还原为原始状态)")
                    st.dataframe(export_final_df.head(100), use_container_width=True)
                    
                    # 导出处理
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        export_final_df.to_excel(writer, index=False, sheet_name='随访名单')
                        workbook = writer.book
                        worksheet = writer.sheets['随访名单']
                        
                        # 🔒 锁定数字长文本，防止Excel变形
                        text_format_cols = ['会员卡号或患者ID', '商品ID编码', '门店编码']
                        for target_col in text_format_cols:
                            if target_col in export_final_df.columns:
                                col_idx = export_final_df.columns.get_loc(target_col) + 1
                                for row in range(2, worksheet.max_row + 1):
                                    cell = worksheet.cell(row=row, column=col_idx)
                                    cell.number_format = '@'
                                
                    processed_data = output.getvalue()
                    file_suffix = "_部分筛选" if selected_products or selected_pharmacies else ""
                    
                    st.download_button(
                        label="📥 点击下载【规范格式】随访 Excel 表",
                        data=processed_data,
                        file_name=f"{target_month_str}月份标准随访记录表{file_suffix}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    except Exception as e:
        st.error(f"💥 程序运行出错。错误原因: {e}")
