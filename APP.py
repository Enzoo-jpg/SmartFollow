import streamlit as st
import pandas as pd
import io

# 设置网页标题和图标
st.set_page_config(page_title="患者随访名单筛选系统", page_icon="📋", layout="centered")

st.title("📋 患者随访名单自动筛选系统")

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
    "DM1121464": "001",  
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
1. **历史销售底表**：上传的 Excel 中必须包含一个名为 **`销售底表`** 的工作表，且表头须含：*销售时间、商品名、药房、门店编码、患者id、规格*。
2. **已完成随访记录表（可选）**：上传后系统会自动抓取其中的 *患者oneId、药品名称、门店、门店编码* 列，并**智能按次数差额扣减**，精准提取出**“漏访待补名单”**。
3. **关于规格说明**：目前系统已升级，**规格不再作为判定新老患者的依据**（即同一患者在同门店买同药品，多规格不重复计算随访）。
4. **自动更名映射**：系统已内置药房名称和门店编码转换功能，上传后会自动进行规范化统一。
""", unsafe_allow_html=True)

st.divider()

# =================【标准模板下载区】=================
st.markdown("### 📥 官方标准模板下载")
with st.expander("👉 如果忘记表头或怕格式有误，请点此展开下载标准模板"):
    df_base_tpl = pd.DataFrame(columns=['销售时间', '商品名', '药房', '门店编码', '患者id', '规格'])
    df_base_tpl.loc[0] = ['2026-05-01', '特诺雅', '国药控股四川专业药房连锁有限公司金牛区一环路西三段药房', '9025007', 'HZ001', '100mg']
    base_out = io.BytesIO()
    with pd.ExcelWriter(base_out, engine='openpyxl') as writer:
        df_base_tpl.to_excel(writer, index=False, sheet_name='销售底表')
    base_tpl_bytes = base_out.getvalue()

    df_hist_tpl = pd.DataFrame(columns=['患者oneId', '药品名称', '门店', '门店编码'])
    df_hist_tpl.loc[0] = ['HZ001', '特诺雅-(古塞奇尤单抗注射液)', '国药控股四川专业药房连锁有限公司金牛区一环路西三段药房', '9025007']
    hist_out = io.BytesIO()
    with pd.ExcelWriter(hist_out, engine='openpyxl') as writer:
        df_hist_tpl.to_excel(writer, index=False, sheet_name='已完成随访记录')
    hist_tpl_bytes = hist_out.getvalue()

    # 优化点：缩短 label 长度，严防复制截断错误
    st.download_button(
        label="📥 下载《1. 销售底表模板》",
        data=base_tpl_bytes,
        file_name="1_销售底表标准模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label="📥 下载《2. 已完成随访模板》",
        data=hist_tpl_bytes,
        file_name="2_已完成随访记录表模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
uploaded_history = st.file_uploader("3. （可选）上传本月已完成的随访记录表 (.xlsx) —— 用于多规格差额补访比对", type=["xlsx"])

if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = [sheet for sheet in xl.sheet_names if "销售底表" in sheet]
        
        if not sheet_name:
            st.error("❌ 错误：未在 Excel 中找到包含“销售底表”字样的工作表！")
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name[0])
            df.columns = df.columns.str.strip()
            
            # 自动识别关键列
            time_col = next((col for col in df.columns if "时间" in col), None)
            name_col = next((col for col in df.columns if "商品名" in col), None)
            pharmacy_col = next((col for col in df.columns if "药房" in col), None)
            code_col = next((col for col in df.columns if "编码" in col), None) 
            id_col = next((col for col in df.columns if "id" in col.lower() or "患者" in col), None)
            spec_col = next((col for col in df.columns if "规格" in col), None)
            
            if not all([time_col, name_col, pharmacy_col, code_col, id_col, spec_col]): 
                st.error("❌ 错误：底表中缺少必要列！请确保包含：销售时间、商品名、药房、门店编码、患者id、规格。")
            else:
                # 初始化核心控制变量
                result_df = pd.DataFrame()
                followup_patients = pd.DataFrame()
                has_history = False
                hist_warning_flag = False
                
                # 把繁重的计算过程包裹在 spinner 中
                with st.spinner("📊 正在读取数据并进行智能随访计算..."):
                    # 自动做药房名字映射
                    df[pharmacy_col] = df[pharmacy_col].astype(str).str.strip().replace(PHARMACY_MAPPING)
                    # 门店编码替换
                    df[code_col] = df[code_col].astype(str).str.strip().replace(CODE_MAPPING)
                    
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
                            '患者唯一KEY': followup_patients['患者唯一KEY'],
                            '商品名': followup_patients[name_col],
                            '药房': followup_patients[pharmacy_col],
                            '门店编码': followup_patients[code_col],  
                            '患者id': followup_patients[id_col],    
                            '规格': followup_patients[spec_col]      
                        })
                        
                        # 比对已完成随访表
                        if uploaded_history is not None:
                            try:
                                df_history = pd.read_excel(uploaded_history)
                                df_history.columns = df_history.columns.str.strip()
                                
                                req_hist_cols = ['患者oneId', '药品名称', '门店', '门店编码']
                                if all(col in df_history.columns for col in req_hist_cols):
                                    df_history['干净药品名'] = df_history['药品名称'].astype(str).str.split('-').str[0].str.strip()
                                    df_history['患者oneId'] = df_history['患者oneId'].astype(str).str.strip()
                                    df_history['门店'] = df_history['门店'].astype(str).str.strip()
                                    df_history['门店编码'] = df_history['门店编码'].astype(str).str.strip()
                                    
                                    df_history['4D_KEY'] = (
                                        df_history['干净药品名'] + "_" + 
                                        df_history['门店'] + "_" + 
                                        df_history['门店编码'] + "_" + 
                                        df_history['患者oneId']
                                    )
                                    df_history['seq'] = df_history.groupby('4D_KEY').cumcount().astype(str)
                                    df_history['match_KEY'] = df_history['4D_KEY'] + "_" + df_history['seq']
                                    
                                    result_df['4D_KEY'] = (
                                        result_df['商品名'].astype(str).str.strip() + "_" + 
                                        result_df['药房'].astype(str).str.strip() + "_" + 
                                        result_df['门店编码'].astype(str).str.strip() + "_" + 
                                        result_df['患者id'].astype(str).str.strip()
                                    )
                                    result_df['seq'] = result_df.groupby('4D_KEY').cumcount().astype(str)
                                    result_df['match_KEY'] = result_df['4D_KEY'] + "_" + result_df['seq']
                                    
                                    result_df = result_df[~result_df['match_KEY'].isin(df_history['match_KEY'])]
                                    result_df = result_df.drop(columns=['4D_KEY', 'seq', 'match_KEY'])
                                    has_history = True
                                else:
                                    hist_warning_flag = True
                            except Exception as hist_e:
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
                        
                        metric_label = f"🚨 {selected_month_display} 需【补随访】老患者差额" if has_history else f"🎉 {selected_month_display} 需随访老患者总数"
                        st.metric(label=metric_label, value=f"{len(result_df)} 条任务")
                        
                        st.markdown("### 📄 随访名单预览 (前100条)")
                        st.dataframe(result_df.head(100), use_container_width=True)
                        
                        # 导出流
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='随访名单')
                        processed_data = output.getvalue()
                        
                        st.download_button(
                            label="📥 点击下载完整随访 Excel 表",
                            data=processed_data,
                            file_name=f"{target_month_str}月份自动随访名单.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
    except Exception as e:
        st.error(f"💥 程序运行出错。错误原因: {e}")
