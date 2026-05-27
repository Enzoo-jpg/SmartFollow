import streamlit as st
import pandas as pd
import io

# 设置网页标题和图标
st.set_page_config(page_title="患者随访名单筛选系统", page_icon="📋", layout="centered")

st.title("📋 患者随访名单自动筛选系统")

# =================【这里是你的药房名字映射配置区】=================
# 💡 小白提示：左边写你底表里“五花八门”的原始药房名，右边写你最终想要的“标准”药房名
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
    "国药控股内江有限公司第二大药房": "国药控股内江有限公司 second大药房",
    "四川环晟大药房有限公司蜀南大道店": "四川环晟大药房有限公司蜀南大道店",
    "国药控股(乐山)川药医药有限公司滨河路店": "国药控股（乐山）川药医药有限公司乐山高新区店",
    "国药控股昊阳绵阳药业有限公司江油匡山路大药房": "国药控股昊阳绵阳药业有限公司江油匡山路大药房",
    "DTC国药控股四川专业药房连锁有限公司攀枝花药房": "国药控股四川专业药房连锁有限公司攀枝花药房",
    "DTC国药控股广安有限公司广安药房": "国药控股广安有限公司广安药房",
    "DTC国药控股四川医药股份有限公司泸州药房": "国药控股四川医药股份有限公司泸州药房",
    "DTC国药控股四川专业药房连锁有限公司资阳药房": "国药控股四川专业药房连锁有限公司资阳药房",
}

# =================【这里是你的门店编码映射配置区】=================
# 💡 小白提示：左边写你底表里不规范或旧的编码，右边写你最终想要的“标准”新编码
CODE_MAPPING = {
    "DM0900444": "9025007",   # 示例：把旧编码映射为新编码
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
    "DM1121464": "001",  # 四川环晟大药房有限公司蜀南大道店，需要后期更改
    "DM0802780": "9027001",  
    "DM1074241": "9030001",  
    "DM0606859": "9025012",  
    "DM0900441": "9011001",  
    "DM0641557": "9007001",  
    "DM1012587": "9025006",  
    # 以后如果有新增的编码更名需求，按照格式一行行加在这里即可
}
# =============================================================

st.markdown("""
### 💡 使用前须知：
1. 上传的 Excel 中必须包含一个名为 **`销售底表`** 的工作表。
2. 该工作表表头必须包含这六列：**`销售时间`**、**`商品名`**、**`药房`**、**`门店编码`**、**`患者id`**、**`规格`**。
3. **自动更名映射**：系统已内置药房名称和门店编码转换功能，上传后会自动进行规范化统一。
""", unsafe_allow_html=True)

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
                st.success("📊 成功读取数据，正在进行规范化转换及随访计算...")
                
                # 自动做药房名字映射
                df[pharmacy_col] = df[pharmacy_col].astype(str).str.strip()
                df[pharmacy_col] = df[pharmacy_col].replace(PHARMACY_MAPPING)
                
                # 🛠️ 优化点1：确保门店编码被当做文本处理，并执行编码映射替换
                df[code_col] = df[code_col].astype(str).str.strip() 
                df[code_col] = df[code_col].replace(CODE_MAPPING)
                
                # 转换为标准日期格式
                df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
                df = df.dropna(subset=[time_col])
                
                df['购买月份'] = df[time_col].dt.strftime('%Y-%m')
                
                # 5 字段拼接组合成患者唯一KEY
                df['患者唯一KEY'] = (
                    df[name_col].astype(str) + "_" + 
                    df[pharmacy_col].astype(str) + "_" + 
                    df[code_col].astype(str) + "_" + 
                    df[id_col].astype(str) + "_" + 
                    df[spec_col].astype(str)
                )
                
                # 计算每个患者的绝对首次购买月份
                first_buy = df.groupby('患者唯一KEY')['购买月份'].min().reset_index()
                first_buy.columns = ['患者唯一KEY', '首次购买月份']
                
                # 核心业务逻辑筛选
                followup_patients = first_buy[first_buy['首次购买月份'] < target_month_str]
                
                if followup_patients.empty:
                    st.info(f"✨ 计算完成：{selected_month_display} 没有任何需要随访的老患者。")
                else:
                    split_cols = followup_patients['患者唯一KEY'].str.split('_', expand=True)
                    
                    # 🛠️ 优化点2：重写导出结构，将“门店编码”完美还原并加入结果表格
                    result_df = pd.DataFrame({
                        '患者唯一KEY': followup_patients['患者唯一KEY'],
                        '商品名': split_cols[0],
                        '药房': split_cols[1],
                        '门店编码': split_cols[2],  # 👈 门店编码被正确提取出来了
                        '患者id': split_cols[3],    # 👈 位置顺延到3
                        '规格': split_cols[4]      # 👈 位置顺延到4
                    })
                    
                    st.divider()
                    st.metric(label=f"🎉 {selected_month_display} 需随访老患者总数", value=f"{len(result_df)} 人")
                    
                    st.markdown("### 📄 随访名单预览 (前100条)")
                    st.dataframe(result_df.head(100), use_container_width=True)
                    
                   # 导出为 Excel 内存流
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
