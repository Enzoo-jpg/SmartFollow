import streamlit as st
import pandas as pd
import io

# 设置网页标题和图标
st.set_page_config(page_title="患者随访名单筛选系统", page_icon="📋", layout="centered")

st.title("📋 患者随访名单自动筛选系统")
st.markdown("""
<style>
    .big-font { font-size:16px !important; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)

# 侧边栏或主页使用说明
st.markdown("""
### 💡 使用前须知：
1. 上传的 Excel 中必须包含一个名为 **`销售底表`** 的工作表。
2. 该工作表表头必须包含这五列：**`销售时间`**、**`商品名`**、**`药房`**、**`患者id`**（或患者ID）、**`规格`**。
3. **数据安全**：Streamlit 仅在内存中即时处理数据，绝不会存储您的任何患者隐私数据。
""", unsafe_allow_html=True)

st.divider()

# 1. 用户选择随访月份
target_month = st.date_input("1. 选择需要随访的月份：", value=pd.to_datetime("2026-05-01"), min_value=pd.to_datetime("2025-10-01"))
target_month_str = target_month.strftime("%Y-%m")

# 2. 用户上传文件
uploaded_file = st.file_uploader("2. 上传历史销售底表 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 读取Excel中的“销售底表”
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = [sheet for sheet in xl.sheet_names if "销售底表" in sheet]
        
        if not sheet_name:
            st.error("❌ 错误：未在 Excel 中找到包含“销售底表”字样的工作表，请检查工作表名称！")
        else:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name[0])
            
            # 标准化列名，防止空格影响
            df.columns = df.columns.str.strip()
            
            # 自动识别关键列
            time_col = next((col for col in df.columns if "时间" in col), None)
            name_col = next((col for col in df.columns if "商品名" in col), None)
            pharmacy_col = next((col for col in df.columns if "药房" in col), None)
            id_col = next((col for col in df.columns if "id" in col.lower() or "患者" in col), None)
            spec_col = next((col for col in df.columns if "规格" in col), None)
            
            if not all([time_col, name_col, pharmacy_col, id_col, spec_col]):
                st.error("❌ 错误：底表中缺少必要列！请确保包含：销售时间、商品名、药房、患者id、规格。")
            else:
                st.success("📊 成功读取底表数据，正在计算，请稍候...")
                
                # 转换为标准日期格式
                df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
                df = df.dropna(subset=[time_col]) # 剔除无时间记录的行
                
                # 生成购买月份列
                df['购买月份'] = df[time_col].dt.strftime('%Y-%m')
                
                # 拼接患者唯一KEY
                df['患者唯一KEY'] = df[name_col].astype(str) + "_" + df[pharmacy_col].astype(str) + "_" + df[id_col].astype(str) + "_" + df[spec_col].astype(str)
                
                # 计算每个患者的绝对首次购买月份
                first_buy = df.groupby('患者唯一KEY')['购买月份'].min().reset_index()
                first_buy.columns = ['患者唯一KEY', '首次购买月份']
                
                # 核心业务逻辑筛选：首次购买月份 < 当前选择的随访月份
                followup_patients = first_buy[first_buy['首次购买月份'] < target_month_str]
                
                if followup_patients.empty:
                    st.info(f"✨ 计算完成：{target_month_str} 月份没有任何需要随访的老患者（所有人在此月皆为新患或尚未开户）。")
                else:
                    # 还原字段
                    # 重新拆分唯一KEY回原来的四个字段
                    split_cols = followup_patients['患者唯一KEY'].str.split('_', expand=True)
                    result_df = pd.DataFrame({
                        '患者唯一KEY': followup_patients['患者唯一KEY'],
                        '商品名': split_cols[0],
                        '药房': split_cols[1],
                        '患者id': split_cols[2],
                        '规格': split_cols[3]
                    })
                    
                    st.divider()
                    st.metric(label=f"🎉 {target_month_str} 月需随访老患者总数", value=f"{len(result_df)} 人")
                    
                    # 预览数据
                    st.markdown("### 📄 随访名单预览 (前100条)")
                    st.dataframe(result_df.head(100), use_container_width=True)
                    
                    # 导出为 Excel 内存流
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name=f'{target_month_str}随访名单')
                    processed_data = output.getvalue()
                    
                    # 下载按钮
                    st.download_button(
                        label="📥 点击下载完整随访 Excel 表",
                        data=processed_data,
                        file_name=f"{target_month_str}月份自动随访名单.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    except Exception as e:
        st.error(f"💥 程序运行出错，请确保数据格式正确。错误原因: {e}")