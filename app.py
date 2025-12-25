import streamlit as st
import pandas as pd
import numpy as np
import statsmodels.formula.api as smf
import io

st.set_page_config(page_title="销售话术指标分析工具", layout="wide")

st.title("📊 销售话术指标 (LMM) 自动分析平台")
st.markdown("""
该工具将基于您上传的数据，自动计算各产品的 **ICC (组内相关系数)**、**相关性** 以及 **线性混合模型 (LMM)**。
""")

# --- 1. 上传文件 ---
uploaded_file = st.file_uploader("请上传您的 Excel 数据文件", type=["xlsx", "csv"])

if uploaded_file:
    # 加载数据
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.write("✅ 数据预览：", df.head(5))

    # --- 2. 参数配置 ---
    st.sidebar.header("分析配置")
    all_cols = df.columns.tolist()
    
    store_col = st.sidebar.selectbox("选择门店ID列 (Group Var)", all_cols)
    metrics = st.sidebar.multiselect("选择6个话术指标 (Metrics)", all_cols, default=all_cols[:6] if len(all_cols)>6 else None)
    products = st.sidebar.multiselect("选择要分析的业绩列 (Products)", all_cols)

    if st.sidebar.button("开始批量分析"):
        if not metrics or not products or not store_col:
            st.error("请确保已选择门店ID、话术指标和至少一个业绩列。")
        else:
            all_corr_list = []
            all_lmm_list = []
            icc_report = []

            progress_bar = st.progress(0)
            
            for idx, prod in enumerate(products):
                # A. 相关性
                correlations = df[metrics + [prod]].corr()[prod].drop(prod)
                corr_df = correlations.to_frame(name='Correlation').reset_index()
                corr_df.columns = ['Metric', 'Correlation']
                corr_df['Product'] = prod
                all_corr_list.append(corr_df)
                
                # B. ICC
                try:
                    null_model = smf.mixedlm(f"Q('{prod}') ~ 1", df, groups=df[store_col]).fit()
                    sigma_between = null_model.cov_re.iloc[0, 0]
                    sigma_within = null_model.scale
                    icc_value = sigma_between / (sigma_between + sigma_within)
                    icc_report.append({'Product': prod, 'ICC': icc_value})
                except:
                    st.warning(f"产品 {prod} ICC 计算失败")

                # C. LMM
                formula = f"Q('{prod}') ~ " + " + ".join([f"Q('{m}')" for m in metrics])
                try:
                    lmm_model = smf.mixedlm(formula, df, groups=df[store_col]).fit()
                    summary_table = lmm_model.summary().tables[1].reset_index()
                    summary_table.columns = ['Metric', 'Coef', 'Std.Err', 'z', 'P_value', '[0.025', '0.975]']
                    lmm_res = summary_table[summary_table['Metric'].str.contains('|'.join(metrics))].copy()
                    lmm_res['Product'] = prod
                    all_lmm_list.append(lmm_res)
                except Exception as e:
                    st.error(f"产品 {prod} LMM 拟合失败: {e}")
                
                progress_bar.progress((idx + 1) / len(products))

            # --- 3. 准备下载文件 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                pd.DataFrame(icc_report).to_excel(writer, sheet_name='ICC_Overview', index=False)
                pd.concat(all_corr_list).pivot(index='Metric', columns='Product', values='Correlation').to_excel(writer, sheet_name='All_Correlations')
                pd.concat(all_lmm_list).to_excel(writer, sheet_name='LMM_Full_Details', index=False)
            
            st.success("🎉 分析完成！")
            st.download_button(
                label="📥 点击下载分析报告",
                data=output.getvalue(),
                file_name="Sales_Analysis_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )