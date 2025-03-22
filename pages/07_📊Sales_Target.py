import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# 设置工作目录并读取数据
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")
# 读取第一个数据源
df = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='raw_sheet',
    skiprows=0,
    usecols='A:AU',
    nrows=100000
).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query(
    'FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query(
    'FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query(
    'Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query(
    'Inv_Month != "Cancel"')

# 读取第二个数据源
df2 = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='SMT_Internal_Sales_Target',
    skiprows=0,
    usecols='A:G',
    nrows=100000)

# 数据预处理
df = df.dropna(subset=['Inv_Yr', 'Inv_Month'])
df['Region'] = df['Region'].replace(['', 'C66 N/A'], pd.NA).dropna()

# 设置页面布局
st.set_page_config(layout="wide")

# ========== 左边栏：Filter设置 ==========
with st.sidebar:
    st.header("Monthly Report")
    
    # Filter设置
    selected_fy = st.selectbox(
        'FY',
        options=df['FY_INV'].unique(),
        index=list(df['FY_INV'].unique()).index('FY 24/25') if 'FY 24/25' in df['FY_INV'].unique() else 0
    )
    selected_yr = st.multiselect('Inv Yr', options=df['Inv_Yr'].unique())
    selected_month = st.multiselect('Inv Month', options=df['Inv_Month'].unique())
    region_options = [r for r in df['Region'].unique() if pd.notna(r)]
    selected_region = st.multiselect('Region', options=region_options)

    st.header("Sales Target_FY")
    selected_target_fy = st.selectbox(
        'FY_INV (Target)',
        options=df2['FY_INV'].unique(),
        index=list(df2['FY_INV'].unique()).index('FY 24/25') if 'FY 24/25' in df2['FY_INV'].unique() else 0
    )

