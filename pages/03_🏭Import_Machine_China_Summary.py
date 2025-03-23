import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly import subplots
import plotly.colors as colors
from streamlit_extras.metric_cards import style_metric_cards

# 網頁基本設定
st.set_page_config(page_title="Sales Dashboard", page_icon=":rainbow:", layout="wide")
st.title(':factory:  Mounter Import Data of China_Analysis')
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)
st.write("by Arthur Chan")

# 資料載入與清洗
@st.cache_data
def load_clean_data():
    # 載入主要資料集
    df = pd.read_excel(
        io='Monthly_report_for_edit.xlsm',
        engine='openpyxl',
        sheet_name='raw_sheet',
        usecols='A:AR',
        dtype={'Inv_Yr': str}
    ).query('Region != "C66 N/A" and FY_Contract != "Cancel" and FY_INV not in ["TBA", "FY 17/18"]')
    
    # 嚴格處理Inv_Yr型別
    df['Inv_Yr'] = pd.to_numeric(df['Inv_Yr'], errors='coerce')
    df = df.dropna(subset=['Inv_Yr'])
    df['Inv_Yr'] = df['Inv_Yr'].astype(int)
    
    # 載入進口資料
    df_import = pd.read_excel(
        io='Machine_Import_data.xlsm',
        sheet_name='raw data',
        usecols='A:G',
        dtype={'YEAR': str}
    )
    df_import = df_import[df_import['YEAR'].str.isnumeric()]
    df_import['YEAR'] = df_import['YEAR'].astype(int)
    
    return df, df_import

df, df_import = load_clean_data()

# 側邊欄篩選器
st.sidebar.divider()
st.sidebar.header(":point_down: Filters")

# 進口年份範圍篩選 (預設全選)
valid_years_import = sorted(df_import['YEAR'].unique())
selected_years = st.sidebar.multiselect(
    "選擇進口年份範圍",
    options=valid_years_import,
    default=valid_years_import
)

# 發票年份篩選 (動態跟隨進口年份)
valid_years_inv = sorted(df['Inv_Yr'].unique())
selected_inv_years = st.sidebar.multiselect(
    "選擇發票年份",
    options=valid_years_inv,
    default=selected_years
)

# 其他篩選器
region = st.sidebar.multiselect("選擇地區", df["Region"].unique())

# 資料篩選
filter_import = df_import[df_import['YEAR'].isin(selected_years)]
filter_smt = df[
    (df['Inv_Yr'].isin(selected_inv_years)) &
    (df['BRAND'].isin(['HELLER', 'PEMTRON', 'YAMAHA']))
]
if region: filter_smt = filter_smt[filter_smt["Region"].isin(region)]

# 主視覺化佈局
col1, col2 = st.columns(2)

with col1:
    if not filter_import.empty:
        df_import_group = filter_import.groupby(['YEAR', 'MONTH']).agg({'台数':'sum', '进口金额（人民币）':'sum'}).reset_index()
        
        fig = subplots.make_subplots(specs=[[{"secondary_y": True}]])
        
        years = sorted(df_import_group['YEAR'].unique())
        color_map = {year: 'orange' if year == max(years) else colors.qualitative.Dark24[i] for i, year in enumerate(years)}
        
        for year in years:
            df_year = df_import_group[df_import_group['YEAR'] == year]
            
            fig.add_trace(go.Bar(
                x=df_year['MONTH'],
                y=df_year['进口金额（人民币）'] / 1e6,
                name=f'{year} 金額',
                marker_color=color_map[year],
                opacity=0.6
            ), secondary_y=False)
            
            fig.add_trace(go.Scatter(
                x=df_year['MONTH'],
                y=df_year['台数'],
                mode='lines+markers+text',
                name=f'{year} 台數',
                line=dict(color=color_map[year], width=3),
                marker=dict(symbol='square', size=10),
                text=df_year['台数'],
                textposition='middle right',
                textfont=dict(color='black', size=15)
            ), secondary_y=True)
        
        fig.update_layout(
            height=600,
            title={
                'text': "📈 China Mounter Import Trend(QTY & CNY Amount)",
                'font': {'size': 24}
            },
            xaxis=dict(
                title='月份',
                tickmode='linear',
                dtick=1,
                gridcolor='black',
                gridwidth=1,
                showgrid=True
            ),
            yaxis=dict(
                title='金額 (百萬人民幣)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                tickformat='.0fM',
                showgrid=True
            ),
            yaxis2=dict(
                title='數量 (台)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                showgrid=True,
                rangemode='tozero'
            ),
            plot_bgcolor='rgba(255,255,255,0.9)',
            paper_bgcolor='rgb(240,240,240)',
            hovermode='x unified',
            showlegend=True
        )
        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("進口數據樞紐分析表", expanded=True):
            pivot_import = filter_import.pivot_table(
                values=["进口金额（人民币）", "台数"],
                index="MONTH",
                columns="YEAR",
                aggfunc="sum",
                margins=True,
                margins_name="總計"
            ).fillna(0)
            
            html = pivot_import.applymap(lambda x: f"{x:,.0f}").to_html(classes='table table-bordered')
            html = html.replace('<th>台数</th>', '<th style="background-color: #90EE90">台数</th>')
            html = html.replace('<th>进口金额（人民币）</th>', '<th style="background-color: #FFA500">进口金额（人民币）</th>')
            st.markdown(f'<div style="zoom:1.1">{html}</div>', unsafe_allow_html=True)
            csv = pivot_import.to_csv(float_format='%.0f').encode('utf-8')
            st.download_button("下載進口數據", csv, "china_import.csv", "text/csv")

with col2:
    if not filter_smt.empty:
        df_smt_group = filter_smt.groupby(['Inv_Yr', 'Inv_Month']).agg({'Item Qty':'sum', 'Before tax Inv Amt (HKD)':'sum'}).reset_index()
        
        fig = subplots.make_subplots(specs=[[{"secondary_y": True}]])
        
        years = sorted(df_smt_group['Inv_Yr'].unique())
        color_map = {year: 'orange' if year == max(years) else colors.qualitative.Plotly[i] for i, year in enumerate(years)}
        
        for year in years:
            df_year = df_smt_group[df_smt_group['Inv_Yr'] == year]
            
            fig.add_trace(go.Bar(
                x=df_year['Inv_Month'],
                y=df_year['Before tax Inv Amt (HKD)'] / 1e6,
                name=f'{year} 金額',
                marker_color=color_map[year],
                opacity=0.6
            ), secondary_y=False)
            
            fig.add_trace(go.Scatter(
                x=df_year['Inv_Month'],
                y=df_year['Item Qty'],
                mode='lines+markers+text',
                name=f'{year} 數量',
                line=dict(color=color_map[year], width=3),
                marker=dict(symbol='square', size=10),
                text=df_year['Item Qty'],
                textposition='middle right',
                textfont=dict(color='black', size=15)
            ), secondary_y=True)
        
        fig.update_layout(
            height=600,
            title={
                'text': "📊 SMT Invoice Trend(YAMAHA/HELLER/PEMTRON QTY & HKD Amount)",
                'font': {'size': 24}
            },
            xaxis=dict(
                title='月份',
                tickmode='linear',
                dtick=1,
                gridcolor='black',
                gridwidth=1,
                showgrid=True
            ),
            yaxis=dict(
                title='金額 (百萬港幣)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                tickformat='.0fM',
                showgrid=True
            ),
            yaxis2=dict(
                title='數量 (件)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                showgrid=True,
                rangemode='tozero'
            ),
            plot_bgcolor='rgba(255,255,255,0.9)',
            paper_bgcolor='rgb(240,240,240)',
            hovermode='x unified',
            showlegend=True
        )
        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("發票數據樞紐分析表", expanded=True):
            # 調整樞紐表結構
            pivot_smt = filter_smt.pivot_table(
                values=["Before tax Inv Amt (HKD)", "Item Qty"],
                index="BRAND",
                columns=["Inv_Yr", "Inv_Month"],
                aggfunc="sum",
                margins=True
            ).fillna(0)
            
            # 按品牌字母降序排序
            pivot_smt = pivot_smt.reindex(sorted(pivot_smt.index, reverse=True))
            
            html = pivot_smt.style.format("{:,.0f}").to_html()
            html = html.replace('<th>HELLER</th>', '<th style="background-color: #FFA500">HELLER</th>')
            html = html.replace('<th>PEMTRON</th>', '<th style="background-color: #ADD8E6">PEMTRON</th>')
            html = html.replace('<th>YAMAHA</th>', '<th style="background-color: #FFB6C1">YAMAHA</th>')
            st.markdown(f'<div style="zoom:1.1">{html}</div>', unsafe_allow_html=True)
            csv = pivot_smt.to_csv(float_format='%.0f').encode('utf-8')
            st.download_button("下載發票數據", csv, "smt_invoice.csv", "text/csv")

# 樣式設定
style_metric_cards(background_color="#FFFFFF", border_left_color="#686664")

# 頁尾
st.markdown("""
<div style="text-align:center;padding:1rem;margin-top:2rem;border-top:2px solid #ddd">
    <p style="color:#666">Developed by Arthur Chan • Data Version: 2024-02</p>
</div>
""", unsafe_allow_html=True)
