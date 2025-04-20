import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly import subplots
import plotly.colors as colors
from streamlit_extras.metric_cards import style_metric_cards
import os
from io import BytesIO

def format_numbers(x):
    """Custom number formatter with K and M abbreviations"""
    try:
        num = float(x)
        if num >= 1_000_000:
            return f"{num/1_000_000:,.0f}M"
        elif num >= 10_000:  # 修改為1萬以上才用K顯示
            return f"{num/1_000:,.0f}K"
        return f"{num:,.0f}"  # 1萬以下直接顯示原始數字
    except:
        return str(x)

# 網頁基本設定
st.set_page_config(page_title="Sales Dashboard", page_icon=":rainbow:", layout="wide")
st.title(':factory:  Mounter Import Data of China_Analysis')
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)
st.write("by Arthur Chan")
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")

# 資料載入與清洗
#@st.cache_data
def load_clean_data():
    df = pd.read_excel(
        io='Monthly_report_for_edit.xlsm',
        engine='openpyxl',
        sheet_name='raw_sheet',
        usecols='A:AR',
        dtype={'Inv_Yr': str}
    ).query('Region != "C66 N/A" and FY_Contract != "Cancel" and FY_INV not in ["TBA", "FY 17/18"]')
    
    df['Inv_Yr'] = pd.to_numeric(df['Inv_Yr'], errors='coerce')
    df = df.dropna(subset=['Inv_Yr'])
    df['Inv_Yr'] = df['Inv_Yr'].astype(int)
    
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

valid_years_import = sorted(df_import['YEAR'].unique())
selected_years = st.sidebar.multiselect(
    "選擇進口年份範圍",
    options=valid_years_import,
    default=valid_years_import
)

valid_years_inv = sorted(df['Inv_Yr'].unique())
selected_inv_years = st.sidebar.multiselect(
    "選擇發票年份",
    options=valid_years_inv,
    default=selected_years
)

region = st.sidebar.multiselect("選擇地區", df["Region"].unique())

filter_import = df_import[df_import['YEAR'].isin(selected_years)]
filter_smt = df[
    (df['Inv_Yr'].isin(selected_inv_years)) &
    (df['BRAND'].isin(['HELLER', 'PEMTRON', 'YAMAHA']))
]
if region: filter_smt = filter_smt[filter_smt["Region"].isin(region)]

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
                marker=dict(symbol='square', size=6, color='black'),
                text=df_year['台数'],
                textposition='middle right',
                textfont=dict(color='black', size=20)
            ), secondary_y=True)
        
        fig.update_layout(
            height=600,
            title={
                'text': "📈 China Mounter Import Trend(QTY & CNY Amount)",
                'font': {'size': 24},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title='月份',
                tickmode='linear',
                dtick=1,
                gridcolor='black',
                gridwidth=1,
                showgrid=True,
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            yaxis=dict(
                title='金額 (百萬人民幣)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                tickformat='.0fM',
                showgrid=True,
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            yaxis2=dict(
                title='數量 (台)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                showgrid=True,
                rangemode='tozero',
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            plot_bgcolor='rgba(255,255,255,0.9)',
            paper_bgcolor='rgb(240,240,240)',
            hovermode='x unified',
            showlegend=True,
            legend=dict(orientation="h",
                font=dict(size=20),
                bgcolor='rgba(255,255,255,0.8)'
            )
        )
        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("進口數據樞紐分析表"):
            pivot_import = filter_import.pivot_table(
                values=["进口金额（人民币）", "台数"],
                index="MONTH",
                columns="YEAR",
                aggfunc="sum",
                margins=True,
                margins_name="總計"
            ).fillna(0)
            
            formatted_pivot = pivot_import.applymap(format_numbers)
            
            html = formatted_pivot.to_html(classes='table table-bordered')
            html = html.replace(
                '<th>台数</th>',
                '<th style="background-color: #90EE90">台数</th>'
            ).replace(
                '<th>进口金额（人民币）</th>',
                '<th style="background-color: #FFA500">进口金额（人民币）</th>'
            )
            st.markdown(f'<div style="zoom:1.1">{html}</div>', unsafe_allow_html=True)
            
            formatted_csv = pivot_import.applymap(format_numbers)
            csv = formatted_csv.to_csv(encoding='utf-8-sig').encode('utf-8-sig')
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
                marker=dict(symbol='square', size=6, color='black'),
                text=df_year['Item Qty'],
                textposition='middle right',
                textfont=dict(color='black', size=20)
            ), secondary_y=True)
        
        fig.update_layout(
            height=600,
            title={
                'text': "📊 SMT Invoice Trend(<span style='color:#FF0000'>YAMAHA</span>/<span style='color:#0000FF'>HELLER</span>/<span style='color:#006400'>PEMTRON</span> INV QTY & INV Amount HKD)",
                'font': {'size': 24},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title='月份',
                tickmode='linear',
                dtick=1,
                gridcolor='black',
                gridwidth=1,
                showgrid=True,
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            yaxis=dict(
                title='金額 (百萬港幣)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                tickformat='.0fM',
                showgrid=True,
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            yaxis2=dict(
                title='數量 (台)',
                gridcolor='rgba(0,0,0,0.3)',
                gridwidth=0.5,
                showgrid=True,
                rangemode='tozero',
                tickfont=dict(color='#333333', size=24),
                title_font=dict(color='#333333', size=27)
            ),
            plot_bgcolor='rgba(255,255,255,0.9)',
            paper_bgcolor='rgb(240,240,240)',
            hovermode='x unified',
            showlegend=True,
            legend=dict(orientation="h",
                font=dict(size=20),
                bgcolor='rgba(255,255,255,0.8)'
            )
        )
        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("發票數據樞紐分析表"):
            pivot_smt = filter_smt.pivot_table(
                values=["Item Qty", "Before tax Inv Amt (HKD)"],
                index=["BRAND", "Inv_Month"],
                columns="Inv_Yr",
                aggfunc="sum",
                margins=True,
                margins_name="總計"
            ).fillna(0)
            
            brand_totals = filter_smt.groupby('BRAND')['Item Qty'].sum().sort_values(ascending=False)
            ordered_brands = brand_totals.index.tolist()
            pivot_smt = pivot_smt.reindex(index=ordered_brands, level='BRAND')
            
            formatted_smt = pivot_smt.applymap(format_numbers)
            
            html = formatted_smt.to_html()
            html = html.replace(
                '<th>HELLER</th>', 
                '<th style="background-color: #FFA500">HELLER</th>'
            ).replace(
                '<th>PEMTRON</th>', 
                '<th style="background-color: #006400; color: white">PEMTRON</th>'
            ).replace(
                '<th>YAMAHA</th>', 
                '<th style="background-color: #FFB6C1">YAMAHA</th>'
            )
            st.markdown(f'<div style="zoom:1.1">{html}</div>', unsafe_allow_html=True)
            
            formatted_csv = pivot_smt.applymap(format_numbers)
            csv = formatted_csv.to_csv(encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("下載發票數據", csv, "smt_invoice.csv", "text/csv")

style_metric_cards(background_color="#FFFFFF", border_left_color="#686664")

st.markdown("""
<div style="text-align:center;padding:1rem;margin-top:2rem;border-top:2px solid #ddd">
    <p style="color:#666">Developed by Arthur Chan • Data Version: 2024-02</p>
</div>
""", unsafe_allow_html=True)









#加pivot table, column只用YAMAHA, PEMTRON, HELLER台數要match圖
################################################################################################################################# 
#-China Mounter Import Trend和SMT Invoice Trend:兩部分用st.column(2)分兩邊
#-所有st.expander都取消
#-China Mounter Import Trend的range bar預設全選
#-China Mounter Import Trend的pivot table中的"台数"這字用綠色背景色；"进口金额（人民币）"這字用綠色背景色
#-China Mounter Import Trend的combine chart加入Title，名為"China Mounter Import Trend(QTY& CNY Amount)，字體跟subheader大小相若
#-China Mounter Import Trend的combine chart背景色不再用紅色，改用淡黃色
#-China Mounter Import Trend的combine chart中的进口金额（人民币）的scale用m代替現在的b，即是以百萬作單位，50000000就用50M去表示
#-SMT Invoice Trend的combine chart加入Title，名為"SMT Invoice Trend(QTY& HKD Amount)"，字體跟subheader大小相若
#-所有combine chart都改用solid的全格網線
#-所有combine chart中，bar和折線再分開一點，折線的每格scale再拆細一半，令到折線可以放高一點
#-所有combine chart中的折線的marker改用solid的正方形
#-所有combine chart中，bar的顏色比例淡一點，另外最新一年的數據的bar和折線都一律用橙色
#-range bar直接改為sidebar filter
#-所有combine chart我要的格網線是正方形一格一格的那種，取消現在那種格線重新設置
#-China Mounter Import Trend的combine chart中"台数"的數字要直接顯示在折線圖的marker之上
#-SMT Invoice Trend的combine chart中"數量"的數字要直接顯示在折線圖的marker之上
#-進口年份範圍sidebar filter預設所有選項
#-發票年份sidebar filter預設要視乎進口年份範圍sidebar filter，例如進口年份範圍sidebar filter預設選擇了2022和2023，那麼發票年份sidebar filter預設就照樣是2022和2023
#-所有combine chart的Title要按我之前指示補回給我
#-China Mounter Import Trend的combine chart中"台数"的數字改為顯示在折線圖的marker右邊
#-SMT Invoice Trend的combine chart中"數量"的數字改為顯示在折線圖的marker右邊
#-China Mounter Import Trend的combine chart中的进口金额（人民币）的scale參考SMT Invoice Trend的-combine chart中bar的scale顯示方式，現在雖然是以百萬作單位，但20000000在圖上scale現在只顯示為200，我想要的是在200後加一個"M"，寫為"20M"
#-之前China Mounter Import Trend(QTY & CNY)和SMT Invoice Trend(QTY & HKD)兩部分的pivot table都消失了，要按原來的模樣不變地補回給我
#-仍未見表格有做到條件格式：進口數據樞紐分析表的pivot table內，"台数"字眼背景色用綠色，"进口金额（人民币）"字眼背景色用橙色
#-仍未見表格有做到條件格式：SMT Invoice Trend的pivot table內，HELLER字眼背景色用橙色，PEMTRON字眼背景色用淺藍色，YAMAHA字眼背景色用淺紅色
#-另外SMT Invoice Trend的pivot table的內，Inv_Month要放在BRAND的位置後面，不要把Inv_Month放在Inv_Yr後面。BRAND和Inv_Month同樣在row，只有Inv_Yr在列
