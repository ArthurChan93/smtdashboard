import streamlit as st
from streamlit_option_menu import option_menu
import plotly.express as px
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
import os
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from streamlit_extras.metric_cards import style_metric_cards
import seaborn as sns
import base64
from io import BytesIO
import plotly.graph_objects as go
from bs4 import BeautifulSoup
import locale
import re
from lxml import etree
from PIL import Image
from streamlit.components.v1 import html
######################################################################################################
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Contract Dashboard",page_icon=":rainbow:",layout="wide")

title_row1, title_row2, title_row3, title_row4 = st.columns(4)


st.title(':pencil: SMT_Contract Dashboard')
#Text Credit
st.write("by Arthur Chan")

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
######################################################################################################
#Create a browser for user to upload
#@st.cache_data
#def load_data(file):
#       data = pd.read_excel(file)
#       return data
#uploaded_file = st.sidebar.file_uploader(":file_folder: Upload monthly report here")
#if uploaded_file is not None:
#        df = load_data(uploaded_file)
#        st.dataframe(df)
#唔show 17/18, cancel, tba資料
#else:
#os.chdir(r"/Users/arthurchan/Library/CloudStorage/OneDrive-個人/Monthly Report")
#os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")

df = pd.read_excel(
               io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='raw_sheet', skiprows=0, usecols='A:AT',nrows=100000,).query(
                    'Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('FY_Contract != "FY 17/18"').query(
                         'FY_Contract != "Cancel"').query('Contract_Yr != "TBA"').query('Contract_Yr != "Cancel"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"')
df_sales_target = pd.read_excel(
               io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='YAMAHA_Sales_Target', skiprows=0, usecols='A:D',nrows=10000,)
######################################################################################################
st.sidebar.write(":orange[Default ALL if no selection] :white_check_mark:")


# Sidebar Filter
# Create FY Contract filter
# Define sidebar filters and create corresponding DataFrames for each filter
# Function to sort values by frequency and then value
def sort_options(series):
    # Drop NaN values, convert to string, count frequencies, and sort
    series = series.dropna().astype(str)
    freq = series.value_counts()
    sorted_options = freq.sort_values(ascending=False).index
    return sorted_options

# Sidebar Filter
FY_Contract_filter = st.sidebar.multiselect(
    "FY_Contract", sort_options(df["FY_Contract"]), default=["FY 24/25", "FY 23/24"]
)
df_FY_Contract = df[df["FY_Contract"].isin(FY_Contract_filter)]

fq_invoice_filter = st.sidebar.multiselect(
    "FQ(Contract)", sort_options(df["FQ(Contract)"])
)
df_fq_invoice = df[df["FQ(Contract)"].isin(fq_invoice_filter)]

Contract_Month_filter = st.sidebar.multiselect(
    "INV MONTH", sort_options(df["Contract_Month"])
)
df_Contract_Month = df[df["Contract_Month"].isin(Contract_Month_filter)]

Region_filter = st.sidebar.multiselect(
    "REGION", sort_options(df["Region"])
)
df_Region = df[df["Region"].isin(Region_filter)]

cost_centre_filter = st.sidebar.multiselect(
    "COST CENTRE", sort_options(df["COST_CENTRE"])
)
df_cost_centre = df[df["COST_CENTRE"].isin(cost_centre_filter)]

brand_filter = st.sidebar.multiselect(
    "BRAND", sort_options(df["BRAND"])
)
df_brand = df[df["BRAND"].isin(brand_filter)]

Ordered_Items_filter = st.sidebar.multiselect(
    "MODEL", sort_options(df["Ordered_Items"])
)
df_Ordered_Items = df[df["Ordered_Items"].isin(Ordered_Items_filter)]

Customer_Name_filter = st.sidebar.multiselect(
    "CUSTOMER", sort_options(df["Customer_Name"])
)
df_Customer_Name = df[df["Customer_Name"].isin(Customer_Name_filter)]

New_Customer_filter = st.sidebar.multiselect(
    "New Customer", sort_options(df["New_Customer"])
)
df_New_Customer = df[df["New_Customer"].isin(New_Customer_filter)]

# Handle different filter combinations
if FY_Contract_filter and fq_invoice_filter and Contract_Month_filter and Region_filter and cost_centre_filter and brand_filter and Ordered_Items_filter and Customer_Name_filter and New_Customer_filter:
    # All filters are selected
    filter_df = df_FY_Contract
    filter_df = filter_df[filter_df["FQ(Contract)"].isin(fq_invoice_filter)]
    filter_df = filter_df[filter_df["Contract_Month"].isin(Contract_Month_filter)]
    filter_df = filter_df[filter_df["Region"].isin(Region_filter)]
    filter_df = filter_df[filter_df["COST_CENTRE"].isin(cost_centre_filter)]
    filter_df = filter_df[filter_df["BRAND"].isin(brand_filter)]
    filter_df = filter_df[filter_df["Ordered_Items"].isin(Ordered_Items_filter)]
    filter_df = filter_df[filter_df["Customer_Name"].isin(Customer_Name_filter)]
    filter_df = filter_df[filter_df["New_Customer"].isin(New_Customer_filter)]
elif not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter and not cost_centre_filter and not brand_filter and not Ordered_Items_filter and not Customer_Name_filter and not New_Customer_filter:
    # No filters are selected
    filter_df = df
else:
    # Other filter combinations
    filter_df = pd.DataFrame(columns=df.columns)  # Create an empty DataFrame

    if FY_Contract_filter:
        filter_df = pd.concat([filter_df, df_FY_Contract])
    
    if fq_invoice_filter:
        if not FY_Contract_filter:
            filter_df = pd.concat([filter_df, df_fq_invoice])
        else:
            filter_df = filter_df[filter_df["FQ(Contract)"].isin(fq_invoice_filter)]
    
    if Contract_Month_filter:
        if not FY_Contract_filter and not fq_invoice_filter:
            filter_df = pd.concat([filter_df, df_Contract_Month])
        else:
            filter_df = filter_df[filter_df["Contract_Month"].isin(Contract_Month_filter)]
    
    if Region_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter:
            filter_df = pd.concat([filter_df, df_Region])
        else:
            filter_df = filter_df[filter_df["Region"].isin(Region_filter)]
    
    if cost_centre_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter:
            filter_df = pd.concat([filter_df, df_cost_centre])
        else:
            filter_df = filter_df[filter_df["COST_CENTRE"].isin(cost_centre_filter)]
    
    if brand_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter and not cost_centre_filter:
            filter_df = pd.concat([filter_df, df_brand])
        else:
            filter_df = filter_df[filter_df["BRAND"].isin(brand_filter)]
    
    if Ordered_Items_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter and not cost_centre_filter and not brand_filter:
            filter_df = pd.concat([filter_df, df_Ordered_Items])
        else:
            filter_df = filter_df[filter_df["Ordered_Items"].isin(Ordered_Items_filter)]
    
    if Customer_Name_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter and not cost_centre_filter and not brand_filter and not Ordered_Items_filter:
            filter_df = pd.concat([filter_df, df_Customer_Name])
        else:
            filter_df = filter_df[filter_df["Customer_Name"].isin(Customer_Name_filter)]
    
    if New_Customer_filter:
        if not FY_Contract_filter and not fq_invoice_filter and not Contract_Month_filter and not Region_filter and not cost_centre_filter and not brand_filter and not Ordered_Items_filter and not Customer_Name_filter:
            filter_df = pd.concat([filter_df, df_New_Customer])
        else:
            filter_df = filter_df[filter_df["New_Customer"].isin(New_Customer_filter)]
#Show the original data table
#st.dataframe(df_selection)
############################################################################################################################################################################################################
#Overall Summary
left_column, middle_column, right_column = st.columns(3)
total_invoice_amount = int(filter_df["Before tax Inv Amt (HKD)"].sum())
with left_column:
      #st.subheader((f":dollar: Total INV AMT before tax: :orange[HKD{total_invoice_amount:,}]"))
      st.markdown(f"""
            <div style='background-color: #d0f0c0; padding: 10px; border-radius: 12px; box-shadow: 5px 5px 15px #b0bec5, -5px -5px 15px #ffffff; text-align: center; font-size: 1.5em;'>
            <strong>💵 Total INV AMT before tax: <span style="color: orange;">HKD{total_invoice_amount:,}</span></strong>
            </div>
            """, unsafe_allow_html=True)



total_gp = int(filter_df["G.P.  (HKD)"].sum())
with middle_column:
     #st.subheader(f":moneybag: Total G.P AMT: :orange[HKD{total_gp:,}]")
      st.markdown(f"""
            <style>
            .subheader-neumorphism {{
            background-color: #d0f0c0;
            padding: 10px;
            border-radius: 12px;
            box-shadow: 5px 5px 15px #b0bec5, -5px -5px 15px #ffffff;
            text-align: center;
            font-size: 1.5em;
            }}
            </style>
            <div class="subheader-neumorphism">
            <strong>💰 Total G.P AMT: <span style="color: orange;">HKD{total_gp:,}</span></strong>
            </div>
            """, unsafe_allow_html=True)


invoice_qty = filter_df[(filter_df['BRAND'] != 'LOCAL SUPPLIER') & (filter_df['BRAND'] != 'SOLDERSTAR')& 
            (filter_df['BRAND'] != 'SHINWA')& (filter_df['BRAND'] != 'SIGMA')& (filter_df['BRAND'] != 'C66 SERVICE')& 
            (filter_df['BRAND'] != 'SHIMADZU')& (filter_df['BRAND'] != 'NUTEK')& (filter_df['BRAND'] != 'SAKI')
            & (filter_df['BRAND'] != 'DEK')
            & (filter_df['FY_Contract'] != 'TBA')& (filter_df['FY_Contract'] != 'Cancel')]

OnlyYAMAHA_HELLER_PEMTRON_qty = invoice_qty['Item Qty'].sum()
header_qty = int(OnlyYAMAHA_HELLER_PEMTRON_qty)  # 使用字符串格式化将数字插入标题中
total_unit_qty = int(header_qty)

with right_column:
     #st.subheader(f":factory: INV Qty(YAMAHA, PEMTRON, HELLER): :orange[{total_unit_qty:,}]")
      st.markdown(f"""
            <style>
            .subheader-neumorphism {{
            background-color: #d0f0c0;
            padding: 10px;
            border-radius: 12px;
            box-shadow: 5px 5px 15px #b0bec5, -5px -5px 15px #ffffff;
            text-align: center;
            font-size: 1.5em;
            }}
            </style>
            <div class="subheader-neumorphism">
            <strong>🏭 INV Qty: <span style="color: orange;">{total_unit_qty:,}</strong>(YAMAHA, PEMTRON, HELLER)</span></strong>
            </div>
            """, unsafe_allow_html=True)

##st.divider()
st.markdown(
    """
    <hr style="border: 3px solid green;">
    """,
    unsafe_allow_html=True
)
############################################################################################################################################################################################################
#Pivot table, 差sub-total, GP%
filter_df["Contract_Month"] = filter_df["Contract_Month"].astype(str)
filter_df["Contract_Yr"] = filter_df["Contract_Yr"].astype(str)
filter_df["Contract_No."] = filter_df["Contract_No."].astype(str)

# add subtotal
# https://morioh.com/a/17278219952b/tutorial-on-data-analysis-with-python-and-pivot-tables-with-pandas

 
############################################################################################################################################################################################################     
#Create tabs after overall summary

 
# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
font-size: 28px;
}
</style>
"""

 
st.write(font_css, unsafe_allow_html=True)
st.markdown(
    """
    <style>
    .stTabs [data-baseweb="tab"] {
        border: 5px solid lightgreen;
        padding: 10px;
        border-radius: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

tab1, tab2, tab3 ,tab4,tab5, tab6= st.tabs([":wedding: 概览",":earth_asia: 地区",":books: 品牌",":handshake: 客户",":alarm_clock: Contract 交付周期",":blue_book: INVOICE明細"])

#TAB 1: Overall category
################################################################################################################################################

with tab1:

       col_1, col_2= st.columns(2)
       with col_1:
#LINE CHART of Overall Contract Amount
             st.subheader(":chart_with_upwards_trend: New Contract趋势_:orange[月份]:")
             InvoiceAmount_df2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","FQ(Contract)","Contract_Month"
                          ], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             InvoiceAmount_df2 = InvoiceAmount_df2.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([InvoiceAmount_df2['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
             fig3 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = InvoiceAmount_df2['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                   FY_Contract_data = InvoiceAmount_df2[InvoiceAmount_df2['FY_Contract'] == FY_Contract]
                   fig3.add_trace(go.Scatter(
                         x=FY_Contract_data['Contract_Month'],
                         y=FY_Contract_data['Before tax Inv Amt (HKD)'],
                         mode='lines+markers+text',
                         name=FY_Contract,
                         text=FY_Contract_data['Before tax Inv Amt (HKD)'],
                         textposition="bottom center",
                         texttemplate='%{text:.3s}',
                         hovertemplate='%{x}<br>%{y:.2f}',
                         marker=dict(size=10)))
                   fig3.update_layout(xaxis=dict(
                         type='category',
                         categoryorder='array',
                         categoryarray=sort_Month_order),
                         yaxis=dict(showticklabels=True),
                         font=dict(family="Arial, Arial", size=12, color="Black"),
                         hovermode='x', showlegend=True,
                         legend=dict(orientation="h",font=dict(size=14)))
             fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))

# 绘制图表

             st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
#############################################################################################################
#FY to FY Quarter Contract Details:
             pvt6 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_Contract"],columns=["FQ(Contract)"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total")
             html11 = pvt6.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             #st.dataframe(pvt6.style.highlight_max(color = 'yellow', axis = 0)
             #                       .format("HKD{:,}"), use_container_width=True)   
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html12 = html11.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html13 = html12.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
             html14 = html13.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
             html15 = html14.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
             html16 = html15.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
             html117 = html16.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html_with_style = str(f'<div style="zoom: 0.8;">{html117}</div>')
             st.markdown(html_with_style, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv2 = pvt6.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv2, file_name='FQ_Sales.csv', mime='text/csv')
################################################################################################################################################
       with col_2:
             st.subheader(":round_pushpin: Contract 百分比:")
# 创建示例数据框
             brand_data = filter_df.round(0).groupby(by=["FY_Contract","COST_CENTRE"],
                     as_index=False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["COST_CENTRE"] = brandinvpie_df["COST_CENTRE"].replace(to_replace=[x for x in brandinvpie_df["COST_CENTRE"
                                       ].unique() if x not in ["C49","C28","C66"]], value="OTHERS")
             brandinvpie_df["COST_CENTRE"] = pd.Categorical(brandinvpie_df["COST_CENTRE"], ["C49","C28","C66"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Before tax Inv Amt (HKD)", names="COST_CENTRE", color="COST_CENTRE", color_discrete_map={
                      "C28": "lightblue", "C49": "orange","C66": "purple"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                   paper_bgcolor='lightgrey',  # 设置背景颜色为浅灰色
                   plot_bgcolor='lightgrey',   # 设置绘图区域背景颜色为浅灰色
                   margin=dict(l=10, r=10, t=10, b=30),  # 设置图表的边距
                   autosize=False,
                   width=600,
                   height=400)
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line=dict(color='black', width=2))
             
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)    



       ##st.divider() 用以下markdown 框線代替divider用以下markdown 框線代替divider
       st.markdown(
            """
            <hr style="border: 3px solid lightblue;">
            """,
            unsafe_allow_html=True
            )        

################################################################################################################################################
#New Section 
#LINE CHART of GP Amount      
       col_3, col_4= st.columns(2)
       with col_3:
             st.subheader(":chart_with_upwards_trend: :green[G.P Amount Trend]_FQ:")
             InvoiceAmount_df2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).groupby(by =
                     ["Contract_Month","FY_Contract"], as_index= False)["G.P.  (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             InvoiceAmount_df2 = InvoiceAmount_df2.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([InvoiceAmount_df2['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
             fig12 = go.Figure()

# 添加每个FY_Contract的折线
             FY_Contract_values = InvoiceAmount_df2['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                   FY_Contract_data = InvoiceAmount_df2[InvoiceAmount_df2['FY_Contract'] == FY_Contract]
                   fig12.add_trace(go.Scatter(
                         x=FY_Contract_data['Contract_Month'],
                         y=FY_Contract_data['G.P.  (HKD)'],
                         mode='lines+markers+text',
                         name=FY_Contract,
                         text=FY_Contract_data['G.P.  (HKD)'],
                         textposition="bottom center",
                         texttemplate='%{text:.3s}',
                         hovertemplate='%{x}<br>%{y:.2f}',
                         marker=dict(size=10)))
                   fig12.update_layout(xaxis=dict(
                         type='category',
                         categoryorder='array',
                         categoryarray=sort_Month_order),
                         yaxis=dict(showticklabels=True),
                         font=dict(family="Arial, Arial", size=12, color="Black"),
                         hovermode='x', showlegend=True,
                         legend=dict(orientation="h",font=dict(size=14)))
             fig12.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig12.update_layout(yaxis_showticklabels = True), use_container_width=True)
#################################################
      #FY to FY Quarter Contract Details:
             pvt16 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).pivot_table(values="G.P.  (HKD)",index=["FY_Contract"],columns=["FQ(Contract)"],
                                   aggfunc="sum",fill_value=0, margins=True,margins_name="Total")
             html17 = pvt16.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html18 = html17.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html19 = html18.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
             html20 = html19.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html21 = html20.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
             html22 = html21.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html23 = html22.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html_with_style = str(f'<div style="zoom: 0.8;">{html23}</div>')
             st.markdown(html_with_style, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv3 = pvt16.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv3, file_name='G.P Amount.csv', mime='text/csv')
################################################################################################################################################
       with col_4:
             st.subheader(":round_pushpin: Contract 百分比:")
# 创建示例数据框
             brand_data = filter_df.round(0).groupby(by=["FY_Contract","COST_CENTRE"],
                     as_index=False)["G.P.  (HKD)"].sum().sort_values(by="G.P.  (HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["COST_CENTRE"] = brandinvpie_df["COST_CENTRE"].replace(to_replace=[x for x in brandinvpie_df["COST_CENTRE"
                                       ].unique() if x not in ["C49","C28","C66"]], value="OTHERS")
             brandinvpie_df["COST_CENTRE"] = pd.Categorical(brandinvpie_df["COST_CENTRE"], ["C49","C28","C66"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="G.P.  (HKD)", names="COST_CENTRE", color="COST_CENTRE", color_discrete_map={
                      "C28": "blue", "C49": "lightgreen","C66": "pink"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                   paper_bgcolor='lightgrey',  # 设置背景颜色为浅灰色
                   plot_bgcolor='lightgrey',   # 设置绘图区域背景颜色为浅灰色
                   margin=dict(l=10, r=10, t=10, b=30),  # 设置图表的边距
                   autosize=False,
                   width=600,
                   height=400)
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line=dict(color='black', width=2))
           
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)    
       
       
       ##st.divider()
       st.markdown(
            """
            <hr style="border: 3px solid lightblue;">
            """,
            unsafe_allow_html=True
            )
####################################################   
       st.subheader(""":globe_with_meridians: Contract Amount Table_:orange[Monthly]:""")

       pvt2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(0).pivot_table(
              index=["FY_Contract","FQ(Contract)","Contract_Yr", "Contract_Month"],
              columns=["COST_CENTRE"],
              values=["Before tax Inv Amt (HKD)"],
              aggfunc="sum",
              fill_value=0,
              margins=True,
              margins_name="Total",
              observed=True)

# 調整columns順序
       columns_order = ["C49", "C28", "C66","Total"]
       pvt2 = pvt2.reindex(columns=columns_order, level=1)

       #使用applymap方法應用格式化
       pvt2 = pvt2.applymap('{:,.0f}'.format)
       html3 = pvt2.to_html(classes='table table-bordered', justify='center')
       html4 = html3.replace('<th>C66</th>', '<th style="background-color: orange">C66</th>')
       html5 = html4.replace('<th>C28</th>', '<th style="background-color: lightblue">C28</th>')
       html6 = html5.replace('<th>NORTH</th>', '<th style="background-color: Khaki">NORTH</th>')
       html7 = html6.replace('<th>C49</th>', '<th style="background-color: lightgreen">C49</th>')

# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
       html8 = html7.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把每個數值置中
       html9 = html8.replace('<td>', '<td style="text-align: middle;">')
# 把REGION值的那列的字改色
       html10 = html9.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
       html11 = html10.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
       html12 = html11.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
       html13 = html12.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
      
# 把所有數值等於或少於0的數值的顏色設為紅色
       html14 = html13.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
       html15 = f'<div style="zoom: 1.2;">{html14}</div>'
       st.markdown(html15, unsafe_allow_html=True)           
 
 
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
       csv1 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
       st.download_button(label='Download Table', data=csv1, file_name='Monthly_Sales.csv', mime='text/csv')

################################################################################
#Pivot table2
       st.subheader(":point_down: Contract Amount Subtotal_:orange[FQ]:clipboard: ")
       #with st.expander(":point_right: click to expand/hide"):
       pvt17 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(
                     values="Before tax Inv Amt (HKD)",
                     index=["FY_Contract","FQ(Contract)"],
                     columns=["COST_CENTRE"],
                     aggfunc="sum",
                     fill_value=0,
                     margins=True,
                     margins_name="Total",
                     observed=True)  # This ensures subtotals are only calculated for existing values)
       desired_order = ["C49", "C28", "C66","Total"]
       pvt17 = pvt17.reindex(columns=desired_order, level=1)

# 定义会计数字格式的格式化函数
       def format_currency(value):
                   return "HKD{:,.0f}".format(value)

# 计算小计行
       subtotal_row = pvt17.groupby(level=0).sum(numeric_only=True)
       subtotal_row.index = pd.MultiIndex.from_product([subtotal_row.index, [""]])
       subtotal_row.name = ("Subtotal", "")  # 小计行索引的名称

# 去除千位數符號並轉換為浮點數
       pvt17 = pvt17.applymap(lambda x: float(str(x).strip('HKD').replace(',', '')))

# 转换为字符串并添加样式
       pvt17 = pvt17.applymap(lambda x: "HKD{:,.0f}".format(x))

# 将小计行与pvt17连接，使用concat函数
       pvt17_concatenated = pd.concat([pvt17, subtotal_row])

# 生成HTML表格
       html_table = pvt17_concatenated.to_html(classes='table table-bordered', justify='center')

# 使用BeautifulSoup处理HTML表格
       soup = BeautifulSoup(html_table, 'html.parser')

# 找到所有的<td>标签，并为小于或等于0的值添加CSS样式
       for td in soup.find_all('td'):
                   value = float(td.text.replace('HKD', '').replace(',', ''))
       if value <= 0:
                   td['style'] = 'color: red;'
      
# 找到所有的<td>标签，并将数值转换为会计数字格式的字符串
       for td in soup.find_all('td'):
                   value = float(td.text.strip('HKD').replace(',', ''))
                   formatted_value = "HKD{:,.0f}".format(value)
                   td.string.replace_with(formatted_value)
# 找到最底部的<tr>标签，并为其添加CSS样式
       last_row = soup.find_all('tr')[-1]
       last_row['style'] = 'background-color: yellow; font-weight: bold;'

# 在特定单元格应用其他样式           
       soup = str(soup)
       soup = soup.replace('<th>C66</th>', '<th style="background-color: orange">C66</th>')
       soup = soup.replace('<th>C28</th>', '<th style="background-color: lightblue">C28</th>')
       soup = soup.replace('<th>HKD0</th>', '<th style="background-color: Khaki">HKD0</th>')
       soup = soup.replace('<th>C49</th>', '<th style="background-color: lightgreen">C49</th>')
       soup = soup.replace('<td>', '<td style="text-align: middle;">')
       soup = soup.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
       soup = soup.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
       soup = soup.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
       soup = soup.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
       soup = soup.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
       html_with_style = str(f'<div style="zoom: 1.2;">{soup}</div>')
       st.markdown(html_with_style, unsafe_allow_html=True)       
       
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
       csv6 = pvt17.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
       st.download_button(label='Download Table', data=csv6, file_name='Cost_Centre_Quarter_Sales.csv', mime='text/csv')
       ##st.divider()
       st.markdown(
            """
            <hr style="border: 3px solid lightblue;">
            """,
            unsafe_allow_html=True
            )
       
############################################################################################################################################################################################################
#TAB 2: Region Category
with tab2:

        one_column, two_column= st.columns(2)
        with one_column:
#All Regional total inv amount BAR CHART
              st.subheader(":bar_chart: Contract Amount_:orange[FY]:")
              category2_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by=["FY_Contract","Region"], 
                       as_index=False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)", ascending=False)
              df_contract_vs_invoice = px.bar(category2_df, x="FY_Contract", y="Before tax Inv Amt (HKD)", color="Region", text_auto='.3s')

# 更改顏色
              colors = {"SOUTH": "orange","EAST": "lightblue","NORTH": "Khaki","WEST": "lightgreen",}
              for trace in df_contract_vs_invoice.data:
                    region = trace.name.split("=")[-1]
                    trace.marker.color = colors.get(region, "blue")

# 更改字體
              df_contract_vs_invoice.update_layout(font=dict(family="Arial", size=14))
              df_contract_vs_invoice.update_traces(marker_line_color='black', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
              df_contract_vs_invoice.update_layout(barmode='group')
# 将图例放在底部
              df_contract_vs_invoice.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.03, xanchor="right", x=1))

# 添加背景色
              background_color2 = 'lightgrey'
              x_range2 = len(category2_df["FY_Contract"].unique())
              background_shapes2 = [dict(
              type='rect',
              xref='x',
              yref='paper',
              x0=i - 0.5,
              y0=0,
              x1=i + 0.5,
              y1=1,
              fillcolor=background_color2,
              opacity=0.1,
              layer='below',
              line=dict(width=5)) for i in range(x_range2)]
             
              df_contract_vs_invoice.update_layout(shapes=background_shapes2, showlegend=True)
          
              st.plotly_chart(df_contract_vs_invoice, use_container_width=True) 

####All Region PIE CHART
        with two_column:
             st.subheader(":round_pushpin: Contract 百分比:")
# 创建示例数据框
             brand_data = filter_df.round(0).groupby(by=["FY_Contract","Region"],
                     as_index=False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)", ascending=False)         
             brandinvpie_df = pd.DataFrame(brand_data)
# 按照指定順序排序 
             brandinvpie_df["Region"] = brandinvpie_df["Region"].replace(to_replace=[x for x in brandinvpie_df["Region"
                                       ].unique() if x not in ["SOUTH","EAST", "WEST", "NORTH"]], value="OTHERS")
             brandinvpie_df["Region"] = pd.Categorical(brandinvpie_df["Region"], ["SOUTH","EAST", "WEST", "NORTH","OTHERS"])
# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Before tax Inv Amt (HKD)", names="Region", color="Region", color_discrete_map={
                      "EAST": "lightblue", "SOUTH": "orange", "WEST": "lightgreen","NORTH": "khaki"})
# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                   paper_bgcolor='lightgrey',  # 设置背景颜色为浅灰色
                   plot_bgcolor='lightgrey',   # 设置绘图区域背景颜色为浅灰色
                   margin=dict(l=10, r=10, t=10, b=30),  # 设置图表的边距
                   autosize=False,
                   width=600,
                   height=400)
# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line=dict(color='black', width=2))
# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
##############################################################################################################################          
# LINE CHART of Regional Comparision
        st.subheader(":chart_with_upwards_trend: Contract Amount Trend_:orange[All Region in one]:")
        InvoiceAmount_df2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by = ["FQ(Contract)","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
        # 使用pivot_table函數來重塑數據，使每個Region成為一個列
        InvoiceAmount_df2 = InvoiceAmount_df2.pivot_table(index="FQ(Contract)", columns="Region", values="Before tax Inv Amt (HKD)", fill_value=0).reset_index()
        # 使用melt函數來恢復原來的長格式，並保留0值
        InvoiceAmount_df2 = InvoiceAmount_df2.melt(id_vars="FQ(Contract)", value_name="Before tax Inv Amt (HKD)", var_name="Region")
        fig2 = px.line(InvoiceAmount_df2,
                       x = "FQ(Contract)",
                       y = "Before tax Inv Amt (HKD)",
                       color='Region',
                       markers=True,
                       text="Before tax Inv Amt (HKD)",
                       color_discrete_map={'SOUTH': 'orange','EAST': 'lightblue',
                                           'NORTH': 'Khaki','WEST': 'lightgreen'})
              # 更新圖表的字體大小和粗細
        fig2.update_layout(font=dict(
                    family="Arial, Arial",
                    size=12,
                    color="Black"))
        fig2.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))
        fig2.update_traces(marker_size=9, textposition="bottom center", texttemplate='%{text:.2s}')
        st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)
###############################################################################################

        one_column, two_column= st.columns(2)
        with one_column:
# LINE CHART of SOUTH CHINA FY/FY
              #st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True
                   )
              st.subheader(":chart_with_upwards_trend: :orange[SOUTH CHINA] Inv Amt Trend_FY to FY:")
              df_Single_south = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "SOUTH"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract",
                                 "Contract_Month"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Contract)" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_south = df_Single_south.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_south['FY_Contract'].unique(), all_fq_invoice_values],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个FY_Contract的折线
              FY_Contract_values = df_Single_south['FY_Contract'].unique()
              for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_south[df_Single_south['FY_Contract'] == FY_Contract]
               fig4.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='rgba(255,165,0,0.3)')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)

             
#SOUTH Region Contract Details FQ_FQ:
              pvt8 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "SOUTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",
                     index=['FY_Contract'],columns=["FQ(Contract)"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)
            
              html76 = pvt8.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: orange">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: orange">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: orange">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: orange">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv11 = pvt8.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv11, file_name='South_Sales.csv', mime='text/csv')
              #st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True
                   )

 

        with two_column:
# LINE CHART of EAST CHINA FY/FY
              #st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)
              st.subheader(":chart_with_upwards_trend: :orange[EAST CHINA] Inv Amt Trend_FY to FY:")
              df_Single_region = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "EAST"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Contract)" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_region = df_Single_region.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_region['FY_Contract'].unique(), all_fq_invoice_values],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
              fig5 = go.Figure()

# 添加每个FY_Contract的折线
              FY_Contract_values = df_Single_region['FY_Contract'].unique()
              for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_region[df_Single_region['FY_Contract'] == FY_Contract]
               fig5.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.2s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig5.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='rgba(0,150,255,0.1)')
             
              fig5.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
              st.plotly_chart(fig5.update_layout(yaxis_showticklabels = True), use_container_width=True)

########################             
#EAST Region Contract Details FQ_FQ:
              pvt9 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "EAST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)
            
              html83 = pvt9.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html84 = html83.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html85 = html84.replace('<th>Q1</th>', '<th style="background-color: lightblue">Q1</th>')
              html86 = html85.replace('<th>Q2</th>', '<th style="background-color: lightblue">Q2</th>')
              html87 = html86.replace('<th>Q3</th>', '<th style="background-color: lightblue">Q3</th>')
              html88 = html87.replace('<th>Q4</th>', '<th style="background-color: lightblue">Q4</th>')
              html89 = html88.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html900 = f'<div style="zoom: 0.7;">{html89}</div>'
             
              st.markdown(html900, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv12 = pvt9.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv12, file_name='East_Sales.csv', mime='text/csv')

              #st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)
################################################# 
        three_column, four_column= st.columns(2)
  
        with three_column:
# LINE CHART of NORTH CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[NORTH CHINA] Inv Amt Trend_FY to FY:")

             df_Single_north = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "NORTH"').query('FY_Contract != "TBA"').round(0).groupby(by=["FY_Contract", "Contract_Month"],
                                as_index=False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Contract)" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_north = df_Single_north.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_north['FY_Contract'].unique(), all_fq_invoice_values],
                               names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
             fig7 = go.Figure()

# 添加每个FY_Contract的折线
             FY_Contract_values = df_Single_north['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_north[df_Single_north['FY_Contract'] == FY_Contract]
               fig7.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig7.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='khaki')
              
             fig7.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig7.update_layout(yaxis_showticklabels = True), use_container_width=True)

#NORTH Region Contract Details FQ_FQ:
             pvt10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "NORTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)

             html62 = pvt10.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html63 = html62.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html64 = html63.replace('<th>Q1</th>', '<th style="background-color: khaki">Q1</th>')
             html65 = html64.replace('<th>Q2</th>', '<th style="background-color: khaki">Q2</th>')
             html66 = html65.replace('<th>Q3</th>', '<th style="background-color: khaki">Q3</th>')
             html67 = html66.replace('<th>Q4</th>', '<th style="background-color: khaki">Q4</th>')
             html68 = html67.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
             html699 = f'<div style="zoom: 0.7;">{html68}</div>'
             st.markdown(html699, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv9 = pvt10.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv9, file_name='North_Sales.csv', mime='text/csv')

             ##st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             
##################################################
        with four_column:
# LINE CHART of WEST CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[WEST CHINA] Inv Amt Trend_FY to FY:")
             df_Single_west = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "WEST"').query('FY_Contract != "TBA"').round(0).groupby(by=["FY_Contract", "Contract_Month"],
                                as_index=False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Contract)" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_west = df_Single_west.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_west['FY_Contract'].unique(), all_fq_invoice_values],
                               names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
             fig6 = go.Figure()

# 添加每个FY_Contract的折线
             FY_Contract_values = df_Single_west['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_west[df_Single_west['FY_Contract'] == FY_Contract]
               fig6.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Before tax Inv Amt (HKD)'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig6.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),
                          paper_bgcolor='lightgreen')
              
             fig6.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig6.update_layout(yaxis_showticklabels = True), use_container_width=True)
#WEST Region Contract Details FQ_FQ:
             pvt18 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "WEST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)
            
             html69 = pvt18.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html70 = html69.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html71 = html70.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
             html72 = html71.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
             html73 = html72.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
             html74 = html73.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
             html75 = html74.replace('<th>Total</th>', '<th style="background-color: lightgreen">Total</th>')
# 放大pivot table
             html766 = f'<div style="zoom: 0.7;">{html75}</div>'

             st.markdown(html766, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv10 = pvt18.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv10, file_name='West_Sales.csv', mime='text/csv')
             ##st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)

        st.subheader(":sunrise: FY Contract Details_:orange[Monthly]:")
# 計算"FQ(Contract)"的subtotal數值
#              filter_df = df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"')
        pvt = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(
              values="Before tax Inv Amt (HKD)",
              index=["FY_Contract","FQ(Contract)","Contract_Yr","Contract_Month"],
              columns=["Region"],
              aggfunc="sum",
              fill_value=0,
              margins=True,
              margins_name="Total",
              observed=True)  # This ensures subtotals are only calculated for existing values)
                    
        desired_order = ["SOUTH", "EAST", "NORTH", "WEST","Total"]
        pvt = pvt.reindex(columns=desired_order, level=1)

# st.dataframe(pvt.style.highlight_max(color='yellow', axis=0).highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
# 使用applymap方法应用格式化
        pvt = pvt.applymap('HKD{:,.0f}'.format)
        html24 = pvt.to_html(classes='table table-bordered', justify='center')
# 放大pivot table
        html24 = f'<div style="zoom: 1.1;">{html24}</div>'
# 將你想要變色的column header找出來，並加上顏色
        html25 = html24.replace('<th>SOUTH</th>', '<th style="background-color: orange">SOUTH</th>')
        html26 = html25.replace('<th>EAST</th>', '<th style="background-color: lightblue">EAST</th>')
        html27 = html26.replace('<th>NORTH</th>', '<th style="background-color: Khaki">NORTH</th>')
        html28 = html27.replace('<th>WEST</th>', '<th style="background-color: lightgreen">WEST</th>')
# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
        html29 = html28.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把每個數值置中
        html30 = html29.replace('<td>', '<td style="text-align: middle;">')
# 把total值的那列的字設為黃色
        html31 = html30.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 把所有數值等於或少於0的數值的顏色設為紅色
        html32 = html31.replace('<td>-', '<td style="color: red;">-')
# 使用Streamlit的markdown來顯示HTML表格
        st.markdown(html32, unsafe_allow_html=True)
             #st.components.v1.html(html10, height=1000)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
        csv4 = pvt.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
        st.download_button(label='Download Table', data=csv4, file_name='Regional_Sales.csv', mime='text/csv')
################################################################################
#Regional inv amount subtotal FQ
        st.subheader(":point_down: Contract Amount Subtotal_:orange[FQ]:clipboard:")
        # 定义CSS样式

 
        with st.expander(":point_right: Click to expand/hide"):
              pvt7 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(
                     values="Before tax Inv Amt (HKD)",
                     index=["FY_Contract","FQ(Contract)"],
                     columns=["Region"],
                     aggfunc="sum",
                     fill_value=0,
                     margins=True,
                     margins_name="Total",
                     observed=True)  # This ensures subtotals are only calculated for existing values)
              desired_order = ["SOUTH", "EAST", "NORTH", "WEST","Total"]
              pvt7 = pvt7.reindex(columns=desired_order)

# 计算小计行
              subtotal_row = pvt7.groupby(level=0).sum(numeric_only=True)
              subtotal_row.index = pd.MultiIndex.from_product([subtotal_row.index, [""]])
              subtotal_row.name = ("Subtotal", "")  # 小计行索引的名称

# 去除千位數符號並轉換為浮點數
              pvt7 = pvt7.applymap(lambda x: float(str(x).strip('HKD').replace(',', '')))

# 转换为字符串并添加样式
              pvt7 = pvt7.applymap(lambda x: "HKD{:,.0f}".format(x))

# 将小计行与pvt7连接，使用concat函数
              pvt7_concatenated = pd.concat([pvt7, subtotal_row])

# 生成HTML表格
              html_table = pvt7_concatenated.to_html(classes='table table-bordered', justify='center')

 
# 使用BeautifulSoup处理HTML表格
              soup = BeautifulSoup(html_table, 'html.parser')

# 找到所有的<td>标签，并为小于或等于0的值添加CSS样式
              for td in soup.find_all('td'):
                    value = float(td.text.replace('HKD', '').replace(',', ''))
              if value <= 0:
                   td['style'] = 'color: red;'
      
# 找到所有的<td>标签，并将数值转换为会计数字格式的字符串
              for td in soup.find_all('td'):
                    value = float(td.text.strip('HKD').replace(',', ''))
                    formatted_value = "HKD{:,.0f}".format(value)
                    td.string.replace_with(formatted_value)
# 找到最底部的<tr>标签，并为其添加CSS样式
              last_row = soup.find_all('tr')[-1]
              last_row['style'] = 'background-color: yellow; font-weight: bold;'

# 在特定单元格应用其他样式           
              soup2 = str(soup)
              soup2 = soup2.replace('<th>SOUTH</th>', '<th style="background-color: orange">SOUTH</th>')
              soup2 = soup2.replace('<th>EAST</th>', '<th style="background-color: lightblue">EAST</th>')
              soup2 = soup2.replace('<th>NORTH</th>', '<th style="background-color: Khaki">NORTH</th>')
              soup2 = soup2.replace('<th>WEST</th>', '<th style="background-color: lightgreen">WEST</th>')
              soup2 = soup2.replace('<td>', '<td style="text-align: middle;">')
              soup2 = soup2.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
              soup2 = soup2.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
              soup2 = soup2.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
              soup2 = soup2.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
              soup2 = soup2.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
              html_with_style2 = str(f'<div style="zoom: 1.2;">{soup2}</div>')
              st.markdown(html_with_style2, unsafe_allow_html=True)       
###########################################################################################################################      
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv5 = pvt7.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv5, file_name='Regional_Quarter_Sales.csv', mime='text/csv') 
############################################################################################################################################ 


#TAB 3: Brand category
with tab3:

       tab4_row2_col1, tab4_row2_col2= st.columns(2)      
       with tab4_row2_col1:
             st.subheader(":sports_medal: 主要品牌 Contract 台数_:orange[FY]:")
             brandinv_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND != "SOLDERSTAR"').query(
                      'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHIMADZU"').query(
                        'BRAND != "OTHERS"').query('BRAND != "SAKI"').query('BRAND != "SAKI"').query('BRAND != "NUTEK"').query(
                        'BRAND != "DEK"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').round(0).groupby(by=["FY_Contract","BRAND_Details"],
                            as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
        # 按照指定顺序排序
#             brandinv_df["BRAND"] = pd.Categorical(brandinv_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER"])
#             brandinv_df = brandinv_df.sort_values("BRAND")
             df_brand = px.bar(brandinv_df, x="FY_Contract", y="Item Qty", color="BRAND_Details", text_auto='.3s')

# Update the traces to display the text above the bars
#             df_brand.update_traces(textposition='inside')

# 更改顏色
             colors = {"PEMTRON": "lightblue","HELLER": "orange","YAMAHA_Mounter": "lightgreen","YAMAHA_Non_Mounter": "khaki",}
             for trace in df_brand.data:
              region = trace.name.split("=")[-1]
              trace.marker.color = colors.get(region, "blue")

# 更改字體
             df_brand.update_layout(font=dict(family="Arial", size=18))
             df_brand.update_traces(marker_line_color='black', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
             df_brand.update_layout(barmode='group')

# 将图例放在底部
             df_brand.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

# 添加背景色
             background_color2 = 'lightgrey'
             x_range2 = len(brandinv_df["FY_Contract"].unique())
             background_shapes2 = [dict(
              type='rect',
              xref='x',
              yref='paper',
              x0=i - 0.5,
              y0=0,
              x1=i + 0.5,
              y1=1,
              fillcolor=background_color2,
              opacity=0.1,
              layer='below',
              line=dict(width=5)) for i in range(x_range2)]
             
             df_brand.update_layout(shapes=background_shapes2, showlegend=True)
# 绘制图表
             st.plotly_chart(df_brand, use_container_width=True)                              
###################################################################
       with tab4_row2_col2:
             st.subheader(":round_pushpin: 主要品牌 Contract 台数_:orange[百分比]:")

# 创建示例数据框
            
             brand_data = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by=["FY_Contract","BRAND_Details"],
                     as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
#             brand_data = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by=["FY_Contract","BRAND"],
#                     as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
            
             brandinvpie_df = pd.DataFrame(brand_data)

# 按照指定順序排序
             brandinvpie_df["BRAND_Details"] = brandinvpie_df["BRAND_Details"].replace(to_replace=[x for x in brandinvpie_df["BRAND_Details"
                                       ].unique() if x not in ["YAMAHA_Mounter","YAMAHA_Non_Mounter", "PEMTRON", "HELLER"]], value="OTHERS")
             brandinvpie_df["BRAND_Details"] = pd.Categorical(brandinvpie_df["BRAND_Details"], ["YAMAHA_Mounter","YAMAHA_Non_Mounter", "PEMTRON", "HELLER","OTHERS"])

#             brandinvpie_df["BRAND"] = brandinvpie_df["BRAND"].replace(to_replace=[x for x in brandinvpie_df["BRAND"
#                                       ].unique() if x not in ["YAMAHA", "PEMTRON", "HELLER"]], value="OTHERS")
#             brandinvpie_df["BRAND"] = pd.Categorical(brandinvpie_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER","OTHERS"])

# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Item Qty", names="BRAND_Details", color="BRAND_Details", color_discrete_map={
                      "PEMTRON": "lightblue", "HELLER": "orange", "YAMAHA_Mounter": "lightgreen","YAMAHA_Non_Mounter": "khaki", "OTHERS":"purple"})

#             df_pie = px.pie(brandinvpie_df, values="Item Qty", names="BRAND", color="BRAND", color_discrete_map={
#                      "PEMTRON": "lightblue", "HELLER": "orange", "YAMAHA": "lightgreen", "OTHERS":"purple"})

# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="v", yanchor="bottom", y=1.02, xanchor="right", x=1),
                   paper_bgcolor='lightgrey',  # 设置背景颜色为浅灰色
                   plot_bgcolor='lightgrey',   # 设置绘图区域背景颜色为浅灰色
                   margin=dict(l=10, r=10, t=10, b=30),  # 设置图表的边距
                   autosize=False,
                   width=600,
                   height=400)

# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line=dict(color='black', width=2))

# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
       
############################################################################################################################################
       one_column, two_column= st.columns(2)
       with one_column:
# LINE CHART of YAMAHA MOUNTER
              ##st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)
              st.subheader(":chart_with_upwards_trend: :green[YAMAHA Mounter] Inv Qty_:orange[Monthly]:")
              df_Single_south = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND_Details == "YAMAHA_Mounter"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract",
                                 "Contract_Month"], as_index= False)["Item Qty"].sum()
# 确保 "Contract_Month" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_south = df_Single_south.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_south['FY_Contract'].unique(), all_fq_invoice_values],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个Contract_Month的折线
              FY_Contract_values = df_Single_south['FY_Contract'].unique()
              for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_south[df_Single_south['FY_Contract'] == FY_Contract]
               fig4.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Item Qty'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Item Qty'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='lightgreen')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)

             
#YAMAHA MOUNTER Contract Details FQ_FQ:
              pvt8 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND_Details == "YAMAHA_Mounter"').round(0).pivot_table(values="Item Qty",
                     index=['FY_Contract'],columns=["FQ(Contract)"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)
            
              html76 = pvt8.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: lightgreen">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: lightgreen">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: lightgreen">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: lightgreen">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv_mounter = pvt8.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv_mounter, file_name='YAMAHA_Mounter_Sales.csv', mime='text/csv')
              ##st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)

       with two_column:
# LINE CHART of YAMAHA_Non_Mounter
              #st.divider()
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)

              st.subheader(":chart_with_upwards_trend: :blue[YAMAHA Non-Mounter] Inv Qty Trend_:orange[Monthly]:")
              df_Single_south = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND_Details == "YAMAHA_Non_Mounter"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract",
                                 "Contract_Month"], as_index= False)["Item Qty"].sum()
# 确保 "Contract_Month" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_south = df_Single_south.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_south['FY_Contract'].unique(), all_fq_invoice_values],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个Contract_Month的折线
              FY_Contract_values = df_Single_south['FY_Contract'].unique()
              for FY_Contract in FY_Contract_values:
               FY_Contract_data = df_Single_south[df_Single_south['FY_Contract'] == FY_Contract]
               fig4.add_trace(go.Scatter(
                          x=FY_Contract_data['Contract_Month'],
                          y=FY_Contract_data['Item Qty'],
                          mode='lines+markers+text',
                          name=FY_Contract,
                          text=FY_Contract_data['Item Qty'],
                          textposition="bottom center",
                          texttemplate='%{text:.3s}',
                          hovertemplate='%{x}<br>%{y:.2f}',
                          marker=dict(size=10)))
               fig4.update_layout(xaxis=dict(
                          type='category',
                          categoryorder='array',
                          categoryarray=all_fq_invoice_values),
                          yaxis=dict(showticklabels=True),
                          font=dict(family="Arial, Arial", size=12, color="Black"),
                          hovermode='x', showlegend=True,
                          legend=dict(orientation="h",font=dict(size=14)),paper_bgcolor='khaki')
             
              fig4.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

              st.plotly_chart(fig4.update_layout(yaxis_showticklabels = True), use_container_width=True)
         
#YAMAHA NON-MOUNTER Contract QTY Details MONTHLY:
              pvt8 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND_Details == "YAMAHA_Non_Mounter"').round(0).pivot_table(values="Item Qty",
                     index=['FY_Contract'],columns=["FQ(Contract)"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_Contract',ascending=True)
            
              html76 = pvt8.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
              html77 = html76.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
              html78 = html77.replace('<th>Q1</th>', '<th style="background-color: khaki">Q1</th>')
              html79 = html78.replace('<th>Q2</th>', '<th style="background-color: khaki">Q2</th>')
              html80 = html79.replace('<th>Q3</th>', '<th style="background-color: khaki">Q3</th>')
              html81 = html80.replace('<th>Q4</th>', '<th style="background-color: khaki">Q4</th>')
              html822 = html81.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
              html833 = f'<div style="zoom: 0.7;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv_non_mounter = pvt8.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv_non_mounter, file_name='YAMAHA_Non_Mounter.csv', mime='text/csv')
              #st.divider() 
              st.markdown(
                   """
                   <hr style="border: 3px solid lightblue;">
                   """,
                   unsafe_allow_html=True)         
###########################################################################################################
      # BAR CHART of BRAND Comparision
       st.subheader(":bar_chart: 主要品牌 Contract 台数_:orange[Quarterly]:")
       brand_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND != "SOLDERSTAR"').query(
                      'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHIMADZU"').query(
                        'BRAND != "OTHERS"').query('BRAND != "SAKI"').query('BRAND != "SAKI"').query('BRAND != "NUTEK"').query(
                        'BRAND != "DEK"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').round(0).groupby(by=["FQ(Contract)",
                            "BRAND_Details"], as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
        # 按照指定顺序排序
#       brand_df["BRAND"] = pd.Categorical(brand_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER"])
#       brand_df = brand_df.sort_values("BRAND")
       # 创建一个包含排序顺序的列表
       sort_FQ_order = ["Q1", "Q2", "Q3", "Q4"]

# 使用plotly绘制柱状图           
       brand_qty = px.bar(brand_df, x="FQ(Contract)", y="Item Qty", color="BRAND_Details", text_auto='.3s')

# 设置x轴的分类顺序
       brand_qty.update_layout(xaxis={"type": "category", "categoryorder": "array", "categoryarray": sort_FQ_order})

# 更改顏色
       colors = {"YAMAHA_Mounter": "lightgreen","YAMAHA_Non_Mounter": "khaki", "PEMTRON": "lightblue","HELLER": "orange"}
       for trace in brand_qty.data:
              brand_color = trace.name.split("=")[-1]
              trace.marker.color = colors.get(brand_color, "blue")

# 更改字體和label
       brand_qty.update_layout(font=dict(family="Arial", size=13.5, color="black"))
       brand_qty.update_traces(marker_line_color='black', textposition='outside', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
       brand_qty.update_layout(barmode='group')

# 将图例放在底部
       brand_qty.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))


# 添加背景色
       background_color = 'lightgrey'
       x_range = len(brand_df['FQ(Contract)'].unique())
       background_shapes = [dict(
              type='rect',
              xref='x',
              yref='paper',
              x0=i - 0.5,
              y0=0,
              x1=i + 0.5,
              y1=1,
              fillcolor=background_color,
              opacity=0.1,
              layer='below',
              line=dict(width=5)) for i in range(x_range)]

# 绘制图表
       brand_qty.update_layout(shapes=background_shapes, showlegend=True)
       st.plotly_chart(brand_qty, use_container_width=True)

# 绘制图表
#       st.plotly_chart(brand_qty, use_container_width=True)
########################################################################################################
       st.subheader(":point_down: 主要品牌 Contract 台数_:orange[FQ Subtotal]:clipboard:")          
       with st.expander(":point_right: click to expand/hide data"):
#Brand Inv Qty by Inv Month:
             #with st.expander("Click to expand"):
             pvt6 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND != "SOLDERSTAR"').query(
                      'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHIMADZU"').query(
                        'BRAND != "OTHERS"').query('BRAND != "SAKI"').query('BRAND != "SAKI"').query('BRAND != "NUTEK"').query(
                        'BRAND != "DEK"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').round(0).pivot_table(
                    values="Item Qty",
                    index=["FY_Contract","FQ(Contract)"],
                    columns=["BRAND_Details"],
                    aggfunc="sum",
                    fill_value=0,
                    margins=True,
                    margins_name="Total",
                    observed=True)  # This ensures subtotals are only calculated for existing values)
            
             desired_order_brand = ["YAMAHA_Mounter","YAMAHA_Non_Mounter", "PEMTRON", "HELLER","Total"]
             pvt6 = pvt6.reindex(columns=desired_order_brand)

# 计算小计行
             subtotal_row = pvt6.groupby(level=0).sum(numeric_only=True)
             subtotal_row.index = pd.MultiIndex.from_product([subtotal_row.index, [""]])
             subtotal_row.name = ("Subtotal", "")  # 小计行索引的名称

# 去除千位數符號並轉換為浮點數
             pvt6 = pvt6.applymap(lambda x: float(str(x).strip('HKD').replace(',', '')))

# 将小计行与pvt7连接，使用concat函数
             pvt6_concatenated = pd.concat([pvt6, subtotal_row])

# 生成HTML表格
             html_table = pvt6_concatenated.to_html(classes='table table-bordered', justify='center')

# 使用BeautifulSoup处理HTML表格
             soup = BeautifulSoup(html_table, 'html.parser')

# 找到所有的<td>标签，并为小于或等于0的值添加CSS样式
             for td in soup.find_all('td'):
                   if value <= 0:
                         td['style'] = 'color: red;'    
# 找到最底部的<tr>标签，并为其添加CSS样式
             last_row = soup.find_all('tr')[-1]
             last_row['style'] = 'background-color: yellow; font-weight: bold;'

# 在特定单元格应用其他样式           
             soup3 = str(soup)
             soup3 = soup3.replace('<th>YAMAHA_Mounter</th>', '<th style="background-color: lightgreen">YAMAHA_Mounter</th>')
             soup3 = soup3.replace('<th>PEMTRON</th>', '<th style="background-color: lightblue">PEMTRON</th>')
             soup3 = soup3.replace('<th>HELLER</th>', '<th style="background-color: orange">HELLER</th>')
             soup3 = soup3.replace('<th>YAMAHA_Non_Mounter</th>', '<th style="background-color: khaki">YAMAHA_Non_Mounter</th>')
             soup3 = soup3.replace('<td>', '<td style="text-align: middle;">')
             soup3 = soup3.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
             soup3 = soup3.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
             soup3 = soup3.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
             soup3 = soup3.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
             soup3 = soup3.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
             html_with_style3 = str(f'<div style="zoom: 1.3;">{soup3}</div>')
             st.markdown(html_with_style3, unsafe_allow_html=True)       
      
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv13 = pvt6.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv13, file_name='Brand_invoice_qty.csv', mime='text/csv')

       #st.divider()
       st.markdown(
            """
            <hr style="border: 3px solid lightblue;">
            """,
            unsafe_allow_html=True
            )
############################################################################################################################################
#Top Product line chart Contract qty trend
       
       left_column, right_column = st.columns(2)
       with left_column:
             st.header(":chart_with_upwards_trend: Main Unit Contract Qty_:orange[Monthly]:")
###########################################################################################################
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM20R Contract Details:
       with left_column:
             
             st.subheader(":red[YAMAHA:] YSM20R")
             df_YSM20 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM20 = df_YSM20.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM20['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig8 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YSM20['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YSM20[df_YSM20['FY_Contract'] == FY_Contract]
                  fig8.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig8.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig8.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig8.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig8.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY YSM20R Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt13 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query(
                   'Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True, margins_name="Total").sort_index(axis=0, ascending=True)
            
             html90 = pvt13.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html91 = html90.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html92 = html91.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html93 = html92.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html94 = html93.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html95 = html94.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html96 = html95.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html97 = f'<div style="zoom: 1.2;">{html96}</div>'
             st.markdown(html97, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv14 = pvt13.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv14, file_name='YSM20 Invocie Qty.csv', mime='text/csv')
##########################################################################################
#Line Chart FY to FY YSM10 Contract Details:
       with right_column:
             st.subheader(":red[YAMAHA:] YSM10")
             df_YSM10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM10 = df_YSM10.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM10['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig9 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YSM10['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YSM10[df_YSM10['FY_Contract'] == FY_Contract]
                  fig9.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig9.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig9.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig9.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig9.update_layout(yaxis_showticklabels = True), use_container_width=True)
#FY to FY YSM10 Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt14 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query(
                   'Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
            
             html98 = pvt14.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html99 = html98.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html100 = html99.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html101 = html100.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html102 = html101.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html103 = html102.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html104 = html103.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html105 = f'<div style="zoom: 1.2;">{html104}</div>'
             st.markdown(html105, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv14 = pvt14.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv14, file_name='YSM10 Invocie Qty.csv', mime='text/csv')
            
#############################################################################################################################################################
#Second row for Top product trend
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM40R Contract Details:
       with left_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True
                  )
             st.subheader(":red[YAMAHA:] YSM40R")
             df_YSM40 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM40R"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM40 = df_YSM40.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM40['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig10 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YSM40['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YSM40[df_YSM40['FY_Contract'] == FY_Contract]
                  fig10.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig10.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig10.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig10.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig10.update_layout(yaxis_showticklabels = True), use_container_width=True)
#FY to FY YSM40R Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt14 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YSM40R"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
             html114 = pvt14.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html115 = html114.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html116 = html115.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html117 = html116.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html118 = html117.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html119 = html118.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html120 = html119.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html121 = f'<div style="zoom: 1.2;">{html120}</div>'
             st.markdown(html121, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv16 = pvt14.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv16, file_name='YSM40 Invocie Qty.csv', mime='text/csv')            
###############################################################################################################################
#Line Chart FY to FY YSi-V Contract Details:
       with right_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             st.subheader(":red[YAMAHA:] YSi-V(AOI)")
             df_YSIV = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSi-V"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSIV = df_YSIV.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSIV['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig11 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YSIV['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YSIV[df_YSIV['FY_Contract'] == FY_Contract]
                  fig11.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig11.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig11.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig11.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig11.update_layout(yaxis_showticklabels = True), use_container_width=True)
#FY to FY YSi-V Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt15 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query(
                   'Ordered_Items == "YSi-V"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
            
             html106 = pvt15.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html107 = html106.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html108 = html107.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html109 = html108.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html110 = html109.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html111 = html110.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html112 = html111.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html113 = f'<div style="zoom: 1.2;">{html112}</div>'
             st.markdown(html113, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv15 = pvt15.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv15, file_name='YSi-V(AOI) Invocie Qty.csv', mime='text/csv')
##############################################################################################################
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YRM10 Contract Details:
       with left_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             st.subheader(":red[YAMAHA:] YRM10")
             df_YRM10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YRM10"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YRM10 = df_YRM10.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YRM10['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig13 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YRM10['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YRM10[df_YRM10['FY_Contract'] == FY_Contract]
                  fig13.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig13.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig13.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig13.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig13.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY YRM10 Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt20 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YRM10"').pivot_table(values="Item Qty",
                     index=["FY_Contract"],columns=["FQ(Contract)"], aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
             html122 = pvt20.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html123 = html122.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html124 = html123.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html125 = html124.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html126 = html125.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html127 = html126.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html128 = html127.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html129 = f'<div style="zoom: 1.2;">{html128}</div>'
             st.markdown(html129, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv16 = pvt20.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv16, file_name='YRM10 Invocie Qty.csv', mime='text/csv')  
       
       #Line Chart FY to FY YRM20 Contract Details:
       with right_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             st.subheader(":red[YAMAHA:] YRM20")
             df_YRM20 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YRM20"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YRM20 = df_YRM20.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YRM20['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig13 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YRM20['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YRM20[df_YRM20['FY_Contract'] == FY_Contract]
                  fig13.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig13.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig13.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig13.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig13.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY YRM20 Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt20 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YRM20"').pivot_table(values="Item Qty",
                     index=["FY_Contract"],columns=["FQ(Contract)"], aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
             html122 = pvt20.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html123 = html122.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html124 = html123.replace('<th>Q1</th>', '<th style="background-color: pink">Q1</th>')
             html125 = html124.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html126 = html125.replace('<th>Q3</th>', '<th style="background-color: pink">Q3</th>')
             html127 = html126.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html128 = html127.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html129 = f'<div style="zoom: 1.2;">{html128}</div>'
             st.markdown(html129, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv16 = pvt20.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv16, file_name='YRM20 Invocie Qty.csv', mime='text/csv')          
#############################################################################################################################    
       left_column, right_column = st.columns(2) 
#Line Chart FY to FY PEMTRON Contract Details:
       with left_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             st.subheader(":blue[PEMTRON:]")
             df_PEMTRON = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND == "PEMTRON"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_PEMTRON = df_PEMTRON.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_PEMTRON['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig14 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_PEMTRON['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_PEMTRON[df_PEMTRON['FY_Contract'] == FY_Contract]
                  fig14.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig14.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig14.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),paper_bgcolor='lightblue')
             fig14.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig14.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY PEMTRON Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt17 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('BRAND == "PEMTRON"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
             html130 = pvt17.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html131 = html130.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html132 = html131.replace('<th>Q1</th>', '<th style="background-color: lightblue">Q1</th>')
             html133 = html132.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html134 = html133.replace('<th>Q3</th>', '<th style="background-color: lightblue">Q3</th>')
             html135 = html134.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html136 = html135.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html137 = f'<div style="zoom: 1.2;">{html136}</div>'
             st.markdown(html137, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv17 = pvt17.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv17, file_name='PEMTRON Main Units Invocie Qty.csv', mime='text/csv') 
########################################################################################################################
#Line Chart FY to FY HELLER Contract Details:
       with right_column:
             #st.divider()
             st.markdown(
                  """
                  <hr style="border: 3px solid lightblue;">
                  """,
                  unsafe_allow_html=True)
             st.subheader(":orange[HELLER:]")
             df_YRM10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('BRAND == "HELLER"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","Contract_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YRM10 = df_YRM10.groupby(["FY_Contract", "Contract_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YRM10['FY_Contract'].unique(), sort_Month_order],
                                   names=['FY_Contract', 'Contract_Month'])).fillna(0).reset_index()
#建立圖表         
             fig15 = go.Figure()
# 添加每个FY_Contract的折线
             FY_Contract_values = df_YRM10['FY_Contract'].unique()
             for FY_Contract in FY_Contract_values:
                  FY_Contract_data = df_YRM10[df_YRM10['FY_Contract'] == FY_Contract]
                  fig15.add_trace(go.Scatter(
                       x=FY_Contract_data['Contract_Month'],
                       y=FY_Contract_data["Item Qty"],
                       mode='lines+markers+text',
                       name=FY_Contract,
                       text=FY_Contract_data["Item Qty"],
                       textposition="bottom center",
                       texttemplate='%{text:.3s}',
                       hovertemplate='%{x}<br>%{y:.2f}',
                       marker=dict(size=10)))
                  fig15.update_layout(xaxis=dict(
                       type='category',
                       categoryorder='array',
                       categoryarray=sort_Month_order),
                       yaxis=dict(showticklabels=True),
                       font=dict(family="Arial, Arial", size=13, color="Black"),
                       hovermode='x', showlegend=True,
                       legend=dict(orientation="h",font=dict(size=14)))
                 
             fig15.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),paper_bgcolor='rgba(255,165,0,0.3)')
             fig15.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig15.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY HELLER Contract Details:
             filter_df["FQ(Contract)"] = pd.Categorical(filter_df["FQ(Contract)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt18 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('BRAND == "HELLER"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
            
             html138 = pvt18.applymap('{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html139 = html138.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html140 = html139.replace('<th>Q1</th>', '<th style="background-color: orange">Q1</th>')
             html141 = html140.replace('<th>Q2</th>', '<th style="background-color: lightgrey">Q2</th>')
             html142 = html141.replace('<th>Q3</th>', '<th style="background-color: orange">Q3</th>')
             html143 = html142.replace('<th>Q4</th>', '<th style="background-color: lightgrey">Q4</th>')
             html144 = html143.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html145 = f'<div style="zoom: 1.2;">{html144}</div>'
             st.markdown(html145, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv18 = pvt18.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
            
             st.download_button(label='Download Table', data=csv18, file_name='HELLER Main Units Invocie Qty.csv', mime='text/csv') 
             
############################################################################################################################################################################################################
#TAB 4: Customer category
with tab4:
#Top Down Customer details Table
      left_column, right_column= st.columns(2)
#BAR CHART Customer List
      with left_column:
              st.subheader(":money_with_wings: :orange[十大客戶 合同金額]_排行图:")            
              customer_qty_line = (filter_df.query('BRAND != "C66 SERVICE"').query('Contract_Yr != "TBA"').query(
                                 'Contract_Month != "TBA"').query('Contract_Month != "Cancel"').groupby(
                                 by=["Customer_Name"])[["Before tax Inv Amt (HKD)"]].sum().sort_values(
                                 by="Before tax Inv Amt (HKD)", ascending=False).head(10))
# 生成颜色梯度
              colors = px.colors.sequential.Blues[::-1]  # 将颜色顺序反转为从深到浅
# 创建条形图
              fig_customer_inv_qty = px.bar(
                  customer_qty_line,
                  x="Before tax Inv Amt (HKD)",
                  y=customer_qty_line.index,
                  text="Before tax Inv Amt (HKD)",
                  orientation="h",
                  color=customer_qty_line.index,
                  color_discrete_sequence=colors[:len(customer_qty_line)],
                  template="plotly_white", text_auto='.3s')

# 更新图表布局和样式
              fig_customer_inv_qty.update_layout(
                  height=600,
                  yaxis=dict(title="Customer_Name"),
                  xaxis=dict(title="Before tax Inv Amt (HKD)"),)
              fig_customer_inv_qty.update_layout(font=dict(family="Arial", size=20))
              fig_customer_inv_qty.update_traces(
                  textposition="inside",
                  marker_line_color="black",
                  marker_line_width=2,
                  opacity=1,showlegend=False,
                  )
             # 显示图表
              st.plotly_chart(fig_customer_inv_qty, use_container_width=True)

##################################################    
      with right_column:
              st.subheader(":medal: :orange[十大客戶 合同金額]_明細:")
    # Create a pivot table with the required filters
              pvt21 = filter_df.query('BRAND != "C66 SERVICE"').query('Contract_Yr != "TBA"').query(
                     'Contract_Month not in ["TBA", "Cancel"]').round(0).pivot_table(
                            values=["Item Qty", "Before tax Inv Amt (HKD)"],
                            index=["Customer_Name", "Region", "BRAND", "Ordered_Items"],
                            aggfunc={"Item Qty": "sum", "Before tax Inv Amt (HKD)": "sum"},
                            fill_value=0)

# Concatenate and sort
              pvt21 = pd.concat([pvt21]).sort_values(by=("Before tax Inv Amt (HKD)"), ascending=False)
              customer_totals = pvt21.groupby(level='Customer_Name').sum()

# Select top 10 customers based on "Before tax Inv Amt (HKD)"
              top_customers = customer_totals.nlargest(10, "Before tax Inv Amt (HKD)")

# Filter original pivot table to include only top customers
              pvt21_top10 = pvt21.loc[top_customers.index]

# Format the "Before tax Inv Amt (HKD)"
              pvt21_top10["Before tax Inv Amt (HKD)"] = pvt21_top10["Before tax Inv Amt (HKD)"].apply(lambda x: f"HKD {x:,.0f}")

# Render as HTML
              st.markdown(pvt21_top10.to_html(), unsafe_allow_html=True)
      
      



# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv19 = pvt21_top10.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv19, file_name='Top 10 Invoiced Projects.csv', mime='text/csv')  


       
      row2_left_column, row2_right_column= st.columns(2)
      with row2_left_column:
              st.subheader(":radio: Top 10 Customer_:blue[Contract Qty]:")            
              customer_line = (filter_df.query('BRAND != "C66 SERVICE"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query(
                             'Contract_Month != "Cancel"').query('BRAND != "C66 SERVICE"').query(
                             'BRAND != "LOCAL SUPPLIER"').query('BRAND != "SIGMA"').groupby(
                             by=["Customer_Name"])[["Item Qty"]].sum().sort_values(by="Item Qty", ascending=False).head(10))
# 生成颜色梯度
              colors = px.colors.sequential.Greens[::-1]  # 将颜色顺序反转为从深到浅
# 创建条形图
              fig_customer = px.bar(
                  customer_line,
                  x="Item Qty",
                  y=customer_line.index,
                  text="Item Qty",
                  orientation="h",
                  color=customer_line.index,
                  color_discrete_sequence=colors[:len(customer_line)],
                  template="plotly_white", text_auto='.3s')

# 更新图表布局和样式
              fig_customer.update_layout(
                  height=400,
                  yaxis=dict(title="Customer_Name"),
                  xaxis=dict(title="Item Qty"),)
              fig_customer.update_layout(font=dict(family="Arial", size=15))
              fig_customer.update_traces(
                  textposition="inside",
                  marker_line_color="black",
                  marker_line_width=2,
                  opacity=1,showlegend=False,
                  )
             # 显示图表
              st.plotly_chart(fig_customer, use_container_width=True)

      with row2_right_column:
             st.subheader(":trophy: :orange[Top Customer List]_Inv Amount& Qty:")
             pvt11 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(2).pivot_table(index=["Customer_Name","Region"],
                    values=["Item Qty","Before tax Inv Amt (HKD)"],
             aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
             st.dataframe(pvt11.style.format("{:,}"), use_container_width=True)
#############################################################################################################################################################################################################
with tab5:      
      tab5_row1_col1, tab5_row1_col2= st.columns(2)      
      with tab5_row1_col1:
             st.subheader(":hourglass_flowing_sand: INV Leadtime Qty_:orange[FY]:")
             brandinv_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query(
                      'BRAND != "C66 SERVICE"').round(0).groupby(by=["FY_Contract","Inv_LeadTime_MonthGroup"],
                            as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
        # 按照指定顺序排序
#             brandinv_df["BRAND"] = pd.Categorical(brandinv_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER"])
#             brandinv_df = brandinv_df.sort_values("BRAND")
             df_brand = px.bar(brandinv_df, x="FY_Contract", y="Item Qty", color="Inv_LeadTime_MonthGroup", text_auto='.3s')

# Update the traces to display the text above the bars
#             df_brand.update_traces(textposition='inside')

# 更改顏色
             colors = {"> 6 months": "red","0-3 months": "green","4-6 months": "khaki",}
             for trace in df_brand.data:
              region = trace.name.split("=")[-1]
              trace.marker.color = colors.get(region, "blue")

# 更改字體
             df_brand.update_layout(font=dict(family="Arial", size=18))
             df_brand.update_traces(marker_line_color='black', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
             df_brand.update_layout(barmode='group')

# 将图例放在底部
             df_brand.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

# 添加背景色
             background_color2 = 'lightgrey'
             x_range2 = len(brandinv_df["FY_Contract"].unique())
             background_shapes2 = [dict(
              type='rect',
              xref='x',
              yref='paper',
              x0=i - 0.5,
              y0=0,
              x1=i + 0.5,
              y1=1,
              fillcolor=background_color2,
              opacity=0.1,
              layer='below',
              line=dict(width=5)) for i in range(x_range2)]
             
             df_brand.update_layout(shapes=background_shapes2, showlegend=True)
# 绘制图表
             st.plotly_chart(df_brand, use_container_width=True) 
###############################################################################################            
      with tab5_row1_col2:
             st.subheader(":round_pushpin: 主要品牌 Contract 台数_:orange[百分比]:")

# 创建示例数据框
             brand_data = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query(
                      'BRAND != "C66 SERVICE"').round(0).groupby(by=["FY_Contract","Inv_LeadTime_MonthGroup"],
                     as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
      
             brandinvpie_df = pd.DataFrame(brand_data)

# 按照指定順序排序
             brandinvpie_df["Inv_LeadTime_MonthGroup"] = brandinvpie_df["Inv_LeadTime_MonthGroup"].replace(to_replace=[x for x in brandinvpie_df["Inv_LeadTime_MonthGroup"
                                       ].unique() if x not in ["> 6 months","0-3 months","4-6 months"]], value="OTHERS")
             brandinvpie_df["Inv_LeadTime_MonthGroup"] = pd.Categorical(brandinvpie_df["Inv_LeadTime_MonthGroup"], ["> 6 months","0-3 months","4-6 months","OTHERS"])

# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Item Qty", names="Inv_LeadTime_MonthGroup", color="Inv_LeadTime_MonthGroup", color_discrete_map={
                      "> 6 months": "red","0-3 months": "green","4-6 months": "khaki"})

#             df_pie = px.pie(brandinvpie_df, values="Item Qty", names="BRAND", color="BRAND", color_discrete_map={
#                      "PEMTRON": "lightblue", "HELLER": "orange", "YAMAHA": "lightgreen", "OTHERS":"purple"})

# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="v", yanchor="bottom", y=1.02, xanchor="right", x=1),
                   paper_bgcolor='lightgrey',  # 设置背景颜色为浅灰色
                   plot_bgcolor='lightgrey',   # 设置绘图区域背景颜色为浅灰色
                   margin=dict(l=10, r=10, t=10, b=30),  # 设置图表的边距
                   autosize=False,
                   width=600,
                   height=400)

# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line=dict(color='black', width=2))

# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)

###############################################################################################  
#Table of > 6 months data
      st.subheader(":ledger: Contract Details_:orange[Monthly]:point_down::")
      filter_df["G.P. %"] = (filter_df["G.P. %"].astype(float) * 100).round(2).astype(str) + "%"
      pvt21 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Inv_LeadTime_MonthGroup == "> 6 months"').query('FY_Contract != "TBA"').round(0).pivot_table(
           values=["Item Qty","Before tax Inv Amt (HKD)","G.P.  (HKD)"],
           index=["Inv_LeadTime_MonthGroup","FY_Contract","Contract_Yr","Contract_Month","Contract_No.","Customer_Name","Ordered_Items","G.P. %"],
           aggfunc={"Item Qty":"sum","Before tax Inv Amt (HKD)": "sum", "G.P.  (HKD)": "sum"},
           fill_value=0, margins=True,margins_name="Total")
            
      # 按"Contract_Yr","Contract_Month"以大到小排序
      pvt21 = pvt21.sort_values(by=["Contract_Yr","Contract_Month"], ascending=False)

# 将"Before tax Inv Amt (HKD)"和"G.P. (HKD)"格式化为会计单位
      pvt21["Before tax Inv Amt (HKD)"] = "HKD " + pvt21["Before tax Inv Amt (HKD)"].astype(int).apply(lambda x: "{:,}".format(x))
      pvt21["G.P.  (HKD)"] = "HKD " + pvt21["G.P.  (HKD)"].astype(int).apply(lambda x: "{:,}".format(x))

# 将"Item Qty"前的"HKD"字眼去掉
      pvt21 = pvt21.rename(columns={"Item Qty": "Item Qty"})

# 调整values部分的显示顺序
      pvt21 = pvt21[["Item Qty", "Before tax Inv Amt (HKD)", "G.P.  (HKD)"]]
      html146 = pvt21.to_html(classes='table table-bordered', justify='center')
      html147 = html146.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
      html148 = html147.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
      html149 = f'<div style="zoom: 0.85;">{html148}</div>'
      st.markdown(html149, unsafe_allow_html=True)


# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
      csv19 = pvt21.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
      st.download_button(label='Download Table', data=csv19, file_name='Projects_INV Leadtime > 6 months.csv', mime='text/csv')      


############################################################################################################################################
#TAB 6 Invocie Details
with tab6:
   filter_df["GP%_of_month"] = (filter_df["GP%_of_month"] * 100).round(2).astype(str) + "%"
   # 定义CSS样式
   st.markdown(
             """
             <style>
             .st-expander > div:first-child {
             font-size: 100px;
             background-color: yellow;
             border: 3px solid black;
             padding: 10px;
             }
             </style>
             """,
             unsafe_allow_html=True)
   with st.expander(":point_right: 打开subtotal明细"): 
    
    st.subheader(":closed_book: Contract Amount Subtotal_:orange[Monthly]:point_down:: ")
#    with st.expander(":point_right: :closed_book: click to expand/ hide the tabe"):
    pvt2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(0).pivot_table(
              index=["FY_Contract","Contract_Yr","FQ(Contract)", "Contract_Month","GP%_of_month"],
              values=["Before tax Inv Amt (HKD)","G.P.  (HKD)"],
              aggfunc="sum",
              fill_value=0,
              margins=True,
              margins_name="Total",
              observed=True)

    pvt2 = pvt2.reindex(level=1)
# 按"Contract_Yr","Contract_Month"以大到小排序
    pvt2 = pvt2.sort_values(by=["FY_Contract","Contract_Yr","FQ(Contract)","Contract_Month"], ascending=False)

       #使用applymap方法應用格式化
    pvt2 = pvt2.applymap('{:,.0f}'.format)
    html3 = pvt2.to_html(classes='table table-bordered', justify='center')

# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
    html8 = html3.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把每個數值置中
    html9 = html8.replace('<td>', '<td style="text-align: middle;">')

# 把所有數值等於或少於0的數值的顏色設為紅色
    html14 = html9.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
    html15 = f'<div style="zoom: 1.1;">{html14}</div>'
    st.markdown(html15, unsafe_allow_html=True)           

# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
    csv1 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
    st.download_button(label='Download Table', data=csv1, file_name='SMT_Monthly_Sales.csv', mime='text/csv')
    #st.divider()
   st.markdown(
            """
            <hr style="border: 3px solid lightblue;">
            """,
            unsafe_allow_html=True
            )

###############################################      
       #FY to FY Quarter Contract Details:
   st.subheader(":ledger: Contract Details_:orange[Monthly]:point_down::")
   pvt21 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).pivot_table(
           values=["Item Qty","Before tax Inv Amt (HKD)","G.P.  (HKD)"],
           index=["FY_Contract","Contract_Yr","Contract_Month","Contract_No.","Customer_Name","Ordered_Items","G.P. %"],
           aggfunc={"Item Qty":"sum","Before tax Inv Amt (HKD)": "sum", "G.P.  (HKD)": "sum"},
           fill_value=0, margins=True,margins_name="Total")
            
      # 按"Contract_Yr","Contract_Month"以大到小排序
   pvt21 = pvt21.sort_values(by=["Contract_Yr","Contract_Month"], ascending=False)

# 将"Before tax Inv Amt (HKD)"和"G.P. (HKD)"格式化为会计单位
   pvt21["Before tax Inv Amt (HKD)"] = "HKD " + pvt21["Before tax Inv Amt (HKD)"].astype(int).apply(lambda x: "{:,}".format(x))
   pvt21["G.P.  (HKD)"] = "HKD " + pvt21["G.P.  (HKD)"].astype(int).apply(lambda x: "{:,}".format(x))

# 将"Item Qty"前的"HKD"字眼去掉
   pvt21 = pvt21.rename(columns={"Item Qty": "Item Qty"})

# 调整values部分的显示顺序
   pvt21 = pvt21[["Item Qty", "Before tax Inv Amt (HKD)", "G.P.  (HKD)"]]
    
   html146 = pvt21.to_html(classes='table table-bordered', justify='center')
   html147 = html146.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
   html148 = html147.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
   html149 = f'<div style="zoom: 0.9;">{html148}</div>'
   st.markdown(html149, unsafe_allow_html=True)


# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
   csv19 = pvt21.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
   st.download_button(label='Download Table', data=csv19, file_name='SMT_Monthly_Sales_Details.csv', mime='text/csv') 

###################################################################################################################   
 

############################################################################################################################################################################################################
#success info
#st.success("Executed successfully")
#st.info("This is an information")
#st.warning("This is a warning")
#st.error("An error occured")

 
 
 
