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
######################################################################################################
 
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")
#Title
st.title(':factory:  Mounter Import Data of China_Analysis')
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
#Text
st.write("by Arthur Chan")
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
 
######################################################################################################
#Create a browser for user to upload
 
#@st.cache_data
#def load_data(file):
#        data = pd.read_excel(file)
#        return data
 
#uploaded_file = st.sidebar.file_uploader(":file_folder: Upload monthly report here")
 
#if uploaded_file is not None:
#df = load_data(uploaded_file)
#st.dataframe(df)
#唔show 17/18, cancel, tba資料
#else:
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"C:\Users\ArthurChan\OneDrive\VS Code\PythonProject_ESE\Sample Excel")


df = pd.read_excel(
       io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='raw_sheet', skiprows=0, usecols='A:AR',
       nrows=10000,).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
 

df_import = pd.read_excel(
       io='Machine_Import_data.xlsx',engine= 'openpyxl',sheet_name='Mounter', skiprows=0, usecols='A:G',nrows=10000,)
 
######################################################################################################
# https://icons.getbootstrap.com/
#Top menu bar
 
#with st.sidebar:
#      selected = option_menu(
#              menu_title=None,
#              options=["Invoice Summary","Contract Summary","Mounter & Non-Mounter"],
#              icons=["database-fill-check","pen-fill","house-fill"],
#              menu_icon="chat-text-fill",
#              default_index=0,
#              orientation="vertical",
#              styles={
#                     "container":{"padding":"0!important","background-color": "navy"},
#                     "icon":{"color":"white", "font-size":"15px"},
#                     "nav-link":{"color":"white","font-size":"18px","text-align":"left","margin":"0px","--hover-color":"orange",},
#                     "nav-link-selected": {"background-color":"orange"},
#              })
 
#if selected == "Invoice Summary":
 
######################################################################################################
# Sidebar Slider
st.sidebar.header(":point_down: Slider for China import data:")
df_import["YEAR"] = df_import["YEAR"].astype(str)
 
#start_yr, end_yr
df_yr= st.sidebar.select_slider('Select a range of year',
            options=df_import["YEAR"].unique(), value=("2022","2023"))
   
#st.sidebar.write('You selected:', start_yr, 'to', end_yr)
 
#New Section      
st.sidebar.divider()
#Sidebar Filter
 
# Create FY Invoice filter
st.sidebar.header(":point_down: Filter for ESE data:")
fy_yr_inv = st.sidebar.multiselect(
        "Select the Financial Year of Invoice",
         options=df["Inv_Yr"].unique(),
         default=[2023,2022],
         )
 
if not fy_yr_inv:
       df2 = df.copy()
else:
       df2 = df[df["Inv_Yr"].isin(fy_yr_inv)]
 
# Create Region filter
region = st.sidebar.multiselect(
        "Select the REGION",
         df2["Region"].unique())
 
if not region:
       df3 = df2.copy()
else:
       df3 = df2[df2["Region"].isin(region)]
 
# Create Cost centre filter
cost_centre = st.sidebar.multiselect(
        "Select the COST CENTRE",
        df3["COST_CENTRE"].unique())
 
if not cost_centre:
       df4 = df3.copy()
else:
       df4 = df3[df3["COST_CENTRE"].isin(cost_centre)]
 
# Create Brand filter
brand = st.sidebar.multiselect(
        "Select the BRAND",
        df4["BRAND"].unique())
 
if not brand:
       df5 = df4.copy()
else:
       df5 = df4[df4["BRAND"].isin(brand)]
 
############################################################################################################################################################################################################
#Restrict the search result according to the filters
 
# No selection
if not fy_yr_inv and not region and not cost_centre and not brand:
       filter_df = df
# Only select Inv_Yr
elif not region and not cost_centre and not brand:
        filter_df = df[df["Inv_Yr"].isin(fy_yr_inv)]
# Only select Region
elif not fy_yr_inv and not cost_centre and not brand:
       filter_df = df[df["Region"].isin(region)]
# Only select cost centre
elif not fy_yr_inv and not region and not brand:
       filter_df = df[df["COST_CENTRE"].isin(cost_centre)]
# Only select brand
elif not fy_yr_inv and not region and not cost_centre:
       filter_df = df[df["BRAND"].isin(brand)]
# Select FY INV & Region
elif fy_yr_inv and region:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["Region"].isin(region)]
# Select 'Inv_Yr', 'COST_CENTRE'
elif fy_yr_inv and cost_centre:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Inv_Yr', 'BRAND'
elif fy_yr_inv and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE'
elif region and cost_centre:
      filter_df = df5[df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Region', 'BRAND'
elif region and brand:
      filter_df = df5[df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'COST_CENTRE', 'BRAND'
elif cost_centre and brand:
      filter_df = df5[df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Inv_Yr', 'Region', 'COST_CENTRE'
elif fy_yr_inv and region and cost_centre:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Inv_Yr', 'Region', 'BRAND'
elif fy_yr_inv and region and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'Inv_Yr', 'COST_CENTRE', 'BRAND'
elif fy_yr_inv and cost_centre and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE', 'BRAND'
elif region and cost_centre and brand:
      filter_df = df5[df["Region"].isin(region) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
 
#Show the original data table
#st.dataframe(df_selection)
############################################################################################################################################################################################################
#TAB 1: Overall category
# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
  font-size: 28px;
}
</style>
"""
st.write(font_css, unsafe_allow_html=True)      
 
#st.subheader(":eight_pointed_black_star: Mounter")
 
#Tab 1 MOUNTER
#with tab1:
#Set variable for slider result
selected_df = df_import[df_import["YEAR"].between(df_yr[0], df_yr[1])]
#MOUNTER IMPORT LINE CHART
row1_left_column, row1_right_column = st.columns(2)
with row1_left_column:
       st.subheader(":radio: :red[China Mounter Import Trend]_:orange[QTY]:")
       df_mounter_import = selected_df.groupby(by = ["MONTH","YEAR"], as_index= False)["QTY"].sum()
 # 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
       df_mounter_import = df_mounter_import.groupby(["YEAR", "MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_mounter_import['YEAR'].unique(), sort_Month_order],
                                   names=['YEAR', 'MONTH'])).fillna(0).reset_index()
       fig2 = go.Figure()      
       years = df_mounter_import["YEAR"].unique().astype(str)  # 獲取唯一年份並轉換為字符串列表
       for year in years:
              df_year = df_mounter_import[df_mounter_import["YEAR"] == year]
              fig2.add_trace(go.Scatter(
                     x=df_year["MONTH"],
                     y=df_year["QTY"],
                     mode='lines+markers+text',
                     marker=dict(size=9),
                     text=df_year["QTY"].apply(lambda x: f'{x:.3f}'),
                     textposition="bottom center",
                     texttemplate='%{text:.3s}',
                     name=year,
                     showlegend=True))
              fig2.update_layout(xaxis=dict(
                     tickmode='linear',
                     tick0=1,
                     dtick=1,
                     ),
                     yaxis=dict(showticklabels=True), 
                     font=dict(family="Arial, Arial", size=14, color="Black"),
                     hovermode='x', showlegend=True,
                     legend=dict(orientation="h",font=dict(size=14)),
                     paper_bgcolor='rgba(255,182,193,0.2)')#圖紙背景色
              fig2.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       
       st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)      
#####################################################################################################################################################
with row1_right_column:
       st.subheader(":radio: :blue[SMT Sales Trend]_:orange[QTY]:")
#LINE CHART of Overall Invoice Qty
       smtqtyAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('BRAND != "SOLDERSTAR"').query(
                                  'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').round(0).groupby(by = ["Inv_Yr","Inv_Month"],
                            as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3,4, 5, 6, 7, 8, 9, 10, 11, 12]
       smtqtyAmount_df2 = smtqtyAmount_df2.groupby(["Inv_Yr", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([smtqtyAmount_df2['Inv_Yr'].unique(), sort_Month_order],
                                   names=['Inv_Yr', 'Inv_Month'])).fillna(0).reset_index()
       fig3 = go.Figure()
# 添加每个Inv_Yr的折线
       fy_inv_values = smtqtyAmount_df2['Inv_Yr'].unique()
       for fy_inv in fy_inv_values:
         fy_inv_data = smtqtyAmount_df2[smtqtyAmount_df2['Inv_Yr'] == fy_inv]
         fig3.add_trace(go.Scatter(
              x=fy_inv_data['Inv_Month'],
              y=fy_inv_data['Item Qty'],
              mode='lines+markers+text',
              name=str(fy_inv),
              text=fy_inv_data['Item Qty'],
              textposition="bottom center",
              texttemplate='%{text:.3s}',
              hovertemplate='%{x}<br>%{y:.2f}',
              marker=dict(size=10)))
         fig3.update_layout(xaxis=dict(
              type='category',
              categoryorder='array',
              categoryarray=sort_Month_order),
              yaxis=dict(showticklabels=True),
              font=dict(family="Arial, Arial", size=14, color="Black"),
              hovermode='x', showlegend=True, 
              legend=dict(orientation="h",font=dict(size=14)),
              paper_bgcolor='rgba(0,150,255,0.1)'
              )
       fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
#####################################################
row1data_left_column, row1data_right_column = st.columns(2)
#with row1data_left_column:
with row1data_left_column:
       with st.expander(":point_right: Click to expand/ hide data"):
              pvt_qty = selected_df.round(0).pivot_table(
                 values=["QTY"],
                 index=["MONTH"],
                 columns=["YEAR"],
                 aggfunc="sum",
                 fill_value=0,
                 margins=True,
                 margins_name="Total",
                 observed=True) 

       #使用applymap方法應用格式化
              pvt_qty = pvt_qty.applymap('{:,.0f}'.format)
              html3 = pvt_qty.to_html(classes='table table-bordered', justify='center')
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
              html15 = f'<div style="zoom: 1.1;">{html14}</div>'
              st.markdown(html15, unsafe_allow_html=True)           
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv6 = pvt_qty.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv6, file_name='China_Mounter_Import_Qty.csv', mime='text/csv')
              st.divider()

###############################################################################################################
with row1data_right_column:
       with st.expander(":point_right: Click to expand/ hide data"):
              filter_df["Inv_Month"] = pd.Categorical(filter_df["Inv_Month"], categories=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
              pvt14 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('BRAND != "SOLDERSTAR"').query(
                      'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').pivot_table(
                      values="Item Qty",index=["Inv_Yr","Inv_Month"],columns=["BRAND"],
                      aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
              
              desired_order = ["YAMAHA", "PEMTRON", "HELLER","Total"]
              pvt14 = pvt14.reindex(columns=desired_order, level=1)

              # 定义会计数字格式的格式化函数
              def format_currency(value):
                     return "{:,.0f}".format(value)
# 计算小计行
              subtotal_row = pvt14.groupby(level=0).sum(numeric_only=True)
              subtotal_row.index = pd.MultiIndex.from_product([subtotal_row.index, [""]])
              subtotal_row.name = ("Subtotal", "")  # 小计行索引的名称
# 去除千位數符號並轉換為浮點數
              pvt14 = pvt14.applymap(lambda x: float(str(x).strip('').replace(',', '')))
# 转换为字符串并添加样式
              pvt14 = pvt14.applymap(lambda x: "{:,.0f}".format(x))
# 将小计行与pvt17连接，使用concat函数
              pvt14_concatenated = pd.concat([pvt14, subtotal_row])
# 生成HTML表格
              html_table = pvt14_concatenated.to_html(classes='table table-bordered', justify='center')
# 使用BeautifulSoup处理HTML表格
              soup = BeautifulSoup(html_table, 'html.parser')

# 找到所有的<td>标签，并为小于或等于0的值添加CSS样式
              for td in soup.find_all('td'):
                     value = float(td.text.replace('', '').replace(',', ''))
              if value <= 0:
                   td['style'] = 'color: red;'
      
# 找到所有的<td>标签，并将数值转换为会计数字格式的字符串
              for td in soup.find_all('td'):
                     value = float(td.text.strip('').replace(',', ''))
                     formatted_value = "{:,.0f}".format(value)
                     td.string.replace_with(formatted_value)
# 找到最底部的<tr>标签，并为其添加CSS样式
              last_row = soup.find_all('tr')[-1]
              last_row['style'] = 'background-color: yellow; font-weight: bold;'

# 在特定单元格应用其他样式           
              soup = str(soup)
              soup = soup.replace('<th>YAMAHA</th>', '<th style="background-color: lightgreen">YAMAHA</th>')
              soup = soup.replace('<th>PEMTRON</th>', '<th style="background-color: lightblue">PEMTRON</th>')
              soup = soup.replace('<th>HELLER</th>', '<th style="background-color: orange">HELLER</th>')
              soup = soup.replace('<td>', '<td style="text-align: middle;">')
              soup = soup.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
              html_with_style = str(f'<div style="zoom: 1.1;">{soup}</div>')
              st.markdown(html_with_style, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv1 = pvt14.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv1, file_name='SMT_invoice_Qty_Yr.csv', mime='text/csv')
              st.divider()

#加pivot table, column只用YAMAHA, PEMTRON, HELLER台數要match圖
################################################################################################################################# 
st.divider()

row2_left_column, row2_right_column = st.columns(2) 
with row2_left_column:
       st.subheader(":money_with_wings: :red[China Mounter Import Trend]_:green[AMOUNT]:")
       df_mounter_import = selected_df.groupby(by = ["MONTH","YEAR"], as_index= False)["Import_Amount(RMB)"].sum()
 # 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
       df_mounter_import = df_mounter_import.groupby(["YEAR", "MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_mounter_import['YEAR'].unique(), sort_Month_order],
                                   names=['YEAR', 'MONTH'])).fillna(0).reset_index()
       fig2 = go.Figure()      
       years = df_mounter_import["YEAR"].unique().astype(str)  # 獲取唯一年份並轉換為字符串列表
       for year in years:
              df_year = df_mounter_import[df_mounter_import["YEAR"] == year]
              fig2.add_trace(go.Scatter(
                     x=df_year["MONTH"],
                     y=df_year["Import_Amount(RMB)"],
                     mode='lines+markers+text',
                     marker=dict(size=9),
                     text=df_year["Import_Amount(RMB)"].apply(lambda x: f'{x:.3f}'),
                     textposition="bottom center",
                     texttemplate='%{text:.3s}',
                     name=year,
                     showlegend=True))
              fig2.update_layout(xaxis=dict(
                     tickmode='linear',
                     tick0=1,
                     dtick=1,
                     ),
                     yaxis=dict(showticklabels=True), 
                     font=dict(family="Arial, Arial", size=14, color="Black"),
                     hovermode='x', showlegend=True,
                     legend=dict(orientation="h",font=dict(size=14)),
                     paper_bgcolor='rgba(255,182,193,0.2)')#圖紙背景色                     
              fig2.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       
       st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)      

##########################################################################################################
with row2_right_column:
       st.subheader(":money_with_wings: :blue[SMT Sales Trend]_:green[AMOUNT]:")
#LINE CHART of Overall Invoice Qty
       smtinvoiceAmount_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).groupby(by = ["Inv_Yr","Inv_Month"],
                            as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3,4, 5, 6, 7, 8, 9, 10, 11, 12]
       smtinvoiceAmount_df = smtinvoiceAmount_df.groupby(["Inv_Yr", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([smtinvoiceAmount_df['Inv_Yr'].unique(), sort_Month_order],
                                   names=['Inv_Yr', 'Inv_Month'])).fillna(0).reset_index()
       fig3 = go.Figure()
# 添加每个Inv_Yr的折线
       fy_inv_values = smtinvoiceAmount_df['Inv_Yr'].unique()
       for fy_inv in fy_inv_values:
         fy_inv_data = smtinvoiceAmount_df[smtinvoiceAmount_df['Inv_Yr'] == fy_inv]
         fig3.add_trace(go.Scatter(
              x=fy_inv_data['Inv_Month'],
              y=fy_inv_data['Before tax Inv Amt (HKD)'],
              mode='lines+markers+text',
              name=str(fy_inv),
              text=fy_inv_data['Before tax Inv Amt (HKD)'],
              textposition="bottom center",
              texttemplate='%{text:.3s}',
              hovertemplate='%{x}<br>%{y:.2f}',
              marker=dict(size=10)))
         fig3.update_layout(xaxis=dict(
              type='category',
              categoryorder='array',
              categoryarray=sort_Month_order),
              yaxis=dict(showticklabels=True),
              font=dict(family="Arial, Arial", size=14, color="Black"),
              hovermode='x', showlegend=True,
              legend=dict(orientation="h",font=dict(size=14)),
              paper_bgcolor='rgba(0,150,255,0.1)')
       fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
############################################################################################################################################################################################################
row2data_left_column, row2data_right_column = st.columns(2) 

#with row2data_left_column:
with row2data_left_column:
       with st.expander(":point_right: Click to expand/ hide data"):
              pvt_amount = selected_df.round(0).pivot_table(
                     values=["Import_Amount(RMB)"],
                     index=["MONTH"],
                     columns=["YEAR"],
                     aggfunc="sum",
                     fill_value=0,
                     margins=True,
                     margins_name="Total",
                     observed=True) 

#使用applymap方法應用格式化
              pvt_amount = pvt_amount.applymap('{:,.0f}'.format)
              html3 = pvt_amount.to_html(classes='table table-bordered', justify='center')
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
              html15 = f'<div style="zoom: 1.1;">{html14}</div>'
              st.markdown(html15, unsafe_allow_html=True)           
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv6 = pvt_amount.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv6, file_name='China_Mounter_Import_Amount.csv', mime='text/csv')
              st.divider()

###################################################################
with row2data_right_column:
       with st.expander(":point_right: Click to expand/ hide data"):
              filter_df["Inv_Month"] = pd.Categorical(filter_df["Inv_Month"], categories=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
              pvt2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').pivot_table(
                      values="Before tax Inv Amt (HKD)",index=["Inv_Month"],columns=["Inv_Yr",],
                      aggfunc="sum",fill_value=0, margins=True,margins_name="Total")

# 定义会计数字格式的格式化函数
#              def format_currency(value):
#                     return "{:,.0f}".format(value)
# 计算小计行
#              subtotal_row = pvt2.groupby(level=0).sum(numeric_only=True)
#              subtotal_row.index = pd.MultiIndex.from_product([subtotal_row.index, [""]])
#              subtotal_row.name = ("Subtotal", "")  # 小计行索引的名称
# 去除千位數符號並轉換為浮點數
              pvt2 = pvt2.applymap(lambda x: float(str(x).strip('').replace(',', '')))
# 转换为字符串并添加样式
#              pvt2 = pvt2.applymap(lambda x: "{:,.0f}".format(x))
# 将小计行与pvt17连接，使用concat函数
#              pvt14_concatenated = pd.concat([pvt2, subtotal_row])
# 生成HTML表格
              html_table = pvt2.to_html(classes='table table-bordered', justify='center')
# 使用BeautifulSoup处理HTML表格
              soup = BeautifulSoup(html_table, 'html.parser')

# 找到所有的<td>标签，并为小于或等于0的值添加CSS样式
              for td in soup.find_all('td'):
                     value = float(td.text.replace('', '').replace(',', ''))
              if value <= 0:
                   td['style'] = 'color: red;'
      
# 找到所有的<td>标签，并将数值转换为会计数字格式的字符串
              for td in soup.find_all('td'):
                     value = float(td.text.strip('').replace(',', ''))
                     formatted_value = "{:,.0f}".format(value)
                     td.string.replace_with(formatted_value)
# 找到最底部的<tr>标签，并为其添加CSS样式
              last_row = soup.find_all('tr')[-1]
              last_row['style'] = 'background-color: yellow; font-weight: bold;'

# 在特定单元格应用其他样式           
              soup = str(soup)
#              soup = soup.replace('<th>YAMAHA</th>', '<th style="background-color: lightgreen">YAMAHA</th>')
#              soup = soup.replace('<th>PEMTRON</th>', '<th style="background-color: lightblue">PEMTRON</th>')
#              soup = soup.replace('<th>HELLER</th>', '<th style="background-color: orange">HELLER</th>')
#              soup = soup.replace('<td>', '<td style="text-align: middle;">')
              soup = soup.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
              html_with_style = str(f'<div style="zoom: 1.1;">{soup}</div>')
              st.markdown(html_with_style, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv2 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv2, file_name='SMT_invoice_Amount.csv', mime='text/csv')

#########################################################################
