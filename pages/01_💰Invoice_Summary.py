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
######################################################################################################
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")

title_row1, title_row2, title_row3, title_row4 = st.columns(4)


st.title(':world_map: SMT_Invoice Dashboard')
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
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
df = pd.read_excel(
               io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='raw_sheet', skiprows=0, usecols='A:AO',nrows=10000,).query(
                    'Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query(
                         'FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
df_sales_target = pd.read_excel(
               io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='YAMAHA_Sales_Target', skiprows=0, usecols='A:D',nrows=10000,)
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


#Sidebar Filter
# Create FY Invoice filter
st.sidebar.header(":point_down:Filter:")
fy_yr_inv = st.sidebar.multiselect(
       "Select the Financial Year of Invoice",
        options=df["FY_INV"].unique(),
        default=["FY 23/24"],
        )

 
if not fy_yr_inv:
      df2 = df.copy()
else:
      df2 = df[df["FY_INV"].isin(fy_yr_inv)]
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
# Only select FY_INV
elif not region and not cost_centre and not brand:
       filter_df = df[df["FY_INV"].isin(fy_yr_inv)]
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
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["Region"].isin(region)]
# Select 'FY_INV', 'COST_CENTRE'
elif fy_yr_inv and cost_centre:
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_INV', 'BRAND'
elif fy_yr_inv and brand:
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE'
elif region and cost_centre:
     filter_df = df5[df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Region', 'BRAND'
elif region and brand:
     filter_df = df5[df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'COST_CENTRE', 'BRAND'
elif cost_centre and brand:
     filter_df = df5[df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'FY_INV', 'Region', 'COST_CENTRE'
elif fy_yr_inv and region and cost_centre:
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_INV', 'Region', 'BRAND'
elif fy_yr_inv and region and brand:
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'FY_INV', 'COST_CENTRE', 'BRAND'
elif fy_yr_inv and cost_centre and brand:
     filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE', 'BRAND'
elif region and cost_centre and brand:
     filter_df = df5[df["Region"].isin(region) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
#Show the original data table
#st.dataframe(df_selection)
############################################################################################################################################################################################################
#Overall Summary

 
left_column, middle_column, right_column = st.columns(3)
total_invoice_amount = int(filter_df["Before tax Inv Amt (HKD)"].sum())
with left_column:
             st.metric(label=":dollar: Total :orange[Invoice] Amount before tax",value=(f"HKD{total_invoice_amount:,}"),)
             total_gp = int(filter_df["G.P.  (HKD)"].sum())
with middle_column:
             st.metric(label=":moneybag: Total :orange[G.P] Amount",value=(f"HKD{total_gp:,}"))

 
total_unit_qty = int(filter_df["Item Qty"].sum())
with right_column:
             st.metric(label=":factory: Total :orange[Invoiced] Machine",value=(f"Qty: {total_unit_qty:,}"))

 
style_metric_cards(background_color="Summer",border_left_color="BLUE",border_size_px=5,box_shadow=True,border_radius_px=1)

 
############################################################################################################################################################################################################
#Pivot table, 差sub-total, GP%

 
filter_df["Inv_Month"] = filter_df["Inv_Month"].astype(str)
filter_df["Inv_Yr"] = filter_df["Inv_Yr"].astype(str)
filter_df["Contract No."] = filter_df["Contract No."].astype(str)
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
tab1, tab2, tab3 ,tab4,tab5= st.tabs([":wedding: Overview",":earth_asia: Region",":blue_book: Invoice Details",":package: Brand",":handshake: Customer"])

#TAB 1: Overall category
################################################################################################################################################
with tab1:

       st.subheader(""":globe_with_meridians: Invoice Amount_:orange[Monthly]:""")

# 将 "Inv_Month" 列转换为 Categorical 数据类型并指定自定义排序
       # 原始程式碼...

# 將 "Inv_Month" 列轉換為 int 資料類型
#       filter_df['Inv_Month'] = filter_df['Inv_Month'].astype(str)

#       month_order = {'4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9, '10': 10, '11': 11, '12': 12, '1': 1, '2': 2, '3': 3}
#       filter_df['Inv_Month'] = filter_df['Inv_Month'].map(lambda x: month_order.get(int(x), x))
####################################################     
 
       pvt2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(
              index=["FY_INV","FQ(Invoice)","Inv_Yr", "Inv_Month"],
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
       pvt2 = pvt2.applymap('HKD{:,.0f}'.format)
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
       html15 = f'<div style="zoom: 1.5;">{html14}</div>'
       st.markdown(html15, unsafe_allow_html=True)           
 
 
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
       csv1 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
       st.download_button(label='Download Table', data=csv1, file_name='Monthly_Sales.csv', mime='text/csv')

################################################################################
#Pivot table2
       with st.expander(":point_right: click to expand"):
             st.subheader(":clipboard: Invoice Amount Subtotal_:orange[FQ]:")
             pvt17 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(
                     values="Before tax Inv Amt (HKD)",
                     index=["FY_INV","FQ(Invoice)"],
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
             html_with_style = str(f'<div style="zoom: 1.5;">{soup}</div>')
             st.markdown(html_with_style, unsafe_allow_html=True)       
       
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv6 = pvt17.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv6, file_name='Cost_Centre_Quarter_Sales.csv', mime='text/csv')
             st.divider()
###################################################################################################
       col_1, col_2= st.columns(2)
       with col_1:
#LINE CHART of Overall Invoice Amount
             st.subheader(":chart_with_upwards_trend: Invoice Amount Trend_:orange[FQ](Available to Show :orange[Multiple FY)]:")
             InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","FQ(Invoice)","Inv_Month"
                          ], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             InvoiceAmount_df2 = InvoiceAmount_df2.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([InvoiceAmount_df2['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
             fig3 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = InvoiceAmount_df2['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                   fy_inv_data = InvoiceAmount_df2[InvoiceAmount_df2['FY_INV'] == fy_inv]
                   fig3.add_trace(go.Scatter(
                         x=fy_inv_data['Inv_Month'],
                         y=fy_inv_data['Before tax Inv Amt (HKD)'],
                         mode='lines+markers+text',
                         name=fy_inv,
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
                         font=dict(family="Arial, Arial", size=12, color="Black"),
                         hovermode='x', showlegend=True,
                         legend=dict(orientation="h",font=dict(size=14)))
                   fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                   st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
#############################################################################################################
#FY to FY Quarter Invoice Details:
                   pvt6 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_INV"],columns=["FQ(Invoice)"],
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
                   st.markdown(html117, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
                   csv2 = pvt6.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
                   st.download_button(label='Download Table', data=csv2, file_name='FQ_Sales.csv', mime='text/csv')
                   st.divider()
################################################################################################################################################
#New Section 
#LINE CHART of GP Amount
       with col_2:
             st.subheader(":chart_with_upwards_trend: :green[G.P Amount Trend]_FQ(Available to Show :green[Multiple FY)]:")
             InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).groupby(by =
                     ["Inv_Month","FY_INV"], as_index= False)["G.P.  (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             InvoiceAmount_df2 = InvoiceAmount_df2.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([InvoiceAmount_df2['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
             fig12 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = InvoiceAmount_df2['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                   fy_inv_data = InvoiceAmount_df2[InvoiceAmount_df2['FY_INV'] == fy_inv]
                   fig12.add_trace(go.Scatter(
                         x=fy_inv_data['Inv_Month'],
                         y=fy_inv_data['G.P.  (HKD)'],
                         mode='lines+markers+text',
                         name=fy_inv,
                         text=fy_inv_data['G.P.  (HKD)'],
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
      #FY to FY Quarter Invoice Details:
             pvt16 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).pivot_table(values="G.P.  (HKD)",index=["FY_INV"],columns=["FQ(Invoice)"],
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
             st.markdown(html23, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv3 = pvt16.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv3, file_name='G.P Amount.csv', mime='text/csv')        
############################################################################################################################################################################################################
#TAB 2: Region Category
with tab2:
       one_column, two_column= st.columns(2)
       with one_column:
        st.subheader(":sunrise: FY Invoice Details_:orange[Monthly]:")
# 計算"FQ(Invoice)"的subtotal數值
#              filter_df = df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"')
        pvt = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(
              values="Before tax Inv Amt (HKD)",
              index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month"],
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
        html24 = f'<div style="zoom: 0.85;">{html24}</div>'
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
        st.subheader(":clipboard: Invoice Amount Subtotal_:orange[FQ]:")
        pvt7 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(
                     values="Before tax Inv Amt (HKD)",
                     index=["FY_INV","FQ(Invoice)"],
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
        html_with_style2 = str(f'<div style="zoom: 1;">{soup2}</div>')
        st.markdown(html_with_style2, unsafe_allow_html=True)       
###########################################################################################################################      
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
        csv5 = pvt7.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
        st.download_button(label='Download Table', data=csv5, file_name='Regional_Quarter_Sales.csv', mime='text/csv')
       
######################################################################################################################
       with two_column:
#All Regional total inv amount BAR CHART
      
        st.subheader(":bar_chart: Invoice Amount_:orange[FY](Available to show :orange[Multiple FY]):")
        category2_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by=["FY_INV","Region"], 
                       as_index=False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)", ascending=False)

        df_contract_vs_invoice = px.bar(category2_df, x="FY_INV", y="Before tax Inv Amt (HKD)", color="Region", text_auto='.3s')

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

        st.plotly_chart(df_contract_vs_invoice, use_container_width=True) 
##############################################################################################################################          
# LINE CHART of Regional Comparision
        st.subheader(":chart_with_upwards_trend: Invoice Amount Trend_:orange[All Region in one]:")
        InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by = ["FQ(Invoice)","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
        # 使用pivot_table函數來重塑數據，使每個Region成為一個列
        InvoiceAmount_df2 = InvoiceAmount_df2.pivot_table(index="FQ(Invoice)", columns="Region", values="Before tax Inv Amt (HKD)", fill_value=0).reset_index()
        # 使用melt函數來恢復原來的長格式，並保留0值
        InvoiceAmount_df2 = InvoiceAmount_df2.melt(id_vars="FQ(Invoice)", value_name="Before tax Inv Amt (HKD)", var_name="Region")
 
        fig2 = px.line(InvoiceAmount_df2,
                       x = "FQ(Invoice)",
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
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[SOUTH CHINA] Inv Amt Trend_FQ(Available to Show :orange[Multiple FY]):")
              df_Single_south = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "SOUTH"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV",
                                 "Inv_Month"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Invoice)" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_south = df_Single_south.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_south['FY_INV'].unique(), all_fq_invoice_values],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
              fig4 = go.Figure()

# 添加每个FY_INV的折线
              fy_inv_values = df_Single_south['FY_INV'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_south[df_Single_south['FY_INV'] == fy_inv]
               fig4.add_trace(go.Scatter(
                          x=fy_inv_data['Inv_Month'],
                          y=fy_inv_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Before tax Inv Amt (HKD)'],
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

             
#SOUTH Region Invoice Details FQ_FQ:
              pvt8 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "SOUTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",
                     index=['FY_INV'],columns=["FQ(Invoice)"],aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_INV',ascending=True)
            
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
              html833 = f'<div style="zoom: 1;">{html822}</div>'

              st.markdown(html833, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv11 = pvt8.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv11, file_name='South_Sales.csv', mime='text/csv')
              st.divider()

 
 
       with two_column:
# LINE CHART of EAST CHINA FY/FY
              st.divider()
              st.subheader(":chart_with_upwards_trend: :orange[EAST CHINA] Inv Amt Trend_FQ(Available to Show :orange[Multiple FY)]:")
              df_Single_region = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "EAST"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Invoice)" 列中的所有值都出现在 df_Single_region 中
              all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
              df_Single_region = df_Single_region.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_region['FY_INV'].unique(), all_fq_invoice_values],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
              fig5 = go.Figure()

# 添加每个FY_INV的折线
              fy_inv_values = df_Single_region['FY_INV'].unique()
              for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_region[df_Single_region['FY_INV'] == fy_inv]
               fig5.add_trace(go.Scatter(
                          x=fy_inv_data['Inv_Month'],
                          y=fy_inv_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Before tax Inv Amt (HKD)'],
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
#EAST Region Invoice Details FQ_FQ:
              pvt9 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "EAST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_INV',ascending=True)
            
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
              html900 = f'<div style="zoom: 1;">{html89}</div>'
             
              st.markdown(html900, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
              csv12 = pvt9.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
              st.download_button(label='Download Table', data=csv12, file_name='East_Sales.csv', mime='text/csv')

              st.divider()
################################################# 
       three_column, four_column= st.columns(2)
  
       with three_column:
# LINE CHART of NORTH CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[NORTH CHINA] Inv Amt Trend_FQ(Available to Show :orange[Multiple FY]):")

             df_Single_north = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "NORTH"').query('FY_INV != "TBA"').round(0).groupby(by=["FY_INV", "Inv_Month"],
                                as_index=False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Invoice)" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_north = df_Single_north.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_north['FY_INV'].unique(), all_fq_invoice_values],
                               names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
             fig7 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = df_Single_north['FY_INV'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_north[df_Single_north['FY_INV'] == fy_inv]
               fig7.add_trace(go.Scatter(
                          x=fy_inv_data['Inv_Month'],
                          y=fy_inv_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Before tax Inv Amt (HKD)'],
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

#NORTH Region Invoice Details FQ_FQ:
             pvt10 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "NORTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_INV',ascending=True)

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
             html699 = f'<div style="zoom: 1;">{html68}</div>'
             st.markdown(html699, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv9 = pvt10.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv9, file_name='North_Sales.csv', mime='text/csv')

             st.divider()
##################################################
       with four_column:
# LINE CHART of WEST CHINA FY/FY
             st.subheader(":chart_with_upwards_trend: :orange[WEST CHINA] Inv Amt Trend_FQ(Available to Show :orange[Multiple FY]):")
             df_Single_west = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "WEST"').query('FY_INV != "TBA"').round(0).groupby(by=["FY_INV", "Inv_Month"],
                                as_index=False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "FQ(Invoice)" 列中的所有值都出现在 df_Single_region 中
             all_fq_invoice_values = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_Single_west = df_Single_west.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_Single_west['FY_INV'].unique(), all_fq_invoice_values],
                               names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
             fig6 = go.Figure()

# 添加每个FY_INV的折线
             fy_inv_values = df_Single_west['FY_INV'].unique()
             for fy_inv in fy_inv_values:
               fy_inv_data = df_Single_west[df_Single_west['FY_INV'] == fy_inv]
               fig6.add_trace(go.Scatter(
                          x=fy_inv_data['Inv_Month'],
                          y=fy_inv_data['Before tax Inv Amt (HKD)'],
                          mode='lines+markers+text',
                          name=fy_inv,
                          text=fy_inv_data['Before tax Inv Amt (HKD)'],
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
#WEST Region Invoice Details FQ_FQ:
             pvt18 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "WEST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
                            aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_values(by='FY_INV',ascending=True)
            
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
             html766 = f'<div style="zoom: 1;">{html75}</div>'

             st.markdown(html766, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv10 = pvt18.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv10, file_name='West_Sales.csv', mime='text/csv')
             st.divider()

############################################################################################################################################ 
#TAB 3 Invocie Details
with tab3:
     tab3_col1, tab3_col2, tab3_col3 = st.columns(3)

     with tab3_col1:
      
            #FY to FY Quarter Invoice Details:
      st.subheader(":ledger: Invoice Amount Subtotal_:orange[FY]:")
      with st.subheader("Click to expand"):
             pvt21 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).pivot_table(
                    values=["Before tax Inv Amt (HKD)","G.P.  (HKD)"],index=["FY_INV"],
                    aggfunc="sum",fill_value=0, margins=True,margins_name="Total")
            
             html146 = pvt21.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center') 
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html147 = html146.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html148 = html147.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             st.markdown(html148, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv19 = pvt21.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
      st.download_button(label='Download Table', data=csv19, file_name='FY_Sales_Total.csv', mime='text/csv') 
#############################################
#FY to FY Quarter Invoice Details:
      st.subheader(":blue_book: Invoice Amount Subtotal_:orange[FQ]:")
      with st.subheader("Click to expand"):
             pvt19 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).pivot_table(values=["Before tax Inv Amt (HKD)","G.P.  (HKD)"],
                     index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month"],aggfunc="sum",fill_value=0, 
                     margins=False,margins_name="Total").sort_index(axis=0, ascending=False)
            
             html55 = pvt19.applymap('HKD{:,.0f}'.format).to_html(classes='table table-bordered', justify='center')
             #st.dataframe(pvt6.style.highlight_max(color = 'yellow', axis = 0)
             #                       .format("HKD{:,}"), use_container_width=True)   
             # 把total值的那行的背景顏色設為黃色，並將字體設為粗體
             html56 = html55.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
             #改column color
             html57 = html56.replace('<th>Inv_Month</th>', '<th style="background-color: orange">Inv_Month</th>')
             html58 = html57.replace('<th>Inv_Yr</th>', '<th style="background-color: orange">Inv_Yr</th>')
             html59 = html58.replace('<th>FQ(Invoice)</th>', '<th style="background-color: orange">FQ(Invoice)</th>')
             html60 = html59.replace('<th>FY_INV</th>', '<th style="background-color: orange">FY_INV</th>')
             html61 = html60.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
             html61 = html61.replace('<th>G.P.  (HKD)</th>', '<th style="background-color: lightblue">G.P.  (HKD)</th>')
             html61 = html61.replace('<th>Before tax Inv Amt (HKD)</th>', '<th style="background-color: yellow">Before tax Inv Amt (HKD)</th>')

             html611= f'<div style="zoom: 0.9;">{html61}</div>'
             st.markdown(html611, unsafe_allow_html=True)
             
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv8 = pvt19.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
      st.download_button(label='Download Table', data=csv8, file_name='Monthly_Sales_Total.csv', mime='text/csv')
 ###################################################################################################################     
     with tab3_col2:
      st.subheader(":green_book: Invoice Details_:orange[Monthly]")
           
      #with st.expander("Click to expand"):
      pvt3 = filter_df.query('FY_INV != "TBA"').query('Inv_Month != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(
             index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month","COST_CENTRE","Region","Contract No.","Customer Name",
             "Ordered_Items"],values=["Item Qty","Before tax Inv Amt (HKD)","G.P.  (HKD)"],
              aggfunc="sum",
              fill_value=0,
              margins=False,
              margins_name="Total",
              observed=True).sort_index(axis=0, ascending=False)  # This ensures subtotals are only calculated for existing values)
    
# 调整值的顺序
      pvt3 = pvt3[["Item Qty", "Before tax Inv Amt (HKD)", "G.P.  (HKD)"]]
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
      csv7 = pvt3.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
      st.download_button(label='Download Table', data=csv7, file_name='Invoice_details.csv', mime='text/csv')


# 使用map方法應用格式化
      pvt3_formatted = pvt3.copy()

# 定義格式化函數
      def format_item_qty(value):
           return "{:.1f}".format(value)

# 使用 apply() 將格式化函數應用於 "Item Qty" 列的每個元素
      pvt3_formatted["Item Qty"] = pvt3_formatted["Item Qty"].apply(format_item_qty)

# 使用 applymap() 將格式化函數應用於 "Before tax Inv Amt (HKD)" 和 "G.P.  (HKD)" 列的每個元素
      pvt3_formatted[["Before tax Inv Amt (HKD)", "G.P.  (HKD)"]] = pvt3_formatted[["Before tax Inv Amt (HKD)", "G.P.  (HKD)"]].applymap('HKD{:,.0f}'.format)


      html46 = pvt3_formatted.to_html(classes='table table-bordered', justify='center')

# 放大pivot table
      html46 = f'<div style="zoom: 0.9;">{html46}</div>'
# 將你想要變色的column header找出來，並加上顏色
      html47 = html46.replace('<th>Before tax Inv Amt (HKD)</th>', '<th style="background-color: yellow">Before tax Inv Amt (HKD)</th>')
      html48 = html47.replace('<th>Inv_Month</th>', '<th style="background-color: orange">Inv_Month</th>')
      html48 = html48.replace('<th>Inv_Yr</th>', '<th style="background-color: orange">Inv_Yr</th>')
      html48 = html48.replace('<th>FQ(Invoice)</th>', '<th style="background-color: orange">FQ(Invoice)</th>')
      html48 = html48.replace('<th>FY_INV</th>', '<th style="background-color: orange">FY_INV</th>')
      html49 = html48.replace('<th>G.P. (HKD)</th>', '<th style="background-color: Khaki">G.P. (HKD)</th>')
      html50 = html49.replace('<th>Item Qty</th>', '<th style="background-color: lightgreen">Item Qty</th>')
      html50 = html50.replace('<th>G.P.  (HKD)</th>', '<th style="background-color: lightblue">G.P.  (HKD)</th>')
# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
      html51 = html50.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把每個數值置中
      html52 = html51.replace('<td>', '<td style="text-align: middle;">')
# 把total值的那列的字設為黃色
      html53 = html52.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 把所有數值等於或少於0的數值的顏色設為紅色
      html54 = html53.replace('<td>-', '<td style="color: red;">-')
# 使用Streamlit的markdown來顯示HTML表格
      st.markdown(html54, unsafe_allow_html=True)
#############################################################################################################      


#TAB 4: Brand category
with tab4:
      # BAR CHART of BRAND Comparision
       st.subheader(":bar_chart: Main Brand Invoice Qty_:orange[Monthly]:")
       brand_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('BRAND != "SOLDERSTAR"').query(
                      'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').round(0).groupby(by=["Inv_Month",
                            "BRAND"], as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
        # 按照指定顺序排序
#       brand_df["BRAND"] = pd.Categorical(brand_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER"])
#       brand_df = brand_df.sort_values("BRAND")
       # 创建一个包含排序顺序的列表
       sort_Month_order = ["4", "5", "6", "7", "8", "9", "10","11","12", "1", "2", "3"]

# 将Inv_Month的数据类型更改为category
#       brand_df["Inv_Month"] = pd.Categorical(brand_df["Inv_Month"], categories=sort_Month_order, ordered=True)

# 使用plotly绘制柱状图           
       brand_qty = px.bar(brand_df, x="Inv_Month", y="Item Qty", color="BRAND", text_auto='.3s')

# 设置x轴的分类顺序
       brand_qty.update_layout(xaxis={"type": "category", "categoryorder": "array", "categoryarray": sort_Month_order})

# 更改顏色
       colors = {"YAMAHA": "lightgreen","PEMTRON": "lightblue","HELLER": "orange"}
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

# 绘制图表
       st.plotly_chart(brand_qty, use_container_width=True)

########################################################################################################      
       tab4_row2_col1, tab4_row2_col2, tab4_row2_col3 = st.columns(3)
       with tab4_row2_col1:
#Brand Inv Qty by Inv Month:
             st.subheader(":clipboard:  Main Brand Invoice Qty_:orange[FQ Subtotal]:")
             #with st.expander("Click to expand"):
             pvt6 = filter_df.query('FY_INV != "TBA"').query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').round(0).pivot_table(
                    values="Item Qty",
                    index=["FY_INV","FQ(Invoice)"],
                    columns=["BRAND"],
                    aggfunc="sum",
                    fill_value=0,
                    margins=True,
                    margins_name="Total",
                    observed=True)  # This ensures subtotals are only calculated for existing values)
            
             desired_order_brand = ["YAMAHA", "PEMTRON", "HELLER","Total"]
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
             soup3 = soup3.replace('<th>YAMAHA</th>', '<th style="background-color: lightgreen">YAMAHA</th>')
             soup3 = soup3.replace('<th>PEMTRON</th>', '<th style="background-color: lightblue">PEMTRON</th>')
             soup3 = soup3.replace('<th>HELLER</th>', '<th style="background-color: orange">HELLER</th>')
             soup3 = soup3.replace('<th>WEST</th>', '<th style="background-color: lightgreen">WEST</th>')
             soup3 = soup3.replace('<td>', '<td style="text-align: middle;">')
             soup3 = soup3.replace('<th>Q1</th>', '<th style="background-color: lightgrey">Q1</th>')
             soup3 = soup3.replace('<th>Q2</th>', '<th style="background-color: pink">Q2</th>')
             soup3 = soup3.replace('<th>Q3</th>', '<th style="background-color: lightgrey">Q3</th>')
             soup3 = soup3.replace('<th>Q4</th>', '<th style="background-color: pink">Q4</th>')
             soup3 = soup3.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 在网页中显示HTML表格
             html_with_style3 = str(f'<div style="zoom: 1;">{soup3}</div>')
             st.markdown(html_with_style3, unsafe_allow_html=True)       
      
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
             csv13 = pvt6.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
             st.download_button(label='Download Table', data=csv13, file_name='Brand_invoice_qty.csv', mime='text/csv')
            
############################################################################################################################################      
       with tab4_row2_col2:
             st.subheader(":sports_medal: Main Brand Invoice Qty_:orange[FY]:")
             brandinv_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('BRAND != "SOLDERSTAR"').query(
                           'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHINWA"').query(
                           'BRAND != "SIGMA"').round(0).groupby(by=["FY_INV","BRAND"],
                            as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
        # 按照指定顺序排序
#             brandinv_df["BRAND"] = pd.Categorical(brandinv_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER"])
#             brandinv_df = brandinv_df.sort_values("BRAND")
             df_brand = px.bar(brandinv_df, x="FY_INV", y="Item Qty", color="BRAND", text_auto='.3s')
# 更改顏色
             colors = {"PEMTRON": "lightblue","HELLER": "orange","YAMAHA": "lightgreen",}
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
             st.plotly_chart(df_brand, use_container_width=True)            
###################################################################
       with tab4_row2_col3:
             st.subheader(":round_pushpin: Main Brand Invoice Qty_:orange[Percentage]:")

# 创建示例数据框
            
             brand_data = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by=["FY_INV","BRAND"],
                     as_index=False)["Item Qty"].sum().sort_values(by="Item Qty", ascending=False)
            
             brandinvpie_df = pd.DataFrame(brand_data)

# 按照指定順序排序

 
             brandinvpie_df["BRAND"] = brandinvpie_df["BRAND"].replace(to_replace=[x for x in brandinvpie_df["BRAND"
                                       ].unique() if x not in ["YAMAHA", "PEMTRON", "HELLER"]], value="OTHERS")
             brandinvpie_df["BRAND"] = pd.Categorical(brandinvpie_df["BRAND"], ["YAMAHA", "PEMTRON", "HELLER","OTHERS"])

# 创建饼状图
             df_pie = px.pie(brandinvpie_df, values="Item Qty", names="BRAND", color="BRAND", color_discrete_map={
                      "PEMTRON": "lightblue", "HELLER": "orange", "YAMAHA": "lightgreen", "OTHERS":"purple"})

# 设置字体和样式
             df_pie.update_layout(
                   font=dict(family="Arial", size=14, color="black"),
                   legend=dict(orientation="v", yanchor="bottom", y=1.02, xanchor="right", x=1))

# 显示百分比标签
             df_pie.update_traces(textposition='outside', textinfo='label+percent', marker_line_width=2,opacity=1)

# 在Streamlit中显示图表
             st.plotly_chart(df_pie, use_container_width=True)
            
###########################################################################################################
#Top Product line chart invoice qty trend
       left_column, middle_column, right_column = st.columns(3)
       with middle_column:
             st.divider()
             st.header(":chart_with_upwards_trend: Main Unit Invoice Qty_:orange[Monthly]:")
###########################################################################################################
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM20R Invoice Details:
       with left_column:
             
             st.subheader(":red[YAMAHA:] YSM20")
             df_YSM20 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM20 = df_YSM20.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM20['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig8 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YSM20['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YSM20[df_YSM20['FY_INV'] == fy_inv]
                  fig8.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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

#FY to FY YSM20R Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt13 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query(
                   'Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#Line Chart FY to FY YSM10 Invoice Details:
       with right_column:
             st.subheader(":red[YAMAHA:] YSM10")
             df_YSM10 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM10 = df_YSM10.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM10['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig9 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YSM10['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YSM10[df_YSM10['FY_INV'] == fy_inv]
                  fig9.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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
#FY to FY YSM10 Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt14 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query(
                   'Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#Line Chart FY to FY YSM40R Invoice Details:
       with left_column:
             st.divider()
             st.subheader(":red[YAMAHA:] YSM40")
             df_YSM40 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM40R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSM40 = df_YSM40.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSM40['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig10 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YSM40['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YSM40[df_YSM40['FY_INV'] == fy_inv]
                  fig10.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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
#FY to FY YSM40R Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt14 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM40R"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#Line Chart FY to FY YSi-V Invoice Details:
       with right_column:
             st.divider()
             st.subheader(":red[YAMAHA:] YSi-V(AOI)")
             df_YSIV = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSi-V"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YSIV = df_YSIV.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YSIV['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig11 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YSIV['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YSIV[df_YSIV['FY_INV'] == fy_inv]
                  fig11.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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
#FY to FY YSi-V Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt15 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query(
                   'Ordered_Items == "YSi-V"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#Line Chart FY to FY YRM Invoice Details:
       with left_column:
             st.divider()
             st.subheader(":red[YAMAHA:] YRM")
             df_YRM = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YRM"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YRM = df_YRM.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YRM['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig13 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YRM['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YRM[df_YRM['FY_INV'] == fy_inv]
                  fig13.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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

#FY to FY YRM Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt20 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Item Qty",
                     index=["FY_INV"],columns=["FQ(Invoice)"], aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=True)
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
            
             st.download_button(label='Download Table', data=csv16, file_name='YRM Invocie Qty.csv', mime='text/csv')          
#############################################################################################################################    
       left_column, right_column = st.columns(2) 
#Line Chart FY to FY PEMTRON Invoice Details:
       with left_column:
             st.divider()
             st.subheader(":blue[PEMTRON:]")
             df_PEMTRON = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('BRAND == "PEMTRON"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_PEMTRON = df_PEMTRON.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_PEMTRON['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig14 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_PEMTRON['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_PEMTRON[df_PEMTRON['FY_INV'] == fy_inv]
                  fig14.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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
                 
             fig14.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig14.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig14.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY PEMTRON Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt17 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('BRAND == "PEMTRON"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#Line Chart FY to FY HELLER Invoice Details:
       with right_column:
             st.divider()
             st.subheader(":orange[HELLER:]")
             df_YRM = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('BRAND == "HELLER"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","Inv_Month"],
                                  as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
             sort_Month_order = ["4", "5", "6", "7", "8", "9", "10", "11", "12", "1", "2", "3"]
             df_YRM = df_YRM.groupby(["FY_INV", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([df_YRM['FY_INV'].unique(), sort_Month_order],
                                   names=['FY_INV', 'Inv_Month'])).fillna(0).reset_index()
#建立圖表         
             fig15 = go.Figure()
# 添加每个FY_INV的折线
             fy_inv_values = df_YRM['FY_INV'].unique()
             for fy_inv in fy_inv_values:
                  fy_inv_data = df_YRM[df_YRM['FY_INV'] == fy_inv]
                  fig15.add_trace(go.Scatter(
                       x=fy_inv_data['Inv_Month'],
                       y=fy_inv_data["Item Qty"],
                       mode='lines+markers+text',
                       name=fy_inv,
                       text=fy_inv_data["Item Qty"],
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
                 
             fig15.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             fig15.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
             st.plotly_chart(fig15.update_layout(yaxis_showticklabels = True), use_container_width=True)

#FY to FY HELLER Invoice Details:
             filter_df["FQ(Invoice)"] = pd.Categorical(filter_df["FQ(Invoice)"], categories=["Q1", "Q2", "Q3", "Q4"])
             pvt18 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('BRAND == "HELLER"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
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
#TAB 5: Customer category
with tab5:
#Top Down Customer details Table
      left_column, right_column= st.columns(2)
#BAR CHART Customer List
      with left_column:
             st.subheader(":radio: Top 10 Customer_:blue[Invoice Qty]:")            
             customer_line = (filter_df.query('BRAND != "C66 SERVICE"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query(
                             'Inv_Month != "Cancel"').query('BRAND != "SOLDERSTAR"').query('BRAND != "C66 SERVICE"').query(
                             'BRAND != "LOCAL SUPPLIER"').query('BRAND != "SHINWA"').query('BRAND != "SIGMA"').groupby(
                             by=["Customer Name"])[["Item Qty"]].sum().sort_values(by="Item Qty", ascending=False).head(10))
# 生成颜色梯度
             colors = px.colors.sequential.Blues[::-1]  # 将颜色顺序反转为从深到浅
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
                  yaxis=dict(title="Customer Name"),
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
##################################################    
      with right_column:
             st.subheader(":money_with_wings: Top 10 Customer_:orange[Invoice Amount]:")            
             customer_qty_line = (filter_df.query('BRAND != "C66 SERVICE"').query('Inv_Yr != "TBA"').query(
                                 'Inv_Month != "TBA"').query('Inv_Month != "Cancel"').groupby(
                                 by=["Customer Name"])[["Before tax Inv Amt (HKD)"]].sum().sort_values(
                                 by="Before tax Inv Amt (HKD)", ascending=False).head(10))
# 生成颜色梯度
             colors = px.colors.sequential.Oranges[::-1]  # 将颜色顺序反转为从深到浅
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
                  height=400,
                  yaxis=dict(title="Customer Name"),
                  xaxis=dict(title="Before tax Inv Amt (HKD)"),)
             fig_customer_inv_qty.update_layout(font=dict(family="Arial", size=15))
             fig_customer_inv_qty.update_traces(
                  textposition="inside",
                  marker_line_color="black",
                  marker_line_width=2,
                  opacity=1,showlegend=False,
                  )
             # 显示图表
             st.plotly_chart(fig_customer_inv_qty, use_container_width=True)
       
      row2_left_column, row2_right_column= st.columns(2)
      with row2_left_column:
             st.subheader(":medal: :orange[Top Customer Purchase List]_Inv Amount& Qty:")
             pvt12 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region","Ordered_Items"],
                    values=["Item Qty","Before tax Inv Amt (HKD)"],
             aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
             st.dataframe(pvt12.style.format("{:,}"), use_container_width=True)
      with row2_right_column:
             st.subheader(":trophy: :orange[Top Customer List]_Inv Amount& Qty:")
             pvt11 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region"],
                    values=["Item Qty","Before tax Inv Amt (HKD)"],
             aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
             st.dataframe(pvt11.style.format("{:,}"), use_container_width=True)
#############################################################################################################################################################################################################


     
      
############################################################################################################################################################################################################
#success info
#st.success("Executed successfully")
#st.info("This is an information")
#st.warning("This is a warning")
#st.error("An error occured")

 
 
 

