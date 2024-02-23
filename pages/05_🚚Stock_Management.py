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

st.title("🌱 Stock Management")
#Text Credit
st.write("by Arthur Chan")

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
######################################################################################################
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#SOUTH STOCK DATA BASE
os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"C:\Users\ArthurChan\OneDrive\VS Code\PythonProject_ESE\Sample Excel")

df_south = pd.read_excel(
               io='stock_list.xlsx',engine= 'openpyxl',sheet_name='Stock_list', skiprows=0, usecols='A:AV',nrows=10000,)

# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
font-size: 28px;
}
</style>
"""

st.write(font_css, unsafe_allow_html=True)
tab1, tab2, tab3= st.tabs(["📙 SOUTH","📘 EAST","📗 NORTH"])
#########################################################################################
with tab1:
             stock_row11, stock_row12, stock_row13, stock_row14, stock_row15= st.columns(5)
             with stock_row13:
                     st.header("🛒 :green[Instock]:")
                     
# BAR CHART of SOUTH Instock MANAGEMENT
             stockrow1_a, stockrow1_b= st.columns(2) 
             with stockrow1_a:
                     st.subheader(":moneybag: :green[有/无定金简易分类]")             
                     df_instock = df_south.query('Stock_Status == "Instock"').groupby(by=["Deposit",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by=["Machine_QTY"], ascending=[False])
# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]

# 使用plotly绘制柱状图           
                     brand_instock = px.bar(df_instock, x="Item", y="Machine_QTY", color="Deposit", 
                            color_discrete_sequence=["blue", "red"], category_orders={"Deposit": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     brand_instock.update_layout(barmode='group')
                     brand_instock.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     brand_instock.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     brand_instock.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.5))       
                     
# 添加背景色
                     background_color = 'lightgreen'
                     x_range = len(df_instock['Item'].unique())
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
                     brand_instock.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(brand_instock, use_container_width=True)

##############################################################################################################################
             with stockrow1_b:  
                     st.subheader(":card_index_dividers: :green[进阶分类]")             
                     df_instock = df_south.query('Stock_Status == "Instock"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by=["Machine_QTY"], ascending=[False])
# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]

# 使用plotly绘制柱状图           
                     brand_instock = px.bar(df_instock, x="Item", y="Machine_QTY", color="Delivery_Status", 
                            color_discrete_sequence=["orange", "lightyellow", "pink", "lightgreen"], category_orders={"Delivery_Status": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     brand_instock.update_layout(barmode='group')
                     brand_instock.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     brand_instock.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     brand_instock.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.6))       
                     
# 添加背景色
                     background_color = 'lightgreen'
                     x_range = len(df_instock['Item'].unique())
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
                     brand_instock.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(brand_instock, use_container_width=True)

####################################################################################
             stock_row2a, stock_row2b, stock_row2c= st.columns(3)
             with stock_row2b:
                     with st.expander(":point_right: Click to expand/ hide data"):
                             pvt = df_south.query('Stock_Status == "Instock"').round(0).pivot_table(
                                     index=["Item","Delivery_Status","Customer_Reserved","Customer_Reserved_Contract_No."],
                                     columns=["客戶送貨期"], 
                                     values=["Machine_QTY"],
                                     aggfunc="sum",
                                     fill_value=0,
                                     margins=True,
                                     margins_name="Total",
                                     observed=True)

       #使用applymap方法應用格式化
                             pvt = pvt.applymap('{:,.0f}'.format)
                             html = pvt.to_html(classes='table table-bordered', justify='center')
                             html = html.replace('<th>-</th>', '<th style="background-color: lightgreen">-</th>')


# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
                             html = html.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把所有數值等於或少於0的數值的顏色設為紅色
                             html = html.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
                             html = f'<div style="zoom: 0.9;">{html}</div>'
                             st.markdown(html, unsafe_allow_html=True)           
 
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
                             csv1 = pvt.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
                             st.download_button(label='Download Table', data=csv1, file_name='Instock_Machine.csv', mime='text/csv')
             
             st.divider()
#####################################################################################
             stock_row11, stock_row12, stock_row13= st.columns(3)
             with stock_row12:
                     st.header("🚢 :orange[Incoming_STK]_:blue[有]_YAMAHA交期:")

# BAR CHART of SOUTH Incoming STOCK MANAGEMENT with YAMAHA shipping schedule
             stockrow2_a, stockrow2_b= st.columns(2) 
             with stockrow2_a:                     
                     st.subheader(":moneybag: :blue[有/无定金简易分类]")
                     
                     df_incoming = df_south.query('Stock_Status == "Incoming_Stock_With_YAMAHA_Schedule"').groupby(by=["Deposit",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)

# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]


# 使用plotly绘制柱状图           
                     incoming_stock = px.bar(df_incoming, x="Item", y="Machine_QTY", color="Deposit", 
                            color_discrete_sequence=["blue", "red"], category_orders={"Deposit": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock.update_layout(barmode='group')
                     incoming_stock.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     incoming_stock.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     incoming_stock.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.5))       
                     
# 添加背景色
                     background_color = 'lightblue'
                     x_range = len(df_incoming['Item'].unique())
                     background_shapes = [dict(
                             type='rect',
                             xref='x',
                             yref='paper',
                             x0=i - 0.5,
                             y0=0,
                             x1=i + 0.5,
                             y1=1,
                             fillcolor=background_color,
                             opacity=0.2,
                             layer='below',
                             line=dict(width=5)) for i in range(x_range)]

# 绘制图表
                     incoming_stock.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(incoming_stock, use_container_width=True)


             with stockrow2_b:                     
                     st.subheader(":card_index_dividers: :blue[进阶分类]")
                     
                     df_incoming = df_south.query('Stock_Status == "Incoming_Stock_With_YAMAHA_Schedule"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)

# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]


# 使用plotly绘制柱状图           
                     incoming_stock = px.bar(df_incoming, x="Item", y="Machine_QTY", color="Delivery_Status", 
                            color_discrete_sequence=["orange", "lightyellow", "pink", "lightgreen"], category_orders={"Delivery_Status": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock.update_layout(barmode='group')
                     incoming_stock.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     incoming_stock.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     incoming_stock.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.6))       
                     
# 添加背景色
                     background_color = 'lightblue'
                     x_range = len(df_incoming['Item'].unique())
                     background_shapes = [dict(
                             type='rect',
                             xref='x',
                             yref='paper',
                             x0=i - 0.5,
                             y0=0,
                             x1=i + 0.5,
                             y1=1,
                             fillcolor=background_color,
                             opacity=0.2,
                             layer='below',
                             line=dict(width=5)) for i in range(x_range)]

# 绘制图表
                     incoming_stock.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(incoming_stock, use_container_width=True)
#####################################################################################
                     
             stock_row3a, stock_row3b, stock_row3c= st.columns(3)
             with stock_row3b:
                     with st.expander(":point_right: Click to expand/ hide data"):
                             
                             pvt2 = df_south.query('Stock_Status == "Incoming_Stock_With_YAMAHA_Schedule"').round(0).pivot_table(
                                     index=["Item","Delivery_Status","ETA_HK","Customer_Reserved","Customer_Reserved_Contract_No."],
                                     columns=["客戶送貨期"], 
                                     values=["Machine_QTY"],
                                     aggfunc="sum",
                                     fill_value=0,
                                     margins=True,
                                     margins_name="Total",
                                     observed=True)
                             
       #使用applymap方法應用格式化
                             
                             pvt2 = pvt2.applymap('{:,.0f}'.format)              
                             html = pvt2.to_html(classes='table table-bordered', justify='center')
                             html = html.replace('<th>-</th>', '<th style="background-color: lightgreen">-</th>')


# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
                             html = html.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把所有數值等於或少於0的數值的顏色設為紅色
                             html = html.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
                             html = f'<div style="zoom: 0.9;">{html}</div>'
                             st.markdown(html, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
                             csv2 = pvt2.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
                             st.download_button(label='Download Table', data=csv2, file_name='Incoming_Machine_with_schedule.csv', mime='text/csv')

             st.divider()

#####################################################################################
             stock_row3a, stock_row3b, stock_row3c= st.columns(3)
             with stock_row3b:
                     st.header("🚢 :orange[Incoming_STK]_:red[无]_YAMAHA交期:")
# BAR CHART of SOUTH STOCK MANAGEMENT No YAMAHA shipping schedule
             stockrow3_a, stockrow3_b= st.columns(2) 
             with stockrow3_a:      
                     st.subheader(":moneybag: :red[有/无定金简易分类]:")
                     
                     df_incoming2 = df_south.query('Stock_Status == "Incoming_Stock_No_YAMAHA_Schedule"').groupby(by=["Deposit",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]

# 使用plotly绘制柱状图           
                     incoming_stock2 = px.bar(df_incoming2, x="Item", y="Machine_QTY", color="Deposit", 
                            color_discrete_sequence=["blue", "red"], category_orders={"Deposit": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock2.update_layout(barmode='group')
                     incoming_stock2.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     incoming_stock2.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     incoming_stock2.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.5))       
                     
# 添加背景色
                     background_color = 'pink'
                     x_range = len(df_incoming2['Item'].unique())
                     background_shapes = [dict(
                             type='rect',
                             xref='x',
                             yref='paper',
                             x0=i - 0.5,
                             y0=0,
                             x1=i + 0.5,
                             y1=1,
                             fillcolor=background_color,
                             opacity=0.2,
                             layer='below',
                             line=dict(width=5)) for i in range(x_range)]

# 绘制图表
                     incoming_stock2.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(incoming_stock2, use_container_width=True)

             with stockrow3_b:      
                     st.subheader(":card_index_dividers: :red[进阶分类]:")
                     
                     df_incoming2 = df_south.query('Stock_Status == "Incoming_Stock_No_YAMAHA_Schedule"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
# 按照要求定义颜色顺序
                     color_order = ["有定金+有預留客戶", "有定金+无客戶送貨期", "无定金+有預留客戶", "无定金+无預留客戶"]

# 使用plotly绘制柱状图           
                     incoming_stock2 = px.bar(df_incoming2, x="Item", y="Machine_QTY", color="Delivery_Status", 
                            color_discrete_sequence=["orange", "lightyellow", "pink", "lightgreen"], category_orders={"Delivery_Status": color_order},
                            text_auto='.0s')

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock2.update_layout(barmode='group')
                     incoming_stock2.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     incoming_stock2.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=12))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=13),)))

# 将图例放在底部
                     incoming_stock2.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=0.5))       
                     
# 添加背景色
                     background_color = 'pink'
                     x_range = len(df_incoming2['Item'].unique())
                     background_shapes = [dict(
                             type='rect',
                             xref='x',
                             yref='paper',
                             x0=i - 0.5,
                             y0=0,
                             x1=i + 0.5,
                             y1=1,
                             fillcolor=background_color,
                             opacity=0.2,
                             layer='below',
                             line=dict(width=5)) for i in range(x_range)]

# 绘制图表
                     incoming_stock2.update_layout(shapes=background_shapes, showlegend=True)
                     st.plotly_chart(incoming_stock2, use_container_width=True)
#####################################################################################
             stock_row4a, stock_row4b, stock_row4c= st.columns(3)
             with stock_row4b:
                     with st.expander(":point_right: Click to expand/ hide data"):
                             pvt3 = df_south.query('Stock_Status == "Incoming_Stock_No_YAMAHA_Schedule"').round(0).pivot_table(
                                     index=["Item","Delivery_Status","ETA_HK","Customer_Reserved","Customer_Reserved_Contract_No."],
                                     
                                     values=["Machine_QTY"],
                                     aggfunc="sum",
                                     fill_value=0,
                                     margins=True,
                                     margins_name="Total",
                                     observed=True)

       #使用applymap方法應用格式化
                             pvt3 = pvt3.applymap('{:,.0f}'.format)
                             html = pvt3.to_html(classes='table table-bordered', justify='center')
                             html = html.replace('<th>-</th>', '<th style="background-color: lightgreen">-</th>')


# 把total值的那行的背景顏色設為黃色，並將字體設為粗體
                             html = html.replace('<tr>\n      <th>Total</th>', '<tr style="background-color: yellow;">\n      <th style="font-weight: bold;">Total</th>')
# 把所有數值等於或少於0的數值的顏色設為紅色
                             html = html.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 放大pivot table
                             html = f'<div style="zoom: 0.9;">{html}</div>'
                             st.markdown(html, unsafe_allow_html=True)
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
                             csv3 = pvt3.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
                             st.download_button(label='Download Table', data=csv3, file_name='Incoming_Machine_no_schedule.csv', mime='text/csv')
 
#####################################################################################################################################################################
#with tab2:
