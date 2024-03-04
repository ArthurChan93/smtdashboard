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
import altair as alt
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
          df_stockstatus = df_south.query('Stock_Status != "Shipped_Stock"').groupby(by=["Stock_Status",
                    "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by=["Machine_QTY"], ascending=[False])

          df_stockstatus1 = pd.DataFrame(df_stockstatus)
# 过滤和排序数据
          df_filtered = df_stockstatus1[df_stockstatus1['Stock_Status'] != 'Shipped_Stock']
          df_filtered = df_filtered.groupby(['Stock_Status', 'Item'], as_index=False)['Machine_QTY'].sum()
          df_filtered = df_filtered.sort_values(by='Machine_QTY', ascending=False)

# 设置颜色和标签
          color_scale = alt.Scale(domain=['Incoming_Stock_With_YAMAHA_Schedule', 'Incoming_Stock_No_YAMAHA_Schedule', 'Instock'],
                        range=['orange', 'lightblue', 'lightgreen'])
          legend_title = 'Stock Status'

# 创建Altair图表
          chart = alt.Chart(df_filtered).mark_bar().encode(
                  x=alt.X('Machine_QTY:Q', title='Machine Quantity'),
                  y=alt.Y('Item:N', sort='-x'), 
                  color=alt.Color('Stock_Status:N', scale=color_scale, legend=alt.Legend(title=legend_title, orient ='top')),
                  tooltip=alt.Tooltip('Machine_QTY:Q')).properties(
                          width=500,
                          height=300)

# 在柱形图内显示数值
          text = chart.mark_text(
                  align='left',
                  baseline='middle',
                  dx=3,fontSize=18).encode(
                  text=alt.Text('Machine_QTY:Q', format='.0f'),
                  color=alt.value('black'))

# 设置图表的布局
          chart_with_text = chart + text

# 将图表和文本合并，并在Streamlit中显示
          st.altair_chart(chart_with_text, use_container_width=True)

















          st.subheader("表格明細:")                
          pvt = df_south.query('Stock_Status != "Shipped_Stock"').round(0).pivot_table(
                              index=["Item"],
                              columns=["Stock_Status"], 
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
          html = html.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')

# 放大pivot table
          html = f'<div style="zoom: 0.9;">{html}</div>'
          st.markdown(html, unsafe_allow_html=True)           
 
# 使用streamlit的download_button方法提供一個下載數據框為CSV檔的按鈕
          csv1 = pvt.to_csv(index=True,float_format='{:,.0f}'.format).encode('utf-8')
          st.download_button(label='Download Table', data=csv1, file_name='Instock_Machine.csv', mime='text/csv')
             
          st.divider()