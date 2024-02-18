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

#Title
with title_row1:
     st.title(":star2: YAMAHA Sales Target_JPY")
#Text Credit
     st.write("by Arthur Chan")
# 加載圖片
image_path = 'LINE.jpg'
#image_path = '/Users/arthurchan/Downloads/Sample/LINE.jpg'
#image_path = '/Users/arthurchan/Downloads/Sample/LINE.jpg'
image = Image.open(image_path)

# 設置目標寬度和高度
target_width = 1000
target_height = 200

# 縮小圖片
resized_image = image.resize((target_width, target_height))

with title_row2:

# 在Streamlit應用程式中顯示縮小後的圖片
     st.image(resized_image, use_column_width=False, output_format='PNG')

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
# Sidebar Slider
     
#     df_sales_target["Year"] = df_sales_target["Year"].astype(str)
 
#start_yr, end_yr
#     df_yr= st.select_slider('Select a range of year',
#            options=df_sales_target["Year"].unique(), value=("2024"))
     
#########################################################################################     
# Filter of year of YAMAHA Sales Target

df_sales_target = pd.read_excel(
          io='Monthly_report_for_edit.xlsm',
          engine='openpyxl',
          sheet_name='YAMAHA_Sales_Target',
          skiprows=0,
          usecols='A:E',
          nrows=10000) 
     
yr_sales_target = st.multiselect(
       "Select the Year of YAMAHA Sales Target",
        options=df_sales_target["Year"].unique(),
        default=[2024],
        )
     
if not yr_sales_target:
      df_sales_target2 = df_sales_target.copy()
     
else:
      df_sales_target2 = df_sales_target[df_sales_target["Year"].isin(yr_sales_target)]


########################################################################################

     
region_sales = df_sales_target2.groupby(by=["Region"]).agg({"Sales_Target": "sum", "Current Achievement": "sum"}
                     ).sort_values(by="Sales_Target", ascending=False)
     
# 按特定顺序排列Region
region_order = ["SOUTH", "EAST", "NORTH", "WEST"]
region_sales = region_sales.reindex(region_order)

     
fig_target = go.Figure()

fig_target.add_trace(
          go.Bar(
               x=region_sales["Current Achievement"],
               y=region_sales.index,
               name="Current Achievement",
               orientation="h",
               marker=dict(color='orange')))


fig_target.add_trace(
          go.Bar(
               x=region_sales["Sales_Target"],
               y=region_sales.index,
               name="Sales_Target",
               orientation="h",
               marker=dict(color='blue'))) 

     
fig_target.update_layout(
          height=400,
          yaxis=dict(title="Region"),
          xaxis=dict(title="Value"),
          barmode='group')
     
fig_target.update_layout(font=dict(family="Arial", size=15, color="black"),paper_bgcolor='lightyellow')
fig_target.update_traces(
          texttemplate='%{x:.3s}',
          textposition="outside",
          marker_line_color="black",
          marker_line_width=2,opacity=1)
     
fig_target.update_yaxes(categoryorder='array', categoryarray=region_order[::-1])  # 反转顺序
     
st.plotly_chart(fig_target, use_container_width=True)
     
     # 计算每个Region的百分比
region_sales['Percentage'] = region_sales['Current Achievement'] / region_sales['Sales_Target'] * 100

# 创建四个饼图
fig_pie_south = go.Figure(go.Pie(
          labels=["Current Achievement", "Sales Target"],
          values=[region_sales.loc["SOUTH", 'Current Achievement'], region_sales.loc["SOUTH", 'Sales_Target']],
          name="SOUTH",
          hole=0.4,
          marker=dict(colors=['orange', 'blue']),
          textposition='outside', textinfo='label+percent', marker_line_width=1,opacity=1))
     
fig_pie_east = go.Figure(go.Pie(
          labels=["Current Achievement", "Sales Target"],
          values=[region_sales.loc["EAST", 'Current Achievement'], region_sales.loc["EAST", 'Sales_Target']],
          name="EAST",
          hole=0.4,
          marker=dict(colors=['orange', 'blue']),
          textposition='outside', textinfo='label+percent',marker_line_width=1,opacity=1))
     
fig_pie_north = go.Figure(go.Pie(
          labels=["Current Achievement", "Sales Target"],
          values=[region_sales.loc["NORTH", 'Current Achievement'], region_sales.loc["NORTH", 'Sales_Target']],
          name="NORTH",
          hole=0.4,
          marker=dict(colors=['orange', 'blue']),
          textposition='outside', textinfo='label+percent',marker_line_width=1,opacity=1))
     
fig_pie_west = go.Figure(go.Pie(
          labels=["Current Achievement", "Sales Target"],
          values=[region_sales.loc["WEST", 'Current Achievement'], region_sales.loc["WEST", 'Sales_Target']],
          name="WEST",
          hole=0.4,
          marker=dict(colors=['orange', 'blue']),
          textposition='outside', textinfo='label+percent', marker_line_width=1,opacity=1))

#############################

#################################
# 更新布局
fig_pie_south.update_layout(height=400, showlegend=False, paper_bgcolor='rgba(255,165,0,0.3)',font=dict(family="Arial", size=18, color="black"))
fig_pie_east.update_layout(height=400, showlegend=False, paper_bgcolor='rgba(0,150,255,0.1)',font=dict(family="Arial", size=18, color="black"))
fig_pie_north.update_layout(height=400, showlegend=False,paper_bgcolor='khaki',font=dict(family="Arial", size=18, color="black"))
fig_pie_west.update_layout(height=400, showlegend=False, paper_bgcolor='lightgreen',font=dict(family="Arial", size=18, color="black"))

# 显示饼图
sales_target_south, sales_target_east, sales_target_north, sales_target_west = st.columns(4)
with sales_target_south:
          st.header("SOUTH")
          st.plotly_chart(fig_pie_south, use_container_width=True)
with sales_target_east:
          st.header("EAST")
          st.plotly_chart(fig_pie_east, use_container_width=True)
with sales_target_north:
          st.header("NORTH")
          st.plotly_chart(fig_pie_north, use_container_width=True)
with sales_target_west:
          st.header("WEST")
          st.plotly_chart(fig_pie_west, use_container_width=True)
