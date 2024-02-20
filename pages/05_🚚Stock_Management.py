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
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"C:\Users\ArthurChan\OneDrive\VS Code\PythonProject_ESE\Sample Excel")

df_south = pd.read_excel(
               io='south_stock_list.xlsx',engine= 'openpyxl',sheet_name='Stock_list', skiprows=0, usecols='A:AS',nrows=10000,)

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
# BAR CHART of SOUTH Instock MANAGEMENT
             st.header("🛒 Instock:")
             stockrow1_a, stockrow1_b= st.columns(2) 
             with stockrow1_a:             
                     df_instock = df_south.query('Stock_Status == "Instock"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
# 按照要求定义颜色顺序
                     color_order = ["有定金+客戶交期", "有定金+无客戶交期", "无定金+有客戶交期", "无定金+无客戶交期"]



# 使用plotly绘制柱状图           
                     brand_instock = px.bar(df_instock, x="Item", y="Machine_QTY", color="Delivery_Status", 
                            color_discrete_sequence=["silver", "orange", "pink", "lightgreen"], 
                            category_orders={"Delivery_Status": color_order}, text_auto='.3s')

# 將barmode設置為"group"以顯示多條棒形圖
                     brand_instock.update_layout(barmode='group')
                     brand_instock.update_traces(marker_line_color='black', textposition='inside', marker_line_width=2,opacity=1)


# 更改字體和label
                     brand_instock.update_layout(font=dict(family="Arial", size=15, color="black"), 
                                         xaxis=dict(title=dict(text="Item", font=dict(size=15))), 
                                         yaxis=dict(title=dict(text="Machine_QTY", font=dict(size=15),)))

# 将图例放在底部
                     brand_instock.update_layout(legend=dict(orientation="h", font=dict(size=17), yanchor="bottom", 
                                    y=1.02, xanchor="right", x=1))

# 绘制图表
                     st.plotly_chart(brand_instock, use_container_width=True)


#####################################################################################
# BAR CHART of SOUTH Incoming STOCK MANAGEMENT
             stockrow2_a, stockrow2_b= st.columns(2) 
             with stockrow2_a:                     
                     st.header("🚢 Incoming_Stock_:blue[WITH]_YAMAHA_Schedule:")
                     
                     df_incoming = df_south.query('Stock_Status == "Incoming_Stock_With_YAMAHA_Schedule"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
 #      sort_Month_order = ["4", "5", "6", "7", "8", "9", "10","11","12", "1", "2", "3"]

# 使用plotly绘制柱状图           
                     incoming_stock = px.bar(df_incoming, x="Item", y="Machine_QTY", color="Delivery_Status", text_auto='.3s')
# 更改字體和label
                     incoming_stock.update_layout(font=dict(family="Arial", size=13.5, color="black"))
                     incoming_stock.update_traces(marker_line_color='black', textposition='outside', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock.update_layout(barmode='group')

# 将图例放在底部
                     incoming_stock.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

# 绘制图表
                     st.plotly_chart(incoming_stock, use_container_width=True)

#####################################################################################
# BAR CHART of SOUTH STOCK MANAGEMENT
             stockrow3_a, stockrow3_b= st.columns(2) 
             with stockrow3_a:      
                     st.header("📅 Incoming_Stock_:red[NO]_YAMAHA_Schedule:")
                     
                     df_incoming = df_south.query('Stock_Status == "Incoming_Stock_No_YAMAHA_Schedule"').groupby(by=["Delivery_Status",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
 #      sort_Month_order = ["4", "5", "6", "7", "8", "9", "10","11","12", "1", "2", "3"]

# 使用plotly绘制柱状图           
                     incoming_stock = px.bar(df_incoming, x="Item", y="Machine_QTY", color="Delivery_Status", text_auto='.3s')
# 更改字體和label
                     incoming_stock.update_layout(font=dict(family="Arial", size=13.5, color="black"))
                     incoming_stock.update_traces(marker_line_color='black', textposition='outside', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
                     incoming_stock.update_layout(barmode='group')

# 将图例放在底部
                     incoming_stock.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

# 绘制图表
                     st.plotly_chart(incoming_stock, use_container_width=True)




 
