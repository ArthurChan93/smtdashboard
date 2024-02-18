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
df_south = pd.read_excel(
               io='south_stock_list.xlsx',engine= 'openpyxl',sheet_name='Stock_list', skiprows=0, usecols='A:AP',nrows=10000,)

# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
font-size: 28px;
}
</style>
"""

st.write(font_css, unsafe_allow_html=True)
tab1, tab2, tab3= st.tabs(["🟠 SOUTH","🔵 EAST","🟢 NORTH"])

with tab1:
      # BAR CHART of SOUTH STOCK MANAGEMENT
       left_column, right_column = st.columns(2)
       with left_column:
             st.subheader("🛄 In Stock:")
             
             df_instock = df_south.query('ETA_MONTH == "已经到货"').groupby(by=["Deposit",
                            "Item"], as_index=False)["Machine_QTY"].sum().sort_values(by="Machine_QTY", ascending=False)
       
 #      sort_Month_order = ["4", "5", "6", "7", "8", "9", "10","11","12", "1", "2", "3"]

# 使用plotly绘制柱状图           
             brand_instock = px.bar(df_instock, x="Item", y="Machine_QTY", color="Deposit", text_auto='.3s')

# 更改顏色
             colors = {"Yes": "lightgreen","No": "pink"}
             for trace in brand_instock.data:
                    brand_color = trace.name.split("=")[-1]
                    trace.marker.color = colors.get(brand_color, "blue")

# 更改字體和label
             brand_instock.update_layout(font=dict(family="Arial", size=13.5, color="black"))
             brand_instock.update_traces(marker_line_color='black', textposition='outside', marker_line_width=2,opacity=1)

# 將barmode設置為"group"以顯示多條棒形圖
             brand_instock.update_layout(barmode='group')

# 将图例放在底部
             brand_instock.update_layout(legend=dict(orientation="h",font=dict(size=14), yanchor="bottom", y=1.02, xanchor="right", x=1))

# 绘制图表
             st.plotly_chart(brand_instock, use_container_width=True)






       with right_column:
             st.subheader("🚢 Incoming_Stock")





 
