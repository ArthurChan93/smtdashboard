# 導入所需的庫
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import os
import streamlit.components.v1 as com #用frame show animation
############################################################################################################################################################################################################
 
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")
# 設置標題
st.title(":robot_face: Sales Projection")

#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
#Text
st.write("by Arthur Chan")
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

######################################################################################################
# 創建一個側邊欄，用於收集用戶的輸入
sidebar = st.sidebar
left_col, mid_col, right_col, col_04, col_05 = st.columns(5)

###########################################################################################################################################################################################################

#os.chdir(r"/Users/arthurchan/Downloads/Sample")
df = pd.read_excel(
io='Sample_excel.xlsx',engine= 'openpyxl',sheet_name='sheet 1', skiprows=0, usecols='A:AO',nrows=10000,).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
 
with mid_col:
    st.subheader(":orange_book: SMT Inv Data for reference")
    pvt = df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(values=["Before tax Inv Amt (HKD)","G.P.  (HKD)"],index=["FY_INV"],
          aggfunc="sum",fill_value=0).sort_index(axis=0, ascending=False)
    
    st.dataframe(pvt.style.highlight_min(color = 'white', axis = 0).format("HKD{:,}"), use_container_width=True)
      
#    html = pvt.map('{:,.1f}'.format, use_container_width=True).to_html()
# 將你想要變色的column header找出來，並加上顏色
#    html1_a = html.replace('<th>Total</th>', '<th style="background-color: yellow">Total</th>')
# 使用Streamlit的markdown來顯示HTML表格              
#    st.markdown(html1_a, unsafe_allow_html=True)

with right_col:
    #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/5c29f284-4d9d-4a86-a754-136cafc5973d/DrlL2cbHPZ.json"></iframe> 去lottie copy iframe code
    com.iframe("https://lottie.host/embed/5c29f284-4d9d-4a86-a754-136cafc5973d/DrlL2cbHPZ.json",height=500, width= 1000)

###########################################################################################################################################################################################################

with left_col:
    st.subheader(" :point_down:  Please select & entry:")

# 在側邊欄中添加一些數字輸入框，讓用戶輸入初始財年, 財年的revenue, gross profit, cost, cash flow，並設置起始值
    fy = st.selectbox(":calendar: :orange[Initial FY] for projection", options=range(2025, 2035), format_func=lambda x: f"FY{x}")
    rev = st.number_input(":moneybag: :orange[Invoice Amount(HKD)]_Initial FY", value=0, format='%d')
    gp = st.number_input(":money_mouth_face: :orange[Gross profit(HKD)]_Initial FY", value=0, format='%d')
    cost = st.number_input(":money_with_wings: :orange[Cost(HKD)]_Initial FY", value=0, format='%d')
    cf = st.number_input(":heavy_dollar_sign: :orange[Cash flow(HKD)]_Initial FY", value=0, format='%d')

# 在側邊欄中添加一些滑動條，讓用戶調整每個財年的revenue增長率，投資報酬增長率，cost增長率，並設置起始值
st.sidebar.subheader("Customization:")
igr = sidebar.slider("Annual :blue[Inv Amount] Growth _(Percentage)", min_value=-100, max_value=100, value=0)
rr = sidebar.slider("Annual :blue[G.P Growth]_(Percentage)", min_value=-100, max_value=100, value=0)
cgr = sidebar.slider("Annual :blue[Cost Increase]_(Percentage)", min_value=-100, max_value=100, value=0)

# 在側邊欄中添加一個數字輸入框，讓用戶輸入預測財年期間
n = sidebar.number_input("Projection(Year)", min_value=1, max_value=10, value=5)

# 在側邊欄中添加一個按鈕，讓用戶按下後進行業續計算
st.sidebar.divider()
button = sidebar.button(":white_check_mark: Start projection")

# 創建一個數據框，用於存儲初始財年的數據
df = pd.DataFrame({"revenue": [rev], "gross profit": [gp], "cost": [cost], "cash flow": [cf]}, index=[fy])

# 創建一個函數，用於根據增長率計算下一財年的數據
def next_year(df, igr, rr, cgr):
    # 獲取當前財年的數據
    rev = df["revenue"].iloc[-1]
    gp = df["gross profit"].iloc[-1]
    cost = df["cost"].iloc[-1]
    cf = df["cash flow"].iloc[-1]
    # 根據增長率計算下一財年的數據
    rev_next = rev * (1 + igr / 100)
    gp_next = gp * (1 + igr / 100)
    cost_next = cost * (1 + cgr / 100)
    cf_next = cf * (1 + rr / 100)
    # 返回下一財年的數據
    return rev_next, gp_next, cost_next, cf_next

# 如果用戶沒有按下按鈕，則顯示一個提示信息，告訴用戶如何進行業續計算
if not button:
    col_01, col_02, col_03, col_04 = st.columns(4)
    with col_01:
        st.info(":point_left: After entering the above information, adjust the parameter in the sidebar, then press the 'Start Prediction' button.")
        labels = ["Revenue", "Gross Profit", "Cost", "Cash Flow"]
        sizes = [rev, gp, cost, cf]
        colors = ["blue", "yellow", "red", "green"]
        fig = px.pie(values=sizes, names=labels, color=colors, title="Pie chart_Initial FY")
        st.plotly_chart(fig)

# 如果用戶按下按鈕，則進行業續計算    
else:
    st.header(":bar_chart: Trend Chart_FY to FY")
    # 用一個循環，根據預測財年期間，計算每一年的數據，並添加到數據框中
    for i in range(n):
        # 計算下一財年的數據
        rev_next, gp_next, cost_next, cf_next = next_year(df, igr, rr, cgr)
        # 獲取下一財年的名稱
        fy_next = df.index[-1] + 1
        # 將下一財年的數據添加到數據框中
        df.loc[fy_next] = [rev_next, gp_next, cost_next, cf_next]
    
    col_01, col_02, col_03 = st.columns(3)

    # 顯示一個折線圖和柱狀圖的混合圖，用於展示預測年間的趨勢圖，並在圖上添加標記和數值，並設置圖顏色為藍色和橙色
    fig, ax = plt.subplots()
    ax.plot(df.index, df["revenue"], label=["revenue"], color="blue", marker="o")
    ax.plot(df.index, df["cost"], label=["cost"], color="orange", marker="o")  
    ax.plot(df.index, df["gross profit"], label="gross profit", color="green", marker="o")
    ax.bar(df.index, df["cash flow"], label="cash flow", color="purple")
    ax.set_xlabel("Year")
    ax.set_ylabel("HKD: 100million")
    
    # 添加数值标签
    ax.legend()
    for x, y in zip(df.index, df["revenue"]):
        ax.text(x, y, f"{y // 1000000:,.0f}m", ha="center", va="bottom")
    for x, y in zip(df.index, df["cost"]):
        ax.text(x, y, f"{y // 1000000:,.0f}m", ha="center", va="bottom")
    for x, y in zip(df.index, df["gross profit"]):
        ax.text(x, y, f"{y // 1000000:,.0f}m", ha="center", va="bottom")

    with col_01:
    # 顯示一個提示信息，告訴用戶業續計算已完成
     st.success("Calculation Completed!")
     st.pyplot(fig)
    # 顯示數據框，並將數據格式化為會計格式，小數點後兩位
     st.dataframe(df.style.format("{:,.0f}"), use_container_width=True)


