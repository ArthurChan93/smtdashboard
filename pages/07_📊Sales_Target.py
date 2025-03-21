import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# 设置工作目录并读取数据
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")
# 读取第一个数据源
df = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='raw_sheet',
    skiprows=0,
    usecols='A:AU',
    nrows=100000
).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query(
    'FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query(
    'FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query(
    'Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query(
    'Inv_Month != "Cancel"')

# 读取第二个数据源
df2 = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='SMT_Internal_Sales_Target',
    skiprows=0,
    usecols='A:G',
    nrows=100000)

# 数据预处理
df = df.dropna(subset=['Inv_Yr', 'Inv_Month'])
df['Region'] = df['Region'].replace(['', 'C66 N/A'], pd.NA).dropna()

# 设置页面布局
st.set_page_config(layout="wide")

# ========== 左边栏：Filter设置 ==========
with st.sidebar:
    st.header("Monthly Report")
    
    # Filter设置
    selected_fy = st.selectbox(
        'FY',
        options=df['FY_INV'].unique(),
        index=list(df['FY_INV'].unique()).index('FY 24/25') if 'FY 24/25' in df['FY_INV'].unique() else 0
    )
    selected_yr = st.multiselect('Inv Yr', options=df['Inv_Yr'].unique())
    selected_month = st.multiselect('Inv Month', options=df['Inv_Month'].unique())
    region_options = [r for r in df['Region'].unique() if pd.notna(r)]
    selected_region = st.multiselect('Region', options=region_options)

    st.header("Sales Target_FY")
    selected_target_fy = st.selectbox(
        'FY_INV (Target)',
        options=df2['FY_INV'].unique(),
        index=list(df2['FY_INV'].unique()).index('FY 24/25') if 'FY 24/25' in df2['FY_INV'].unique() else 0
    )

# ========== 主内容区域 ==========
main_col1, main_col2 = st.columns(2)

# 左边内容：Monthly Report Analysis
with main_col1:
    st.markdown(
        "<h3 style='background-color: #D8BFD8; padding: 10px; border-radius: 5px;'>Monthly Report Data</h3>",
        unsafe_allow_html=True
    )
    
    # 应用Filter
    filtered_df = df[
        (df['FY_INV'] == selected_fy) &
        (df['Inv_Yr'].isin(selected_yr if selected_yr else df['Inv_Yr'].unique())) &
        (df['Inv_Month'].isin(selected_month if selected_month else df['Inv_Month'].unique())) &
        (df['Region'].isin(selected_region if selected_region else region_options))
    ]
    
    # 总Inv Amt
    total_inv = filtered_df['Before tax Inv Amt (HKD)'].sum()
    st.metric("Total Inv Amt (HKD)", f"{total_inv:,.2f}")

    # 饼图部分
    col1, col2 = st.columns(2)
    with col1:
        cost_center_colors = {'C49': '#90EE90', 'C28': '#87CEEB', 'C66': '#FFB6C1'}
        fig1 = px.pie(filtered_df, names='COST_CENTRE', values='Before tax Inv Amt (HKD)',
                     color='COST_CENTRE', color_discrete_map=cost_center_colors)
        fig1.update_traces(
            textposition='inside', 
            textinfo='percent+label', 
            texttemplate='%{percent:.2%}', 
            textfont_size=28
        )
        fig1.update_layout(
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        fig2 = px.pie(filtered_df, names='Region', values='Before tax Inv Amt (HKD)')
        fig2.update_traces(
            textposition='inside', 
            textinfo='percent+label', 
            texttemplate='%{percent:.2%}', 
            textfont_size=28
        )
        fig2.update_layout(
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Pivot Table
    region_order = ['SOUTH', 'EAST', 'NORTH', 'WEST']
    cost_center_order = ['C49', 'C28', 'C66']
    pivot_df = filtered_df.pivot_table(
        values='Before tax Inv Amt (HKD)',
        index='COST_CENTRE',
        columns='Region',
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name='Total'
    ).reindex(index=cost_center_order + ['Total'], columns=region_order + ['Total'], fill_value=0)
    
    # 样式设置
    styled_df = pivot_df.style.format("{:,.2f}").apply(
        lambda x: ['background: #FFFF00' if x.name == 'Total' else '' for _ in x], axis=1)
    
    st.write(styled_df)
    csv = pivot_df.to_csv().encode('utf-8')
    st.download_button("Download Pivot Table", csv, 'monthly_report.csv', 'text/csv')

# 右边内容：Sales Target Analysis
with main_col2:
    st.markdown(
        "<h3 style='background-color: #FFFFE0; padding: 10px; border-radius: 5px;'>Sales Target</h3>",
        unsafe_allow_html=True
    )
    
    # 应用Filter
    filtered_df2 = df2[df2['FY_INV'] == selected_target_fy]
    
    # 总Sales Target
    total_target = filtered_df2['Total _Sales_Target(Inv Amt HKD)'].sum()
    st.metric("Total Sales Target (HKD)", f"{total_target:,.2f}")

    # 饼图部分
    col3, col4 = st.columns(2)
    with col3:
        fig3 = px.pie(filtered_df2, names='COST_CENTRE', values='Total _Sales_Target(Inv Amt HKD)',
                     color='COST_CENTRE', color_discrete_map=cost_center_colors)
        fig3.update_traces(
            textposition='inside', 
            textinfo='percent+label', 
            texttemplate='%{percent:.2%}', 
            textfont_size=28
        )
        fig3.update_layout(
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        region_colors = {'SOUTH': 'orange', 'EAST': 'blue', 'NORTH': '#D8BFD8', 'WEST': 'green'}
        fig4 = px.pie(filtered_df2.melt(id_vars=['FY_INV', 'COST_CENTRE'], 
                                      value_vars=['SOUTH', 'EAST', 'NORTH', 'WEST'],
                                      var_name='Region', value_name='Sales Target(HKD)'),
                     names='Region', values='Sales Target(HKD)',
                     color='Region', color_discrete_map=region_colors)
        fig4.update_traces(
            textposition='inside', 
            textinfo='percent+label', 
            texttemplate='%{percent:.2%}', 
            textfont_size=28
        )
        fig4.update_layout(
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig4, use_container_width=True)

    # Sales Target Pivot Table
    # 修复销售目标数据透视表
    target_pivot = filtered_df2.melt(
        id_vars=['FY_INV', 'COST_CENTRE'],
        value_vars=['SOUTH', 'EAST', 'NORTH', 'WEST'],
        var_name='Region',
        value_name='Sales Target(HKD)'
    ).pivot_table(
        values='Sales Target(HKD)',
        index='COST_CENTRE',
        columns='Region',
        aggfunc='sum',
        margins=True,
        margins_name='Total'
    ).reindex(index=cost_center_order + ['Total'], columns=region_order + ['Total'], fill_value=0)
    
    # 样式设置
    styled_target = target_pivot.style.format("{:,.2f}").apply(
        lambda x: ['background: #FFFF00' if x.name == 'Total' else '' for _ in x], axis=1)
    
    st.write(styled_target)
    csv2 = target_pivot.to_csv().encode('utf-8')
    st.download_button("Download Target Pivot", csv2, 'sales_target.csv', 'text/csv')

# ========== Regional Comparison ==========
st.markdown(
    "<h2 style='background-color: #FFFF00; text-align: center; padding: 10px; border-radius: 5px; font-size: 24px;'>Regional Comparison</h2>",
    unsafe_allow_html=True
)

cost_centers = sorted(filtered_df2['COST_CENTRE'].unique(), 
                     key=lambda x: ['C49', 'C28', 'C66'].index(x) if x in ['C49', 'C28', 'C66'] else 3)

for cost_center in cost_centers:
    # 标题居中设置
    bg_color = {
        'C49': '#90EE90', 
        'C28': '#87CEEB', 
        'C66': '#FFB6C1'
    }.get(cost_center, 'white')
    
    st.markdown(
        f"<h4 style='background-color: {bg_color}; padding: 10px; border-radius: 5px; text-align: center; font-size: 20px;'>{cost_center}</h4>", 
        unsafe_allow_html=True
    )
    
    # 数据准备
    actual_data = filtered_df[filtered_df['COST_CENTRE'] == cost_center]
    target_data = filtered_df2[filtered_df2['COST_CENTRE'] == cost_center]
    
    merged = pd.DataFrame({
        'Region': region_order,
        'Actual Inv Amt(HKD)': [actual_data[actual_data['Region'] == r]['Before tax Inv Amt (HKD)'].sum() for r in region_order],
        'Sales Target(HKD)': [target_data[r].sum() for r in region_order]
    })
    merged['Difference(HKD)'] = merged['Actual Inv Amt(HKD)'] - merged['Sales Target(HKD)']
    merged['Achievement%'] = (merged['Actual Inv Amt(HKD)'] / merged['Sales Target(HKD)'] * 100).round(2)

    # 生成带达标指示的图表
    fig = go.Figure()
    
    # 添加达标状态标记
    for idx, region in enumerate(merged['Region']):
        actual = merged.loc[idx, 'Actual Inv Amt(HKD)']
        target = merged.loc[idx, 'Sales Target(HKD)']
        percentage = merged.loc[idx, 'Achievement%']
        
        if actual >= target:
            fig.add_annotation(
                x=region,
                y=max(actual, target) * 1.3,
                text=f"✅ 达标，已实现{percentage}%",
                showarrow=False,
                font=dict(color="green", size=14, family="Arial Bold"),
                bgcolor="white",
                bordercolor="green"
            )
        else:
            fig.add_annotation(
                x=region,
                y=max(actual, target) * 1.3,
                text=f"❌ 未达标，已实现{percentage}%",
                showarrow=False,
                font=dict(color="red", size=14, family="Arial Bold"),
                bgcolor="white",
                bordercolor="red"
            )
    
    # Actual Bar
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Actual Inv Amt(HKD)'],
        name='Actual Inv Amt(HKD)',
        marker_color='#1f77b4',
        marker_line_color='black',
        marker_line_width=1,
        text=merged['Actual Inv Amt(HKD)'].apply(lambda x: f"{x/1e6:.1f}M" if x >= 1e6 else f"{x/1e3:.0f}K"),
        textposition='outside',
        textfont=dict(
            size=16,
            family='Arial Bold'
        )
    ))
    
    # Target Bar
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Sales Target(HKD)'],
        name='Sales Target(HKD)',
        marker_color='#FFD700',
        marker_line_color='black',
        marker_line_width=1,
        text=merged['Sales Target(HKD)'].apply(lambda x: f"{x/1e6:.1f}M" if x >= 1e6 else f"{x/1e3:.0f}K"),
        textposition='inside',
        textfont=dict(
            size=16,
            family='Arial Bold'
        )
    ))
    
    # Difference Bar
    colors = ['#2ca02c' if diff >=0 else '#ff7f0e' for diff in merged['Difference(HKD)']]
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Difference(HKD)'].abs(),
        name='Difference(HKD)',
        marker_color=colors,
        marker_line_color='black',
        marker_line_width=1,
        text=merged['Difference(HKD)'].apply(lambda x: f"+{x/1e6:.1f}M" if x >=0 else f"-{abs(x)/1e6:.1f}M"),
        textposition='outside',
        textfont=dict(
            size=14,
            family='Arial Bold'
        )
    ))
    
    fig.update_layout(
        # 移除标题
        #title=None,
        barmode='group',
        yaxis_title="Amount (HKD)",
        xaxis_title="Region",
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
        height=800,
        margin=dict(t=100, b=120)  # 调整顶部边距
    )
    st.plotly_chart(fig, use_container_width=True)


# 幫我設計一個用於streamlit的python程式，順序完成以下要求，給我完整CODE
# 
# 讀取以下位置的excel，不需要用家自行上傳
# 
# os.chdir(r"/Users/arthurchan/Downloads/Sample") 
# 
# #os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")
# 
# 第一個數據源:
# df = pd.read_excel(
# io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='raw_sheet', skiprows=0, usecols='A:AU’,nrows=100000,).query(
# 'Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query(
# 'FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
# 
# Excel內的數據列內容我先給你說明:
# -E列名稱是"FY_INV"，意思是FY
# -F列名稱是"Inv_Yr"，意思是Inv Yr
# -G列名稱是"Inv_Month"，意思是Inv Month
# -R列名稱是"COST_CENTRE"，意思是COST CENTRE
# -S列名稱是"Region"，意思是Region
# -AB列名稱是"Before tax Inv Amt (HKD)"，意思是Inv Amt(HKD)
# 
# 第二個數據源:
# df2 = pd.read_excel(
# io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='SMT_Internal_Sales_Target', skiprows=0, usecols='A:G',nrows=100000)
# 
# -A列名稱是"FY_INV"，意思是FY，此列跟第一個數據源中E列的”FY_INV”是一樣意思
# -B列名稱是"COST_CENTRE"，意思是COST CENTRE，此列跟第一個數據源中R列的”COST_CENTRE”是一樣意思
# -C列名稱是"Total _Sales_Target(Inv Amt HKD)"，意思是Sales Target Amt(HKD)
# -D列名稱是"SOUTH"，此列代表的是第一個數據源中"Region"的其中一個Region"SOUTH"，此列數據都是數字，是代表該Region在該FY中的Sales Target Amt(HKD)
# -E列名稱是"EAST"，此列代表的是第一個數據源中"Region"的其中一個Region""EAST"，此列數據都是數字，是代表該Region在該FY中的Sales Target Amt(HKD)
# -F列名稱是"NORTH"，此列代表的是第一個數據源中"Region"的其中一個Region"NORTH"，此列數據都是數字，是代表該Region在該FY中的Sales Target Amt(HKD)
# -G列名稱是"WEST"，此列代表的是第一個數據源中"Region"的其中一個Region"WEST"，此列數據都是數字，是代表該Region在該FY中的Sales Target Amt(HKD)
# 
# 首先我想達到以下效果:
# 1. 在左上出現容許用家Filter一個或更多的選項的filter，數據來自第一個數據源。最終會有以下四個filter，數據結果要有齊四個filter同時存在的不同filter可能性
# -FY的filter，預設值是"FY 24/25”
# -Inv Yr的filter，不用有預設值
# -Inv Month的filter，不用有預設值
# -Region的filter，不用有預設值，如果選項有"C66 N/A"或空白則把它們在filter中隱去
# -加一個header告訴用家數據來自”Monthly Report” 
# 
# 2. 在左上filter下方顯示使用第一個數據源的filter下的總Inv Amt(HKD)的數字，另外要用兩個平排的pie chart，一個指出不同COST CENTRE各佔了多少百分比，另一個指出各Region各佔了多少百分比，再在下方加一個再用pivot table指出不同COST_CENTRE下各Region的Inv Amt(HKD)是多少金額（用會計單位）
# -如果Inv Yr和Inv Month列中出現"Cancel"或"TBA"或"nan”或空白，就無視該些數據，在數據計算中亦不要計算
# -Pivot Table用html，另要提供download button給用家下載表格
# -pie chart中如果有”C49”就用淺綠色，如果有”C28”就用淺藍色，如果有”C66”就用淺紫色，百分比數字要小數點後2個位，不同部分的百分比數字不要疊在一起，數字亦不能太小，legend放圖中的上方
# -Pie chart不能過大，要縮小到與其他顯示的表格相若 
# 
# 3. 在右邊加入容許用家Filter一個或更多的選項的filter，數據來自第二個數據源。最終會有以下一個sidebar filter
# -FY_INV的filter，預設值是"FY 25/26”
# -加一個header告訴用家這是”Sales Target_FY”
# 
# 在右上filter下方顯示使用第二個數據源的filter下的總Sales Target Amt(HKD)的數字，另外要用兩個平排的pie chart，一個指出不同COST_CENTRE各佔了多少百分比，另一個指出”SOUTH”，”EAST”，”NORTH”和”WEST”各佔了多少百分比，再在下方加一個pivot table指出不同COST_CENTRE下”SOUTH”，”EAST”，”NORTH”和”WEST”各自的Sales Target Amt(HKD)是多少金額（用會計單位）
# -Pivot Table用html，另要提供download button給用家下載表格
# -pie chart中如果有”C49”就用淺綠色，如果有”C28”就用淺藍色，如果有”C66”就用淺紫色，百分比數字要小數點後2個位，不同部分的百分比數字不要疊在一起，數字亦不能太小，legend放圖中的上方
# -pie chart中”SOUTH”用橙色，”EAST”用藍色，”NORTH”用淺紫色，”WEST”用綠色，百分比數字要小數點後2個位，百分比數字不要疊在一起，legend放圖中的上方
# -Pie chart不能過大，要縮小到與其他顯示的表格相若
# 
# 再由上而下出現多個縱向BAR chart，視乎第二個數據源中的"COST_CENTRE”列中有多少個獨立內容，如果只有”C49”和”C28”兩個COST_CENTRE，就出現兩個縱向BAR chart
# Bar chart排位次序任何時候都要”C49”排第一個，”C28”排第二個
# 同時每個bar chart內的數據由第一個數據源和第二個數據源組成：
# -第一個數據源的filter下，用Bar顯示各"Region"的Inv Amt(HKD)的數字，如果"Region"選項有"C66 N/A"或空白則不用顯示它們的Bar
# -Inv Amt(HKD)的數字用簡寫，如果是1000000就用1M顯示，如果是50000000就用50M顯示，如此類推
# -Bar chart中"Region”的BAR顏色，”SOUTH”用橙色，”EAST”用藍色，”NORTH”用淺紫色，”WEST”用綠色
# -Bar chart中BAR顯示由左至右依次是”SOUTH”，”EAST”，”NORTH”和”WEST”
# -由於第二個數據源本來已有”SOUTH”，”EAST”，”NORTH”和”WEST”，就把這些跟第一個數據源的"Region”的BAR配對；即是如果第一個數據源的"Region”有”SOUTH”，就把第二個數據源的”SOUTH”在第二個數據源的filter下的Sales Target Amt(HKD)數字用bar顯示在第一個數據源的”SOUTH”bar的緊貼右邊，每個"Region”都這樣做，如此類推
# -Sales Target Amt(HKD)數字同樣用簡寫，如果是1000000就用1M顯示，如果是50000000就用50M顯示，如此類推
# -全部bar配上黑框，legend放右邊
# 
# 再加入以下元素
# -pivot table上的Region的排序，左至右依次是"SOUTH"和"EAST"和"NORTH"和"WEST"；另外COST_CENTRE中，"C49"放最高，"C28"放第二，"C66"放最下
# -Monthly Report Analysis和Sales Target Analysis兩部分用st.columns(2)分開平排顯示，Regional Comparison部分則不用
# -Regional Comparison內不同"COST_CENTRE”的subheader都分別用不同背景色標記，如果是"C49"的用綠色，"C28"的用藍色，"C66"的用粉色
# -Bar Chart內不同"Region"的"Actual"bar與"Target"bar要平排，不要放在同一條bar，"Target"bar一律用黃色
# -Monthly Report Analysis和Sales Target Analysis的Header各自配上背景色區分，Monthly Report Analysis用淺紫色，Sales Target Analysis用黃色
# -每條bar都要配上黑框
# -Bar Chart內不同"Region"的bar, bar內照樣寫上數字，不過要簡寫，例如之前提到的10M
# -Bar Chart內不同"Region"的bar”除了有Actual"bar與"Target"bar要平排，加多一條名為Difference的bar，就是Actual與Target的相差值，如果該"Region”的Actual值大於Target值，背景色用綠色；如果該"Region”的Actual值少於Target值，背景色用橙色

#-Regional Comparison標題文字要由放在左邊改為置中，並配上黃色背景色，下方每個Region的標題文字也要置中
#-Bar Chart的legend放在圖上方
#-Bar Chart內不同"Region"的bar”如果Actual值大於Target值就等於達到sales target；如果該"Region”的Actual值少於Target值就等於達不到sales target；請在現有基本上下加入最合適的方式讓用家一眼就看出該地區是否達標
#-不同"COST_CENTRE”的subheader標題文字要由放在左邊改為置中
#-Bar Chart內是否達標的標示不要檔住Sales Target Amt(HKD)數字
#不同"COST_CENTRE”的subheader標題文字要置中
#-Bar Chart內的Sales Target Amt(HKD)數字字體大一點，並且粗體
#-Bar Chart內是否達標的標示不要檔住Sales Target Amt(HKD)數字
#-Bar Chart的legend中的difference要一同顯示達標與不達標的兩種顏色
#-由於每個地區的sales target最終只會有達標與不達標的其中一個結果，因此每個"Region"只能有三條bar。不達標的地區就不要顯示達標bar或留位置給達標bar了，同樣地達標的地區就不要顯示未達標bar或留位置給未達標bar了；只是在legend裡顯示有兩種色的difference就好。
#每個pivot table都加入總計，每個pivot table除了行有總計，列都要加入總計，計總一行配以黃色背景色；另外每個bar chart的title都取消不用title
