import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ========== 數據初始化 ==========
os.chdir(r"/Users/arthurchan/Downloads/Sample")
#os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")
df = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='raw_sheet',
    usecols='A:AU'
).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query(
    'FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query(
    'FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query(
    'Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query(
    'Inv_Month != "Cancel"')

df2 = pd.read_excel(
    io='Monthly_report_for_edit.xlsm',
    engine='openpyxl',
    sheet_name='SMT_Internal_Sales_Target',
    usecols='A:G')

# ========== 數據預處理 ==========
df = df.dropna(subset=['Inv_Yr', 'Inv_Month'])
df['Region'] = df['Region'].replace(['', 'C66 N/A'], pd.NA).dropna()

# ========== 頁面配置 ==========
st.set_page_config(layout="wide")

# ========== 全局樣式設定 ==========
region_order = ['SOUTH', 'EAST', 'NORTH', 'WEST']
cost_center_order = ['C49', 'C28', 'C66']
large_font_style = dict(
    xaxis=dict(
        tickfont=dict(size=24),
        titlefont=dict(size=28),
        categoryorder='array',
        categoryarray=region_order,
        type='category'
    ),
    yaxis=dict(
        tickfont=dict(size=24),
        titlefont=dict(size=28)
    ),
    legend=dict(
        font=dict(size=35),
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="center",
        x=0.5
    ),
    margin=dict(t=100, b=120),
    height=800,
    annotationdefaults=dict(font=dict(size=20))
)

# ========== 左側邊欄 ==========
with st.sidebar:
    st.header("Monthly Report")
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
        index=0
    )

# ========== 主內容區 ==========
main_col1, main_col2 = st.columns(2)

# 左側主內容區
with main_col1:
    st.markdown(
        f"<h3 style='background-color: #D8BFD8; padding: 10px; border-radius: 5px;'>"
        f"Monthly Report Data - {selected_fy}</h3>",
        unsafe_allow_html=True
    )
    
    filtered_df = df[
        (df['FY_INV'] == selected_fy) &
        (df['Inv_Yr'].isin(selected_yr if selected_yr else df['Inv_Yr'].unique())) &
        (df['Inv_Month'].isin(selected_month if selected_month else df['Inv_Month'].unique())) &
        (df['Region'].isin(selected_region if selected_region else region_options))
    ]
    
    total_inv = filtered_df['Before tax Inv Amt (HKD)'].sum()
    st.subheader(f"Total Inv Amt (HKD): {total_inv:,.2f}")

    col1, col2 = st.columns(2)
    with col1:
        cost_center_colors = {'C49': '#90EE90', 'C28': '#87CEEB', 'C66': '#FFB6C1'}
        fig1 = px.pie(filtered_df, names='COST_CENTRE', values='Before tax Inv Amt (HKD)',
                     color='COST_CENTRE', color_discrete_map=cost_center_colors)
        fig1.update_traces(textposition='inside', textinfo='percent+label', texttemplate='%{percent:.2%}', textfont_size=40)
        #fig1.update_layout(**large_font_style)
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        fig2 = px.pie(filtered_df, names='Region', values='Before tax Inv Amt (HKD)')
        fig2.update_traces(textposition='inside', textinfo='percent+label', texttemplate='%{percent:.2%}', textfont_size=40)
        #fig2.update_layout(**large_font_style)
        st.plotly_chart(fig2, use_container_width=True)

    pivot_df = filtered_df.pivot_table(
        values='Before tax Inv Amt (HKD)',
        index='COST_CENTRE',
        columns='Region',
        aggfunc='sum',
        fill_value=0,
        margins=True,
        margins_name='Total'
    ).reindex(index=cost_center_order + ['Total'], columns=region_order + ['Total'], fill_value=0)
    
    styled_df = pivot_df.style.format("{:,.2f}").apply(
        lambda x: ['background: #FFFF00' if x.name == 'Total' else '' for _ in x], axis=1)
    st.write(styled_df)
    st.download_button("Download Pivot Table", pivot_df.to_csv().encode('utf-8'), 'monthly_report.csv', 'text/csv')

# 右側主內容區
with main_col2:
    st.markdown(
        f"<h3 style='background-color: #FFFFE0; padding: 10px; border-radius: 5px;'>"
        f"Sales Target - {selected_target_fy}</h3>",
        unsafe_allow_html=True
    )
    
    filtered_df2 = df2[df2['FY_INV'] == selected_target_fy]
    total_target = filtered_df2['Total _Sales_Target(Inv Amt HKD)'].sum()
    st.subheader(f"Total Sales Target (HKD): {total_target:,.2f}")

    col3, col4 = st.columns(2)
    with col3:
        fig3 = px.pie(filtered_df2, names='COST_CENTRE', values='Total _Sales_Target(Inv Amt HKD)',
                     color='COST_CENTRE', color_discrete_map=cost_center_colors)
        fig3.update_traces(textposition='inside', textinfo='percent+label', texttemplate='%{percent:.2%}', textfont_size=40)
    #    fig3.update_layout(**large_font_style)
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        region_colors = {'SOUTH': 'orange', 'EAST': 'blue', 'NORTH': '#D8BFD8', 'WEST': 'green'}
        fig4 = px.pie(filtered_df2.melt(id_vars=['FY_INV', 'COST_CENTRE'], 
                                      value_vars=region_order,
                                      var_name='Region', value_name='Sales Target(HKD)'),
                     names='Region', values='Sales Target(HKD)',
                     color='Region', color_discrete_map=region_colors)
        fig4.update_traces(textposition='inside', textinfo='percent+label', texttemplate='%{percent:.2%}', textfont_size=40)
    #    fig4.update_layout(**large_font_style)
        st.plotly_chart(fig4, use_container_width=True)

    target_pivot = filtered_df2.melt(
        id_vars=['FY_INV', 'COST_CENTRE'],
        value_vars=region_order,
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
    
    styled_target = target_pivot.style.format("{:,.2f}").apply(
        lambda x: ['background: #FFFF00' if x.name == 'Total' else '' for _ in x], axis=1)
    st.write(styled_target)
    st.download_button("Download Target Pivot", target_pivot.to_csv().encode('utf-8'), 'sales_target.csv', 'text/csv')

# ========== Regional Comparison 模組 ==========
st.markdown(
    "<h2 style='background-color: #FFFF00; text-align: center; padding: 10px; border-radius: 5px; font-size: 24px;'>Regional Comparison</h2>",
    unsafe_allow_html=True
)

# ========== 所有成本中心區域比較 ==========
st.markdown(
    "<h4 style='background-color: #D8BFD8; padding: 10px; border-radius: 5px; text-align: center; font-size: 30px;'>All Cost Centres(C49, C28, C66)</h4>", 
    unsafe_allow_html=True
)

all_actual = (
    filtered_df.groupby('Region')['Before tax Inv Amt (HKD)']
    .sum()
    .reindex(region_order, fill_value=0)
    .reset_index()
)

all_target = (
    filtered_df2.melt(
        id_vars=['FY_INV', 'COST_CENTRE'],
        value_vars=region_order,
        var_name='Region',
        value_name='Sales Target(HKD)'
    )
    .groupby('Region')['Sales Target(HKD)']
    .sum()
    .reindex(region_order, fill_value=0)
    .reset_index()
)

merged_all = pd.merge(all_actual, all_target, on='Region', how='outer')
merged_all['Difference(HKD)'] = merged_all['Before tax Inv Amt (HKD)'] - merged_all['Sales Target(HKD)']
merged_all['Achievement%'] = (merged_all['Before tax Inv Amt (HKD)'] / merged_all['Sales Target(HKD)'] * 100).round(2)

fig_all = go.Figure()

for idx, row in merged_all.iterrows():
    actual = row['Before tax Inv Amt (HKD)']
    target = row['Sales Target(HKD)']
    percentage = row['Achievement%']
    
    if actual >= target:
        fig_all.add_annotation(
            x=row['Region'],
            y=max(actual, target) * 1.3,
            text=f"✅ 达标，已实现{percentage}%",
            showarrow=False,
            font=dict(color="green", size=20, family="Arial Bold"),
            bgcolor="white",
            bordercolor="green"
        )
    else:
        fig_all.add_annotation(
            x=row['Region'],
            y=max(actual, target) * 1.3,
            text=f"❌ 未达标，已实现{percentage}%",
            showarrow=False,
            font=dict(color="red", size=20, family="Arial Bold"),
            bgcolor="white",
            bordercolor="red"
        )

fig_all.add_trace(go.Bar(
    x=merged_all['Region'],
    y=merged_all['Before tax Inv Amt (HKD)'],
    name=f'Actual ({selected_fy})',  # 圖例保留FY信息
    marker_color='#1f77b4',
    text=[f"{x/1e6:.1f}M" if x >=1e6 else f"{x/1e3:.0f}K" for x in merged_all['Before tax Inv Amt (HKD)']],  # 移除FY標籤
    textposition='outside',
    textfont=dict(size=20, family='Arial Bold')
))

fig_all.add_trace(go.Bar(
    x=merged_all['Region'],
    y=merged_all['Sales Target(HKD)'],
    name=f'Target ({selected_target_fy})',  # 圖例保留FY信息
    marker_color='#FFD700',
    text=[f"{x/1e6:.1f}M" if x >=1e6 else f"{x/1e3:.0f}K" for x in merged_all['Sales Target(HKD)']],  # 移除FY標籤
    textposition='inside',
    textfont=dict(size=20, family='Arial Bold')
))

fig_all.add_trace(go.Bar(
    x=merged_all['Region'],
    y=merged_all['Difference(HKD)'].abs(),
    name='Difference',
    marker_color=['#2ca02c' if d >=0 else '#ff7f0e' for d in merged_all['Difference(HKD)']],
    text=[f"+{x/1e6:.1f}M" if x >=0 else f"-{abs(x)/1e6:.1f}M" for x in merged_all['Difference(HKD)']],
    textposition='outside',
    textfont=dict(size=20, family='Arial Bold')
))

fig_all.update_layout(
    barmode='group',
    yaxis_title="Amount (HKD)",
    **large_font_style
)
st.plotly_chart(fig_all, use_container_width=True)

# ========== 分成本中心區域比較 ==========
cost_centers = sorted(filtered_df2['COST_CENTRE'].unique(), 
                     key=lambda x: ['C49', 'C28', 'C66'].index(x) if x in ['C49', 'C28', 'C66'] else 3)

for cost_center in cost_centers:
    bg_color = {
        'C49': '#90EE90', 
        'C28': '#87CEEB', 
        'C66': '#FFB6C1'
    }.get(cost_center, 'white')
    
    st.markdown(
        f"<h4 style='background-color: {bg_color}; padding: 10px; border-radius: 5px; text-align: center; font-size: 30px;'>{cost_center}</h4>", 
        unsafe_allow_html=True
    )
    
    actual_data = filtered_df[filtered_df['COST_CENTRE'] == cost_center]
    target_data = filtered_df2[filtered_df2['COST_CENTRE'] == cost_center]
    
    merged = pd.DataFrame({
        'Region': region_order,
        'Actual Inv Amt(HKD)': [actual_data[actual_data['Region'] == r]['Before tax Inv Amt (HKD)'].sum() for r in region_order],
        'Sales Target(HKD)': [target_data[r].sum() for r in region_order]
    })
    merged['Difference(HKD)'] = merged['Actual Inv Amt(HKD)'] - merged['Sales Target(HKD)']
    merged['Achievement%'] = (merged['Actual Inv Amt(HKD)'] / merged['Sales Target(HKD)'] * 100).round(2)

    fig = go.Figure()
    
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
                font=dict(color="green", size=20, family="Arial Bold"),
                bgcolor="white",
                bordercolor="green"
            )
        else:
            fig.add_annotation(
                x=region,
                y=max(actual, target) * 1.3,
                text=f"❌ 未达标，已实现{percentage}%",
                showarrow=False,
                font=dict(color="red", size=20, family="Arial Bold"),
                bgcolor="white",
                bordercolor="red"
            )
    
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Actual Inv Amt(HKD)'],
        name=f'Actual ({selected_fy})',  # 圖例保留FY信息
        marker_color='#1f77b4',
        text=[f"{x/1e6:.1f}M" if x >=1e6 else f"{x/1e3:.0f}K" for x in merged['Actual Inv Amt(HKD)']],  # 移除FY標籤
        textposition='outside',
        textfont=dict(size=20, family='Arial Bold')
    ))
    
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Sales Target(HKD)'],
        name=f'Target ({selected_target_fy})',  # 圖例保留FY信息
        marker_color='#FFD700',
        text=[f"{x/1e6:.1f}M" if x >=1e6 else f"{x/1e3:.0f}K" for x in merged['Sales Target(HKD)']],  # 移除FY標籤
        textposition='inside',
        textfont=dict(size=20, family='Arial Bold')
    ))
    
    fig.add_trace(go.Bar(
        x=merged['Region'],
        y=merged['Difference(HKD)'].abs(),
        name='Difference',
        marker_color=['#2ca02c' if d >=0 else '#ff7f0e' for d in merged['Difference(HKD)']],
        text=[f"+{x/1e6:.1f}M" if x >=0 else f"-{abs(x)/1e6:.1f}M" for x in merged['Difference(HKD)']],
        textposition='outside',
        textfont=dict(size=20, family='Arial Bold')
    ))
    
    fig.update_layout(
        barmode='group',
        yaxis_title="Amount (HKD)",
        **large_font_style
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
    
#-在Regional Comparison部分，每個COST_CENTRE的header字體加大一半
#-在Regional Comparison部分，除了現在的分開COST_CENTRE去顯示，另外加入一個部分是所有COST_CENTRE與Sales Target對比，此新部分放在Regional Comparison部分的最上面
#-在Regional Comparison部分，"All Cost Centres vs Sales Target"的header改為"All Cost Centres"
#-在Regional Comparison部分，"All Cost Centres"部分跟其他單一COST_CENTRE部分做法一樣，分"SOUTH"和"EAST"和"NORTH"和"WEST"去展示
#-在Regional Comparison部分，所有COST_CENTRE與Sales Target對比的部分，Legend和提示是否達標兩部分的構成和方法和格式要跟其他單一COST_CENTRE部分做法一樣
#-在Regional Comparison部分，"All Cost Centres"部分，"SOUTH"和"EAST"和"NORTH"和"WEST"各自的amount(HKD)數值展示該地區的所有COST_CENTRE加起來的invoice數值就可
#-在Regional Comparison部分，"All Cost Centres"部分，"SOUTH"和"EAST"和"NORTH"和"WEST"各自的amount(HKD)invoice數值我不想在同一bar內分開展示幾個cost centre的數值，只展示總數就可
#-在Regional Comparison部分，"All Cost Centres"部分，排列次序要按"SOUTH"和"EAST"和"NORTH"和"WEST"排，Sales Targe和Difference都要用棒形圖，格式參考其他單一COST_CENTRE比較圖就可
#- Monthly Report Data部分要有標示去顯示現在此部分的數據是屬於哪一個FY，例如在Monthly Report的sidebar filter中已選了 FY 22/23，就標示"FY 22/23"，所以標示就是視乎Monthly Report的sidebar ilter中filter的FY是什麼，就變成什麼
#- Sales Target部分要有標示去顯示現在此部分的數據是屬於哪一個FY_INV (Target)，例如在Sales Target_FY的sidebar filter中已選了 FY 22/23，就標示"FY 22/23"，所以標示就是視乎Sales Target_FY的sidefilter中filter的是什麼，就變成什麼
#- 在Regional Comparison部分，所有柱圖的scale和column字體都加大一半
#- Regional Comparison所有柱圖內的Actual inv amt的bar都要有標示該bar的數據是屬於哪一個FY。Actual inv amt的bar要視乎Monthly Report的sidebar ilter中filter的FY是什麼，如果在Monthly Report的sidebar filter中的FY已選了 FY 22/23，就每條Actual inv amt的bar標示"FY 22/23"
#- Regional Comparison所有柱圖內的sales target的bar都要有標示該bar的數據是屬於哪一個FY_INV (Target)。sales target的bar要視乎Sales Target_FY的sidefilter中filter的是什麼，如果在FY_INV (Target)的sidebar filter中的FY已選了 FY 22/23，就每條sales target的bar標示"FY 22/23"
