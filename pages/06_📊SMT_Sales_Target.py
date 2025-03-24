import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from io import StringIO
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
        index=0  # 修改处：移除预设值
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
                font=dict(size=24),
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
                font=dict(size=24),
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
                font=dict(size=24),
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
                font=dict(size=24),
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        st.plotly_chart(fig4, use_container_width=True)

    # Sales Target Pivot Table
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
        barmode='group',
        yaxis_title="Amount (HKD)",
        xaxis_title="Region",
        showlegend=True,
        legend=dict(
            orientation="h",
            font=dict(size=24),
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
        height=800,
        margin=dict(t=100, b=120)
    )
    st.plotly_chart(fig, use_container_width=True)
