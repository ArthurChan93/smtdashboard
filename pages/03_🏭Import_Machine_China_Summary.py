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
######################################################################################################
 
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")
#Title
st.title(':factory:  Mounter Import Data of China_Analysis')
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
#Text
st.write("by Arthur Chan")
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)
 
######################################################################################################
#Create a browser for user to upload
 
#@st.cache_data
#def load_data(file):
#        data = pd.read_excel(file)
#        return data
 
#uploaded_file = st.sidebar.file_uploader(":file_folder: Upload monthly report here")
 
#if uploaded_file is not None:
#df = load_data(uploaded_file)
#st.dataframe(df)
#唔show 17/18, cancel, tba資料
#else:
#os.chdir(r"/Users/arthurchan/Downloads/Sample")
 
df = pd.read_excel(
       io='Monthly_report_for_edit.xlsm',engine= 'openpyxl',sheet_name='raw_sheet', skiprows=0, usecols='A:AO',
       nrows=10000,).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
 

df_import = pd.read_excel(
       io='Machine_Import_data.xlsx',engine= 'openpyxl',sheet_name='Mounter', skiprows=0, usecols='A:G',nrows=10000,)
 
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
# Sidebar Slider
st.sidebar.header(":point_down: Slider for China import data:")
df_import["YEAR"] = df_import["YEAR"].astype(str)
 
#start_yr, end_yr
df_yr= st.sidebar.select_slider('Select a range of year',
            options=df_import["YEAR"].unique(), value=("2022","2023"))
   
#st.sidebar.write('You selected:', start_yr, 'to', end_yr)
 
#New Section      
st.sidebar.divider()
#Sidebar Filter
 
# Create FY Invoice filter
st.sidebar.header(":point_down: Filter for ESE data:")
fy_yr_inv = st.sidebar.multiselect(
        "Select the Financial Year of Invoice",
         options=df["Inv_Yr"].unique(),
         default=[2023,2022],
         )
 
if not fy_yr_inv:
       df2 = df.copy()
else:
       df2 = df[df["Inv_Yr"].isin(fy_yr_inv)]
 
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
# Only select Inv_Yr
elif not region and not cost_centre and not brand:
        filter_df = df[df["Inv_Yr"].isin(fy_yr_inv)]
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
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["Region"].isin(region)]
# Select 'Inv_Yr', 'COST_CENTRE'
elif fy_yr_inv and cost_centre:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Inv_Yr', 'BRAND'
elif fy_yr_inv and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE'
elif region and cost_centre:
      filter_df = df5[df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Region', 'BRAND'
elif region and brand:
      filter_df = df5[df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'COST_CENTRE', 'BRAND'
elif cost_centre and brand:
      filter_df = df5[df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Inv_Yr', 'Region', 'COST_CENTRE'
elif fy_yr_inv and region and cost_centre:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Inv_Yr', 'Region', 'BRAND'
elif fy_yr_inv and region and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'Inv_Yr', 'COST_CENTRE', 'BRAND'
elif fy_yr_inv and cost_centre and brand:
      filter_df = df5[df["Inv_Yr"].isin(fy_yr_inv) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE', 'BRAND'
elif region and cost_centre and brand:
      filter_df = df5[df["Region"].isin(region) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
 
#Show the original data table
#st.dataframe(df_selection)
############################################################################################################################################################################################################
#TAB 1: Overall category
# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
  font-size: 28px;
}
</style>
"""
st.write(font_css, unsafe_allow_html=True)      
 
#st.subheader(":eight_pointed_black_star: Mounter")
 
#Tab 1 MOUNTER
#with tab1:
#Set variable for slider result
selected_df = df_import[df_import["YEAR"].between(df_yr[0], df_yr[1])]
#MOUNTER IMPORT LINE CHART
row1_left_column, row1_right_column = st.columns(2)
with row1_left_column:
       st.subheader(":radio: :red[China Mounter Import Trend]_:orange[QTY]:")
       df_mounter_import = selected_df.groupby(by = ["MONTH","YEAR"], as_index= False)["QTY"].sum()
       fig2 = px.line(df_mounter_import,
                            x= "MONTH",
                            y = "QTY",
                            color='YEAR',
                            symbol="YEAR",
                            markers=True,
                            text="QTY",
                            )
       fig2.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig2.update_layout(yaxis_showticklabels = False), use_container_width=True)
#################################################################################################################################
with row1_right_column:
       st.subheader(":radio: :blue[SMT Sales Trend]_:orange[QTY]:")
#LINE CHART of Overall Invoice Qty
       smtqtyAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('BRAND != "SOLDERSTAR"').query(
                                  'BRAND != "C66 SERVICE"').query('BRAND != "LOCAL SUPPLIER"').round(0).groupby(by = ["Inv_Yr","Inv_Month"],
                            as_index= False)["Item Qty"].sum()
# 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3,4, 5, 6, 7, 8, 9, 10, 11, 12]
       smtqtyAmount_df2 = smtqtyAmount_df2.groupby(["Inv_Yr", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([smtqtyAmount_df2['Inv_Yr'].unique(), sort_Month_order],
                                   names=['Inv_Yr', 'Inv_Month'])).fillna(0).reset_index()
       fig3 = go.Figure()
# 添加每个Inv_Yr的折线
       fy_inv_values = smtqtyAmount_df2['Inv_Yr'].unique()
       for fy_inv in fy_inv_values:
         fy_inv_data = smtqtyAmount_df2[smtqtyAmount_df2['Inv_Yr'] == fy_inv]
         fig3.add_trace(go.Scatter(
              x=fy_inv_data['Inv_Month'],
              y=fy_inv_data['Item Qty'],
              mode='lines+markers+text',
              name=str(fy_inv),
              text=fy_inv_data['Item Qty'],
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
              hovermode='x', showlegend=True)
       fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)
#####################################################
       #加pivot table, column只用YAMAHA, PEMTRON, HELLER台數要match圖
################################################################################################################################# 
row2_left_column, row2_right_column = st.columns(2) 
with row2_left_column:
       st.subheader(":money_with_wings: :red[China Mounter Import Trend]_:orange[AMOUNT]:")
       df_mounter_import = selected_df.groupby(by = ["MONTH","YEAR"], as_index= False)["Import_Amount(RMB)"].sum()
 # 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3,4, 5, 6, 7, 8, 9, 10, 11, 12]
       df_mounter_import = df_mounter_import.groupby(["YEAR", "MONTH"]).sum().reindex(pd.MultiIndex.from_product([df_mounter_import['YEAR'].unique(), sort_Month_order],
                                   names=['YEAR', 'MONTH'])).fillna(0).reset_index()
       fig2 = go.Figure()      
       years = df_mounter_import["YEAR"].unique().astype(str)  # 獲取唯一年份並轉換為字符串列表
       for year in years:
              df_year = df_mounter_import[df_mounter_import["YEAR"] == year]
              fig2.add_trace(go.Scatter(
                     x=df_year["MONTH"],
                     y=df_year["Import_Amount(RMB)"],
                     mode='lines+markers+text',
                     marker=dict(size=9),
                     text=df_year["Import_Amount(RMB)"].apply(lambda x: f'{x:.3f}'),
                     textposition="bottom center",
                     texttemplate='%{text:.3s}',
                     name=year,
                     showlegend=True))
              fig2.update_layout(xaxis=dict(
                     tickmode='linear',
                     tick0=1,
                     dtick=1,
                     title="MONTH"),
                     yaxis=dict(showticklabels=True), 
                     font=dict(family="Arial, Arial", size=12, color="Black"),
                     hovermode='x', showlegend=True,
                     legend=dict(orientation="h"))
              fig2.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       
       st.plotly_chart(fig2.update_layout(yaxis_showticklabels = True), use_container_width=True)      

##########################################################################################################
with row2_right_column:
       st.subheader(":radio: :blue[SMT Sales Trend]_:orange[AMOUNT]:")
#LINE CHART of Overall Invoice Qty
       smtinvoiceAmount_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).groupby(by = ["Inv_Yr","Inv_Month"],
                            as_index= False)["Before tax Inv Amt (HKD)"].sum()
# 确保 "Inv Month" 列中的所有值都出现
       sort_Month_order = [1, 2, 3,4, 5, 6, 7, 8, 9, 10, 11, 12]
       smtinvoiceAmount_df = smtinvoiceAmount_df.groupby(["Inv_Yr", "Inv_Month"]).sum().reindex(pd.MultiIndex.from_product([smtinvoiceAmount_df['Inv_Yr'].unique(), sort_Month_order],
                                   names=['Inv_Yr', 'Inv_Month'])).fillna(0).reset_index()
       fig3 = go.Figure()
# 添加每个Inv_Yr的折线
       fy_inv_values = smtinvoiceAmount_df['Inv_Yr'].unique()
       for fy_inv in fy_inv_values:
         fy_inv_data = smtinvoiceAmount_df[smtinvoiceAmount_df['Inv_Yr'] == fy_inv]
         fig3.add_trace(go.Scatter(
              x=fy_inv_data['Inv_Month'],
              y=fy_inv_data['Before tax Inv Amt (HKD)'],
              mode='lines+markers+text',
              name=str(fy_inv),
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
              hovermode='x', showlegend=True)
       fig3.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = True), use_container_width=True)











with st.expander(":point_right: Click to expand"):
              left_column, right_column = st.columns(2)
              with right_column:
                     pvt_amount = selected_df.round(0).pivot_table(values=["Import_Amount(RMB)"],index=["MONTH"],columns=["YEAR"],
                     aggfunc="sum",fill_value=0).sort_index(axis=0, ascending=False)
                     st.dataframe(pvt_amount.style.format("RMB{:,}"), use_container_width=True)
              with left_column:
                     pvt_qty = selected_df.round(0).pivot_table(values=["QTY"],index=["MONTH"],columns=["YEAR"],
                     aggfunc="sum",fill_value=0).sort_index(axis=0, ascending=False)
                     st.dataframe(pvt_qty.style.format("{:,}"), use_container_width=True)
 
############################################################################################################################################################################################################
st.title(":star: Top YAMAHA Mounter Inv Trend_FQ to FQ:")
 
left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM20R Invoice Details:
with left_column:
              st.subheader(":radio:")
              InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","Inv_Yr"],
                                   as_index= False)["Item Qty"].sum()
              fig = px.line(InvoiceAmount_df2,
              x= "FQ(Invoice)",
              y = "Item Qty",
              color='Inv_Yr',
              symbol="Inv_Yr",
              title="YSM20R INV QTY",
              markers=True,
              text="Item Qty",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["Inv_Yr"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
 
#Line Chart FY to FY YSM20R Invoice Amount Details:
with right_column:
              st.subheader(":money_with_wings:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","Inv_Yr"],
                                   as_index= False)["Before tax Inv Amt (HKD)"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)",
              color='Inv_Yr',
              symbol="Inv_Yr",
              title="YSM20R INV AMOUNT(HKD)",
              markers=True,
              text="Before tax Inv Amt (HKD)",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Invoice Amount Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.round(0).query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Before tax Inv Amt (HKD)",
                            index=["Inv_Yr"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
#New Section      
st.divider()

left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM10 Invoice Qty Details:
with left_column:
              st.subheader(":radio:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","Inv_Yr"],
                                   as_index= False)["Item Qty"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Item Qty",
              color='Inv_Yr',
              symbol="Inv_Yr",
              title="YSM10 INV QTY",
              markers=True,
              text="Item Qty",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Invoice QTY Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["Inv_Yr"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
 
#Line Chart FY to FY YSM10 Invoice Amount Details:
with right_column:
              st.subheader(":money_with_wings:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","Inv_Yr"],
                                   as_index= False)["Before tax Inv Amt (HKD)"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)",
              color='Inv_Yr',
              symbol="Inv_Yr",
              title="YSM10 INV AMOUNT(HKD)",
              markers=True,
              text="Before tax Inv Amt (HKD)",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Invoice Amount Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.round(0).query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Before tax Inv Amt (HKD)",
                            index=["Inv_Yr"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
