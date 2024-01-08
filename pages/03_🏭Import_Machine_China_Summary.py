import streamlit as st
from streamlit_option_menu import option_menu
import plotly.express as px
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
import os
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
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
st.sidebar.header("Slider for China import data:")
df_import["YEAR"] = df_import["YEAR"].astype(str)
 
#start_yr, end_yr
df_yr= st.sidebar.select_slider('Select a range of year',
            options=df_import["YEAR"].unique(), value=("2021","2023"))
   
#st.sidebar.write('You selected:', start_yr, 'to', end_yr)
 
#New Section      
st.sidebar.divider()
#Sidebar Filter
 
# Create FY Invoice filter
st.sidebar.header("Filter for ESE data:")
fy_yr_inv = st.sidebar.multiselect(
        "Select the Financial Year of Invoice",
         options=df["FY_INV"].unique(),
         default=["FY 23/24","FY 22/23"],
         )
 
if not fy_yr_inv:
       df2 = df.copy()
else:
       df2 = df[df["FY_INV"].isin(fy_yr_inv)]
 
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
# Only select FY_INV
elif not region and not cost_centre and not brand:
        filter_df = df[df["FY_INV"].isin(fy_yr_inv)]
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
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["Region"].isin(region)]
# Select 'FY_INV', 'COST_CENTRE'
elif fy_yr_inv and cost_centre:
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_INV', 'BRAND'
elif fy_yr_inv and brand:
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE'
elif region and cost_centre:
      filter_df = df5[df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Region', 'BRAND'
elif region and brand:
      filter_df = df5[df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'COST_CENTRE', 'BRAND'
elif cost_centre and brand:
      filter_df = df5[df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'FY_INV', 'Region', 'COST_CENTRE'
elif fy_yr_inv and region and cost_centre:
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_INV', 'Region', 'BRAND'
elif fy_yr_inv and region and brand:
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'FY_INV', 'COST_CENTRE', 'BRAND'
elif fy_yr_inv and cost_centre and brand:
      filter_df = df5[df["FY_INV"].isin(fy_yr_inv) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
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
  font-size: 24px;
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
left_column, right_column = st.columns(2)
with left_column:
       st.subheader(":radio: :orange[Qty] Trend_YR to YR:")
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
 
with right_column:
              st.subheader(":money_with_wings: :orange[Amount] Trend_YR to YR:")
              df_mounter_import = selected_df.groupby(by = ["MONTH","YEAR"], as_index= False)["Import_Amount(RMB)"].sum()
              fig2 = px.line(df_mounter_import,
                             x= "MONTH",
                             y = "Import_Amount(RMB)",
                             color='YEAR',
                             symbol="YEAR",
                             markers=True,
                             text="Import_Amount(RMB)",
                             )
              fig2.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig2.update_layout(yaxis_showticklabels = False), use_container_width=True)      

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
              InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"],
                                   as_index= False)["Item Qty"].sum()
              fig = px.line(InvoiceAmount_df2,
              x= "FQ(Invoice)",
              y = "Item Qty",
              color='FY_INV',
              symbol="FY_INV",
              title="YSM20R INV QTY",
              markers=True,
              text="Item Qty",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
 
#Line Chart FY to FY YSM20R Invoice Amount Details:
with right_column:
              st.subheader(":money_with_wings:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"],
                                   as_index= False)["Before tax Inv Amt (HKD)"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)",
              color='FY_INV',
              symbol="FY_INV",
              title="YSM20R INV AMOUNT(HKD)",
              markers=True,
              text="Before tax Inv Amt (HKD)",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Invoice Amount Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.round(0).query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Before tax Inv Amt (HKD)",
                            index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
#New Section      
st.divider()

left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM10 Invoice Qty Details:
with left_column:
              st.subheader(":radio:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"],
                                   as_index= False)["Item Qty"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Item Qty",
              color='FY_INV',
              symbol="FY_INV",
              title="YSM10 INV QTY",
              markers=True,
              text="Item Qty",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Invoice QTY Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
 
#Line Chart FY to FY YSM10 Invoice Amount Details:
with right_column:
              st.subheader(":money_with_wings:")
              InvoiceAmount_df3 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"],
                                   as_index= False)["Before tax Inv Amt (HKD)"].sum()
              fig = px.line(InvoiceAmount_df3,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)",
              color='FY_INV',
              symbol="FY_INV",
              title="YSM10 INV AMOUNT(HKD)",
              markers=True,
              text="Before tax Inv Amt (HKD)",
              )
              fig.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Invoice Amount Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt1 = filter_df.round(0).query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Before tax Inv Amt (HKD)",
                            index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt1.style.format("{:,}"), use_container_width=True)
