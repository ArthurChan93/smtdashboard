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

######################################################################################################
 
# emojis https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
#Webpage config& tab name& Icon
st.set_page_config(page_title="Sales Dashboard",page_icon=":rainbow:",layout="wide")
#Title
st.title(':pencil: SMT_Contract Dashboard')
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
#         df = load_data(uploaded_file)
#         st.dataframe(df)
 
 #唔show 17/18, cancel, tba資料
#else:
#        os.chdir(r"/Users/arthurchan/Downloads/Sample")
df = pd.read_excel(
                io='Sample_excel.xlsx',engine= 'openpyxl',sheet_name='sheet 1', skiprows=0, usecols='A:AO',nrows=10000,).query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').query('Region != "C66 N/A"').query('FY_Contract != "FY 16/17"').query('FY_Contract != "FY 17/18"')
 
######################################################################################################
# https://icons.getbootstrap.com/
#Top menu bar
 
#with st.sidebar:
#      selected = option_menu(
#              menu_title=None,
#              options=["Contract Summary","Contract Summary","Mounter & Non-Mounter"],
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
 
#if selected == "Contract Summary":
 
######################################################################################################
#Sidebar Filter
 
# Create FY Contract filter
st.sidebar.header("Filter:")
fy_yr_contract = st.sidebar.multiselect(
        "Select the Financial Year of Contract",
         options=df["FY_Contract"].unique(),
         default=["FY 23/24"],
         )

if not fy_yr_contract:
       df2 = df.copy()
else:
       df2 = df[df["FY_Contract"].isin(fy_yr_contract)]
 
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
if not fy_yr_contract and not region and not cost_centre and not brand:
       filter_df = df
# Only select FY_Contract
elif not region and not cost_centre and not brand:
        filter_df = df[df["FY_Contract"].isin(fy_yr_contract)]
# Only select Region
elif not fy_yr_contract and not cost_centre and not brand:
       filter_df = df[df["Region"].isin(region)]
# Only select cost centre
elif not fy_yr_contract and not region and not brand:
       filter_df = df[df["COST_CENTRE"].isin(cost_centre)]
# Only select brand
elif not fy_yr_contract and not region and not cost_centre:
       filter_df = df[df["BRAND"].isin(brand)]
# Select FY INV & Region
elif fy_yr_contract and region:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df5["Region"].isin(region)]
# Select 'FY_Contract', 'COST_CENTRE'
elif fy_yr_contract and cost_centre:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_Contract', 'BRAND'
elif fy_yr_contract and brand:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE'
elif region and cost_centre:
      filter_df = df5[df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'Region', 'BRAND'
elif region and brand:
      filter_df = df5[df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'COST_CENTRE', 'BRAND'
elif cost_centre and brand:
      filter_df = df5[df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'FY_Contract', 'Region', 'COST_CENTRE'
elif fy_yr_contract and region and cost_centre:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df["Region"].isin(region) & df5["COST_CENTRE"].isin(cost_centre)]
# Select 'FY_Contract', 'Region', 'BRAND'
elif fy_yr_contract and region and brand:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df["Region"].isin(region) & df5["BRAND"].isin(brand)]
# Select 'FY_Contract', 'COST_CENTRE', 'BRAND'
elif fy_yr_contract and cost_centre and brand:
      filter_df = df5[df["FY_Contract"].isin(fy_yr_contract) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
# Select 'Region', 'COST_CENTRE', 'BRAND'
elif region and cost_centre and brand:
      filter_df = df5[df["Region"].isin(region) & df["COST_CENTRE"].isin(cost_centre) & df5["BRAND"].isin(brand)]
 
#Show the original data table
#st.dataframe(df_selection)
 
############################################################################################################################################################################################################ 
#Overall Summary

left_column, middle_column, right_column = st.columns(3)
total_contract_amount = int(filter_df["Before tax Inv Amt (HKD)"].sum())
with left_column:
              st.metric(label=":dollar: Total :orange[Contract] Amount before tax",value=(f"HKD{total_contract_amount:,}"),)
  
total_NewSign_item = int(filter_df["Item Qty"].sum())
with middle_column:
              st.metric(label=":factory: :orange[Signed] Main Unit Qty",value=(f"{total_NewSign_item:,}"))

total_NewCustomer_qty = int(filter_df["New Customer Qty(Contract month)"].sum())
with right_column:
              st.metric(label=":face_with_cowboy_hat: New Customer Qty",value=(f"{total_NewCustomer_qty:,}"))

style_metric_cards(background_color="Summer",border_left_color="GREEN",border_size_px=5,box_shadow=True,border_radius_px=1)

############################################################################################################################################################################################################  
#Pivot table, 差sub-total, GP%

 
filter_df["Contract_Month"] = filter_df["Contract_Month"].astype(str)
filter_df["Contract_Yr"] = filter_df["Contract_Yr"].astype(str)
filter_df["Contract No."] = filter_df["Contract No."].astype(str)
# add subtotal
# https://morioh.com/a/17278219952b/tutorial-on-data-analysis-with-python-and-pivot-tables-with-pandas

############################################################################################################################################################################################################        
#Create tabs after overall summary

# Make the tab font bigger
font_css = """
<style>
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p {
  font-size: 24px;
}
</style>
"""

st.write(font_css, unsafe_allow_html=True)
tab1, tab2, tab3 ,tab4= st.tabs([":earth_asia: Overall",":dart: Region",":package: Brand",":handshake: Customer"])

#TAB 1: Overall category

with tab1:
       left_column, right_column = st.columns(2)
       with left_column:
               st.subheader(":green_book: New Contract Details_:orange[Contract Month]:")
               pvt = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(values=["Before tax Inv Amt (HKD)"],index=["FY_Contract","FQ(Contract)","Contract_Yr","Contract_Month"],
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
               
              #add subtotals row to pivot table
              #(pd.concat([pvt, 
              #            pvt.query('Contract_Month != "Total"').sum(level=0).assign(FY_Contract='subtotal')
              #           .set_index("FY_Contract", append=True)])
              #   .sort_index())
               
               st.dataframe(pvt.style.highlight_max(color = 'yellow', axis = 0).format("HKD{:,}"), use_container_width=True)
       
       properties = {"border": "1px solid black", "width": "65px", "text-align": "center"}
           
       with right_column:
               st.subheader(":globe_with_meridians: New Contract Details_:orange[All Cost Centre]:")
               pvt2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(0).pivot_table(index=["FY_Contract","FQ(Contract)","Contract_Yr","Contract_Month"],columns=["COST_CENTRE"],values="Before tax Inv Amt (HKD)",
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
               st.dataframe(pvt2.style.highlight_max(color = 'yellow', axis = 0).format("HKD{:,}").set_properties(**properties).highlight_between(left=-1, right=1, props='font-weight:bold;color:red'), use_container_width=True)

               #FY to FY Quarter Contract Details:
       st.subheader(":point_down: Contract Amount :orange[Subtotal]_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt6 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_Contract"],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
              st.dataframe(pvt6.style.highlight_max(color = 'yellow', axis = 0).format("HKD{:,}"), use_container_width=True) 
       
       st.divider()

#LINE CHART of Overall Contract Amount
 # 加底部sum table
       st.subheader(":chart_with_upwards_trend: Contract Amount Trend_FQ to FQ(Available to show :orange[Multiple FY]):")
       ContractAmount_df2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FQ(Contract)","FY_Contract"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig_overall_Contract = px.line(ContractAmount_df2,
              x= "FQ(Contract)",
              y = "Before tax Inv Amt (HKD)", 
              color='FY_Contract',
              symbol="FY_Contract",
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig_overall_Contract.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig_overall_Contract.update_layout(yaxis_showticklabels = False), use_container_width=True)

#New Section       
       st.divider()
############################################################################################################################################################################################################  
#TAB 2: Region Category
with tab2:
       st.subheader(":sunrise: FY Contract Details_:orange[Contract Month]:")
       pvt = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_Contract","FQ(Contract)","Contract_Yr","Contract_Month"],columns=["Region"],
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt.style.highlight_max(color = 'yellow', axis = 0).highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)

       left_column, middle_column, right_column = st.columns(3)
       
#Regional total inv amount% PIE CHART
       with left_column:
          st.subheader(":round_pushpin: Regional Contract Amount_:orange[Percentage Distribution]:")
          fig1 = px.pie(filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"'), values= "Before tax Inv Amt (HKD)", names = "Region", hole = 0.5)
          fig1.update_traces(text = filter_df["Region"], textposition= "inside")
          st.plotly_chart(fig1,use_container_width=True)
       
#Regional total inv amount BAR CHART
       with middle_column:
              category_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by = ["Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)",ascending=False)
              st.subheader(":bar_chart: Total Regional Contract Amount:")
              fig = px.bar(category_df, x = "Region", y = "Before tax Inv Amt (HKD)",text_auto='.2s')
              fig.update_traces(text = filter_df["Region"], textposition= "inside")
              fig.update_traces(marker_color = 'lightgreen', marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(fig, use_container_width=True, height = 150, weight = 300)
              
#BAR CHART of Regional Contract AMOUNT FY to FY
       with right_column:              
              st.subheader(":bar_chart: Regional Contract Amount_FY to FY(Available to show :orange[Multiple FY]):")
              category2_df = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by = ["FY_Contract","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)",ascending=False)
              df_contract = px.bar(category2_df, x="FY_Contract", y="Before tax Inv Amt (HKD)", color="Region",text_auto='.2s')
              df_contract.update_traces(marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(df_contract, use_container_width=True)
              #with st.expander("Regional_Contract_Amount(HKD)_View_Actual_Figures"):
                #st.write(category_df.style.background_gradient(cmap="Greens"))
                #csv = category_df.to_csv(index = False).encode('utf-8')
                #st.download_button("Download Data", data = csv, file_name= "Regional_Sales.csv", mime = "text/csv",
                #            help = 'Click here to download the data as a Excel file')
              
       #FY Quarter Regional Contract Details:
       st.subheader(":point_down: Regional Contract Amount Subtotal_FY to FY:")
       with st.expander("Click to expand"):
              pvt7 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["Region"],columns=["FY_Contract"],
              aggfunc="sum",fill_value=0).sort_values(by="Region",ascending=False)
              st.dataframe(pvt7.style.format("HKD{:,}"), use_container_width=True)
       
       st.divider()

# LINE CHART of Regional Comparision
       st.subheader(":chart_with_upwards_trend: All Region Contract Amount Trend_FY to FY(Available to show :orange[Multiple Regions]):")
       ContractAmount_df2 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).groupby(by = ["FQ(Contract)","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig2 = px.line(ContractAmount_df2,
               x = "FQ(Contract)",
               y = "Before tax Inv Amt (HKD)",
               color='Region',
               markers=True,
               text="Before tax Inv Amt (HKD)",
               )
       fig2.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig2, use_container_width=True)

#All Region Contract Details FQ_FQ:
       st.subheader(":point_down: Regional Contract Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt7 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["Region"],columns=["FQ(Contract)"],
              aggfunc="sum",fill_value=0).sort_values(by="Region",ascending=False)
              st.dataframe(pvt7.style.highlight_max(color = 'yellow', axis = 0).highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of SOUTH CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[SOUTH CHINA] Contract Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "SOUTH"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","FQ(Contract)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig4 = px.line(df_Single_region,
              x= "FQ(Contract)",
              y = "Before tax Inv Amt (HKD)", 
              color='FY_Contract',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig4.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}',marker_color = 'orange',)
       st.plotly_chart(fig4.update_layout(yaxis_showticklabels = False), use_container_width=True)
#SOUTH Region Contract Details FQ_FQ:
       st.subheader(":point_down: SOUTH CHINA Contract Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt8 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "SOUTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_Contract',ascending=False)
              st.dataframe(pvt8.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
#LINE CHART of EAST CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[EAST CHINA] Contract Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "EAST"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","FQ(Contract)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig5 = px.line(df_Single_region,
              x= "FQ(Contract)",
              y = "Before tax Inv Amt (HKD)", 
              color='FY_Contract',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig5.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}',marker_color = 'orange',)
       st.plotly_chart(fig5.update_layout(yaxis_showticklabels = False), use_container_width=True)
#EAST Region Contract Details FQ_FQ:
       st.subheader(":point_down: EAST CHINA Contract Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt9 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "EAST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_Contract',ascending=False)
              st.dataframe(pvt9.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of WEST CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[WEST CHINA] Contract Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "WEST"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","FQ(Contract)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig6 = px.line(df_Single_region,
              x= "FQ(Contract)",
              y = "Before tax Inv Amt (HKD)", 
              color='FY_Contract',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig6.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}',marker_color = 'orange',)
       st.plotly_chart(fig6.update_layout(yaxis_showticklabels = False), use_container_width=True)
#WEST Region Contract Details FQ_FQ:
       st.subheader(":point_down: WEST CHINA Contract Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "WEST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_Contract',ascending=False)
              st.dataframe(pvt10.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of NORTH CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[NORTH CHINA] Contract Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "NORTH"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FY_Contract","FQ(Contract)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig7 = px.line(df_Single_region,
              x= "FQ(Contract)",
              y = "Before tax Inv Amt (HKD)", 
              color='FY_Contract',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig7.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}',marker_color = 'orange',)
       st.plotly_chart(fig7.update_layout(yaxis_showticklabels = False), use_container_width=True)
#NORTH Region Contract Details FQ_FQ:
       st.subheader(":point_down: NORTH CHINA Contract Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Region == "NORTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_Contract'],columns=["FQ(Contract)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_Contract',ascending=False)
              st.dataframe(pvt10.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
############################################################################################################################################################################################################    
#TAB 3: Brand category 
with tab3:
       #Brand Contract Qty by Contract Month:
       st.subheader(":point_down: Brand Item Signed Qty_:orange[Contract Month]:")
       #with st.expander("Click to expand"):
       pvt6 = filter_df.query('FY_Contract != "TBA"').query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').pivot_table(values="Item Qty",index=["FY_Contract","FQ(Contract)","Contract_Yr","Contract_Month","BRAND","Ordered_Items"],
       aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt6.style.format("{:,}").highlight_max(color = 'yellow', axis = 0), use_container_width=True) 

       left_column, middle_column , right_column= st.columns(3)
       with left_column:
              st.subheader(":bar_chart: Overall Contract_Brand Qty:")
              sales_by_brand_line = (
              filter_df.query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').groupby(by=["BRAND"])[["Item Qty"]].sum().sort_values(by="Item Qty")
       )
              fig_product_sales = px.bar(
              sales_by_brand_line,
              x=["Item Qty"],
              y=sales_by_brand_line.index,
              text_auto='.2s',
              orientation="h",
              color_discrete_sequence=["#008388"]*len(sales_by_brand_line),
              template="plotly_white"
              )
              fig_product_sales.update_traces(text = filter_df["BRAND"])
              fig_product_sales.update_traces(marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(fig_product_sales, use_container_width=True)
              
       with middle_column:
                     st.subheader(":trophy: :orange[Brand List]_Contract Amount& Qty:")
                     pvt4 = filter_df.pivot_table(index=["BRAND"],values=["Before tax Inv Amt (HKD)","Item Qty"],
                     aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
                     st.dataframe(pvt4, use_container_width=True)
              
       with right_column:
                     st.subheader(":sports_medal: :orange[Item List]_Contract Amount& Qty:")
                     pvt5 = filter_df.round(0).pivot_table(index=["Ordered_Items"],values=["Before tax Inv Amt (HKD)","Item Qty"],
                     aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
                     st.dataframe(pvt5, use_container_width=True)

       st.subheader(":clipboard: Brand Contract Qty_FY to FY(Available to show :orange[Multiple FY]):")
       pvt8 = filter_df.pivot_table(index=["BRAND"],values="Item Qty",columns=["FY_Contract"],
       aggfunc="sum",fill_value=0,).sort_index(axis=0, ascending=False)#.sort_values(by="Item Qty",ascending=False)
       st.dataframe(pvt8, use_container_width=True)

#Top Product line chart contract qty trend
       left_column, middle_column, right_column = st.columns(3)
       with middle_column:
              st.subheader(":chart_with_upwards_trend: :orange[YAMAHA Contract Qty] Trend_FQ to FQ:")

       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM20R Contract Details:
       with left_column:
              df_YSM20 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FQ(Contract)","FY_Contract"], 
                                   as_index= False)["Item Qty"].sum()
              fig8 = px.line(df_YSM20,
              x= "FQ(Contract)",
              y = "Item Qty", 
              color='FY_Contract',
              symbol="FY_Contract",
              title="YSM20R",
              markers=True, 
              
              text="Item Qty", 
              )
              fig8.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig8.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Contract Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt13 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt13.style.format("{:,}"), use_container_width=True)

#Line Chart FY to FY YSM10 Contract Details:
       with right_column:
              df_YSM10 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FQ(Contract)","FY_Contract"], 
                                   as_index= False)["Item Qty"].sum()
              fig9 = px.line(df_YSM10,
              x= "FQ(Contract)",
              y = "Item Qty", 
              color='FY_Contract',
              symbol="FY_Contract",
              title="YSM10",
              markers=True, 
              text="Item Qty", 
              )
              fig9.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig9.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Contract Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt14 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt14.style.format("{:,}"), use_container_width=True)


#Second row for Top product trend
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM40R Contract Details:
       with left_column:
              df_YSM40 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSM40R"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FQ(Contract)","FY_Contract"], 
                                   as_index= False)["Item Qty"].sum()
              fig10 = px.line(df_YSM40,
              x= "FQ(Contract)",
              y = "Item Qty", 
              color='FY_Contract',
              symbol="FY_Contract",
              title="YSM40R",
              markers=True, 
              text="Item Qty", 
              )
              fig10.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig10.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM40R Contract Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt14 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YSM40R"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt14.style.format("{:,}"), use_container_width=True)


#Line Chart FY to FY YSi-V Contract Details:
       with right_column:
              df_YSIV = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('Ordered_Items == "YSi-V"').query('FY_Contract != "TBA"').round(0).groupby(by = ["FQ(Contract)","FY_Contract"], 
                                   as_index= False)["Item Qty"].sum()
              fig11 = px.line(df_YSIV,
              x= "FQ(Contract)",
              y = "Item Qty", 
              color='FY_Contract',
              symbol="FY_Contract",
              title="YSi-V",
              markers=True, 
              text="Item Qty", 
              )
              fig11.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig11.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM40R Contract Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt15 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').query('FY_Contract != "TBA"').query('Ordered_Items == "YSi-V"').pivot_table(values="Item Qty",index=["FY_Contract"],columns=["FQ(Contract)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt15.style.format("{:,}"), use_container_width=True)         
############################################################################################################################################################################################################  
#TAB 4: Customer category
with tab4:
       #Top Down Customer details Table 
       st.subheader(":point_down: Customer Contract Details_:orange[Contract Month]")
       #with st.expander("Click to expand"):
       pvt3 = filter_df.query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').pivot_table(index=["FY_Contract","FQ(Contract)","Contract_Yr",
       "Contract_Month","COST_CENTRE","Region","Contract No.","Customer Name","Ordered_Items"],
       values=["Before tax Inv Amt (HKD)","Item Qty"],
       aggfunc="sum",
       fill_value=0,
       margins=True,
       margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt3.round(0).style.highlight_max(color = 'yellow', axis = 0).format("{:,}"), use_container_width=True)

       left_column, middle_column , right_column= st.columns(3)
       #BAR CHART Customer List 
       with left_column:
              st.subheader(":bar_chart: :orange[Top 10 Customer]_Contract Amount:")
              customer_line = (
              filter_df.query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').groupby(by=["Customer Name"])[["Before tax Inv Amt (HKD)"]]
              .sum().sort_values(by="Before tax Inv Amt (HKD)", ascending = False).head(10)
       )
              fig_customer = px.bar(
              customer_line,
              x=["Before tax Inv Amt (HKD)"],
              y=customer_line.index,
              text_auto='.2s',
              orientation="h",
              color_discrete_sequence=["orange"]*len(customer_line),
              template="plotly_white",
              height=400,
              )
              fig_customer.update_traces(text = filter_df["Customer Name"])
              fig_customer.update_traces(marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(fig_customer, use_container_width=True, ascending = False)
       
       with middle_column:
              st.subheader(":trophy: :orange[Top Customer List]_Contract Amount& Qty:")
              pvt11 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region"],
                     values=["Item Qty","Before tax Inv Amt (HKD)"],
              aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
              st.dataframe(pvt11.style.format("{:,}"), use_container_width=True)

       with right_column:
              st.subheader(":medal: :orange[Top Customer Purchase List]_Contract Amount& Qty:")
              pvt12 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region","Ordered_Items"],
                     values=["Item Qty","Before tax Inv Amt (HKD)"],
              aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
              st.dataframe(pvt12.style.format("{:,}"), use_container_width=True)

# LINE CHART of New Customer Qty Regional Comparison
#       st.subheader(":chart_with_upwards_trend: New Customer Qty Trend_FQ to FQ(Region):")
#       NewCustomerAmount_df = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').round(0).groupby(by = ["FQ(Contract)","Region"], 
#               as_index= False)["New Customer Qty"].sum()
#       fig4 = px.line(NewCustomerAmount_df,
#               x = "FQ(Contract)",
#               y = "New Customer Qty",
#               color='Region',
#               markers=True,
#               text="New Customer Qty",
#               )
#       st.plotly_chart(fig4, use_container_width=True)
       
# BAR CHART of New Customer Qty Regional Comparison
       st.subheader(":chart_with_upwards_trend: :orange[New Customer Qty]_FY to FY(Available to show :orange[Multiple FY]):")  
       df_NewCustomerAmount_bar = filter_df.query('Contract_Yr != "TBA"').query('Contract_Month != "TBA"').query('Contract_Month != "Cancel"').groupby(by = ["FY_Contract","Region"], 
               as_index= False)["New Customer Qty(Contract month)"].sum()    
       fig5= px.bar(df_NewCustomerAmount_bar, x="FY_Contract", y="New Customer Qty(Contract month)", color="Region", barmode="group",color_continuous_scale = 'viridis',text_auto = True)
       fig5.update_traces(textposition='outside')

       colors = ["lightblue","green","yellow","purple"]
       fig5.update_traces(marker_line_color = 'black', showlegend=True, 
                  marker_line_width = 2, opacity = 1)
       st.plotly_chart(fig5, use_container_width=True)
#Pivot table FY to FY New Customer QTY:
       with st.expander(":point_right: Click to expand"):
              pvt13 = filter_df.query('FY_Contract != "TBA"').query('FY_Contract != "Cancel"').pivot_table(values="New Customer Qty(Contract month)",index=["FY_Contract"],columns=["Region"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
              st.dataframe(pvt13.style.format("{:,}"), use_container_width=True)  









############################################################################################################################################################################################################  



#fig3, ax = plt.subplots(figsize = (9,4))
#ax.plot(filter_df.Contract_Month, filter_df.Contract_Month, linewidth = 2,color='red',maker='o',markersize = 6, label ='South')
#ax.plot(filter_df.Contract_Month, filter_df.East, linewidth = 2,color='blue',maker='o',markersize = 6, label ='East')
#ax.plot(filter_df.Contract_Month, filter_df.West, linewidth = 2,color='green',maker='o',markersize = 6, label ='West')
#ax.plot(filter_df.Contract_Month, filter_df.North, linewidth = 2,color='yellow',maker='o',markersize = 6, label ='North')
#ax.set_title("Monthly Contract Amount HKD by Region")
#ax.set_xlabel('Contract Month')
#ax.set_ylabel('Contract Amount HKD')
#plt.xticks(rotation = 45)
#ax.legend(loc = 'lower left', fontsize = 10)
#st.pyplot(fig3)
 

#success info
#st.success("Executed successfully")
#st.info("This is an information")
#st.warning("This is a warning")
#st.error("An error occured")

 
######################################################################################################

 
