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
st.title(':world_map: SMT_Invoice Dashboard')
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
                io='Monthly_report_for_edit.xlsx',engine= 'openpyxl',sheet_name='list', skiprows=0, usecols='A:AO',nrows=10000,).query('Region != "C66 N/A"').query('FY_Contract != "Cancel"').query('FY_INV != "TBA"').query('FY_INV != "FY 17/18"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Yr != "Cancel"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"')
 

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
#Sidebar Filter
 
# Create FY Invoice filter
st.sidebar.header("Filter:")
fy_yr_inv = st.sidebar.multiselect(
        "Select the Financial Year of Invoice",
         options=df["FY_INV"].unique(),
         default=["FY 23/24"],
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
#Overall Summary

left_column, middle_column, right_column = st.columns(3)
total_invoice_amount = int(filter_df["Before tax Inv Amt (HKD)"].sum())
with left_column:
              st.metric(label=":dollar: Total :orange[Invoice] Amount before tax",value=(f"HKD{total_invoice_amount:,}"),)
  
total_gp = int(filter_df["G.P.  (HKD)"].sum())
with middle_column:
              st.metric(label=":moneybag: Total :orange[G.P] Amount",value=(f"HKD{total_gp:,}"))

total_unit_qty = int(filter_df["Item Qty"].sum())
with right_column:
              st.metric(label=":factory: Total :orange[Invoiced] Machine",value=(f"Qty: {total_unit_qty:,}"))

style_metric_cards(background_color="Summer",border_left_color="BLUE",border_size_px=5,box_shadow=True,border_radius_px=1)

############################################################################################################################################################################################################  
#Pivot table, 差sub-total, GP%

 
filter_df["Inv_Month"] = filter_df["Inv_Month"].astype(str)
filter_df["Inv_Yr"] = filter_df["Inv_Yr"].astype(str)
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
               st.subheader(":blue_book: Invoice Details_:orange[Inv Month]:")
               pvt = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(values=["Before tax Inv Amt (HKD)","G.P.  (HKD)"],index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month"],
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
               
              #add subtotals row to pivot table
              #(pd.concat([pvt, 
              #            pvt.query('Inv_Month != "Total"').sum(level=0).assign(FY_INV='subtotal')
              #           .set_index("FY_INV", append=True)])
              #   .sort_index())
               col_loc_1 = filter_df.columns.get_loc('FQ(Invoice)')
               
               highlight_col=lambda x: ['background: red' if x.name in ['FQ(Invoice)']
                                   else '' for i in x]
##               def redden(x):
#                      if x == "Q4":
#                             f'&lt;span class="significant" style="color: red;"&gt;{x}&lt;/span&gt;'
#                             return str(x)
#                      else:
#                             return None
               st.dataframe(pvt.style
                                     .highlight_max(color = 'yellow', axis = 0)
                                 
                                     .format("HKD{:,}"), use_container_width=True)

#.background_gradient(subset=["FQ(Invoice)"],cmap='Blues')      

              #headers={"selector":"th",
              #         "props": [("background-color", "dodgerblue"), ("color", "black")]}
                                      
               
       with right_column:
               st.subheader(""":globe_with_meridians: Invoice Details_:orange[All Cost Centre]:""")
               pvt2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month"],columns=["COST_CENTRE"],values="Before tax Inv Amt (HKD)",
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
               st.dataframe(pvt2.style                                                      
                            .highlight_max(color = 'yellow', axis = 0).format("HKD{:,}"), use_container_width=True)






       #FY to FY Quarter Invoice Details:
       st.subheader(":point_down: Invoice Amount :orange[Subtotal]_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt6 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
              st.dataframe(pvt6.style.highlight_max(color = 'yellow', axis = 0)
                                     .format("HKD{:,}"), use_container_width=True) 
       
       st.divider()

#LINE CHART of Overall Invoice Amount
 # 加底部sum table，accounting format
       st.subheader(":chart_with_upwards_trend: Inv Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY)]:")
       InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig3 = px.line(InvoiceAmount_df2,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)", 
              #title= "Inv Amount FY to FY(Multiple FY selection available)",
              color='FY_INV',
              symbol="FY_INV",
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig3.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig3.update_layout(yaxis_showticklabels = False), use_container_width=True)

#New Section       
       st.divider()
############################################################################################################################################################################################################  
#TAB 2: Region Category
with tab2:
       st.subheader(":sunrise: FY Invoice Details_:orange[Inv Month]:")
       pvt = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month"],columns=["Region"],
                            aggfunc="sum",
                            fill_value=0,
                            margins=True,
                            margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt.style.highlight_max(color = 'yellow', axis = 0).highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)

       left_column, middle_column, right_column = st.columns(3)
       
#Regional total inv amount% PIE CHART
       with left_column:
          st.subheader(":round_pushpin: Regional Invoice_:orange[Percentage Distribution]:")
          fig1 = px.pie(filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"'), values= "Before tax Inv Amt (HKD)", names = "Region", hole = 0.5)
          fig1.update_traces(text = filter_df["Region"], textposition= "inside")
          st.plotly_chart(fig1,use_container_width=True)
       
#Regional total inv amount BAR CHART
       with middle_column:
              category_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by = ["Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)",ascending=False)
              st.subheader(":bar_chart: Total Regional Invoice Amount:")
              fig = px.bar(category_df, x = "Region", y = "Before tax Inv Amt (HKD)",text_auto='.2s')
              fig.update_traces(text = filter_df["Region"], textposition= "inside")
              fig.update_traces(marker_color = 'orange', marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(fig, use_container_width=True, height = 150, weight = 300)
              
#PAR CHART of Regional INVOICE AMOUNT FY to FY
       with right_column:              
              st.subheader(":bar_chart: Regional Inv Amount_FY to FY(Available to show :orange[Multiple FY]):")
              category2_df = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by = ["FY_INV","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum().sort_values(by="Before tax Inv Amt (HKD)",ascending=False)
              df_contract_vs_invoice = px.bar(category2_df, x="FY_INV", y="Before tax Inv Amt (HKD)", color="Region",text_auto='.2s')
              df_contract_vs_invoice.update_traces(marker_line_color = 'black',
              marker_line_width = 2, opacity = 1)
              st.plotly_chart(df_contract_vs_invoice, use_container_width=True)
              #with st.expander("Regional_Invoice_Amount(HKD)_View_Actual_Figures"):
                #st.write(category_df.style.background_gradient(cmap="Greens"))
                #csv = category_df.to_csv(index = False).encode('utf-8')
                #st.download_button("Download Data", data = csv, file_name= "Regional_Sales.csv", mime = "text/csv",
                #            help = 'Click here to download the data as a Excel file')
              
       #FY Quarter Regional Invoice Details:
       st.subheader(":point_down: Regional Invoice Amount Subtotal_FY to FY:")
       with st.expander("Click to expand"):
              pvt7 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["Region"],columns=["FY_INV"],
              aggfunc="sum",fill_value=0).sort_values(by="Region",ascending=False)
              st.dataframe(pvt7.style.format("HKD{:,}"), use_container_width=True)
       
       st.divider()

# LINE CHART of Regional Comparision
       st.subheader(":chart_with_upwards_trend: All Region Inv Amount Trend_FY to FY(Available to show :orange[Multiple Regions]):")
       InvoiceAmount_df2 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).groupby(by = ["FQ(Invoice)","Region"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig2 = px.line(InvoiceAmount_df2,
               x = "FQ(Invoice)",
               y = "Before tax Inv Amt (HKD)",
               color='Region',
               markers=True,
               text="Before tax Inv Amt (HKD)",
               )
       fig2.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig2, use_container_width=True)

#All Region Invoice Details FQ_FQ:
       st.subheader(":point_down: All Region Invoice Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt7 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=["Region"],columns=["FQ(Invoice)"],
              aggfunc="sum",fill_value=0).sort_values(by="Region",ascending=False)
              st.dataframe(pvt7.style.highlight_max(color = 'yellow', axis = 0).highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of SOUTH CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[SOUTH CHINA] Inv Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "SOUTH"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","FQ(Invoice)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig4 = px.line(df_Single_region,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)", 
              #title= "Inv Amount FY to FY(Multiple FY selection available)",
              color='FY_INV',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig4.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig4.update_layout(yaxis_showticklabels = False), use_container_width=True)
#SOUTH Region Invoice Details FQ_FQ:
       st.subheader(":point_down: SOUTH CHINA Invoice Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt8 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "SOUTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_INV',ascending=False)
              st.dataframe(pvt8.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of EAST CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[EAST CHINA] Inv Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY)]:")
       df_Single_region = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "EAST"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","FQ(Invoice)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig5 = px.line(df_Single_region,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)", 
              #title= "Inv Amount FY to FY(Multiple FY selection available)",
              color='FY_INV',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig5.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig5.update_layout(yaxis_showticklabels = False), use_container_width=True)
#EAST Region Invoice Details FQ_FQ:
       st.subheader(":point_down: EAST CHINA Invoice Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt9 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "EAST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_INV',ascending=False)
              st.dataframe(pvt9.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of WEST CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[WEST CHINA] Inv Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "WEST"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","FQ(Invoice)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig6 = px.line(df_Single_region,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)", 
              #title= "Inv Amount FY to FY(Multiple FY selection available)",
              color='FY_INV',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig6.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig6.update_layout(yaxis_showticklabels = False), use_container_width=True)
#WEST Region Invoice Details FQ_FQ:
       st.subheader(":point_down: WEST CHINA Invoice Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt10 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "WEST"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_INV',ascending=False)
              st.dataframe(pvt10.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
       st.divider()
# LINE CHART of NORTH CHINA FY/FY
       st.subheader(":chart_with_upwards_trend: :orange[NORTH CHINA] Inv Amount Trend_FQ to FQ(Available to Show :orange[Multiple FY]):")
       df_Single_region = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "NORTH"').query('FY_INV != "TBA"').round(0).groupby(by = ["FY_INV","FQ(Invoice)"], as_index= False)["Before tax Inv Amt (HKD)"].sum()
       fig7 = px.line(df_Single_region,
              x= "FQ(Invoice)",
              y = "Before tax Inv Amt (HKD)", 
              #title= "Inv Amount FY to FY(Multiple FY selection available)",
              color='FY_INV',
              markers=True, 
              text="Before tax Inv Amt (HKD)", 
              )
       fig7.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
       st.plotly_chart(fig7.update_layout(yaxis_showticklabels = False), use_container_width=True)
#NORTH Region Invoice Details FQ_FQ:
       st.subheader(":point_down: NORTH CHINA Invoice Amount Subtotal_FQ to FQ:")
       with st.expander("Click to expand"):
              pvt10 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Region == "NORTH"').round(0).pivot_table(values="Before tax Inv Amt (HKD)",index=['FY_INV'],columns=["FQ(Invoice)"],
              aggfunc="sum",fill_value=0).sort_values(by='FY_INV',ascending=False)
              st.dataframe(pvt10.style.highlight_between(left=-1, right=1, props='font-weight:bold;color:red').format("HKD{:,}"), use_container_width=True)
############################################################################################################################################################################################################    
#TAB 3: Brand category 
with tab3:
       #Brand Inv Qty by Inv Month:
       st.subheader(":point_down: Brand Item Inv Qty_:orange[Inv Month]:")
       #with st.expander("Click to expand"):
       pvt6 = filter_df.query('FY_INV != "TBA"').query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').round(0).pivot_table(values="Item Qty",index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month","BRAND","Ordered_Items"],
       aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt6.style.format("{:,}").highlight_max(color = 'yellow', axis = 0), use_container_width=True) 

       left_column, middle_column , right_column= st.columns(3)
       with left_column:
              st.subheader(":bar_chart: Overall Sales_Brand Qty:")
              sales_by_brand_line = (
              filter_df.query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').groupby(by=["BRAND"])[["Item Qty"]].sum().sort_values(by="Item Qty")
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
                     st.subheader(":trophy: :orange[Brand List]_Inv Amount& Qty:")
                     pvt4 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(index=["BRAND"],values=["Before tax Inv Amt (HKD)","Item Qty"],
                     aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
                     st.dataframe(pvt4, use_container_width=True)
              
              with right_column:
                     st.subheader(":sports_medal: :orange[Item List]_Inv Amount& Qty:")
                     pvt5 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(index=["Ordered_Items"],values=["Before tax Inv Amt (HKD)","Item Qty"],
                     aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
                     st.dataframe(pvt5, use_container_width=True)

       st.subheader(":clipboard: Brand Inv Qty_FY to FY(Available to show :orange[Multiple FY]):")
       pvt8 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(index=["BRAND"],values="Item Qty",columns=["FY_INV"],
       aggfunc="sum",fill_value=0,).sort_index(axis=0, ascending=False)#.sort_values(by="Item Qty",ascending=False)
       st.dataframe(pvt8, use_container_width=True)


#Top Product line chart invoice qty trend
       left_column, middle_column, right_column = st.columns(3)
       with middle_column:
              st.subheader(":chart_with_upwards_trend: :orange[YAMAHA Inv Qty] Trend_FQ to FQ:")

       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM20R Invoice Details:
       with left_column:
              df_YSM20 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM20R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"], 
                                   as_index= False)["Item Qty"].sum()
              fig8 = px.line(df_YSM20,
              x= "FQ(Invoice)",
              y = "Item Qty", 
              color='FY_INV',
              symbol="FY_INV",
              title="YSM20R",
              markers=True, 
              
              text="Item Qty", 
              )
              fig8.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig8.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM20R Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt13 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM20R"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt13.style.format("{:,}"), use_container_width=True)
#Line Chart FY to FY YSM10 Invoice Details:
       with right_column:
              df_YSM10 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM10"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"], 
                                   as_index= False)["Item Qty"].sum()
              fig9 = px.line(df_YSM10,
              x= "FQ(Invoice)",
              y = "Item Qty", 
              color='FY_INV',
              symbol="FY_INV",
              title="YSM10",
              markers=True, 
              text="Item Qty", 
              )
              fig9.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig9.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM10 Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt14 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM10"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt14.style.format("{:,}"), use_container_width=True)


#Second row for Top product trend
       left_column, right_column = st.columns(2)
#Line Chart FY to FY YSM40R Invoice Details:
       with left_column:
              df_YSM40 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSM40R"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"], 
                                   as_index= False)["Item Qty"].sum()
              fig10 = px.line(df_YSM40,
              x= "FQ(Invoice)",
              y = "Item Qty", 
              color='FY_INV',
              symbol="FY_INV",
              title="YSM40R",
              markers=True, 
              text="Item Qty", 
              )
              fig10.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig10.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM40R Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt14 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSM40R"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt14.style.format("{:,}"), use_container_width=True)


#Line Chart FY to FY YSi-V Invoice Details:
       with right_column:
              df_YSIV = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('Ordered_Items == "YSi-V"').query('FY_INV != "TBA"').round(0).groupby(by = ["FQ(Invoice)","FY_INV"], 
                                   as_index= False)["Item Qty"].sum()
              fig11 = px.line(df_YSIV,
              x= "FQ(Invoice)",
              y = "Item Qty", 
              color='FY_INV',
              symbol="FY_INV",
              title="YSi-V",
              markers=True, 
              text="Item Qty", 
              )
              fig11.update_traces(marker_size=9, textposition="top center", texttemplate='%{text:.2s}')
              st.plotly_chart(fig11.update_layout(yaxis_showticklabels = False), use_container_width=True)
#FY to FY YSM40R Invoice Details:
              with st.expander(":point_right:  Click to expand"):
                     pvt15 = filter_df.query('FY_INV != "TBA"').query('FY_INV != "Cancel"').query('FY_INV != "TBA"').query('Ordered_Items == "YSi-V"').pivot_table(values="Item Qty",index=["FY_INV"],columns=["FQ(Invoice)"],
                     aggfunc="sum",fill_value=0, margins=True,margins_name="Total").sort_index(axis=0, ascending=False)
                     st.dataframe(pvt15.style.format("{:,}"), use_container_width=True)
         
############################################################################################################################################################################################################  
#TAB 4: Customer category
with tab4:
       #Top Down Customer details Table 
       st.subheader(":point_down: Customer Invoice Details_:orange[Inv Month]")
       #with st.expander("Click to expand"):
       pvt3 = filter_df.query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(0).pivot_table(index=["FY_INV","FQ(Invoice)","Inv_Yr","Inv_Month","COST_CENTRE","Region","Contract No.","Customer Name","Ordered_Items"],values=["Before tax Inv Amt (HKD)","Item Qty"],
       aggfunc="sum",
       fill_value=0,
       margins=True,
       margins_name="Total").sort_index(axis=0, ascending=False)
       st.dataframe(pvt3.style.highlight_max(color = 'yellow', axis = 0).format("{:,}"), use_container_width=True)

       left_column, middle_column , right_column= st.columns(3)
#BAR CHART Customer List 
       with left_column:
              st.subheader(":bar_chart: :orange[Top 10 Customer]_Invoice Amount:")
              customer_line = (
              filter_df.query('BRAND != "C66 SERVICE"').query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').groupby(by=["Customer Name"])[["Before tax Inv Amt (HKD)"]]
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
              st.subheader(":trophy: :orange[Top Customer List]_Inv Amount& Qty:")
              pvt11 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region"],
                     values=["Item Qty","Before tax Inv Amt (HKD)"],
              aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
              st.dataframe(pvt11.style.format("{:,}"), use_container_width=True)

       with right_column:
              st.subheader(":medal: :orange[Top Customer Purchase List]_Inv Amount& Qty:")
              pvt12 = filter_df.query('Product_Type != "SERVICE/ PARTS"').query('Inv_Yr != "TBA"').query('Inv_Month != "TBA"').query('Inv_Month != "Cancel"').round(2).pivot_table(index=["Customer Name","Region","Ordered_Items"],
                     values=["Item Qty","Before tax Inv Amt (HKD)"],
              aggfunc="sum",fill_value=0,).sort_values(by="Item Qty",ascending=False)
              st.dataframe(pvt12.style.format("{:,}"), use_container_width=True)

          
############################################################################################################################################################################################################  



#fig3, ax = plt.subplots(figsize = (9,4))
#ax.plot(filter_df.Inv_Month, filter_df.Inv_Month, linewidth = 2,color='red',maker='o',markersize = 6, label ='South')
#ax.plot(filter_df.Inv_Month, filter_df.East, linewidth = 2,color='blue',maker='o',markersize = 6, label ='East')
#ax.plot(filter_df.Inv_Month, filter_df.West, linewidth = 2,color='green',maker='o',markersize = 6, label ='West')
#ax.plot(filter_df.Inv_Month, filter_df.North, linewidth = 2,color='yellow',maker='o',markersize = 6, label ='North')
#ax.set_title("Monthly Invoice Amount HKD by Region")
#ax.set_xlabel('Invoice Month')
#ax.set_ylabel('Invoice Amount HKD')
#plt.xticks(rotation = 45)
#ax.legend(loc = 'lower left', fontsize = 10)
#st.pyplot(fig3)
 

#success info
#st.success("Executed successfully")
#st.info("This is an information")
#st.warning("This is a warning")
#st.error("An error occured")

 
######################################################################################################

 
