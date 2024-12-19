import streamlit as st
import pandas as pd

# Function to process the uploaded files
def process_files(south_file, east_file):
    # Read the sheets from the uploaded files
    south_df = pd.read_excel(south_file, sheet_name='Stock_list', usecols=['ETA_Month', 'Item', 'Customer Reserved', 'Machine_QTY', 'ETA HK'])
    stk_df = pd.read_excel(east_file, sheet_name='STK', usecols=['ETA ', 'MACHINE TYPE', 'QTY', 'Customer', '到货情况'])
    ind_df = pd.read_excel(east_file, sheet_name='IND', usecols=['ETA ', 'MACHINE TYPE', 'QTY', 'Customer', 'ETA '])

    # Process the data according to the rules provided
    south_df.rename(columns={'ETA_Month': 'Status', 'Item': 'Model', 'Customer Reserved': 'Customer', 'Machine_QTY': 'QTY', 'ETA HK': 'Incoming'}, inplace=True)
    stk_df.rename(columns={'到货情况': 'Status', 'MACHINE TYPE': 'Model', 'ETA ': 'Incoming'}, inplace=True)
    ind_df.rename(columns={'ETA ': 'Incoming', 'MACHINE TYPE': 'Model'}, inplace=True)

    # Combine the dataframes
    combined_df = pd.concat([south_df, stk_df, ind_df], ignore_index=True)

    # Standardize the 'Status' column
    combined_df['Status'] = combined_df['Status'].replace({'已到货': 'STOCK', '已经到货': 'STOCK', '已发货': None})
    combined_df.dropna(subset=['Status'], inplace=True)

    # Standardize the 'Incoming' column
    combined_df['Incoming'] = combined_df['Incoming'].apply(lambda x: 'TBA' if 'TBA' in str(x) else pd.to_datetime(x, errors='coerce').strftime('%b-%y') if pd.notnull(pd.to_datetime(x, errors='coerce')) else x)

    # Standardize the 'Model' column
    combined_df['Model'] = combined_df['Model'].replace({'YSi-V(DL)': 'YSi-V', 'YSi-V(SL)': 'YSi-V', 
                                                         'YSM20R-2': 'YSM20R', 'YSM20R(PV)-2': 'YSM20R', 
                                                         'YSM20R(SV)-2': 'YSM20R', 'YSM20R(PV)-1': 'YSM20R',
                                                         'YSM10 96': 'YSM10'})

    return combined_df

# Function to add subtotals and grand total
def add_totals(df):
    subtotals = df.groupby('Status')['QTY'].sum().reset_index()
    subtotals['Model'] = "Subtotal"
    
    grand_total = pd.DataFrame([['Grand Total', 'Subtotal', df['QTY'].sum()]], columns=['Status', 'Model', 'QTY'])
    
    df_with_totals = pd.concat([df, subtotals], ignore_index=True)
    
    return pd.concat([df_with_totals, grand_total], ignore_index=True)

# Function to style the dataframe
def style_dataframe(df):
    def highlight_cells(val):
        color = ''
        if val == 'STOCK':
            color = 'background-color: lightgreen'
        elif val == 'Subtotal' or val == 'Grand Total':
            color = 'background-color: yellow'
        elif val == 'TBA':
            color = 'background-color: lightcoral'  # Light orange color
        return color

    # Apply highlight and remove decimal places for QTY
    styled_df = df.style.applymap(highlight_cells, subset=['Status', 'Model'])
    styled_df.format({'QTY': '{:.0f}'})  # Remove decimal points for QTY
    return styled_df

# Streamlit app layout
st.title('Machine Inventory and Incoming Status')

# Custom CSS for styling
st.markdown("""
    <style>
    .south-uploader {
        background-color: #FFE4B5;
        padding: 10px;
        border-radius: 5px;
    }
    .east-uploader {
        background-color: #ADD8E6;
        padding: 10px;
        border-radius: 5px;
    }
    .combine-button {
        color: red !important;
        border: 2px solid red !important;
        background: none;
        padding: 5px 10px;
        border-radius: 5px;
        cursor: pointer;
    }
    .download-button {
        color: blue !important;
        border: 2px solid blue !important;
        background: none;
        padding: 5px 10px;
        border-radius: 5px;
        cursor: pointer;
    }
    .report-title {
        font-size: 20px;
        font-weight: bold;
        text-decoration: underline;
    }
    </style>
    """, unsafe_allow_html=True)

# File uploaders with custom background colors
st.markdown('<div class="south-uploader">', unsafe_allow_html=True)
st.subheader('South STK info')
south_file = st.file_uploader('Upload South Stock List Excel file', type='xlsx', key='south')
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="east-uploader">', unsafe_allow_html=True)
st.subheader('EAST & WEST & NORTH STK info')
east_file = st.file_uploader('Upload East China-STK Machine IND Order info Excel file', type='xlsx', key='east')
st.markdown('</div>', unsafe_allow_html=True)

# Combine button
if st.button('Combine', key='combine_button'):
    if south_file and east_file:
        combined_df = process_files(south_file, east_file)
        
        # Generate the summary report
        stock_summary = combined_df[combined_df['Status'] == "STOCK"].groupby(['Model']).agg({'QTY': 'sum'}).reset_index()
        stock_summary['Status'] = "STOCK"
        stock_summary = stock_summary.sort_values(by='QTY', ascending=False)
        
        incoming_summary = combined_df[(combined_df['Status'] != "STOCK") & (combined_df['Incoming'] != "TBA")].groupby(['Incoming', 'Model']).agg({'QTY': 'sum'}).reset_index()
        incoming_summary.rename(columns={'Incoming': 'Status'}, inplace=True)
        incoming_summary = incoming_summary.sort_values(by='QTY', ascending=False)
        
        tba_summary = combined_df[combined_df['Incoming'] == "TBA"].groupby(['Model']).agg({'QTY': 'sum'}).reset_index()
        tba_summary['Status'] = "TBA"
        tba_summary = tba_summary.sort_values(by='QTY', ascending=False)
        
        summary = pd.concat([stock_summary, incoming_summary, tba_summary], ignore_index=True)
        summary = summary[['Status', 'Model', 'QTY']]
        
        summary_with_totals = add_totals(summary)
        
        # Display the summary report
        st.markdown('<div class="report-title">Summary Report</div>', unsafe_allow_html=True)
        st.markdown(style_dataframe(summary_with_totals).to_html(index=False), unsafe_allow_html=True)
        st.download_button('Download Summary Report', summary_with_totals.to_csv(index=False), file_name='summary_report.csv', 
                           key='download_button', help='Download the summary report')
    else:
        st.warning('Please upload both files to proceed.')
