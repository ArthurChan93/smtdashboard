import streamlit as st
import pandas as pd

# Function to process the uploaded files
def process_files(south_file, east_file):
    # Read the sheets from the uploaded files
    south_df = pd.read_excel(south_file, sheet_name='Stock_list', usecols=['ETA_Month', 'Item', 'Customer Reserved', 'Machine_QTY', 'ETA HK'])
    stk_df = pd.read_excel(east_file, sheet_name='STK', usecols=['ETA ', 'MACHINE TYPE', 'QTY', 'Customer', '到货情况'])
    ind_df = pd.read_excel(east_file, sheet_name='IND', usecols=['ETA ', 'MACHINE TYPE', 'QTY', 'Customer', 'ETA '])

    # Rename columns for consistency
    south_df.rename(columns={'ETA_Month': 'Status', 'Item': 'Model', 'Customer Reserved': 'Customer', 'Machine_QTY': 'QTY', 'ETA HK': 'Incoming'}, inplace=True)
    stk_df.rename(columns={'到货情况': 'Status', 'MACHINE TYPE': 'Model', 'ETA ': 'Incoming'}, inplace=True)
    ind_df.rename(columns={'ETA ': 'Incoming', 'MACHINE TYPE': 'Model'}, inplace=True)

    # Filter out rows with "已发货"
    south_df = south_df[~south_df['Status'].str.contains('已发货', na=False)]
    stk_df = stk_df[~stk_df['Status'].str.contains('已发货', na=False)]

    # Combine the dataframes
    combined_df = pd.concat([south_df, stk_df, ind_df], ignore_index=True)

    # Standardize the 'Status' column
    combined_df['Status'] = combined_df['Status'].replace({'已到货': 'STOCK', '已经到货': 'STOCK', '已发货': None})
    combined_df.dropna(subset=['Status'], inplace=True)

    # Standardize the 'Incoming' column
    combined_df['Incoming'] = combined_df['Incoming'].apply(
        lambda x: 'TBA' if 'TBA' in str(x) else 
                  pd.to_datetime(x, errors='coerce').strftime('%b-%y').upper() if pd.notnull(pd.to_datetime(x, errors='coerce')) else x
    )

    # Further classification of 'Status' based on 'Incoming' column
    def classify_status(row):
        if row['Status'] == 'STOCK':
            return 'STOCK'
        elif 'TBA' in str(row['Incoming']):
            return 'TBA'
        elif pd.notnull(row['Incoming']) and '-' in str(row['Incoming']):
            return f"{row['Incoming']} Incoming"
        return None

    combined_df['Status'] = combined_df.apply(classify_status, axis=1)
    combined_df.dropna(subset=['Status'], inplace=True)

    # Standardize the 'Model' column
    combined_df['Model'] = combined_df['Model'].replace({
        'YSi-V(DL)': 'YSi-V', 'YSi-V(SL)': 'YSi-V',
        'YSM20R-2': 'YSM20R', 'YSM20R(PV)-2': 'YSM20R',
        'YSM20R-1': 'YSM20R', 'YSM20R(SV)-2': 'YSM20R',
        'YSM20R(PV)-1': 'YSM20R', 'YSM10 96': 'YSM10'
    })

    # Remove rows with missing or invalid 'Model'
    combined_df.dropna(subset=['Model'], inplace=True)

    return combined_df

# Function to pivot the data into the desired format
def pivot_data(df):
    # Group by Status and Model, summing QTY
    grouped = df.groupby(['Status', 'Model']).agg({'QTY': 'sum'}).reset_index()

    # Pivot the data
    pivot = grouped.pivot(index='Status', columns='Model', values='QTY').fillna(0).astype(int)  # Fill NaN with 0, convert to int
    pivot['Subtotal'] = pivot.sum(axis=1)  # Add Subtotal column

    # Add Grand Total row
    grand_total = pd.DataFrame(pivot.sum(axis=0)).T
    grand_total['Status'] = 'Grand Total'
    pivot = pd.concat([pivot.reset_index(), grand_total], ignore_index=True)

    # Ensure 'STOCK' is the first row and remove duplicates in categories
    unique_statuses = pivot['Status'].dropna().unique().tolist()
    if 'STOCK' in unique_statuses:
        unique_statuses.remove('STOCK')
    ordered_statuses = ['STOCK'] + unique_statuses
    pivot['Status'] = pd.Categorical(pivot['Status'], categories=ordered_statuses, ordered=True)
    pivot.sort_values(by='Status', inplace=True)
    pivot.reset_index(drop=True, inplace=True)

    return pivot

# Function to style the dataframe
def style_dataframe(df):
    def highlight_cells(val):
        if 'Subtotal' in str(val):
            return 'background-color: yellow'  # Yellow background for Subtotal
        elif 'STOCK' in str(val):
            return 'background-color: lightgreen'  # Light green background for STOCK
        elif 'TBA' in str(val):
            return 'background-color: orange'  # Orange background for TBA        
        elif val == 'Subtotal' or val == 'Grand Total':
            return 'background-color: yellow'  # Yellow background for Subtotal and Grand Total
        elif 'Incoming' in str(val):
            return 'background-color: lightpink'  # Pink background for Incoming
        return ''
        
    # Apply highlight styling to all relevant columns
    styled_df = df.style.applymap(highlight_cells)  # Apply to all cells
    # Format numeric columns
    for col in df.columns[1:]:
        styled_df.format({col: '{:.0f}'})  # Remove decimal points for QTY
    return styled_df



    # Apply highlight styling to the 'Status' and 'Subtotal' columns
    styled_df = df.style.applymap(highlight_cells, subset=['Status'])
    styled_df.applymap(highlight_cells, subset=['Subtotal'])

    # Format numeric columns
    for col in df.columns[1:]:
        styled_df.format({col: '{:.0f}'})  # Remove decimal points for QTY
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
        pivoted_df = pivot_data(combined_df)

        # Display the summary report
        st.markdown('<div class="report-title">Summary Report</div>', unsafe_allow_html=True)
        st.markdown(style_dataframe(pivoted_df).to_html(index=False), unsafe_allow_html=True)
        st.download_button('Download Summary Report', pivoted_df.to_csv(index=False), file_name='summary_report.csv', 
                           key='download_button', help='Download the summary report')
    else:
        st.warning('Please upload both files to proceed.')
