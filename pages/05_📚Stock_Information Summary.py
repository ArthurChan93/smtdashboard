import streamlit as st
import pandas as pd
import os

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

    # Custom sort function for Status
    def sort_status(status):
        if status == 'STOCK':
            return (0, None)  # STOCK always first
        if status == 'Grand Total':
            return (float('inf'), None)  # Grand Total always last
        if status == 'TBA':
            return (2, None)  # TBA second to last
        if 'Incoming' in status or 'OUT' in status:
            parts = status.split(' ')[0]  # Extract MM-YY part
            date = pd.to_datetime(parts, format='%b-%y', errors='coerce')
            if pd.notnull(date):
                return (1, date)  # Sort by date
        return (3, None)  # Other statuses

    # Add a custom sort key and sort
    pivot['SortKey'] = pivot['Status'].apply(sort_status)
    pivot = pivot.sort_values(by='SortKey').drop(columns=['SortKey'])
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

# Streamlit app layout
st.title('📊Machine Inventory and Incoming Status')

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
st.subheader(':orange[South]: STK info')
south_file = st.file_uploader('1️⃣ Upload South stock information file:point_down:', type='xlsx', key='south')
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="east-uploader">', unsafe_allow_html=True)
st.subheader(':blue[EAST & WEST & NORTH]: STK info')
east_file = st.file_uploader('2️⃣ Upload SHA stock information file:point_down:', type='xlsx', key='east')
st.markdown('</div>', unsafe_allow_html=True)

# Combine button
left_column, right_column = st.columns(2)
with left_column:
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


with right_column:
    # 添加文件上傳器
    monthly_report_file = st.file_uploader('3️⃣ Upload :red[Monthly Report_edit]: :point_down:', type=['xlsx', 'xlsm'], key='monthly_report')
    if monthly_report_file:
        # 當文件已成功上傳後，重新執行“Combine”功能
        if south_file and east_file:
            combined_df = process_files(south_file, east_file)
            pivoted_df = pivot_data(combined_df)
            
            # 創建新表格，刪除 `Status` 列中包含 `TBA` 的行並存儲
            tba_row = pivoted_df[pivoted_df['Status'] == 'TBA'].copy()
            modified_pivoted_df = pivoted_df[pivoted_df['Status'] != 'TBA'].copy()
            
            # 隱藏 Subtotal 列
            if 'Subtotal' in modified_pivoted_df.columns:
                modified_pivoted_df.drop(columns=['Subtotal'], inplace=True)
            
            # 在每個 `Incoming` 狀態之下添加一行 `OUT`
            incoming_indices = modified_pivoted_df.index[modified_pivoted_df['Status'].str.contains('Incoming')].tolist()
            rows_to_insert = []
            for idx in incoming_indices:
                out_row = modified_pivoted_df.iloc[idx].copy()
                # 修改 Status 為 "MMM-YY OUT"
                incoming_date = out_row['Status'].split(' ')[0]
                out_row['Status'] = f"{incoming_date} OUT"
                # 修改數量為負數
                for col in modified_pivoted_df.columns[1:]:
                    out_row[col] = -1 * out_row[col]
                rows_to_insert.append((idx + 1, out_row))
            
            # 插入行
            for idx, row in rows_to_insert[::-1]:  # 必須倒序插入，否則索引會錯亂
                modified_pivoted_df = pd.concat([modified_pivoted_df.iloc[:idx], row.to_frame().T, modified_pivoted_df.iloc[idx:]]).reset_index(drop=True)
            
            # 如果 Monthly Report 文件被提供，進一步處理
            monthly_df = pd.read_excel(monthly_report_file, sheet_name='raw_sheet')
            
            # 遍歷每個 `Incoming` 狀態，提取對應數據並填充 `OUT` 行
            for idx, row in modified_pivoted_df.iterrows():
                if 'Incoming' in row['Status']:
                    # 提取年月
                    incoming_date = row['Status'].split(' ')[0]
                    month_year = pd.to_datetime(incoming_date, format='%b-%y', errors='coerce')
                    if pd.notnull(month_year):
                        year = month_year.year
                        month = month_year.month
                        
                        # 過濾 `Monthly Report` 數據
                        filtered_df = monthly_df[
                            (monthly_df['Inv_Yr'] == year) &
                            (monthly_df['Inv_Month'] == month) &
                            (monthly_df['Ordered_Items'].isin(pivoted_df.columns[1:]))  # 排除非 Model 列
                        ]
                        # 計算每個 Model 的數量總和
                        model_sums = filtered_df.groupby('Ordered_Items')['Item Qty'].sum()
                        # 填充 `OUT` 行
                        for model, qty in model_sums.items():
                            modified_pivoted_df.loc[idx + 1, model] = -1 * qty  # +1 是對應 `OUT` 行，並顯示為負數
            
            # 將 NaN 替換為 0，轉換為整數
            modified_pivoted_df.fillna(0, inplace=True)
            for col in modified_pivoted_df.columns[1:]:
                modified_pivoted_df[col] = modified_pivoted_df[col].astype(int)
            
            # 移除已有的 Grand Total 行（如果存在）
            modified_pivoted_df = modified_pivoted_df[modified_pivoted_df['Status'] != 'Grand Total']
            
            # 計算新的 Grand Total 列，包括 TBA 行的數值
            grand_total_row = modified_pivoted_df.iloc[:, 1:].sum(axis=0)
            if not tba_row.empty:
                grand_total_row += tba_row.iloc[:, 1:].sum(axis=0)
            grand_total_row = grand_total_row.round(0).astype(int)  # 確保無小數點
            grand_total_row['Status'] = 'Grand Total'
            
            # 添加新的 Grand Total 行
            modified_pivoted_df = pd.concat([modified_pivoted_df, grand_total_row.to_frame().T], ignore_index=True)
            
            # **新增的功能：生成 "Balance" 表格**
            balance_df = modified_pivoted_df[modified_pivoted_df['Status'].str.contains('Incoming')].copy()
            balance_df['Status'] = balance_df['Status'].str.replace('Incoming', 'Balance')
            
            # 修正累積總和邏輯
            for idx, row in balance_df.iterrows():
                # 找到對應的 "Incoming" 行
                status_balance = row['Status']
                incoming_status = status_balance.replace('Balance', 'Incoming')
                out_status = status_balance.replace('Balance', 'OUT')

                # 找到對應行的索引
                incoming_idx = modified_pivoted_df[modified_pivoted_df['Status'] == incoming_status].index[0]
                out_idx = modified_pivoted_df[modified_pivoted_df['Status'] == out_status].index[0]

                # 計算從頭到當前行的加總
                cumulative_sums = modified_pivoted_df.iloc[:out_idx + 1, 1:].sum(axis=0)
                for col in balance_df.columns[1:]:
                    balance_df.loc[idx, col] = cumulative_sums[col]

            # 填充 NaN 值並確保數據為整數
            balance_df.fillna(0, inplace=True)
            for col in balance_df.columns[1:]:
                balance_df[col] = balance_df[col].astype(int)

            # **將 TBA 行添加到 Balance 表格最底部**
            if not tba_row.empty:
                tba_row['Status'] = 'TBA'
                tba_row.fillna(0, inplace=True)
                for col in tba_row.columns[1:]:
                    tba_row[col] = tba_row[col].astype(int)
                balance_df = pd.concat([balance_df, tba_row], ignore_index=True)

            # **為 Balance 行設置粉紅色背景，並為 TBA 行設置橙色背景**
            def style_dataframe_with_balance_and_tba(df):
                styled_df = df.style.apply(
                    lambda x: [
                        'background-color: pink' if 'Balance' in v else 'background-color: orange' if 'TBA' in v else ''
                        for v in x
                    ],
                    subset=['Status']
                )
                return styled_df

            # **顯示第一個表格**
            st.markdown('<div class="report-title">Modified Report</div>', unsafe_allow_html=True)
            st.markdown(style_dataframe(modified_pivoted_df).to_html(index=False), unsafe_allow_html=True)

            # **提供兩個表格的下載選項**
            st.download_button(
                'Download Modified Report',
                modified_pivoted_df.to_csv(index=False),
                file_name='modified_report.csv',
                key='download_modified_button',
                help='Download the modified report'
            )

            # **顯示新生成的 "Balance" 表格**
            st.markdown('<div class="report-title">Balance Report</div>', unsafe_allow_html=True)
            st.markdown(style_dataframe_with_balance_and_tba(balance_df).to_html(index=False), unsafe_allow_html=True)
            
            st.download_button(
                'Download Balance Report',
                balance_df.to_csv(index=False),
                file_name='balance_report.csv',
                key='download_balance_button',
                help='Download the balance report'
            )
        else:
            st.warning('Please upload both South and East stock information files and combine first.')



