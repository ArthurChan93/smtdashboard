import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# 設定頁面配置
st.set_page_config(layout="wide", page_title="合約分析系統")

# 固定讀取Excel文件
try:
    # 設定工作路徑
 #  os.chdir(r"D:\ArthurChan\OneDrive - Electronic Scientific Engineering Ltd\Monthly report(one drive)")
    
    # 直接讀取Excel文件並轉換關鍵欄位為字串
    df = pd.read_excel(
        io='Monthly_report_for_edit.xlsm',
        engine='openpyxl',
        sheet_name='raw_sheet',
        skiprows=0,
        usecols='A:AU',
        nrows=100000,
        dtype={'FY_Contract': str, 'FY_INV': str, 'Inv_Yr': str, 'Inv_Month': str}  # 新增型別指定
    )
    
    # 轉換Region欄位為字串
    df['Region'] = df['Region'].astype(str)
    
    # 應用過濾條件
    df = df.query('Region != "C66 N/A"'
                ).query('FY_Contract != "Cancel"'
                      ).query('FY_INV != "TBA"'
                            ).query('FY_INV != "FY 17/18"'
                                  ).query('FY_INV != "Cancel"'
                                        ).query('Inv_Yr != "TBA"'
                                              ).query('Inv_Yr != "Cancel"'
                                                    ).query('Inv_Month != "TBA"'
                                                          ).query('Inv_Month != "Cancel"')
    
    # 欄位名稱映射
    columns_mapping = {
        'FY_INV': '財年',
        'COST_CENTRE': '成本中心',
        'Curr': '合同貨幣',
        'Signed Contract Amt': '開票總金額(非港元)',
        'Before tax Inv Amt (HKD)': '開票總金額(港元)',
        'Inv_Yr': '開票年份',
        'Inv_Month': '開票月份',
        'Region': '地區',
        'Delivery_Terms': '運輸條款'
    }
    df = df.rename(columns=columns_mapping)
    
    # 數據清洗流程(強化型別轉換)
    invalid_values = ['Cancel', 'TBA', 'nan', '', np.nan]
    df['開票年份'] = pd.to_numeric(df['開票年份'], errors='coerce').replace(invalid_values, pd.NA)
    df['開票月份'] = df['開票月份'].astype(str).str.strip().replace(invalid_values, pd.NA)
    df['地區'] = df['地區'].astype(str).replace(invalid_values, pd.NA)
    df['運輸條款'] = df['運輸條款'].astype(str).replace(invalid_values, pd.NA)
    
    # 移除無效資料
    df = df.dropna(subset=['開票年份', '開票月份', '地區'])
    df['開票年份'] = df['開票年份'].astype(int)
    
    # 側邊欄篩選器(強化型別處理)
    with st.sidebar:
        st.header("篩選條件")
        
        # 匯率設定
        jpy_rate = st.number_input(
            "JPY兌人民幣匯率", 
            min_value=0.0, 
            value=0.045,
            format="%.4f"
        )
        
        # 財年篩選(確保字串比較)
        all_fy = sorted(df['財年'].astype(str).unique())
        selected_fy = st.multiselect(
            "選擇財年", 
            options=all_fy,
            default=['FY 24/25']
        )
        
        # 開票年份篩選(確保整數型別)
        valid_years = sorted(df['開票年份'].unique())
        selected_years = st.multiselect(
            "選擇開票年份",
            options=valid_years,
            default=[2024, 2025]
        )
        
        # 開票月份篩選(統一字串處理)
        valid_months = sorted(
            df['開票月份'].astype(str).unique(),
            key=lambda x: int(x) if x.isdigit() else 0
        )
        selected_months = st.multiselect(
            "選擇開票月份",
            options=valid_months,
            default=[str(m) for m in range(1,13) if m != 3]
        )
        
        # 地區篩選
        all_regions = sorted(df['地區'].unique())
        selected_regions = st.multiselect(
            "選擇地區",
            options=all_regions,
            default=['SOUTH']
        )
        
        # 運輸條款篩選
        all_delivery_terms = sorted(df['運輸條款'].dropna().unique())
        selected_delivery_terms = st.multiselect(
            "選擇運輸條款",
            options=all_delivery_terms
        )
    
    # 篩選條件組合(統一型別處理)
    filter_conditions = []
    if selected_fy: 
        filter_conditions.append(df['財年'].astype(str).isin(selected_fy))
    if selected_years: 
        filter_conditions.append(df['開票年份'].isin(selected_years))
    if selected_months: 
        filter_conditions.append(df['開票月份'].astype(str).isin(selected_months))
    if selected_regions: 
        filter_conditions.append(df['地區'].isin(selected_regions))
    if selected_delivery_terms: 
        filter_conditions.append(df['運輸條款'].isin(selected_delivery_terms))
    
    # 應用篩選
    if filter_conditions:
        filtered_df = df[np.logical_and.reduce(filter_conditions)]
    else:
        filtered_df = df
    
    # 主顯示區(保持不變)
    col1, col2 = st.columns([1, 3])
    
    with col1:
        st.subheader("總開票金額")
        total_hkd = filtered_df['開票總金額(港元)'].sum()
        st.metric("總金額 (HKD)", f"HK$ {total_hkd:,.2f}")
        
        st.subheader("各Cost Centre金額")
        cost_center_sum = filtered_df.groupby('成本中心')['開票總金額(港元)'].sum().reset_index()
        st.dataframe(
            cost_center_sum.style.format({'開票總金額(港元)': "HK$ {:,.2f}"}),
            hide_index=True
        )
    
    with col2:
        st.subheader("合同貨幣分布")
        if not filtered_df.empty:
            currency_sum = filtered_df.groupby('合同貨幣')['開票總金額(非港元)'].sum().reset_index()
            total = currency_sum['開票總金額(非港元)'].sum()
            currency_sum['百分比 (%)'] = (currency_sum['開票總金額(非港元)'] / total * 100).round(2)
            
            # 創建餅圖
            fig, ax = plt.subplots(figsize=(6, 4))
            
            # 顏色配置
            color_mapping = {
                'JPY': '#90EE90',  # 淺綠色
                'RMB': '#FFA07A'   # 淺橙色
            }
            base_colors = ['#66b3ff', '#ff9999', '#c2c2f0', '#ffb3e6', '#c4e8c4']
            
            # 生成顏色列表
            colors = []
            for currency in currency_sum['合同貨幣']:
                if currency in color_mapping:
                    colors.append(color_mapping[currency])
                else:
                    colors.append(base_colors[len(colors) % len(base_colors)])
            
            def autopct_format(pct):
                return f'{pct:.2f}%' if pct >= 1 else ''
            
            wedges, texts, autotexts = ax.pie(
                currency_sum['開票總金額(非港元)'],
                autopct=autopct_format,
                startangle=90,
                colors=colors,
                pctdistance=0.8,
                textprops={'color': 'black', 'weight': 'bold'},
                wedgeprops={'width': 0.4, 'linewidth': 1, 'edgecolor': 'white'}
            )
            
            ax.legend(wedges, currency_sum['合同貨幣'], title="合同貨幣", loc="upper right", bbox_to_anchor=(1.3, 1))
            ax.axis('equal')
            st.pyplot(fig)

            # JPY計算公式
            if 'JPY' in currency_sum['合同貨幣'].values:
                jpy_amount = currency_sum.loc[currency_sum['合同貨幣'] == 'JPY', '開票總金額(非港元)'].values[0]
                currency_sum['JPY 0.35%計算'] = currency_sum['合同貨幣'].apply(
                    lambda x: jpy_amount * 0.0035 if x == 'JPY' else None
                )
                currency_sum['JPY兌人民幣'] = currency_sum['合同貨幣'].apply(
                    lambda x: currency_sum.loc[currency_sum['合同貨幣'] == 'JPY', 'JPY 0.35%計算'].values[0] * jpy_rate if x == 'JPY' else None
                )
            
            st.dataframe(
                currency_sum.style.format({
                    '開票總金額(非港元)': "{:,.2f}",
                    '百分比 (%)': "{:.2f}%",
                    'JPY 0.35%計算': "JP¥ {:,.3f}",
                    'JPY兌人民幣': "¥ {:,.2f}" 
                }),
                hide_index=True,
                height=400
            )
        else:
            st.warning("沒有符合篩選條件的數據")

except Exception as e:
    st.error(f"檔案讀取失敗，錯誤訊息：{str(e)}")
    st.info("請確認以下事項：")
    st.write("1. 檔案路徑是否正確：D:\\ArthurChan\\OneDrive - Electronic Scientific Engineering Ltd\\Monthly report(one drive)\\Monthly_report_for_edit.xlsm")
    st.write("2. Excel文件是否包含'sheet_name='raw_sheet''工作表")
    st.write("3. Excel文件A:AU欄位格式是否符合預期")
    st.write("4. 確認Excel欄位無混合資料型別(數值與文字混用)")
