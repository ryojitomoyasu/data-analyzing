import streamlit as st
import pandas as pd
import numpy as np
import sys
import os

# Ensure root directory is in path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.append(parent_dir)

try:
    from data_processor import clean_data
except ImportError:
    st.warning("data_processor module not found. Some features may be limited.")
    def clean_data(df): return df

# Page Configuration
st.set_page_config(page_title="Order Calculation", page_icon="📦", layout="wide")

st.title("📦 発注数量計算システム (Order Calculation)")
st.markdown("""
このページでは、過去の売上実績と現在の在庫・入荷予定を基に、
**最適な発注数量（Suggested Order Quantity）** を自動計算します。
""")

# 1. Access Sales Data from Session State (INDEPENDENT)
if "df_order_sales" not in st.session_state:
    st.session_state.df_order_sales = None

# Option to Load from Main Dashboard (Convenience copy)
if st.session_state.df_order_sales is None and "df" in st.session_state and st.session_state.df is not None:
    # Small button in sidebar or main area? Keeping it here for visibility if empty
    if st.button("📥 Main Dashboardのデータをコピーして使用"):
        st.session_state.df_order_sales = st.session_state.df.copy()
        st.success("Main Dashboardのデータをコピーしました。")
        st.rerun()

# Always allow updating Sales Data via Expander
# [Refactor] Sales History is OPTIONAL if the user provides Avg Sales in the inventory file.
with st.expander("📂 売上履歴データ (Sales History) を変更・ロード", expanded=False):
    st.info("""
    **こちらは「過去の売上実績」用です。**
    適正在庫の基準となる「平均月販」を自動計算する場合に使用します。
    ※ すでにExcelファイルに「平均月販」が含まれている場合は、このデータをロードする必要はありません。
    """)
    sales_file = st.file_uploader("売上履歴データ (Sales History)", type=["xlsx", "csv"], key="sales_upload_order_page_main")
    if sales_file:
        try:
            if sales_file.name.endswith('.csv'):
                df_load = pd.read_csv(sales_file)
            else:
                df_load = pd.read_excel(sales_file)
            
            st.session_state.df_order_sales = clean_data(df_load)
            st.success("✅ 売上データを読み込みました！（Order Calc専用）")
            st.rerun()
        except Exception as e:
            st.error(f"データ読み込みエラー: {e}")

df_sales = st.session_state.df_order_sales



# 2. Inventory Data Upload
# [Refactor] Wrap in Expander to match Sales History UI
st.divider() 
# Initialize Session State for Inventory if not exists
if 'df_inventory' not in st.session_state:
    st.session_state.df_inventory = None
if 'df_inventory_name' not in st.session_state:
    st.session_state.df_inventory_name = ""

with st.expander("📂 在庫・入荷状況 (Inventory Status) を変更・ロード", expanded=st.session_state.df_inventory is None):
    st.info("""
    **こちらは「現在の在庫状況」用です。**
    商品ごとの現在の在庫数と、発注済みの入荷予定数を含むファイルをアップロードしてください。
    """)
    inventory_file = st.file_uploader("現在在庫・入荷データ (Current Stock Status)", type=["xlsx", "csv"])
    
    if inventory_file:
        # Only process if it's a new file or not yet loaded
        if st.session_state.df_inventory is None or st.session_state.df_inventory_name != inventory_file.name:
            try:
                if inventory_file.name.endswith('.csv'):
                    df_inv_load = pd.read_csv(inventory_file)
                else:
                    df_inv_load = pd.read_excel(inventory_file)
                
                st.session_state.df_inventory = df_inv_load
                st.session_state.df_inventory_name = inventory_file.name
                st.success(f"在庫データ読み込み完了: {inventory_file.name}")
                st.rerun()
            except Exception as e:
                st.error(f"ファイル読み込みエラー: {e}")

# 3. Calculation Logic
if st.session_state.df_inventory is not None:
    df_inv = st.session_state.df_inventory
    st.success(f"✅ 使用中の在庫データ: {st.session_state.df_inventory_name} ({len(df_inv)} 件)")
    
    try:
        inv_columns = df_inv.columns.tolist()

        # --- Column Mapping Section (Left / Sidebar) ---
        with st.sidebar:
            st.header("🛠️ 列のマッピング設定")
            st.info("アップロードしたファイルのどの列が各項目に当たるかを選択してください。")
            
            # Default guesses (Helper)
            def get_index(options, keywords):
                for i, opt in enumerate(options):
                    if any(k in str(opt).lower() for k in keywords):
                        return i
                return 0

            # 0. Item Number (New)
            idx_id = get_index(inv_columns, ['id', 'no', 'code', 'number', '番号', 'コード'])
            inv_id_col = st.selectbox("商品コード/Item# (Item No.)", inv_columns, index=idx_id, key="map_id")

            # 1. Item Name
            idx_item = get_index(inv_columns, ['item', 'product', 'description', 'name', '商品', '品名'])
            inv_item_col = st.selectbox("商品名 (Item Name)", inv_columns, index=idx_item, key="map_item")
            
            # 2. Current Stock
            idx_stock = get_index(inv_columns, ['current', 'stock', '在庫', '现有'])
            current_stock_col = st.selectbox("現在在庫 (Current Stock)", inv_columns, index=idx_stock, key="map_stock")
            
            # 3. Incoming
            idx_incoming = get_index(inv_columns, ['incoming', 'po', '入荷', '発注残'])
            incoming_col = st.selectbox("入荷予定 (Incoming)", inv_columns, index=idx_incoming, key="map_incoming")
            
            st.divider()
            
            # 4. Average Monthly Sales Source
            avg_sales_source = st.radio("平均月販の参照元", ["売上データから自動計算", "アップロードファイルの列を使用"], key="map_source")
            
            calc_window = "直近3ヶ月" # Default
            avg_sales_inv_col = None
            
            if avg_sales_source == "売上データから自動計算":
                st.caption("Sales Dashboardのデータを使用します")
                calc_window = st.selectbox("平均月販の算出期間", ["直近3ヶ月", "直近6ヶ月", "12ヶ月", "全期間"], index=0, key="map_window")
                # Parameters for Calc
                safety_months = st.number_input("適正在庫月数 (Safety Stock Months)", min_value=0.1, max_value=12.0, value=1.5, step=0.1, key="map_safety_calc")
                
            else:
                st.caption("ファイル内の数値を平均月販として使用")
                idx_avg = get_index(inv_columns, ['avg', 'sales', 'average', '月販', '平均'])
                avg_sales_inv_col = st.selectbox("平均月販列 (Avg Sales Column)", inv_columns, index=idx_avg, key="map_avg_col")
                safety_months = st.number_input("適正在庫月数 (Safety Stock Months)", min_value=0.1, max_value=12.0, value=1.5, step=0.1, key="map_safety_file")

        # Main Info Display
        st.info(f"""
        **計算設定:**
        - **適正在庫:** 月販 × {safety_months} ヶ月分
        - **計算式:** (平均月販 × {safety_months}) - ({current_stock_col} + {incoming_col})
        """)

        # Proceed with Logic
        if not inv_item_col:
            st.error("商品名列が選択されていません。")
        else:
            # Normalize Item Names for Merge
            df_inv['Item_Join'] = df_inv[inv_item_col].astype(str).str.strip()
            
            # Get Avg Sales
            if avg_sales_source == "売上データから自動計算":
                # Ensure Sales Data is loaded
                if df_sales is None:
                    st.error("⚠️ 「売上データから自動計算」を選択しましたが、売上履歴データがロードされていません。")
                    st.info("ページ上部の「📂 売上履歴データ」からファイルをロードするか、平均月販の参照元を「アップロードファイルの列を使用」に変更してください。")
                    st.stop()
                
                # Identify Columns (Simple heuristic matching app.py logic)
                item_col = next((col for col in df_sales.columns if any(x in col.lower() for x in ['item', 'product', '商品', '品名'])), None)
                val_col = next((col for col in df_sales.columns if any(x in col.lower() for x in ['value', 'sales', 'amount', '売上', '金額'])), None)
                month_col = next((col for col in df_sales.columns if any(x in col.lower() for x in ['month', '月'])), None)
                year_col = next((col for col in df_sales.columns if any(x in col.lower() for x in ['year', '年'])), None)
                
                if not item_col or not val_col:
                     st.error("⚠️ 売上データから「商品名」または「売上数量/金額」の列を特定できませんでした。")
                     st.stop()

                df_sales['Item_Join'] = df_sales[item_col].astype(str).str.strip()
                # ... (Existing Calculation Logic) ...
                if year_col and month_col:
                    df_sales['Period'] = df_sales[year_col].astype(str) + '-' + df_sales[month_col].astype(str).str.zfill(2)
                    periods = sorted(df_sales['Period'].unique())
                    if "3ヶ月" in calc_window: target_periods = periods[-3:]
                    elif "6ヶ月" in calc_window: target_periods = periods[-6:]
                    elif "12ヶ月" in calc_window: target_periods = periods[-12:]
                    else: target_periods = periods
                    
                    df_calc = df_sales[df_sales['Period'].isin(target_periods)]
                    months_count = len(target_periods)
                else:
                    df_calc = df_sales
                    months_count = 1 
                
                avg_sales = df_calc.groupby('Item_Join')[val_col].sum().reset_index()
                avg_sales.rename(columns={val_col: 'Total_Sales_Window'}, inplace=True)
                avg_sales['Avg_Monthly_Sales'] = avg_sales['Total_Sales_Window'] / max(1, months_count)
                
                # Merge
                df_merged = pd.merge(df_inv, avg_sales, on='Item_Join', how='left')
                df_merged['Avg_Monthly_Sales'] = df_merged['Avg_Monthly_Sales'].fillna(0)
                
            else:
                # Use Column from Inventory File
                df_merged = df_inv.copy()
                # Ensure numeric
                df_merged[avg_sales_inv_col] = pd.to_numeric(df_merged[avg_sales_inv_col], errors='coerce').fillna(0)
                df_merged['Avg_Monthly_Sales'] = df_merged[avg_sales_inv_col]

            # Ensure Numeric for Stock Cols
            df_merged[current_stock_col] = pd.to_numeric(df_merged[current_stock_col], errors='coerce').fillna(0)
            df_merged[incoming_col] = pd.to_numeric(df_merged[incoming_col], errors='coerce').fillna(0)
            
            # Calculate
            df_merged['Required_Stock'] = df_merged['Avg_Monthly_Sales'] * safety_months
            df_merged['Suggested_Order'] = df_merged['Required_Stock'] - (df_merged[current_stock_col] + df_merged[incoming_col])
            df_merged['Suggested_Order'] = df_merged['Suggested_Order'].apply(lambda x: max(0, x))
            
            # Formatting Display
            st.divider()
            st.subheader("3. 計算結果 (Calculated Orders)")
            
            # Select relevant columns
            # Added inv_id_col to the front
            disp_cols = [inv_id_col, inv_item_col, 'Avg_Monthly_Sales', current_stock_col, incoming_col, 'Required_Stock', 'Suggested_Order']
            
            # Handle duplicate columns if ID/Name/etc are the same
            disp_cols = list(dict.fromkeys(disp_cols)) # Remove duplicates preserving order

            df_res = df_merged[disp_cols].copy()
            
            # Renaming for clarity
            df_res.rename(columns={
                inv_id_col: 'Item#',
                inv_item_col: '商品名 (Name)',
                'Avg_Monthly_Sales': '平均月販 (Avg Sales)',
                'Required_Stock': '必要在庫数 (Required)',
                'Suggested_Order': '発注推奨数 (To Order)'
            }, inplace=True)
            
            # Format Dictionary
            format_dict = {
                'Item#': '{:.0f}', # No decimals for ID
                '平均月販 (Avg Sales)': '{:,.1f}',
                '必要在庫数 (Required)': '{:,.1f}', 
                '発注推奨数 (To Order)': '{:,.0f}',
                current_stock_col: '{:,.0f}',
                incoming_col: '{:,.0f}'
            }
            
            st.dataframe(
                df_res.style.format(format_dict).hide(axis="index"),
                use_container_width=True,
                hide_index=True # Explicitly hide index for Streamlit
            )
            
            csv = df_res.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 発注リストをダウンロード (CSV)", csv, "order_calculation_result.csv", "text/csv")
            
    except Exception as e:
        st.error(f"ファイル処理エラー: {e}")
else:
    st.info("👈 上記のアップローダーから在庫データをアップロードしてください。アップロード後に計算が開始されます。")
    st.markdown("""
    **アップロードするファイルのサンプル形式:**
    | Item Code | Item Name | Current Stock | Incoming |
    | :--- | :--- | :--- | :--- |
    | A001 | Product A | 100 | 50 |
    | B002 | Product B | 20 | 0 |
    """)
