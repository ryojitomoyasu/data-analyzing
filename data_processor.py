import streamlit as st
import pandas as pd
import numpy as np

@st.cache_data
def clean_data(df):
    """
    Cleans the uploaded dataframe: cleans numeric columns and converts date columns.
    Also attempts to find the correct header row if the first row is not the header.
    """
    # 1. Header Detection Strategy
    # Look for a row that contains typical column names like '日付' and '売上'
    search_keywords = ['日付', 'date', '売上', 'sales', '商品', 'product', 'カテゴリ', 'category', 'ブランド', 'brand']
    
    # Check if current columns already look good
    current_match_count = sum(1 for col in df.columns if any(k in str(col).lower() for k in search_keywords))
    
    # If few matches, scan first 10 rows to see if substantial header exists
    if current_match_count < 2:
        for i, row in df.head(10).iterrows():
            # Convert row values to string and check for keywords
            row_str = row.astype(str).str.lower()
            match_count = sum(1 for val in row_str if any(k in val for k in search_keywords))
            
            if match_count >= 2:
                # Found a better header! Reload/Adjust DataFrame
                # Set this row as columns
                new_header = row
                df = df.iloc[i+1:].copy()
                df.columns = new_header
                df = df.reset_index(drop=True)
                break

    # Clean numeric columns (remove commas, currency symbols, etc.)
    for col in df.columns:
        # Check if the column name itself is valid (not NaN or empty)
        if pd.isna(col) or str(col).strip() == '':
            continue
            
        if df[col].dtype == 'object':
            try:
                # Remove common non-numeric chars and convert to float
                # We save original first to check later if it was a date
                series_str = df[col].astype(str)
                # Clean FY prefix if present
                cleaned = series_str.str.replace(r'(?i)^\s*(?:fy|fiscal\s*year)\s*', '', regex=True)
                cleaned = cleaned.str.replace(r'[$,¥,]', '', regex=True).str.strip()
                cleaned = cleaned.replace('', '0')
                # If conversion successful for most items, keep it
                numeric_val = pd.to_numeric(cleaned, errors='coerce')
                # Only update if we didn't turn everything into NaN (unless it was already empty)
                if numeric_val.notna().sum() >= (len(cleaned) * 0.5):
                    df[col] = numeric_val
            except:
                pass
    
    # Auto-detect and convert datetime
    date_cols = [col for col in df.columns if isinstance(col, str) and any(x in col for x in ['日', 'date', '時点', 'time', '月'])]
    for col in date_cols:
        try:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        except:
            pass
            
    # Drop rows where critical columns are NaN (e.g. empty rows from Excel)
    df = df.dropna(how='all')
    
    return df

@st.cache_data
def calculate_kpis(df):
    """
    Calculates basic KPIs for the dashboard.
    Assumes existence of columns like 'Sales' or '売上', 'Date' or '日付'.
    """
    sales_col = next((col for col in df.columns if '売上' in col or 'sales' in col.lower()), None)
    date_col = next((col for col in df.columns if '日付' in col or 'date' in col.lower()), None)
    
    if not sales_col:
        return {}

    total_sales = df[sales_col].sum()
    avg_order_value = df[sales_col].mean()
    
    kpis = {
        "total_sales": total_sales,
        "avg_order_value": avg_order_value,
        "record_count": len(df)
    }
    
    # Monthly growth if date exists
    if date_col and pd.api.types.is_datetime64_any_dtype(df[date_col]):
        df_sorted = df.sort_values(date_col)
        monthly_sales = df_sorted.resample('M', on=date_col)[sales_col].sum()
        if len(monthly_sales) >= 2:
            prev_month = monthly_sales.iloc[-2]
            current_month = monthly_sales.iloc[-1]
            growth = ((current_month - prev_month) / prev_month) * 100 if prev_month != 0 else 0
            kpis["monthly_growth"] = growth
            
    return kpis

@st.cache_data
def infer_industry(df):
    """
    Infers the industry from product or category names (Legacy backup).
    """
    cat_col = next((col for col in df.columns if 'カテゴリ' in col or 'category' in col.lower() or '商品' in col or 'product' in col.lower()), None)
    if not cat_col:
        return "汎用"
    
    sample_items = df[cat_col].dropna().unique()[:10]
    return ", ".join(sample_items)

@st.cache_data
def calculate_abc_analysis(df):
    """
    Performs ABC analysis based on sales by category/product.
    Returns a dataframe with cumulative percentage and rank.
    """
    sales_col = next((col for col in df.columns if '売上' in col or 'sales' in col.lower()), None)
    cat_col = next((col for col in df.columns if 'カテゴリ' in col or 'category' in col.lower() or '商品' in col or 'product' in col.lower()), None)
    
    if not sales_col or not cat_col:
        return None
    
    # Group and sort
    abc_df = df.groupby(cat_col)[sales_col].sum().reset_index()
    abc_df = abc_df.sort_values(by=sales_col, ascending=False)
    
    # Calculate cumulative metrics
    total_sales = abc_df[sales_col].sum()
    abc_df['cumulative_sales'] = abc_df[sales_col].cumsum()
    abc_df['cumulative_percent'] = (abc_df['cumulative_sales'] / total_sales) * 100
    
    # Assign ABC rank
    def assign_rank(p):
        if p <= 70: return 'A'
        elif p <= 90: return 'B'
        else: return 'C'
        
    abc_df['rank'] = abc_df['cumulative_percent'].apply(assign_rank)
    return abc_df
