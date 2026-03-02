import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import plotly.graph_objects as go
import sys
import os
import json
import io

# Ensure we can import from parent directory
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from data_processor import clean_data
except ImportError:
    # Fallback if running from root
    from data_processor import clean_data

# Page Config
st.set_page_config(page_title="Sales Dashboard", layout="wide", page_icon="📊")

# Custom CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .stMetric {
        background: rgba(255, 255, 255, 0.7) !important;
        backdrop-filter: blur(10px);
        border-radius: 15px !important;
        padding: 20px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    .main-title {
        font-size: 2.5rem;
        font-weight: 800;
        color: #1e3a8a;
        margin-bottom: 1rem;
    }
    
    /* Compact MultiSelect Tags - Readable Size */
    span[data-baseweb="tag"] {
        background-color: #1f2d53 !important; /* Navy Blue */
        color: white !important;
        font-size: 0.85rem !important;
        height: auto !important;
        padding: 4px 8px !important;
        margin: 2px !important;
    }
    span[data-baseweb="tag"] span {
        padding-right: 6px !important;
    }
    /* Let the container grow naturally */
    div[data-baseweb="select"] ul {
        padding: 0 !important;
        margin: 0 !important;
    }
    
    /* Dropdown Menu Items - Allow Wrap & Full Width */
    [data-baseweb="menu"] {
        max-width: 100% !important; 
        width: auto !important;
    }
    li[role="option"] {
        height: auto !important;
        min-height: 40px !important;
    }
    li[role="option"] > div {
        white-space: normal !important;
        overflow-wrap: break-word !important;
        line-height: 1.4 !important;
        padding-top: 8px !important;
        padding-bottom: 8px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Localization Logic ---
if "language" not in st.session_state:
    st.session_state.language = "日本語"

# Language Selector (Top of Sidebar)
st.sidebar.subheader("🌐 Language")
st.session_state.language = st.sidebar.radio(
    "Select Language",
    ["日本語", "English"],
    index=0 if st.session_state.language == "日本語" else 1,
    horizontal=True,
    label_visibility="collapsed"
)

# Translation Dictionary
translations = {
    "日本語": {
        "app_title": "BizStrategy",
        "title": "📊 売上分析ダッシュボード",
        "drop_data_info": "👋 分析を開始するには、データをアップロードしてください。",
        "file_uploader_label": "Excel または CSV ファイルをドロップ",
        "loading": "データを読み込み・解析中...",
        "success_load": "✅ データ読み込み完了！ダッシュボードを作成します。",
        "error_load": "読み込みエラー: ",
        "load_new_data": "🔄 新しいデータを読み込む",
        "column_settings": "🛠️ データ列の設定",
        "label_brand": "【ブランド】列を選択",
        "label_branch": "【支店 (Branch)】列を選択",
        "label_category": "【カテゴリ】列を選択",
        "label_year": "【年 (Year)】列を選択",
        "label_month": "【月 (Month)】列を選択",
        "label_value": "【売上・数値】列を選択",
        "label_item": "【商品名】列を選択 (任意)",
        "warn_col_same": "⚠️ ブランド列と売上列が同じです。正しい列を選択してください。",
        "filter_header": "🔍 フィルター",
        "filter_year": "対象年",
        "filter_month": "対象月",
        "filter_item": "対象商品",
        "filter_category": "カテゴリー",
        "select_branch": "支店を選択 (Branch)",
        "year_summary": "📊 年別売上サマリー",
        "year_sales_suffix": "年 売上",
        "no_data": "データがありません",
        "avg_transaction": "💳 平均取引単価",
        "no1_brand": "🏆 No.1 ブランド",
        "trend_chart": "📈 月次売上推移",
        "pie_chart": "🥧 売上構成比",
        "detail_table": "📋 詳細ピボットテーブル",
        "top_products": "🥇 売上TOP商品",
        "assign_columns_info": "👈 サイドバーで、分析に使用するデータ列を割り当ててください。",
        "current_preview": "現在のデータプレビュー:",
        "area_branch_sales": "エリア・支店別売上",
        "branch_sales_comparison": "支店別売上対比",
        "monthly_trend_header": "カテゴリー別　月次売上推移",
        "category_sales_section": "カテゴリー別売上",
        "treemap_title": "カテゴリー別・商品別 売上構成比",
        "treemap_chart_title": "売上構成比: カテゴリー -> 商品 (上位5)",
        "monthly_comparison_title": "月次売上推移 (対前年比)",
        "trend_chart_title": "月次売上推移",
        "growth_matrix_title": "成長マトリクス",
        "product_trend_title": "商品別 月次売上推移",
        "select_category_prompt": "表示するカテゴリーを選択してください（デフォルト: 売上上位5カテゴリー）",
        "select_category_label": "表示するカテゴリーを選択",
        "select_product_prompt": "表示する商品を選択してください（デフォルト: 売上上位5商品）",
        "select_product_label": "表示する商品を選択",
        "reset_button": "戻る (Reset)",
        "current_year_label": "今期",
        "prev_year_label": "前期",
        "label_company": "【会社 (Company)】列を選択",
        "warn_branch_not_set": "※ 支店列が未設定です",
        "item_popover_label": "✅ 選択メニュー",
        "item_bulk_action": "###### 📦 一括操作",
        "btn_select_all": "全選択",
        "btn_clear_all": "全解除",
        "item_search_header": "###### 🔍 検索して操作",
        "item_search_placeholder": "キーワード検索 (例: Frugra)",
        "item_search_count": "件ヒット (チェックで選択/解除)",
        "btn_apply": "OK (適用)",
        "warn_no_item": "⚠️ 商品が選択されていません",
        "info_no_item_col": "商品名列が未設定のため、商品フィルターは表示されません。",
        "company_branch_sales_header": "🏪 会社支店別売上 (Company Branch Sales)",
        "metric_total_sales": "🏢 全体売上 (Total)",
        "metric_branch_sales": "🏪 支店売上",
        "metric_forecast": "📈 着地見込 (Forecast)",
        "company_analysis_header": "🏢 全体売上 (Company Analysis)",
        "company_analysis_caption": "会社別の売上構成と推移",
        "company_sales_table_header": "会社別売上対比",
        "company_sales_chart_header": "会社別売上比較",
        "label_sales": "売上",
        "label_growth": "成長率 (%)",
        "label_total": "合計",
        "label_trend": "トレンド",
        "chain_customer_sales_header": "🏪 チェーン店・カスタマー別売上",
        "select_analysis_unit": "分析する単位を選択 (Select Analysis Level)",
        "unit_sales_comparison": "別売上比較",
        "warn_no_valid_col": "⚠️ 分析可能な列が見つかりません。",
        "customer_strategy_header": "👥 顧客・商品詳細分析 (Customer Strategy)",
        "customer_strategy_desc": "チェーンストアや個店（Customer）単位で、売上推移や商品動向を分析します。",
        "select_analysis_level_caption": "分析するレベル（列）を選んでください（例: Chain Store, Customer Name）",
        "select_target_col": "分析対象の列 (Customer Column)",
        "analysis_metric_label": "分析指標 (Analysis Metric)",
        "customer_search_label": "顧客名検索",
        "yoy_comparison_header": "年度別比較分析 (YoY Comparison)",
        "warn_insufficient_year": "比較可能な年度データが不足しています。",
        "category_breakdown_header": "Category Breakdown",
        "warn_no_cat_col": "カテゴリ列が設定されていません",
        "breakdown_header": "Breakdown",
        "warn_no_agg_col": "集計基準列が見つかりません",
        "cust_total_trend_title": "Total Trend",
        "product_opp_matrix_header": "💡 商品機会マトリクス",
        "matrix_no_data": "データ不足のためマトリクスを表示できません",
        "matrix_no_prev": "前年データがないため成長率を計算できません",
        "cat_trend_header": "📈 カテゴリ別 月次売上推移",
        "prod_trend_header": "📈 商品別 月次売上推移",
        "warn_no_cust_data": "No data for this customer.",
        "growth_driver_header": "🚀 商品・カテゴリー別 成長要因分析 (Growth Driver)",
        "growth_driver_desc": "指定した商品・カテゴリーの売上を「最も伸ばしたお店 (Top 10)」を分析します。",
        "analysis_axis_label": "分析切り口",
        "select_target_label": "分析対象を選択",
        "warn_axis_col_missing": "選択した切り口の列が見つかりません",
        "ranking_unit_label": "集計単位 (Ranking Unit)",
        "warn_agg_col_missing": "集計可能な列が見つかりません",
        "growth_top_100_header": "の売上を伸ばした",
        "scroll_caption": "※ リストをスクロールすると下位の店舗も確認できます。",
        "warn_year_data_missing": "比較対象の年度データが不足しています",
        "dist_analysis_header": "📊 Distribution & Coverage Analysis",
        "select_row_axis": "分析対象 (行 / Group By)",
        "select_unit_axis": "分布単位 (カウント対象 / Count Unique)",
        "select_trend_months": "📅 対象月",
        "popover_dist_label": "🔍 分析対象商品を選択",
        "dist_filter_main_header": "###### 📦 データ絞り込み",
        "dist_search_placeholder": "商品を検索 (例: Pizza)",
        "new_growth_analysis_header": "📈 新規獲得分析 (New Growth Analysis)",
        "new_growth_caption": "※ 直近には購入がなく、対象月に新たに購入があった先をリストアップします。",
        "new_growth_count_label": "🎯 新規獲得数",
        "info_no_new_growth": "✨ 新規獲得はありませんでした",
        "error_agg_type": "集計エラー: 列の型が対応していない可能性があります。",
        "warn_ai_api_needed": "⚠️ AI機能を利用するには、Google API Keyの設定が必要です。",
        "success_ai_connected": "✅ 接続成功！ AI機能が有効化されました。",
        "error_ai_connect": "接続エラー: ",
        "ai_chat_placeholder": "このデータについて質問する... (例: なぜ売上が伸び悩んでいる？来月の対策は？)",
        "cust_monthly_trend_title": "カテゴリー別 月次売上推移",
        "yoy_summary_table_title": "前年対比 実績集計表",
        "col_sales_suffix": "実績",
        "col_diff": "前年差異",
        "col_yoy": "前年比 (%)",
        "store_yoy_title": "店舗別 実績対比",
        "store_monthly_title": "店舗別 月次売上集計",
        "select_store_monthly": "詳細を確認する店舗を選択"
    },
    "English": {
        "app_title": "BizStrategy",
        "title": "📊 Sales Analytics Dashboard",
        "drop_data_info": "👋 Please upload data to start analysis.",
        "file_uploader_label": "Drop Excel or CSV file here",
        "loading": "Loading and parsing data...",
        "success_load": "✅ Data loaded! Creating dashboard.",
        "error_load": "Load Error: ",
        "load_new_data": "🔄 Load New Data",
        "column_settings": "🛠️ Column Settings",
        "label_brand": "Select [Brand] Column",
        "label_branch": "Select [Branch] Column",
        "label_category": "Select [Category] Column",
        "label_year": "Select [Year] Column",
        "label_month": "Select [Month] Column",
        "label_value": "Select [Value/Sales] Column",
        "label_item": "Select [Item] Column (Optional)",
        "warn_col_same": "⚠️ Brand and Value columns are the same.",
        "filter_header": "🔍 Filters",
        "filter_year": "Filter Years",
        "filter_month": "Filter Months",
        "filter_item": "Filter Items",
        "filter_category": "Filter Categories",
        "select_branch": "Select Branch (Sidebar)",
        "year_summary": "📊 Sales Summary by Year",
        "year_sales_suffix": " Sales",
        "no_data": "No Data Available",
        "avg_transaction": "💳 Avg Transaction",
        "no1_brand": "🏆 Top Brand",
        "trend_chart": "📈 Monthly Trend",
        "pie_chart": "🥧 Sales Composition",
        "detail_table": "📋 Detailed Pivot Table",
        "top_products": "🥇 Top Products",
        "assign_columns_info": "👈 Please map data columns in the sidebar.",
        "current_preview": "Current Data Preview:",
        "area_branch_sales": "Area/Branch Sales",
        "branch_sales_comparison": "Branch Sales Comparison",
        "monthly_trend_header": "Monthly Sales Trend by Category",
        "category_sales_section": "Category Sales",
        "treemap_title": "Sales Share by Category & Item",
        "treemap_chart_title": "Sales Share: Category -> Item (Top 5)",
        "monthly_comparison_title": "Monthly Sales Comparison (YoY)",
        "trend_chart_title": "Monthly Sales Trend",
        "growth_matrix_title": "Growth Matrix",
        "product_trend_title": "Monthly Sales Trend by Product",
        "select_category_prompt": "Select categories to display (Default: Top 5)",
        "select_category_label": "Select Categories",
        "select_product_prompt": "Select products to display (Default: Top 5)",
        "select_product_label": "Select Products",
        "reset_button": "Reset View",
        "current_year_label": "Current Year",
        "prev_year_label": "Previous Year",
        "label_company": "Select [Company] Column",
        "warn_branch_not_set": "※ Branch Column not set",
        "item_popover_label": "✅ Selection Menu",
        "item_bulk_action": "###### 📦 Bulk Actions",
        "btn_select_all": "Select All",
        "btn_clear_all": "Clear All",
        "item_search_header": "###### 🔍 Search & Select",
        "item_search_placeholder": "Search items...",
        "item_search_count": "items found",
        "btn_apply": "Apply Selection",
        "warn_no_item": "⚠️ No items selected.",
        "info_no_item_col": "Item filter hidden (column not mapped).",
        "company_branch_sales_header": "🏪 Sales Overview (Company/Branch)",
        "metric_total_sales": "🏢 Total Revenue",
        "metric_branch_sales": "🏪 Branch Revenue",
        "metric_forecast": "📈 Annual Forecast",
        "company_analysis_header": "🏢 Company Analysis",
        "company_analysis_caption": "Revenue composition and trends by company.",
        "company_sales_table_header": "Revenue Comparison by Company",
        "company_sales_chart_header": "Revenue Comparison Chart",
        "label_sales": "Sales",
        "label_growth": "Growth (%)",
        "label_total": "Total",
        "label_trend": "Trend",
        "chain_customer_sales_header": "🏪 Chain/Customer Sales Analysis",
        "select_analysis_unit": "Select Analysis Unit",
        "unit_sales_comparison": " Sales Comparison",
        "warn_no_valid_col": "⚠️ No valid analysis columns found.",
        "customer_strategy_header": "👥 Customer Strategy Analysis",
        "customer_strategy_desc": "Analyze sales trends and product mix by Chain Store or Customer.",
        "select_analysis_level_caption": "Select Analysis Level (e.g., Chain Store, Customer Name)",
        "select_target_col": "Analysis Column (Customer Column)",
        "analysis_metric_label": "Analysis Metric",
        "customer_search_label": "Search Customer",
        "yoy_comparison_header": "YoY Comparison Analysis",
        "warn_insufficient_year": "Insufficient year data for comparison.",
        "category_breakdown_header": "Category Breakdown",
        "warn_no_cat_col": "Category column not set.",
        "breakdown_header": "Breakdown",
        "warn_no_agg_col": "Aggregation column not found.",
        "cust_total_trend_title": "Total Trend",
        "product_opp_matrix_header": "💡 Product Opportunity Matrix",
        "matrix_no_data": "Insufficient data for matrix.",
        "matrix_no_prev": "Cannot calculate growth (No previous year data).",
        "cat_trend_header": "📈 Monthly Category Trends",
        "prod_trend_header": "📈 Monthly Product Trends",
        "warn_no_cust_data": "No data for this customer.",
        "growth_driver_header": "🚀 Growth Driver Analysis",
        "growth_driver_desc": "Analyze which customers drove the most growth for a specific Item/Category.",
        "analysis_axis_label": "Analysis Axis",
        "select_target_label": "Select Target",
        "warn_axis_col_missing": "Selected axis column not found.",
        "ranking_unit_label": "Ranking Unit",
        "warn_agg_col_missing": "Aggregation column not found.",
        "growth_top_100_header": "Top Growers for",
        "scroll_caption": "※ Scroll list to see more.",
        "warn_year_data_missing": "Insufficient year data for comparison.",
        "dist_analysis_header": "📊 Distribution & Coverage Analysis",
        "select_row_axis": "Row Axis (Group By)",
        "select_unit_axis": "Count Unit (Unique Count)",
        "select_trend_months": "📅 Target Months",
        "popover_dist_label": "🔍 Filter Analysis Items",
        "dist_filter_main_header": "###### 📦 Filter Data",
        "dist_search_placeholder": "Search Items (e.g. Pizza)",
        "new_growth_analysis_header": "📈 New Growth Analysis",
        "new_growth_caption": "※ Lists customers with no purchase in prior period but purchase in target month.",
        "new_growth_count_label": "🎯 New Acquisitions",
        "info_no_new_growth": "✨ No new acquisitions in this period.",
        "error_agg_type": "Aggregation Error: Column types may be incompatible.",
        "warn_ai_api_needed": "⚠️ Google API Key required for AI features.",
        "success_ai_connected": "✅ AI Features Connected!",
        "error_ai_connect": "Connection Error: ",
        "ai_chat_placeholder": "Ask about this data... (e.g. Why is sales flat? Action plan?)",
        "cust_monthly_trend_title": "Category Monthly Sales Trend",
        "yoy_summary_table_title": "YoY Performance Summary",
        "col_sales_suffix": " Sales",
        "col_diff": "Diff",
        "col_yoy": "YoY (%)",
        "store_yoy_title": "Store YoY Comparison",
        "store_monthly_title": "Store Monthly Breakdown",
        "select_store_monthly": "Select Store (Month Detail)"
    }
}

t = translations[st.session_state.language]

# Title Layout handled below
# st.markdown(f'<h1 class="main-title">{t["title"]}</h1>', unsafe_allow_html=True)

# --- Data Loading Logic ---
if "df" not in st.session_state:
    st.session_state.df = None

# Persistence Constants
CACHE_DIR = "data_cache"
CACHE_FILE = os.path.join(CACHE_DIR, "current_data.pkl")

# Ensure cache dir exists
if not os.path.exists(CACHE_DIR):
    os.makedirs(CACHE_DIR)

# Try to restore from cache if no data in session
if st.session_state.df is None and os.path.exists(CACHE_FILE):
    try:
        st.session_state.df = pd.read_pickle(CACHE_FILE)
        st.toast("🔄 以前のデータを復元しました", icon="📂")
    except Exception as e:
        st.error(f"キャッシュ読み込みエラー: {e}")

# If no data, show uploader directly on this page
if st.session_state.df is None:
    st.info(t["drop_data_info"])
    uploaded_file = st.file_uploader(t["file_uploader_label"], type=["xlsx", "csv"])
    
    if uploaded_file:
        try:
            with st.spinner(t["loading"]):
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                # Clean Data
                df = clean_data(df)
                st.session_state.df = df
                
                # Save to Cache
                df.to_pickle(CACHE_FILE)
                
                st.success(t["success_load"])
                # No need to rerun if we just set the state and continue, 
                # but if layout depends on it, ensure it's not looping.
                st.rerun()
        except Exception as e:
            st.error(f"{t['error_load']}{e}")
            st.stop()
    else:
        st.stop() # Stop execution here until file is uploaded

# Data exists
df = st.session_state.df.copy()

# Sidebar Data Reset
if st.sidebar.button(t["load_new_data"]):
    if st.session_state.df is not None:
        st.session_state.df = None
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
        st.rerun()

st.sidebar.divider()
st.sidebar.subheader(t["column_settings"])

# --- Column Mapping Logic ---
all_columns = list(df.columns)

# Config Persistence
CACHE_DIR = "data_cache" # Ensure consistency
CONFIG_FILE = os.path.join(CACHE_DIR, "column_config.json")
saved_config = {}
if os.path.exists(CONFIG_FILE):
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            saved_config = json.load(f)
    except:
        pass

def save_col_config():
    config = {
        "company": st.session_state.get("col_company"),
        "branch": st.session_state.get("col_branch"),
        "year": st.session_state.get("col_year"),
        "month": st.session_state.get("col_month"),
        "category": st.session_state.get("col_category"),
        "value": st.session_state.get("col_value"),
        "item": st.session_state.get("col_item")
    }
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False)

# --- Filter Persistence Logic ---
FILTER_CONFIG_FILE = os.path.join(CACHE_DIR, "filter_config.json")
saved_filters = {}
if os.path.exists(FILTER_CONFIG_FILE):
    try:
        with open(FILTER_CONFIG_FILE, 'r', encoding='utf-8') as f:
            saved_filters = json.load(f)
    except:
        pass

def save_filter_config():
    # Save current filter states from session_state
    filter_config = {
        "selected_companies": st.session_state.get("filter_company_key", []),
        "selected_branch": st.session_state.get("filter_branch_key", []),
        "selected_years": st.session_state.get("filter_year_key", []),
        "selected_months": st.session_state.get("selected_months_state", []),
        "selected_categories": st.session_state.get("filter_category_key", []),
        # Item filter might be large, but saving it is requested. 
        # Using the key defined downstream or standardized key.
        # Note: Item filter keys are dynamic based on item_col. We try to catch generic one or active one.
        # We'll handle specific item key saving in the callback wrapper if needed, or stick to generic keys.
        "selected_months": st.session_state.get("selected_months_state", []),
        "yoy_current_year": st.session_state.get("yoy_current_key"),
        "yoy_prev_year": st.session_state.get("yoy_prev_key"),
    }
    # For dynamic item key, we need to know the current item_col. 
    # But save_filter_config is called on change.
    # Let's rely on st.session_state directly in the dict construction below if practical.
    
    with open(FILTER_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(filter_config, f, ensure_ascii=False)

def get_filter_default(key, current_options, default_val=None):
    """
    Retrieve saved filter list. Validates that saved items still exist in current_options.
    If no save or no valid items, returns default_val (or None/Empty list).
    """
    saved_list = saved_filters.get(key)
    if saved_list is None:
        return default_val if default_val is not None else []
    
    # Validation: Only keep items that exist in current options
    # options might be list of strings or int.
    # Convert all to string for comparison safety if needed, 
    # but multiselect usually handles types matching.
    # Assuming options are the source of truth.
    valid_items = [i for i in saved_list if i in current_options]
    
    # If valid_items is empty but saved_list was not (meaning data changed), 
    # we default to Empty (All) to avoid errors.
    return valid_items if valid_items else (default_val if default_val is not None else [])

def get_filter_value(key, current_options, default_val):
    """
    Retrieve saved single value. Validates it exists in current_options.
    """
    saved_val = saved_filters.get(key)
    if saved_val is not None and saved_val in current_options:
        return saved_val
    return default_val

# Improved defaults (Auto-detect)
# Added filter to ignore 'Unnamed' columns to prevent bad defaults from artifacts
def is_valid_col(c):
    return 'Unnamed' not in str(c)

auto_branch = next((col for col in all_columns if is_valid_col(col) and ('branch' in str(col).lower() or '支店' in str(col) or 'jfc' in str(col).lower() or 'cust' in str(col).lower() or 'client' in str(col).lower() or '得意先' in str(col) or '顧客' in str(col))), None)
auto_year = next((col for col in all_columns if is_valid_col(col) and ('year' in str(col).lower() or '年' in str(col) or 'fy' in str(col).lower())), None)
auto_month = next((col for col in all_columns if is_valid_col(col) and ('month' in str(col).lower() or '月' in str(col))), None)
auto_category = next((col for col in all_columns if is_valid_col(col) and ('category' in str(col).lower() or 'categor' in str(col).lower() or 'カテゴリ' in str(col) or 'brand' in str(col).lower() or 'ブランド' in str(col) or 'class' in str(col).lower() or '分類' in str(col))), None)
auto_value = next((col for col in all_columns if is_valid_col(col) and ('value' in str(col).lower() or '売上' in str(col) or 'sales' in str(col).lower() or 'amount' in str(col).lower() or 'rev' in str(col).lower() or '金額' in str(col))), None)
auto_item = next((col for col in all_columns if is_valid_col(col) and ('item' in str(col).lower() or '商品' in str(col) or 'product' in str(col).lower() or 'sku' in str(col).lower() or 'desc' in str(col).lower())), None)

def get_default(saved_key, auto_default, options):
    val = saved_config.get(saved_key)
    # Validate: Check exists AND is a "valid" logical column (not Unnamed artifacts)
    if val in options and is_valid_col(val):
        return val
    return auto_default

def get_index(val, options):
    return options.index(val) if val in options else 0

# Company (New Request)
# User Request Order: Branch, Year, Month, Category, Value, Item
# Insert Company before Branch logic or adjacent? 
# Let's put it first as it is the highest level
opt_company = ["(なし)"] + all_columns
auto_company = next((col for col in all_columns if is_valid_col(col) and ('company' in str(col).lower() or '会社' in str(col) or 'corp' in str(col).lower() or '法人' in str(col))), None)
def_company = get_default("company", auto_company, opt_company)
idx_company = get_index(def_company, opt_company)
if def_company is None and auto_company:
     idx_company = get_index(auto_company, opt_company)

company_col = st.sidebar.selectbox(t["label_company"], opt_company, index=idx_company, key="col_company", on_change=save_col_config)

# Branch
opt_branch = ["(なし)"] + all_columns
def_branch = get_default("branch", auto_branch, opt_branch)
# Fallback index logic for branch (special case with (None))
idx_branch = get_index(def_branch, opt_branch)
if def_branch is None and auto_branch: # If no save, and auto found, use auto
     idx_branch = get_index(auto_branch, opt_branch)
elif def_branch is None:
     idx_branch = 0

branch_col = st.sidebar.selectbox(t["label_branch"], opt_branch, index=idx_branch, key="col_branch", on_change=save_col_config)

# Year
def_year = get_default("year", auto_year, all_columns)
year_col = st.sidebar.selectbox(t["label_year"], all_columns, index=get_index(def_year, all_columns), key="col_year", on_change=save_col_config)

# Month
def_month = get_default("month", auto_month, all_columns)
month_col = st.sidebar.selectbox(t["label_month"], all_columns, index=get_index(def_month, all_columns), key="col_month", on_change=save_col_config)

# Category
def_category = get_default("category", auto_category, all_columns)
category_col = st.sidebar.selectbox(t["label_category"], all_columns, index=get_index(def_category, all_columns), key="col_category", on_change=save_col_config)
# Implicitly map Brand to Category for internal logic
brand_col = category_col

# Value
def_value = get_default("value", auto_value, all_columns)
val_col_name = st.sidebar.selectbox(t["label_value"], all_columns, index=get_index(def_value, all_columns), key="col_value", on_change=save_col_config)

# Item
opt_item = ["(なし)"] + all_columns
def_item = get_default("item", auto_item, opt_item)
idx_item = get_index(def_item, opt_item)
if def_item is None and auto_item:
    idx_item = get_index(auto_item, opt_item)

item_col = st.sidebar.selectbox(t["label_item"], opt_item, index=idx_item, key="col_item", on_change=save_col_config)

if category_col == val_col_name:
    st.warning(t["warn_col_same"])

# Proceed with Dashboard Generation using selected columns
if brand_col and year_col and month_col and val_col_name:
    
    # --- Data Preprocessing ---
    # Robust Normalization of Year and Month to prevent mismatch (e.g. "2024" vs "2024.0" vs " 2024 ")
    if year_col and year_col != "(なし)":
        df[year_col] = df[year_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    
    if month_col and month_col != "(なし)":
        df[month_col] = df[month_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    df[val_col_name] = pd.to_numeric(df[val_col_name], errors='coerce').fillna(0)
    
    # Month Ordering (Robust)
    month_map = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        'january': 1, 'february': 2, 'march': 3, 'april': 4, 'june': 6,
        'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12
    }
    
    if df[month_col].dtype == 'object':
        # Normalize: Lower case and strip
        # Improved Logic: Try to extract leading number first (e.g. "01_May" -> 1)
        # If no leading number, fallback to name map
        
        # 1. Try Map First (for pure names)
        df['MonthNum'] = df[month_col].astype(str).str.lower().str.strip().str.replace(r'[.,]', '', regex=True).map(month_map)
        
        # 2. Fill NaN with regex extraction
        mask_nan = df['MonthNum'].isna()
        if mask_nan.any():
            extracted_num = df.loc[mask_nan, month_col].astype(str).str.extract(r'^(\d+)', expand=False)
            df.loc[mask_nan, 'MonthNum'] = pd.to_numeric(extracted_num, errors='coerce')
        
        # Final cleanup
        df['MonthNum'] = df['MonthNum'].fillna(0).astype(int)
        
        # If mapping failed completely, try numeric conversion
        if df['MonthNum'].sum() == 0:
             df['MonthNum'] = pd.to_numeric(df[month_col], errors='coerce').fillna(0).astype(int)
    else:
        df['MonthNum'] = pd.to_numeric(df[month_col], errors='coerce').fillna(0).astype(int)

    # Filters Sidebar
    st.sidebar.divider()
    st.sidebar.subheader(t["filter_header"])
    
    unique_years = sorted(df[year_col].unique())
    unique_months_df = df[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')
    unique_months = unique_months_df[month_col].tolist()
    
    # Sidebar filters
    # Sidebar filters
    st.sidebar.divider()
    st.sidebar.subheader("🔍 Filters")
    
    # --- Company Filter (Highest Level) ---
    if company_col and company_col != "(なし)":
        unique_companies = sorted(df[company_col].astype(str).unique())
        
        # Helper buttons
        c_com_all, c_com_clr = st.sidebar.columns(2)
        if c_com_all.button("Select All", key="btn_company_all"):
            st.session_state["filter_company_key"] = unique_companies
            save_filter_config()
            st.rerun()
        if c_com_clr.button("Clear", key="btn_company_clr"):
             st.session_state["filter_company_key"] = []
             save_filter_config()
             st.rerun()

        selected_companies = st.sidebar.multiselect(
            t["label_company"],
            options=unique_companies,
            default=get_filter_default("selected_companies", unique_companies, []),
            placeholder="Selected: All (Default)",
            key="filter_company_key",
            on_change=save_filter_config
        )
        if not selected_companies:
            st.sidebar.markdown("###### ✅ All")
    else:
        selected_companies = "All"

    # --- Branch Filter (Moved here to be sticky) ---
    if branch_col and branch_col != "(なし)":
        unique_branches = sorted(df[branch_col].astype(str).unique())
        
        # Helper: Select All / Clear Buttons
        c_sel_all, c_sel_clr = st.sidebar.columns(2)
        if c_sel_all.button("Select All", key="btn_branch_all"):
            st.session_state["filter_branch_key"] = unique_branches
            save_filter_config()
            st.rerun()
        if c_sel_clr.button("Clear", key="btn_branch_clr"):
            st.session_state["filter_branch_key"] = []
            save_filter_config()
            st.rerun()
            
        selected_branch = st.sidebar.multiselect(
            t["select_branch"], 
            options=unique_branches, 
            default=get_filter_default("selected_branch", unique_branches, []), 
            placeholder="Selected: All (Default)",
            key="filter_branch_key",
            on_change=save_filter_config
        )
        if not selected_branch:
             st.sidebar.markdown("###### ✅ All")
    else:
        selected_branch = "All"
        if branch_col == "(なし)":
             st.sidebar.caption(t["warn_branch_not_set"])



    selected_years = st.sidebar.multiselect(
        t["filter_year"], 
        options=unique_years, 
        default=get_filter_default("selected_years", unique_years, unique_years), 
        key="filter_year_key",
        on_change=save_filter_config
    )

    # --- Month Filter Enhanced UI ---
    # User request: "Make it same format as Branch. Delete selection menu."
    selected_months = st.sidebar.multiselect(
        t["filter_month"], 
        options=unique_months, 
        default=get_filter_default("selected_months", unique_months, []), 
        placeholder="Selected: All (Default)",
        key='selected_months_state',
        on_change=save_filter_config,
        label_visibility="collapsed"
    )
    if not selected_months:
         st.sidebar.markdown("###### ✅ All")
    
    # --- Category Filter ---
    # Moved above Item Filter as requested
    if category_col and category_col != "(なし)":
        unique_categories = sorted(df[category_col].astype(str).unique())
        
        # Helper: Select All / Clear Buttons
        c_cat_all, c_cat_clr = st.sidebar.columns(2)
        if c_cat_all.button("Select All", key="btn_category_all"):
             st.session_state["filter_category_key"] = unique_categories
             save_filter_config()
             st.rerun()
        if c_cat_clr.button("Clear", key="btn_category_clr"):
             st.session_state["filter_category_key"] = []
             save_filter_config()
             st.rerun()

        selected_categories = st.sidebar.multiselect(
            t["filter_category"], 
            options=unique_categories, 
            default=get_filter_default("selected_categories", unique_categories, unique_categories), 
            key="filter_category_key",
            on_change=save_filter_config
        )
        if not selected_categories:
             st.sidebar.markdown("###### ✅ All")
    else:
        selected_categories = "All"

    selected_items = "All"
    if item_col and item_col != "(なし)":
        unique_items = sorted(df[item_col].astype(str).unique())
        
        # Initialize Selection DataFrame in Session State
        # Re-initialize if unique_items count changed (e.g. data reload)
        # Note: Ideally we track data hash, but simply checking if all items exist is good enough or reset on new unique_items size
        # Session State for Item Selection
        if 'selected_items_state' not in st.session_state:
            st.session_state['selected_items_state'] = unique_items # Default All


        # Helper functions for buttons
        def select_all_items():
            st.session_state['selected_items_state'] = unique_items
        
        def clear_all_items():
            st.session_state['selected_items_state'] = []

        # Filter UI Layout
        st.sidebar.markdown(f"**{t['filter_item']}**")
        
        # Popover for Bulk Actions & Search
        with st.sidebar.popover(t["item_popover_label"], help="一括選択・検索メニュー", use_container_width=True):
            st.markdown(t["item_bulk_action"])
            c_p1, c_p2 = st.columns(2)
            if c_p1.button(t["btn_select_all"], key="btn_all_sel", use_container_width=True):
                st.session_state['selected_items_state'] = unique_items
                if 'search_editor' in st.session_state: del st.session_state['search_editor']
                st.rerun()
            if c_p2.button(t["btn_clear_all"], key="btn_all_clr", use_container_width=True):
                st.session_state['selected_items_state'] = []
                if 'search_editor' in st.session_state: del st.session_state['search_editor']
                st.rerun()

            st.divider()
            
            st.divider()
            
            st.markdown(t["item_search_header"])
            search_word = st.text_input("Search", placeholder=t["item_search_placeholder"], label_visibility="collapsed")
            
            # Prepare DataFrame for Editor
            current_selection = set(st.session_state.get('selected_items_state', []))
            
            df_search = pd.DataFrame({
                'Selected': [item in current_selection for item in unique_items],
                'Item': unique_items
            })

            # Filter by search word
            if search_word:
                df_search = df_search[df_search['Item'].astype(str).str.contains(search_word, case=False, na=False)]
            
            # Show count
            st.caption(f"{len(df_search)} {t['item_search_count']}")

            # Data Editor for Granular Selection
            edited_search_df = st.data_editor(
                df_search,
                column_config={
                    "Selected": st.column_config.CheckboxColumn(
                        "Sel.",
                        width="small",
                        default=False 
                    ),
                    "Item": st.column_config.TextColumn(
                        "Item Name",
                        width=400, # Fixed width to see content
                        disabled=True
                    )
                },
                disabled=["Item"],
                hide_index=True,
                use_container_width=True,
                height=min((len(df_search)+1)*35 + 10, 400), # Dynamic height
                key="search_editor"
            )

            # OK Button to Apply Changes
            if st.button(t["btn_apply"], key="btn_apply_changes", use_container_width=True, type="primary"):
                # Calculate new selection based on the FINAL state of the editor when button is clicked
                new_selection_set = current_selection.copy()
                view_items = df_search['Item'].tolist()
                new_state_map = dict(zip(edited_search_df['Item'], edited_search_df['Selected']))
                
                for item in view_items:
                    if new_state_map.get(item, False):
                        new_selection_set.add(item)
                    else:
                        new_selection_set.discard(item) # Remove if exists
                
                # Update State
                st.session_state['selected_items_state'] = sorted(list(new_selection_set))
                # Note: We technically don't need to delete 'search_editor' here because the inputs naturally match now,
                # but clearing it ensures clean slate.
                if 'search_editor' in st.session_state: del st.session_state['search_editor']
                st.rerun()

        # Standard Multiselect (Matches other filters)
        # Label is handled above, so use empty label here to avoid double label, but keep space tight
        # Update key to be dynamic based on item_col to force reset when column changes
        selected_items = st.sidebar.multiselect(
            t["filter_item"], 
            options=unique_items, 
            default=unique_items, # Select ALL by default when key changes
            key=f'selected_items_state_{item_col}',
            label_visibility="collapsed"
        )

        if not selected_items:
            st.sidebar.warning(t["warn_no_item"])

    elif item_col == "(なし)":
         st.sidebar.info(t["info_no_item_col"])
         selected_items = "All"

    # --- Top Title ---
    st.markdown(f'<h1 class="main-title" style="margin-top:0;">{t["title"]}</h1>', unsafe_allow_html=True)

    # Apply Filters for MAIN DASHBOARD (Charts etc)
    # Start with Year and Month
    # Apply Filters for MAIN DASHBOARD (Charts etc)
    # Start with Year and Month
    # Robustness: Ensure we compare Strings vs Strings to handle Text/Int mismatches
    
    df_filtered = df.copy()
    
    # Filter by Company
    if selected_companies and selected_companies != "All" and company_col and company_col != "(なし)":
         df_filtered = df_filtered[df_filtered[company_col].astype(str).isin(selected_companies)]

    # Filter by Category
    if selected_categories and selected_categories != "All" and category_col and category_col != "(なし)":
         df_filtered = df_filtered[df_filtered[category_col].astype(str).isin(selected_categories)]

    # Filter by Year
    if selected_years:
        df_filtered = df_filtered[df_filtered[year_col].astype(str).isin([str(y) for y in selected_years])]
    
    # Filter by Month (Empty = All)
    if selected_months:
        df_filtered = df_filtered[df_filtered[month_col].astype(str).isin([str(m) for m in selected_months])]
    
    # Apply Item Filter if active
    if selected_items != "All" and item_col and item_col != "(なし)":
        df_filtered = df_filtered[df_filtered[item_col].astype(str).isin(selected_items)]

    # Capture Pre-Branch Data for Benchmark Table
    df_pre_branch_filter = df_filtered.copy()

    # Apply Branch Filter to main dashboard
    if selected_branch and branch_col and branch_col != "(なし)" and selected_branch != "All":
        df_filtered = df_filtered[df_filtered[branch_col].astype(str).isin(selected_branch)]
    
    # --- [NEW] Export Feature (Sidebar) ---
    st.sidebar.subheader("📥 Export Report")
    if st.sidebar.button("Generate Excel Report", key="btn_export_excel"):
        with st.spinner("Generating Sales Review Report..."):
            try:
                output = io.BytesIO()
                # Use xlsxwriter
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    
                    # --- Styles (Shared) ---
                    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': 'black', 'align': 'left'})
                    fmt_section_header = workbook.add_format({'bold': True, 'bg_color': 'black', 'font_color': 'white', 'align': 'left', 'border': 1})
                    fmt_col_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                    fmt_text = workbook.add_format({'border': 1})
                    
                    # Data Formats
                    fmt_curr = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bold': True, 'font_color': '#0070C0'}) 
                    fmt_prev = workbook.add_format({'num_format': '#,##0', 'border': 1, 'font_color': '#984807'})
                    fmt_plug = workbook.add_format({'num_format': '#,##0', 'border': 1, 'italic': True})
                    fmt_ratio = workbook.add_format({'num_format': '0.0%', 'border': 1})
                    
                    # Red conditional
                    fmt_pct_red = workbook.add_format({'num_format': '0.0%', 'border': 1, 'font_color': 'red', 'bold': True})
                    fmt_pct_green = workbook.add_format({'num_format': '0.0%', 'border': 1, 'font_color': 'black'})
                    fmt_plug_red = workbook.add_format({'num_format': '#,##0', 'border': 1, 'font_color': 'red'})
                    
                    # Highlight for Selected Branch
                    fmt_highlight_text = workbook.add_format({'border': 1, 'bg_color': '#FFFF00', 'bold': True})
                    fmt_highlight_prev = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00', 'font_color': '#984807'})
                    fmt_highlight_curr = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00', 'font_color': '#0070C0', 'bold': True})
                    fmt_highlight_pct_red = workbook.add_format({'num_format': '0.0%', 'border': 1, 'bg_color': '#FFFF00', 'font_color': 'red', 'bold': True})
                    fmt_highlight_pct_green = workbook.add_format({'num_format': '0.0%', 'border': 1, 'bg_color': '#FFFF00', 'font_color': 'black'})
                    fmt_highlight_plug = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00', 'italic': True})
                    fmt_highlight_plug_red = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFFF00', 'font_color': 'red'})
                    
                    # --- Data Setup ---
                    avail_years_exp = sorted(df_pre_branch_filter[year_col].unique(), key=lambda x: str(x))
                    if len(avail_years_exp) >= 2:
                         curr_y = avail_years_exp[-1]
                         prev_y = avail_years_exp[-2]
                    elif len(avail_years_exp) == 1:
                         curr_y = avail_years_exp[0]
                         prev_y = None
                    else:
                         curr_y, prev_y = None, None
                    
                    # Aggregation Helper
                    def get_metrics_df(df_target, grp_col):
                         if not grp_col or df_target.empty: return pd.DataFrame()
                         agg = df_target.groupby([grp_col, year_col])[val_col_name].sum().unstack(fill_value=0)
                         if curr_y and curr_y not in agg.columns: agg[curr_y] = 0
                         if prev_y and prev_y not in agg.columns: agg[prev_y] = 0
                         
                         if curr_y and prev_y:
                             agg = agg[[prev_y, curr_y]]
                             agg['vs LY (%)'] = agg.apply(lambda r: (r[curr_y]/r[prev_y] - 1) if r[prev_y]!=0 else 0, axis=1)
                             agg['Plug'] = agg[curr_y] - agg[prev_y]
                             total_curr = agg[curr_y].sum()
                             agg['Ratio'] = agg.apply(lambda r: r[curr_y]/total_curr if total_curr!=0 else 0, axis=1)
                             agg = agg.sort_values(curr_y, ascending=False)
                         return agg

                    # --- Sheet Generator Function ---
                    def write_review_sheet(sheet_name, breakdown_col, breakdown_label):
                        ws = workbook.add_worksheet(sheet_name)
                        row = 0
                        
                        # Title
                        sel_br_name = selected_branch[0] if isinstance(selected_branch, list) and selected_branch else selected_branch
                        if sel_br_name == "All" or isinstance(sel_br_name, list): sel_br_name = "All Branch"
                        
                        ws.write(row, 0, f"【{t.get('app_title', 'BizStrategy')}】Sales Review (by {sel_br_name})", fmt_title)
                        row += 1
                        
                        # Global Summary
                        all_curr = df_pre_branch_filter[df_pre_branch_filter[year_col]==curr_y][val_col_name].sum() if curr_y else 0
                        all_prev = df_pre_branch_filter[df_pre_branch_filter[year_col]==prev_y][val_col_name].sum() if prev_y else 0
                        all_growth = (all_curr/all_prev - 1) if all_prev!=0 else 0
                        all_plug = all_curr - all_prev
                        
                        headers = [breakdown_label, f"FY'{str(prev_y)[-2:]} Sales", f"FY'{str(curr_y)[-2:]} Sales", "vs LY (%)", "Plug", "Ratio"]
                        
                        # Fix: Add Header Row for Global Summary
                        ws.write(row, 0, "", fmt_section_header) # Placeholder or empty
                        ws.write_row(row, 1, headers[1:], fmt_col_header)
                        row += 1
                        
                        # Determine Company Label
                        comp_label = "Wismettac" # Default
                        if company_col and selected_companies and selected_companies != "All":
                             if isinstance(selected_companies, list):
                                 if len(selected_companies) == 1:
                                     comp_label = str(selected_companies[0])
                                 else:
                                     comp_label = f"{len(selected_companies)} Companies"
                             else:
                                 comp_label = str(selected_companies)
                                 
                        ws.write(row, 0, f"{comp_label} All Branch", fmt_section_header)
                        ws.write(row, 1, all_prev, fmt_prev)
                        ws.write(row, 2, all_curr, fmt_curr)
                        ws.write(row, 3, all_growth/100, fmt_pct_red if all_growth<0 else fmt_pct_green)
                        ws.write(row, 4, all_plug, fmt_plug_red if all_plug<0 else fmt_plug)
                        ws.write(row, 5, 1.0, fmt_ratio)
                        row += 2
                        
                        # Section 1: All Branch List (Always Branch)
                        ws.merge_range(row, 0, row, 5, "■ Actual Sales (Branch Comparison)", fmt_section_header)
                        row += 1
                        ws.write_row(row, 0, ["Branch"] + headers[1:], fmt_col_header)
                        row += 1
                        
                        if branch_col:
                            df_all_br = get_metrics_df(df_pre_branch_filter, branch_col)
                            start_row = row
                            
                            for idx, r in df_all_br.iterrows():
                                 # Highlight Check
                                 is_selected = False
                                 if isinstance(selected_branch, list):
                                     is_selected = str(idx) in [str(s) for s in selected_branch]
                                 elif selected_branch != "All":
                                     is_selected = str(idx) == str(selected_branch)
                                 
                                 f_txt = fmt_highlight_text if is_selected else fmt_text
                                 f_pre = fmt_highlight_prev if is_selected else fmt_prev
                                 f_cur = fmt_highlight_curr if is_selected else fmt_curr
                                 f_pct = (fmt_highlight_pct_red if r['vs LY (%)'] < 0 else fmt_highlight_pct_green) if is_selected else (fmt_pct_red if r['vs LY (%)'] < 0 else fmt_pct_green)
                                 f_plg = (fmt_highlight_plug_red if r['Plug'] < 0 else fmt_highlight_plug) if is_selected else (fmt_plug_red if r['Plug'] < 0 else fmt_plug)
                                 f_rat = fmt_highlight_pct_green if is_selected else fmt_ratio
                                 
                                 ws.write(row, 0, str(idx), f_txt)
                                 ws.write(row, 1, r[prev_y], f_pre)
                                 ws.write(row, 2, r[curr_y], f_cur)
                                 ws.write(row, 3, r['vs LY (%)']/100, f_pct)
                                 ws.write(row, 4, r['Plug'], f_plg)
                                 ws.write(row, 5, r['Ratio'], f_rat)
                                 row += 1
                            
                            ws.conditional_format(start_row, 1, row-1, 1, {'type': 'data_bar', 'bar_color': '#FFC000'})
                            ws.conditional_format(start_row, 2, row-1, 2, {'type': 'data_bar', 'bar_color': '#00B0F0'})
                            
                            # Total Row
                            ws.write(row, 0, "Total", fmt_section_header)
                            ws.write(row, 1, df_all_br[prev_y].sum(), fmt_prev)
                            ws.write(row, 2, df_all_br[curr_y].sum(), fmt_curr)
                            # Safe Div check for growth?
                            br_tot_prev = df_all_br[prev_y].sum()
                            ws.write(row, 3, (df_all_br[curr_y].sum()/br_tot_prev-1) if br_tot_prev else 0, fmt_pct_green)
                            ws.write(row, 4, df_all_br['Plug'].sum(), fmt_plug)
                            ws.write(row, 5, 1.0, fmt_ratio)
                            row += 2
    
                        # Section 2: Global Sales (Breakdown)
                        ws.merge_range(row, 0, row, 5, f"■ Actual Sales by {breakdown_label} (By All Branch)", fmt_section_header)
                        row += 1
                        ws.write_row(row, 0, [breakdown_label] + headers[1:], fmt_col_header)
                        row += 1
                        
                        if breakdown_col:
                            df_all_item = get_metrics_df(df_pre_branch_filter, breakdown_col).head(50)
                            start_row = row
                            for idx, r in df_all_item.iterrows():
                                 ws.write(row, 0, str(idx), fmt_text)
                                 ws.write(row, 1, r[prev_y], fmt_prev)
                                 ws.write(row, 2, r[curr_y], fmt_curr)
                                 ws.write(row, 3, r['vs LY (%)']/100, fmt_pct_red if r['vs LY (%)'] < 0 else fmt_pct_green)
                                 ws.write(row, 4, r['Plug'], fmt_plug_red if r['Plug'] < 0 else fmt_plug)
                                 ws.write(row, 5, r['Ratio'], fmt_ratio)
                                 row += 1
                            if row > start_row:
                                ws.conditional_format(start_row, 1, row-1, 1, {'type': 'data_bar', 'bar_color': '#FFC000'})
                                ws.conditional_format(start_row, 2, row-1, 2, {'type': 'data_bar', 'bar_color': '#00B0F0'})
                            
                            ws.write(row, 0, "Total (Top 50)", fmt_section_header)
                            ws.write(row, 1, df_all_item[prev_y].sum(), fmt_prev)
                            ws.write(row, 2, df_all_item[curr_y].sum(), fmt_curr)
                            ws.write(row, 4, df_all_item['Plug'].sum(), fmt_plug)
                            row += 2
                            
                        # Section 3: Local Sales (Breakdown)
                        if selected_branch != "All":
                            ws.merge_range(row, 0, row, 5, f"■ Actual Sales by {breakdown_label} (By {sel_br_name})", fmt_section_header)
                            row += 1
                            ws.write_row(row, 0, [breakdown_label] + headers[1:], fmt_col_header)
                            row += 1
                            
                            if breakdown_col:
                                df_loc_item = get_metrics_df(df_filtered, breakdown_col).head(50)
                                start_row = row
                                for idx, r in df_loc_item.iterrows():
                                     ws.write(row, 0, str(idx), fmt_text)
                                     ws.write(row, 1, r[prev_y], fmt_prev)
                                     ws.write(row, 2, r[curr_y], fmt_curr)
                                     ws.write(row, 3, r['vs LY (%)']/100, fmt_pct_red if r['vs LY (%)'] < 0 else fmt_pct_green)
                                     ws.write(row, 4, r['Plug'], fmt_plug_red if r['Plug'] < 0 else fmt_plug)
                                     ws.write(row, 5, r['Ratio'], fmt_ratio)
                                     row += 1
                                if row > start_row:
                                    ws.conditional_format(start_row, 1, row-1, 1, {'type': 'data_bar', 'bar_color': '#FFC000'})
                                    ws.conditional_format(start_row, 2, row-1, 2, {'type': 'data_bar', 'bar_color': '#00B0F0'})
                                
                                ws.write(row, 0, "Total (Top 50)", fmt_section_header)
                                ws.write(row, 1, df_loc_item[prev_y].sum(), fmt_prev)
                                ws.write(row, 2, df_loc_item[curr_y].sum(), fmt_curr)
                                ws.write(row, 4, df_loc_item['Plug'].sum(), fmt_plug)
    
                        ws.set_column('A:A', 35)
                        ws.set_column('B:F', 14)

                    # --- Generate Sheets ---
                    # 1. Item Sheet
                    item_bk_col = item_col if (item_col and item_col != "(なし)") else category_col
                    write_review_sheet("Sales Review (Item)", item_bk_col, "Item")
                    
                    # 2. Category Sheet (New)
                    if category_col and category_col != "(なし)":
                        write_review_sheet("Sales Review (Class)", category_col, "Category")

                output.seek(0)
                st.sidebar.download_button(
                    label="📥 Download Sales Review",
                    data=output,
                    file_name="Sales_Review_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.sidebar.success("Review Generated!")
            except Exception as e:
                import traceback
                st.sidebar.error(f"Export Error: {e}")
                st.sidebar.text(traceback.format_exc())
    
    # --- Top KPIs by Year ---
    

    # --- Top KPIs by Year (Redesigned: Total vs Branch YoY) ---
    st.divider()
    st.subheader(t["company_branch_sales_header"])

    # Determine Current/Previous Year for KPI Context
    # Determine Current/Previous Year for KPI Context
    if selected_years:
        numeric_years = []
        for y in selected_years:
            try:
                # Handle 2025, "2025", 2025.0, "2025.0", "FY2025"
                s_val = str(y)
                # Remove FY prefix if present
                s_val = s_val.upper().replace("FISCAL YEAR", "").replace("FY", "").strip()
                val = float(s_val)
                numeric_years.append(int(val))
            except (ValueError, TypeError):
                continue
                
        if numeric_years:
            current_year = max(numeric_years)
            prev_year = current_year - 1
        else:
            current_year = None
            prev_year = None
    else:
        current_year = None
        prev_year = None

    if current_year is not None:
        # Helper to get sales for a specific year from a DF
        def get_sales_for_year(df_in, target_year):
            # Group by year string normalization
            # Note: year_col in df is already normalized to string in Preprocessing step
            # We compare with str(target_year). Robustly handle "FY" in data if not cleaned.
            
            # 1. Try exact numbers match (astype str)
            mask = df_in[year_col].astype(str).str.replace(r'\.0$', '', regex=True) == str(target_year)
            
            if not mask.any():
                # 2. Try removing "FY" prefix from data column
                cleaned_col = df_in[year_col].astype(str).str.replace(r'(?i)^\s*(?:fy|fiscal\s*year)\s*', '', regex=True)
                cleaned_col = cleaned_col.str.replace(r'\.0$', '', regex=True).str.strip()
                mask = cleaned_col == str(target_year)
                
            row = df_in[mask][val_col_name].sum()
            return row

        # 1. Total Sales (All Branches)
        total_sales_curr = get_sales_for_year(df_pre_branch_filter, current_year)
        total_sales_prev = get_sales_for_year(df_pre_branch_filter, prev_year)
        
        total_delta = None
        if total_sales_prev > 0:
            total_delta = f"{(total_sales_curr - total_sales_prev) / total_sales_prev:.1%}"
            
        # 2. Branch Sales (Selected Branch)
        # Only meaningful if a branch filter is applied. 
        # If "All" is selected, df_filtered == df_pre_branch_filter.
        branch_sales_curr = get_sales_for_year(df_filtered, current_year)
        branch_sales_prev = get_sales_for_year(df_filtered, prev_year)
        
        branch_delta = None
        if branch_sales_prev > 0:
            branch_delta = f"{(branch_sales_curr - branch_sales_prev) / branch_sales_prev:.1%}"
            
        # Branch Label
        if not selected_branch or selected_branch == "All":
             disp_branch_label = "All Branches"
        else:
             if isinstance(selected_branch, list):
                 if len(selected_branch) == 1:
                     disp_branch_label = selected_branch[0]
                 else:
                     disp_branch_label = f"{len(selected_branch)} Branches Selected"
             else:
                 disp_branch_label = str(selected_branch)

        # 3. Forecast Logic
        # Calculate Forecast details based on CURRENT view (Filtered)
        # Forecast is usually relevant for the current selection (Branch)
        forecast_text = "N/A"
        forecast_val = 0
        months_count = df_filtered[df_filtered[year_col].astype(str) == str(current_year)][month_col].nunique()
        if months_count > 0 and branch_sales_curr > 0:
             forecast_val = (branch_sales_curr / months_count) * 12
             forecast_text = f"{forecast_val:,.0f}"

        # Display Scorecards
        kpi_cols = st.columns(3)
        
        with kpi_cols[0]:
            st.metric(f"{t['metric_total_sales']} - {current_year}", f"{total_sales_curr:,.0f}", delta=total_delta)
            
        with kpi_cols[1]:
            st.metric(f"{t['metric_branch_sales']} ({disp_branch_label}) - {current_year}", f"{branch_sales_curr:,.0f}", delta=branch_delta)
            
        with kpi_cols[2]:
             if forecast_val > 0:
                 st.metric(label=f"{t['metric_forecast']} {current_year}", value=forecast_text, delta=f"Based on {months_count} months")

    else:
        st.warning("Please select a valid Year to see KPIs.")


    # --- [NEW] Company Structure Analysis (Requested Position: Below Scorecards) ---
    if company_col and company_col != "(なし)":
        st.divider()
        
        # 1. Determine Company Label for Header
        if not selected_companies or selected_companies == "All":
             disp_comp_label = "All Companies"
        else:
             if isinstance(selected_companies, list):
                 if len(selected_companies) == 1:
                     disp_comp_label = selected_companies[0]
                 else:
                     disp_comp_label = f"{len(selected_companies)} Companies"
             else:
                 disp_comp_label = str(selected_companies)

        st.subheader(f"{t['company_analysis_header']} - {disp_comp_label}")
        st.caption(t["company_analysis_caption"])
        
        # Prepare Data for Company Analysis
        # We need a base DF that respects Month/Item filters but NOT Global Year filters (so we can choose years)
        df_comp_base = df.copy()
        if selected_months:
            df_comp_base = df_comp_base[df_comp_base[month_col].astype(str).isin([str(m) for m in selected_months])]
        if selected_items != "All" and item_col and item_col != "(なし)":
            df_comp_base = df_comp_base[df_comp_base[item_col].astype(str).isin(selected_items)]
        # Also respect Company Filter if set (which it is, by df_filtered logic usually, but here we want to allow comparison?)
        # User said "Select Company name written on side". If specific company selected, we show that.
        if selected_companies and selected_companies != "All":
             df_comp_base = df_comp_base[df_comp_base[company_col].astype(str).isin(selected_companies)]
        
        # [Sync Request] Respect Sidebar Year Filter
        if selected_years:
             df_comp_base = df_comp_base[df_comp_base[year_col].astype(str).isin([str(y) for y in selected_years])]

        # Year Selection for Company Analysis
        avail_years_comp = sorted(df_comp_base[year_col].unique(), reverse=True)
        
        if not avail_years_comp:
             st.warning("No data for Company Analysis.")
        else:
             # Default Logic
             def_c_comp = current_year if (current_year and current_year in avail_years_comp) else avail_years_comp[0]
             def_p_comp = prev_year if (prev_year and prev_year in avail_years_comp) else (avail_years_comp[1] if len(avail_years_comp) > 1 else avail_years_comp[0])

             cy_c, py_c, sp_c = st.columns([1, 1, 4])
             with cy_c:
                 sel_cy_comp = st.selectbox("今期 (Current)", options=avail_years_comp, index=avail_years_comp.index(def_c_comp), key="comp_cy")
             with py_c:
                 sel_py_comp = st.selectbox("前期 (Previous)", options=avail_years_comp, index=avail_years_comp.index(def_p_comp), key="comp_py")

             # Filter for Pivot
             df_comp_viz = df_comp_base[df_comp_base[year_col].isin([sel_cy_comp, sel_py_comp])]
             
             # Pivot Table: Company vs Year
             comp_pivot = df_comp_viz.groupby([company_col, year_col])[val_col_name].sum().unstack(fill_value=0)

             # Shared Styling Helper
             def style_growth_simple(v):
                 try:
                    val = float(v.strip('%')) if isinstance(v, str) else float(v)
                    return 'color: green; font-weight: bold;' if val >= 0 else 'color: red; font-weight: bold;'
                 except:
                    return ''
             
             # Calculate Stats
             curr = sel_cy_comp
             prev = sel_py_comp
             
             # Ensure Cols
             if curr not in comp_pivot.columns: comp_pivot[curr] = 0
             if prev not in comp_pivot.columns: comp_pivot[prev] = 0
             
             comp_pivot['vs LY (%)'] = comp_pivot.apply(lambda r: ((r[curr] - r[prev])/r[prev]*100) if r[prev]!=0 else 0, axis=1)
             comp_pivot['Diff'] = comp_pivot[curr] - comp_pivot[prev]
             
             # Sorting (Desc by Current Year)
             comp_pivot = comp_pivot.sort_values(curr, ascending=False)
             
             # Add Grand Total Row
             total_curr = comp_pivot[curr].sum()
             total_prev = comp_pivot[prev].sum()
             total_diff = total_curr - total_prev
             total_growth = ((total_curr - total_prev) / total_prev * 100) if total_prev != 0 else 0
             
             # Create a Total DataFrame
             # User Request: Order 2024 (Prev) then 2025 (Curr)
             df_total_row = pd.DataFrame({
                 prev: [total_prev],
                 curr: [total_curr],
                 'vs LY (%)': [total_growth],
                 'Diff': [total_diff]
             }, index=['Total'])
             
             # Reorder columns for comp_pivot to match: Prev, Curr, vs LY, Diff
             # Existing columns: [curr, prev, 'vs LY (%)', 'Diff'] (may vary)
             cols_ordered = [prev, curr, 'vs LY (%)', 'Diff']
             comp_pivot = comp_pivot[cols_ordered]
             
             # Concatenate Total at the top
             comp_pivot_disp = pd.concat([df_total_row, comp_pivot])
             
             # Display Table
             c_comp1, c_comp2 = st.columns([1, 1])
             with c_comp1:
                 st.write(f"**{t['company_sales_table_header']} ({prev} vs {curr})**")
                 st.dataframe(
                     comp_pivot_disp.style.format("{:,.0f}", subset=[c for c in comp_pivot_disp.columns if c != 'vs LY (%)'])
                                          .format("{:.1f}%", subset=['vs LY (%)'])
                                          .format("{:.1f}%", subset=['vs LY (%)'])
                                          .applymap(style_growth_simple, subset=['vs LY (%)'])
                                          .apply(lambda x: ['font-weight: bold' if x.name == 'Total' else '' for i in x], axis=1),
                     use_container_width=True,
                     height=400 # Fixed height to allow scrolling for "Top 10 visible"
                 )
             
             with c_comp2:
                 # Bar Chart (Comparison: Current vs Previous)
                 st.write(f"**{t['company_sales_chart_header']} ({prev} vs {curr})**")
                 
                 # Prepare data for chart (Remove Total row for chart if we used display df, but we use original pivot or filtered)
                 # Use comp_top_n for chart to avoid clutter if too many?
                 # User didn't strictly say "Limit Chart", but "Top 10 visible in Pivot".
                 # If we show ALL in chart, it might be unreadable. Let's show Top 20 for chart or All if small.
                 chart_data = comp_pivot.head(20).reset_index() # Top 20 Companies
                 
                 # Melt for Grouped Bar
                 chart_data_melted = chart_data.melt(id_vars=[company_col], value_vars=[prev, curr], var_name='Year', value_name='Sales')
                 
                 fig_comp = px.bar(
                     chart_data_melted, 
                     x=company_col, 
                     y='Sales', 
                     color='Year', 
                     barmode='group',
                     text_auto='.2s',
                     title=f"{t['company_sales_chart_header']} (Top 20)", 
                     template="plotly_white",
                     color_discrete_map={str(prev): '#94a3b8', str(curr): '#1e3a8a'} # Gray for Prev, Blue for Curr
                 )
                 fig_comp.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                 st.plotly_chart(fig_comp, use_container_width=True)


    # --- [NEW] Branch & Category Analysis Section ---
    # We want to allow independent Year choice here, so we need a DF that is NOT filtered by the global 'selected_years'.
    # However, it SHOULD still respect 'selected_months' and 'selected_items' if that's the user intent?
    # Usually, "Comparison" implies comparing the same months/items across different years.
    # So we construct a new partial filter base.
    
    df_analysis_base = df.copy()
    # Apply Month Filter
    if selected_months:
        df_analysis_base = df_analysis_base[df_analysis_base[month_col].astype(str).isin([str(m) for m in selected_months])]
    # Apply Item Filter
    if selected_items != "All" and item_col and item_col != "(なし)":
        df_analysis_base = df_analysis_base[df_analysis_base[item_col].astype(str).isin(selected_items)]
        
    # [Sync Request] Respect Sidebar Year Filter
    if selected_years:
        df_analysis_base = df_analysis_base[df_analysis_base[year_col].astype(str).isin([str(y) for y in selected_years])]

    # [Fix] Respect Company Filter for Branch Analysis
    if selected_companies and selected_companies != "All":
         df_analysis_base = df_analysis_base[df_analysis_base[company_col].astype(str).isin(selected_companies)]

    # [Fix] Respect Category Filter for Branch Analysis
    if selected_categories and selected_categories != "All" and category_col and category_col != "(なし)":
         df_analysis_base = df_analysis_base[df_analysis_base[category_col].astype(str).isin(selected_categories)]

    if branch_col and branch_col != "(なし)":
         # Branch Label for Title
         if not selected_branch or selected_branch == "All":
              disp_branch_label_title = "All Branches"
         else:
              if isinstance(selected_branch, list):
                  # User Request: Show actual names
                  disp_branch_label_title = ", ".join(str(b) for b in selected_branch)
              else:
                  disp_branch_label_title = str(selected_branch)

         # --- Year Selection Helpers ---
         # Get all available years from the base data (respecting Item/Month filters)
         avail_years = sorted(df_analysis_base[year_col].unique(), reverse=True)
         
         if not avail_years:
             st.warning("No data available for the selected filters.")
         else:
             st.divider()
             st.subheader(f"🏢 エリア・支店別売上 (Area/Branch Sales) - {disp_branch_label_title}")
             
             # Default Current = Max Year
             def_c = avail_years[0]
             # Default Previous = Global Prev (if exists) else (2nd Max or Max)
             def_p = prev_year if (prev_year and prev_year in avail_years) else (avail_years[1] if len(avail_years) > 1 else avail_years[0])
             
             # Layout: Current (Left) | Previous (Right) | Spacer (swapped and made smaller)
             yc1, yc2, yc3 = st.columns([0.7, 0.7, 4.6])
             with yc1:
                 sec_curr_year = st.selectbox(t["current_year_label"], options=avail_years, index=avail_years.index(def_c), key="sec_cy")
             with yc2:
                 sec_prev_year = st.selectbox(t["prev_year_label"], options=avail_years, index=avail_years.index(def_p), key="sec_py")
             
             # Calculate Data for these specific years
             # We filter df_analysis_base for these 2 years ONLY to optimize? Or just GroupBy?
             # GroupBy is safer.
             df_analysis_filtered = df_analysis_base[df_analysis_base[year_col].isin([sec_curr_year, sec_prev_year])]
             


             # 1. Branch Comparison Table
             # Use df_analysis_filtered (All Branches)
             df_br_comp = df_analysis_filtered.groupby([branch_col, year_col])[val_col_name].sum().unstack(fill_value=0)
             
             # Keys from selection
             sc_key = sec_curr_year
             sp_key = sec_prev_year
             
             # Ensure columns exist in pivot (if one year has no sales)
             if sc_key not in df_br_comp.columns: df_br_comp[sc_key] = 0
             if sp_key not in df_br_comp.columns: df_br_comp[sp_key] = 0

             # metrics
             df_br_comp['vs LY (%)'] = df_br_comp.apply(lambda row: ((row[sc_key] - row[sp_key]) / row[sp_key] * 100) if row[sp_key] != 0 else (0 if row[sc_key] == 0 else 100), axis=1)
             df_br_comp['Diff'] = df_br_comp[sc_key] - df_br_comp[sp_key]
             
             df_br_comp = df_br_comp.sort_values(sc_key, ascending=False)
             
             # Total Row
             total_c = df_br_comp[sc_key].sum()
             total_p = df_br_comp[sp_key].sum()
             total_d = total_c - total_p
             total_g = ((total_c - total_p) / total_p * 100) if total_p != 0 else (0 if total_c == 0 else 100)
             
             df_total = pd.DataFrame({sc_key: [total_c], sp_key: [total_p], 'vs LY (%)': [total_g], 'Diff': [total_d]}, index=['Total'])
             df_br_comp = pd.concat([df_total, df_br_comp])
             
             # Format
             curr_col = sc_key 
             prev_col = sp_key
             
             # Handle duplicate year selection (e.g. 2025 vs 2025) which causes Styler error
             if sp_key == sc_key:
                 df_disp = df_br_comp[[sp_key, sc_key, 'vs LY (%)', 'Diff']].copy()
                 # Manually rename duplicate columns
                 df_disp.columns = [f"{sp_key} (Prev)", f"{sc_key} (Curr)", 'vs LY (%)', 'Diff']
                 format_dict = {f"{sp_key} (Prev)": "{:,.0f}", f"{sc_key} (Curr)": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}
             else:
                 disp_cols = {sp_key: f"{sp_key}", sc_key: f"{sc_key}"}
                 df_disp = df_br_comp[[sp_key, sc_key, 'vs LY (%)', 'Diff']].rename(columns=disp_cols)
                 format_dict = {f"{sp_key}": "{:,.0f}", f"{sc_key}": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}

             # Recode highlight function to reuse closures?
             # Recode highlight function to reuse closures?
             def highlight_rows_local(row):
                 # Only highlight based on Branch selection, NOT color text unless handled by subset
                 style = ''
                 is_bold = False
                 is_yellow = False
                 
                 # Check selection

                 
                 # [Fix] Check Column Value instead of Index Name (since we reset index for resizing)
                 # Handle Branch Name
                 row_name = row.get(branch_col, '')
                 
                 if row_name == 'Total':
                     is_bold = True
                     style += 'background-color: #f0f2f6;' 
                 
                 is_sel = False
                 if isinstance(selected_branch, list):
                     if row_name in selected_branch: is_sel = True
                 elif isinstance(selected_branch, str) and selected_branch != "All":
                      if row_name == selected_branch: is_sel = True
                 
                 if is_sel:
                     is_bold = True
                     is_yellow = True
                 
                 if is_bold: style += 'font-weight: bold;'
                 if is_yellow: style += 'background-color: #ffffcc; color: black;'
                 
                 growth_c = 'black'
                 try:
                    v = float(row['vs LY (%)'])
                    if v >= 0: growth_c = 'green'
                    else: growth_c = 'red'
                 except: pass
                 
                 return [style + (f'color: {growth_c};' if c == 'vs LY (%)' else '') for c in row.index]
             
             # [Fix] Reset Index to make "Column A" (Branch) resizable
             df_disp = df_disp.reset_index()
             # If index name wasn't set on 'Total' row creation, fill generic if needed, 
             # but groupby usually sets it. Rename to "Branch" for display if it's the internal col name
             if branch_col in df_disp.columns:
                 pass # Good
             else:
                 # Fallback if reset_index created 'index'
                 df_disp.rename(columns={'index': branch_col}, inplace=True)

             st.dataframe(
                 df_disp.style.format(format_dict).apply(highlight_rows_local, axis=1).hide(axis="index"),
                 use_container_width=True,
                 height=400,
                 hide_index=True
             )

             # --- [NEW] Visualizations: Branch YoY Bar Chart & Category Treemap ---
             
             # --- Visualizations: Branch YoY Bar Chart & Monthly Sales Trend ---
             
             # Stacked Layout (No Columns)
             st.write(f"**{t['branch_sales_comparison']}**")
             
             # Aggregate data by Branch and Year for the Top Section Year Selection
             df_chart_br = df_analysis_filtered.groupby([branch_col, year_col])[val_col_name].sum().reset_index()
             df_chart_br.rename(columns={branch_col: "Branch", year_col: "Year", val_col_name: "Sales"}, inplace=True)
             
             # Filter based on df_br_comp keys to respect sort order
             order_branches = [idx for idx in df_br_comp.index if idx != 'Total']
             
             fig_br_bar = px.bar(
                 df_chart_br, 
                 x="Branch", 
                 y="Sales", 
                 color="Year", 
                 barmode='group',
                 text_auto='.2s',
                 # User Request: Swap Order (Left: Prev, Right: Curr)
                 category_orders={"Branch": order_branches, "Year": [str(sec_prev_year), str(sec_curr_year)]},
                 template="plotly_white",
                 # Use local variables for colors
                 color_discrete_map={str(sec_prev_year): '#94a3b8', str(sec_curr_year): '#1e3a8a'}
             )
             fig_br_bar.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
             st.plotly_chart(fig_br_bar, use_container_width=True)

             # --- [NEW] Branch Sales Share Treemap (Requested) ---
             st.markdown(f"**{t.get('branch_treemap_title', '支店別売上シェア')} ({sec_curr_year})**")
             # Filter for Current Year
             df_br_tm_src = df_analysis_filtered[df_analysis_filtered[year_col] == sec_curr_year].copy()
             
             if not df_br_tm_src.empty:
                 try:
                     # Group by Branch and Category
                     # Check if category_col exists
                     grp_cols_br_tm = [branch_col]
                     path_cols_br_tm = [branch_col]
                     
                     # Fill NaN for Branch
                     df_br_tm_src[branch_col] = df_br_tm_src[branch_col].fillna("Unknown")
                     
                     if category_col and category_col != "(なし)":
                         grp_cols_br_tm.append(category_col)
                         path_cols_br_tm.append(category_col)
                         # Fill NaN for Category if handling
                         df_br_tm_src[category_col] = df_br_tm_src[category_col].fillna("Unknown")
                     
                     
                     df_br_tm = df_br_tm_src.groupby(grp_cols_br_tm)[val_col_name].sum().reset_index()
                     
                     # [Fix] Treemap requires POSITIVE values. Filter out <= 0.
                     df_br_tm = df_br_tm[df_br_tm[val_col_name] > 0]
                     
                     if not df_br_tm.empty:
                         # Treemap
                         fig_br_tm = px.treemap(
                             df_br_tm, 
                             path=path_cols_br_tm, 
                             values=val_col_name,
                             # color=val_col_name, # Removed to use distinct colors per Branch
                             # color_continuous_scale='Blues',
                             template="plotly_white"
                         )
                         fig_br_tm.update_traces(textinfo="label+value+percent entry", textposition="middle center", textfont=dict(size=14))
                         fig_br_tm.update_layout(margin=dict(t=20, l=10, r=10, b=10))
                         st.plotly_chart(fig_br_tm, use_container_width=True)
                     else:
                         st.info("No positive sales data for Treemap.")
                         
                 except Exception as e:
                     st.error(f"Treemap Error: {e}")
             else:
                 st.info("No data for Treemap.")

             # --- Relocated Monthly Sales Comparison (YoY) ---
             st.markdown(f"**{t['monthly_comparison_title']}**")
             # Group by Year, MonthNum to get temporal order but plot x=Month, color=Year
             groupby_cols = [year_col, month_col, 'MonthNum']
             agg_series = df_filtered.groupby(groupby_cols)[val_col_name].sum()
             
             y_col_for_plot = val_col_name
             if val_col_name in groupby_cols:
                  y_col_for_plot = f"{t['label_total']} {val_col_name}"
             
             trend_data = agg_series.reset_index(name=y_col_for_plot, allow_duplicates=True).sort_values('MonthNum')
             trend_data = trend_data.loc[:, ~trend_data.columns.duplicated()]
             trend_data[year_col] = trend_data[year_col].astype(str)
             
             filter_label = "All"
             if isinstance(selected_branch, list) and selected_branch:
                  if len(selected_branch) > 3: filter_label = f"{len(selected_branch)} Selected"
                  else: filter_label = ", ".join(selected_branch)
             elif selected_branch != "All":
                  filter_label = str(selected_branch)
                  
             fig_trend = px.bar(
                 trend_data, x=month_col, y=y_col_for_plot, color=year_col, barmode='group',
                 title=f"{t['monthly_comparison_title']} - {filter_label}", template="plotly_white"
             )
             
             unique_months_sorted = df[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')[month_col].tolist()
             fig_trend.update_xaxes(categoryorder='array', categoryarray=unique_months_sorted)
             
             try:
                  avail_years_trend = sorted(trend_data[year_col].unique())
                  if avail_years_trend:
                      latest_trend_year = avail_years_trend[-1]
                      df_latest_trend = trend_data[trend_data[year_col] == latest_trend_year].copy().sort_values('MonthNum')
                      if len(df_latest_trend) > 1:
                          x_idx = np.arange(len(df_latest_trend))
                          y_vals = df_latest_trend[y_col_for_plot].values
                          z = np.polyfit(x_idx, y_vals, 1)
                          p = np.poly1d(z)
                          fig_trend.add_trace(go.Scatter(x=df_latest_trend[month_col], y=p(x_idx), mode='lines', name=f"{t['label_trend']} ({latest_trend_year})", line=dict(color='red', width=2, dash='dash')))
             except Exception as e: print(f"Trendline Error: {e}")
             
             st.plotly_chart(fig_trend, use_container_width=True)
             
             # 2. Category Breakdown (Selected Branch + specific years)
             if category_col and category_col != "(なし)":
                 branch_label = "All Branches"
                 # Filtering for the category table:
                 # It needs to filter by the NEW 'section' years AND 'selected_branch'
                 df_cat_base = df_analysis_filtered.copy()
                 
                 # Apply Branch Filter (Global)
                 if isinstance(selected_branch, list) and len(selected_branch) > 0:
                     df_cat_base = df_cat_base[df_cat_base[branch_col].isin(selected_branch)]
                     if len(selected_branch) > 3: branch_label = f"{len(selected_branch)} Selected"
                     else: branch_label = ", ".join(selected_branch)
                 elif isinstance(selected_branch, str) and selected_branch != "All":
                     df_cat_base = df_cat_base[df_cat_base[branch_col] == selected_branch]
                     branch_label = selected_branch
                 
                 st.divider()
                 st.subheader(f"🏢 {t['category_sales_section']} - {branch_label}")
                 st.markdown(f"**{t['category_sales_section']}**")
                 
                 df_cat_comp = df_cat_base.groupby([category_col, year_col])[val_col_name].sum().unstack(fill_value=0)
                 
                 if sc_key not in df_cat_comp.columns: df_cat_comp[sc_key] = 0
                 if sp_key not in df_cat_comp.columns: df_cat_comp[sp_key] = 0
                 
                 # [NEW] Filter out rows where BOTH years are 0
                 df_cat_comp = df_cat_comp[(df_cat_comp[sc_key] != 0) | (df_cat_comp[sp_key] != 0)]
                 
                 df_cat_comp['vs LY (%)'] = df_cat_comp.apply(lambda row: ((row[sc_key] - row[sp_key]) / row[sp_key] * 100) if row[sp_key] != 0 else (0 if row[sc_key] == 0 else 100), axis=1)
                 df_cat_comp['Diff'] = df_cat_comp[sc_key] - df_cat_comp[sp_key]
                 df_cat_comp = df_cat_comp.sort_values(sc_key, ascending=False)
                 
                 # Create Total Row for Category Table
                 cat_total_c = df_cat_comp[sc_key].sum()
                 cat_total_p = df_cat_comp[sp_key].sum()
                 cat_total_d = cat_total_c - cat_total_p
                 cat_total_g = ((cat_total_c - cat_total_p) / cat_total_p * 100) if cat_total_p != 0 else (0 if cat_total_c == 0 else 100)
                 
                 df_cat_total = pd.DataFrame({sc_key: [cat_total_c], sp_key: [cat_total_p], 'vs LY (%)': [cat_total_g], 'Diff': [cat_total_d]}, index=['Total'])
                 df_cat_comp = pd.concat([df_cat_total, df_cat_comp])
                 
                 # Display Category Table
                 # Display Category Table
                 if sp_key == sc_key:
                     df_disp_cat = df_cat_comp[[sp_key, sc_key, 'vs LY (%)', 'Diff']].copy()
                     df_disp_cat.columns = [f"{sp_key} (Prev)", f"{sc_key} (Curr)", 'vs LY (%)', 'Diff']
                     format_dict_cat = {f"{sp_key} (Prev)": "{:,.0f}", f"{sc_key} (Curr)": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}
                 else:
                     df_disp_cat = df_cat_comp[[sp_key, sc_key, 'vs LY (%)', 'Diff']]
                     format_dict_cat = {f"{sp_key}": "{:,.0f}", f"{sc_key}": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}

                 # Recode highlight function for Category Table
                 def highlight_rows_cat_func(row):
                     style = ''
                     is_bold = False
                     
                     # Check Column Value (Category)
                     row_name = row.get(category_col, '')
                     
                     if row_name == 'Total':
                         is_bold = True
                         style += 'background-color: #f0f2f6;' 
                     
                     if is_bold: style += 'font-weight: bold;'
                     
                     growth_c = 'black'
                     try:
                        v = float(row['vs LY (%)'])
                        if v >= 0: growth_c = 'green'
                        else: growth_c = 'red'
                     except: pass
                     
                     return [style + (f'color: {growth_c};' if c == 'vs LY (%)' else '') for c in row.index]

                 # [Fix] Reset Index to make "Column A" (Category) resizable
                 df_disp_cat = df_disp_cat.reset_index()
                 if category_col not in df_disp_cat.columns:
                      df_disp_cat.rename(columns={'index': category_col}, inplace=True)
 
                 st.dataframe(
                     df_disp_cat.style.format(format_dict_cat).apply(highlight_rows_cat_func, axis=1).hide(axis="index"),
                     use_container_width=True,
                     hide_index=True
                 )
                 
                 
                 # --- Visualizations below Table ---
                 # --- Category/Item Share Treemap (Full Width now) ---
                 st.markdown(f"**{t['treemap_title']}**")
                 
                 if category_col and item_col:
                    grp_cols_tm = [category_col, item_col]
                    
                    # [Fix] Use Section-Specific Data (Current Year & Selected Branch) instead of Global df_filtered
                    # This ensures consistency with the tables above.
                    df_tm_src = df_analysis_base[df_analysis_base[year_col] == sec_curr_year]
                    
                    # Apply Branch Filter to Treemap Source
                    if isinstance(selected_branch, list) and len(selected_branch) > 0:
                         df_tm_src = df_tm_src[df_tm_src[branch_col].isin(selected_branch)]
                    elif isinstance(selected_branch, str) and selected_branch != "All":
                         df_tm_src = df_tm_src[df_tm_src[branch_col] == selected_branch]
                    
                    # 1. Aggregate Sales by Category and Item
                    cat_item_share = df_tm_src.groupby(grp_cols_tm)[val_col_name].sum().reset_index()
                    
                    # 2. Limit to Top 5 Items per Category
                    # Sort by Category and Sales (descending)
                    cat_item_share = cat_item_share.sort_values([category_col, val_col_name], ascending=[True, False])
                    
                    # Take Top 5
                    top_5_per_cat = cat_item_share.groupby(category_col).head(5)
                    
                    # Optional: Calculate "Others" for each category? 
                    # To keep it simple and clean as requested ("Top 5 is fine"), we can just show Top 5.
                    # Adding "Others" ensures percentages are correct relative to the Category total, 
                    # but might still clutter if every category has an "Others". 
                    # The user said "Top 5 is fine", so let's try just showing Top 5 first.
                    # However, for a Treemap, if we filter data, the "parent" box size will shrink to sum(children).
                    # This might misrepresent the Category's actual size relative to others.
                    # A better approach for Treemap is usually Global Top 5 Categories -> Top 5 Items, 
                    # OR Keep all Categories, but only show Top 5 Items in them.
                    
                    # Let's create a visual df that includes "Others" to maintain Category size correctness.
                    final_df_list = []
                    for cat, group in cat_item_share.groupby(category_col):
                        if len(group) > 5:
                            top_g = group.head(5)
                            other_val = group.iloc[5:][val_col_name].sum()
                            other_row = pd.DataFrame({category_col: [cat], item_col: ["Others"], val_col_name: [other_val]})
                            final_df_list.append(pd.concat([top_g, other_row], ignore_index=True))
                        else:
                            final_df_list.append(group)
                    
                    if final_df_list:
                        cat_item_tree_data = pd.concat(final_df_list, ignore_index=True)
                    else:
                        cat_item_tree_data = pd.DataFrame(columns=[category_col, item_col, val_col_name])

                    
                    fig_treemap = px.treemap(
                        cat_item_tree_data, path=[category_col, item_col], values=val_col_name,
                        title=f"{t['treemap_chart_title']}", template="plotly_white"
                    )
                    fig_treemap.update_traces(textinfo="label+value+percent entry", textposition="middle center", textfont=dict(size=14))
                    fig_treemap.update_layout(margin=dict(t=40, l=10, r=10, b=10))
                    st.plotly_chart(fig_treemap, use_container_width=True, config={'displayModeBar': True}, key=f"treemap_chart_{st.session_state.get('treemap_key_ver', 0)}")
                 else:
                    st.info("Category or Item column missing.")

                 # --- Visualizations below Table (Moved Down) ---
                 # Monthly Trend Chart (Full Width)
                 st.markdown(f"**{t['monthly_trend_header']}**")
                 
                 # [Fix] Use BOTH Years for Trend Chart (df_cat_base is already filtered to 2 years)
                 df_trend_base = df_cat_base 
                 
                 if not df_trend_base.empty:
                     # Calculate Top Categories for Default Selection (Based on Current Year Sales)
                     df_curr_only = df_trend_base[df_trend_base[year_col] == sec_curr_year]
                     top_cats_trend = df_curr_only.groupby(category_col)[val_col_name].sum().sort_values(ascending=False).head(5).index.tolist()
                     all_cats_sorted = df_curr_only.groupby(category_col)[val_col_name].sum().sort_values(ascending=False).index.tolist()
                     
                     # Initialize session_state if not present
                     if "cat_trend_viz_key" not in st.session_state:
                         st.session_state.cat_trend_viz_key = top_cats_trend
                     
                     st.caption(t['select_category_prompt'])
                     selected_cats_trend = st.multiselect(
                         t['select_category_label'],
                         options=all_cats_sorted,
                         key="cat_trend_viz_key"
                     )
                     
                     if selected_cats_trend:
                         df_trend_viz = df_trend_base[df_trend_base[category_col].isin(selected_cats_trend)].copy() # Copy to avoid SettingWithCopy
                         
                         # Ensure Year is String for Discrete Color/Dash mapping
                         df_trend_viz[year_col] = df_trend_viz[year_col].astype(str)
                         
                         grp_cols = [month_col, category_col, year_col]
                         if 'MonthNum' in df_trend_viz.columns: grp_cols.append('MonthNum')
                         
                         trend_agg = df_trend_viz.groupby(grp_cols)[val_col_name].sum().reset_index()
                         
                         if 'MonthNum' in trend_agg.columns: trend_agg = trend_agg.sort_values('MonthNum')
                         
                         # Explicitly define order
                         sorted_months_cat = df_trend_viz[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')[month_col].tolist()
                         
                         # --- Combo Chart Logic ---
                         # 1. Bar Chart Data (Current Year Only)
                         df_trend_bar = trend_agg[trend_agg[year_col] == str(sec_curr_year)]
                         
                         # 2. Line Chart Data (YoY % Calculation)
                         # Need Total Sales per Month for Current and Previous Year
                         # [Fix] Calculate directly from filtered df_trend_viz to ensure consistency
                         grp_cols_tot = [month_col, year_col]
                         if 'MonthNum' in df_trend_viz.columns: grp_cols_tot.append('MonthNum')
                         
                         df_trend_totals = df_trend_viz.groupby(grp_cols_tot)[val_col_name].sum().unstack(level=year_col, fill_value=0).reset_index()
                         
                         if 'MonthNum' in df_trend_totals.columns:
                             df_trend_totals = df_trend_totals.sort_values('MonthNum')
                         
                         # Correctly handle columns if unstack resulted in year columns
                         # Ensure we use string keys for column access
                         s_curr = str(sec_curr_year)
                         s_prev = str(sec_prev_year)
                         
                         # Ensure columns exist (fill 0 if missing)
                         # Explicitly check string columns
                         cols_str_tot = {str(c): c for c in df_trend_totals.columns}
                         
                         col_curr_tot = cols_str_tot.get(s_curr)
                         col_prev_tot = cols_str_tot.get(s_prev)
                         
                         if not col_curr_tot:
                              df_trend_totals[s_curr] = 0
                              col_curr_tot = s_curr
                         if not col_prev_tot:
                              df_trend_totals[s_prev] = 0
                              col_prev_tot = s_prev
                              
                         df_trend_totals['YoY %'] = df_trend_totals.apply(lambda r: ((r[col_curr_tot] - r[col_prev_tot]) / r[col_prev_tot] * 100) if r[col_prev_tot] != 0 else (0 if r[col_curr_tot] == 0 else 100), axis=1)
                         
                         # --- [NEW] Summary Table Logic (Category) ---
                         # Prepare data for table
                         df_cat_summary = df_trend_totals[[month_col, col_curr_tot, col_prev_tot, 'YoY %']].copy()
                         df_cat_summary['Diff'] = df_cat_summary[col_curr_tot] - df_cat_summary[col_prev_tot]
                         
                         # Ensure sorted by MonthNum if available
                         if 'MonthNum' in df_trend_totals.columns:
                             df_cat_summary['MonthNum'] = df_trend_totals['MonthNum']
                             df_cat_summary = df_cat_summary.sort_values('MonthNum')
                         
                         # Restore sorting if explicit sort list exists
                         if sorted_months_cat:
                             df_cat_summary[month_col] = pd.Categorical(df_cat_summary[month_col], categories=sorted_months_cat, ordered=True)
                             df_cat_summary = df_cat_summary.sort_values(month_col)

                         # Transpose: Rows = Metrics, Cols = Months
                         df_cat_summary_t = df_cat_summary.set_index(month_col)[[col_curr_tot, col_prev_tot, 'Diff', 'YoY %']].T
                         
                         # Add Total Column
                         cat_total_curr = df_cat_summary[col_curr_tot].sum()
                         cat_total_prev = df_cat_summary[col_prev_tot].sum()
                         cat_total_diff = cat_total_curr - cat_total_prev
                         cat_total_yoy = (cat_total_diff / cat_total_prev * 100) if cat_total_prev != 0 else 0
                         
                         df_cat_summary_t['Total'] = [cat_total_curr, cat_total_prev, cat_total_diff, cat_total_yoy]
                         
                         # Rename Index
                         df_cat_summary_t.index = [f"{sec_curr_year}実績", f"{sec_prev_year}実績", "前年差異", "前年比 (%)"]
                         
                         # Display Table
                         st.markdown("##### 前年対比 実績集計表")
                         
                         # Formatting function (reuse style logic if possible, or redefine)
                         def style_cat_summary(styler):
                             styler.format("{:,.0f}", subset=pd.IndexSlice[[f"{sec_curr_year}実績", f"{sec_prev_year}実績", "前年差異"], :])
                             styler.format("{:,.1f}%", subset=pd.IndexSlice[["前年比 (%)"], :])
                             styler.applymap(lambda v: 'color: red;' if v < 0 else 'color: blue;', subset=pd.IndexSlice[["前年差異", "前年比 (%)"], :])
                             return styler

                         st.dataframe(df_cat_summary_t.style.pipe(style_cat_summary), use_container_width=True)
                         # --- End Summary Table ---
                             
                         # Create Figure
                         fig_trend = go.Figure()
                         
                         # Add Stacked Bars for Categories (Current Year)
                         # Explicitly iterate to ensure legend and colors align with overall palette if simpler, 
                         # or use px to generate traces and add them. Using px is easier for colors.
                         
                         fig_base = px.bar(
                             df_trend_bar, x=month_col, y=val_col_name, color=category_col, 
                             category_orders={month_col: sorted_months_cat},
                             template="plotly_white",
                             text_auto='.2s'
                         )
                         for trace in fig_base.data:
                             fig_trend.add_trace(trace)
                             
                         # Add Line Trace (YoY %)
                         fig_trend.add_trace(
                             go.Scatter(
                                 x=df_trend_totals[month_col], 
                                 y=df_trend_totals['YoY %'], 
                                 name=f"YoY % vs {sec_prev_year}", 
                                 mode='lines+markers', 
                                 yaxis='y2',
                                 line=dict(color='red', width=3)
                             )
                         )
                         
                         # Layout with Dual Axis
                         fig_trend.update_layout(
                             title=f"{t['trend_chart_title']} ({sec_curr_year}) & YoY Growth (%)",
                             yaxis=dict(title=t['label_sales'], showgrid=False),
                             yaxis2=dict(
                                 title="YoY Growth (%)", 
                                 overlaying='y', 
                                 side='right', 
                                 showgrid=True, # Grid for % is usually helpful
                                 tickformat='.0f',
                                 ticksuffix='%'
                             ),
                             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                             barmode='relative', # Ensure stacking
                             template="plotly_white",
                             xaxis=dict(categoryorder='array', categoryarray=sorted_months_cat)
                         )
                         st.plotly_chart(fig_trend, use_container_width=True)
                     else:
                         st.info("カテゴリーを選択してください。")
                 else: st.info("No data for current year trend.")
                 
                 # Growth Matrix Chart (Full Width, Below)
                 st.markdown(f"**{t['growth_matrix_title']}**")
                 
                 # [Fix] Exclude 'Total' row and ensure Category column name is restored
                 df_bubble = df_cat_comp[df_cat_comp.index != 'Total'].copy()
                 df_bubble.index.name = category_col # Restore index name if lost during concat
                 df_bubble = df_bubble.reset_index()
                 
                 df_bubble = df_bubble[df_bubble[sc_key] > 0]
                 if not df_bubble.empty:
                     fig_bubble = px.scatter(df_bubble, x=sc_key, y="vs LY (%)", size=sc_key, color=category_col, title=f"{t['growth_matrix_title']} ({sec_curr_year})", template="plotly_white", labels={sc_key: t['label_sales'], "vs LY (%)": t['label_growth']})
                     fig_bubble.add_hline(y=100, line_dash="dash", line_color="gray")
                     st.plotly_chart(fig_bubble, use_container_width=True)
                 else: st.info("No data for growth matrix.")
                         
                 # --- [NEW] Product Monthly Trend (Requested) ---
                 st.markdown(f"**{t['product_trend_title']} - {branch_label}**")
                 
                 if item_col:
                         # 2. Filter for BOTH Years (Current & Previous) to allow YoY calc
                         # Note: Top 5 selection should still be based on Current Year Sales
                         df_prod_trend_viz = df_cat_base[df_cat_base[year_col].isin([sec_curr_year, sec_prev_year])]
                         
                         if not df_prod_trend_viz.empty:
                             # Calculate Top Products based on CURRENT YEAR sales
                             df_prod_curr = df_prod_trend_viz[df_prod_trend_viz[year_col] == sec_curr_year]
                             
                             if not df_prod_curr.empty:
                                 top_prods = df_prod_curr.groupby(item_col)[val_col_name].sum().sort_values(ascending=False).head(5).index.tolist()
                                 all_prods_sorted = df_prod_curr.groupby(item_col)[val_col_name].sum().sort_values(ascending=False).index.tolist()
                                 
                                 # Initialize session_state if not present
                                 if "prod_trend_viz_key" not in st.session_state:
                                     st.session_state.prod_trend_viz_key = top_prods

                                 st.caption(t['select_product_prompt'])
                                 
                                 # Multiselect
                                 selected_prods_trend = st.multiselect(
                                     t['select_product_label'],
                                     options=all_prods_sorted,
                                     key="prod_trend_viz_key"
                                 )
                                 
                                 if selected_prods_trend:
                                     # Filter main df for selected products
                                     df_prod_sel = df_prod_trend_viz[df_prod_trend_viz[item_col].isin(selected_prods_trend)].copy()
                                     
                                     # Ensure Year is String
                                     df_prod_sel[year_col] = df_prod_sel[year_col].astype(str)
                                     
                                     # --- Stacked Bar (Curr) + Line (Prev) Logic ---
                                     
                                     # 1. Prepare Data for Current Year (Stacked Bars by Item)
                                     # Filter for Current Year and Selected Products
                                     df_prod_curr_items = df_prod_sel[df_prod_sel[year_col] == str(sec_curr_year)]
                                     
                                     # Group by Month and Item for Stacked Bars
                                     grp_cols_curr = [month_col, item_col]
                                     if 'MonthNum' in df_prod_curr_items.columns: grp_cols_curr.append('MonthNum')
                                     
                                     prod_curr_agg = df_prod_curr_items.groupby(grp_cols_curr)[val_col_name].sum().reset_index()
                                     if 'MonthNum' in prod_curr_agg.columns: prod_curr_agg = prod_curr_agg.sort_values('MonthNum')
                                     
                                     sorted_months_curr = prod_curr_agg.sort_values('MonthNum' if 'MonthNum' in prod_curr_agg.columns else month_col)[month_col].unique().tolist()

                                     # 2. Prepare Data for Previous Year (Total Line)
                                     # Filter for Previous Year and Selected Products
                                     df_prod_prev_total = df_prod_sel[df_prod_sel[year_col] == str(sec_prev_year)]
                                     
                                     # Group by Month only (Total)
                                     grp_cols_prev = [month_col]
                                     if 'MonthNum' in df_prod_prev_total.columns: grp_cols_prev.append('MonthNum')
                                     
                                     prod_prev_agg = df_prod_prev_total.groupby(grp_cols_prev)[val_col_name].sum().reset_index()
                                     if 'MonthNum' in prod_prev_agg.columns: prod_prev_agg = prod_prev_agg.sort_values('MonthNum')
                                     
                                     # --- [NEW] Summary Table Logic ---
                                     # Aggregate Current Year by Month (Total of selected items)
                                     prod_curr_total_agg = prod_curr_agg.groupby(month_col, as_index=False)[val_col_name].sum()
                                     if 'MonthNum' in prod_curr_agg.columns:
                                         # Merge MonthNum back for sorting
                                         month_order = prod_curr_agg[[month_col, 'MonthNum']].drop_duplicates()
                                         prod_curr_total_agg = prod_curr_total_agg.merge(month_order, on=month_col, how='left').sort_values('MonthNum')
                                     
                                     # Rename val cols for merge
                                     curr_view = prod_curr_total_agg[[month_col, val_col_name]].rename(columns={val_col_name: 'Current'})
                                     prev_view = prod_prev_agg[[month_col, val_col_name]].rename(columns={val_col_name: 'Previous'})
                                     
                                     # Merge Current and Previous
                                     df_summary = pd.merge(curr_view, prev_view, on=month_col, how='outer').fillna(0)
                                     
                                     # Restore sorting if possible
                                     if sorted_months_curr:
                                         df_summary[month_col] = pd.Categorical(df_summary[month_col], categories=sorted_months_curr, ordered=True)
                                         df_summary = df_summary.sort_values(month_col)
                                     
                                     # Calculate Metrics
                                     df_summary['Diff'] = df_summary['Current'] - df_summary['Previous']
                                     df_summary['YoY %'] = df_summary.apply(lambda r: ((r['Current'] - r['Previous']) / r['Previous'] * 100) if r['Previous'] != 0 else (0 if r['Current'] == 0 else 100), axis=1)
                                     
                                     # Transpose for Display (Rows = Metrics, Cols = Months)
                                     df_summary_t = df_summary.set_index(month_col).T
                                     
                                     # Add Total Column (Sum for Diff/Val, Calc for %)
                                     total_curr = df_summary['Current'].sum()
                                     total_prev = df_summary['Previous'].sum()
                                     total_diff = total_curr - total_prev
                                     total_yoy = (total_diff / total_prev * 100) if total_prev != 0 else 0
                                     
                                     df_summary_t['Total'] = [total_curr, total_prev, total_diff, total_yoy]
                                     
                                     # Rename Index for display
                                     df_summary_t.index = [f"{sec_curr_year}実績", f"{sec_prev_year}実績", "前年差異", "前年比 (%)"]
                                     
                                     # Display Table
                                     st.markdown("##### 前年対比 実績集計表")
                                     
                                     # Formatting
                                     def style_summary(styler):
                                         styler.format("{:,.0f}", subset=pd.IndexSlice[[f"{sec_curr_year}実績", f"{sec_prev_year}実績", "前年差異"], :])
                                         styler.format("{:,.1f}%", subset=pd.IndexSlice[["前年比 (%)"], :])
                                         # Color Diff and YoY%
                                         styler.applymap(lambda v: 'color: red;' if v < 0 else 'color: blue;', subset=pd.IndexSlice[["前年差異", "前年比 (%)"], :])
                                         return styler

                                     st.dataframe(df_summary_t.style.pipe(style_summary), use_container_width=True)
                                     
                                     # --- End Summary Table ---
                                     
                                     fig_prod = go.Figure()
                                     
                                     # Trace 1: Stacked Bars for Current Year (Loop through items)
                                     items_in_curr = prod_curr_agg[item_col].unique()
                                     for item in items_in_curr:
                                         item_data = prod_curr_agg[prod_curr_agg[item_col] == item]
                                         fig_prod.add_trace(
                                             go.Bar(
                                                 x=item_data[month_col],
                                                 y=item_data[val_col_name],
                                                 name=str(item),
                                                 texttemplate='%{y:.2s}',
                                                 textposition='auto',
                                                 # Let Plotly handle colors automatically for items
                                             )
                                         )
                                         
                                     # Trace 2: Line for Previous Year Total
                                     fig_prod.add_trace(
                                         go.Scatter(
                                             x=prod_prev_agg[month_col],
                                             y=prod_prev_agg[val_col_name],
                                             name=f"{sec_prev_year} Total (Selected)",
                                             mode='lines+markers+text', # Add text
                                             text=prod_prev_agg[val_col_name],
                                             texttemplate='%{text:.2s}',
                                             textposition='top center',
                                             line=dict(color='gray', width=3, dash='dash'),
                                             marker=dict(size=8)
                                         )
                                     )
                                     
                                     # Layout
                                     fig_prod.update_layout(
                                         title=f"{t['product_trend_title']} ({sec_curr_year} Stacked vs {sec_prev_year} Total)",
                                         yaxis=dict(title=t['label_sales'], showgrid=True),
                                         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                         barmode='stack', # Stacked bars
                                         template="plotly_white",
                                         xaxis=dict(categoryorder='array', categoryarray=sorted_months_curr if sorted_months_curr else None)
                                     )
                                     
                                     st.plotly_chart(fig_prod, use_container_width=True, key=f"prod_trend_{str(selected_prods_trend)}")
                                     
                                 else:
                                     st.info("商品を選択してください。")
                                     
                                 # [DEBUG]
                                 with st.expander("Debug Info (Product Trend)"):
                                     st.write("Selected Products:", selected_prods_trend)
                                     if 'prod_prev_agg' in locals():
                                          st.write(f"Line Data ({sec_prev_year}):", prod_prev_agg[[month_col, val_col_name]].to_dict('records'))
                                     if 'prod_curr_agg' in locals():
                                          st.write(f"Bar Data ({sec_curr_year}):", prod_curr_agg.head())
                                     

                                         
                             else:
                                 st.info(f"{sec_curr_year}年度の商品データがありません。")
                         else:
                             st.info(f"データがありません。")
                 else:
                     st.warning("⚠️ 商品（Item）列が設定されていません。")


    # --- [NEW] Chain Store / Customer Sales Section ---
    st.divider()
    st.subheader(t["chain_customer_sales_header"])
    
    # Column Selection: Let user choose Chain Store, Customer, or Branch
    chain_customer_col_opts = []
    for col in all_columns:
        if is_valid_col(col) and col not in [year_col, month_col, category_col, item_col, val_col_name]:
            chain_customer_col_opts.append(col)
    
    # Try to auto-detect Chain Store or Customer columns
    default_chain_idx = 0
    chain_guess = next((c for c in chain_customer_col_opts if 'chain' in str(c).lower() or 'チェーン' in str(c)), None)
    if chain_guess:
        default_chain_idx = chain_customer_col_opts.index(chain_guess)
    elif branch_col and branch_col in chain_customer_col_opts:
        default_chain_idx = chain_customer_col_opts.index(branch_col)
    
    if chain_customer_col_opts:
        selected_chain_customer_col = st.selectbox(
            t["select_analysis_unit"],
            options=chain_customer_col_opts,
            index=default_chain_idx,
            key="chain_customer_col_selector",
            help="Select the unit for specific analysis (e.g. Chain Store, Customer, Branch)"
        )
        
        st.markdown(f"**{selected_chain_customer_col}{t['unit_sales_comparison']}**")
        
        # Use the same filtered data as category breakdown
        df_chain_comp_base = df_analysis_filtered.copy()
        
        # Apply Branch Filter if specific branches are selected (from sidebar)
        if branch_col and branch_col != "(なし)":
            if isinstance(selected_branch, list) and len(selected_branch) > 0:
                df_chain_comp_base = df_chain_comp_base[df_chain_comp_base[branch_col].isin(selected_branch)]
            elif isinstance(selected_branch, str) and selected_branch != "All":
                df_chain_comp_base = df_chain_comp_base[df_chain_comp_base[branch_col] == selected_branch]
        
        # Group by Selected Column and Year
        df_chain_comp = df_chain_comp_base.groupby([selected_chain_customer_col, year_col])[val_col_name].sum().unstack(fill_value=0)
        
        # Ensure columns exist
        if sec_curr_year not in df_chain_comp.columns: df_chain_comp[sec_curr_year] = 0
        if sec_prev_year not in df_chain_comp.columns: df_chain_comp[sec_prev_year] = 0
        
        # Filter out rows where BOTH years are 0
        df_chain_comp = df_chain_comp[(df_chain_comp[sec_curr_year] != 0) | (df_chain_comp[sec_prev_year] != 0)]
        
        # Calculate metrics (Growth Rate)
        df_chain_comp['vs LY (%)'] = df_chain_comp.apply(
            lambda row: ((row[sec_curr_year] - row[sec_prev_year]) / row[sec_prev_year] * 100) if row[sec_prev_year] != 0 else (0 if row[sec_curr_year] == 0 else 100), 
            axis=1
        )
        df_chain_comp['Diff'] = df_chain_comp[sec_curr_year] - df_chain_comp[sec_prev_year]
        df_chain_comp = df_chain_comp.sort_values(sec_curr_year, ascending=False)
        
        # Display Table
        cc_key = f"FY{sec_curr_year}"
        cp_key = f"FY{sec_prev_year}"
        
        # Handle duplicate year selection
        if sec_curr_year == sec_prev_year:
            cc_disp_key = f"{cc_key} (Curr)"
            cp_disp_key = f"{cp_key} (Prev)"
            
            # Create display DF by selecting the same column twice (since curr == prev)
            # This creates duplicate column names temporarily which we strictly adhere to overwrite immediately
            # Note: sec_curr_year is used for both
            df_disp_chain = df_chain_comp[[sec_curr_year, sec_curr_year, 'vs LY (%)', 'Diff']].copy()
            df_disp_chain.columns = [cp_disp_key, cc_disp_key, 'vs LY (%)', 'Diff']
            format_dict_chain = {f"{cp_disp_key}": "{:,.0f}", f"{cc_disp_key}": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}
            
            df_disp_chain = df_disp_chain.reset_index()
        else:
            df_chain_comp = df_chain_comp.rename(columns={sec_curr_year: cc_key, sec_prev_year: cp_key})
            format_dict_chain = {f"{cp_key}": "{:,.0f}", f"{cc_key}": "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff": "{:,.0f}"}
            df_disp_chain = df_chain_comp[[cp_key, cc_key, 'vs LY (%)', 'Diff']].reset_index()

        # [Fix] Reset index for resizing
        df_disp_chain = df_disp_chain.reset_index(drop=True) if 'index' in df_disp_chain.columns else df_disp_chain
        
        # Styling Function
        def style_chain_growth(styler):
            styler.format(format_dict_chain)
            # Color YoY%: Green if > 0, Red if < 0
            styler.applymap(lambda v: 'color: green;' if v > 0 else ('color: red;' if v < 0 else None), subset=['vs LY (%)'])
            return styler
        
        st.dataframe(
            df_disp_chain.style.pipe(style_chain_growth).hide(axis="index"),
            use_container_width=True,
            hide_index=True,
            column_config={
                selected_chain_customer_col: st.column_config.TextColumn(
                    selected_chain_customer_col,
                    width="medium"
                )
            }
        )
    else:
        st.warning(t["warn_no_valid_col"])
    
    
    st.divider()
    st.subheader(t["customer_strategy_header"])
    st.markdown(t["customer_strategy_desc"])

    # 0. Select Analysis Axis (Customer Identifier)
    # User wants to chose "Chain Store" or "Customer"
    st.caption(t["select_analysis_level_caption"])
    # Default to Branch Col as fallback, but try to find "Chain" or "Customer"
    strat_col_opts = [c for c in all_columns if is_valid_col(c)]
    
    # Auto-detect "Chain" or "Customer"
    auto_strat_idx = 0
    if branch_col in strat_col_opts:
        auto_strat_idx = strat_col_opts.index(branch_col)
    
    # refine default logic if "Chain" exists
    if "strat_col_axis" in st.session_state and st.session_state.strat_col_axis in strat_col_opts:
        auto_strat_idx = strat_col_opts.index(st.session_state.strat_col_axis)
    else:
        chain_guess = next((c for c in strat_col_opts if 'Chain' in str(c) or 'チェーン' in str(c)), None)
        if chain_guess:
            auto_strat_idx = strat_col_opts.index(chain_guess)
         
    target_cust_col = st.selectbox(t["select_target_col"], strat_col_opts, index=auto_strat_idx, key="strat_col_axis")
    
    # --- [NEW] Analysis Metric Selector (Sales vs Qty) ---
    # Find potential numeric columns
    num_cols = df_filtered.select_dtypes(include=['number']).columns.tolist()
    # Try to identify Sales vs Qty vs Others (Price, Profit, etc)
    # Broaden keywords to capture variety
    cand_sales = [c for c in num_cols if any(k in c.lower() for k in ['val', 'sales', 'amount', 'rev', '売上', '金額'])]
    cand_qty = [c for c in num_cols if any(k in c.lower() for k in ['qty', 'quant', 'case', 'vol', '数', 'ケース'])]
    cand_other = [c for c in num_cols if any(k in c.lower() for k in ['price', 'cost', 'profit', 'margin', '単価', '原価', '利益'])]
    
    # Default options
    metric_options = []
    if val_col_name in num_cols: # Primary
        metric_options.append(val_col_name)
    
    # Add unique candidates
    for c in cand_sales + cand_qty + cand_other:
        if c not in metric_options:
            metric_options.append(c)
            
    # Fallback
    if not metric_options:
        metric_options = num_cols

    # Default to current global val_col_name or persistent session state
    if "strat_metric_sel" in st.session_state and st.session_state.strat_metric_sel in metric_options:
        def_metric_idx = metric_options.index(st.session_state.strat_metric_sel)
    else:
        def_metric_idx = metric_options.index(val_col_name) if val_col_name in metric_options else 0
        
    strat_val_col = st.radio(t["analysis_metric_label"], metric_options, index=def_metric_idx, horizontal=True, key="strat_metric_sel")

    # 1. Customer Selection
    # Populate options based on selected column
    avail_customers = []
    if target_cust_col:
        # Get unique values, sorted
        avail_customers = sorted(df_filtered[target_cust_col].astype(str).unique().tolist())
        
    selected_customer_strat = []
    if target_cust_col:
        # Get unique values, sorted. Use global 'df' if available to prevent options disappearing on filter, otherwise df_filtered.
        # This ensures persistence of selection even if currently filtered out by date etc.
        source_df = df if 'df' in locals() else df_filtered
        avail_customers = sorted(source_df[target_cust_col].astype(str).unique().tolist())
        
        # Check for column switch to reset selection
        if "last_strat_col" not in st.session_state:
             st.session_state.last_strat_col = target_cust_col
        
        if st.session_state.last_strat_col != target_cust_col:
             st.session_state.strat_cust_sel = [] # Reset selection
             st.session_state.last_strat_col = target_cust_col
             
        # Initialize key if missing
        if "strat_cust_sel" not in st.session_state:
             st.session_state.strat_cust_sel = []

        selected_customer_strat = st.multiselect(
            f"{t['customer_search_label']} ({target_cust_col})", 
            options=avail_customers,
            key="strat_cust_sel", 
            help="Type to search (multiple allowed)"
        )

    if selected_customer_strat:
        # Filter Data for this Customer
        df_strat = df_filtered[df_filtered[target_cust_col].astype(str).isin(selected_customer_strat)]
        
        # Helper for display name
        cust_label_short = ", ".join(selected_customer_strat)
        if len(cust_label_short) > 30:
            cust_label_short = f"{len(selected_customer_strat)} Selected"
        
        if not df_strat.empty:
            
            # --- [NEW] Advanced Annual Comparison (Requested Layout) ---
            st.markdown(f"**{t['yoy_comparison_header']}**")
            
            # 1. Year Selection (Current vs Previous)
            avail_years = sorted(df_strat[year_col].unique(), reverse=True)
            if len(avail_years) < 2:
                st.warning(t["warn_insufficient_year"])
                default_curr, default_prev = (avail_years[0], None) if avail_years else (None, None)
            else:
                default_curr = avail_years[0]
                default_prev = avail_years[1]
                
            # Swap positions: Current (Left) | Previous (Right) and make smaller
            c_year1, c_year2, c_year3 = st.columns([0.7, 0.7, 4.6])
            with c_year1:
                cur_year_sel = st.selectbox(t["current_year_label"], avail_years, index=avail_years.index(default_curr) if default_curr in avail_years else 0, key="yoy_curr_sel")
            with c_year2:
                prev_year_sel = st.selectbox(t["prev_year_label"], avail_years, index=avail_years.index(default_prev) if default_prev in avail_years else 0, key="yoy_prev_sel")
                
            # 2. Tabs for Category / Branch Logic
            # "Branch" in this context: 
            # If target_cust_col is "Chain Store", then "Branch" could be "Customer Name" (individual store).
            # If target_cust_col is "Customer Name", then "Branch" doesn't exist, maybe "Prop" (Property)?
            # Let's try to detect a sub-dimension.
            
            sub_dim_col = None
            
            # Priority: Granular Store Name > Branch
            sub_dim_candidates = ['Customer Name', 'Store Name', 'Customer', '店舗名', '店名']
            
            # Find first match
            found_store_col = None
            for cand in sub_dim_candidates:
                # Case insensitive check
                match = next((c for c in df_strat.columns if str(c).lower() == cand.lower()), None)
                if match and match != target_cust_col:
                    found_store_col = match
                    break
            
            if found_store_col:
                sub_dim_col = found_store_col
            elif branch_col and branch_col != target_cust_col:
                 sub_dim_col = branch_col
            
            # Shared Definitions
            display_cols = {cur_year_sel: f"FY{cur_year_sel}", prev_year_sel: f"FY{prev_year_sel}"}
            
            def style_growth(v):
                try:
                    val = float(str(v).replace('%',''))
                    return 'color: green' if val > 0 else 'color: red' if val < 0 else None
                except:
                    return None
            
            # --- [Moved] Total Monthly Sales Summary ---
            st.markdown(f"**{t['cust_monthly_trend_title']} {cust_label_short}**")
            
            # Filter for Current and Previous Year
            df_cust_trend_base = df_strat[df_strat[year_col].isin([cur_year_sel, prev_year_sel])].copy()
            cust_summary = None
            month_order = []

            if not df_cust_trend_base.empty:
                 # --- Summary Table Preparation ---
                cust_summary = df_cust_trend_base.groupby([month_col, year_col])[val_col_name].sum().unstack(fill_value=0)
                
                # Ensure columns
                if cur_year_sel not in cust_summary.columns: cust_summary[cur_year_sel] = 0
                if prev_year_sel not in cust_summary.columns: cust_summary[prev_year_sel] = 0
                
                # Sort by MonthNum
                if 'MonthNum' in df_filtered.columns:
                    month_map = df_filtered[[month_col, 'MonthNum']].drop_duplicates().set_index(month_col)['MonthNum']
                    cust_summary['MonthNum'] = cust_summary.index.map(month_map)
                    cust_summary = cust_summary.sort_values('MonthNum').drop(columns=['MonthNum'])
                else:
                    # Fallback sort
                    try:
                        cust_summary = cust_summary.loc[sorted_months_curr] if 'sorted_months_curr' in locals() else cust_summary
                    except:
                        pass

                # Calculate summary metrics
                cust_summary['Diff'] = cust_summary[cur_year_sel] - cust_summary[prev_year_sel]
                cust_summary['YoY %'] = cust_summary.apply(lambda row: (row['Diff'] / row[prev_year_sel] * 100) if row[prev_year_sel] != 0 else 0, axis=1)
                
                # Transpose for Display
                cust_summary_t = cust_summary.T
                
                # Add Total Column
                total_c = cust_summary[cur_year_sel].sum()
                total_p = cust_summary[prev_year_sel].sum()
                total_d = total_c - total_p
                total_y = (total_d / total_p * 100) if total_p != 0 else 0
                cust_summary_t['Total'] = [total_p, total_c, total_d, total_y] 
                
                cust_summary_t = cust_summary_t.reindex([cur_year_sel, prev_year_sel, 'Diff', 'YoY %'])
                cust_summary_t.index = [f"{cur_year_sel}{t['col_sales_suffix']}", f"{prev_year_sel}{t['col_sales_suffix']}", t['col_diff'], t['col_yoy']]
                
                # Get month order for chart
                month_order = cust_summary.index.tolist()

                # Style Function
                def style_cust_summary(styler):
                    styler.format("{:,.0f}", subset=pd.IndexSlice[[f"{cur_year_sel}{t['col_sales_suffix']}", f"{prev_year_sel}{t['col_sales_suffix']}", t['col_diff']], :])
                    styler.format("{:,.1f}%", subset=pd.IndexSlice[[t['col_yoy']], :])
                    styler.applymap(lambda v: 'color: red;' if v < 0 else 'color: blue;', subset=pd.IndexSlice[[t['col_diff'], t['col_yoy']], :])
                    return styler
                
                st.markdown(f"##### {t['yoy_summary_table_title']}")
                st.dataframe(cust_summary_t.style.pipe(style_cust_summary), use_container_width=True)

            # --- Aggregation Tables ---
            st.markdown(f"**{t['category_breakdown_header']}**")
            
            # 2.1 Category Performance
            if category_col and category_col != "(なし)":
                br_perf = df_strat.groupby([category_col, year_col])[strat_val_col].sum().unstack(fill_value=0)
                
                # Ensure columns exist
                if cur_year_sel not in br_perf.columns: br_perf[cur_year_sel] = 0
                if prev_year_sel not in br_perf.columns: br_perf[prev_year_sel] = 0
                
                # Calc YoY
                br_perf['vs LY (%)'] = br_perf.apply(lambda x: ((x[cur_year_sel] - x[prev_year_sel])/x[prev_year_sel]*100) if x[prev_year_sel]!=0 else 0, axis=1)
                br_perf['Diff (Plug)'] = br_perf[cur_year_sel] - br_perf[prev_year_sel]
                
                # Sort
                br_perf = br_perf.sort_values(cur_year_sel, ascending=False)
                
                # Display
                if prev_year_sel == cur_year_sel:
                    key_prev = f"FY{prev_year_sel} (Prev)"
                    key_curr = f"FY{cur_year_sel} (Curr)"
                    br_perf_disp = br_perf[[prev_year_sel, cur_year_sel, 'vs LY (%)', 'Diff (Plug)']].copy()
                    br_perf_disp.columns = [key_prev, key_curr, 'vs LY (%)', 'Diff (Plug)']
                    format_dict_yoy = {key_prev: "{:,.0f}", key_curr: "{:,.0f}", "vs LY (%)": "{:,.1f}%", "Diff (Plug)": "{:,.0f}"}
                else:
                    br_perf_disp = br_perf[[prev_year_sel, cur_year_sel, 'vs LY (%)', 'Diff (Plug)']].rename(columns=display_cols)
                    format_dict_yoy = {
                        f"FY{prev_year_sel}": "{:,.0f}",
                        f"FY{cur_year_sel}": "{:,.0f}",
                        "vs LY (%)": "{:,.1f}%",
                        "Diff (Plug)": "{:,.0f}"
                    }

                # Add Total Row
                br_perf_disp = br_perf_disp.reset_index()
                # Calculate footer sums handled ? No pandas styler handles it usually or we append.
                # Appending Total row for clarity
                # Sum numeric cols
                # sum_row = br_perf_disp.sum(numeric_only=True)
                # sum_row[category_col] = 'Total'
                # Recalc % for Total
                # Need raw sums first
                # tot_curr = br_perf[cur_year_sel].sum()
                # tot_prev = br_perf[prev_year_sel].sum()
                # tot_diff = tot_curr - tot_prev
                # tot_yoy = (tot_diff / tot_prev * 100) if tot_prev != 0 else 0
                
                # Map back to display columns
                # This is getting complex to overwrite, let's keep it simple for now as requested (just enable multiselect)
                # Just showing table is fine.

                st.dataframe(
                    br_perf_disp.style.format(format_dict_yoy).map(style_growth, subset=["vs LY (%)"]).hide(axis="index"),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info(t["warn_no_cat_col"])
            
            # 2.2 Store YoY Performance (Summary)
            if sub_dim_col:
                st.markdown(f"**{t['store_yoy_title']} ({sub_dim_col})**")
                
                # Group by Store and Year
                store_perf = df_strat.groupby([sub_dim_col, year_col])[strat_val_col].sum().unstack(fill_value=0)
                
                # Check Years
                if cur_year_sel in store_perf.columns and prev_year_sel in store_perf.columns:
                     # Calculate Metrics
                     store_perf['Diff'] = store_perf[cur_year_sel] - store_perf[prev_year_sel]
                     store_perf['Growth %'] = (store_perf['Diff'] / store_perf[prev_year_sel] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                     
                     # Sort by Current Year Sales
                     store_perf = store_perf.sort_values(cur_year_sel, ascending=False)
                     
                     # Select & Rename Columns
                     disp_cols_s = [prev_year_sel, cur_year_sel, 'Diff', 'Growth %']
                     store_perf_disp = store_perf[disp_cols_s].copy()
                     store_perf_disp.columns = [f"{prev_year_sel}{t['col_sales_suffix']}", f"{cur_year_sel}{t['col_sales_suffix']}", t['col_diff'], t['col_yoy']]
                     
                     # Format
                     format_dict_s = {
                         f"{prev_year_sel}{t['col_sales_suffix']}": "{:,.0f}",
                         f"{cur_year_sel}{t['col_sales_suffix']}": "{:,.0f}",
                         t['col_diff']: "{:+,.0f}",
                         t['col_yoy']: "{:+.1f}%"
                     }
                     
                     st.dataframe(
                         store_perf_disp.style.format(format_dict_s).bar(subset=[t['col_diff']], color=['#ffcccc', '#ccffcc'], align='mid').map(style_growth, subset=[t['col_yoy']]),
                         use_container_width=True
                     )
                else:
                     st.info(t["warn_year_data_missing"])
            
            st.divider()
            # 2.2 Branch/Sub-dim Performance (Monthly)
            if sub_dim_col:
                st.markdown(f"**{t['store_monthly_title']} ({sub_dim_col} - {cur_year_sel})**")
                group_col = sub_dim_col
                
                # Check if grouping is possible
                if group_col in df_strat.columns:
                    # Filter for Current Year only for this wide view
                    df_sub_curr = df_strat[df_strat[year_col] == cur_year_sel]
                    
                    if not df_sub_curr.empty:
                        # Pivot: Index=Store, Columns=Month, Values=Sales
                        br_monthly = df_sub_curr.groupby([group_col, month_col])[strat_val_col].sum().unstack(fill_value=0)
                        
                        # Sort Columns (Months)
                        if 'MonthNum' in df_filtered.columns:
                            # Create a sorter based on global MonthNum
                            month_map = df_filtered[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')
                            sorted_months = [m for m in month_map[month_col] if m in br_monthly.columns]
                            br_monthly = br_monthly[sorted_months]
                        else:
                            # Fallback to existing sort list if available or alphabetical
                            # Assuming columns are already reasonable or use standard sort
                            pass

                        # Add Total Column
                        br_monthly['Total'] = br_monthly.sum(axis=1)
                        
                        # Sort Rows by Total
                        br_monthly = br_monthly.sort_values('Total', ascending=False)
                        
                        # Format
                        format_dict_br = {col: "{:,.0f}" for col in br_monthly.columns}
                        
                        # Reset index to show Store Name as column
                        br_monthly_disp = br_monthly.reset_index()
                        
                        st.dataframe(
                            br_monthly_disp.style.format(format_dict_br).background_gradient(cmap="Blues", subset=br_monthly.columns[:-1], axis=None),
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        st.info(f"{cur_year_sel}{t['year_sales_suffix']} {t['no_data']}")

                else:
                    st.info(f"{t['warn_no_agg_col']} ({group_col})")
            else:
                st.info(f"{t['warn_no_sub_dim_col']}")
            
            # --- [Refactor] Full Width Layout for Better Visibility ---
            
            # --- Chart Preparation (Full Width) ---
            # 1. Stacked Bars (Current Year, Grouped by Category)
            if not df_cust_trend_base.empty and cust_summary is not None:
                df_curr_cat = df_cust_trend_base[df_cust_trend_base[year_col] == cur_year_sel]
                # Agg check
                curr_cat_agg = df_curr_cat.groupby([month_col, category_col])[val_col_name].sum().reset_index()
                
                # month_order is defined above
                
                fig_cust = go.Figure()
                
                # Add Traces for each Category
                for cat in sorted(curr_cat_agg[category_col].unique()):
                    cat_data = curr_cat_agg[curr_cat_agg[category_col] == cat]
                    fig_cust.add_trace(go.Bar(
                        x=cat_data[month_col],
                        y=cat_data[val_col_name],
                        name=cat,
                        texttemplate='%{y:.2s}'
                    ))
                
                # 2. Line (Previous Year Total)
                prev_total_agg = df_cust_trend_base[df_cust_trend_base[year_col] == prev_year_sel].groupby(month_col)[val_col_name].sum().reindex(month_order).fillna(0).reset_index()
                
                fig_cust.add_trace(go.Scatter(
                    x=prev_total_agg[month_col],
                    y=prev_total_agg[val_col_name],
                    name=f"{prev_year_sel} {t['col_sales_suffix']}",
                    mode='lines+markers',
                    line=dict(color='gray', width=2, dash='dash'),
                    marker=dict(size=6, symbol='circle')
                ))
                
                fig_cust.update_layout(
                    title=f"{t['cust_monthly_trend_title']} ({cur_year_sel} vs {prev_year_sel})",
                    barmode='stack',
                    yaxis=dict(title=t['label_sales'], showgrid=True),
                    xaxis=dict(categoryorder='array', categoryarray=month_order),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    template="plotly_white",
                    hovermode="x unified"
                )
                st.plotly_chart(fig_cust, use_container_width=True)
            else:
                st.info("データがありません。")

            st.divider()

            # 3. Product Opportunity Matrix (Bubble)
            st.markdown(f"**{t['product_opp_matrix_header']} ({strat_val_col} vs {t['label_growth']})**")
            # Need Current Year vs Previous Year Comparison
            # Assuming current year is max year in data or selected context?
            # Use years selected above
            current_year = cur_year_sel
            prev_year = prev_year_sel
            
            if current_year == prev_year:
                st.warning("Product Matrix requires two different years for comparison.")
                current_year = None # Skip rendering
                
            if current_year and prev_year:
                # Filter df for years
                df_strat_yoy = df_strat[df_strat[year_col].isin([current_year, prev_year])]
                # Pivot
                item_perf = df_strat_yoy.groupby([item_col, year_col])[strat_val_col].sum().unstack(fill_value=0)
                if current_year in item_perf.columns and prev_year in item_perf.columns:
                    item_perf['Growth %'] = ((item_perf[current_year] - item_perf[prev_year]) / item_perf[prev_year] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                    item_perf = item_perf.reset_index()
                    # Filter out zero sales in current year
                    item_perf = item_perf[item_perf[current_year] > 0]
                    
                    if not item_perf.empty:
                        fig_bub = px.scatter(
                            item_perf, 
                            x=current_year, 
                            y='Growth %', 
                            size=current_year, 
                            hover_name=item_col,
                            color='Growth %',
                            title=f"{t['product_opp_matrix_header']} ({cust_label_short})",
                            template="plotly_white",
                            labels={current_year: f"Current {strat_val_col}", 'Growth %': "YoY Growth (%)"}
                        )
                        fig_bub.add_hline(y=0, line_dash="dash", line_color="gray")
                        st.plotly_chart(fig_bub, use_container_width=True)
                    else:
                        st.info(t["matrix_no_data"])
                else:
                    st.info(t["matrix_no_prev"])
            else:
                    st.info(t["warn_year_data_missing"])

            # 4. [Moved Up] Category Monthly Trend (Requested)
            if category_col and category_col != "(なし)":
                 st.markdown(f"**{t['cat_trend_header']} ({selected_customer_strat})**")
                 # Default Top 5 Categories
                 top_cats_strat = df_strat.groupby(category_col)[strat_val_col].sum().sort_values(ascending=False).head(5).index.tolist()
                 # Ensure defaults are in options
                 cat_opts = sorted(df_strat[category_col].unique())
                 # Filter defaults to be safe
                 top_cats_strat = [c for c in top_cats_strat if c in cat_opts]
                 
                 sel_cats_strat = st.multiselect(t['select_category_label'], cat_opts, default=top_cats_strat, key="strat_cat_multi")
                 
                 if sel_cats_strat:
                     df_strat_cats = df_strat[df_strat[category_col].isin(sel_cats_strat)]
                     # Group by Month + Category
                     df_cat_trend = df_strat_cats.groupby([month_col, category_col])[strat_val_col].sum().reset_index()
                     # Sort by MonthNum
                     if 'MonthNum' in df_strat.columns:
                          df_cat_trend = df_cat_trend.merge(df_strat[[month_col, 'MonthNum']].drop_duplicates(), on=month_col, how='left').sort_values('MonthNum')
                     
                     # Explicitly define order
                     sorted_months_strat_cat = df_cat_trend[month_col].unique().tolist()

                     fig_cat_trend = px.line(
                         df_cat_trend, 
                         x=month_col, 
                         y=strat_val_col, 
                         color=category_col, 
                         markers=True, 
                         title="Category Monthly Sales Trend", 
                         template="plotly_white",
                         category_orders={month_col: sorted_months_strat_cat}
                     )
                     st.plotly_chart(fig_cat_trend, use_container_width=True)

            # 5. Product Monthly Trend (Multi-line)
            st.markdown(f"**{t['prod_trend_header']} ({selected_customer_strat})**")
            # Default Top 5 Items
            top_items_strat = df_strat.groupby(item_col)[val_col_name].sum().sort_values(ascending=False).head(5).index.tolist()
            sel_items_strat = st.multiselect(t['select_product_label'], sorted(df_strat[item_col].unique()), default=top_items_strat, key="strat_item_multi")
            
            if sel_items_strat:
                df_strat_items = df_strat[df_strat[item_col].isin(sel_items_strat)]
                # Group by Month + Item
                df_item_trend = df_strat_items.groupby([month_col, item_col])[val_col_name].sum().reset_index()
                # Sort by MonthNum
                # Sort by MonthNum
                if 'MonthNum' in df_strat.columns:
                     df_item_trend = df_item_trend.merge(df_strat[[month_col, 'MonthNum']].drop_duplicates(), on=month_col, how='left').sort_values('MonthNum')
                     
                # Correctly Sort X-Axis for Plotly
                # We need to explicitly tell Plotly the order of categories, otherwise it might default to trace order or alphabetical
                sorted_months = df_item_trend[month_col].unique().tolist()
                
                fig_item_trend = px.line(
                    df_item_trend, 
                    x=month_col, 
                    y=val_col_name, 
                    color=item_col, 
                    markers=True, 
                    title="Product Monthly Sales Trend", 
                    template="plotly_white",
                    category_orders={month_col: sorted_months} # FORCE ORDER
                )
                st.plotly_chart(fig_item_trend, use_container_width=True)

        else:
            st.warning(t["warn_no_cust_data"])

    st.divider()

    
    # --- [NEW] Growth Driver Analysis (Outside of Customer Filter) ---
    st.subheader(t["growth_driver_header"])
    st.caption(t["growth_driver_desc"])

    # 1. Select Axis
    # Default to Item
    axis_opts = {"商品 (Item)": item_col, "カテゴリー (Category)": category_col}
    # Add Brand if available? (target_cust_col might be brand in some contexts, but usually 'Brand' column)
    # Let's check columns for 'Brand'
    brand_candidate = [c for c in df_filtered.columns if "brand" in str(c).lower()]
    if brand_candidate:
        axis_opts["ブランド (Brand)"] = brand_candidate[0]

    c_gd1, c_gd2, c_gd3 = st.columns([1, 1, 1])
    with c_gd1:
        gd_axis_label = st.radio(t["analysis_axis_label"], list(axis_opts.keys()), horizontal=True, key="gd_axis_radio")
        gd_axis_col = axis_opts[gd_axis_label]
        
    # 2. Select Target Value
    # Get unique values for the selected column
    if gd_axis_col:
        gd_opts = sorted(df_filtered[gd_axis_col].astype(str).unique().tolist())
        with c_gd2:
            gd_target = st.selectbox(t["select_target_label"], gd_opts, key="gd_target_sel")
    else:
        gd_target = None
        st.warning(t["warn_axis_col_missing"])

    # 3. Ranking Unit Selection
    # User Request: "Select between Chain Store and Customer"
    # Chain Store -> Rank Chains
    # Customer -> Rank Individual Stores (Customer Name / Store Name)
    
    rank_opts = {}
    if target_cust_col:
        rank_opts["チェーンストア (Chain Store)"] = target_cust_col
    
    # Logic to find the best "Customer/Store" column
    # Priority: "Customer Name" > "Store Name" > "Customer" > branch_col
    store_col_candidates = ["Customer Name", "Store Name", "Customer", "Store", "店舗名", "店名"]
    best_store_col = None
    
    # 1. Check specific candidates
    for c in store_col_candidates:
        # Case insensitive check
        match = next((col for col in df_filtered.columns if str(col).lower() == c.lower()), None)
        if match:
            best_store_col = match
            break
    
    # 2. Fallback to branch_col if valid and not same as target_cust_col
    if not best_store_col and branch_col and branch_col != target_cust_col:
        best_store_col = branch_col
        
    if best_store_col:
        rank_opts["店舗・個店 (Customer)"] = best_store_col
        
    # If columns are same (e.g. only Chain exists), just show Chain
    if not best_store_col and target_cust_col:
        # Only Chain available
        pass 
        
    with c_gd3:
        if rank_opts:
            gd_rank_label = st.radio(t["ranking_unit_label"], list(rank_opts.keys()), horizontal=True, key="gd_rank_unit")
            gd_loc_col = rank_opts[gd_rank_label]
        else:
            gd_loc_col = None
            st.error(t["warn_agg_col_missing"])

    if gd_target and gd_loc_col:
        # Filter Data by Target
        df_gd = df_filtered[df_filtered[gd_axis_col].astype(str) == gd_target]
        
        # Calculate Growth by Location
        # Needs Year Col
        avail_years_gd = sorted(df_gd[year_col].unique(), reverse=True)
        if len(avail_years_gd) >= 2:
            curr_y = avail_years_gd[0]
            prev_y = avail_years_gd[1]
            
            # Group by Aggregation Unit
            gd_perf = df_gd.groupby([gd_loc_col, year_col])[strat_val_col].sum().unstack(fill_value=0)
            
            if curr_y in gd_perf.columns and prev_y in gd_perf.columns:
                gd_perf['Growth Amount'] = gd_perf[curr_y] - gd_perf[prev_y]
                gd_perf['Growth %'] = (gd_perf['Growth Amount'] / gd_perf[prev_y] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                
                # Sort by Growth Amount DESC (Top Growers)
                # User Request: "Scroll to see more" -> Increase limit from 10 to 100
                gd_top_n = gd_perf.sort_values('Growth Amount', ascending=False).head(100)
                
                # Display
                st.markdown(f"**🏆 {t['growth_top_100_header']} {gd_target} ({gd_rank_label} Top 100)**")
                st.caption(t["scroll_caption"])
                
                # Format
                gd_disp = gd_top_n[[prev_y, curr_y, 'Growth Amount', 'Growth %']].copy()
                gd_disp = gd_disp.rename(columns={prev_y: f"{prev_y} Sales", curr_y: f"{curr_y} Sales"})
                
                # Style function for Pos/Neg
                def style_pos_neg(v):
                    try:
                        val = float(v)
                        if val > 0:
                            return 'color: green; font-weight: bold;'
                        elif val < 0:
                            return 'color: red; font-weight: bold;'
                        return ''
                    except:
                        return ''
                
                # Use column config for better name display
                st.dataframe(
                    gd_disp.style.format({
                        f"{prev_y} Sales": "{:,.0f}", 
                        f"{curr_y} Sales": "{:,.0f}", 
                        "Growth Amount": "{:+,.0f}", 
                        "Growth %": "{:+.1f}%"
                    })
                    .bar(subset=['Growth Amount'], color=['#ffcccc', '#ccffcc'], align='mid')
                    .map(style_pos_neg, subset=['Growth Amount', 'Growth %']),
                    use_container_width=True,
                    height=500, # Enable scrolling
                    column_config={
                        gd_loc_col: st.column_config.TextColumn(
                            "Name",
                            width="medium", 
                            help=gd_loc_col
                        )
                    }
                )
            else:
                st.info(t["warn_year_data_missing"])
                
            # --- [NEW] Drill Down: Monthly Sales for Selected Store ---
            if 'gd_top_n' in locals() and not gd_top_n.empty:
                st.markdown("---")
                st.caption("👇 Select a store to see monthly details")
                
                # Selection
                gd_store_opts = gd_top_n.index.tolist()
                selected_gd_store = st.selectbox(
                    f"{t['select_store_monthly']} ({gd_rank_label})", 
                    gd_store_opts, 
                    key="gd_store_detail_sel"
                )
                
                if selected_gd_store:
                    st.markdown(f"**🗓️ {selected_gd_store} - Monthly Sales Trend ({gd_target})**")
                    
                    # Filter for this store
                    df_gd_store = df_gd[df_gd[gd_loc_col] == selected_gd_store]
                    
                    # Pivot Table (Month x Year)
                    gd_monthly = df_gd_store.groupby([month_col, year_col])[strat_val_col].sum().unstack(fill_value=0)
                    
                    # Sort Columns (Months)
                    if 'MonthNum' in df_filtered.columns:
                         month_map = df_filtered[[month_col, 'MonthNum']].drop_duplicates().set_index(month_col)['MonthNum']
                         gd_monthly['MonthNum'] = gd_monthly.index.map(month_map)
                         gd_monthly = gd_monthly.sort_values('MonthNum').drop(columns=['MonthNum'])
                    else:
                         pass # Fallback sort
                         
                    # Calculate Total
                    gd_monthly['Total'] = gd_monthly.sum(axis=1)
                    
                    # Transpose: Rows = Years, Cols = Months
                    gd_monthly_t = gd_monthly.T
                    
                    # Format
                    if curr_y in gd_monthly_t.columns and prev_y in gd_monthly_t.columns:
                        # Ensure nice order
                        cols = sorted(gd_monthly_t.columns, reverse=True) # Descending years
                        gd_monthly_t = gd_monthly_t[cols]
                        
                        # Add Diff Row (Optional but good)
                        gd_monthly_t['Diff'] = gd_monthly_t[curr_y] - gd_monthly_t[prev_y]
                        gd_monthly_t['YoY %'] = (gd_monthly_t['Diff'] / gd_monthly_t[prev_y] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                        
                        # Reorder to show rows: [Prev Year, Curr Year, Diff, YoY]
                        # Actually transposing made Years into Columns. 
                        # Rows are Months.
                        # Wait, user asked for "Table". usually Years as rows is better for simple check, OR Months as columns.
                        # Previous tables were: Rows=Stores, Cols=Months.
                        # Here we have 1 Store. So Rows=Years, Cols=Months works best?
                        # Or Rows=Months, Cols=Years.
                        # Let's check "Store Monthly Breakdown" format: Rows=Store, Cols=Jan..Dec.
                        # Since we have only 1 Store, maybe Rows=Years (2024, 2025) and Cols=Jan..Dec.
                        
                        gd_m_display = gd_monthly.T # Cols are Years. Rows are Months.
                        # We want Rows to be Years.
                        gd_m_display = gd_monthly.T # Wait, unstack gives Years as Columns.
                        # So gd_monthly has Index=Month, Columns=Years.
                        # gd_monthly.T has Index=Years, Columns=Months.
                        
                        # Add Total Column (already added to gd_monthly?)
                        # No, gd_monthly['Total'] sum across Years. That's not what we want. We want sum across months.
                        # Let's redo.
                        
                        # Pivot: Index=Year, Columns=Month
                        gd_daily_view = df_gd_store.groupby([year_col, month_col])[strat_val_col].sum().unstack(fill_value=0)
                         
                        # Sort Columns (Months)
                        if 'MonthNum' in df_filtered.columns:
                            # Create a sorter
                            m_map = df_filtered[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')
                            sorted_ms = [m for m in m_map[month_col] if m in gd_daily_view.columns]
                            gd_daily_view = gd_daily_view[sorted_ms]
                            
                        # Add Total Column
                        gd_daily_view['Total'] = gd_daily_view.sum(axis=1)
                        
                        # Add YoY Row? No, Years are rows.
                        # We can add a Diff Row?
                        # If we have 2024 and 2025.
                        # diff_row = gd_daily_view.loc[curr_y] - gd_daily_view.loc[prev_y]
                        # This works well.
                        
                        if curr_y in gd_daily_view.index and prev_y in gd_daily_view.index:
                             diff_ser = gd_daily_view.loc[curr_y] - gd_daily_view.loc[prev_y]
                             diff_ser.name = 'Diff Amount'
                             
                             yoy_ser = (diff_ser / gd_daily_view.loc[prev_y] * 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                             yoy_ser.name = 'Growth %'
                             
                             gd_daily_view = pd.concat([gd_daily_view, diff_ser.to_frame().T, yoy_ser.to_frame().T])
                        
                        # Format
                        def style_gd(styler):
                             styler.format("{:,.0f}", subset=gd_daily_view.columns) # All numeric
                             # Percentages for Growth % row
                             styler.format("{:+.1f}%", subset=pd.IndexSlice[['Growth %'], :])
                             return styler

                        st.dataframe(
                             gd_daily_view.style.pipe(style_gd).background_gradient(cmap="Blues", subset=pd.IndexSlice[gd_daily_view.index[:-2], gd_daily_view.columns[:-1]], axis=None),
                             use_container_width=True
                        )
                    else:
                         # Simple view if not 2 years
                         gd_daily_view = df_gd_store.groupby([year_col, month_col])[strat_val_col].sum().unstack(fill_value=0)
                         gd_daily_view['Total'] = gd_daily_view.sum(axis=1)
                         st.dataframe(gd_daily_view.style.format("{:,.0f}"), use_container_width=True)
            else:
                 pass
        else:
             st.info(t["warn_year_data_missing"])
        
    st.divider()

    # --- Row 3: Item Ranking (MOVED TO BRANCH SALES SECTION) ---

    # This section has been moved to appear after "Product Monthly Trend" in Branch Sales
    
    st.subheader(t["dist_analysis_header"])
    
    # Dynamic Column Selection for Analysis
    # User wants to choose Header (Group By) and probably the Count Unit
    
    all_cols_dist = sorted([str(c) for c in df.columns])
    # Defaults
    def_row_idx = all_cols_dist.index(item_col) if item_col in all_cols_dist else 0
    def_unit_idx = all_cols_dist.index(branch_col) if branch_col in all_cols_dist else 0
    
    # Prepare Data for Trend Selector (from filtered scope)
    # Drop NaNs to prevent 'nan' option
    available_trends_df = df_filtered[[month_col, 'MonthNum']].dropna(subset=[month_col]).drop_duplicates().sort_values('MonthNum')
    available_trend_months = available_trends_df[month_col].tolist()
    
    # User request: Default to ALL months selected
    default_trends = available_trend_months 


    # Layout: 3 Columns
    c_axis_1, c_axis_2, c_trend = st.columns([1, 1, 2])
    
    with c_axis_1:
        axis_row = st.selectbox(t["select_row_axis"], all_cols_dist, index=def_row_idx, key="dist_axis_row")
    with c_axis_2:
        axis_unit = st.selectbox(t["select_unit_axis"], all_cols_dist, index=def_unit_idx, key="dist_axis_unit")
    with c_trend:
        trend_months = st.multiselect(t["select_trend_months"], options=available_trend_months, default=default_trends, key="dist_trend_months")


    # --- Advanced Item Filter (Always available to filter the DATA context) ---
    # The user might want to analyze "Branches" but only for "Pizza". So this filter remains relevant.
    
    # Popover for selection
    # Only relevant if we have items to filter.
    if item_col and item_col != "(なし)":
        # Exclude NaNs from item list to prevent type issues
        all_dist_items = sorted(df_filtered[item_col].dropna().unique().tolist())
        # Calculate top 20 items by sales for default selection
        top_20_dist = df_filtered.groupby(item_col)[val_col_name].sum().nlargest(20).index.tolist()
        
        if 'dist_selected_items' not in st.session_state:
            st.session_state['dist_selected_items'] = top_20_dist
        
        with st.popover(t["popover_dist_label"], help="Filter Analysis Scope", use_container_width=True):
             st.markdown(t["dist_filter_main_header"])
             c_d1, c_d2 = st.columns(2)
             if c_d1.button(t["btn_select_all"], key="btn_dist_all_sel", use_container_width=True):
                 st.session_state['dist_selected_items'] = all_dist_items
                 st.rerun()
             if c_d2.button(t["btn_clear_all"], key="btn_dist_all_clr", use_container_width=True):
                 st.session_state['dist_selected_items'] = []
                 st.rerun()
                 
             st.divider()
             search_dist = st.text_input("Filter", placeholder=t["dist_search_placeholder"], label_visibility="collapsed", key="search_dist_input")
             
             current_dist_sel = set(st.session_state.get('dist_selected_items', []))
             df_dist_search = pd.DataFrame({
                 'Selected': [i in current_dist_sel for i in all_dist_items],
                 'Item': all_dist_items
             })
             df_dist_search['Selected'] = df_dist_search['Selected'].astype(bool)
             
             if search_dist:
                 df_dist_search = df_dist_search[df_dist_search['Item'].astype(str).str.contains(search_dist, case=False, na=False)]
             
             st.caption(f"{len(df_dist_search)} {t['item_search_count']}")
             
             edited_dist_df = st.data_editor(
                 df_dist_search,
                 column_config={
                     "Selected": st.column_config.CheckboxColumn("Sel.", width="small", default=False),
                     "Item": st.column_config.TextColumn("Item Name", width=300, disabled=True)
                 },
                 disabled=["Item"],
                 hide_index=True,
                 key="dist_item_editor",
                 use_container_width=True
             )
             
             if st.button(t["btn_apply"], key="btn_apply_dist_changes", use_container_width=True, type="primary"):
                 new_dist_selection = set(st.session_state.get('dist_selected_items', []))
                 view_dist_items = df_dist_search['Item'].tolist()
                 new_dist_state_map = dict(zip(edited_dist_df['Item'], edited_dist_df['Selected']))
                 
                 for item in view_dist_items:
                     if new_dist_state_map.get(item, False):
                         new_dist_selection.add(item)
                     else:
                         new_dist_selection.discard(item)
                 
                 st.session_state['dist_selected_items'] = sorted(list(new_dist_selection))
                 if 'dist_item_editor' in st.session_state: del st.session_state['dist_item_editor']
                 st.rerun()


        # --- Apply Context Filter ---
        # If the Axis Row is Item, we filter the ROWS to show using the selection.
        # If Axis Row is NOT Item, we filter the DATA using the selection, then Group By Axis Row.
        
        target_items = st.session_state.get('dist_selected_items', [])
        
        # Filter Data first
        df_target_context = df_filtered.copy()
        if item_col and item_col != "(なし)" and target_items:
             df_target_context = df_target_context[df_target_context[item_col].isin(target_items)]
        elif item_col and item_col != "(なし)" and not target_items:
             st.warning(t["warn_no_item"])
             df_target_context = pd.DataFrame()

        # --- Aggregation Logic ---
        if not df_target_context.empty:
            st.markdown(f"###### Analysis: {axis_row} x {axis_unit} Count")
            
            # --- Monthly Trend Selector (Moved Up) ---
            # Using trend_months from top block
            
            # Decide aggregation approach
            stats_pivot = pd.DataFrame() # Init
            
            try:
                if trend_months:
                    # --- Monthly Trend Mode ---
                    # Filter for selected trend months
                    df_trend = df_target_context[df_target_context[month_col].isin(trend_months)]
                    
                    if not df_trend.empty:
                        # Group by [Row, Month, MonthNum] -> Count Unit only
                        stats_base = df_trend.groupby([axis_row, month_col, 'MonthNum'])[axis_unit].nunique().reset_index()
                        
                        # Pivot: Index=Row, Cols=Month, Values=Count
                        stats_pivot = stats_base.pivot(index=axis_row, columns=month_col, values=axis_unit).fillna(0)
                        
                        # Sort Columns by MonthNum
                        # Get valid month names present in pivot
                        valid_cols = [c for c in stats_pivot.columns]
                        # Sort valid_cols based on MonthNum order from available_trends_df
                        # Create a mapping
                        month_map = dict(zip(available_trends_df[month_col], available_trends_df['MonthNum']))
                        sorted_cols = sorted(valid_cols, key=lambda x: month_map.get(x, 0))
                        
                        stats_pivot = stats_pivot[sorted_cols]
                        
                        # Add Total Sales for sorting (optional, user wanted Count only but sorting context helps)
                        # User requested NO Sales columns. Just Count.
                        # Maybe add "Total Count" for sorting rows?
                        stats_pivot['Total Count'] = stats_pivot.sum(axis=1)
                        stats_pivot = stats_pivot.sort_values('Total Count', ascending=False)
                        
                        # Rename columns for clarity? No, Month names are fine.
                        
                        st.dataframe(stats_pivot.style.format("{:,.0f}"), use_container_width=True)
                        
                        # --- New Growth Analysis (Start vs End Month) ---
                        if len(sorted_cols) >= 2:
                            end_m = sorted_cols[-1]
                            prior_months = sorted_cols[:-1] # All months except the last one
                            
                            start_m_str = prior_months[0]
                            last_prior_str = prior_months[-1]
                            range_str = f"{start_m_str} ～ {last_prior_str}" if len(prior_months) > 1 else start_m_str
                            
                            
                            st.write(f"**{t['new_growth_analysis_header']}: {end_m} [Strict Mode]**")
                            st.caption(t["new_growth_caption"])
                            
                            # Get Unique Units for Prior Period and End Month
                            # using df_trend which is already filtered by Item selection
                            users_prior = set(df_trend[df_trend[month_col].isin(prior_months)][axis_unit].unique())
                            users_end = set(df_trend[df_trend[month_col] == end_m][axis_unit].unique())
                            
                            new_users = list(users_end - users_prior)
                            
                            # Display Total Count prominently but compactly as requested
                            st.markdown(f"#### {t['new_growth_count_label']}: {len(new_users):,} ")
                            
                            if new_users:
                                # Get details for these new users in the End Month
                                df_new_growth = df_trend[
                                    (df_trend[month_col] == end_m) & 
                                    (df_trend[axis_unit].isin(new_users))
                                ]
                                
                                # Aggregate to show summary
                                # Group by User -> List Items, Sum Sales
                                agg_dict = {
                                    val_col_name: 'sum',
                                    item_col: lambda x: ", ".join(sorted(x.unique())[:5]) # Show top 5 items
                                }
                                
                                # Identify Business Division Column (Robust Search)
                                business_div_col = None # No fallback to branch_col as per user request
                                potential_div_names = ["Business Division", "business division", "Division", "division", "Channel", "channel"]
                                found_div_col = None
                                
                                # Check for explicit column name in dataset
                                for col in df_new_growth.columns:
                                    if col in potential_div_names:
                                        found_div_col = col
                                        break
                                
                                if found_div_col:
                                     business_div_col = found_div_col
                                
                                has_div = business_div_col and business_div_col != "(なし)" and business_div_col in df_new_growth.columns
                                
                                if has_div:
                                    agg_dict[business_div_col] = lambda x: x.mode()[0] if not x.mode().empty else x.iloc[0] # Most frequent or first
                                
                                growth_summary = df_new_growth.groupby(axis_unit).agg(agg_dict).reset_index()
                                
                                # Rename Columns
                                sales_label = f'Sales ({end_m})'
                                # create explicit label
                                full_sales_label = f"{sales_label} [{val_col_name}]"
                                
                                rename_map = {
                                    val_col_name: full_sales_label,
                                    item_col: 'Purchased Items (Sample)'
                                }
                                if has_div:
                                    rename_map[business_div_col] = 'Business Division'
                                    
                                growth_summary = growth_summary.rename(columns=rename_map)
                                
                                # Reorder Columns
                                cols_order = [axis_unit, full_sales_label, 'Purchased Items (Sample)']
                                if has_div:
                                    cols_order = ['Business Division'] + cols_order
                                    
                                growth_summary = growth_summary[cols_order].sort_values(full_sales_label, ascending=False)
                                
                                # Display
                                st.dataframe(
                                    growth_summary.style.format({full_sales_label: "{:,.0f}"}),
                                    use_container_width=True,
                                    height=400 # Scrollable fixed height (approx 10-12 rows)
                                )
                            else:
                                st.info(f"{t['info_no_new_growth']} ({start_m_str} -> {end_m})")
                
                elif len(selected_years) > 1:
                    # YoY Comparison Mode (Existing Logic)
                    stats_base = df_target_context.groupby([axis_row, year_col]).agg({
                        axis_unit: 'nunique',
                        val_col_name: 'sum'
                    }).reset_index()
                    
                    # Pivot: Index=Row, Cols=Year, Values=[Count, Sales]
                    stats_pivot = stats_base.pivot(index=axis_row, columns=year_col, values=[axis_unit, val_col_name])
                    
                    # Flatten MultiIndex Columns
                    new_cols = []
                    for level0, level1 in stats_pivot.columns:
                        metric_name = "Count" if level0 == axis_unit else "Sales"
                        new_cols.append(f"{metric_name} {level1}")
                    
                    stats_pivot.columns = new_cols
                    
                    # Add Total Sales for sorting
                    sales_cols = [c for c in new_cols if "Sales" in c]
                    stats_pivot['Total Sales'] = stats_pivot[sales_cols].sum(axis=1)
                    
                    # Sort
                    stats_final = stats_pivot.sort_values('Total Sales', ascending=False)
                    
                    # Limit
                    if len(stats_final) > 50:
                         st.caption(f"※ Top 50 / {len(stats_final)} rows shown")
                         stats_final = stats_final.head(50)

                    st.dataframe(stats_final.style.format("{:,.0f}"), use_container_width=True)
                    
                else:
                    # Single Year / Standard Mode (No Pivot Needed)
                    stats = df_target_context.groupby(axis_row).agg({
                        axis_unit: 'nunique',
                        val_col_name: 'sum'
                    }).reset_index()
                    
                    stats.columns = [axis_row, f'{axis_unit} Count', 'Total Sales']
                    
                    # Handling Division by Zero for Avg
                    stats['Avg Sales'] = stats.apply(
                        lambda x: x['Total Sales'] / x[f'{axis_unit} Count'] if x[f'{axis_unit} Count'] > 0 else 0, axis=1
                    )
                    
                    # Sort by Sales
                    stats = stats.sort_values('Total Sales', ascending=False)
                    
                    # Limit
                    if len(stats) > 50:
                        st.caption(f"※ Top 50 / {len(stats)} rows shown")
                        stats = stats.head(50)
                    
                    st.dataframe(
                        stats.style
                        .format({"Total Sales": "{:,.0f}", f"{axis_unit} Count": "{:,.0f}", "Avg Sales": "{:,.0f}"})
                        .bar(subset=[f'{axis_unit} Count'], color='#d65f5f'),
                        use_container_width=True
                    )
                
            except Exception as e:
                st.error(f"{t['error_agg_type']}\n{e}")

        st.divider()
        
        # --- AI Strategy Insight Section ---
        st.markdown("### 🤖 AI Strategy Insight & Consultant Chat")
        
        # Prepare Data Summary for AI
        # Re-calculate or retrieve key metrics safe for this scope
        
        # 1. Total Sales (Current View)
        current_total_sales = df_filtered[val_col_name].sum()
        
        # 2. YoY Growth (Overall)
        # Helper to get value matching index type (int vs str)
        def safe_get_year_val(series, key):
             if key is None: return 0
             if key in series.index: return series[key]
             if str(key) in series.index: return series[str(key)]
             return 0

        # We can use the sales_by_year series calculated above
        # Note: sales_by_year was removed from main KPI section, so calculate it locally here
        sales_by_year = df_filtered.groupby(year_col)[val_col_name].sum().sort_index()
        
        # Assuming current_year and prev_year are still valid in scope (they are top-level in this filtering block)
        current_y_sales = safe_get_year_val(sales_by_year, current_year)
        prev_y_sales = safe_get_year_val(sales_by_year, prev_year)
        
        yoy_text = "N/A"
        if prev_y_sales > 0:
            growth_r = ((current_y_sales - prev_y_sales) / prev_y_sales) * 100
            yoy_text = f"{growth_r:+.1f}%"
            
        kpi_text = f"Total Sales: {current_total_sales:,.0f} (YoY: {yoy_text})"
        
        # 3. Forecast
        # Re-calc forecast just to be safe
        forecast_disp = "N/A"
        if current_year:
             curr_sales_part = df_filtered[df_filtered[year_col] == current_year][val_col_name].sum()
             months_c = df_filtered[df_filtered[year_col] == current_year][month_col].nunique()
             if months_c > 0:
                 proj = (curr_sales_part / months_c) * 12
                 forecast_disp = f"{proj:,.0f}"
                 
        forecast_text = f"Year-End Forecast: {forecast_disp}"
        
        forecast_text = f"Year-End Forecast: {forecast_disp}"
        
        # --- AI Settings & Token Control ---
        # User requested Full Data access. 
        # Removed Slider limitation logic.
        ai_max_rows = None # No limit logic, but we still define var if needed or just skip
        
        # --- [CRITICAL] Prepare Raw Data Context for AI ---
        # User explicitly requested "Gemini to see Raw Data" and "Remove Limit".
        # We will pass a CSV sample of the FILTERED data (ALL ROWS).
        
        # Select relevant columns for context to save tokens/noise
        ctx_cols = [c for c in df_filtered.columns if c in [month_col, item_col, category_col, branch_col, val_col_name, year_col, target_cust_col]]
        # If specific columns aren't found, fallback to all (or reasonable subset)
        if not ctx_cols:
             ctx_cols = df_filtered.columns.tolist()
        
        # Full Data usage
        df_ctx_safe = df_filtered[ctx_cols]
        # No head/truncation
        # if len(df_ctx_safe) > ai_max_rows: ... (REMOVED)
             
        raw_data_csv = df_ctx_safe.to_csv(index=False)

        data_summary_for_ai = f"""
        【分析対象データ (Raw Data / CSV format)】
        以下は現在のダッシュボードでフィルタリングされた実際の売上データ（最大{ai_max_rows}行）です。
        このデータを分析し、傾向や異常値を読み取ってください。
        
        Shape: {df_ctx_safe.shape}
        Columns: {list(df_ctx_safe.columns)}

        ```csv
        {raw_data_csv}
        ```
        
        【KPI概要】
        {kpi_text}
        {forecast_text}
        """
        
        # 1. API Key Check & AI Engine Initialization
        if "ai_sales_insight" not in st.session_state:
            st.session_state["ai_sales_insight"] = None
        
        ai_ready = False
        
        # Check if engine exists
        if 'ai_engine' in st.session_state and st.session_state['ai_engine'] is not None:
             ai_ready = True
        else:
             # Try to init with environment key
             try:
                 from ai_engine import AIEngine
                 st.session_state['ai_engine'] = AIEngine()
                 ai_ready = True
             except ValueError:
                 ai_ready = False
        
        # Display Logic
        if not ai_ready:
             st.warning(t["warn_ai_api_needed"])
             api_key_input = st.text_input("🔑 Google Gemini API Key", type="password", help="Enter your Gemini API Key here.", key="db_api_key_input_persistent")
             
             if api_key_input:
                 try:
                     from ai_engine import AIEngine
                     st.session_state['ai_engine'] = AIEngine(api_key=api_key_input)
                     st.success(t["success_ai_connected"])
                     st.rerun()
                 except Exception as e:
                     st.error(f"{t['error_ai_connect']}{e}")
        
        else:
            # AI IS READY -> Show Features
            
            # Feature 1: Insight Button
            if st.button("✨ Generate AI Insight (Click to Analyze)", key="btn_gen_insight", type="primary"):
                with st.spinner("AI is analyzing your sales data..."):
                    insight = st.session_state['ai_engine'].generate_sales_insight(data_summary_for_ai)
                    st.session_state["ai_sales_insight"] = insight
            
            if st.session_state["ai_sales_insight"]:
                st.info(st.session_state["ai_sales_insight"], icon="💡")
            
            # Feature 2: Consultant Chat Interface
            st.markdown("#### 💬 Ask the AI Consultant")
            
            # Chat History
            if "chat_history_sales" not in st.session_state:
                st.session_state["chat_history_sales"] = []
                
            # Display History
            for msg in st.session_state["chat_history_sales"]:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                    
            # Chat Input
            if prompt := st.chat_input(t["ai_chat_placeholder"]):
                # Display User Message
                st.session_state["chat_history_sales"].append({"role": "user", "content": prompt})
                with st.chat_message("user"):
                    st.markdown(prompt)
                    
                # Generate Response
                with st.chat_message("assistant"):
                    with st.spinner("Thinking..."):
                        # Contextualize with Dashboard Data
                        response_text = st.session_state['ai_engine'].get_strategy_advice(
                            user_query=prompt,
                            internal_data_summary=data_summary_for_ai,
                            external_trends="General Retail Trends (2025)", # Simplified for now
                            history=st.session_state["chat_history_sales"][:-1]
                        )
                        
                        st.markdown(response_text)
                        st.session_state["chat_history_sales"].append({"role": "assistant", "content": response_text})



else:
    st.info(t["assign_columns_info"])
    st.write(t["current_preview"])
    st.dataframe(df.head())
