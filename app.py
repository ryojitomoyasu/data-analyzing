import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io


# ローカル環境でだけ .env を読む
if os.path.exists(".env"):
    from dotenv import load_dotenv
    load_dotenv()
    
API_KEY = os.getenv("API_KEY")
if not API_KEY:
    st.error("API_KEY is missing. Please set it in Secrets / environment variables.")
    st.stop()

from data_processor import clean_data, calculate_kpis, infer_industry, calculate_abc_analysis
from ai_engine import AIEngine

# Page Config
st.set_page_config(page_title="BizStrategy AI Partner", layout="wide", page_icon="🚀")

# Premium CSS for Glassmorphism and Animations
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }

    /* Glassmorphism for containers */
    .stMetric, .report-box, .stChatFloatingInputContainer, [data-testid="stExpander"] {
        background: rgba(255, 255, 255, 0.7) !important;
        backdrop-filter: blur(10px);
        border-radius: 15px !important;
        border: 1px solid rgba(255, 255, 255, 0.3);
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.07);
        padding: 20px !important;
        transition: transform 0.3s ease;
    }
    
    .stMetric:hover {
        transform: translateY(-5px);
    }

    /* Custom Header */
    .main-title {
        font-size: 3rem;
        font-weight: 800;
        background: linear-gradient(to right, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }

    .sub-title {
        color: #4b5563;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }

    /* Sidebar glassmorphism */
    [data-testid="stSidebar"] {
        background-color: rgba(255, 255, 255, 0.4) !important;
        backdrop-filter: blur(15px);
        border-right: 1px solid rgba(255, 255, 255, 0.2);
    }

    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .stPlotlyChart {
        animation: fadeIn 0.8s ease-out;
    }
    
    .report-box {
        border-left: 5px solid #3b82f6 !important;
        animation: fadeIn 0.5s ease-out;
    }

    /* Chat Bubbles */
    .chat-bubble {
        padding: 12px 18px;
        border-radius: 20px;
        margin-bottom: 10px;
        max-width: 80%;
        line-height: 1.5;
        font-size: 0.95rem;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    .user-bubble {
        background: #3b82f6;
        color: white;
        margin-left: auto;
        border-bottom-right-radius: 2px;
    }
    .assistant-bubble {
        background: white;
        color: #1f2937;
        margin-right: auto;
        border-bottom-left-radius: 2px;
        border: 1px solid #e5e7eb;
    }

    /* Action Buttons */
    .stDownloadButton button {
        border-radius: 10px !important;
        background: #ffffff !important;
        color: #3b82f6 !important;
        border: 1px solid #3b82f6 !important;
        transition: all 0.3s ease;
    }
    .stDownloadButton button:hover {
        background: #3b82f6 !important;
        color: #ffffff !important;
    }

    </style>
    """, unsafe_allow_html=True)

# Load Env
load_dotenv()

# Session State Initialization
if "ai_engine" not in st.session_state:
    st.session_state.ai_engine = None
if "external_report" not in st.session_state:
    st.session_state.external_report = ""
if "messages" not in st.session_state:
    st.session_state.messages = []
if "df" not in st.session_state:
    st.session_state.df = None

# Sidebar
with st.sidebar:
    st.title("⚙️ 設定")
    
    # API Key Input
    api_key_input = st.text_input("Gemini API Key", type="password", help="`.env` に設定がない場合、こちらに入力してください。")
    
    # Init AI Engine
    try:
        current_key = api_key_input or os.getenv("GOOGLE_API_KEY")
        if current_key and (st.session_state.ai_engine is None or api_key_input):
            st.session_state.ai_engine = AIEngine(api_key=current_key)
            st.success("✅ AI Engine Ready (Gemini 2.0 Flash)")
    except Exception as e:
        st.error(f"AI Engine Error: {e}")

    st.divider()
    
    # File Uploader
    st.subheader("📊 データの読み込み")
    uploaded_file = st.file_uploader("Excel or CSV", type=["xlsx", "csv"])
    
    if st.button("💡 サンプルデータをロード"):
        if os.path.exists("dummy_sales.csv"):
            st.session_state.df = pd.read_csv("dummy_sales.csv")
            st.session_state.df = clean_data(st.session_state.df)
            st.toast("サンプルデータを読み込みました")
            st.rerun()
        else:
            st.error("サンプルデータが見つかりません。")

# Main Content
st.markdown('<h1 class="main-title">🚀 BizStrategy AI Partner</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Gemini 2.0 Flash とデータで、次の一手を。</p>', unsafe_allow_html=True)

# Use session state df if available
if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            st.session_state.df = pd.read_csv(uploaded_file)
        else:
            st.session_state.df = pd.read_excel(uploaded_file)
        st.session_state.df = clean_data(st.session_state.df)
    except Exception as e:
        st.error(f"エラー: {e}")

df = st.session_state.df

if df is not None:
    st.success(f"データ読み込み完了: {len(df)} 件")
    
    # Display Data (Optional/Truncated)
    with st.expander("🔍 データプレビュー"):
        st.dataframe(df.head(10), use_container_width=True)

    # KPI Dashboard
    kpis = calculate_kpis(df)
    if kpis:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📊 売上総額", f"{kpis['total_sales']:,.0f} 円")
        c2.metric("💰 平均単価", f"{kpis['avg_order_value']:,.0f} 円")
        c3.metric("📝 レコード数", f"{kpis['record_count']:,.0f} 件")
        if "monthly_growth" in kpis:
            growth = kpis['monthly_growth']
            c4.metric("📈 前月比", f"{growth:.1f} %", delta=f"{growth:.1f}%")

    # Visualization
    st.divider()
    st.subheader("📈 トレンド・構成分析")
    
    v_col1, v_col2 = st.columns(2)
    
    sales_col = next((col for col in df.columns if '売上' in col or 'sales' in col.lower()), None)
    date_col = next((col for col in df.columns if '日付' in col or 'date' in col.lower()), None)
    cat_col = next((col for col in df.columns if 'カテゴリ' in col or 'category' in col.lower() or '商品' in col or 'product' in col.lower()), None)

    if date_col and sales_col:
        df_trend = df.sort_values(date_col)
        fig_trend = px.line(df_trend, x=date_col, y=sales_col, title="売上推移", template="plotly_white")
        fig_trend.update_traces(line_color='#3b82f6', line_width=3)
        v_col1.plotly_chart(fig_trend, use_container_width=True)

    if cat_col and sales_col:
        df_cat = df.groupby(cat_col)[sales_col].sum().reset_index()
        fig_pie = px.pie(df_cat, values=sales_col, names=cat_col, title="構成比分析", hole=0.4, template="plotly_white")
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        v_col2.plotly_chart(fig_pie, use_container_width=True)


    # ABC Analysis Visualization
    st.divider()
    st.subheader("📊 重点分析 (ABC分析 / パレート図)")
    abc_df = calculate_abc_analysis(df)
    if abc_df is not None:
        c_p1, c_p2 = st.columns([2, 1])
        
        # Pareto Chart
        fig_abc = px.bar(abc_df, x=cat_col, y=sales_col, color='rank', 
                         color_discrete_map={'A':'#ef4444', 'B':'#f59e0b', 'C':'#10b981'},
                         title="商品/カテゴリ別売上ランキング (ABCランク)")
        
        # Add line for cumulative percentage
        import plotly.graph_objects as go
        fig_abc.add_trace(go.Scatter(x=abc_df[cat_col], y=abc_df['cumulative_percent'], 
                                     name='累積構成比', yaxis='y2', line=dict(color='#1e3a8a', width=3)))
        
        fig_abc.update_layout(
            yaxis2=dict(title='累積構成比 (%)', overlaying='y', side='right', range=[0, 105]),
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        c_p1.plotly_chart(fig_abc, use_container_width=True)
        
        with c_p2:
            st.markdown("### ランク別概況")
            a_items = len(abc_df[abc_df['rank']=='A'])
            b_items = len(abc_df[abc_df['rank']=='B'])
            st.info(f"**ランクA (最重点)**: {a_items} 品目\n\n売上の70%を占める主力商品です。在庫切れ厳禁、販売強化の優先順位が最も高い領域です。")
            st.warning(f"**ランクB (準重点)**: {b_items} 品目\n\n売上の20%を占めます。Aへの昇格、あるいは維持のための管理が必要です。")

    # Brand Pivot Analysis
    st.divider()
    st.subheader("📈 ブランド別・月別売上推移 (Pivot Analysis)")
    
    # Identify Columns based on User's Data Structure
    # "Brand", "Year", "Month", "Value" are the expected columns from the screenshot
    brand_col = next((col for col in df.columns if 'Brand' in str(col) or 'ブランド' in str(col)), None)
    year_col = next((col for col in df.columns if 'Year' in str(col) or '年' in str(col)), None)
    month_col = next((col for col in df.columns if 'Month' in str(col) or '月' in str(col)), None)
    val_col_name = next((col for col in df.columns if 'Value' in str(col) or '売上' in str(col) or 'Sales' in str(col)), None)

    if brand_col and year_col and month_col and val_col_name:
        # Create a copy for manipulation
        df_p = df.copy()
        
        # Month Sort Order Mapping (Handle Apr, May strings etc.)
        month_order = {
            'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
            'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
            'January': 1, 'February': 2, 'March': 3, 'April': 4, 'June': 6,
            'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        
        # Ensure Value is numeric
        df_p[val_col_name] = pd.to_numeric(df_p[val_col_name], errors='coerce').fillna(0)
        
        # Add Month Number for sorting if Month is string
        if df_p[month_col].dtype == 'object':
            # Try mapping English month names
            df_p['MonthNum'] = df_p[month_col].map(month_order)
            # If map failed (maybe Numbers as strings or Japanese), try numeric conversion
            if df_p['MonthNum'].isnull().all():
                 df_p['MonthNum'] = pd.to_numeric(df_p[month_col], errors='coerce')
        else:
            df_p['MonthNum'] = df_p[month_col]

        # Fill missing month nums with 0 to allow sorting
        df_p['MonthNum'] = df_p['MonthNum'].fillna(0).astype(int)

        # Create Pivot Table
        # Request: Column: Brand, Year  Row: Month, Value: Value
        try:
            pivot_table = df_p.pivot_table(
                index=[month_col], 
                columns=[brand_col, year_col], 
                values=val_col_name, 
                aggfunc='sum',
                fill_value=0
            )
            
            # Sort by Month Num
            # To do this cleanly with the pivot table having Month as index name
            # We sort the source first or reindex the pivot
            unique_months = df_p[[month_col, 'MonthNum']].drop_duplicates().sort_values('MonthNum')
            sorted_index = unique_months[month_col].tolist()
            # Filter to only months actually in the pivot index
            sorted_index = [m for m in sorted_index if m in pivot_table.index]
            pivot_table = pivot_table.reindex(sorted_index)

            # Visualization
            tab1, tab2 = st.tabs(["📋 Pivot Table (詳細)", "📈 推移グラフ"])
            
            with tab1:
                st.write("#### ブランド・年別 月次売上表")
                # Format numbers with commas
                st.dataframe(pivot_table.style.format("{:,.0f}"), use_container_width=True)
                
                csv_p = pivot_table.to_csv().encode('utf-8-sig')
                st.download_button("📥 PivotデータをCSVでダウンロード", csv_p, "custom_pivot.csv", "text/csv")

            with tab2:
                # For plotting, simpler flat structure is often better
                # Group by Year, Month, Brand
                df_grouped = df_p.groupby([year_col, month_col, 'MonthNum', brand_col])[val_col_name].sum().reset_index()
                df_grouped = df_grouped.sort_values(['Year', 'MonthNum'])
                
                # Create a composite label for X-axis or Color
                df_grouped['Period'] = df_grouped[year_col].astype(str) + '-' + df_grouped[month_col].astype(str)
                
                fig_custom = px.line(
                    df_grouped,
                    x=month_col,
                    y=val_col_name,
                    color=brand_col,
                    line_dash=year_col, # Differentiate years by dash style
                    title="ブランド別 月次売上推移 (年比較)",
                    markers=True,
                    template="plotly_white"
                )
                # Ensure x-axis is sorted correctly
                fig_custom.update_xaxes(categoryorder='array', categoryarray=sorted_index)
                
                st.plotly_chart(fig_custom, use_container_width=True)

        except Exception as e:
            st.error(f"Pivot Table作成エラー: {e}")
            st.write("データ形式の確認: ")
            st.write(df_p.head())

    else:
        st.warning("必要な列（Brand, Year, Month, Value）が完全に見つかりませんでした。列名を確認してください。")

    # External Trend Analyzer
    st.divider()
    st.subheader("🌍 業界トレンド分析 (Gemini 2.0 Flash Search)")
    
    # Report Export Section
    with st.sidebar:
        st.divider()
        st.subheader("📄 レポート出力")
        if st.session_state.df is not None:
            # Generate Report Text
            report_text = f"# BizStrategy AI Partner 分析レポート\n\n"
            report_text += f"## 1. 内部データサマリー\n"
            report_text += f"- 売上総額: {kpis.get('total_sales', 0):,.0f}円\n"
            report_text += f"- レコード数: {kpis.get('record_count', 0)}件\n"
            if "monthly_growth" in kpis:
                report_text += f"- 前月比成長率: {kpis['monthly_growth']:.2f}%\n"
            
            if st.session_state.external_report:
                report_text += f"\n## 2. 外部環境トレンド\n{st.session_state.external_report}\n"
            
            if st.session_state.messages:
                report_text += f"\n## 3. 戦略チャット履歴\n"
                for m in st.session_state.messages:
                    role = "ユーザー" if m["role"] == "user" else "AIコンサルタント"
                    report_text += f"**{role}**: {m['content']}\n\n"
            
            st.download_button(
                label="📥 分析レポート(MD)を保存",
                data=report_text,
                file_name="biz_strategy_report.md",
                mime="text/markdown"
            )

    if st.button("🔍 最新トレンドを検索・分析する"):
        if st.session_state.ai_engine:
            with st.spinner("Google Search を利用して業界動向を調査中..."):
                # Use the new smart industry inference
                sample_cats = df[cat_col].dropna().unique()[:5] if cat_col else ["一般"]
                industry = st.session_state.ai_engine.infer_industry_smart(", ".join(map(str, sample_cats)))
                st.info(f"推論された業界: {industry}")
                
                report = st.session_state.ai_engine.get_industry_trends(industry)
                st.session_state.external_report = report
        else:
            st.warning("APIキーを設定してください。")

    if st.session_state.external_report:
        st.markdown(f'<div class="report-box"><strong>外部環境分析レポート:</strong><br><br>{st.session_state.external_report}</div>', unsafe_allow_html=True)

    # Strategy Chat
    st.divider()
    st.subheader("💬 戦略壁打ちチャット (Context Aware)")
    
    # Scrollable chat area
    chat_container = st.container()
    with chat_container:
        for message in st.session_state.messages:
            role_class = "user-bubble" if message["role"] == "user" else "assistant-bubble"
            st.markdown(f'<div class="chat-bubble {role_class}">{message["content"]}</div>', unsafe_allow_html=True)

    if prompt := st.chat_input("現在のデータに基づいた戦略を相談する"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.markdown(f'<div class="chat-bubble user-bubble">{prompt}</div>', unsafe_allow_html=True)

        if st.session_state.ai_engine:
            with chat_container:
                # Placeholder for assistant response animation
                with st.spinner("思考中..."):
                    # Prepare data summary
                    data_summary = f"売上総額: {kpis.get('total_sales', 0)}, 平均単価: {kpis.get('avg_order_value', 0)}, レコード数: {kpis.get('record_count', 0)}"
                    if "monthly_growth" in kpis:
                        data_summary += f", 前月比: {kpis['monthly_growth']:.2f}%"
                    
                    ext_report = st.session_state.external_report or "未取得（最新トレンドは検索されていません）"
                    
                    # Get response with history
                    history = []
                    for m in st.session_state.messages[:-1]:
                        role = "user" if m["role"] == "user" else "model"
                        history.append({"role": role, "parts": [m["content"]]})
                    
                    response = st.session_state.ai_engine.get_strategy_advice(prompt, data_summary, ext_report, history=history)
                    st.markdown(f'<div class="chat-bubble assistant-bubble">{response}</div>', unsafe_allow_html=True)
                    st.session_state.messages.append({"role": "assistant", "content": response})
        else:
            st.error("APIキーが必要です。")

else:
    st.info("👈 サイドバーから csv または xlsx ファイルをアップロードするか、サンプルデータをロードしてください。")
