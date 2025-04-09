import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
import datetime
import sys
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

# ページ設定を最初に行う
st.set_page_config(layout="wide", page_title="エスコ売上実績分析ダッシュボード")

# 事業部ごとの色を設定（グローバル変数として定義）
DEPT_COLORS = {
    '1.在来': '#FFD700',  # 濃い黄
    '2.ＳＯ': '#FF8C00',  # 濃いオレンジ
    '3.ＳＳ': '#43A047',  # 緑
    '4.教材': '#FF0000',  # 赤
    '5.ス介': '#1E88E5'   # 青
}

def check_password():
    """パスワード認証を行う関数"""
    def password_entered():
        if st.session_state["password"] == "esco2025":  # パスワードを設定
            st.session_state["password_correct"] = True
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("""
        <style>
        .password-container {
            max-width: 400px;
            margin: 100px auto;
            padding: 20px;
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        </style>
        """, unsafe_allow_html=True)
        
        with st.container():
            st.markdown('<div class="password-container">', unsafe_allow_html=True)
            st.markdown("### 🔐 パスワード認証")
            st.text_input(
                "パスワードを入力してください",
                type="password",
                on_change=password_entered,
                key="password"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        return False
    
    return st.session_state["password_correct"]

# パスワード認証チェック
if not check_password():
    st.stop()

# サイドバーの設定
st.sidebar.title("メニュー")
st.sidebar.markdown("""
### 📊 分析メニュー

1. [全体分析](/)
   - 全体の売上推移
   - 部門別分析
   - グループ別分析

2. [商品別分析](/商品別分析)
   - 商品ごとの売上分析
   - カテゴリー別分析
   - 新規採用商品の分析
""")

st.sidebar.markdown("---")
st.sidebar.info("© 2025 商品本部PS課")

# メインページのコンテンツ
st.title("エスコ売上実績分析ダッシュボード")

# Add custom CSS for better styling
st.markdown("""
<style>
    /* モダンなカラーパレット */
    :root {
        --primary: #1E88E5;
        --primary-light: #e3f2fd;
        --secondary: #43A047;
        --accent: #FF5722;
        --background: #ffffff;
        --surface: #f8f9fa;
        --text: #2c3e50;
        --text-light: #6c757d;
    }

    /* メインヘッダーのデザイン改善 */
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: var(--primary);
        text-align: center;
        margin: 2rem 0;
        padding: 2rem;
        background: linear-gradient(135deg, var(--primary-light) 0%, var(--surface) 100%);
        border-radius: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border: 1px solid rgba(30, 136, 229, 0.1);
    }

    /* サブヘッダーのデザイン改善 */
    .sub-header {
        font-size: 1.8rem;
        font-weight: 600;
        color: var(--text);
        margin: 2rem 0 1.5rem;
        padding: 1rem 1.5rem;
        background: var(--surface);
        border-radius: 15px;
        border-left: 5px solid var(--primary);
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }

    /* メトリックカードのデザイン改善 */
    .metric-card {
        background: var(--background);
        border-radius: 15px;
        padding: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.1);
    }

    /* アップロードセクションのデザイン改善 */
    .upload-section {
        background: linear-gradient(135deg, var(--primary-light) 0%, #ffffff 100%);
        padding: 2.5rem;
        border-radius: 20px;
        margin: 2rem 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid rgba(30, 136, 229, 0.1);
    }

    /* タブのデザイン改善 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        background: var(--surface);
        padding: 1rem;
        border-radius: 15px;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    .stTabs [data-baseweb="tab"] {
        height: auto;
        padding: 0.75rem 1.5rem;
        background: var(--background);
        border-radius: 10px;
        border: 1px solid rgba(0, 0, 0, 0.05);
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        background: var(--primary) !important;
        color: white !important;
        box-shadow: 0 2px 4px rgba(30, 136, 229, 0.2);
    }

    /* メトリック値のデザイン改善 */
    div[data-testid="stMetricValue"] {
        font-size: 2.2rem;
        font-weight: 700;
        color: var(--primary);
        background: linear-gradient(135deg, var(--primary) 0%, #64B5F6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 1.2rem;
        font-weight: 500;
        color: var(--text-light);
    }

    /* ラジオボタンのスタイル */
    .stRadio > div {
        display: flex !important;
        flex-direction: row !important;
        align-items: center !important;
        gap: 1rem !important;
    }
    
    .stRadio > div[role="radiogroup"] > label {
        padding: 0.5rem 1rem !important;
        background-color: white !important;
        border: 1px solid #e0e0e0 !important;
        border-radius: 4px !important;
        cursor: pointer !important;
        color: #2c3e50 !important;
        font-weight: 500 !important;
        font-size: 1rem !important;
    }
    
    .stRadio > div[role="radiogroup"] > label:hover {
        background-color: #f8f9fa !important;
    }
    
    .stRadio > div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }
    
    .stRadio > div[role="radiogroup"] > label > div:last-child {
        margin: 0 !important;
    }

    /* データフレームのデザイン改善 */
    .stDataFrame {
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }

    /* セレクトボックスのデザイン改善 */
    .stSelectbox {
        border-radius: 10px;
    }
    .stSelectbox > div > div {
        background: var(--background);
        border-radius: 10px;
        border: 1px solid rgba(0, 0, 0, 0.1);
    }

    /* ダウンロードボタンのデザイン改善 */
    .stDownloadButton {
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        background: var(--primary);
        color: white;
        border: none;
        box-shadow: 0 2px 4px rgba(30, 136, 229, 0.2);
        transition: all 0.2s ease;
    }
    .stDownloadButton:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 6px rgba(30, 136, 229, 0.3);
    }

    /* Add creator info style */
    .creator-info {
        text-align: right;
        color: var(--text-light);
        font-size: 0.9rem;
        margin-top: 0.5rem;
        padding-right: 1rem;
        font-style: italic;
    }
</style>
""", unsafe_allow_html=True)

# App title and description
st.markdown('<div class="main-header">📊 エスコ売上分析ダッシュボード</div>', unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; padding: 1rem; margin-bottom: 2rem;">
このダッシュボードでは、売上データを分析し、部門ごとのパフォーマンスを視覚化します。<br>
年次・月次の売上、総差、総差率などの主要指標を確認できます。
</div>
<div style="text-align: right; color: var(--text-light); font-size: 0.9rem; font-style: italic; margin-bottom: 2rem;">
    Created by 商品戦略部PS課
</div>
""", unsafe_allow_html=True)

# File upload section
st.markdown('<div class="upload-section">', unsafe_allow_html=True)
st.markdown("""
<h2 style="color: #1E88E5; margin-bottom: 1rem;">📁 データのアップロード</h2>
<p style="margin-bottom: 1rem;">Excelファイルをアップロードしてください</p>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["xlsx", "xls", "csv"])
st.markdown('</div>', unsafe_allow_html=True)

# Function to process data
def process_data(df):
    # Clean and prepare data
    df.columns = ['年月度', 'N事業名', '売上', '総差']
    
    # Extract year and month from 年月度
    df['年'] = df['年月度'].astype(str).str[:4]
    df['月'] = df['年月度'].astype(str).str[4:6]
    
    # Calculate 総差率 (margin rate)
    df['総差率'] = (df['総差'] / df['売上'] * 100).round(2)
    
    return df

# Function to demonstrate with sample data if no file is uploaded
def load_sample_data():
    # Create sample data similar to the image
    data = {
        '年月度': [201601, 201602, 201602, 201602, 201603, 201603, 201603, 201604, 201605, 201605, 201606],
        'N事業名': ['1.在来', '1.在来', '2.SO', '3.SS', '1.在来', '2.SO', '3.SS', '1.在来', '1.在来', '2.SO', '1.在来'],
        '売上': [58697, 3430096, 24967, 2620, 7718692, 19837, 4324, 1570069, 1854943, 6217, 3698368],
        '総差': [9737, 250726, 4224, 386, 723878, 3388, 744, 221050, 241806, 1068, 490729]
    }
    return pd.DataFrame(data)

# Demo data or uploaded data
if uploaded_file is not None:
    try:
        # ファイル名のチェック
        if "エスコ" not in uploaded_file.name or "衞藤" not in uploaded_file.name:
            st.error("承認されていないファイル形式です")
            df = load_sample_data()
            df = process_data(df)
        else:
            # Determine file type and read accordingly
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            df = process_data(df)
            st.success("データが正常に読み込まれました！")
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        df = load_sample_data()
        df = process_data(df)
        st.info("サンプルデータを使用して表示します。アップロードしたデータで分析するには、Excelファイルをアップロードしてください。")
else:
    df = load_sample_data()
    df = process_data(df)
    st.info("サンプルデータを使用して表示します。アップロードしたデータで分析するには、Excelファイルをアップロードしてください。")

# Create tabs for different views
tab1, tab2, tab3 = st.tabs(["📈 全体分析", "🏢 事業部別分析", "📊 詳細データ"])

with tab1:
    st.markdown('<div class="sub-header">🚀 全体の売上・総差分析</div>', unsafe_allow_html=True)
    
    # Aggregated metrics
    total_sales = df['売上'].sum()
    total_margin = df['総差'].sum()
    avg_margin_rate = (total_margin / total_sales * 100).round(2)
    
    # 期間情報を計算
    min_date = df['年月度'].min()
    max_date = df['年月度'].max()
    min_year = str(min_date)[:4]
    min_month = str(min_date)[4:6]
    max_year = str(max_date)[:4]
    max_month = str(max_date)[4:6]
    
    # 期間表示を追加
    st.markdown(f"### 📅 期間: {min_year}年{min_month}月 ～ {max_year}年{max_month}月")
    
    # Display KPIs in columns with 千円 unit
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("総売上", f"¥{int(total_sales/1000):,}千円")
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("総利益", f"¥{int(total_margin/1000):,}千円")
        st.markdown('</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("平均総差率", f"{avg_margin_rate}%")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # 事業部の対応関係を定義（コードと表示名の対応）
    business_types_code = list(df['N事業名'].unique())  # 実際にデータに存在する事業部コード
    business_types_map = {
        '1.在来': '在来事業',
        '2.ＳＯ': 'ＳＯ事業',
        '3.ＳＳ': 'ＳＳ事業',
        '4.教材': '教育教材',
        '5.ス介': 'スマート介護'
    }
    
    # スマート事業の定義（2.ＳＯ、3.ＳＳ、4.教材、5.ス介）
    smart_business_codes = ['2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']
    
    # 色の設定
    colors_map = {
        '1.在来': '#FFD700',  # 濃い黄
        '2.ＳＯ': '#FF8C00',  # 濃いオレンジ
        '3.ＳＳ': '#43A047',  # 緑
        '4.教材': '#FF0000',  # 赤
        '5.ス介': '#1E88E5'   # 青
    }
    
    # Yearly comparison
    st.markdown('<div class="sub-header">📆 年次比較</div>', unsafe_allow_html=True)
    
    # Create tabs for different business categories
    yearly_tab1, yearly_tab2, yearly_tab3 = st.tabs(["📊 ALL", "🏢 在来事業", "💡 スマート事業"])
    
    with yearly_tab1:
        # 年次データをピボットで準備
        yearly_pivot = pd.pivot_table(
            df,
            index='年',
            columns='N事業名',
            values='売上',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # 積み上げ棒グラフを作成
        fig_yearly = go.Figure()
        
        for i, code in enumerate(business_types_code):
            if code in yearly_pivot.columns:
                label = business_types_map.get(code, code)
                color = colors_map.get(code, '#CCCCCC')
                
                fig_yearly.add_trace(
                    go.Bar(
                        x=yearly_pivot['年'],
                        y=yearly_pivot[code] / 1000,  # 千円単位に変換
                        name=label,
                        marker_color=color,
                        hovertemplate='<b>%{x}年</b><br>' + label + ': ¥%{y:,.0f}千<extra></extra>'
                    )
                )
        
        # 総差率の計算
        yearly_all = df.groupby('年').agg({
        '売上': 'sum',
        '総差': 'sum'
    }).reset_index()
    
        yearly_all['総差率'] = np.where(
            yearly_all['売上'] > 0,
            (yearly_all['総差'] / yearly_all['売上'] * 100).round(2),
            0
        )
        
        # 総差率の線グラフを追加
        fig_yearly.add_trace(
            go.Scatter(
                x=yearly_all['年'],
                y=yearly_all['総差率'],
                name='総差率',
                mode='lines+markers+text',  # テキストを追加
                line=dict(color='#FF0000', width=2),  # 赤色に設定
                yaxis='y2',
                text=yearly_all['総差率'].apply(lambda x: f'{x:.2f}%'),  # 総差率の値を表示
                textposition='top center',  # テキストの位置を上部中央に
                textfont=dict(size=14, color='#FF0000'),  # テキストのフォント設定
                hovertemplate='<b>%{x}年</b><br>総差率: %{y:.2f}%<extra></extra>'
            )
        )
        
        # レイアウト設定
        fig_yearly.update_layout(
            title='全体：年次売上・総差率の推移',
            barmode='stack',
            height=600,
            yaxis=dict(
                title='金額（千円）',
                side='left',
                tickfont=dict(size=14),
                range=[0, 400000]  # Y軸の最高値を4億（400,000千円）に設定
            ),
            yaxis2=dict(
                title='総差率 (%)',
                side='right',
                overlaying='y',
                showgrid=False,
                range=[0, max(yearly_all['総差率']) * 1.3],
                tickfont=dict(size=14)
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5,
                font=dict(size=14)
            ),
            margin=dict(t=100, b=80),
            annotations=[
                dict(
                    x=yearly_pivot['年'].iloc[i],
                    y=yearly_pivot.iloc[i, 1:].sum() / 1000 + 20000,  # 積み上げ棒の上部に表示
                    text=f"¥{int(yearly_pivot.iloc[i, 1:].sum() / 1000):,}千",
                    font=dict(size=16, color='black', family='Arial'),
                    showarrow=False
                ) for i in range(len(yearly_pivot))
            ]
        )
        
        st.plotly_chart(fig_yearly, use_container_width=True)
        
        # 全体の年次売上・総差一覧表を追加（横向きに表示）
        st.markdown("#### 全体：年次売上・総差一覧")
        
        # 表示用のデータフレームを作成
        yearly_table = yearly_all[['年', '売上', '総差', '総差率']].copy()
        yearly_table['売上_千'] = yearly_table['売上'] / 1000
        yearly_table['総差_千'] = yearly_table['総差'] / 1000
        
        # 横向きのテーブルに変換
        yearly_wide = pd.DataFrame()
        yearly_wide['指標'] = ['売上（千円）', '総差（千円）', '総差率（%）']
        
        for year in yearly_table['年'].unique():
            year_data = yearly_table[yearly_table['年'] == year]
            yearly_wide[year] = [
                f"{int(year_data['売上_千'].values[0]):,}",
                f"{int(year_data['総差_千'].values[0]):,}",
                f"{year_data['総差率'].values[0]:.2f}"
            ]
        
        # インデックスを表示せずに右寄せテーブルを表示
        st.markdown(
            """
            <style>
            table td:nth-child(n+2) {
                text-align: right;
            }
            </style>
            """, 
            unsafe_allow_html=True
        )
        st.table(yearly_wide.set_index('指標'))
    
    with yearly_tab2:
        # Filter data for 在来
        zarai_yearly = df[df['N事業名'] == '1.在来'].groupby('年').agg({
            '売上': 'sum',
            '総差': 'sum'
        }).reset_index()
        
        # 総差率の計算と千円単位への変換
        zarai_yearly['総差率'] = np.where(
            zarai_yearly['売上'] > 0,
            (zarai_yearly['総差'] / zarai_yearly['売上'] * 100).round(2),
            0
        )
        zarai_yearly['売上_千'] = zarai_yearly['売上'] / 1000
        zarai_yearly['総差_千'] = zarai_yearly['総差'] / 1000
        
        # Create bar chart with year comparison for 在来
        fig_yearly_zarai = go.Figure()
        
        fig_yearly_zarai.add_trace(
        go.Bar(
                x=zarai_yearly['年'],
                y=zarai_yearly['売上_千'],
            name='売上',
            marker_color='#1E88E5',
                text=zarai_yearly['売上_千'].apply(lambda x: f'¥{x:,.0f}千'),
                textposition='outside',
                textfont=dict(size=11, color='black', family='Arial'),
                marker_line_width=1,
                marker_line_color='black',
                hovertemplate='<b>%{x}年</b><br>売上: ¥%{y:,.0f}千<extra></extra>'
            )
        )
        
        fig_yearly_zarai.add_trace(
        go.Bar(
                x=zarai_yearly['年'],
                y=zarai_yearly['総差_千'],
            name='総差',
            marker_color='#43A047',
                text=zarai_yearly['総差_千'].apply(lambda x: f'¥{x:,.0f}千'),
                textposition='inside',
                textfont=dict(size=11, color='black', family='Arial'),
                marker_line_width=1,
                marker_line_color='black',
                hovertemplate='<b>%{x}年</b><br>総差: ¥%{y:,.0f}千<extra></extra>'
            )
        )
        
        fig_yearly_zarai.add_trace(
        go.Scatter(
                x=zarai_yearly['年'],
                y=zarai_yearly['総差率'],
                mode='lines+markers+text',
            name='総差率',
            yaxis='y2',
                line=dict(color='#FF5722', width=2),
                marker=dict(size=8),
                text=zarai_yearly['総差率'].apply(lambda x: f'{x:.2f}%'),
                textposition='top center',
                textfont=dict(size=11, color='#FF5722', family='Arial'),
                hovertemplate='<b>%{x}年</b><br>総差率: %{y:.2f}%<extra></extra>'
            )
        )
        
        fig_yearly_zarai.update_layout(
            title='在来事業：年次売上・総差・総差率の推移',
            barmode='group',
            height=550,
            yaxis=dict(
                title='金額（千円）',
                side='left',
                tickfont=dict(size=14)
            ),
            yaxis2=dict(
                title='総差率 (%)',
                side='right',
                overlaying='y',
                showgrid=False,
                tickfont=dict(size=14)
            ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
                x=0.5,
                font=dict(size=14)
            ),
            margin=dict(t=100, b=80)
        )
        
        st.plotly_chart(fig_yearly_zarai, use_container_width=True)

    with yearly_tab3:
        # Filter data for Smart business (更新された定義を使用)
        smart_yearly = df[df['N事業名'].isin(smart_business_codes)].groupby('年').agg({
        '売上': 'sum',
        '総差': 'sum'
    }).reset_index()
    
        # 総差率の計算と千円単位への変換
        smart_yearly['総差率'] = np.where(
            smart_yearly['売上'] > 0,
            (smart_yearly['総差'] / smart_yearly['売上'] * 100).round(2),
            0
        )
        smart_yearly['売上_千'] = smart_yearly['売上'] / 1000
        smart_yearly['総差_千'] = smart_yearly['総差'] / 1000
        
        # Create bar chart with year comparison for Smart business
        fig_yearly_smart = go.Figure()
        
        fig_yearly_smart.add_trace(
        go.Bar(
                x=smart_yearly['年'],
                y=smart_yearly['売上_千'],
            name='売上',
            marker_color='#1E88E5',
                text=smart_yearly['売上_千'].apply(lambda x: f'¥{x:,.0f}千'),
            textposition='outside',
                textfont=dict(size=11, color='black', family='Arial'),
                marker_line_width=1,
                marker_line_color='black',
                hovertemplate='<b>%{x}年</b><br>売上: ¥%{y:,.0f}千<extra></extra>'
        )
    )
    
        fig_yearly_smart.add_trace(
        go.Bar(
                x=smart_yearly['年'],
                y=smart_yearly['総差_千'],
            name='総差',
            marker_color='#43A047',
                text=smart_yearly['総差_千'].apply(lambda x: f'¥{x:,.0f}千'),
            textposition='inside',
                textfont=dict(size=11, color='black', family='Arial'),
                marker_line_width=1,
                marker_line_color='black',
                hovertemplate='<b>%{x}年</b><br>総差: ¥%{y:,.0f}千<extra></extra>'
        )
    )
    
        fig_yearly_smart.add_trace(
        go.Scatter(
                x=smart_yearly['年'],
                y=smart_yearly['総差率'],
            mode='lines+markers+text',
            name='総差率',
            yaxis='y2',
                line=dict(color='#FF5722', width=2),
                marker=dict(size=8),
                text=smart_yearly['総差率'].apply(lambda x: f'{x:.2f}%'),
            textposition='top center',
                textfont=dict(size=11, color='#FF5722', family='Arial'),
            hovertemplate='<b>%{x}年</b><br>総差率: %{y:.2f}%<extra></extra>'
        )
    )
    
        fig_yearly_smart.update_layout(
            title='スマート事業：年次売上・総差・総差率の推移',
        barmode='group',
            height=550,
        yaxis=dict(
                title='金額（千円）',
                side='left',
                tickfont=dict(size=14)
        ),
        yaxis2=dict(
            title='総差率 (%)',
            side='right',
            overlaying='y',
                showgrid=False,
                tickfont=dict(size=14)
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
                x=0.5,
                font=dict(size=14)
            ),
            margin=dict(t=100, b=80)
        )
        
        st.plotly_chart(fig_yearly_smart, use_container_width=True)
        
        # スマート事業の年次売上・総差一覧表を追加（横向きに表示）
        st.markdown("#### スマート事業：年次売上・総差一覧")
        
        # 横向きのテーブルに変換
        smart_wide = pd.DataFrame()
        smart_wide['指標'] = ['売上（千円）', '総差（千円）']
        
        for year in smart_yearly['年'].unique():
            year_data = smart_yearly[smart_yearly['年'] == year]
            smart_wide[year] = [
                f"{int(year_data['売上_千'].values[0]):,}",
                f"{int(year_data['総差_千'].values[0]):,}"
            ]
        
        # インデックスを表示せずにテーブルを表示
        st.table(smart_wide.set_index('指標'))
    
    # Monthly trend analysis
    st.markdown('<div class="sub-header">📅 月次推移</div>', unsafe_allow_html=True)
    
    # Create tabs for different business categories
    trend_tab1, trend_tab2, trend_tab3 = st.tabs(["📊 ALL", "🏢 在来事業", "💡 スマート事業"])
    
    with trend_tab1:
        import numpy as np
        import pandas as pd
        
        # デバッグ表示を削除し、本番用にする
        
        # 月次データを準備
        df_monthly = df.copy()
        df_monthly['年月'] = df_monthly['年'] + '/' + df_monthly['月']
        all_months = sorted(df_monthly['年月'].unique())
        
        # 集計用のピボットデータを作成
        sales_pivot = pd.pivot_table(
            df_monthly,
            index='年月',
            columns='N事業名',
            values='売上',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # 積み上げ棒グラフを作成
        fig = go.Figure()
        
        for i, code in enumerate(business_types_code):
            if code in sales_pivot.columns:
                label = business_types_map.get(code, code)
                color = colors_map.get(code, '#CCCCCC')
                
                fig.add_trace(
                    go.Bar(
                        x=sales_pivot['年月'],
                        y=sales_pivot[code] / 1000,  # 千円単位に変換
                        name=label,
                        marker_color=color,
                        hovertemplate='<b>%{x}</b><br>' + label + ': ¥%{y:,.0f}千<extra></extra>'
                    )
                )
        
        # 総差率の計算
        monthly_all = df_monthly.groupby('年月').agg({
            '売上': 'sum',
            '総差': 'sum'
        }).reset_index()
        
        monthly_all['総差率'] = np.where(
            monthly_all['売上'] > 0,
            (monthly_all['総差'] / monthly_all['売上'] * 100).round(2),
            0
        )
        
        # レイアウト設定
        fig.update_layout(
            title='事業部別売上（積み上げ棒グラフ）',
            barmode='stack',  # ここで積み上げモードを指定
            height=500,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        
        # 積み上げ棒グラフを表示
        st.plotly_chart(fig, use_container_width=True)
        
        # 総差率のグラフ（別グラフとして表示）
        fig_rate = go.Figure()
        
        fig_rate.add_trace(
            go.Scatter(
                x=monthly_all['年月'],
                y=monthly_all['総差率'],
                mode='lines+markers',
                name='総差率',
                line=dict(color='#FF5722', width=3),
                marker=dict(size=10),
                hovertemplate='<b>%{x}</b><br>総差率: %{y:.2f}%<extra></extra>'
            )
        )
        
        fig_rate.update_layout(
            title='全体：月次総差率 (%)',
            height=300,
            yaxis=dict(title='総差率 (%)'),
            showlegend=True
        )
        
        st.plotly_chart(fig_rate, use_container_width=True)
    
    with trend_tab2:
        # Filter data for 在来
        zarai_data = df[df['N事業名'] == '1.在来'].copy()
        monthly_zarai = zarai_data.groupby(['年', '月']).agg({
            '売上': 'sum',
            '総差': 'sum'
        }).reset_index()
        
        monthly_zarai['総差率'] = (monthly_zarai['総差'] / monthly_zarai['売上'] * 100).round(2)
        monthly_zarai['年月'] = monthly_zarai['年'] + '/' + monthly_zarai['月']
        monthly_zarai['売上_千'] = monthly_zarai['売上'] / 1000
        monthly_zarai['総差_千'] = monthly_zarai['総差'] / 1000
        
        # Create subplots for 在来
        fig_zarai = make_subplots(
            rows=2, 
            cols=1, 
            subplot_titles=('在来事業：月次売上と総差（千円）', '在来事業：月次総差率'), 
            vertical_spacing=0.15,
            row_heights=[0.6, 0.4]
        )
        
        # Add bar chart for sales
        fig_zarai.add_trace(
            go.Bar(
                x=monthly_zarai['年月'],
                y=monthly_zarai['売上_千'],
                name='売上',
                marker_color='#1E88E5',
                hovertemplate='<b>%{x}</b><br>売上: ¥%{y:,.0f}千<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Add bar chart for margin
        fig_zarai.add_trace(
            go.Bar(
                x=monthly_zarai['年月'],
                y=monthly_zarai['総差_千'],
                name='総差',
                marker_color='#43A047',
                hovertemplate='<b>%{x}</b><br>総差: ¥%{y:,.0f}千<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Add line chart for margin rate
        fig_zarai.add_trace(
            go.Scatter(
                x=monthly_zarai['年月'],
                y=monthly_zarai['総差率'],
                mode='lines+markers',
                name='総差率',
                line=dict(color='#FF5722', width=3),
                marker=dict(size=10),
                hovertemplate='<b>%{x}</b><br>総差率: %{y:.2f}%<extra></extra>'
            ),
            row=2, col=1
        )
        
        # Update layout
        fig_zarai.update_layout(
            height=700,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            )
        )
        
        st.plotly_chart(fig_zarai, use_container_width=True)
    
    with trend_tab3:
        # Filter data for Smart business (更新された定義を使用)
        smart_data = df[df['N事業名'].isin(smart_business_codes)].copy()
        monthly_smart = smart_data.groupby(['年', '月']).agg({
            '売上': 'sum',
            '総差': 'sum'
        }).reset_index()
        
        monthly_smart['総差率'] = (monthly_smart['総差'] / monthly_smart['売上'] * 100).round(2)
        monthly_smart['年月'] = monthly_smart['年'] + '/' + monthly_smart['月']
        monthly_smart['売上_千'] = monthly_smart['売上'] / 1000
        monthly_smart['総差_千'] = monthly_smart['総差'] / 1000
        
        # Create subplots for Smart business
        fig_smart = make_subplots(rows=2, cols=1, 
                           subplot_titles=('スマート事業：月次売上と総差（千円）', 'スマート事業：月次総差率'), 
                           vertical_spacing=0.15,
                           row_heights=[0.6, 0.4])
        
        # Add bar chart for sales
        fig_smart.add_trace(
            go.Bar(
                x=monthly_smart['年月'],
                y=monthly_smart['売上_千'],
                name='売上',
                marker_color='#1E88E5',
                hovertemplate='<b>%{x}</b><br>売上: ¥%{y:,.0f}千<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Add bar chart for margin
        fig_smart.add_trace(
            go.Bar(
                x=monthly_smart['年月'],
                y=monthly_smart['総差_千'],
                name='総差',
                marker_color='#43A047',
                hovertemplate='<b>%{x}</b><br>総差: ¥%{y:,.0f}千<extra></extra>'
            ),
            row=1, col=1
        )
        
        # Add line chart for margin rate
        fig_smart.add_trace(
            go.Scatter(
                x=monthly_smart['年月'],
                y=monthly_smart['総差率'],
                mode='lines+markers',
                name='総差率',
                line=dict(color='#FF5722', width=3),
                marker=dict(size=10),
                hovertemplate='<b>%{x}</b><br>総差率: %{y:.2f}%<extra></extra>'
            ),
            row=2, col=1
        )
    
    # Update layout
        fig_smart.update_layout(
        height=700,
            showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
            )
    )
    
        st.plotly_chart(fig_smart, use_container_width=True)

with tab2:
    st.markdown('<div class="sub-header">🏢 事業部別分析</div>', unsafe_allow_html=True)
    
    # 年度選択をマルチセレクトボタン形式に変更
    available_years = sorted(df['年'].unique())
    selected_year = st.multiselect(
        "分析する年度を選択してください:",
        options=available_years,
        default=[available_years[-1]]  # デフォルトで最新の年度を選択
    )

    if not selected_year:  # 年度が選択されていない場合
        st.warning('年度を選択してください')
        st.stop()
    
    # 選択された年度のデータでフィルタリング
    df_year = df[df['年'].isin(selected_year)]
    
    # Aggregate by business unit for selected years
    bu_data = df_year.groupby('N事業名').agg({
        '売上': 'sum',
        '総差': 'sum'
    }).reset_index()
    
    bu_data['総差率'] = (bu_data['総差'] / bu_data['売上'] * 100).round(2)
    
    # 事業部の順序を定義
    business_order = ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']
    
    # データを事業部の順序でソート
    bu_data['N事業名'] = pd.Categorical(bu_data['N事業名'], categories=business_order, ordered=True)
    bu_data = bu_data.sort_values('N事業名')
    
    # Display business unit metrics
    st.markdown(f'<div class="sub-header">📊 {", ".join(selected_year)}年度 事業部別パフォーマンス</div>', unsafe_allow_html=True)
    
    # Create pie charts for sales and margin distribution
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        fig = px.pie(
            bu_data, 
            values='売上', 
            names='N事業名',
            title=f'{", ".join(selected_year)}年度 事業部別売上割合',
            color='N事業名',
            color_discrete_map=colors_map,
            category_orders={'N事業名': business_order},
            hole=0.4
        )
        fig.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            hovertemplate='<b>%{label}</b><br>売上: ¥%{value:,.0f}<br>割合: %{percent}<extra></extra>'
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        fig = px.pie(
            bu_data, 
            values='総差', 
            names='N事業名',
            title=f'{", ".join(selected_year)}年度 事業部別総差割合',
            color='N事業名',
            color_discrete_map=colors_map,
            category_orders={'N事業名': business_order},
            hole=0.4
        )
        fig.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            hovertemplate='<b>%{label}</b><br>総差: ¥%{value:,.0f}<br>割合: %{percent}<extra></extra>'
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Comparison of business units
    st.markdown(f'<div class="sub-header">📈 {", ".join(selected_year)}年度 事業部比較</div>', unsafe_allow_html=True)
    
    fig = go.Figure()
    
    # Add bars for sales and margin
    fig.add_trace(
        go.Bar(
            x=bu_data['N事業名'],
            y=bu_data['売上'] / 1000,  # 千円単位に変更
            name='売上',
            marker_color='#1E88E5',
            text=bu_data['売上'].apply(lambda x: f'¥{int(x/1000):,}千'),  # テキスト表示も千円単位に
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>売上: ¥%{y:,.0f}千<extra></extra>'  # ホバー表示も千円単位に
        )
    )
    
    fig.add_trace(
        go.Bar(
            x=bu_data['N事業名'],
            y=bu_data['総差'] / 1000,  # 千円単位に変更
            name='総差',
            marker_color='#43A047',
            text=bu_data['総差'].apply(lambda x: f'¥{int(x/1000):,}千'),  # テキスト表示も千円単位に
            textposition='inside',
            hovertemplate='<b>%{x}</b><br>総差: ¥%{y:,.0f}千<extra></extra>'  # ホバー表示も千円単位に
        )
    )
    
    # Add markers for margin rate
    fig.add_trace(
        go.Scatter(
            x=bu_data['N事業名'],
            y=bu_data['総差率'],
            mode='markers+text',
            name='総差率',
            yaxis='y2',
            marker=dict(
                color='#FF0000',  # 赤色に変更
                size=16,
                symbol='diamond'
            ),
            text=bu_data['総差率'].apply(lambda x: f'{x:.2f}%'),
            textposition='top right',
            hovertemplate='<b>%{x}</b><br>総差率: %{y:.2f}%<extra></extra>'
        )
    )
    
    # Update layout
    fig.update_layout(
        title='事業部別売上・総差・総差率',
        barmode='group',
        height=500,
        yaxis=dict(
            title='金額 (千円)',  # 単位を千円に変更
            side='left'
        ),
        yaxis2=dict(
            title='総差率 (%)',
            side='right',
            overlaying='y',
            showgrid=False
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
        margin=dict(l=60, r=60, t=80, b=60)
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Monthly trend by business unit
    st.markdown('<div class="sub-header">📅 事業部別月次推移</div>', unsafe_allow_html=True)
    
    # 事業部別月次データの作成
    bu_monthly = df.copy()
    bu_monthly['年月'] = bu_monthly['年'] + '/' + bu_monthly['月']
    bu_monthly = bu_monthly.groupby(['N事業名', '年月']).agg({
        '売上': 'sum',
        '総差': 'sum'
    }).reset_index()
    
    # 総差率の計算
    bu_monthly['総差率'] = (bu_monthly['総差'] / bu_monthly['売上'] * 100).round(2)
    
    # スタイルを追加
    st.markdown("""
        <style>
        div.row-widget.stRadio > div {
            flex-direction: row;
            align-items: center;
            gap: 1rem;
            background-color: #f8f9fa;
            padding: 1rem;
            border-radius: 10px;
        }
        div.row-widget.stRadio > div[role="radiogroup"] > label {
            background: white;
            padding: 10px 20px;
            border-radius: 5px;
            border: 1px solid #1E88E5;
            cursor: pointer;
            margin: 0.2rem;
        }
        div.row-widget.stRadio > div[role="radiogroup"] > label:hover {
            background: #e3f2fd;
        }
        div.row-widget.stRadio > div[role="radiogroup"] > label[data-baseweb="radio"] > div {
            display: none;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # ドロップダウンをラジオボタンに変更
    metric_option = st.radio(
        "表示する指標を選択してください:",
        ["売上", "総差", "総差率"],
        horizontal=True  # 水平に配置
    )
    
    # Create line chart for selected metric by business unit
    if metric_option in ['売上', '総差']:
        y_values = bu_monthly[metric_option] / 1000  # 千円単位に変更
    else:
        y_values = bu_monthly[metric_option]  # 総差率はそのまま

    fig = px.line(
        bu_monthly,
        x='年月',
        y=y_values,
        color='N事業名',
        markers=True,
        title=f'事業部別 {metric_option} 月次推移',
        color_discrete_map=colors_map,
        category_orders={'N事業名': business_order}
    )
    
    # Update layout
    fig.update_layout(
        height=500,
        xaxis_title='年月',
        yaxis_title=f'{metric_option}{"（千円）" if metric_option in ["売上", "総差"] else "（%）"}',  # 単位を追加
        legend_title='事業部',
        hovermode='x unified',
        margin=dict(l=60, r=60, t=80, b=60)
    )
    
    # Format hover template based on metric
    if metric_option == '売上' or metric_option == '総差':
        fig.update_traces(
            hovertemplate='<b>%{x}</b><br>%{y:,.0f}千円<extra>%{fullData.name}</extra>'  # 千円単位に変更
        )
    else:  # For 総差率
        fig.update_traces(
            hovertemplate='<b>%{x}</b><br>%{y:.2f}%<extra>%{fullData.name}</extra>'
        )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Heat map of performance
    st.markdown('<div class="sub-header">🔥 パフォーマンスヒートマップ</div>', unsafe_allow_html=True)
    
    # ドロップダウンをラジオボタンに変更
    pivot_option = st.radio(
        "ヒートマップに表示する指標を選択してください:",
        ["売上", "総差", "総差率"],
        horizontal=True  # 水平に配置
    )
    
    # Create pivot table
    if pivot_option in ['売上', '総差']:
        pivot_data = df[df['年'].isin(selected_year)].pivot_table(  # 選択された年度でフィルタリング
            index='N事業名',
            columns='月',
            values=pivot_option,
            aggfunc='sum'
        ).reset_index()
        pivot_data.iloc[:, 1:] = pivot_data.iloc[:, 1:] / 1000  # 数値列を千円単位に変更
    else:
        pivot_data = df[df['年'].isin(selected_year)].pivot_table(  # 選択された年度でフィルタリング
        index='N事業名',
        columns='月',
        values=pivot_option,
        aggfunc='sum'
    ).reset_index()
    
    # Sort columns if they are months
    if pivot_data.columns.dtype == 'object' and len(pivot_data.columns) > 1:
        try:
            month_columns = [col for col in pivot_data.columns if col != 'N事業名']
            month_columns.sort()
            sorted_columns = ['N事業名'] + month_columns
            pivot_data = pivot_data[sorted_columns]
        except:
            pass
    
    # Prepare data for heatmap
    heatmap_data = pivot_data.set_index('N事業名')
    
    # Convert to long format for plotly
    heatmap_long = pd.melt(
        pivot_data, 
        id_vars=['N事業名'], 
        var_name='月', 
        value_name=pivot_option
    )
    
    # Create heatmap
    fig = px.imshow(
        heatmap_data,
        labels=dict(x="月", y="事業部", color=f"{pivot_option}（{'千円' if pivot_option in ['売上', '総差'] else '%'}）"),
        x=heatmap_data.columns,
        y=heatmap_data.index,
        color_continuous_scale='Blues' if pivot_option != '総差率' else 'RdYlGn',
        title=f'事業部別 {pivot_option} ヒートマップ（{selected_year[0]}年度）'  # 年度を追加
    )
    
    # Add annotations
    annotations = []
    for i, row in enumerate(heatmap_data.index):
        for j, col in enumerate(heatmap_data.columns):
            value = heatmap_data.iloc[i, j]
            if pd.notna(value):
                if pivot_option == '売上' or pivot_option == '総差':
                    text = f'{value:,.0f}'
                else:  # For 総差率
                    text = f'{value:.2f}%'
                annotations.append(
                    dict(
                        x=col,
                        y=row,
                        text=text,
                        showarrow=False,
                        font=dict(color='black' if value < (heatmap_data.max().max() / 2) else 'white')
                    )
                )
    
    # Add subtitle with unit information
    subtitle = f"※数値の単位: {'千円' if pivot_option in ['売上', '総差'] else '%'}"
    
    fig.update_layout(
        annotations=annotations + [
            dict(
                x=0.5,
                y=-0.15,
                xref="paper",
                yref="paper",
                text=subtitle,
                showarrow=False,
                font=dict(size=12),
                align="center"
            )
        ],
        height=400,
        margin=dict(l=60, r=60, t=80, b=80)  # 下部マージンを増やして単位表示のスペースを確保
    )
    
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.markdown('<div class="sub-header">📊 詳細データ</div>', unsafe_allow_html=True)
    
    # Filter options
    st.markdown('<div class="sub-header" style="font-size: 1.5rem;">🔍 データフィルタ</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-card" style="padding: 2rem;">', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    
    with col1:
        # Get unique years and months
        years = sorted(df['年'].unique())
        selected_years = st.multiselect(
            "年を選択:",
            options=years,
            default=years
        )
    
    with col2:
        # Get unique business units
        bus = sorted(df['N事業名'].unique())
        selected_bus = st.multiselect(
            "事業部を選択:",
            options=bus,
            default=bus
        )
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Filter data
    filtered_df = df[df['年'].isin(selected_years) & df['N事業名'].isin(selected_bus)]
    
    # Show filtered data
    st.markdown('<div class="sub-header" style="font-size: 1.5rem;">📋 フィルタリングされたデータ</div>', unsafe_allow_html=True)
    st.dataframe(filtered_df)
    
    # Download button for filtered data
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, sheet_name='フィルタデータ', index=False)
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['フィルタデータ']
        
        # Add some cell formats
        format_header = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        format_money = workbook.add_format({'num_format': '¥#,##0'})
        format_percent = workbook.add_format({'num_format': '0.00%'})
        
        # Write the column headers with the defined format
        for col_num, value in enumerate(filtered_df.columns.values):
            worksheet.write(0, col_num, value, format_header)
            
            # Set column width based on content
            if value in ['売上', '総差']:
                worksheet.set_column(col_num, col_num, 15, format_money)
            elif value == '総差率':
                worksheet.set_column(col_num, col_num, 10, format_percent)
            else:
                worksheet.set_column(col_num, col_num, 12)
    
    buffer.seek(0)
    
    st.download_button(
        label="データをExcelにダウンロード",
        data=buffer,
        file_name="売上分析データ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Add footer with instructions
st.markdown("""
---
<div style="background-color: #f8f9fa; padding: 2rem; border-radius: 10px; margin-top: 2rem;">
    <h3 style="color: #1E88E5;">💡 使用方法</h3>
    <ol style="margin-top: 1rem;">
        <li>エクセルファイルをアップロードしてください（フォーマット: 年月度, N事業名, 売上, 総差）</li>
        <li>「全体分析」タブでは年次・月次の全体的な売上と総差の推移が確認できます</li>
        <li>「事業部別分析」タブでは各事業部のパフォーマンスを比較できます</li>
        <li>「詳細データ」タブではデータをフィルタリングしてExcelにエクスポートできます</li>
    </ol>
</div>
""", unsafe_allow_html=True)