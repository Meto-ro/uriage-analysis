import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import datetime
import sys
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

# 事業部ごとの色を設定（グローバル変数として定義）
DEPT_COLORS = {
    '1.在来': '#FFD700',  # 濃い黄
    '2.ＳＯ': '#FF8C00',  # 濃いオレンジ
    '3.ＳＳ': '#43A047',  # 緑
    '4.教材': '#FF0000',  # 赤
    '5.ス介': '#1E88E5'   # 青
}

def export_to_gsheet(df_grouped, month, dept_sales, group_sales, cross_analysis):
    """Googleスプレッドシートにデータを出力する関数"""
    try:
        # NaN値を0に置換
        df_grouped = df_grouped.fillna(0)
        dept_sales = dept_sales.fillna(0)
        group_sales = group_sales.fillna(0)
        cross_analysis = cross_analysis.fillna(0)
        
        # 数値を文字列に変換（NaN対策）
        for df in [df_grouped, dept_sales, group_sales, cross_analysis]:
            for col in df.select_dtypes(include=['float64', 'int64']).columns:
                df[col] = df[col].astype(str)

        # 認証情報の設定
        credentials_dict = dict(st.secrets["gcp_service_account"])
        scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # 認証情報をJSONファイルとして一時的に保存
        with open('temp_credentials.json', 'w') as f:
            json.dump(credentials_dict, f)
        
        # 認証の実行
        credentials = ServiceAccountCredentials.from_json_keyfile_name('temp_credentials.json', scope)
        gc = gspread.authorize(credentials)
        
        # スプレッドシートの作成
        spreadsheet_title = f"売上分析_{month}"
        sh = gc.create(spreadsheet_title)
        
        # 一時ファイルの削除
        os.remove('temp_credentials.json')
        
        # 編集者権限を持つユーザーのリスト
        editors = [
            'metoh@jointex.jp',  # あなたのアカウント
            # 他の編集者のメールアドレスをここに追加
        ]
        
        # 特定のユーザーに編集権限を付与
        for editor in editors:
            sh.share(
                editor,
                perm_type='user',
                role='writer',
                notify=False
            )
        
        # 全員に閲覧権限を付与
        sh.share(
            None,
            perm_type='anyone',
            role='reader',
            notify=False
        )
        
        # 各シートにデータを書き込む
        # サマリーシート
        worksheet = sh.get_worksheet(0)
        worksheet.update_title('サマリー')
        worksheet.update([df_grouped.columns.values.tolist()] + df_grouped.values.tolist())
        
        # 事業部別分析
        worksheet = sh.add_worksheet(title='事業部別分析', rows=str(len(dept_sales)+1), cols=str(len(dept_sales.columns)))
        worksheet.update([dept_sales.columns.values.tolist()] + dept_sales.values.tolist())
        
        # 商品群分析
        worksheet = sh.add_worksheet(title='商品群分析', rows=str(len(group_sales)+1), cols=str(len(group_sales.columns)))
        worksheet.update([group_sales.columns.values.tolist()] + group_sales.values.tolist())
        
        # クロス分析
        worksheet = sh.add_worksheet(title='クロス分析', rows=str(len(cross_analysis)+1), cols=str(len(cross_analysis.columns)))
        worksheet.update([cross_analysis.columns.values.tolist()] + cross_analysis.values.tolist())
        
        # スプレッドシートのURLを返す
        return sh.url
        
    except Exception as e:
        st.error(f"Googleスプレッドシートへの出力中にエラーが発生しました: {e}")
        return None

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

def main():
    # パスワード認証
    if not check_password():
        return
        
    # ページ設定
    st.set_page_config(
        page_title="エスコ商品売上分析ダッシュボード",
        page_icon="📊",
        layout="wide"
    )

    # カスタムCSS
    st.markdown("""
        <style>
        /* 基本レイアウト */
        .main {
            background-color: #f5f5f5;
            padding: 0rem 1rem;
        }
        
        /* コンテナスタイル */
        .container {
            background-color: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 2rem;
        }
        
        /* タイトルコンテナ */
        .title-container {
            background-color: #1a237e;
            color: white;
            padding: 2rem;
            border-radius: 10px;
            margin-bottom: 2rem;
            text-align: center;
        }
        
        /* メトリクス */
        .stMetric {
            background-color: white !important;
            padding: 1rem !important;
            border-radius: 10px !important;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
        }
        
        /* データフレーム */
        .dataframe {
            background-color: white !important;
            padding: 1rem !important;
            border-radius: 10px !important;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
        }
        
        /* 見出し */
        h1 {
            color: white !important;
            font-size: 2.5rem !important;
            font-weight: bold !important;
            margin: 0 !important;
        }
        
        h2 {
            color: #1a237e !important;
            font-size: 1.8rem !important;
            margin-bottom: 1rem !important;
        }
        
        h3 {
            color: #283593 !important;
            font-size: 1.5rem !important;
            margin-bottom: 0.8rem !important;
        }
        
        /* タブ */
        .stTabs [data-baseweb="tab-list"] {
            gap: 2rem;
            background-color: transparent !important;
        }
        
        .stTabs [data-baseweb="tab"] {
            height: 4rem;
            white-space: pre-wrap;
            background-color: white !important;
            border-radius: 4px;
            color: #1a237e !important;
            font-weight: bold !important;
            border: none !important;
            padding: 1rem !important;
        }
        
        .stTabs [aria-selected="true"] {
            background-color: #1a237e !important;
            color: white !important;
        }
        
        /* ボタン */
        .stButton > button {
            background-color: #1a237e !important;
            color: white !important;
            font-weight: bold !important;
            border: none !important;
            padding: 0.5rem 1rem !important;
            border-radius: 5px !important;
        }
        
        .stButton > button:hover {
            background-color: #283593 !important;
        }
        
        /* 印刷用スタイル */
        @media print {
            .main {
                padding: 0 !important;
            }
            .stDeployButton, 
            .stToolbar, 
            .stSpinner, 
            [data-testid="stDecoration"], 
            [data-testid="stToolbar"],
            .exportButton {
                display: none !important;
            }
            .element-container {
                break-inside: avoid;
            }
            .chart-container {
                break-inside: avoid;
                page-break-inside: avoid;
            }
            @page {
                margin: 1cm;
                size: A4 portrait;
            }
        }
        </style>
    """, unsafe_allow_html=True)

    # タイトルセクション
    st.markdown("""
        <div style='background-color: #1a237e; padding: 2rem; border-radius: 10px; margin-bottom: 2rem; text-align: center;'>
            <h1 style='color: white; font-size: 2.5rem; font-weight: bold; margin: 0;'>📊 エスココード商品売上分析ダッシュボード</h1>
        </div>
        <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
            <p style='margin-bottom: 1rem;'>このダッシュボードでは、エスココード商品の売上データを分析し、各販売チャネルでの売上状況を可視化します。</p>
            <p style='margin-bottom: 1rem;'>主な機能：</p>
            <ul style='margin-bottom: 1rem;'>
                <li>商品別の売上、総差、総差率などの主要指標の確認</li>
                <li>販売チャネル別の売上構成と商品カテゴリーのクロス分析</li>
                <li>商品群ごとの売上傾向分析</li>
            </ul>
            <p style='margin-bottom: 1rem;'>※ このデータは各カタログの商品採用選定の参考資料としてご活用いただけます。</p>
            <p style='margin-bottom: 1rem;'>アイコンの説明：</p>
            <ul style='margin-bottom: 1rem;'>
                <li>📦 商品別分析：商品ごとの売上・総差分析</li>
                <li>🏢 事業部別分析：事業部ごとの売上構成比と実績</li>
                <li>📊 総合分析：商品群別のクロス分析</li>
                <li>🆕 新規採用商品：新規採用商品の分析（商品名の前に🆕マークが付きます）</li>
                <li>💰 売上額　💹 総差額　📦 商品数　📈 総差率</li>
            </ul>
            <p style='color: #666; font-size: 0.9rem; text-align: right; margin: 1rem 0 0 0;'>Created by 商品戦略部PS課</p>
        </div>
    """, unsafe_allow_html=True)

    print("Python実行パス:", sys.executable)
    
    # ファイルアップロード
    st.markdown("""
        <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
            <h3 style='margin-top: 0; color: #1a237e;'>📁 データ入力</h3>
        </div>
    """, unsafe_allow_html=True)
    uploaded_file = st.file_uploader("エクセルファイルをアップロードしてください", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            # ファイル名のチェック
            if "エスコ" not in uploaded_file.name or "衞藤" not in uploaded_file.name:
                st.error("承認されていないファイル形式です")
                st.stop()
            
            # 全ての列を一旦文字列として読み込む
            df = pd.read_excel(
                uploaded_file,
                dtype=str
            )
            
            # カラム名のマッピング
            column_mapping = {
                '商品コード': '商品コード',
                '商品名': '商品漢字名',
                '群コード': '群２',
                '事業部名': 'Ｎ事業名',
                'N数': '出荷数',
                '売上額': '売上',
                '総差': '総差',
                '年': '年',
                '受注開始': '受注開始'
            }
            
            # カラム名の変更
            df = df.rename(columns=column_mapping)
            
            # 欠損値を0に置換
            df = df.fillna(0)
            
            # 数値型への変換（カンマを除去してから変換）
            numeric_columns = ['出荷数', '売上', '総差']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')
            
            # 受注開始年の抽出（上4桁を取得）
            df['受注開始年'] = df['受注開始'].str[:4]
            
            # 新規採用商品の判定
            df['新規採用'] = (df['年'] == df['受注開始年'])
            
            # 事業名の確認
            if 'Ｎ事業名' in df.columns:
                # 事業名の正規化（全角/半角の違いなどを統一）
                standard_business_names = ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']
                business_name_mapping = {}
                unique_business_names = df['Ｎ事業名'].unique()
                
                for name in unique_business_names:
                    mapped_name = name
                    if name in standard_business_names:
                        mapped_name = name
                    elif '在来' in name:
                        mapped_name = '1.在来'
                    elif 'SO' in name or 'ＳＯ' in name:
                        mapped_name = '2.ＳＯ'
                    elif 'SS' in name or 'ＳＳ' in name:
                        mapped_name = '3.ＳＳ'
                    elif '教材' in name:
                        mapped_name = '4.教材'
                    elif 'ス介' in name or 'スマート介護' in name or 'ｽﾏｰﾄ介護' in name:
                        mapped_name = '5.ス介'
                    business_name_mapping[name] = mapped_name
                
                # 事業名を正規化
                df['Ｎ事業名'] = df['Ｎ事業名'].map(lambda x: business_name_mapping.get(x, x))
                
                # マッピングされなかった事業名があれば警告
                unmapped_names = df[~df['Ｎ事業名'].isin(standard_business_names)]['Ｎ事業名'].unique()
                if len(unmapped_names) > 0:
                    st.warning(f"以下の事業名が標準形式にマッピングできませんでした: {', '.join(unmapped_names)}")
            
            # 解析月の表示
            default_year = df['年'].iloc[0] if not df['年'].empty else datetime.datetime.now().strftime("%Y")
            month = st.text_input("解析対象月 (YYYY-MM)", 
                                f"{default_year}-{datetime.datetime.now().strftime('%m')}")

        except Exception as e:
            st.error(f"データ読み込み中にエラーが発生しました: {e}")
            return

        # データの読み込み後、新規採用商品の表示名を修正
        df['表示用商品名'] = df.apply(lambda row: f"🆕 {row['商品漢字名']}" if row['新規採用'] else row['商品漢字名'], axis=1)

        # 商品コード別の集計
        df_grouped = df.groupby('商品コード').agg({
            '商品漢字名': 'first',
            '表示用商品名': 'first',
            '群２': 'first',
            '商品群２名': 'first',
            'Ｎ事業名': 'first',
            '出荷数': 'sum',
            '売上': 'sum',
            '総差': 'sum',
            '新規採用': 'first'
        }).reset_index()
        
        # 数値を整数型に変換
        df_grouped['出荷数'] = df_grouped['出荷数'].astype(int)
        df_grouped['売上'] = df_grouped['売上'].astype(int)
        df_grouped['総差'] = df_grouped['総差'].astype(int)
        
        # 基本統計量の表示をより魅力的に
        st.markdown("""
            <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
            <h3 style='margin-top: 0;'>📈 基本データ概要</h3>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("💰 総売上額", f"{df_grouped['売上'].sum() // 1000:,}千円",
                     delta=None, delta_color="normal")
        with col2:
            st.metric("💹 総差額", f"{df_grouped['総差'].sum() // 1000:,}千円",
                     delta=None, delta_color="normal")
        with col3:
            st.metric("📦 商品数", f"{len(df_grouped):,}点",
                     delta=None, delta_color="normal")
        with col4:
            st.metric("📊 平均総差率", f"{df_grouped['総差'].sum() / df_grouped['売上'].sum():.1%}",
                     delta=None, delta_color="normal")

        # タブのスタイリング
        st.markdown("""
            <div style='height: 2rem;'></div>
        """, unsafe_allow_html=True)
        
        tab1, tab2, tab3, tab4 = st.tabs(["📦 商品別分析", "🏢 事業部別分析", "📊 総合分析", "🆕 新規採用商品分析"])
        
        with tab1:
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>🏆 売上高ランキング（商品別）</h3>
                </div>
            """, unsafe_allow_html=True)
            
            # 商品別売上ランキング
            top_products = df_grouped.sort_values('売上', ascending=False).head(20)
            top_products = top_products.reset_index(drop=True)
            top_products.index = top_products.index + 1
            
            # 売上と総差を千円単位に変換（数値型を維持）
            top_products['売上'] = (top_products['売上'] / 1000).astype(int)
            top_products['総差'] = (top_products['総差'] / 1000).astype(int)
            
            # 売上高グラフ
            fig = px.bar(
                top_products,
                x='売上',
                y='表示用商品名',  # 商品漢字名から表示用商品名に変更
                color='Ｎ事業名',
                title=f'売上上位20商品 ({month})',
                orientation='h',
                height=800,
                text='売上',
                color_discrete_map=DEPT_COLORS,
                category_orders={"Ｎ事業名": ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']}
            )
            
            # グラフのレイアウト調整
            fig.update_layout(
                xaxis_title='売上（千円）',
                yaxis_title='商品名',
                yaxis={'categoryorder':'total ascending'},  # 売上額順に並べ替え
                showlegend=True,
                legend_title='N事業名',
                legend=dict(
                    yanchor="middle",   # 凡例を中央に
                    y=0.5,             # 凡例の垂直位置を中央に
                    xanchor="right",   # 凡例を右寄せに
                    x=1.15             # 凡例の水平位置をグラフの右側に少し離して配置
                ),
                margin=dict(l=20, r=100, t=40, b=20),  # 右側の余白を増やして凡例用のスペースを確保
                uniformtext=dict(mode='hide', minsize=8)  # テキストサイズの最小値を設定
            )
            
            # 売上額の表示フォーマット調整
            fig.update_traces(
                texttemplate='¥%{text:,d}',
                textposition='inside',
                insidetextanchor='start',    # テキストを左寄せに
                textangle=0                  # テキストを水平に
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 詳細ランキング表
            # 総差率を計算（小数点第1位まで表示）
            top_products['総差率'] = (top_products['総差'] / top_products['売上'] * 100).round(1)
            
            st.dataframe(
                top_products[['商品コード', '表示用商品名', '商品群２名', '出荷数', '売上', '総差', '総差率']],
                hide_index=False,
                height=800,
                column_config={
                    "商品コード": st.column_config.TextColumn("商品コード"),
                    "表示用商品名": st.column_config.TextColumn("商品名"),
                    "商品群２名": st.column_config.TextColumn("商品群"),
                    "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                    "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                    "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                    "総差率": st.column_config.NumberColumn("総差率", format="%.1f%%")
                }
            )
        
        with tab2:
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>🏢 事業部別分析</h3>
                </div>
            """, unsafe_allow_html=True)
            
            # 事業部別集計
            dept_sales = df_grouped.groupby('Ｎ事業名').agg({
                '売上': 'sum',
                '総差': 'sum',
                '出荷数': 'sum',
                '商品コード': 'count'
            }).reset_index()

            # 事業部名に含まれる数字で並び替え
            dept_order = {'1.在来': 1, '2.ＳＯ': 2, '3.ＳＳ': 3, '4.教材': 4, '5.ス介': 5}
            dept_sales['sort_order'] = dept_sales['Ｎ事業名'].map(lambda x: dept_order.get(x, 99))
            dept_sales = dept_sales.sort_values('sort_order')
            
            # 売上と総差を千円単位に変換（数値型を維持）
            dept_sales['売上'] = (dept_sales['売上'] / 1000).astype(int)
            dept_sales['総差'] = (dept_sales['総差'] / 1000).astype(int)

            # 総差率を計算（小数点第1位まで表示）
            dept_sales['総差率'] = (dept_sales['総差'] / dept_sales['売上'] * 100).round(1)

            # 事業部名の表示用マッピング
            display_names = {
                '1.在来': '1.在来',
                '2.ＳＯ': '2.ＳＯ',
                '3.ＳＳ': '3.ＳＳ',
                '4.教材': '4.教材',
                '5.ス介': '5.ス介'
            }
            dept_sales['表示名'] = dept_sales['Ｎ事業名'].map(lambda x: display_names.get(x, x))
            
            # 事業部別売上構成比
            fig = px.pie(
                dept_sales.sort_values('sort_order'), # sort_orderでソート
                values='売上', 
                names='表示名',
                title=f"事業部別売上構成比 ({month})",
                hole=0.4,
                color='表示名',
                color_discrete_map=DEPT_COLORS,
                category_orders={"表示名": ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']}  # 凡例の順序を指定
            )

            # グラフのレイアウト調整
            fig.update_layout(
                showlegend=True,
                legend=dict(
                    orientation="v",
                    yanchor="middle",
                    y=0.5,
                    xanchor="right",
                    x=1.1
                )
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 事業部別詳細表
            st.dataframe(
                dept_sales.rename(columns={'商品コード': '商品数'})[['Ｎ事業名', '売上', '総差', '総差率', '出荷数', '商品数']],
                hide_index=True,
                column_config={
                    "Ｎ事業名": st.column_config.TextColumn("事業部名"),
                    "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                    "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                    "総差率": st.column_config.NumberColumn("総差率", format="%.1f%%"),
                    "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                    "商品数": st.column_config.NumberColumn("商品数", format="%d")
                }
            )
            
            # 各事業部のTOP5商品
            st.subheader("事業部別売上上位商品")
            
            # タブの作成（ドロップダウンの代わりに）
            dept_tabs = st.tabs([dept for dept in dept_sales['Ｎ事業名'].tolist()])
            
            # 各事業部のタブ内容
            for idx, dept_tab in enumerate(dept_tabs):
                with dept_tab:
                    selected_dept = dept_sales['Ｎ事業名'].tolist()[idx]
                    
                    # 選択された事業部のデータのみをフィルタリング
                    top_dept_products = df_grouped[df_grouped['Ｎ事業名'] == selected_dept].sort_values('売上', ascending=False).head(10)
                    
                    # 売上順位を追加
                    top_dept_products = top_dept_products.reset_index(drop=True)
                    top_dept_products.index = top_dept_products.index + 1
                    
                    # 売上を千円単位に変換（数値型を維持）
                    top_dept_products['売上'] = (top_dept_products['売上'] / 1000).astype(int)
                    top_dept_products['総差'] = (top_dept_products['総差'] / 1000).astype(int)
                    
                    # 総差率を計算（小数点第1位まで必ず表示）
                    top_dept_products['総差率'] = (top_dept_products['総差'] / top_dept_products['売上'] * 100).apply(lambda x: float(f"{x:.1f}"))
                    
                    # 商品名と商品群名の空白を除去し、組み合わせた列を作成
                    top_dept_products['商品漢字名'] = top_dept_products['商品漢字名'].str.strip()
                    top_dept_products['商品群２名'] = top_dept_products['商品群２名'].str.strip()
                    top_dept_products['表示用商品名'] = top_dept_products['表示用商品名'].str.strip()
                    
                    # 売上の高い順にソート
                    top_dept_products = top_dept_products.sort_values('売上', ascending=False)
                    
                    # グラフの作成
                    fig = px.bar(
                        top_dept_products,
                        x='売上',
                        y='表示用商品名',
                        title=f"{selected_dept}の売上上位10商品 ({month})",
                        orientation='h',
                        height=500,
                        text='売上',
                        color_discrete_sequence=[DEPT_COLORS[selected_dept]]
                    )
                    
                    # グラフのレイアウト調整
                    fig.update_layout(
                        xaxis_title='売上（千円）',
                        yaxis_title='商品名',
                        yaxis={'categoryorder':'total ascending'},
                        showlegend=False,
                        margin=dict(l=300, r=20, t=40, b=20),
                        uniformtext=dict(mode='hide', minsize=8)
                    )
                    
                    # 売上額の表示フォーマット調整
                    fig.update_traces(
                        texttemplate='¥%{text:,d}',
                        textposition='inside',
                        insidetextanchor='start',
                        textangle=0
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 事業部別のTOP10商品表示をデータフレームでも表示
                    st.dataframe(
                        top_dept_products[['商品コード', '表示用商品名', '出荷数', '売上', '総差', '総差率']],
                        hide_index=False,
                        column_config={
                            "index": st.column_config.NumberColumn(
                                "売上順位",
                                help="売上金額順の順位"
                            ),
                            "表示用商品名": st.column_config.TextColumn(
                                "商品名",
                                width="large",
                                help="商品の名称"
                            ),
                            "売上": st.column_config.NumberColumn(
                                "売上（千円）",
                                format="%d"
                            ),
                            "総差": st.column_config.NumberColumn(
                                "総差（千円）",
                                format="%d"
                            ),
                            "総差率": st.column_config.NumberColumn(
                                "総差率",
                                format="%.1f%%"
                            )
                        }
                    )
        
        with tab3:
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>📊 総合分析</h3>
                </div>
            """, unsafe_allow_html=True)
            
            # 群２別分析
            group_sales = df_grouped.groupby('商品群２名').agg({
                '売上': 'sum',
                '総差': 'sum',
                '出荷数': 'sum'
            }).reset_index()
            group_sales = group_sales.sort_values('売上', ascending=False)
            
            # 売上と総差を千円単位に変換（数値型を維持）
            group_sales['売上'] = (group_sales['売上'] / 1000).astype(int)
            group_sales['総差'] = (group_sales['総差'] / 1000).astype(int)
            
            # 総差率を計算（小数点第1位まで必ず表示）
            group_sales['総差率'] = (group_sales['総差'] / group_sales['売上'] * 100).apply(lambda x: float(f"{x:.1f}"))
            
            fig = px.bar(
                group_sales.head(10),
                x='売上',
                y='商品群２名',
                title=f"商品群別売上TOP10 ({month})",
                orientation='h',
                height=500,
                text=['¥{:,d}\n{:.1f}%'.format(s, m) for s, m in zip(group_sales.head(10)['売上'], group_sales.head(10)['総差率'])]
            )
            
            # グラフのレイアウト調整
            fig.update_layout(
                xaxis_title='売上（千円）',
                yaxis_title='商品群名',
                yaxis={'categoryorder':'total ascending'},  # 売上額順に並べ替え
                showlegend=False,
                margin=dict(l=200, r=20, t=40, b=20)
            )
            
            # 売上額と総差率の表示フォーマット調整
            fig.update_traces(
                textposition='inside',
                insidetextanchor='start',
                textangle=0
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # 群２別売上のデータフレーム表示
            st.dataframe(
                group_sales.head(10),
                hide_index=True,
                column_config={
                    "商品群２名": st.column_config.TextColumn("商品群名"),
                    "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                    "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                    "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                    "総差率": st.column_config.NumberColumn("総差率", format="%.1f")
                }
            )
            
            # クロス分析の追加
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>📊 商品カテゴリーと事業部のクロス分析</h3>
                </div>
            """, unsafe_allow_html=True)
            
            # 事業部順序を定義
            dept_order_list = ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']
            
            # クロス集計用のピボットテーブルを作成（売上を千円単位に）
            cross_analysis = pd.pivot_table(
                df_grouped,
                values='売上',
                index='商品群２名',
                columns='Ｎ事業名',
                aggfunc='sum',
                fill_value=0
            ) // 1000

            # 事業部の順序を指定
            existing_depts = [dept for dept in dept_order_list if dept in cross_analysis.columns]
            
            # 存在しない事業部がある場合は警告を表示
            missing_depts = [dept for dept in dept_order_list if dept not in cross_analysis.columns]
            if missing_depts:
                st.warning(f"以下の事業部のデータが見つかりませんでした: {', '.join(missing_depts)}")
            
            # 存在する事業部のみを使用
            cross_analysis = cross_analysis[existing_depts]

            # 行と列の合計を計算
            cross_analysis['合計'] = cross_analysis.sum(axis=1)
            row_totals = cross_analysis['合計']
            col_totals = cross_analysis[existing_depts].sum()

            # パーセンテージの計算
            percentage_data = cross_analysis[existing_depts].div(cross_analysis['合計'], axis=0) * 100

            # 商品群を売上合計で降順ソート
            cross_analysis = cross_analysis.loc[row_totals.sort_values(ascending=False).index]
            percentage_data = percentage_data.loc[row_totals.sort_values(ascending=False).index]

            # 上位10商品群のみを表示
            cross_analysis = cross_analysis.head(10)
            percentage_data = percentage_data.head(10)

            # 表題を追加
            st.markdown(f"#### 商品カテゴリーと事業部のクロス分析詳細（TOP10）")
            st.markdown(f"※ {month}の実績")

            # ヒートマップの作成
            fig = go.Figure()

            # テキスト表示用のデータを作成
            text_data = []
            for i in range(len(cross_analysis)):
                row = []
                for j in range(len(existing_depts)):
                    sales = cross_analysis[existing_depts].values[i][j]
                    pct = percentage_data.values[i][j]
                    row.append(f"売上：¥{sales:,}K<br>構成比：{pct:.1f}％")  # テキストの説明を追加
                text_data.append(row)

            # 売上額のヒートマップ
            fig.add_trace(go.Heatmap(
                z=cross_analysis[existing_depts].values,
                x=existing_depts,
                y=cross_analysis.index,
                text=text_data,
                texttemplate="%{text}",
                customdata=percentage_data.values,
                textfont={"size": 12},
                hoverongaps=False,
                colorscale='Blues',
                showscale=True,
                colorbar=dict(
                    title=dict(
                        text='売上（千円）',
                        side='top'
                    ),
                    orientation='h',
                    y=1.0,
                    x=0.5,
                    len=1.0,
                    thickness=25,
                    tickfont=dict(size=10)
                )
            ))

            # レイアウトの調整
            fig.update_layout(
                title=dict(
                    text=f"商品カテゴリーと事業部のクロス分析 ({month})<br><span style='font-size:12px'>※ 各セルの表示：売上（千円）/ 商品群内での構成比（%）</span>",
                    font=dict(size=16),
                    y=0.92,  # タイトル位置を元に戻す
                    x=0.4,
                    xanchor='center'
                ),
                height=800,
                margin=dict(
                    l=20,
                    r=250,
                    t=130,  # 上部マージンを確保
                    b=100
                ),
                xaxis=dict(
                    title=dict(
                        text='事業部',
                        font=dict(size=12),
                        standoff=50
                    ),
                    tickangle=0,
                    tickfont=dict(size=11),
                    side='bottom'
                ),
                yaxis=dict(
                    title=dict(
                        text='商品群',
                        font=dict(size=12),
                        standoff=10
                    ),
                    tickfont=dict(size=11),
                    side='right'
                ),
                font=dict(size=12),
                annotations=[  # 凡例の説明を追加
                    dict(
                        x=1.15,
                        y=-0.15,
                        xref="paper",
                        yref="paper",
                        text="【数値の見方】<br>売上：各セルの売上金額（千円）<br>構成比：商品群における事業部別の売上構成比（%）<br>色の濃さ：売上金額の大きさを表現",
                        showarrow=False,
                        font=dict(size=11),
                        align="left"
                    )
                ]
            )

            # ヒートマップの表示
            st.plotly_chart(fig, use_container_width=True)

            # クロス分析の説明と総評の追加
            st.markdown("""
            ### クロス分析の見方
            - 各セルには売上金額（千円）と構成比（%）を表示
            - 構成比は各商品群の総売上に対する事業部別の割合
            - 色の濃さは売上金額の大きさを表現
            - TOP10の商品群のみを表示し、重要な部分に焦点を当て
            """)

            # データに基づく分析ロジックの追加
            def generate_analysis_summary(cross_analysis, percentage_data):
                summary = {
                    'dept_insights': [],
                    'product_insights': []
                }
                
                # 事業部別の主力商品群分析
                for dept in existing_depts:
                    dept_data = cross_analysis[dept]
                    total_sales = dept_data.sum()
                    top_categories = dept_data.nlargest(3)
                    top_categories_names = cross_analysis[cross_analysis[dept].isin(top_categories)].index.tolist()
                    
                    # 事業部の特徴を分析
                    if total_sales > 0:  # 売上が存在する場合のみ
                        summary['dept_insights'].append(
                            f"- {dept}の主力商品群：{', '.join(top_categories_names[:3])} "
                            f"（売上：{top_categories.sum():,}千円）"
                        )
                
                # 商品群の売上規模分析
                for product in cross_analysis.index:
                    sales_data = cross_analysis.loc[product]
                    total_sales = sales_data.sum()
                    max_dept = sales_data.idxmax()
                    max_sales = sales_data.max()
                    
                    if total_sales > 8000:  # 売上8,000千円以上の主力商品群
                        summary['product_insights'].append(
                            f"- {product}は{max_dept}が最も売上が高く（{max_sales:,}千円）、"
                            f"全体で{total_sales:,}千円の売上規模"
                        )
                
                return summary

            # 分析サマリーの生成
            analysis = generate_analysis_summary(cross_analysis, percentage_data)

            st.markdown("""
            ### データに基づく分析総評
            
            #### 事業部別の主力商品群""")
            for insight in analysis['dept_insights']:
                st.markdown(insight)
            
            st.markdown("""
            #### 売上規模の大きい商品群の状況""")
            for insight in analysis['product_insights']:
                st.markdown(insight)
            
            st.markdown("""
            #### 分析のポイント
            - 各事業部の主力となっている商品群を把握し、その特徴を理解
            - 売上規模の大きい商品群がどの事業部で強みを持っているかを確認
            - 商品群と事業部の関係性から、各事業部の事業特性を把握
            """)

            # データフレームでの詳細表示
            st.subheader("商品カテゴリーと事業部のクロス分析詳細（TOP10）")
            st.markdown("※ 下表の数値は全て売上金額（単位：千円）を表しています")
            
            # 表示用のデータフレーム作成
            display_df = cross_analysis[existing_depts].copy()
            display_df['合計'] = cross_analysis['合計']
            
            st.dataframe(
                display_df,
                column_config={col: st.column_config.NumberColumn(
                    f"{col}（千円）" if col != "合計" else f"{col}（千円）",  # 列名に単位を追加
                    format="%d"
                ) for col in display_df.columns}
            )
        
        with tab4:
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>🆕 新規採用商品分析</h3>
                </div>
            """, unsafe_allow_html=True)
            
            # 新規採用商品の分析
            new_products = df[df['新規採用'] == True]
            
            # 新規採用商品の集計
            new_products_grouped = new_products.groupby('商品コード').agg({
                '商品漢字名': 'first',
                '表示用商品名': 'first',  # 表示用商品名を追加
                '群２': 'first',
                '商品群２名': 'first',
                'Ｎ事業名': 'first',
                '出荷数': 'sum',
                '売上': 'sum',
                '総差': 'sum'
            }).reset_index()
            
            # 数値を整数型に変換
            new_products_grouped['出荷数'] = new_products_grouped['出荷数'].astype(int)
            new_products_grouped['売上'] = new_products_grouped['売上'].astype(int)
            new_products_grouped['総差'] = new_products_grouped['総差'].astype(int)
            
            # 基本統計量の表示をより魅力的に
            st.markdown("""
                <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                <h3 style='margin-top: 0;'>📈 基本データ概要</h3>
                </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("💰 総売上額", f"{new_products_grouped['売上'].sum() // 1000:,}千円",
                         delta=None, delta_color="normal")
            with col2:
                st.metric("💹 総差額", f"{new_products_grouped['総差'].sum() // 1000:,}千円",
                         delta=None, delta_color="normal")
            with col3:
                st.metric("📦 商品数", f"{len(new_products_grouped):,}点",
                         delta=None, delta_color="normal")
            with col4:
                st.metric("📊 平均総差率", f"{new_products_grouped['総差'].sum() / new_products_grouped['売上'].sum():.1%}",
                         delta=None, delta_color="normal")

            # タブのスタイリング
            st.markdown("""
                <div style='height: 2rem;'></div>
            """, unsafe_allow_html=True)
            
            tab1, tab2, tab3 = st.tabs(["📦 商品別分析", "🏢 事業部別分析", "📊 総合分析"])
            
            with tab1:
                st.markdown("""
                    <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                    <h3 style='margin-top: 0;'>🏆 売上高ランキング（商品別）</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                # 商品別売上ランキング
                top_products = new_products_grouped.sort_values('売上', ascending=False).head(20)
                top_products = top_products.reset_index(drop=True)
                top_products.index = top_products.index + 1
                
                # 売上と総差を千円単位に変換（数値型を維持）
                top_products['売上'] = (top_products['売上'] / 1000).astype(int)
                top_products['総差'] = (top_products['総差'] / 1000).astype(int)
                
                # 売上高グラフ
                fig = px.bar(
                    top_products,
                    x='売上',
                    y='表示用商品名',  # 商品漢字名から表示用商品名に変更
                    color='Ｎ事業名',
                    title=f'売上上位20商品 ({month})',
                    orientation='h',
                    height=800,
                    text='売上',
                    color_discrete_map=DEPT_COLORS,
                    category_orders={"Ｎ事業名": ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']}
                )
                
                # グラフのレイアウト調整
                fig.update_layout(
                    xaxis_title='売上（千円）',
                    yaxis_title='商品名',
                    yaxis={'categoryorder':'total ascending'},  # 売上額順に並べ替え
                    showlegend=True,
                    legend_title='N事業名',
                    legend=dict(
                        yanchor="middle",   # 凡例を中央に
                        y=0.5,             # 凡例の垂直位置を中央に
                        xanchor="right",   # 凡例を右寄せに
                        x=1.15             # 凡例の水平位置をグラフの右側に少し離して配置
                    ),
                    margin=dict(l=20, r=100, t=40, b=20),  # 右側の余白を増やして凡例用のスペースを確保
                    uniformtext=dict(mode='hide', minsize=8)  # テキストサイズの最小値を設定
                )
                
                # 売上額の表示フォーマット調整
                fig.update_traces(
                    texttemplate='¥%{text:,d}',
                    textposition='inside',
                    insidetextanchor='start',    # テキストを左寄せに
                    textangle=0                  # テキストを水平に
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # 詳細ランキング表
                # 総差率を計算（小数点第1位まで表示）
                top_products['総差率'] = (top_products['総差'] / top_products['売上'] * 100).round(1)
                
                st.dataframe(
                    top_products[['商品コード', '表示用商品名', '商品群２名', '出荷数', '売上', '総差', '総差率']],
                    hide_index=False,
                    height=800,
                    column_config={
                        "商品コード": st.column_config.TextColumn("商品コード"),
                        "表示用商品名": st.column_config.TextColumn("商品名"),
                        "商品群２名": st.column_config.TextColumn("商品群"),
                        "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                        "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                        "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                        "総差率": st.column_config.NumberColumn("総差率", format="%.1f%%")
                    }
                )
            
            with tab2:
                st.markdown("""
                    <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                    <h3 style='margin-top: 0;'>🏢 事業部別分析</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                # 事業部別集計
                dept_sales = new_products_grouped.groupby('Ｎ事業名').agg({
                    '売上': 'sum',
                    '総差': 'sum',
                    '出荷数': 'sum',
                    '商品コード': 'count'
                }).reset_index()

                # 事業部名に含まれる数字で並び替え
                dept_order = {'1.在来': 1, '2.ＳＯ': 2, '3.ＳＳ': 3, '4.教材': 4, '5.ス介': 5}
                dept_sales['sort_order'] = dept_sales['Ｎ事業名'].map(lambda x: dept_order.get(x, 99))
                dept_sales = dept_sales.sort_values('sort_order')
                
                # 売上と総差を千円単位に変換（数値型を維持）
                dept_sales['売上'] = (dept_sales['売上'] / 1000).astype(int)
                dept_sales['総差'] = (dept_sales['総差'] / 1000).astype(int)

                # 総差率を計算（小数点第1位まで表示）
                dept_sales['総差率'] = (dept_sales['総差'] / dept_sales['売上'] * 100).round(1)

                # 事業部名の表示用マッピング
                display_names = {
                    '1.在来': '1.在来',
                    '2.ＳＯ': '2.ＳＯ',
                    '3.ＳＳ': '3.ＳＳ',
                    '4.教材': '4.教材',
                    '5.ス介': '5.ス介'
                }
                dept_sales['表示名'] = dept_sales['Ｎ事業名'].map(lambda x: display_names.get(x, x))
                
                # 事業部別売上構成比
                fig = px.pie(
                    dept_sales.sort_values('sort_order'), # sort_orderでソート
                    values='売上', 
                    names='表示名',
                    title=f"事業部別売上構成比 ({month})",
                    hole=0.4,
                    color='表示名',
                    color_discrete_map=DEPT_COLORS,
                    category_orders={"表示名": ['1.在来', '2.ＳＯ', '3.ＳＳ', '4.教材', '5.ス介']}  # 凡例の順序を指定
                )

                # グラフのレイアウト調整
                fig.update_layout(
                    showlegend=True,
                    legend=dict(
                        orientation="v",
                        yanchor="middle",
                        y=0.5,
                        xanchor="right",
                        x=1.1
                    )
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # 事業部別詳細表
                st.dataframe(
                    dept_sales.rename(columns={'商品コード': '商品数'})[['Ｎ事業名', '売上', '総差', '総差率', '出荷数', '商品数']],
                    hide_index=True,
                    column_config={
                        "Ｎ事業名": st.column_config.TextColumn("事業部名"),
                        "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                        "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                        "総差率": st.column_config.NumberColumn("総差率", format="%.1f%%"),
                        "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                        "商品数": st.column_config.NumberColumn("商品数", format="%d")
                    }
                )
                
                # 各事業部のTOP5商品
                st.subheader("事業部別売上上位商品")
                
                # タブの作成（ドロップダウンの代わりに）
                dept_tabs = st.tabs([dept for dept in dept_sales['Ｎ事業名'].tolist()])
                
                # 各事業部のタブ内容
                for idx, dept_tab in enumerate(dept_tabs):
                    with dept_tab:
                        selected_dept = dept_sales['Ｎ事業名'].tolist()[idx]
                        
                        # 選択された事業部のデータのみをフィルタリング
                        top_dept_products = new_products_grouped[new_products_grouped['Ｎ事業名'] == selected_dept].sort_values('売上', ascending=False).head(10)
                        
                        # 売上順位を追加
                        top_dept_products = top_dept_products.reset_index(drop=True)
                        top_dept_products.index = top_dept_products.index + 1
                        
                        # 売上を千円単位に変換（数値型を維持）
                        top_dept_products['売上'] = (top_dept_products['売上'] / 1000).astype(int)
                        top_dept_products['総差'] = (top_dept_products['総差'] / 1000).astype(int)
                        
                        # 総差率を計算（小数点第1位まで必ず表示）
                        top_dept_products['総差率'] = (top_dept_products['総差'] / top_dept_products['売上'] * 100).apply(lambda x: float(f"{x:.1f}"))
                        
                        # 商品名と商品群名の空白を除去し、組み合わせた列を作成
                        top_dept_products['商品漢字名'] = top_dept_products['商品漢字名'].str.strip()
                        top_dept_products['商品群２名'] = top_dept_products['商品群２名'].str.strip()
                        top_dept_products['表示用商品名'] = top_dept_products['表示用商品名'].str.strip()
                        
                        # 売上の高い順にソート
                        top_dept_products = top_dept_products.sort_values('売上', ascending=False)
                        
                        # グラフの作成
                        fig = px.bar(
                            top_dept_products,
                            x='売上',
                            y='表示用商品名',
                            title=f"{selected_dept}の売上上位10商品 ({month})",
                            orientation='h',
                            height=500,
                            text='売上',
                            color_discrete_sequence=[DEPT_COLORS[selected_dept]]
                        )
                        
                        # グラフのレイアウト調整
                        fig.update_layout(
                            xaxis_title='売上（千円）',
                            yaxis_title='商品名',
                            yaxis={'categoryorder':'total ascending'},
                            showlegend=False,
                            margin=dict(l=300, r=20, t=40, b=20),
                            uniformtext=dict(mode='hide', minsize=8)
                        )
                        
                        # 売上額の表示フォーマット調整
                        fig.update_traces(
                            texttemplate='¥%{text:,d}',
                            textposition='inside',
                            insidetextanchor='start',
                            textangle=0
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # 事業部別のTOP10商品表示をデータフレームでも表示
                        st.dataframe(
                            top_dept_products[['商品コード', '表示用商品名', '出荷数', '売上', '総差', '総差率']],
                            hide_index=False,
                            column_config={
                                "index": st.column_config.NumberColumn(
                                    "売上順位",
                                    help="売上金額順の順位"
                                ),
                                "表示用商品名": st.column_config.TextColumn(
                                    "商品名",
                                    width="large",
                                    help="商品の名称"
                                ),
                                "売上": st.column_config.NumberColumn(
                                    "売上（千円）",
                                    format="%d"
                                ),
                                "総差": st.column_config.NumberColumn(
                                    "総差（千円）",
                                    format="%d"
                                ),
                                "総差率": st.column_config.NumberColumn(
                                    "総差率",
                                    format="%.1f%%"
                                )
                            }
                        )
            
            with tab3:
                st.markdown("""
                    <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
                    <h3 style='margin-top: 0;'>📊 総合分析</h3>
                    </div>
                """, unsafe_allow_html=True)
                
                # 群２別分析
                group_sales = new_products_grouped.groupby('商品群２名').agg({
                    '売上': 'sum',
                    '総差': 'sum',
                    '出荷数': 'sum'
                }).reset_index()
                group_sales = group_sales.sort_values('売上', ascending=False)
                
                # 売上と総差を千円単位に変換（数値型を維持）
                group_sales['売上'] = (group_sales['売上'] / 1000).astype(int)
                group_sales['総差'] = (group_sales['総差'] / 1000).astype(int)
                
                # 総差率を計算（小数点第1位まで必ず表示）
                group_sales['総差率'] = (group_sales['総差'] / group_sales['売上'] * 100).apply(lambda x: float(f"{x:.1f}"))
                
                fig = px.bar(
                    group_sales.head(10),
                    x='売上',
                    y='商品群２名',
                    title=f"商品群別売上TOP10 ({month})",
                    orientation='h',
                    height=500,
                    text=['¥{:,d}\n{:.1f}%'.format(s, m) for s, m in zip(group_sales.head(10)['売上'], group_sales.head(10)['総差率'])]
                )
                
                # グラフのレイアウト調整
                fig.update_layout(
                    xaxis_title='売上（千円）',
                    yaxis_title='商品群名',
                    yaxis={'categoryorder':'total ascending'},  # 売上額順に並べ替え
                    showlegend=False,
                    margin=dict(l=200, r=20, t=40, b=20)
                )
                
                # 売上額と総差率の表示フォーマット調整
                fig.update_traces(
                    textposition='inside',
                    insidetextanchor='start',
                    textangle=0
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # 群２別売上のデータフレーム表示
                st.dataframe(
                    group_sales.head(10),
                    hide_index=True,
                    column_config={
                        "商品群２名": st.column_config.TextColumn("商品群名"),
                        "売上": st.column_config.NumberColumn("売上（千円）", format="%d"),
                        "総差": st.column_config.NumberColumn("総差（千円）", format="%d"),
                        "出荷数": st.column_config.NumberColumn("出荷数", format="%d"),
                        "総差率": st.column_config.NumberColumn("総差率", format="%.1f")
                    }
                )
        
        # データダウンロード機能
        st.markdown("""
            <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
            <h3 style='margin-top: 0;'>💾 分析データのダウンロード</h3>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # エクセルファイル作成
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # サマリーシート
                summary_data = pd.DataFrame({
                    '項目': ['解析対象月', '総売上額', '総差額', '平均総差率', '商品数'],
                    '値': [
                        month,
                        f"¥{new_products_grouped['売上'].sum() // 1000:,}千円",
                        f"¥{new_products_grouped['総差'].sum() // 1000:,}千円",
                        f"{new_products_grouped['総差'].sum() / new_products_grouped['売上'].sum():.1%}",
                        f"{len(new_products_grouped):,}点"
                    ]
                })
                summary_data.to_excel(writer, sheet_name='サマリー', index=False)
                
                # 商品別売上ランキング
                product_ranking = new_products_grouped.sort_values('売上', ascending=False)[
                    ['商品コード', '商品漢字名', '商品群２名', 'Ｎ事業名', '出荷数', '売上', '総差']
                ]
                product_ranking['売上順位'] = range(1, len(product_ranking) + 1)
                product_ranking['総差率'] = product_ranking.apply(
                    lambda row: row['総差'] / row['売上'] if row['売上'] != 0 else 0,
                    axis=1
                )
                product_ranking.to_excel(writer, sheet_name='商品別売上ランキング', index=False)
                
                # 事業部別分析
                dept_sales.to_excel(writer, sheet_name='事業部別分析', index=False)
                
                # 商品群分析
                group_sales.to_excel(writer, sheet_name='商品群分析', index=False)
            
            buffer.seek(0)
            
            # エクセルダウンロードボタン
            st.download_button(
                label="📊 詳細分析レポートをエクセルでダウンロード",
                data=buffer,
                file_name=f"売上分析詳細レポート_{month}.xlsx",
                mime="application/vnd.ms-excel"
            )
        
        with col2:
            # Googleスプレッドシートへの出力ボタン
            if st.button("📊 Googleスプレッドシートに出力"):
                with st.spinner('Googleスプレッドシートに出力中...'):
                    sheet_url = export_to_gsheet(new_products_grouped, month, dept_sales, group_sales, cross_analysis)
                    if sheet_url:
                        st.success('Googleスプレッドシートへの出力が完了しました！')
                        st.markdown(f'[スプレッドシートを開く]({sheet_url})', unsafe_allow_html=True)

        # 総評コメントのスタイリング
        st.markdown("""
            <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 2rem;'>
            <h3 style='margin-top: 0;'>📝 総評コメント</h3>
            </div>
        """, unsafe_allow_html=True)

        # 総評コメント
        st.subheader("総評コメント")

        # 基本データの計算
        top_group = group_sales.iloc[0]['商品群２名']
        top_group_sales = group_sales.iloc[0]['売上']
        total_sales = group_sales['売上'].sum()
        top_group_ratio = top_group_sales / total_sales * 100

        top_dept = dept_sales.iloc[0]['Ｎ事業名']
        top_dept_sales = dept_sales.iloc[0]['売上']
        top_dept_ratio = top_dept_sales / total_sales * 100

        avg_margin = group_sales['総差'].sum() / total_sales * 100

        # 追加の分析データ
        top_margin_group = group_sales.sort_values('総差率', ascending=False).iloc[0]
        low_margin_group = group_sales.sort_values('総差率').iloc[0]
        
        # 商品群の集中度分析
        top3_groups = group_sales.head(3)
        top3_ratio = top3_groups['売上'].sum() / total_sales * 100

        comment = f"""
        ### {month} 月次売上分析レポート

        #### 📊 全体概況
- 総売上規模：**¥{total_sales:,}千円**
- 平均総差率：**{avg_margin:.1f}%**

#### 💹 商品群分析
- 売上トップ商品群「**{top_group}**」が全体の**{top_group_ratio:.1f}%**を占める
- 上位3商品群で全体の**{top3_ratio:.1f}%**を占め{' 、商品群の集中度が高い' if top3_ratio > 50 else ' 、商品群が分散している'}
- 最高総差率：**{top_margin_group['商品群２名']}**（**{top_margin_group['総差率']:.1f}%**）
- 最低総差率：**{low_margin_group['商品群２名']}**（**{low_margin_group['総差率']:.1f}%**）

#### 📈 事業部分析
- **{top_dept}**が売上の**{top_dept_ratio:.1f}%**を占め、主力事業部として機能
- 事業部間の売上格差は{' 大きく、リスク分散の検討が必要' if top_dept_ratio > 40 else ' 比較的小さく、バランスが取れている'}

#### 💡 重点施策の提案
1. **収益性改善**：総差率の低い「{low_margin_group['商品群２名']}」の原価構造を見直し
2. **商品戦略**：{'上位3商品群への依存度を下げるため、新規商品群の育成を検討' if top3_ratio > 50 else '主力商品群のさらなる強化と、新規商品群の開拓を並行して推進'}
3. **事業部戦略**：{'事業部間の売上平準化に向けた施策の検討' if top_dept_ratio > 40 else '各事業部の強みを活かした成長戦略の推進'}

➡️ 特に注目すべき点：
- {'商品群の集中リスクへの対応' if top3_ratio > 50 else '商品群ポートフォリオの最適化'}
- {'収益性の改善（特に低総差率商品群）' if avg_margin < 15 else '高収益体質の維持・強化'}
- {'事業部間の連携強化による相乗効果の創出' if top_dept_ratio > 40 else '各事業部の独自性を活かした成長戦略の推進'}
"""

        st.markdown(comment)

        # PDFエクスポートの説明を最下部に移動
        st.markdown("""
            <div style='background-color: white; padding: 1.5rem; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-top: 2rem;'>
            <h3 style='margin-top: 0;'>📄 PDFエクスポート手順</h3>
            <ol style='margin-bottom: 0;'>
                <li>ブラウザの印刷機能を使用（Ctrl+P または ⌘+P）</li>
                <li>出力先を「PDFとして保存」に設定</li>
                <li>用紙サイズ：A4</li>
                <li>余白：最小</li>
                <li>背景のグラフィックスを印刷：オン</li>
            </ol>
            </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()