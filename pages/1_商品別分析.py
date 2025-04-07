st.set_page_config(layout="wide", page_title="商品別分析 - エスコ売上実績分析")

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

# パスワード認証の状態をチェック
if "password_correct" not in st.session_state or not st.session_state["password_correct"]:
    st.stop()

# Google Sheets APIの認証設定
try:
    # 環境変数から認証情報を取得
    credentials_json = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    
    # 認証情報の設定
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_json, scope)
    gc = gspread.authorize(credentials)
    
except Exception as e:
    st.error("Google Sheets APIの認証に失敗しました。")
    st.error(f"エラー: {str(e)}")
    st.stop()

# メインコンテンツ
st.title("商品別分析")

st.info("🚧 このページは現在開発中です。近日中に機能が追加される予定です。")

st.subheader("機能一覧")
st.markdown("""
- 商品ごとの売上推移分析
- カテゴリー別分析
- 新規採用商品の分析
""")

# データ読み込み用の関数
@st.cache_data(ttl=3600)  # 1時間キャッシュ
def load_data():
    try:
        # スプレッドシートを開く
        SPREADSHEET_KEY = st.secrets["spreadsheet_key"]
        worksheet = gc.open_by_key(SPREADSHEET_KEY).worksheet('売上データ')
        
        # データを取得してDataFrameに変換
        data = worksheet.get_all_values()
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # データの前処理
        # ... 必要な前処理をここに追加 ...
        
        return df
    except Exception as e:
        st.error("データの読み込みに失敗しました。")
        st.error(f"エラー: {str(e)}")
        return None

# データの読み込み
df = load_data()

if df is not None:
    # データ分析と可視化のコード
    st.write("データの読み込みが完了しました。分析を開始します。")
    # ... 以降の分析コード ...

# ... existing code from sh_analysis.py ... 