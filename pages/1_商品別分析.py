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

st.set_page_config(layout="wide", page_title="商品別分析 - エスコ売上実績分析")

# ... existing code from sh_analysis.py ... 