import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io

def check_password():
    """パスワード認証を行う関数"""
    def password_entered():
        if st.session_state["password"] == "esco2024":  # パスワードを設定
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

st.set_page_config(layout="wide", page_title="エスコ売上実績分析ダッシュボード")

# ... existing code from all_analysis.py ... 