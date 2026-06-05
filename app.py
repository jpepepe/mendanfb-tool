# -*- coding: utf-8 -*-
import streamlit as st

st.set_page_config(
    page_title="人材紹介 FB ツール",
    page_icon="💬",
    layout="wide"
)

pg = st.navigation([
    st.Page("pages/01_💬_初回面談FBツール.py",      title="初回面談FBツール",   icon="💬"),
    st.Page("pages/02_💼_求人提案FBツール.py",      title="求人提案FBツール",   icon="💼"),
    st.Page("pages/03_📊_初回面談ダッシュボード.py", title="初回面談ダッシュボード", icon="📊"),
    st.Page("pages/04_💼_求人提案ダッシュボード.py", title="求人提案ダッシュボード", icon="💼"),
])
pg.run()
