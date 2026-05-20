import streamlit as st
import pandas as pd
import VIX  # 載入您的爬蟲程式

st.set_page_config(page_title="個人槓桿控制面板", layout="wide")
st.title("📈 個人槓桿控制 - 台股 VIX 監控面板")
st.markdown("這是我專屬的 VIX 自動化監控網頁，資料來源為期交所最新數據。")
st.divider()

# 讓雲端伺服器自動去抓取最新資料 (快取 1 小時)
@st.cache_data(ttl=3600)
def load_data():
    # 呼叫 VIX.py 裡面的下載功能，產生 CSV
    VIX.download_vix_web_table()
    # 讀取剛剛產生的 CSV
    df = pd.read_csv("VIX_Recent_2Months_Data.csv")
    return df

try:
    with st.spinner('正在從期交所抓取最新資料中...'):
        df = load_data()
    
    date_col = next((col for col in df.columns if '日期' in col), None)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col])
        df = df.sort_values(by=date_col)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📊 VIX 近期走勢圖")
        close_col = next((col for col in df.columns if '收盤' in col), None)
        if date_col and close_col:
            chart_data = df.set_index(date_col)[close_col]
            st.line_chart(chart_data)
        else:
            st.warning("找不到日期或收盤價欄位，無法繪製圖表。")

    with col2:
        st.subheader("📋 詳細數據表格")
        st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error(f"抓取資料發生錯誤：{e}")