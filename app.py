import streamlit as st
import pandas as pd
import json
import os
from extractor import extract_pack_data, extract_price_data
from pdf_generator import generate_packing_list, generate_invoice
from google_sheets_updater import update_google_sheet

st.set_page_config(page_title="Invoice & Packing List Generator", layout="wide")

st.title("🍎 蘋果出貨 Invoice & Packing List 生成工具")

st.sidebar.header("⚙️ 環境設定 / Settings")

if 'api_key_valid' not in st.session_state:
    st.session_state.api_key_valid = False

api_key = st.sidebar.text_input("輸入 Gemini API Key", type="password")
if st.sidebar.button("確認 API Key"):
    if api_key:
        st.session_state.api_key_valid = True
        st.sidebar.success("API Key 已確認！")
    else:
        st.session_state.api_key_valid = False
        st.sidebar.error("請輸入 Key!")

st.sidebar.markdown("---")
st.sidebar.subheader("Google Sheet 設定")
sheet_url = st.sidebar.text_input("總表 Google Sheet URL")
google_creds = st.sidebar.text_area("Google Service Account Credentials (JSON)", help="貼上 Google 提供給您的 JSON 憑證")

st.header("1. 匯入資料 (Upload Data)")
col1, col2 = st.columns(2)
with col1:
    pack_file = st.file_uploader("上傳 Pack-Sample (PDF 檔案)", type=["pdf"])
with col2:
    price_file = st.file_uploader("上傳 Price-Sample (PDF 檔案)", type=["pdf"])
    
order_no = st.text_input("輸入 注文番號 (e.g., USN 1031)")

if 'pack_data' not in st.session_state:
    st.session_state.pack_data = []
if 'price_data' not in st.session_state:
    st.session_state.price_data = []
if 'pdf_generated' not in st.session_state:
    st.session_state.pdf_generated = False
if 'pl_bytes' not in st.session_state:
    st.session_state.pl_bytes = None
if 'inv_bytes' not in st.session_state:
    st.session_state.inv_bytes = None

if st.button("利用 AI 辨識資料 (Extract Data)"):
    if not api_key:
        st.error("請先在左側欄位輸入 Gemini API Key 並按下確認！")
    else:
        with st.spinner("辨識 Pack 檔案中..."):
            if pack_file:
                try:
                    res = extract_pack_data(api_key, pack_file.read())
                    st.session_state.pack_data = res
                    st.success("Pack 辨識成功！")
                except Exception as e:
                    st.error(f"Pack 辨識失敗: {e}")
        
        with st.spinner("辨識 Price 檔案中..."):
            if price_file:
                try:
                    res = extract_price_data(api_key, price_file.read())
                    st.session_state.price_data = res
                    st.success("Price 辨識成功！")
                except Exception as e:
                    st.error(f"Price 辨識失敗: {e}")

st.header("2. 預覽與編輯資料 (Preview & Edit)")
st.caption("您可以在下方表格直接修改辨識錯誤的數字，或新增/刪除行。")

pack_df = pd.DataFrame(st.session_state.pack_data)
if pack_df.empty:
    pack_df = pd.DataFrame(columns=["variety", "grade", "size", "quantity"])
edited_pack = st.data_editor(pack_df, num_rows="dynamic", key="pack_editor")

price_df = pd.DataFrame(st.session_state.price_data)
if price_df.empty:
    price_df = pd.DataFrame(columns=["variety", "grade", "size", "price"])
edited_price = st.data_editor(price_df, num_rows="dynamic", key="price_editor")

st.header("3. 產生檔案 (Generate Files)")

# Separate the Generate logic from the Download logic to prevent download buttons from disappearing.
if st.button("生成 PDF 與更新總表"):
    if edited_pack.empty:
        st.error("沒有 Pack 資料，無法產生檔案。")
    elif not order_no:
        st.error("請輸入 注文番號！")
    else:
        with st.spinner("產生檔案中..."):
            try:
                pack_list = edited_pack.to_dict('records')
                price_list = edited_price.to_dict('records')
                
                os.makedirs("output", exist_ok=True)
                pl_path = os.path.join("output", "PackingList_Output.pdf")
                inv_path = os.path.join("output", "Invoice_Output.pdf")
                
                generate_packing_list(pack_list, order_no, pl_path)
                generate_invoice(pack_list, price_list, order_no, inv_path)
                
                # Save to session_state so they persist after re-run
                with open(pl_path, "rb") as f:
                    st.session_state.pl_bytes = f.read()
                with open(inv_path, "rb") as f:
                    st.session_state.inv_bytes = f.read()
                
                st.session_state.pdf_generated = True
                
                # Google Sheets Update
                if sheet_url and google_creds:
                    try:
                        import datetime
                        date_str = datetime.date.today().strftime("%Y/%m/%d")
                        update_google_sheet(google_creds, sheet_url, order_no, date_str, 0)
                        st.success("成功登記至 Google Sheet！")
                    except Exception as ge:
                        st.error(f"寫入 Google Sheet 失敗（請確認您的 JSON 和網址是否正確）: {ge}")
                else:
                    st.warning("您沒有填寫 Google Sheet 資訊，跳過總表更新。")
                    
                st.success("🎉 PDF 產生成功！請往下捲動下載檔案。")
                
            except Exception as e:
                st.error(f"產生過程中發生錯誤: {e}")

# If we have generated PDFs in session state, always show them
if st.session_state.pdf_generated and st.session_state.pl_bytes and st.session_state.inv_bytes:
    st.markdown("### 下載區 (Downloads)")
    st.download_button("📥 下載 Packing List PDF", st.session_state.pl_bytes, file_name=f"{order_no}_PackingList.pdf", mime="application/pdf")
    st.download_button("📥 下載 Invoice PDF", st.session_state.inv_bytes, file_name=f"{order_no}_Invoice.pdf", mime="application/pdf")
