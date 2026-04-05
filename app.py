import streamlit as st
import pandas as pd
import json
import os
from extractor import extract_pack_data, extract_price_data
from pdf_generator import generate_packing_list, generate_invoice
from excel_updater import update_excel_master

st.set_page_config(page_title="Invoice & Packing List Generator", layout="wide")

st.title("🍎 蘋果出貨 Invoice & Packing List 生成工具")

st.sidebar.header("環境設定 / Settings")
api_key = st.sidebar.text_input("輸入 Gemini API Key", type="password")

st.header("1. 匯入資料 (Upload Data)")

col1, col2 = st.columns(2)
with col1:
    pack_file = st.file_uploader("上傳 Pack-Sample (PDF 檔案)", type=["pdf"])
with col2:
    price_file = st.file_uploader("上傳 Price-Sample (PDF 檔案)", type=["pdf"])
    
order_no = st.text_input("輸入 注文番號 (e.g., USN 1031)")
excel_master = st.file_uploader("上傳 總表 Excel (可選)", type=["xlsx"])

if 'pack_data' not in st.session_state:
    st.session_state.pack_data = []
if 'price_data' not in st.session_state:
    st.session_state.price_data = []

if st.button("利用 AI 辨識資料 (Extract Data)"):
    if not api_key:
        st.error("請先在左側輸入 Gemini API Key")
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
st.write("您可以在下方表格直接修改辨識錯誤的數字，或新增/刪除行")

pack_df = pd.DataFrame(st.session_state.pack_data)
if pack_df.empty:
    pack_df = pd.DataFrame(columns=["variety", "grade", "size", "quantity"])
edited_pack = st.data_editor(pack_df, num_rows="dynamic", key="pack_editor")

price_df = pd.DataFrame(st.session_state.price_data)
if price_df.empty:
    price_df = pd.DataFrame(columns=["variety", "grade", "size", "price"])
edited_price = st.data_editor(price_df, num_rows="dynamic", key="price_editor")

st.header("3. 產生檔案 (Generate Files)")
if st.button("生成 PDF 與更新總表"):
    if edited_pack.empty:
        st.error("沒有 Pack 資料，無法產生檔案。")
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
                
                st.success("PDF 產生成功！")
                
                with open(pl_path, "rb") as f:
                    st.download_button("下載 Packing List PDF", f, file_name="PackingList.pdf")
                with open(inv_path, "rb") as f:
                    st.download_button("下載 Invoice PDF", f, file_name="Invoice.pdf")
                
                # Excel Update Logic
                if excel_master and order_no:
                    excel_path = os.path.join("output", "master.xlsx")
                    with open(excel_path, "wb") as f:
                        f.write(excel_master.getbuffer())
                    
                    import datetime
                    date_str = datetime.date.today().strftime("%Y/%m/%d")
                    update_excel_master(excel_path, order_no, date_str, 0)
                    with open(excel_path, "rb") as f:
                        st.download_button("下載更新後的總表 Excel", f, file_name=f"{order_no}_master.xlsx")
                
            except Exception as e:
                st.error(f"產生過程中發生錯誤: {e}")
