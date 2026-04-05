import streamlit as st
import pandas as pd
import json
import os
from extractor import extract_pack_data, extract_price_data
from pdf_generator import generate_packing_list, generate_invoice
from google_sheets_updater import update_google_sheet, load_saved_data

st.set_page_config(page_title="Invoice & Packing List Generator", layout="wide")

st.title("🍎 蘋果出貨 Invoice & Packing List 生成工具")

st.sidebar.header("⚙️ 環境設定 / Settings")

st.sidebar.header("🔐 系統解鎖")
app_password = st.sidebar.text_input("請輸入系統密碼", type="password")

if app_password == "unis5888":
    st.sidebar.success("授權成功！環境變數已自動載入。")
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        sheet_url = st.secrets["GOOGLE_SHEET_URL"]
        google_creds = st.secrets["GOOGLE_CREDS_JSON"]
        st.session_state.api_key_valid = True
    except Exception as e:
        st.sidebar.error("⚠️ 尚未設定好 Streamlit Secrets，請到後台設定。")
        api_key, sheet_url, google_creds = "", "", ""
        st.session_state.api_key_valid = False
else:
    if app_password:
        st.sidebar.error("密碼錯誤！")
    st.info("👈 請先在左側輸入密碼以解鎖並使用工具。")
    st.stop()


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
    elif not pack_file and not price_file:
         st.warning("請先上傳 PDF 檔案！")
    else:
        if pack_file:
            with st.spinner("辨識 Pack 檔案中..."):
                try:
                    res = extract_pack_data(api_key, pack_file.read())
                    st.session_state.pack_data = res
                    st.success("Pack 辨識成功！")
                except Exception as e:
                    st.error(f"Pack 辨識失敗: {e}")
        
        if price_file:
            with st.spinner("辨識 Price 檔案中..."):
                try:
                    res = extract_price_data(api_key, price_file.read())
                    st.session_state.price_data = res
                    st.success("Price 辨識成功！")
                except Exception as e:
                    st.error(f"Price 辨識失敗: {e}")

st.markdown("---")
st.markdown("**(可選) 從 Google Sheet 載入之前已儲存的資料：**")
col_load1, col_load2 = st.columns([3,1])
with col_load1:
    load_order = st.text_input("輸入想載入的 注文番號 (例如: USN 1031)", key="load_order")
with col_load2:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("載入雲端資料"):
        if not sheet_url or not google_creds:
            st.error("請在左側設定 Google Sheet 網址與 JSON！")
        else:
            with st.spinner("載入中..."):
                data = load_saved_data(google_creds, sheet_url, load_order)
                if data:
                    st.session_state.pack_data = data["pack"]
                    st.session_state.price_data = data["price"]
                    st.success("資料載入成功！請在下方預覽區確認。")
                else:
                    st.warning("找不到該番號的儲存資料。")

st.header("2. 預覽與編輯資料 (Preview & Edit)")
st.caption("您可以先利用「批量更正」功能把同一個錯字統一替換，或是直接在下方表格做個別修改與新增/刪除行。")

with st.expander("🛠️ 批量更正 品種 (Variety) 與 等級 (Grade) 名稱", expanded=False):
    all_vars = sorted(list(set([i.get('variety','') for i in st.session_state.pack_data] + [i.get('variety','') for i in st.session_state.price_data])))
    all_grades = sorted(list(set([i.get('grade','') for i in st.session_state.pack_data] + [i.get('grade','') for i in st.session_state.price_data])))
    
    colA, colB = st.columns(2)
    with colA:
        st.write("🍏 品種 (Variety) 更正")
        var_df = pd.DataFrame([{"目前名稱": v, "更正為": v} for v in all_vars if str(v).strip() != ""])
        new_vars = st.data_editor(var_df, use_container_width=True, hide_index=True) if not var_df.empty else pd.DataFrame()
        
    with colB:
        st.write("🏅 等級 (Grade) 更正")
        grade_df = pd.DataFrame([{"目前名稱": g, "更正為": g} for g in all_grades if str(g).strip() != ""])
        new_grades = st.data_editor(grade_df, use_container_width=True, hide_index=True) if not grade_df.empty else pd.DataFrame()

    if st.button("執行批量更正 (Apply Bulk Rename)"):
        var_map = {row["目前名稱"]: row["更正為"] for _, row in new_vars.iterrows()} if not new_vars.empty else {}
        grade_map = {row["目前名稱"]: row["更正為"] for _, row in new_grades.iterrows()} if not new_grades.empty else {}
        
        for item in st.session_state.pack_data:
            if item.get("variety") in var_map: item["variety"] = var_map[item["variety"]]
            if item.get("grade") in grade_map: item["grade"] = grade_map[item["grade"]]
        
        for item in st.session_state.price_data:
            if item.get("variety") in var_map: item["variety"] = var_map[item["variety"]]
            if item.get("grade") in grade_map: item["grade"] = grade_map[item["grade"]]
            
        st.success("批量更正成功！下方表格已更新。")
        st.rerun()

st.markdown("---")
pack_df = pd.DataFrame(st.session_state.pack_data)
if pack_df.empty:
    pack_df = pd.DataFrame(columns=["variety", "grade", "size", "quantity"])
edited_pack = st.data_editor(pack_df, num_rows="dynamic", key="pack_editor")

price_df = pd.DataFrame(st.session_state.price_data)
if price_df.empty:
    price_df = pd.DataFrame(columns=["variety", "grade", "size", "price"])
edited_price = st.data_editor(price_df, num_rows="dynamic", key="price_editor")


st.header("3. 產生檔案 (Generate Files)")

if st.button("生成 PDF 與寫入總表"):
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
                        update_google_sheet(google_creds, sheet_url, order_no, date_str, pack_list, price_list)
                        st.success("成功更新至 Google Sheet 總表並儲存歸檔！備註欄已上連結！")
                    except Exception as ge:
                        st.error(f"寫入 Google Sheet 失敗（請確認您的 JSON 和網址是否正確）: {ge}")
                else:
                    st.warning("您沒有填寫 Google Sheet 資訊，跳過總表更新。")
                    
                st.success("🎉 PDF 產生成功！請往下捲動下載檔案。")
                
            except Exception as e:
                st.error(f"產生過程中發生錯誤: {e}")

if st.session_state.pdf_generated and st.session_state.pl_bytes and st.session_state.inv_bytes:
    st.markdown("### 下載區 (Downloads)")
    st.download_button("📥 下載 Packing List PDF", st.session_state.pl_bytes, file_name=f"{order_no}_PackingList.pdf", mime="application/pdf")
    st.download_button("📥 下載 Invoice PDF", st.session_state.inv_bytes, file_name=f"{order_no}_Invoice.pdf", mime="application/pdf")
