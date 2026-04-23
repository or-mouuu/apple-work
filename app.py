import streamlit as st
import pandas as pd
import json
import os
from extractor import extract_pack_data, extract_price_data, pdf_to_images
from pdf_generator import generate_packing_list, generate_invoice, preprocess_invoice_data
from google_sheets_updater import update_google_sheet, load_saved_data

st.set_page_config(page_title="蘋果出貨文件生成工具", layout="wide")

st.title("🍎 蘋果出貨文件生成工具")

st.sidebar.header("🔐 系統解鎖")
app_password = st.sidebar.text_input("請輸入系統密碼", type="password")

if app_password == "unis5888":
    st.sidebar.success("授權成功！系統環境已自動載入。")
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        sheet_url = st.secrets["GOOGLE_SHEET_URL"]
        google_creds = st.secrets["GOOGLE_CREDS_JSON"]
        st.session_state.api_key_valid = True
    except Exception as e:
        st.sidebar.error("⚠️ 尚未設定好參數，請到後台設定 Secrets。")
        api_key, sheet_url, google_creds = "", "", ""
        st.session_state.api_key_valid = False
else:
    if app_password:
        st.sidebar.error("密碼錯誤！")
    st.info("👈 請先在左側輸入密碼以解鎖並使用工具。")
    st.stop()


st.header("第一步：匯入原始資料")
col1, col2 = st.columns(2)
with col1:
    pack_file = st.file_uploader("上傳「重量紀錄」掃描檔 (PDF)", type=["pdf"])
with col2:
    price_file = st.file_uploader("上傳「價格紀錄」掃描檔 (PDF)", type=["pdf"])
    
order_no = st.text_input("輸入 注文番號 (例如: USN 1031)")

if 'pack_data' not in st.session_state:
    st.session_state.pack_data = []
if 'row_totals' not in st.session_state:
    st.session_state.row_totals = []
if 'price_data' not in st.session_state:
    st.session_state.price_data = []
if 'pdf_generated' not in st.session_state:
    st.session_state.pdf_generated = False
if 'pl_bytes' not in st.session_state:
    st.session_state.pl_bytes = None
if 'inv_bytes' not in st.session_state:
    st.session_state.inv_bytes = None

if st.button("利用 AI 自動辨識資料"):
    if not api_key:
        st.error("請先在左側欄位確認系統設定！")
    elif not pack_file and not price_file:
         st.warning("請先上傳 PDF 檔案！")
    else:
        if pack_file:
            with st.spinner("辨識重量紀錄中..."):
                try:
                    pack_bytes = pack_file.read()
                    st.session_state.pack_pdf_bytes = pack_bytes
                    res = extract_pack_data(api_key, pack_bytes)
                    st.session_state.pack_data = res.get("pack_data", [])
                    st.session_state.row_totals = res.get("row_totals", [])
                    st.success("重量紀錄辨識成功！")
                except Exception as e:
                    st.error(f"辨識失敗: {e}")
        
        if price_file:
            with st.spinner("辨識價格紀錄中..."):
                try:
                    price_bytes = price_file.read()
                    st.session_state.price_pdf_bytes = price_bytes
                    res = extract_price_data(api_key, price_bytes, st.session_state.pack_data)
                    st.session_state.price_data = res
                    st.success("價格紀錄辨識成功！")
                except Exception as e:
                    st.error(f"辨識失敗: {e}")

st.markdown("<br>", unsafe_allow_html=True)
with st.expander("☁️ 載入歷史雲端資料 (若需調閱以前產生過的明細再點開此區塊)", expanded=False):
    st.caption("從 Google 總表中抓取您過去儲存過的所有數字。")
    col_load1, col_load2 = st.columns([3,1])
    with col_load1:
        load_order = st.text_input("輸入想調閱的 注文番號", key="load_order")
    with col_load2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("載入歷史紀錄"):
            if not sheet_url or not google_creds:
                st.error("未找到正確的連線設定！")
            else:
                with st.spinner("努力載入中..."):
                    data = load_saved_data(google_creds, sheet_url, load_order)
                    if data:
                        st.session_state.pack_data = data["pack"]
                        st.session_state.price_data = data["price"]
                        st.success("資料載入成功！請在下方預覽區確認。")
                    else:
                        st.warning("找不到該番號的儲存資料。")

st.markdown("---")
st.header("第二步：預覽與編輯結果")
st.caption("您可以先利用「批量更正」功能把同一個錯字統一替換，或者直接在下方表格做個別數字修改以及新增/刪除資料。")

with st.expander("🛠️ 批量更正「品種」與「等級」名稱", expanded=False):
    all_vars = sorted(list(set([i.get('variety','') for i in st.session_state.pack_data] + [i.get('variety','') for i in st.session_state.price_data])))
    all_grades = sorted(list(set([i.get('grade','') for i in st.session_state.pack_data] + [i.get('grade','') for i in st.session_state.price_data])))
    
    colA, colB = st.columns(2)
    with colA:
        st.write("🍏 品種統一更正")
        var_df = pd.DataFrame([{"目前名稱": v, "更正為": v} for v in all_vars if str(v).strip() != ""])
        new_vars = st.data_editor(var_df, use_container_width=True, hide_index=True) if not var_df.empty else pd.DataFrame()
        
    with colB:
        st.write("🏅 等級統一更正")
        grade_df = pd.DataFrame([{"目前名稱": g, "更正為": g} for g in all_grades if str(g).strip() != ""])
        new_grades = st.data_editor(grade_df, use_container_width=True, hide_index=True) if not grade_df.empty else pd.DataFrame()

    if st.button("執行統一更正"):
        var_map = {row["目前名稱"]: row["更正為"] for _, row in new_vars.iterrows()} if not new_vars.empty else {}
        grade_map = {row["目前名稱"]: row["更正為"] for _, row in new_grades.iterrows()} if not new_grades.empty else {}
        
        for item in st.session_state.pack_data:
            if item.get("variety") in var_map: item["variety"] = var_map[item["variety"]]
            if item.get("grade") in grade_map: item["grade"] = grade_map[item["grade"]]
        
        for item in st.session_state.price_data:
            if item.get("variety") in var_map: item["variety"] = var_map[item["variety"]]
            if item.get("grade") in grade_map: item["grade"] = grade_map[item["grade"]]
            
        st.success("批量更正成功！下方表格已同步更新。")
        st.rerun()

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 📦 重量紀錄")
pack_df = pd.DataFrame(st.session_state.pack_data)
if pack_df.empty:
    pack_df = pd.DataFrame(columns=["variety", "grade", "size", "quantity"])

col_pack_img, col_pack_data = st.columns(2)

with col_pack_img:
    st.caption("📄 原始掃描檔")
    if st.session_state.get('pack_pdf_bytes'):
        try:
            images = pdf_to_images(st.session_state.pack_pdf_bytes)
            for img in images:
                st.image(img, use_container_width=True)
        except Exception as e:
            st.error("圖片載入失敗")
    else:
        st.info("尚無上傳的掃描檔。")

with col_pack_data:
    st.caption("✏️ AI 辨識結果 (可編輯)")
    edited_pack = st.data_editor(pack_df, num_rows="dynamic", key="pack_editor", use_container_width=True, height=600)

if st.session_state.get('row_totals'):
    st.markdown("### 🔍 出荷數加總驗證")
    st.caption("驗證您目前表格中的加總與掃描圖上的出荷數是否一致。")
    
    temp_pack = edited_pack.copy()
    temp_pack['quantity'] = pd.to_numeric(temp_pack['quantity'], errors='coerce').fillna(0)
    calc_df = temp_pack.groupby(['variety', 'grade'], sort=False)['quantity'].sum().reset_index()
    totals_df = pd.DataFrame(st.session_state.row_totals)
    
    if not totals_df.empty:
        merged_df = pd.merge(calc_df, totals_df, on=['variety', 'grade'], how='outer')
        merged_df.rename(columns={'quantity': '表格加總 (Quantity)', 'expected_total': '原始出荷數 (Expected)'}, inplace=True)
        merged_df['表格加總 (Quantity)'] = merged_df['表格加總 (Quantity)'].fillna(0).astype(int)
        merged_df['原始出荷數 (Expected)'] = merged_df['原始出荷數 (Expected)'].fillna(0).astype(int)
        
        # 加上合計列
        total_calc = merged_df['表格加總 (Quantity)'].sum()
        total_exp = merged_df['原始出荷數 (Expected)'].sum()
        total_row = pd.DataFrame([{
            'variety': '合計', 
            'grade': '', 
            '表格加總 (Quantity)': total_calc, 
            '原始出荷數 (Expected)': total_exp
        }])
        merged_df = pd.concat([merged_df, total_row], ignore_index=True)
        
        def highlight_diff(row):
            if row['表格加總 (Quantity)'] != row['原始出荷數 (Expected)']:
                return ['background-color: #ffcccc'] * len(row)
            return ['background-color: #ccffcc'] * len(row)
            
        st.dataframe(merged_df.style.apply(highlight_diff, axis=1), use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("### 💰 價格紀錄")

price_df = pd.DataFrame(st.session_state.price_data)
if price_df.empty:
    price_df = pd.DataFrame(columns=["variety", "grade", "size", "price"])

col_price_img, col_price_data = st.columns(2)

with col_price_img:
    st.caption("📄 原始掃描檔")
    if st.session_state.get('price_pdf_bytes'):
        try:
            images = pdf_to_images(st.session_state.price_pdf_bytes)
            for img in images:
                st.image(img, use_container_width=True)
        except Exception as e:
            st.error("圖片載入失敗")
    else:
        st.info("尚無上傳的掃描檔。")

with col_price_data:
    st.caption("✏️ AI 辨識結果 (可編輯)")
    edited_price = st.data_editor(price_df, num_rows="dynamic", key="price_editor", use_container_width=True, height=600)
    
    # Missing Price Warning
    try:
        current_pack = edited_pack.to_dict('records')
        current_price = edited_price.to_dict('records')
        processed_for_warning = preprocess_invoice_data(current_pack, current_price)
        missing_prices = [item for item in processed_for_warning if item.get('_price', 0) == 0]
        
        if missing_prices:
            missing_names = sorted(list(set([f"{m.get('variety')} {m.get('grade')}" for m in missing_prices])))
            st.warning(f"⚠️ **注意：以下品項目前缺少價格對應，計算將會是 ¥0：**\n\n" + ", ".join(missing_names) + "\n\n您可以選擇在上方表格手動新增對應的價格，或是勾選下方「在產生 Invoice 時自動排除缺少價格的品項」。")
    except Exception as e:
        pass


st.markdown("---")
st.header("第三步：產生最終檔案與寫入總表")

st.markdown("### 📝 封面資訊 (Cover Page Details)")
with st.expander("展開編輯封面資訊 (預設為截圖中的文字)", expanded=False):
    import datetime
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown("**SHIPPER**")
        shipper_name = st.text_input("Name", "UNIS CO.,LTD.", key="shipper_name")
        shipper_addr1 = st.text_input("Address 1", "3-4-16 DENEN HIROSAKI CITY AOMORI JAPAN", key="shipper_addr1")
        shipper_addr2 = st.text_input("Address 2", "036-8086", key="shipper_addr2")
        shipper_tel = st.text_input("TEL", "+81-172-55-8975", key="shipper_tel")
        shipper_fax = st.text_input("FAX", "+81-172-55-8976", key="shipper_fax")
        
        st.markdown("**CONSIGNEE**")
        consignee_name = st.text_input("Consignee Name", "S.N.K. TRADING CO.,LTD.", key="consignee_name")
        consignee_addr1 = st.text_input("Consignee Address 1", "11F., NO.131, FUCHENG 2ND ST., FENGSHAN DIST.,", key="consignee_addr1")
        consignee_addr2 = st.text_input("Consignee Address 2", "KAOHSIUNG 830640 TAIWAN", key="consignee_addr2")
        consignee_tel = st.text_input("Consignee TEL", "+886-7-8117189", key="consignee_tel")
        consignee_fax = st.text_input("Consignee FAX", "+886-7-8117189", key="consignee_fax")
        
    with col_c2:
        date = st.text_input("DATE", datetime.date.today().strftime('%Y/%m/%d'), key="date")
        booking_agent = st.text_input("BOOKING AGENT", "WAN HAI LINES", key="booking_agent")
        booking_no = st.text_input("BOOKING NO.", "008 EA 22992", key="booking_no")
        shipped_per = st.text_input("SHIPPED PER MV", "WAN HAI 376 S004", key="shipped_per")
        from_port = st.text_input("FROM", "YOKOHAMA", key="from_port")
        to_port = st.text_input("TO", "KEELUNG", key="to_port")
        on_or_about = st.text_input("ON OR ABOUT", "2024/11/2", key="on_or_about")
        
        st.markdown("**REMARKS**")
        origin = st.text_input("ORIGIN", "AOMORI", key="origin")
        brand = st.text_input("BRAND", "SHICHIFUKUJIN", key="brand")
        pallet_count = st.text_input("PALLET (數量)", "21", key="pallet_count")
        
    # pallet_weight_total moved out
cover_info = {
    "shipper_name": shipper_name,
    "shipper_addr1": shipper_addr1,
    "shipper_addr2": shipper_addr2,
    "shipper_tel": shipper_tel,
    "shipper_fax": shipper_fax,
    "consignee_name": consignee_name,
    "consignee_addr1": consignee_addr1,
    "consignee_addr2": consignee_addr2,
    "consignee_tel": consignee_tel,
    "consignee_fax": consignee_fax,
    "date": date,
    "booking_agent": booking_agent,
    "booking_no": booking_no,
    "shipped_per": shipped_per,
    "from_port": from_port,
    "to_port": to_port,
    "on_or_about": on_or_about,
    "origin": origin,
    "brand": brand,
    "pallet": pallet_count,
    "pallet_weight": pallet_weight_total
}

col_w1, col_w2 = st.columns(2)
with col_w1:
    case_weight = st.number_input("設定每櫃淨重 (Net Weight / Case), 預設 11kg", value=11.0, step=0.1)
with col_w2:
    pallet_weight_total = st.number_input("Other packaging materials (pallet etc.) KG", value=189.0, step=1.0)

exclude_zero = st.checkbox("在產生 Invoice 時自動排除缺少價格 (¥0) 的品項", value=True)

if st.button("生成 PDF 檔案並更新總表"):
    if edited_pack.empty:
        st.error("沒有重量紀錄資料，無法產生檔案。")
    elif not order_no:
        st.error("請在上方輸入注文番號！")
    else:
        with st.spinner("產生檔案中..."):
            try:
                pack_list = edited_pack.to_dict('records')
                price_list = edited_price.to_dict('records')
                
                os.makedirs("output", exist_ok=True)
                pl_path = os.path.join("output", "PackingList_Output.pdf")
                inv_path = os.path.join("output", "Invoice_Output.pdf")
                
                generate_packing_list(pack_list, order_no, case_weight, cover_info, pl_path)
                generate_invoice(pack_list, price_list, order_no, cover_info, inv_path, exclude_zero_price=exclude_zero)
                
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
                        st.success("成功更新至 Google 總表並儲存歸檔！")
                    except Exception as ge:
                        st.error(f"寫入總表失敗（請聯絡開發人員確認憑證是否過期）: {ge}")
                else:
                    st.warning("偵測不到 Google 連線資訊，自動跳過總表更新。")
                    
                st.success("🎉 PDF 產生完畢！請往下捲動下載檔案。")
                
            except Exception as e:
                st.error(f"產生過程中發生意料之外的錯誤: {e}")

if st.session_state.pdf_generated and st.session_state.pl_bytes and st.session_state.inv_bytes:
    st.markdown("### 檔案下載區")
    st.download_button("📥 下載 Packing List (裝箱單)", st.session_state.pl_bytes, file_name=f"{order_no}_PackingList.pdf", mime="application/pdf")
    st.download_button("📥 下載 Invoice (明細單)", st.session_state.inv_bytes, file_name=f"{order_no}_Invoice.pdf", mime="application/pdf")
