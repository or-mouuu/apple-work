import gspread
from google.oauth2.service_account import Credentials
import json
import pandas as pd

def authorize_gspread(json_creds_str):
    creds_dict = json.loads(json_creds_str)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def update_google_sheet(json_creds_str, sheet_url, order_no, date_str, pack_data, price_data):
    try:
        gc = authorize_gspread(json_creds_str)
        sh = gc.open_by_url(sheet_url)
        
        # 1. Save data to Saved_Data tab
        try:
            saved_ws = sh.worksheet("Saved_Data")
        except gspread.exceptions.WorksheetNotFound:
            saved_ws = sh.add_worksheet(title="Saved_Data", rows="1000", cols="5")
            saved_ws.append_row(["OrderNo", "Date", "PackJSON", "PriceJSON", "Link"])

        pack_json = json.dumps(pack_data, ensure_ascii=False)
        price_json = json.dumps(price_data, ensure_ascii=False)
        
        # Check if order_no already exists in Saved_Data to update or append
        order_col = []
        try:
            order_col = saved_ws.col_values(1)
        except:
            pass
            
        data_sheet_id = saved_ws.id
        row_index = len(order_col) + 1
        link_url = f"{sheet_url}#gid={data_sheet_id}&range=A{row_index}"
        
        if order_no in order_col:
            idx = order_col.index(order_no) + 1
            saved_ws.update_cell(idx, 2, date_str)
            saved_ws.update_cell(idx, 3, pack_json)
            saved_ws.update_cell(idx, 4, price_json)
            link_url = f"{sheet_url}#gid={data_sheet_id}&range=A{idx}"
        else:
            saved_ws.append_row([order_no, date_str, pack_json, price_json, link_url])
        
        # 2. Update Master Sheet (assumed to be sheet1)
        ws = sh.sheet1
        
        # Headers: 製作日期[1] 注文番號[2] 植檢申請番號[3] 進口商[4] 包裝廠[5] 品種類數[6] 重量[7] 備考[8]
        # Find next empty row in column A or B
        col_b = ws.col_values(2) # 注文番號 is col 2
        next_row = len(col_b) + 1
        if next_row < 2:
            next_row = 2
            
        # Calculate variety count and total weight
        varieties = set([item.get('variety', '') for item in pack_data])
        total_case = sum([int(item.get('quantity', 0)) for item in pack_data])
        total_weight = total_case * 11.5 # Net weight
        
        ws.update_cell(next_row, 1, date_str)       # 製作日期
        ws.update_cell(next_row, 2, order_no)       # 注文番號
        # 3, 4, 5 leave blank
        ws.update_cell(next_row, 6, len(varieties)) # 品種類數
        ws.update_cell(next_row, 7, f"{total_weight:.1f}") # 重量
        ws.update_cell(next_row, 8, f'=HYPERLINK("{link_url}", "View Data")') # 備考 (Link to Saved_Data)

        return True
    except Exception as e:
        raise Exception(f"Failed to update Google Sheet: {str(e)}")

def load_saved_data(json_creds_str, sheet_url, order_no):
    try:
        gc = authorize_gspread(json_creds_str)
        sh = gc.open_by_url(sheet_url)
        saved_ws = sh.worksheet("Saved_Data")
        records = saved_ws.get_all_records()
        for rec in records:
            if str(rec.get("OrderNo", "")) == order_no:
                return {
                    "pack": json.loads(rec.get("PackJSON", "[]")),
                    "price": json.loads(rec.get("PriceJSON", "[]"))
                }
        return None
    except Exception as e:
        return None
