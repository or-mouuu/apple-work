import gspread
from google.oauth2.service_account import Credentials
import json

def update_google_sheet(json_creds_str, sheet_url, order_no, date_str, total_quantity):
    try:
        creds_dict = json.loads(json_creds_str)
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(credentials)
        
        # Open by url
        sh = gc.open_by_url(sheet_url)
        worksheet = sh.sheet1 # or explicitly get by rules
        
        # We need to find the first empty row in Column A to write the order_no
        col_a = worksheet.col_values(1)
        next_row = len(col_a) + 1
        
        worksheet.update_cell(next_row, 1, order_no)
        worksheet.update_cell(next_row, 2, date_str)
        # Update more columns if necessary
        return True
    except Exception as e:
        raise Exception(f"Failed to update Google Sheet: {str(e)}")
