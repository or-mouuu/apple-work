import os
import io
import json
from PIL import Image
import pypdfium2 as pdfium
from google import genai
from google.genai import types

def pdf_to_images(pdf_bytes):
    pdf = pdfium.PdfDocument(pdf_bytes)
    images = []
    for i in range(len(pdf)):
        page = pdf[i]
        bitmap = page.render(scale=2)
        images.append(bitmap.to_pil())
    return images

def extract_pack_data(api_key, pdf_bytes):
    try:
        images = pdf_to_images(pdf_bytes)
        client = genai.Client(api_key=api_key)
        
        prompt = """
        This is a shipping advice (pack-sample) tracking apple qualities and quantities.
        The columns usually contain grades (等級, e.g. 勝, 赤特選, 黒特選) and variety (品名, e.g. シナノスイート, 名月).
        The size columns are typically numbers like 20玉, 22玉, 24玉, etc.
        
        CRITICAL TABLE ALIGNMENT TASK:
        Because handwriting and tables can be hard to read, you must first do a step-by-step reasoning in the "row_reasoning" field.
        1. Identify the exact header sizes. COUNT exactly how many size columns there are and carefully list them all in order (e.g., 20, 22, 24, 26, 28, 32, 36, 40, 46, 50, 56).
        2. Pay intense attention to the VERTICAL alignment of numbers under these headers.
        3. Do NOT mix up rows. Read STRICTLY horizontally between the horizontal grid lines.
        4. For EVERY single row, explicitly trace each column from left to right, STARTING FROM THE VERY FIRST SIZE COLUMN (e.g., 20).
        You MUST list out every single size header for every row, even if it is empty. 
        For example, if headers are 20, 22, 24, 26:
        "Row 1: 20=blank, 22=1, 24=1, 26=2..."
        "Row 2: 20=2, 22=4, 24=9, 26=26..."
        DO NOT skip the first column (e.g., 20) in your reasoning, especially if previous rows had it blank.
        5. A common error is shifting numbers left or right when some columns are blank. Check the vertical grid lines carefully.
        6. On the far right of the table, there is an "出荷数" (Total Quantity) column. Trace it for each row.
        7. Only after tracing all columns, put the final non-empty size/quantities into the "extracted_data" list, and the row total into the "row_totals" list.

        You must ONLY return a valid JSON object with the following schema:
        {
          "row_reasoning": [
            "string detailing the sizes found in each row"
          ],
          "extracted_data": [
            {
              "variety": "string",
              "grade": "string",
              "size": "string (just the number, e.g., '22')",
              "quantity": "integer"
            }
          ],
          "row_totals": [
            {
              "variety": "string",
              "grade": "string",
              "expected_total": "integer (from the 出荷数 column)"
            }
          ]
        }
        """
        
        # Pass all pages to the model
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[*images, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            )
        )
        data = json.loads(response.text)
        return {
            "pack_data": data.get("extracted_data", []),
            "row_totals": data.get("row_totals", [])
        }
    except Exception as e:
        raise Exception(f"Failed to extract pack data: {str(e)}")

def extract_price_data(api_key, pdf_bytes, pack_data=None):
    try:
        images = pdf_to_images(pdf_bytes)
        client = genai.Client(api_key=api_key)
        
        reference_notes = ""
        if pack_data:
            varieties = list(set([item.get('variety') for item in pack_data if item.get('variety')]))
            grades = list(set([item.get('grade') for item in pack_data if item.get('grade')]))
            reference_notes = f"""
        CRITICAL INSTRUCTION FOR HANDWRITING: 
        Because handwriting can be messy, please use the following lists of varieties and grades (extracted from the weight record) to match and correct the handwritten item names in the price record. Pick the closest match from these lists for "variety" and "grade" if possible:
        Expected Varieties: {varieties}
        Expected Grades: {grades}
        """

        prompt = f"""
        This is a handwritten price list (price-sample) tracking apple prices.
        CRITICAL INSTRUCTION: There are typically two tables on the page. The top one usually contains your purchase prices, and the BOTTOM one is the CNF table with actual selling prices.
        You MUST ONLY extract prices from the BOTTOM (CNF) table. Ignore the top table completely.
        {reference_notes}
        There are sections for different varieties (like USN-1031 Shinano Sweet vs others like 名月).
        The columns are sizes (e.g. 32pup, 36p, 40p).
        The rows are grades (e.g. 勝, 赤特選).
        Extract this into a flat JSON list. Each object in the list should represent the price for one size of one grade for a variety.
        If a size has 'up' or 'pup', just use that string.
        You must ONLY return the raw JSON array. Do not wrap it in markdown. Do not include any other text.
        Schema for each object:
        {{
          "variety": "string",
          "grade": "string",
          "size": "string (e.g., '32', '28p', '28pup')",
          "price": "integer"
        }}
        """
        
        # Pass all pages to the model
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[*images, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
            )
        )
        return json.loads(response.text)
    except Exception as e:
        raise Exception(f"Failed to extract price data: {str(e)}")
