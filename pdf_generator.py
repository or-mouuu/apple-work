from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from collections import defaultdict
import datetime

def generate_packing_list(data, order_no, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    
    # Simple Title Header
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2.0, height - 2*cm, "Attached sheet for P/L")
    
    c.setFont("Helvetica", 10)
    c.drawString(2*cm, height - 3*cm, f"REF.NO.: {order_no}")
    c.drawRightString(width - 2*cm, height - 3*cm, f"DATE: {datetime.date.today().strftime('%Y/%m/%d')}")
    c.drawRightString(width - 2*cm, height - 3.5*cm, f"PAGE: 1/1")
    
    c.line(2*cm, height - 3.7*cm, width - 2*cm, height - 3.7*cm)
    c.line(2*cm, height - 3.8*cm, width - 2*cm, height - 3.8*cm)
    
    c.drawString(2*cm, height - 4.5*cm, "FRESH APPLE")
    c.drawString(2*cm, height - 5.0*cm, "NO MARK")
    
    c.drawString(4*cm, height - 4.2*cm, "MARKS AND NOS")
    c.drawString(8*cm, height - 4.2*cm, "DESCRIPTION")
    c.drawString(11*cm, height - 4.2*cm, "QUANTITY (CASE)")
    c.drawString(14*cm, height - 4.2*cm, "NET")
    c.drawString(17*cm, height - 4.2*cm, "GROSS")
    
    y = height - 6.0*cm
    
    # Sort and Print
    # Combine same variety, grade, size
    for item in data:
        y -= 0.5*cm
        if y < 2*cm:
            c.showPage()
            y = height - 2*cm
        
        c.drawString(2*cm, y, str(item.get('variety', '')))
        c.drawString(5*cm, y, str(item.get('grade', '')))
        c.drawString(8*cm, y, str(item.get('size', '')) + " p")
        qty = item.get('quantity', 0)
        c.drawString(11*cm, y, str(qty))
        c.drawString(14*cm, y, f"{qty * 11.5:.1f}")
        c.drawString(17*cm, y, f"{qty * 13.0:.1f}")

    c.line(2*cm, y - 0.5*cm, width - 2*cm, y - 0.5*cm)
    c.save()
    return output_path

def generate_invoice(data, price_data, order_no, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2.0, height - 2*cm, "Attached sheet for I/V")
    
    c.setFont("Helvetica", 10)
    c.drawString(2*cm, height - 3*cm, f"REF.NO.: {order_no}")
    c.drawRightString(width - 2*cm, height - 3*cm, f"DATE: {datetime.date.today().strftime('%Y/%m/%d')}")
    c.drawRightString(width - 2*cm, height - 3.5*cm, f"PAGE: 1/1")
    
    c.line(2*cm, height - 3.7*cm, width - 2*cm, height - 3.7*cm)
    c.line(2*cm, height - 3.8*cm, width - 2*cm, height - 3.8*cm)
    
    c.drawString(2*cm, height - 4.5*cm, "FRESH APPLE")
    c.drawString(2*cm, height - 5.0*cm, "NO MARK")
    c.drawString(14*cm, height - 5.0*cm, "C&F: Keelung")
    
    c.drawString(4*cm, height - 4.2*cm, "MARKS AND NOS")
    c.drawString(8*cm, height - 4.2*cm, "DESCRIPTION")
    c.drawString(11*cm, height - 4.2*cm, "QUANTITY")
    c.drawString(14*cm, height - 4.2*cm, "UNIT PRICE")
    c.drawString(17*cm, height - 4.2*cm, "AMOUNT")
    
    y = height - 6.0*cm
    
    total_qty = 0
    total_amt = 0
    
    # Map prices
    price_map = {}
    for p in price_data:
        k = f"{p.get('variety')}_{p.get('grade')}_{p.get('size')}"
        price_map[k] = p.get('price', 0)
    
    for item in data:
        y -= 0.5*cm
        if y < 2*cm:
            c.showPage()
            y = height - 2*cm
            
        var = str(item.get('variety', ''))
        grade = str(item.get('grade', ''))
        size = str(item.get('size', ''))
        qty = item.get('quantity', 0)
        
        # Look for price mapping
        p_val = price_map.get(f"{var}_{grade}_{size}", 0)
        # simplistic fallback
        if p_val == 0:
            for p in price_data:
                if p.get('variety') == var and p.get('grade') == grade:
                    p_val = p.get('price', 0)
                    break
                    
        amt = qty * p_val
        total_qty += qty
        total_amt += amt
        
        c.drawString(2*cm, y, var)
        c.drawString(5*cm, y, grade)
        c.drawString(8*cm, y, size + " p")
        c.drawString(11*cm, y, f"{qty} cs")
        c.drawString(14*cm, y, f"¥{p_val:,.0f}")
        c.drawString(17*cm, y, f"¥{amt:,.0f}")
        
    c.line(2*cm, y - 0.5*cm, width - 2*cm, y - 0.5*cm)
    c.drawString(8*cm, y - 1*cm, "TOTAL")
    c.drawString(11*cm, y - 1*cm, f"{total_qty} cs")
    c.drawString(17*cm, y - 1*cm, f"¥{total_amt:,.0f}")

    c.save()
    return output_path
