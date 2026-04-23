from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import datetime
import os
import re

font_path = "NotoSansTC-Regular.ttf"

def register_font():
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont('NotoSansTC', font_path))
            return 'NotoSansTC'
        except Exception:
            pass
    return 'Helvetica'

def normalize(s):
    return str(s).replace(" ", "").replace("p", "").replace("up", "").lower()

def get_price(var, grade, size, price_data):
    n_grade = normalize(grade)
    n_var = normalize(var)
    n_size = normalize(size)
    
    target_full = n_var + n_grade
    
    # 1. Strict match on grade and size
    for p in price_data:
        p_var = normalize(p.get('variety', ''))
        if p_var and n_var and p_var != n_var: continue
            
        p_grade = normalize(p.get('grade', ''))
        p_size = normalize(p.get('size', ''))
        
        # Check grade match
        grade_match = False
        p_full = p_var + p_grade
        if p_grade and p_grade in target_full: grade_match = True
        elif target_full in p_full: grade_match = True
        
        if grade_match and p_size == n_size:
            return int(p.get('price', 0))
            
    # 2. Fallback: match only on grade and variety
    for p in price_data:
        p_var = normalize(p.get('variety', ''))
        if p_var and n_var and p_var != n_var: continue
            
        p_grade = normalize(p.get('grade', ''))
        p_full = p_var + p_grade
        
        if p_grade and p_grade in target_full:
            return int(p.get('price', 0))
        elif target_full in p_full:
            return int(p.get('price', 0))
            
    return 0

def draw_cover_page(c, doc_type, cover_info, totals, width, height, cjk_font):
    c.setFont("Times-Roman", 16)
    if doc_type == "PACKINGLIST":
        c.drawCentredString(width/2.0, height - 2.5*cm, "PACKINGLIST")
    else:
        c.drawCentredString(width/2.0, height - 2.5*cm, "I N V O I C E")
        
    c.setFont("Times-Roman", 10)
    c.setLineWidth(1)
    
    x1, x_mid, x_mid2, x2 = 1.5*cm, 10.5*cm, 15*cm, width - 1.5*cm
    y0 = height - 3*cm
    y1 = y0 - 3*cm
    y2 = y1 - 3.5*cm
    y3 = y2 - 3.5*cm
    
    # Horizontal lines across the whole table
    c.line(x1, y1, x2, y1)
    c.line(x1, y2, x2, y2)
    
    # Right column specific horizontal lines
    y_ref_bottom = y0 - 1*cm
    y_booking_bottom = y0 - 2*cm
    
    c.line(x_mid, y_ref_bottom, x2, y_ref_bottom)
    c.line(x_mid, y_booking_bottom, x2, y_booking_bottom)
    
    c.rect(x1, y3, x2 - x1, y0 - y3)
    c.line(x_mid, y3, x_mid, y0)
    c.line(x_mid2, y_booking_bottom, x_mid2, y0)
    
    c.drawString(x1 + 0.1*cm, y0 - 0.4*cm, "SHIPPER:")
    c.drawString(x1 + 0.3*cm, y0 - 0.8*cm, cover_info.get("shipper_name", ""))
    c.drawString(x1 + 0.3*cm, y0 - 1.2*cm, cover_info.get("shipper_addr1", ""))
    c.drawString(x1 + 0.3*cm, y0 - 1.6*cm, cover_info.get("shipper_addr2", ""))
    c.drawString(x1 + 1.5*cm, y0 - 2.0*cm, "TEL: " + cover_info.get("shipper_tel", ""))
    c.drawString(x1 + 1.5*cm, y0 - 2.4*cm, "FAX: " + cover_info.get("shipper_fax", ""))
    
    c.drawString(x1 + 0.1*cm, y1 - 0.4*cm, "CONSIGNEE:")
    c.drawString(x1 + 0.3*cm, y1 - 0.8*cm, cover_info.get("consignee_name", ""))
    c.drawString(x1 + 0.3*cm, y1 - 1.2*cm, cover_info.get("consignee_addr1", ""))
    c.drawString(x1 + 0.3*cm, y1 - 1.6*cm, cover_info.get("consignee_addr2", ""))
    c.drawString(x1 + 1.5*cm, y1 - 2.0*cm, "TEL: " + cover_info.get("consignee_tel", ""))
    c.drawString(x1 + 1.5*cm, y1 - 2.4*cm, "FAX: " + cover_info.get("consignee_fax", ""))
    
    c.drawString(x1 + 0.1*cm, y2 - 0.4*cm, "NOTIFY PATY:")
    
    c.drawString(x_mid + 0.1*cm, y0 - 0.4*cm, "REF.NO.")
    c.drawCentredString((x_mid + x_mid2)/2, y0 - 0.8*cm, cover_info.get("order_no", ""))
    c.drawString(x_mid2 + 0.1*cm, y0 - 0.4*cm, "DATE")
    c.drawRightString(x2 - 0.2*cm, y0 - 0.8*cm, cover_info.get("date", ""))
    
    c.drawString(x_mid + 0.1*cm, y_ref_bottom - 0.4*cm, "BOOKING AGNET")
    c.drawCentredString((x_mid + x_mid2)/2, y_ref_bottom - 0.8*cm, cover_info.get("booking_agent", ""))
    c.drawString(x_mid2 + 0.1*cm, y_ref_bottom - 0.4*cm, "BOOKING NO.")
    c.drawRightString(x2 - 0.2*cm, y_ref_bottom - 0.8*cm, cover_info.get("booking_no", ""))
    
    c.drawString(x_mid + 0.1*cm, y_booking_bottom - 0.4*cm, "SHIPPED PER MV")
    c.drawCentredString(x_mid + 3.5*cm, y_booking_bottom - 0.4*cm, cover_info.get("shipped_per", ""))
    c.drawString(x_mid + 0.5*cm, y_booking_bottom - 0.9*cm, "FROM: " + cover_info.get("from_port", ""))
    c.drawString(x_mid2 + 0.5*cm, y_booking_bottom - 0.9*cm, "TO: " + cover_info.get("to_port", ""))
    c.drawString(x_mid + 0.1*cm, y1 + 0.4*cm, "ON OR ABOUT")
    c.drawCentredString(x_mid + 3.5*cm, y1 + 0.1*cm, cover_info.get("on_or_about", ""))
    
    c.drawString(x_mid + 0.1*cm, y1 - 0.4*cm, "REMARKS:")
    c.drawString(x_mid + 1*cm, y1 - 0.9*cm, "ORIGIN: " + cover_info.get("origin", ""))
    c.drawString(x_mid + 1*cm, y1 - 1.4*cm, "BRAND: " + cover_info.get("brand", ""))
    c.drawString(x_mid2 + 1*cm, y1 - 1.4*cm, "ICE BOX")
    c.drawString(x_mid + 1*cm, y1 - 1.9*cm, "PALLET: " + str(cover_info.get("pallet", "")))
    
    c.drawString(x_mid + 0.1*cm, y2 - 0.4*cm, "ALSO NOTIFY:")
    
    y_tbl = y3 - 0.5*cm
    c.line(x1, y_tbl, x2, y_tbl)
    c.line(x1, y_tbl - 0.1*cm, x2, y_tbl - 0.1*cm)
    
    c.drawString(x1 + 1*cm, y_tbl - 0.5*cm, "MARKS AND NOS")
    c.drawCentredString(x1 + 7*cm, y_tbl - 0.5*cm, "DESCRIPTION")
    
    if doc_type == "PACKINGLIST":
        c.drawCentredString(x1 + 10*cm, y_tbl - 0.5*cm, "QUANTITY")
        c.drawCentredString(x1 + 13.5*cm, y_tbl - 0.5*cm, "NET")
        c.drawCentredString(x1 + 16.5*cm, y_tbl - 0.5*cm, "GROSS")
    else:
        c.drawCentredString(x1 + 10*cm, y_tbl - 0.5*cm, "QUANTITY")
        c.drawCentredString(x1 + 13.5*cm, y_tbl - 0.5*cm, "UNIT PRICE")
        c.drawCentredString(x1 + 16.5*cm, y_tbl - 0.5*cm, "AMOUNT")
        
    c.line(x1, y_tbl - 0.7*cm, x2, y_tbl - 0.7*cm)
    
    c.drawString(x1, y_tbl - 1.3*cm, "FRESH APPLE")
    c.drawString(x1 + 1*cm, y_tbl - 1.8*cm, "NO MARK")
    
    if doc_type == "PACKINGLIST":
        c.drawCentredString(x1 + 7*cm, y_tbl - 1.8*cm, "KG")
        c.drawCentredString(x1 + 10*cm, y_tbl - 1.8*cm, "CASE")
        c.drawCentredString(x1 + 13.5*cm, y_tbl - 1.8*cm, "KG")
        c.drawCentredString(x1 + 16.5*cm, y_tbl - 1.8*cm, "KG")
    else:
        c.setFont("Times-Bold", 10)
        c.drawCentredString(x1 + 15*cm, y_tbl - 1.3*cm, "C&F: Keelung")
        c.line(x1 + 12*cm, y_tbl - 1.4*cm, x2, y_tbl - 1.4*cm)
        c.setFont("Times-Roman", 10)
        
    c.drawCentredString(x1 + 7*cm, y_tbl - 7*cm, "As per attachment")
    
    y_tot = 5*cm
    c.line(x1, y_tot, x2, y_tot)
    c.setFont("Times-Bold", 10)
    c.drawString(x1 + 1*cm, y_tot - 0.5*cm, "TOTAL")
    
    if doc_type == "PACKINGLIST":
        c.drawCentredString(x1 + 10*cm, y_tot - 0.5*cm, f"{totals.get('qty', 0)}")
        c.drawCentredString(x1 + 13.5*cm, y_tot - 0.5*cm, f"{totals.get('net', 0):.1f}")
        c.drawCentredString(x1 + 16.5*cm, y_tot - 0.5*cm, f"{totals.get('gross', 0):.1f}")
        
        c.setFont("Times-Roman", 10)
        c.drawString(x1 + 7*cm, y_tot - 1.5*cm, "Pallte and wrapping material etc.")
        c.drawCentredString(x1 + 16.5*cm, y_tot - 1.5*cm, f"{totals.get('pallet', 0):.1f}")
        c.line(x1 + 6.5*cm, y_tot - 1.7*cm, x2, y_tot - 1.7*cm)
        
        c.setFont("Times-Bold", 10)
        c.drawString(x1 + 7*cm, y_tot - 2.2*cm, "TOTAL GROSS WEIGHT")
        c.drawCentredString(x1 + 16.5*cm, y_tot - 2.2*cm, f"{totals.get('gross', 0) + totals.get('pallet', 0):.1f}")
    else:
        c.drawCentredString(x1 + 10*cm, y_tot - 0.5*cm, f"{totals.get('qty', 0)}")
        c.drawCentredString(x1 + 16.5*cm, y_tot - 0.5*cm, f"¥{totals.get('amount', 0):,.0f}")
        
    c.showPage()


def generate_packing_list(data, order_no, case_weight, cover_info, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    cjk_font = register_font()
    
    total_qty = sum(int(item.get('quantity', 0)) for item in data)
    total_net = total_qty * case_weight
    total_gross = total_qty * (case_weight + 1.0)
    
    totals = {
        'qty': total_qty,
        'net': total_net,
        'gross': total_gross,
        'pallet': float(cover_info.get("pallet_weight", 0))
    }
    
    cover_info['order_no'] = order_no
    draw_cover_page(c, "PACKINGLIST", cover_info, totals, width, height, cjk_font)
    
    c.setFont("Times-Bold", 16)
    c.drawCentredString(width/2.0, height - 2*cm, "Attached sheet for P/L")
    
    c.setFont("Times-Roman", 10)
    c.drawString(2*cm, height - 3*cm, f"REF.NO.: {order_no}")
    c.drawRightString(width - 2*cm, height - 3*cm, f"DATE: {datetime.date.today().strftime('%Y/%m/%d')}")
    c.drawRightString(width - 2*cm, height - 3.5*cm, f"PAGE: 1/1")
    
    c.setLineWidth(1)
    c.line(2*cm, height - 3.7*cm, width - 2*cm, height - 3.7*cm)
    c.line(2*cm, height - 3.8*cm, width - 2*cm, height - 3.8*cm)
    
    c.drawString(3.5*cm, height - 4.3*cm, "MARKS AND NOS")
    c.drawCentredString(10*cm, height - 4.3*cm, "DESCRIPTION")
    c.drawCentredString(13.5*cm, height - 4.3*cm, "QUANTITY (CASE)")
    c.drawCentredString(16*cm, height - 4.3*cm, "NET")
    c.drawCentredString(18.5*cm, height - 4.3*cm, "GROSS")
    
    c.line(2*cm, height - 4.5*cm, width - 2*cm, height - 4.5*cm)
    c.drawString(2*cm, height - 5.0*cm, "FRESH APPLE")
    c.drawString(2*cm, height - 5.5*cm, "NO MARK")
    
    y = height - 6.5*cm
    
    last_combo = None
    
    for item in data:
        if y < 2*cm:
            c.showPage()
            c.setFont("Times-Roman", 10)
            y = height - 2*cm
            
        var = str(item.get('variety', '')).strip()
        grade = str(item.get('grade', '')).strip()
        combo = f"{var} {grade}".strip()
        size = str(item.get('size', ''))
        qty = int(item.get('quantity', 0))
        net = qty * case_weight
        gross = qty * (case_weight + 1.0)
        
        if combo != last_combo:
            c.setDash(1, 2)
            c.line(2*cm, y - 0.1*cm, 7.5*cm, y - 0.1*cm)
            c.setDash()
            
            c.setFont(cjk_font, 10)
            c.drawRightString(7.5*cm, y, combo)
            last_combo = combo
            
        c.setFont("Times-Roman", 10)
        c.drawCentredString(10*cm, y, f"{size} p")
        c.drawCentredString(13.5*cm, y, str(qty))
        c.drawCentredString(16*cm, y, f"{net:.1f}")
        c.drawCentredString(18.5*cm, y, f"{gross:.1f}")
        
        y -= 0.6*cm

    c.setDash()
    c.line(2*cm, y - 0.2*cm, width - 2*cm, y - 0.2*cm)
    c.save()
    return output_path

def preprocess_invoice_data(data, price_data):
    def find_best_rule(var, grade, size, price_list):
        n_grade = normalize(grade)
        n_var = normalize(var)
        n_size = normalize(size)
        target_full = n_var + n_grade
        
        matched_prices = []
        for p in price_list:
            p_var = normalize(p.get('variety', ''))
            if p_var and n_var and p_var != n_var: continue
                
            p_grade = normalize(p.get('grade', ''))
            p_full = p_var + p_grade
            
            if p_grade and p_grade in target_full:
                matched_prices.append(p)
            elif target_full in p_full:
                matched_prices.append(p)
                
        for p in matched_prices:
            p_size = normalize(p.get('size', ''))
            if p_size == n_size:
                return p, "exact"
                
        try:
            m = re.search(r'\d+', str(size))
            s_val = int(m.group()) if m else -1
        except:
            s_val = -1
            
        best_pup_rule = None
        best_max = 9999
        
        for p in matched_prices:
            raw_s = str(p.get('size', '')).lower()
            if 'up' in raw_s:
                try:
                    m = re.search(r'\d+', raw_s)
                    if m:
                        p_val = int(m.group())
                        if s_val != -1 and s_val <= p_val:
                            if p_val < best_max:
                                best_max = p_val
                                best_pup_rule = p
                except:
                    pass
        if best_pup_rule:
            return best_pup_rule, "pup"
            
        return None, "none"

    grouped = {}
    for item in data:
        var = str(item.get('variety', '')).strip()
        grade = str(item.get('grade', '')).strip()
        size = str(item.get('size', ''))
        qty = int(item.get('quantity', 0))
        
        rule, rule_type = find_best_rule(var, grade, size, price_data)
        
        if rule and rule_type == "pup":
            rule_size = str(rule.get('size', ''))
            key = (var, grade, f"pup_group_{rule_size}")
            if key not in grouped:
                grouped[key] = {
                    'variety': var, 'grade': grade, 
                    'sizes_in_group': [], 'quantity': 0, 'price': int(rule.get('price', 0)), 'is_pup': True
                }
            m = re.search(r'\d+', size)
            if m:
                grouped[key]['sizes_in_group'].append(int(m.group()))
            grouped[key]['quantity'] += qty
        else:
            key = (var, grade, f"exact_{size}")
            if key not in grouped:
                price_val = int(rule.get('price', 0)) if rule else get_price(var, grade, size, price_data)
                grouped[key] = {
                    'variety': var, 'grade': grade, 'size_display': f"{size} p",
                    'quantity': 0, 'price': price_val, 'is_pup': False
                }
            grouped[key]['quantity'] += qty

    result = []
    for k, v in grouped.items():
        if v.get('is_pup'):
            sizes = sorted(v['sizes_in_group'])
            if sizes:
                min_s = sizes[0]
                max_s = sizes[-1]
                if min_s == max_s:
                    v['size_display'] = f"{min_s} p"
                else:
                    v['size_display'] = f"{min_s}p - {max_s}p"
            else:
                v['size_display'] = "pup"
        
        result.append({
            'variety': v['variety'],
            'grade': v['grade'],
            'size': v['size_display'],
            'quantity': v['quantity'],
            '_price': v['price']
        })
    return result

def generate_invoice(data, price_data, order_no, cover_info, output_path, exclude_zero_price=False):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    cjk_font = register_font()
    
    processed_data = preprocess_invoice_data(data, price_data)
    if exclude_zero_price:
        processed_data = [item for item in processed_data if item.get('_price', 0) > 0]
        
    total_qty = sum(int(item.get('quantity', 0)) for item in processed_data)
    total_amt = sum(int(item.get('quantity', 0)) * int(item.get('_price', 0)) for item in processed_data)
    
    totals = {
        'qty': total_qty,
        'amount': total_amt
    }
    
    cover_info['order_no'] = order_no
    draw_cover_page(c, "INVOICE", cover_info, totals, width, height, cjk_font)
    
    c.setFont("Times-Bold", 16)
    c.drawCentredString(width/2.0, height - 2*cm, "Attached sheet for I/V")
    
    c.setFont("Times-Roman", 10)
    c.drawString(2*cm, height - 3*cm, f"REF.NO.: {order_no}")
    c.drawRightString(width - 2*cm, height - 3*cm, f"DATE: {datetime.date.today().strftime('%Y/%m/%d')}")
    c.drawRightString(width - 2*cm, height - 3.5*cm, f"PAGE: 1/1")
    
    c.setLineWidth(1)
    c.line(2*cm, height - 3.7*cm, width - 2*cm, height - 3.7*cm)
    c.line(2*cm, height - 3.8*cm, width - 2*cm, height - 3.8*cm)
    
    c.drawString(3.5*cm, height - 4.3*cm, "MARKS AND NOS")
    c.drawCentredString(10*cm, height - 4.3*cm, "DESCRIPTION")
    c.drawCentredString(13*cm, height - 4.3*cm, "QUANTITY")
    c.drawCentredString(15.5*cm, height - 4.3*cm, "UNIT PRICE")
    c.drawRightString(19*cm, height - 4.3*cm, "AMOUNT")
    
    c.line(2*cm, height - 4.5*cm, width - 2*cm, height - 4.5*cm)
    c.drawString(2*cm, height - 5.0*cm, "FRESH APPLE")
    c.drawString(2*cm, height - 5.5*cm, "NO MARK")
    c.drawRightString(width - 2*cm, height - 5.0*cm, "C&F : Keelung")
    c.line(16*cm, height - 5.2*cm, width - 2*cm, height - 5.2*cm)
    
    y = height - 6.5*cm
    
    last_combo = None
    
    for item in processed_data:
        if y < 2*cm:
            c.showPage()
            c.setFont("Times-Roman", 10)
            y = height - 2*cm
            
        var = str(item.get('variety', '')).strip()
        grade = str(item.get('grade', '')).strip()
        combo = f"{var} {grade}".strip()
        display_size = str(item.get('size', ''))
        qty = int(item.get('quantity', 0))
        
        p_val = item.get('_price', 0)
        amt = qty * p_val
        
        if combo != last_combo:
            c.setDash(1, 2)
            c.line(2*cm, y - 0.1*cm, 7.5*cm, y - 0.1*cm)
            c.setDash()
            
            c.setFont(cjk_font, 10)
            c.drawRightString(7.5*cm, y, combo)
            last_combo = combo
            
        c.setFont("Times-Roman", 10)
        c.drawCentredString(10*cm, y, display_size)
        c.drawCentredString(13*cm, y, f"{qty} cs")
        c.drawCentredString(15.5*cm, y, f"¥{p_val:,.0f}")
        c.drawRightString(19*cm, y, f"¥{amt:,.0f}")
        
        y -= 0.6*cm
        
    c.line(2*cm, y - 0.2*cm, width - 2*cm, y - 0.2*cm)
    c.drawString(10*cm, y - 0.8*cm, "TOTAL")
    c.drawCentredString(13*cm, y - 0.8*cm, f"{total_qty} cs")
    c.drawRightString(19*cm, y - 0.8*cm, f"¥{total_amt:,.0f}")

    c.save()
    return output_path
