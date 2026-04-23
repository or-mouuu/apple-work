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

def generate_packing_list(data, order_no, case_weight, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    cjk_font = register_font()
    
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
    data = sorted(data, key=lambda x: (str(x.get('variety', '')), str(x.get('grade', ''))))
    
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

def generate_invoice(data, price_data, order_no, output_path, exclude_zero_price=False):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    cjk_font = register_font()
    
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
    total_qty = 0
    total_amt = 0
    
    processed_data = preprocess_invoice_data(data, price_data)
    
    if exclude_zero_price:
        processed_data = [item for item in processed_data if item.get('_price', 0) > 0]
        
    processed_data = sorted(processed_data, key=lambda x: (str(x.get('variety', '')), str(x.get('grade', ''))))
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
        total_qty += qty
        total_amt += amt
        
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
