import streamlit as st
import pdfplumber
import pandas as pd
import re
import datetime
import io
import os
import textwrap
import itertools
import traceback

# Web app design
st.set_page_config(
    page_title="Virbac Statement Converter", 
    page_icon="📄", 
    layout="wide"
)

st.title("📄 Virbac Account Statement Converter (Smart Tracking)")
st.markdown("CFA Team sathi: PDF upload kara ani **Excel + PDF** donhi milva.")

uploaded_files = st.file_uploader(
    "Yethe PDF file upload kara", 
    type="pdf", 
    accept_multiple_files=True
)

def process_pdf_logic(uploaded_file):
    uploaded_file.seek(0)
    
    run_datetime_obj = datetime.datetime.now()
    run_datetime = run_datetime_obj.strftime("%d/%m/%Y %I:%M %p")
    
    period, customer_no, customer_name = "", "", ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            first_page = pdf.pages[0]
            text_layout = first_page.extract_text(layout=True) or ""
            clean_text = " ".join(text_layout.split())
            
            rx_period = r'(?:Account.*?date|period from).*?(\d{2}/\d{2}/\d{2,4})\s*(?:to|-)\s*(\d{2}/\d{2}/\d{2,4})'
            p_match = re.search(rx_period, clean_text, re.IGNORECASE)
            
            if p_match: 
                period = f"from {p_match.group(1)} To {p_match.group(2)}"
            
            rx_cust = r'(?:Payer|Customer No).*?(\d{6})'
            c_match = re.search(rx_cust, clean_text, re.IGNORECASE)
            
            if c_match: 
                customer_no = c_match.group(1).strip()
                for p in pdf.pages[:1]:
                    pg_text = p.extract_text() or ""
                    lines = [l.strip() for l in pg_text.split('\n') if l.strip()]
                    for i, line in enumerate(lines):
                        if customer_no in line:
                            end_idx = min(i + 7, len(lines))
                            for j in range(i + 1, end_idx):
                                candidate = lines[j].strip()
                                cand_lower = candidate.lower()
                                if len(candidate) < 4: 
                                    continue
                                    
                                ignore_keywords = [
                                    'date', 'time', 'page', 'statement', 
                                    'accounting', 'period', 'payer', 
                                    'customer', 'limit', 'opening', 
                                    'bal', 'dt', 'balance'
                                ]
                                
                                if any(x in cand_lower for x in ignore_keywords):
                                    continue
                                    
                                customer_name = candidate
                                break
                        if customer_name: 
                            break
    except Exception as e:
        err_msg = f"PDF वाचताना त्रुटी (Reading Error): {e}"
        return None, None, None, None, None, err_msg

    extracted_rows = []
    try:
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                p_text = page.extract_text()
                if p_text:
                    for line in p_text.split('\n'): 
                        extracted_rows.append(line.strip())
    except Exception as e:
        err_msg = f"PDF एक्स्ट्रॅक्ट करताना त्रुटी: {e}"
        return None, None, None, None, None, err_msg

    final_data = []
    running_balance = 0.0
    opening_balance = 0.0
    found_opening = False
    
    total_billed_dr = 0.0
    total_paid_cr = 0.0
    
    s_inv = 0.0
    pay = 0.0
    reco = 0.0
    c_oth = 0.0
    c_brk = 0.0
    g_ret = 0.0
    d_not = 0.0
    tcs = 0.0
    tds = 0.0
    tech_b = 0.0
    n_tech_b = 0.0

    try:
        for idx, row_text in enumerate(extracted_rows):
            row_upper = row_text.upper()
            
            if "OPENING BALANCE" in row_upper and not found_opening:
                amounts = re.findall(r'-?\s*\(?\s*[\d,]+\.\d{2}\s*\)?', row_text)
                if amounts:
                    amount_str = amounts[-1]
                    has_cr_str = 'CR' in row_upper
                    has_bracket = '(' in amount_str
                    has_minus = '-' in amount_str
                    is_cr = has_cr_str or has_bracket or has_minus
                    
                    if not is_cr:
                        for offset in range(1, 3):
                            if idx + offset < len(extracted_rows):
                                nxt_line = extracted_rows[idx + offset]
                                nxt_upper = nxt_line.upper().strip()
                                if nxt_upper in ['CR', '(CR)', 'CR.']:
                                    is_cr = True
                                    break
                                elif re.search(r'\d', nxt_upper):
                                    break 
                                    
                    clean_amt = re.sub(r'[^\d.]', '', amount_str)
                    val = float(clean_amt)
                    opening_balance = val
                    running_balance = -val if is_cr else val
                    found_opening = True
                    
                    d_val = "" if is_cr else val
                    c_val = val if is_cr else ""
                    
                    final_data.append({
                        "Date": "", "Type": "OPENING BAL", "Doc No": "", 
                        "Chq No": "", "Debit": d_val, "Credit": c_val, 
                        "Balance": round(running_balance, 2), "Remarks": ""
                    })
                continue

            date_match = re.search(r'\b(\d{2}/\d{2}/\d{2})\b', row_text)
            if date_match:
                skip_words = ["DATE:", "ACCOUNTING DATE", "TIME:", "PAGE"]
                if any(x in row_upper for x in skip_words): 
                    continue
                    
                date = date_match.group(1)
                amounts = re.findall(r'-?\(?[\d,]+\.\d{2}\)?', row_text)
                
                if not amounts: 
                    continue
                    
                amount_str = amounts[-1]
                has_cr_str = 'CR' in row_upper
                has_bracket = '(' in amount_str
                has_minus = '-' in amount_str
                is_cr = has_cr_str or has_bracket or has_minus
                
                clean_amt = re.sub(r'[^\d.]', '', amount_str)
                val = float(clean_amt)
                
                if val == 0.0: 
                    continue
                
                doc_no_match = re.search(r'\b\d{9,10}\b', row_text)
                doc_no = doc_no_match.group(0) if doc_no_match else ""
                
                t_type = "Other"
                s_type = "Other"
                remarks = ""
                
                is_tds_reco = "TDSRECO" in row_upper and doc_no.startswith('000') and is_cr
                is_tds_cr = "TDS CREDIT NOTE" in row_upper
                
                if is_tds_reco or is_tds_cr: 
                    t_type = "TDS Credit Note"
                    s_type = "TDS Credit Note"
                elif "CBOU199" in row_upper: 
                    t_type = "TECHNICAL BOUNCED"
                    s_type = "TECHNICAL BOUNCED"
                elif any(c in row_upper for c in ["CBOU101", "CBOU102", "CBOU110"]): 
                    t_type = "NON TECHNICAL BOUNCED"
                    s_
