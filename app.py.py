import streamlit as st
import pdfplumber
import pandas as pd
import re
import datetime
import io
import os
import textwrap
import itertools

# Web app design
st.set_page_config(page_title="Virbac Statement Converter", page_icon="📄", layout="wide")

st.title("📄 Virbac Account Statement Converter (Smart Tracking & Dashboards)")
st.markdown("CFA Team sathi: PDF upload kara ani **Excel + PDF** donhi format milva.")

uploaded_files = st.file_uploader("Yethe PDF file upload kara", type="pdf", accept_multiple_files=True)

def process_pdf_logic(uploaded_file):
    run_datetime_obj = datetime.datetime.now()
    run_datetime = run_datetime_obj.strftime("%d/%m/%Y %I:%M %p")
    today_date = run_datetime_obj.date()
    
    period, customer_no, customer_name = "", "", ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            first_page = pdf.pages[0]
            text_layout = first_page.extract_text(layout=True) or ""
            clean_text = " ".join(text_layout.split())
            
            p_match = re.search(r'(?:Account.*?date|period from).*?(\d{2}/\d{2}/\d{2,4})\s*(?:to|-)\s*(\d{2}/\d{2}/\d{2,4})', clean_text, re.IGNORECASE)
            if p_match: period = f"from {p_match.group(1)} To {p_match.group(2)}"
            
            c_match = re.search(r'(?:Payer|Customer No).*?(\d{6})', clean_text, re.IGNORECASE)
            if c_match: 
                customer_no = c_match.group(1).strip()
                for p in pdf.pages[:1]:
                    lines = [l.strip() for l in (p.extract_text() or "").split('\n') if l.strip()]
                    for i, line in enumerate(lines):
                        if customer_no in line:
                            for j in range(i + 1, min(i + 7, len(lines))):
                                candidate = lines[j].strip()
                                cand_lower = candidate.lower()
                                if len(candidate) < 4: continue
                                if any(x in cand_lower for x in ['date', 'time', 'page', 'statement', 'accounting', 'period', 'payer', 'customer', 'limit', 'opening', 'bal', 'dt', 'balance']):
                                    continue
                                customer_name = candidate
                                break
                        if customer_name: break
    except Exception as e:
        return None, None, None, None, None, f"PDF vachtana error: {e}"

    extracted_rows = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            p_text = page.extract_text()
            if p_text:
                for line in p_text.split('\n'): extracted_rows.append(line.strip())

    final_data, running_balance, opening_balance, found_opening = [], 0.0, 0.0, False
    total_billed_dr, total_paid_cr = 0.0, 0.0
    s_inv, pay, reco, c_oth, c_brk, g_ret, d_not, tcs, tds, tech_b, n_tech_b = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

    for idx, row_text in enumerate(extracted_rows):
        row_upper = row_text.upper()
        
        if "OPENING BALANCE" in row_upper and not found_opening:
            amounts = re.findall(r'-?\s*\(?\s*[\d,]+\.\d{2}\s*\)?', row_text)
            if amounts:
                amount_str = amounts[-1]
                is_cr = 'CR' in row_upper or '(' in amount_str or '-' in amount_str
                
                if not is_cr:
                    for offset in range(1, 3):
                        if idx + offset < len(extracted_rows):
                            next_line = extracted_rows[idx + offset].upper().strip()
                            if next_line == 'CR' or next_line == '(CR)' or next_line == 'CR.':
                                is_cr = True
                                break
                            elif re.search(r'\d', next_line):
                                break 
                                
                val = float(re.sub(r'[^\d.]', '', amount_str))
                opening_balance = val
                running_balance = -val if is_cr else val
                found_opening = True
                final_data.append({"Date": "", "Type": "OPENING BAL", "Doc No": "", "Chq No": "", "Debit": val if not is_cr else "", "Credit": val if is_cr else "", "Balance": round(running_balance, 2), "Remarks": ""})
            continue

        date_match = re.search(r'\b(\d{2}/\d{2}/\d{2})\b', row_text)
        if date_match:
            if any(x in row_upper for x in ["DATE:", "ACCOUNTING DATE", "TIME:", "PAGE"]): continue
            date, amounts = date_match.group(1), re.findall(r'-?\(?[\d,]+\.\d{2}\)?', row_text)
            if not amounts: continue
            amount_str = amounts[-1]
            is_cr = 'CR' in row_upper or '(' in amount_str or '-' in amount_str
            val = float(re.sub(r'[^\d.]', '', amount_str))
            if val == 0.0: continue
            
            doc_no_match = re.search(r'\b\d{9,10}\b', row_text)
            doc_no = doc_no_match.group(0) if doc_no_match else ""
            
            t_type, s_type = "Other", "Other"
            remarks = ""
            
            if ("TDSRECO" in row_upper and doc_no.startswith('000') and is_cr) or ("TDS CREDIT NOTE" in row_upper): 
                t_type, s_type = "TDS Credit Note", "TDS Credit Note"
            elif "CBOU199" in row_upper: 
                t_type, s_type = "TECHNICAL BOUNCED", "TECHNICAL BOUNCED"
            elif any(code in row_upper for code in ["CBOU101", "
