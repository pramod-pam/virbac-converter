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
            
            p_match = re.search(
                r'(?:Account.*?date|period from).*?(\d{2}/\d{2}/\d{2,4})\s*(?:to|-)\s*(\d{2}/\d{2}/\d{2,4})', 
                clean_text, re.IGNORECASE
            )
            if p_match: 
                period = f"from {p_match.group(1)} To {p_match.group(2)}"
            
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
                                ignore_keywords = [
                                    'date', 'time', 'page', 'statement', 
                                    'accounting', 'period', 'payer', 'customer', 
                                    'limit', 'opening', 'bal', 'dt', 'balance'
                                ]
                                if any(x in cand_lower for x in ignore_keywords):
                                    continue
                                customer_name = candidate
                                break
                        if customer_name: break
    except Exception as e:
        return None, None, None, None, None, f"PDF वाचताना त्रुटी (Reading Error): {e}"

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
        return None, None, None, None, None, f"PDF एक्स्ट्रॅक्ट करताना त्रुटी: {e}"

    final_data, running_balance, opening_balance = [], 0.0, 0.0
    found_opening = False
    total_billed_dr, total_paid_cr = 0.0, 0.0
    s_inv, pay, reco, c_oth, c_brk, g_ret = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    d_not, tcs, tds, tech_b, n_tech_b = 0.0, 0.0, 0.0, 0.0, 0.0

    try:
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
                                if next_line in ['CR', '(CR)', 'CR.']:
                                    is_cr = True
                                    break
                                elif re.search(r'\d', next_line):
                                    break 
                                    
                    val = float(re.sub(r'[^\d.]', '', amount_str))
                    opening_balance = val
                    running_balance = -val if is_cr else val
                    found_opening = True
                    final_data.append({
                        "Date": "", "Type": "OPENING BAL", "Doc No": "", 
                        "Chq No": "", "Debit": val if not is_cr else "", 
                        "Credit": val if is_cr else "", 
                        "Balance": round(running_balance, 2), "Remarks": ""
                    })
                continue

            date_match = re.search(r'\b(\d{2}/\d{2}/\d{2})\b', row_text)
            if date_match:
                if any(x in row_upper for x in ["DATE:", "ACCOUNTING DATE", "TIME:", "PAGE"]): 
                    continue
                date = date_match.group(1)
                amounts = re.findall(r'-?\(?[\d,]+\.\d{2}\)?', row_text)
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
                elif any(code in row_upper for code in ["CBOU101", "CBOU102", "CBOU110"]): 
                    t_type, s_type = "NON TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"
                elif "RECONC" in row_upper: 
                    t_type, s_type = "Reconciliation", "Reconciliation"
                elif any(x in row_upper for x in ["CHQ", "PAYMENT", "DD-NEFT", "NEFT"]) or (doc_no.startswith('000') and is_cr): 
                    t_type, s_type = "PAYMENT", "PAYMENT"
                elif "INVOICE" in row_upper:
                    if doc_no.startswith('3'): 
                        t_type, s_type = "Credit Note(Brakage Expiry)", "Credit Note(Brakage Expiry)"
                    elif doc_no.startswith('4'): 
                        t_type, s_type = "Credit Note(Others)", "Credit Note(Others)"
                    elif doc_no.startswith('5'): 
                        t_type, s_type = "Debit Note", "Debit Note"
                    elif doc_no.startswith('8'): 
                        t_type, s_type = "TCS Debit Note", "TCS Debit Note"
                    else: 
                        t_type, s_type = ("Goods Return Invoice", "Goods Return Invoice") if is_cr else ("Sales Invoice", "Sales Invoice")

                chq_no = ""
                if s_type in ["PAYMENT", "TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"] or any(x in row_upper for x in ["CHQ", "NEFT", "DD", "RTGS"]):
                    tokens = row_text.split()
                    for token in tokens:
                        if re.match(r'\d{2}/\d{2}/\d{2}', token) or token == date: continue
                        if re.search(r'\.\d{2}\)?$', token) or re.search(r'\d,\d', token): continue
                        clean_token = re.sub(r'[^A-Za-z0-9]', '', token)
                        if not clean_token or clean_token == doc_no: continue
                        if clean_token.isdigit() and len(clean_token) == 6:
                            chq_no = token; break
                        elif re.match(r'^[A-Za-z0-9]{8,25}$', clean_token):
                            ignore_words = [
                                "PAYMENT", "RECONC", "INVOICE", "OPENING", 
                                "BALANCE", "CLOSING", "TECHNICAL", "BOUNCED", "NON"
                            ]
                            if clean_token.upper() not in ignore_words and not clean_token.upper().startswith("CBOU"):
                                chq_no = token; break

                debit, credit = (val, 0.0) if not is_cr else (0.0, val)
                total_billed_dr += debit
                total_paid_cr += credit

                if s_type == "Sales Invoice": s_inv += val
                elif s_type == "PAYMENT": pay += val
                elif s_type == "Reconciliation": reco += (debit - credit)
                elif s_type == "Credit Note(Others)": c_oth += val
                elif s_type == "Credit Note(Brakage Expiry)": c_brk += val
                elif s_type == "Goods Return Invoice": g_ret += val
                elif s_type == "Debit Note": d_not += val
                elif s_type == "TCS Debit Note": tcs += val
                elif s_type == "TDS Credit Note": tds += val
                elif s_type == "TECHNICAL BOUNCED": tech_b += val
                elif s_type == "NON TECHNICAL BOUNCED": n_tech_b += val

                final_data.append({
                    "Date": date, "Type": t_type, "Doc No": doc_no, "Chq No": chq_no, 
                    "Debit": debit if debit > 0 else "", 
                    "Credit": credit if credit > 0 else "", 
                    "Balance": "", "Remarks": remarks
                })

        if final_data:
            reordered_data = []
            current_date = None
            date_block = []
            
            # ==== CRITICAL FIX: PERFECT SORTING ORDER FOR RECONCILIATIONS ====
            def row_sort_key(x):
                t = str(x.get('Type', ''))
                if 'OPENING BAL' in t: 
                    return (0, '', 0)
                elif 'PAYMENT' in t or 'BOUNCED' in t: 
                    return (2, str(x.get('Chq No', '')), 0)
                elif 'Advance Utilized' in t or ('Reconciliation' in t and str(x.get('Debit', '')) != ""):
                    return (3, '1', 0)
                elif 'Bill Settled' in t or ('Reconciliation' in t and str(x.get('Credit', '')) != "" and 'Advance Carried Forward' not in t):
