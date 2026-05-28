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
    run_datetime_obj = datetime.datetime.now()
    run_datetime = run_datetime_obj.strftime("%d/%m/%Y %I:%M %p")
    today_date = run_datetime_obj.date()
    
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
        return None, None, None, None, None, f"PDF vachtana error: {e}"

    extracted_rows = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            p_text = page.extract_text()
            if p_text:
                for line in p_text.split('\n'): 
                    extracted_rows.append(line.strip())

    final_data, running_balance, opening_balance = [], 0.0, 0.0
    found_opening = False
    total_billed_dr, total_paid_cr = 0.0, 0.0
    s_inv, pay, reco, c_oth, c_brk, g_ret = 0.0, 0.
