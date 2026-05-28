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
