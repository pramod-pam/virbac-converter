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
                    return (3, '2', 0)
                elif 'Advance Carried Forward' in t:
                    return (3, '3', 0)
                else:
                    return (1, '', 0)

            for r in final_data:
                if r['Type'] == 'OPENING BAL':
                    reordered_data.append(r)
                    continue
                    
                if r['Date'] != current_date:
                    if date_block:
                        date_block.sort(key=row_sort_key)
                        reordered_data.extend(date_block)
                    current_date = r['Date']
                    date_block = [r]
                else:
                    date_block.append(r)
                    
            if date_block:
                date_block.sort(key=row_sort_key)
                reordered_data.extend(date_block)
                
            final_data = reordered_data
            
            current_bal = -opening_balance if final_data[0].get("Credit") else opening_balance
            for r in final_data:
                if r['Type'] == 'OPENING BAL':
                    r['Balance'] = round(current_bal, 2)
                    continue
                d = float(r['Debit']) if r['Debit'] != "" else 0.0
                c = float(r['Credit']) if r['Credit'] != "" else 0.0
                current_bal += (d - c)
                r['Balance'] = round(current_bal, 2)
                
            running_balance = current_bal

        if final_data:
            for i in range(len(final_data)):
                if "BOUNCED" in final_data[i]["Type"] and final_data[i]["Chq No"] == "":
                    b_amt = final_data[i]["Debit"] if final_data[i]["Debit"] != "" else final_data[i]["Credit"]
                    b_doc = final_data[i]["Doc No"]
                    for j in range(i - 1, -1, -1):
                        if final_data[j]["Type"] == "PAYMENT":
                            p_amt = final_data[j]["Credit"] if final_data[j]["Credit"] != "" else final_data[j]["Debit"]
                            p_doc = final_data[j]["Doc No"]
                            if (b_amt != "" and b_amt == p_amt) or (b_doc != "" and b_doc == p_doc):
                                if final_data[j]["Chq No"]:
                                    final_data[i]["Chq No"] = final_data[j]["Chq No"]  
                                    break
            
            doc_amounts = {}
            doc_details_for_pending = {}
            for r in final_data:
                t_check = r["Type"]
                if t_check not in ["PAYMENT", "Reconciliation", "OPENING BAL", "CLOSING BAL"] and "BOUNCED" not in t_check:
                    if r["Doc No"]:
                        amt = r["Debit"] if r["Debit"] != "" else r["Credit"]
                        if amt != "":
                            doc_amounts[r["Doc No"]] = float(amt)
                            doc_details_for_pending[r["Doc No"]] = {
                                "Date": r["Date"], "Type": r["Type"], "Amt": float(amt)
                            }

            inv_balances = {k: v for k, v in doc_amounts.items()}
            adv_balances = {}
            for r in final_data:
                if r["Doc No"] and r["Doc No"].startswith('000'):
                    c_val = float(r["Credit"]) if r["Credit"] != "" else 0.0
                    if c_val > 0:
                        adv_balances[r["Doc No"]] = c_val

            adj_credits_map = {}
            reconc_dr_map = {}
            reconc_return_debits_by_date = {}  
            
            for temp_r in final_data:
                if temp_r["Type"] == "Reconciliation":
                    if temp_r["Credit"] != "":
                        c_val = float(temp_r["Credit"])
                        temp_doc = temp_r.get("Doc No", "")
                        if temp_doc and not temp_doc.startswith('000'):
                            adj_credits_map[c_val] = temp_doc
                    if temp_r["Debit"] != "":
                        d_val = float(temp_r["Debit"])
                        temp_doc = temp_r.get("Doc No", "")
                        reconc_dr_map[(temp_r["Date"], d_val)] = temp_doc
                        
                        t_cat = doc_details_for_pending.get(temp_doc, {}).get("Type", "")
                        is_ret = temp_doc.startswith(('22', '3', '4', '9'))
                        is_ret_cat = "Return" in t_cat or "Credit Note" in t_cat
                        if temp_doc and (is_ret or is_ret_cat):
                            dt = temp_r["Date"]
                            if dt not in reconc_return_debits_by_date:
                                reconc_return_debits_by_date[dt] = []
                            reconc_return_debits_by_date[dt].append(d_val)

            chq_stats = {}
            bounced_stats = {}
            
            for r in final_data:
                if r.get("Chq No"):
                    key = (r["Date"], r["Chq No"])
                    c_amt = float(r["Credit"]) if r["Credit"] != "" else 0.0
                    d_amt = float(r["Debit"]) if r["Debit"] != "" else 0.0
                    
                    if "PAYMENT" in r["Type"]:
                        if key not in chq_stats:
                            chq_stats[key] = {'total_c': 0.0, 'total_d': 0.0, 'count': 0, 'seen': 0}
                        chq_stats[key]['total_c'] += c_amt
                        chq_stats[key]['total_d'] += d_amt
                        chq_stats[key]['count'] += 1
                    
                    elif "BOUNCED" in r["Type"]:
                        if key not in bounced_stats:
                            bounced_stats[key] = {'total_c': 0.0, 'total_d': 0.0, 'count': 0, 'seen': 0}
                        bounced_stats[key]['total_c'] += c_amt
                        bounced_stats[key]['total_d'] += d_amt
                        bounced_stats[key]['count'] += 1

            grouped_data = []
            for r in final_data:
                doc_no = r.get("Doc No", "")
                if "PAYMENT" in r["Type"]:
                    amt = float(r["Credit"]) if r["Credit"] != "" else (float(r["Debit"]) if r["Debit"] != "" else 0.0)
                    
                    status_tag = ""
                    if doc_no and doc_no in inv_balances:
                        inv_balances[doc_no] -= amt
                        pending = inv_balances[doc_no]
                        status_tag = " [CLEARED]" if pending <= 0.5 else f" [Pend: {int(pending):,}]"
                    
                    orig_amt = doc_amounts.get(doc_no, None)
                    orig_amt_str = f" | Inv Amt: {int(orig_amt):,}" if orig_amt else ""
                    adj_formatted = f"{int(amt):,}"
                    
                    doc_type = "Inv"
                    if doc_no.startswith(('3','4','9')): doc_type = "Cr Note"
                    elif doc_no.startswith('5'): doc_type = "Dr Note"
                    elif doc_no.startswith('8'): doc_type = "TCS Dr Note"
                    elif doc_no.startswith('000'): doc_type = "Adv"
                    
                    if doc_no and doc_no.startswith('000'):
                        r["Remarks"] = f"Total Advance Received: {int(amt):,}"
                    elif doc_no:
                        r["Remarks"] = f"{doc_type}: {doc_no}{orig_amt_str} | Adj: {adj_formatted}{status_tag}"
                    else:
                        r["Remarks"] = f"Adj: {adj_formatted}"
                    
                    r["Type"] = "PAYMENT"
                    grouped_data.append(r)
                    
                    if r["Chq No"]:
                        key = (r["Date"], r["Chq No"])
                        if key in chq_stats and chq_stats[key]['count'] >= 1:
                            chq_stats[key]['seen'] += 1
                            if chq_stats[key]['seen'] == chq_stats[key]['count']:
                                total_c = chq_stats[key]['total_c']
                                total_d = chq_stats[key]['total_d']
                                net_amt = abs(total_c - total_d)
                                total_adj = total_c if total_c > total_d else total_d
                                
                                grouped_data.append({
                                    "Date": "", "Type": "-> CHQ/NEFT SUMMARY", "Doc No": "", "Chq No": key[1],  
                                    "Debit": "", "Credit": "", "Balance": "", 
                                    "Remarks": f"Total Inv Adj: {int(total_adj):,} | Total Chq/NEFT Amt: {int(net_amt):,}"
                                })

                elif "BOUNCED" in r["Type"]:
                    grouped_data.append(r)
                    
                    if r["Chq No"]:
                        key = (r["Date"], r["Chq No"])
                        if key in bounced_stats and bounced_stats[key]['count'] >= 1:
                            bounced_stats[key]['seen'] += 1
                            if bounced_stats[key]['seen'] == bounced_stats[key]['count']:
                                total_c = bounced_stats[key]['total_c']
                                total_d = bounced_stats[key]['total_d']
                                net_amt = abs(total_c - total_d)
                                
                                grouped_data.append({
                                    "Date": "", "Type": "-> BOUNCED CHQ/NEFT SUMMARY", "Doc No": "", "Chq No": key[1],  
                                    "Debit": "", "Credit": "", "Balance": "", 
                                    "Remarks": f"Total Bounced Chq/NEFT Amt: {int(net_amt):,}"
                                })

                elif "Reconciliation" in r["Type"]:
                    amt = float(r["Credit"]) if r["Credit"] != "" else (float(r["Debit"]) if r["Debit"] != "" else 0.0)
                    is_debit_entry = r["Debit"] != ""
                    doc_category = doc_details_for_pending.get(doc_no, {}).get("Type", "")
                    
                    if is_debit_entry:
                        if doc_no in inv_balances:
                            inv_balances[doc_no] += amt 
                        
                        is_ret = doc_no.startswith(('22', '3', '4', '9'))
                        is_ret_cat = "Return" in doc_category or "Credit Note" in doc_category
                        if doc_no and (is_ret or is_ret_cat):
                            r["Type"] = ">> System Adj (Reconciliation)"
                            r["Remarks"] = "Adjusted from Return | System entry (Ignore)"
                        elif doc_no and doc_no.startswith('000'):
                            r["Type"] = ">> Advance Utilized"
                            if doc_no in adv_balances:
                                adv_balances[doc_no] -= amt 
                                rem_adv_bal = adv_balances[doc_no]
                                r["Remarks"] = f"Deducted from Adv: {doc_no} | Rem Adv Bal: {int(rem_adv_bal):,}"
                            else:
                                if abs(amt - abs(opening_balance)) < 0.5:
                                    r["Remarks"] = f"Deducted from Opening Balance | Adj: {int(amt):,}"
                                else:
                                    r["Remarks"] = f"Deducted from Old Adv (Opening Bal) | Adj: {int(amt):,}"
                        else:
                            r["Type"] = "Reconciliation"
                            r["Remarks"] = f"Debit Adj: {doc_no} | Amt: {int(amt):,}"
                    else:
                        source_doc = reconc_dr_map.get((r["Date"], amt),
