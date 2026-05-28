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
                    s_type = "NON TECHNICAL BOUNCED"
                elif "RECONC" in row_upper: 
                    t_type = "Reconciliation"
                    s_type = "Reconciliation"
                elif any(x in row_upper for x in ["CHQ", "PAYMENT", "DD-NEFT", "NEFT"]):
                    t_type = "PAYMENT"
                    s_type = "PAYMENT"
                elif doc_no.startswith('000') and is_cr:
                    t_type = "PAYMENT"
                    s_type = "PAYMENT"
                elif "INVOICE" in row_upper:
                    if doc_no.startswith('3'): 
                        t_type = "Credit Note(Brakage Expiry)"
                        s_type = "Credit Note(Brakage Expiry)"
                    elif doc_no.startswith('4'): 
                        t_type = "Credit Note(Others)"
                        s_type = "Credit Note(Others)"
                    elif doc_no.startswith('5'): 
                        t_type = "Debit Note"
                        s_type = "Debit Note"
                    elif doc_no.startswith('8'): 
                        t_type = "TCS Debit Note"
                        s_type = "TCS Debit Note"
                    elif is_cr:
                        t_type = "Goods Return Invoice"
                        s_type = "Goods Return Invoice"
                    else: 
                        t_type = "Sales Invoice"
                        s_type = "Sales Invoice"

                chq_no = ""
                pay_bounced_types = [
                    "PAYMENT", "TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"
                ]
                has_chq_word = any(x in row_upper for x in ["CHQ", "NEFT", "DD", "RTGS"])
                
                if s_type in pay_bounced_types or has_chq_word:
                    tokens = row_text.split()
                    for token in tokens:
                        if re.match(r'\d{2}/\d{2}/\d{2}', token) or token == date: 
                            continue
                        if re.search(r'\.\d{2}\)?$', token) or re.search(r'\d,\d', token): 
                            continue
                            
                        clean_token = re.sub(r'[^A-Za-z0-9]', '', token)
                        if not clean_token or clean_token == doc_no: 
                            continue
                            
                        if clean_token.isdigit() and len(clean_token) == 6:
                            chq_no = token
                            break
                        elif re.match(r'^[A-Za-z0-9]{8,25}$', clean_token):
                            ignore_words = [
                                "PAYMENT", "RECONC", "INVOICE", "OPENING", 
                                "BALANCE", "CLOSING", "TECHNICAL", "BOUNCED", "NON"
                            ]
                            c_tok_up = clean_token.upper()
                            if c_tok_up not in ignore_words and not c_tok_up.startswith("CBOU"):
                                chq_no = token
                                break

                debit = 0.0 if is_cr else val
                credit = val if is_cr else 0.0
                
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

                d_val = debit if debit > 0 else ""
                c_val = credit if credit > 0 else ""
                
                final_data.append({
                    "Date": date, "Type": t_type, "Doc No": doc_no, 
                    "Chq No": chq_no, "Debit": d_val, "Credit": c_val, 
                    "Balance": "", "Remarks": remarks
                })

        if final_data:
            reordered_data = []
            current_date = None
            date_block = []
            
            def row_sort_key(x):
                t = str(x.get('Type', ''))
                if 'OPENING BAL' in t: 
                    return (0, '', 0)
                elif 'PAYMENT' in t or 'BOUNCED' in t: 
                    return (2, str(x.get('Chq No', '')), 0)
                
                is_adv_util = 'Advance Utilized' in t
                is_recon_dr = 'Reconciliation' in t and str(x.get('Debit', '')) != ""
                if is_adv_util or is_recon_dr:
                    return (3, '1', 0)
                    
                is_bill_set = 'Bill Settled' in t
                is_recon_cr = 'Reconciliation' in t and str(x.get('Credit', '')) != ""
                is_not_cf = 'Advance Carried Forward' not in t
                if is_bill_set or (is_recon_cr and is_not_cf):
                    return (3, '2', 0)
                    
                if 'Advance Carried Forward' in t:
                    return (3, '3', 0)
                    
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
            
            current_bal = opening_balance
            if final_data[0].get("Credit"):
                current_bal = -opening_balance
                
            for r in final_data:
                if r['Type'] == 'OPENING BAL':
                    r['Balance'] = round(current_bal, 2)
                    continue
                d_str = r['Debit']
                c_str = r['Credit']
                d = float(d_str) if d_str != "" else 0.0
                c = float(c_str) if c_str != "" else 0.0
                current_bal += (d - c)
                r['Balance'] = round(current_bal, 2)
                
            running_balance = current_bal

        if final_data:
            for i in range(len(final_data)):
                typ = final_data[i]["Type"]
                chq = final_data[i]["Chq No"]
                if "BOUNCED" in typ and chq == "":
                    d_str = final_data[i]["Debit"]
                    c_str = final_data[i]["Credit"]
                    b_amt = d_str if d_str != "" else c_str
                    b_doc = final_data[i]["Doc No"]
                    
                    for j in range(i - 1, -1, -1):
                        if final_data[j]["Type"] == "PAYMENT":
                            pd_str = final_data[j]["Debit"]
                            pc_str = final_data[j]["Credit"]
                            p_amt = pc_str if pc_str != "" else pd_str
                            p_doc = final_data[j]["Doc No"]
                            
                            amt_match = (b_amt != "" and b_amt == p_amt)
                            doc_match = (b_doc != "" and b_doc == p_doc)
                            
                            if amt_match or doc_match:
                                if final_data[j]["Chq No"]:
                                    final_data[i]["Chq No"] = final_data[j]["Chq No"]  
                                    break
            
            doc_amounts = {}
            doc_details_for_pending = {}
            
            for r in final_data:
                t_check = r["Type"]
                skip_types = ["PAYMENT", "Reconciliation", "OPENING BAL", "CLOSING BAL"]
                
                if t_check not in skip_types and "BOUNCED" not in t_check:
                    if r["Doc No"]:
                        d_str = r["Debit"]
                        c_str = r["Credit"]
                        amt = d_str if d_str != "" else c_str
                        if amt != "":
                            doc_amounts[r["Doc No"]] = float(amt)
                            doc_details_for_pending[r["Doc No"]] = {
                                "Date": r["Date"], 
                                "Type": r["Type"], 
                                "Amt": float(amt)
                            }

            inv_balances = {k: v for k, v in doc_amounts.items()}
            adv_balances = {}
            
            for r in final_data:
                d_no = r["Doc No"]
                if d_no and d_no.startswith('000'):
                    c_str = r["Credit"]
                    c_val = float(c_str) if c_str != "" else 0.0
                    if c_val > 0:
                        adv_balances[d_no] = c_val

            adj_credits_map = {}
            reconc_dr_map = {}
            reconc_return_debits_by_date = {}  
            
            for temp_r in final_data:
                if temp_r["Type"] == "Reconciliation":
                    c_str = temp_r["Credit"]
                    d_str = temp_r["Debit"]
                    temp_doc = temp_r.get("Doc No", "")
                    
                    if c_str != "":
                        c_val = float(c_str)
                        if temp_doc and not temp_doc.startswith('000'):
                            adj_credits_map[c_val] = temp_doc
                            
                    if d_str != "":
                        d_val = float(d_str)
                        dt_val_tuple = (temp_r["Date"], d_val)
                        reconc_dr_map[dt_val_tuple] = temp_doc
                        
                        det_dict = doc_details_for_pending.get(temp_doc, {})
                        t_cat = det_dict.get("Type", "")
                        
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
                c_no = r.get("Chq No")
                if c_no:
                    key = (r["Date"], c_no)
                    c_str = r["Credit"]
                    d_str = r["Debit"]
                    c_amt = float(c_str) if c_str != "" else 0.0
                    d_amt = float(d_str) if d_str != "" else 0.0
                    
                    if "PAYMENT" in r["Type"]:
                        if key not in chq_stats:
                            chq_stats[key] = {'tc': 0.0, 'td': 0.0, 'cnt': 0, 'seen': 0}
                        chq_stats[key]['tc'] += c_amt
                        chq_stats[key]['td'] += d_amt
                        chq_stats[key]['cnt'] += 1
                    
                    elif "BOUNCED" in r["Type"]:
                        if key not in bounced_stats:
                            bounced_stats[key] = {'tc': 0.0, 'td': 0.0, 'cnt': 0, 'seen': 0}
                        bounced_stats[key]['tc'] += c_amt
                        bounced_stats[key]['td'] += d_amt
                        bounced_stats[key]['cnt'] += 1

            grouped_data = []
            for r in final_data:
                doc_no = r.get("Doc No", "")
                
                if "PAYMENT" in r["Type"]:
                    c_str = r["Credit"]
                    d_str = r["Debit"]
                    amt = float(c_str) if c_str != "" else (float(d_str) if d_str != "" else 0.0)
                    
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
                        rem_str = f"{doc_type}: {doc_no}{orig_amt_str} | Adj: {adj_formatted}{status_tag}"
                        r["Remarks"] = rem_str
                    else:
                        r["Remarks"] = f"Adj: {adj_formatted}"
                    
                    r["Type"] = "PAYMENT"
                    grouped_data.append(r)
                    
                    c_no = r.get("Chq No")
                    if c_no:
                        key = (r["Date"], c_no)
                        if key in chq_stats and chq_stats[key]['cnt'] >= 1:
                            chq_stats[key]['seen'] += 1
                            if chq_stats[key]['seen'] == chq_stats[key]['cnt']:
                                t_c = chq_stats[key]['tc']
                                t_d = chq_stats[key]['td']
                                net_amt = abs(t_c - t_d)
                                t_adj = t_c if t_c > t_d else t_d
                                
                                grouped_data.append({
                                    "Date": "", "Type": "-> CHQ/NEFT SUMMARY", 
                                    "Doc No": "", "Chq No": key[1],  
                                    "Debit": "", "Credit": "", "Balance": "", 
                                    "Remarks": f"Total Inv Adj: {int(t_adj):,} | Total Chq/NEFT Amt: {int(net_amt):,}"
                                })

                elif "BOUNCED" in r["Type"]:
                    grouped_data.append(r)
                    c_no = r.get("Chq No")
                    
                    if c_no:
                        key = (r["Date"], c_no)
                        if key in bounced_stats and bounced_stats[key]['cnt'] >= 1:
                            bounced_stats[key]['seen'] += 1
                            if bounced_stats[key]['seen'] == bounced_stats[key]['cnt']:
                                t_c = bounced_stats[key]['tc']
                                t_d = bounced_stats[key]['td']
                                net_amt = abs(t_c - t_d)
                                
                                grouped_data.append({
                                    "Date": "", "Type": "-> BOUNCED CHQ/NEFT SUMMARY", 
                                    "Doc No": "", "Chq No": key[1],  
                                    "Debit": "", "Credit": "", "Balance": "", 
                                    "Remarks": f"Total Bounced Chq/NEFT Amt: {int(net_amt):,}"
                                })

                elif "Reconciliation" in r["Type"]:
                    c_str = r["Credit"]
                    d_str = r["Debit"]
                    amt = float(c_str) if c_str != "" else (float(d_str) if d_str != "" else 0.0)
                    
                    is_debit_entry = d_str != ""
                    det_dict = doc_details_for_pending.get(doc_no, {})
                    doc_category = det_dict.get("Type", "")
                    
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
                                rem_str = f"Deducted from Adv: {doc_no} | Rem Adv Bal: {int(rem_adv_bal):,}"
                                r["Remarks"] = rem_str
                            else:
                                if abs(amt - abs(opening_balance)) < 0.5:
                                    r["Remarks"] = f"Deducted from Opening Balance | Adj: {int(amt):,}"
                                else:
                                    r["Remarks"] = f"Deducted from Old Adv (Opening Bal) | Adj: {int(amt):,}"
                        else:
                            r["Type"] = "Reconciliation"
                            r["Remarks"] = f"Debit Adj: {doc_no} | Amt: {int(amt):,}"
                    else:
                        dt_amt_tuple = (r["Date"], amt)
                        source_doc = reconc_dr_map.get(dt_amt_tuple, "")
                        
                        is_ret_src = source_doc.startswith(('22', '3', '4', '9'))
                        src_det = doc_details_for_pending.get(source_doc, {})
                        is_ret_src_cat = "Return" in src_det.get("Type", "")
                        is_from_return = source_doc and (is_ret_src or is_ret_src_cat)
                        
                        if not is_from_return:
                            ret_debits = reconc_return_debits_by_date.get(r["Date"], [])
                            if ret_debits:
                                sum_ret = sum(ret_debits)
                                diff_amt = abs(sum_ret - amt)
                                
                                if diff_amt < 0.5:
                                    is_from_return = True
                                elif len(ret_debits) <= 15: 
                                    for combo_len in range(2, len(ret_debits) + 1):
                                        found_combo = False
                                        for combo in itertools.combinations(ret_debits, combo_len):
                                            combo_sum = sum(combo)
                                            combo_diff = abs(combo_sum - amt)
                                            if combo_diff < 0.5:
                                                is_from_return = True
                                                found_combo = True
                                                break
                                        if found_combo:
                                            break
                        
                        if doc_no in inv_balances:
                            inv_balances[doc_no] -= amt
                            pending = inv_balances[doc_no]
                            status_tag = " [CLEARED]" if pending <= 0.5 else f" [Pend: {int(pending):,}]"
                        else:
                            status_tag = ""

                        if doc_no and doc_no.startswith('000'):
                            r["Type"] = ">> Advance Carried Forward"
                            r["Remarks"] = f"Advance Carried Forward: {doc_no} | Amt: {int(amt):,}"
                        elif is_from_return:
                            r["Type"] = ">> Bill Settled (Reconciliation)"
                            orig_amt = doc_amounts.get(doc_no, None)
                            orig_amt_str = f" | Amt: {int(orig_amt):,}" if orig_amt else ""
                            r["Remarks"] = f"Bill Settled Against Return | Inv: {doc_no}{orig_amt_str} | Adj: {int(amt):,}{status_tag}"
                        else:
                            r["Type"] = ">> Bill Settled (Reconciliation)"
                            orig_amt = doc_amounts.get(doc_no, None)
                            orig_amt_str = f" | Amt: {int(orig_amt):,}" if orig_amt else ""
                            doc_type = "Inv"
                            if doc_no.startswith(('3','4','9')): doc_type = "Cr Note"
                            elif doc_no.startswith('5'): doc_type = "Dr Note"
                            elif doc_no.startswith('8'): doc_type = "TCS Dr Note"
                            r["Remarks"] = f"{doc_type}: {doc_no}{orig_amt_str} | Adj: {int(amt):,}{status_tag}"
                            
                grouped_data.append(r)
            else:
                if "Goods Return Invoice" in r["Type"]:
                    r["Remarks"] = "Material Returned by Customer"
                elif "Credit Note(Brakage Expiry)" in r["Type"]:
                    r["Remarks"] = "Credit Note Issued for Breakage/Expiry"
                elif "Credit Note(Others)" in r["Type"]:
                    r["Remarks"] = "Credit Note Issued (Others)"
                grouped_data.append(r)

            final_data = grouped_data
            
            re_sorted_final = []
            current_date_block = []
            curr_date = None
            
            for r in final_data:
                chk_t = str(r.get("Type", ""))
                skip_sort = ["OPENING BAL", "CLOSING BAL", "-> BATCH SETTLEMENT SUMMARY"]
                
                if chk_t in skip_sort:
                    if current_date_block:
                        current_date_block.sort(key=row_sort_key)
                        re_sorted_final.extend(current_date_block)
                        current_date_block = []
                    re_sorted_final.append(r)
