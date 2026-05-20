import streamlit as st
import pdfplumber
import pandas as pd
import re
import datetime
import io
import os
import textwrap

# Web app design
st.set_page_config(page_title="Virbac Statement Converter", page_icon="📄", layout="wide")

st.title("📄 Virbac Account Statement Converter (Dual Dashboard)")
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

    for row_text in extracted_rows:
        row_upper = row_text.upper()
        if "OPENING BALANCE" in row_upper and not found_opening:
            amounts = re.findall(r'-?\(?[\d,]+\.\d{2}\)?', row_text)
            if amounts:
                amount_str = amounts[-1]
                is_cr = 'CR' in row_upper or '(' in amount_str or '-' in amount_str
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
            elif any(code in row_upper for code in ["CBOU101", "CBOU102", "CBOU110"]): 
                t_type, s_type = "NON TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"
            elif "RECONC" in row_upper: 
                t_type, s_type = "RECONCILIATION", "RECONCILIATION"
            elif any(x in row_upper for x in ["CHQ", "PAYMENT", "DD-NEFT", "NEFT"]) or (doc_no.startswith('000') and is_cr): 
                t_type, s_type = "PAYMENT", "PAYMENT"
            elif "INVOICE" in row_upper:
                if doc_no.startswith('3'): t_type, s_type = "Credit Note(Brakage Expiry)", "Credit Note(Brakage Expiry)"
                elif doc_no.startswith('4'): t_type, s_type = "Credit Note(Others)", "Credit Note(Others)"
                elif doc_no.startswith('5'): t_type, s_type = "Debit Note", "Debit Note"
                elif doc_no.startswith('8'): t_type, s_type = "TCS Debit Note", "TCS Debit Note"
                else: t_type, s_type = ("Goods Return Invoice", "Goods Return Invoice") if is_cr else ("Sales Invoice", "Sales Invoice")

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
                        ignore_words = ["PAYMENT", "RECONC", "INVOICE", "OPENING", "BALANCE", "CLOSING", "TECHNICAL", "BOUNCED", "NON"]
                        if clean_token.upper() not in ignore_words and not clean_token.upper().startswith("CBOU"):
                            chq_no = token; break

            debit, credit = (val, 0.0) if not is_cr else (0.0, val)
            running_balance += (debit - credit)
            total_billed_dr += debit
            total_paid_cr += credit

            if s_type == "Sales Invoice": s_inv += val
            elif s_type == "PAYMENT": pay += val
            elif s_type == "RECONCILIATION": reco += (debit - credit)
            elif s_type == "Credit Note(Others)": c_oth += val
            elif s_type == "Credit Note(Brakage Expiry)": c_brk += val
            elif s_type == "Goods Return Invoice": g_ret += val
            elif s_type == "Debit Note": d_not += val
            elif s_type == "TCS Debit Note": tcs += val
            elif s_type == "TDS Credit Note": tds += val
            elif s_type == "TECHNICAL BOUNCED": tech_b += val
            elif s_type == "NON TECHNICAL BOUNCED": n_tech_b += val

            final_data.append({"Date": date, "Type": t_type, "Doc No": doc_no, "Chq No": chq_no, "Debit": debit if debit > 0 else "", "Credit": credit if credit > 0 else "", "Balance": round(running_balance, 2), "Remarks": remarks})

    if final_data:
        # Bounced Cheque Match Logic
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
            if r["Type"] not in ["PAYMENT", "RECONCILIATION", "OPENING BAL", "CLOSING BAL"] and "BOUNCED" not in r["Type"]:
                if r["Doc No"]:
                    amt = r["Debit"] if r["Debit"] != "" else r["Credit"]
                    if amt != "":
                        doc_amounts[r["Doc No"]] = float(amt)
                        doc_details_for_pending[r["Doc No"]] = {"Date": r["Date"], "Type": r["Type"], "Amt": float(amt)}

        inv_balances = {k: v for k, v in doc_amounts.items()}
        chq_stats = {}
        for r in final_data:
            if "PAYMENT" in r["Type"] and r["Chq No"]:
                key = r["Chq No"]
                c_amt = float(r["Credit"]) if r["Credit"] != "" else 0.0
                d_amt = float(r["Debit"]) if r["Debit"] != "" else 0.0
                if key not in chq_stats:
                    chq_stats[key] = {'total_c': 0.0, 'total_d': 0.0, 'count': 0, 'seen': 0, 'last_bal': 0.0, 'date': r['Date']}
                chq_stats[key]['total_c'] += c_amt
                chq_stats[key]['total_d'] += d_amt
                chq_stats[key]['count'] += 1
                chq_stats[key]['last_bal'] = r['Balance'] 

        grouped_data = []
        skip_indices = set()
        
        for i, r in enumerate(final_data):
            if i in skip_indices: continue
            if "PAYMENT" in r["Type"] and r["Chq No"]:
                key = r["Chq No"]
                if chq_stats[key]['seen'] == 0:
                    total_c = chq_stats[key]['total_c']
                    total_d = chq_stats[key]['total_d']
                    net_amt = abs(total_c - total_d)
                    grouped_data.append({
                        "Date": chq_stats[key]['date'], "Type": "PAYMENT (Total)", "Doc No": "", "Chq No": key,  
                        "Debit": total_d if total_d > 0 else "", "Credit": total_c if total_c > 0 else "", 
                        "Balance": chq_stats[key]['last_bal'], "Remarks": f"Net NEFT/Chq Amt: {int(net_amt):,}"
                    })
                for j in range(i, len(final_data)):
                    if final_data[j]["Chq No"] == key and "PAYMENT" in final_data[j]["Type"]:
                        child_r = final_data[j]
                        doc_no = child_r.get("Doc No", "")
                        amt = float(child_r["Credit"]) if child_r["Credit"] != "" else (float(child_r["Debit"]) if child_r["Debit"] != "" else 0.0)
                        status_tag = ""
                        if doc_no and doc_no in inv_balances:
                            inv_balances[doc_no] -= amt
                            pending = inv_balances[doc_no]
                            status_tag = "[CLEARED]" if pending <= 0.5 else f"[Pend: {int(pending):,}]"
                        orig_amt = doc_amounts.get(doc_no, None)
                        orig_amt_str = f"Inv Amt: {int(orig_amt):,}" if orig_amt else ""
                        doc_type = "Cr Note" if doc_no.startswith(('3','4')) else "Dr Note" if doc_no.startswith('5') else "TCS Dr Note" if doc_no.startswith('8') else "Adv" if doc_no.startswith('000') else "Inv"
                        grouped_data.append({
                            "Date": "", "Type": f"    -> Adj {doc_type}" if doc_no else "    -> On Acct", 
                            "Doc No": doc_no, "Chq No": "", "Debit": child_r["Debit"], "Credit": child_r["Credit"], 
                            "Balance": "", "Remarks": " | ".join([p for p in [orig_amt_str, status_tag] if p])
                        })
                        skip_indices.add(j)
                        chq_stats[key]['seen'] += 1
            elif "RECONCILIATION" in r["Type"]:
                doc_no = r.get("Doc No", "")
                if doc_no:
                    amt = float(r["Credit"]) if r["Credit"] != "" else (float(r["Debit"]) if r["Debit"] != "" else 0.0)
                    status_tag = ""
                    if doc_no in inv_balances:
                        inv_balances[doc_no] -= amt
                        pending = inv_balances[doc_no]
                        status_tag = "[CLEARED]" if pending <= 0.5 else f"[Pend: {int(pending):,}]"
                    orig_amt = doc_amounts.get(doc_no, None)
                    orig_amt_str = f"Amt: {int(orig_amt):,}" if orig_amt else ""
                    doc_type = "Cr Note" if doc_no.startswith(('3','4')) else "Dr Note" if doc_no.startswith('5') else "TCS Dr Note" if doc_no.startswith('8') else "Adv" if doc_no.startswith('000') else "Inv"
                    r["Remarks"] = " | ".join([p for p in [f"Adj {doc_type}", orig_amt_str, status_tag] if p])
                grouped_data.append(r)
            else: grouped_data.append(r)

        final_data = grouped_data
        final_data.append({"Date": "", "Type": "CLOSING BAL", "Doc No": "", "Chq No": "", "Debit": "", "Credit": "", "Balance": round(running_balance, 2), "Remarks": ""})
        
        # --- PENDING BILLS WITH AGING LOGIC ---
        pending_invoices = []
        total_pending_amt = 0.0
        for doc, pending_amt in inv_balances.items():
            if pending_amt > 0.5:
                details = doc_details_for_pending.get(doc, {"Date": "01/01/00", "Type": "Invoice", "Amt": pending_amt})
                age_days = "N/A"
                try:
                    inv_date_obj = datetime.datetime.strptime(details["Date"], "%d/%m/%y").date()
                    age_days = (today_date - inv_date_obj).days
                except: pass
                
                pending_invoices.append({
                    "Date": details["Date"], "Doc No": doc, "Type": details["Type"],
                    "Billed Amt": details["Amt"], "Pending Amt": round(pending_amt, 2),
                    "Age (Days)": age_days
                })
                total_pending_amt += pending_amt

        header_info = {"Time": run_datetime, "Period": period, "CustomerNo": customer_no, "CustomerName": customer_name}
        
        # जुना डिटेल समरी तक्ता परत आणला (Re-added Detailed Summary Info)
        summary_info = [
            ("OPENING BAL", opening_balance), ("Sales Invoice", s_inv), ("PAYMENT", pay), ("RECONCILIATION", reco),
            ("Credit Note(Others)", c_oth), ("Credit Note(Brakage Expiry)", c_brk), ("Goods Return Invoice", g_ret),
            ("Debit Note", d_not), ("TCS Debit Note", tcs), ("TDS Credit Note", tds),
            ("TECHNICAL BOUNCED", tech_b), ("NON TECHNICAL BOUNCED", n_tech_b), ("CLOSING BAL", running_balance)
        ]
        
        dash_info = {
            "Opening Bal": opening_balance, "Billed (Dr)": total_billed_dr, 
            "Paid / Adj (Cr)": total_paid_cr, "Closing Bal": running_balance, "Total Pending": total_pending_amt
        }
        return final_data, header_info, summary_info, dash_info, pending_invoices, None
    return None, None, None, None, None, "Data sapadla nahi."

def get_pdf_download_fpdf(final_data, header_info, summary_info, dash_info, pending_invoices, manual_name="", cfa_name=""):
    from fpdf import FPDF
    class PDF(FPDF):
        def footer(self):
            self.set_y(-15); self.set_font('Arial', 'I', 9)
            self.cell(0, 10, f'Page {self.page_no()} of {{nb}}', 0, 0, 'C')
    pdf = PDF(); pdf.alias_nb_pages(); pdf.set_auto_page_break(auto=False); pdf.add_page()
    def safe_str(s): return str(s).encode('latin1', 'ignore').decode('latin1')
    logo_file = next((f for f in ["logo.png", "Logo.png", "logo.jpg"] if os.path.exists(f)), None)
    if logo_file: pdf.image(logo_file, x=85, y=5, w=40); pdf.ln(15)
    else: pdf.ln(5)
    cust_name = manual_name if manual_name.strip() else header_info['CustomerName']
    cfa_text = f" | CFA Name: {cfa_name}" if cfa_name.strip() else ""
    pdf.set_font("Arial", 'B', 13); pdf.cell(190, 6, txt="VIRBAC - STATEMENT OF ACCOUNT", ln=True, align='C')
    pdf.set_font("Arial", size=10); pdf.cell(190, 6, txt=safe_str(f"Time: {header_info['Time']} | Period: {header_info['Period']}"), ln=True, align='C')
    pdf.cell(190, 6, txt=safe_str(f"Customer No: {header_info['CustomerNo']} | Customer Name: {cust_name}{cfa_text}"), ln=True, align='C'); pdf.ln(5)
    
    # 1. NEW TOP SUMMARY DASHBOARD
    dash_w = 190 / 5; dash_headers = ["Opening Bal", "Billed (Dr) +", "Paid / Adj (Cr) -", "Closing Bal =", "Total Pending"]
    pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(220, 235, 255)
    for h in dash_headers: pdf.cell(dash_w, 6, safe_str(h), border=1, align='C', fill=True)
    pdf.ln(); pdf.set_font("Arial", 'B', 9)
    dash_vals = [f"{int(dash_info['Opening Bal']):,}", f"{int(dash_info['Billed (Dr)']):,}", f"{int(dash_info['Paid / Adj (Cr)']):,}", f"{int(dash_info['Closing Bal']):,}", f"{int(dash_info['Total Pending']):,}"]
    for i, v in enumerate(dash_vals):
        if i == 4: pdf.set_text_color(200, 0, 0)
        pdf.cell(dash_w, 8, safe_str(v), border=1, align='C'); pdf.set_text_color(0, 0, 0)
    pdf.ln(8)
    
    # 2. OLD DETAILED TRANSACTION SUMMARY
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(100, 6, "Transaction Type", border=1, align='L')
    pdf.cell(40, 6, "Amount (INR)", border=1, ln=True, align='R')
    pdf.set_font("Arial", size=9)
    for row in summary_info:
        pdf.cell(100, 6, safe_str(row[0]), border=1, align='L')
        pdf.cell(40, 6, f"{int(float(row[1])):,}", border=1, ln=True, align='R')
    pdf.ln(5)
    
    # 3. MAIN TABLE
    col_widths = [13, 28, 16, 31, 15, 15, 17, 55] 
    headers = ["Date", "Type", "Doc No", "Chq/NEFT No", "Billed (Dr)", "Paid (Cr)", "Balance", "Remarks"]
    pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(240, 240, 240)
    for i in range(len(headers)): pdf.cell(col_widths[i], 6, safe_str(headers[i]), border=1, align='C', fill=True)
    pdf.ln()
    
    for r in final_data:
        wrapped_remarks = textwrap.wrap(safe_str(r['Remarks']), width=40) or [""]
        row_height = len(wrapped_remarks) * 6
        if pdf.get_y() + row_height > 275:
            pdf.add_page(); pdf.set_font("Arial", 'B', 8)
            for i in range(len(headers)): pdf.cell(col_widths[i], 6, safe_str(headers[i]), border=1, align='C', fill=True)
            pdf.ln()
        
        y_start = pdf.get_y(); temp_x = 10
        type_str = str(r['Type']).strip()
        is_payment = type_str in ["PAYMENT", "PAYMENT (Total)", "CLOSING BAL"]
        is_bounced = "BOUNCED" in type_str.upper()
        
        # Style Logic (Red color for Bounced Checks)
        if is_bounced: pdf.set_text_color(200, 0, 0); pdf.set_font("Arial", 'B', 7)
        elif is_payment: pdf.set_fill_color(240, 245, 250); pdf.set_font("Arial", 'B', 7)
        elif "-> Adj" in type_str or "-> On Acct" in type_str: pdf.set_font("Arial", 'I', 7); pdf.set_text_color(70, 70, 70)
        else: pdf.set_font("Arial", size=7); pdf.set_text_color(0, 0, 0)
        
        if is_payment: pdf.rect(temp_x, y_start, sum(col_widths), row_height, 'DF')
        else:
            cur_rect_x = temp_x
            for w in col_widths: pdf.rect(cur_rect_x, y_start, w, row_height, 'D'); cur_rect_x += w
            
        pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[0], 6, safe_str(r['Date']), align='C'); temp_x += col_widths[0]
        pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[1], 6, safe_str(r['Type'])[:28], align='L'); temp_x += col_widths[1]
        if "PAYMENT (Total)" in type_str:
            merged_w = col_widths[2] + col_widths[3]
            pdf.set_xy(temp_x, y_start); pdf.cell(merged_w, 6, safe_str(r['Chq No'])[:35], align='C'); temp_x += merged_w
        else:
            pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[2], 6, safe_str(r['Doc No']), align='C'); temp_x += col_widths[2]
            pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[3], 6, safe_str(r['Chq No'])[:30], align='C'); temp_x += col_widths[3]
        
        pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[4], 6, f"{int(float(r['Debit'])):,}" if r['Debit']!="" else "", align='R'); temp_x += col_widths[4]
        pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[5], 6, f"{int(float(r['Credit'])):,}" if r['Credit']!="" else "", align='R'); temp_x += col_widths[5]
        pdf.set_xy(temp_x, y_start); pdf.cell(col_widths[6], 6, f"{int(float(r['Balance'])):,}" if r['Balance']!="" else "", align='R'); temp_x += col_widths[6]
        for i, line in enumerate(wrapped_remarks):
            pdf.set_xy(temp_x, y_start + (i * 6)); pdf.cell(col_widths[7], 6, line, align='L')
        pdf.set_y(y_start + row_height); pdf.set_text_color(0, 0, 0)
        
    # Outstanding Table with AGING
    if pending_invoices:
        pdf.ln(5); pdf.set_font("Arial", 'B', 9); pdf.set_fill_color(255, 204, 204) 
        pdf.cell(190, 6, "OUTSTANDING / PENDING BILLS SUMMARY", border=1, ln=True, align='C', fill=True)
        p_col_widths = [20, 35, 45, 30, 30, 30]; p_headers = ["Date", "Doc No", "Type", "Billed (INR)", "Pending (INR)", "Age (Days)"]
        pdf.set_font("Arial", 'B', 8)
        for i in range(len(p_headers)): pdf.cell(p_col_widths[i], 6, safe_str(p_headers[i]), border=1, align='C')
        pdf.ln(); pdf.set_font("Arial", size=8); total_pending = 0.0
        for p in pending_invoices:
            if pdf.get_y() > 260: pdf.add_page()
            pdf.cell(p_col_widths[0], 6, safe_str(p['Date']), border=1, align='C')
            pdf.cell(p_col_widths[1], 6, safe_str(p['Doc No']), border=1, align='C')
            pdf.cell(p_col_widths[2], 6, safe_str(p['Type'])[:25], border=1, align='C')
            pdf.cell(p_col_widths[3], 6, f"{int(float(p['Billed Amt'])):,}", border=1, align='R')
            pdf.cell(p_col_widths[4], 6, f"{int(float(p['Pending Amt'])):,}", border=1, align='R')
            
            # Age highlight Logic
            if isinstance(p['Age (Days)'], int) and p['Age (Days)'] > 30: pdf.set_font("Arial", 'B', 8)
            pdf.cell(p_col_widths[5], 6, safe_str(p['Age (Days)']), border=1, align='C'); pdf.set_font("Arial", size=8)
            pdf.ln(); total_pending += p['Pending Amt']
            
        pdf.set_font("Arial", 'B', 8); pdf.cell(sum(p_col_widths[:4]), 6, "Total Outstanding:", border=1, align='R')
        pdf.cell(p_col_widths[4], 6, f"{int(total_pending):,}", border=1, align='R')
        pdf.cell(p_col_widths[5], 6, "", border=1, ln=True)

    if pdf.get_y() > 240: pdf.add_page()
    pdf.ln(5); pdf.set_font("Arial", 'I', 8); pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(190, 5, safe_str("Note: This is a computer generated statement. If you have any queries or discrepancies regarding this statement, please contact our CFA / Accounts team within 7 days of receipt. Thank you for your business!"), align='L')
    pdf.ln(5); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", 'B', 11)
    pdf.cell(190, 6, "For Virbac Animal Health India Pvt Ltd", ln=True); pdf.ln(5)
    if cfa_name.strip(): pdf.cell(190, 6, safe_str(f"Authorized Signatory: {cfa_name}"), ln=True)
    pdf.ln(5); pdf.cell(190, 6, "Signature                  Place: ___________      Date: ___________", ln=True)
    return bytes(pdf.output(dest='S').encode('latin1', 'ignore'))

def get_excel_download(final_data, header_info, summary_info, dash_info, pending_invoices, manual_name="", cfa_name=""):
    output = io.BytesIO(); cust_name = manual_name if manual_name.strip() else header_info['CustomerName']
    excel_data = [{"Date": r["Date"], "Type": r["Type"], "Doc No": r["Doc No"], "Chq/NEFT No": r["Chq No"], "Billed Amt (Dr)": r["Debit"], "Paid Amt (Cr)": r["Credit"], "Balance": r["Balance"], "Remarks": r["Remarks"]} for r in final_data]
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([["Report Time:", header_info["Time"]], ["Heading:", "STATEMENT OF ACCOUNT"], ["Period:", header_info["Period"]], ["Customer No:", header_info["CustomerNo"]], ["Customer Name:", cust_name], ["CFA Name:", cfa_name]]).to_excel(writer, sheet_name='Statement', index=False, header=False)
        
        # Dashboard in Excel
        pd.DataFrame([["Opening Bal", "Billed (Dr)", "Paid / Adj (Cr)", "Closing Bal", "Total Pending"], [dash_info['Opening Bal'], dash_info['Billed (Dr)'], dash_info['Paid / Adj (Cr)'], dash_info['Closing Bal'], dash_info['Total Pending']]]).to_excel(writer, sheet_name='Statement', index=False, header=False, startrow=8)
        
        # Old Detailed Summary Table
        pd.DataFrame(summary_info, columns=["Transaction Type", "Amount (INR)"]).to_excel(writer, sheet_name='Statement', index=False, startrow=12)
        
        # Main Data
        start_main = 12 + len(summary_info) + 3
        pd.DataFrame(excel_data).to_excel(writer, sheet_name='Statement', index=False, startrow=start_main)
        
        if pending_invoices:
            start_row_pending = start_main + len(excel_data) + 3
            pd.DataFrame([["OUTSTANDING / PENDING BILLS SUMMARY (with Aging)"]]).to_excel(writer, sheet_name='Statement', index=False, header=False, startrow=start_row_pending)
            pd.DataFrame(pending_invoices).to_excel(writer, sheet_name='Statement', index=False, startrow=start_row_pending + 1)
    return output.getvalue()

if uploaded_files:
    for file in uploaded_files:
        data, h_info, s_info, dash_info, pending_inv, err = process_pdf_logic(file)
        if err: st.error(err)
        else:
            st.success(f"✅ {file.name} ready!")
            st.write("### 📊 Dashboard Summary")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Opening Bal", f"₹ {int(dash_info['Opening Bal']):,}")
            c2.metric("Billed (Dr)", f"₹ {int(dash_info['Billed (Dr)']):,}")
            c3.metric("Paid (Cr)", f"₹ {int(dash_info['Paid / Adj (Cr)']):,}")
            c4.metric("Closing Bal", f"₹ {int(dash_info['Closing Bal']):,}")
            c5.metric("Total Pending", f"₹ {int(dash_info['Total Pending']):,}", delta_color="inverse")
            st.write("---")
            col_in1, col_in2 = st.columns(2)
            with col_in1: manual_name = st.text_input("Customer Name (optional):", key=f"cust_{file.name}")
            with col_in2: cfa_name = st.text_input("CFA Name (optional):", key=f"cfa_{file.name}")
            if st.button("✅ फाईल तयार करा (Prepare Files)", key=f"btn_{file.name}"): st.session_state[f"ready_{file.name}"] = True
            if st.session_state.get(f"ready_{file.name}", False):
                st.write("---")
                col1, col2 = st.columns(2)
                with col1: st.download_button("📥 Excel Download", get_excel_download(data, h_info, s_info, dash_info, pending_inv, manual_name, cfa_name), f"{file.name}.xlsx", key=f"dl_xl_{file.name}")
                with col2: st.download_button("📥 PDF Download", get_pdf_download_fpdf(data, h_info, s_info, dash_info, pending_inv, manual_name, cfa_name), f"{file.name}.pdf", key=f"dl_pdf_{file.name}")
