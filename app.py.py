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

st.set_page_config(page_title="Virbac Statement Converter", page_icon="📄", layout="wide")
st.title("📄 Virbac Account Statement Converter (Smart Tracking)")
st.markdown("CFA Team sathi: PDF upload kara ani **Excel + PDF** donhi milva.")

uploaded_files = st.file_uploader("Upload PDF file here", type="pdf", accept_multiple_files=True)

def process_pdf_logic(uploaded_file):
    uploaded_file.seek(0)
    run_datetime = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
    period, customer_no, customer_name = "", "", ""
    
    with pdfplumber.open(uploaded_file) as pdf:
        first_page = pdf.pages[0]
        clean_text = " ".join((first_page.extract_text(layout=True) or "").split())
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
                            if len(candidate) < 4: continue
                            ign_kw = ['date', 'time', 'page', 'statement', 'accounting', 'period', 'payer', 'customer', 'limit', 'opening', 'bal', 'dt', 'balance']
                            if any(x in candidate.lower() for x in ign_kw): continue
                            customer_name = candidate
                            break
                    if customer_name: break

    extracted_rows = []
    uploaded_file.seek(0)
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            p_text = page.extract_text()
            if p_text:
                for line in p_text.split('\n'): extracted_rows.append(line.strip())

    final_data, running_balance, opening_balance, found_opening = [], 0.0, 0.0, False
    total_billed_dr, total_paid_cr = 0.0, 0.0
    s_inv, pay, reco, c_oth, c_brk, g_ret, d_not, tcs, tds, tech_b, n_tech_b = [0.0]*11

    for idx, row_text in enumerate(extracted_rows):
        row_upper = row_text.upper()
        if "OPENING BALANCE" in row_upper and not found_opening:
            amounts = re.findall(r'-?\s*\(?\s*[\d,]+\.\d{2}\s*\)?', row_text)
            if amounts:
                amt_str = amounts[-1]
                is_cr = 'CR' in row_upper or '(' in amt_str or '-' in amt_str
                if not is_cr:
                    for offset in range(1, 3):
                        if idx + offset < len(extracted_rows):
                            nxt = extracted_rows[idx + offset].upper().strip()
                            if nxt in ['CR', '(CR)', 'CR.']: is_cr = True; break
                            elif re.search(r'\d', nxt): break 
                val = float(re.sub(r'[^\d.]', '', amt_str))
                opening_balance = val
                running_balance = -val if is_cr else val
                found_opening = True
                final_data.append({"Date": "", "Type": "OPENING BAL", "Doc No": "", "Chq No": "", "Debit": "" if is_cr else val, "Credit": val if is_cr else "", "Balance": round(running_balance, 2), "Remarks": ""})
            continue

        date_match = re.search(r'\b(\d{2}/\d{2}/\d{2})\b', row_text)
        if date_match:
            if any(x in row_upper for x in ["DATE:", "ACCOUNTING DATE", "TIME:", "PAGE"]): continue
            date = date_match.group(1)
            amounts = re.findall(r'-?\(?[\d,]+\.\d{2}\)?', row_text)
            if not amounts: continue
            amt_str = amounts[-1]
            is_cr = 'CR' in row_upper or '(' in amt_str or '-' in amt_str
            val = float(re.sub(r'[^\d.]', '', amt_str))
            if val == 0.0: continue
            
            doc_match = re.search(r'\b\d{9,10}\b', row_text)
            doc_no = doc_match.group(0) if doc_match else ""
            t_type, s_type, remarks = "Other", "Other", ""
            
            if ("TDSRECO" in row_upper and doc_no.startswith('000') and is_cr) or ("TDS CREDIT NOTE" in row_upper): t_type = s_type = "TDS Credit Note"
            elif "CBOU199" in row_upper: t_type = s_type = "TECHNICAL BOUNCED"
            elif any(c in row_upper for c in ["CBOU101", "CBOU102", "CBOU110"]): t_type = s_type = "NON TECHNICAL BOUNCED"
            elif "RECONC" in row_upper: t_type = s_type = "Reconciliation"
            elif any(x in row_upper for x in ["CHQ", "PAYMENT", "DD-NEFT", "NEFT"]) or (doc_no.startswith('000') and is_cr): t_type = s_type = "PAYMENT"
            elif "INVOICE" in row_upper:
                if doc_no.startswith('3'): t_type = s_type = "Credit Note(Brakage Expiry)"
                elif doc_no.startswith('4'): t_type = s_type = "Credit Note(Others)"
                elif doc_no.startswith('5'): t_type = s_type = "Debit Note"
                elif doc_no.startswith('8'): t_type = s_type = "TCS Debit Note"
                elif is_cr: t_type = s_type = "Goods Return Invoice"
                else: t_type = s_type = "Sales Invoice"

            chq_no = ""
            if s_type in ["PAYMENT", "TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"] or any(x in row_upper for x in ["CHQ", "NEFT", "DD", "RTGS"]):
                for token in row_text.split():
                    if re.match(r'\d{2}/\d{2}/\d{2}', token) or token == date: continue
                    if re.search(r'\.\d{2}\)?$', token) or re.search(r'\d,\d', token): continue
                    ctok = re.sub(r'[^A-Za-z0-9]', '', token)
                    if not ctok or ctok == doc_no: continue
                    if ctok.isdigit() and len(ctok) == 6: chq_no = token; break
                    elif re.match(r'^[A-Za-z0-9]{8,25}$', ctok):
                        ig = ["PAYMENT", "RECONC", "INVOICE", "OPENING", "BALANCE", "CLOSING", "TECHNICAL", "BOUNCED", "NON"]
                        if ctok.upper() not in ig and not ctok.upper().startswith("CBOU"): chq_no = token; break

            debit, credit = (0.0, val) if is_cr else (val, 0.0)
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

            final_data.append({"Date": date, "Type": t_type, "Doc No": doc_no, "Chq No": chq_no, "Debit": debit if debit > 0 else "", "Credit": credit if credit > 0 else "", "Balance": "", "Remarks": remarks})

    if not final_data: return None, None, None, None, None, "Data sapadla nahi (No data found)."

    def rsk(x):
        t = str(x.get('Type', ''))
        if 'OPENING BAL' in t: return (0, '', 0)
        elif 'PAYMENT' in t or 'BOUNCED' in t: return (2, str(x.get('Chq No', '')), 0)
        elif 'Advance Utilized' in t or ('Reconciliation' in t and str(x.get('Debit', '')) != ""): return (3, '1', 0)
        elif 'Bill Settled' in t or ('Reconciliation' in t and str(x.get('Credit', '')) != "" and 'Advance Carried Forward' not in t): return (3, '2', 0)
        elif 'Advance Carried Forward' in t: return (3, '3', 0)
        return (1, '', 0)

    reord = []
    curr_dt = None
    blk = []
    for r in final_data:
        if r['Type'] == 'OPENING BAL': reord.append(r); continue
        if r['Date'] != curr_dt:
            if blk: blk.sort(key=rsk); reord.extend(blk)
            curr_dt = r['Date']
            blk = [r]
        else: blk.append(r)
    if blk: blk.sort(key=rsk); reord.extend(blk)
    final_data = reord
    
    cbal = -opening_balance if final_data[0].get("Credit") else opening_balance
    for r in final_data:
        if r['Type'] == 'OPENING BAL': r['Balance'] = round(cbal, 2); continue
        cbal += (float(r['Debit'] or 0) - float(r['Credit'] or 0))
        r['Balance'] = round(cbal, 2)
    running_balance = cbal

    for i in range(len(final_data)):
        if "BOUNCED" in final_data[i]["Type"] and not final_data[i]["Chq No"]:
            bamt = final_data[i]["Debit"] or final_data[i]["Credit"]
            bdoc = final_data[i]["Doc No"]
            for j in range(i - 1, -1, -1):
                if final_data[j]["Type"] == "PAYMENT":
                    pamt = final_data[j]["Credit"] or final_data[j]["Debit"]
                    if (bamt and bamt == pamt) or (bdoc and bdoc == final_data[j]["Doc No"]):
                        if final_data[j]["Chq No"]: final_data[i]["Chq No"] = final_data[j]["Chq No"]; break

    inv_bal = {}
    doc_det = {}
    for r in final_data:
        tc = r["Type"]
        if tc not in ["PAYMENT", "Reconciliation", "OPENING BAL", "CLOSING BAL"] and "BOUNCED" not in tc:
            if r["Doc No"]:
                amt = r["Debit"] or r["Credit"]
                if amt:
                    inv_bal[r["Doc No"]] = float(amt)
                    doc_det[r["Doc No"]] = {"Date": r["Date"], "Type": tc, "Amt": float(amt)}

    adv_bal = {r["Doc No"]: float(r["Credit"]) for r in final_data if r["Doc No"] and r["Doc No"].startswith('000') and r["Credit"]}
    rec_dr_map = {}
    ret_deb_dt = {}  
    
    for r in final_data:
        if r["Type"] == "Reconciliation" and r["Debit"]:
            dval = float(r["Debit"])
            doc = r.get("Doc No", "")
            rec_dr_map[(r["Date"], dval)] = doc
            tcat = doc_det.get(doc, {}).get("Type", "")
            if doc and (doc.startswith(('22', '3', '4', '9')) or "Return" in tcat or "Credit Note" in tcat):
                ret_deb_dt.setdefault(r["Date"], []).append(dval)

    chq_s, bnc_s = {}, {}
    for r in final_data:
        cn = r.get("Chq No")
        if cn:
            k = (r["Date"], cn)
            camt = float(r["Credit"] or 0)
            damt = float(r["Debit"] or 0)
            if "PAYMENT" in r["Type"]:
                chq_s.setdefault(k, {'tc':0, 'td':0, 'cnt':0, 'seen':0})
                chq_s[k]['tc'] += camt; chq_s[k]['td'] += damt; chq_s[k]['cnt'] += 1
            elif "BOUNCED" in r["Type"]:
                bnc_s.setdefault(k, {'tc':0, 'td':0, 'cnt':0, 'seen':0})
                bnc_s[k]['tc'] += camt; bnc_s[k]['td'] += damt; bnc_s[k]['cnt'] += 1

    grp = []
    for r in final_data:
        dn = r.get("Doc No", "")
        if "PAYMENT" in r["Type"]:
            amt = float(r["Credit"] or r["Debit"] or 0)
            stag = ""
            if dn and dn in inv_bal:
                inv_bal[dn] -= amt
                stag = " [CLEARED]" if inv_bal[dn] <= 0.5 else f" [Pend: {int(inv_bal[dn]):,}]"
            oamt = doc_det.get(dn, {}).get("Amt")
            ostr = f" | Inv Amt: {int(oamt):,}" if oamt else ""
            dtp = "Cr Note" if dn.startswith(('3','4','9')) else "Dr Note" if dn.startswith('5') else "TCS Dr Note" if dn.startswith('8') else "Adv" if dn.startswith('000') else "Inv"
            r["Remarks"] = f"Total Advance Received: {int(amt):,}" if dn.startswith('000') else f"{dtp}: {dn}{ostr} | Adj: {int(amt):,}{stag}" if dn else f"Adj: {int(amt):,}"
            r["Type"] = "PAYMENT"
            grp.append(r)
            cn = r.get("Chq No")
            if cn and chq_s.get((r["Date"], cn)):
                k = (r["Date"], cn)
                chq_s[k]['seen'] += 1
                if chq_s[k]['seen'] == chq_s[k]['cnt']:
                    grp.append({"Date": "", "Type": "-> CHQ/NEFT SUMMARY", "Doc No": "", "Chq No": cn, "Debit": "", "Credit": "", "Balance": "", "Remarks": f"Total Inv Adj: {int(max(chq_s[k]['tc'], chq_s[k]['td'])):,} | Total Chq/NEFT Amt: {int(abs(chq_s[k]['tc'] - chq_s[k]['td'])):,}"})
        elif "BOUNCED" in r["Type"]:
            grp.append(r)
            cn = r.get("Chq No")
            if cn and bnc_s.get((r["Date"], cn)):
                k = (r["Date"], cn)
                bnc_s[k]['seen'] += 1
                if bnc_s[k]['seen'] == bnc_s[k]['cnt']:
                    grp.append({"Date": "", "Type": "-> BOUNCED CHQ/NEFT SUMMARY", "Doc No": "", "Chq No": cn, "Debit": "", "Credit": "", "Balance": "", "Remarks": f"Total Bounced Chq/NEFT Amt: {int(abs(bnc_s[k]['tc'] - bnc_s[k]['td'])):,}"})
        elif "Reconciliation" in r["Type"]:
            amt = float(r["Credit"] or r["Debit"] or 0)
            if r["Debit"]:
                if dn in inv_bal: inv_bal[dn] += amt 
                tcat = doc_det.get(dn, {}).get("Type", "")
                if dn and (dn.startswith(('22', '3', '4', '9')) or "Return" in tcat or "Credit Note" in tcat):
                    r["Type"] = ">> System Adj (Reconciliation)"
                    r["Remarks"] = "Adjusted from Return | System entry (Ignore)"
                elif dn and dn.startswith('000'):
                    r["Type"] = ">> Advance Utilized"
                    if dn in adv_bal:
                        adv_bal[dn] -= amt 
                        r["Remarks"] = f"Deducted from Adv: {dn} | Rem Adv Bal: {int(adv_bal[dn]):,}"
                    else:
                        r["Remarks"] = f"Deducted from Opening Balance | Adj: {int(amt):,}" if abs(amt - abs(opening_balance)) < 0.5 else f"Deducted from Old Adv (Opening Bal) | Adj: {int(amt):,}"
                else: r["Type"] = "Reconciliation"; r["Remarks"] = f"Debit Adj: {dn} | Amt: {int(amt):,}"
            else:
                sdoc = rec_dr_map.get((r["Date"], amt), "")
                sdet = doc_det.get(sdoc, {}).get("Type", "")
                is_ret = sdoc and (sdoc.startswith(('22', '3', '4', '9')) or "Return" in sdet)
                if not is_ret and ret_deb_dt.get(r["Date"]):
                    rd = ret_deb_dt[r["Date"]]
                    if abs(sum(rd) - amt) < 0.5: is_ret = True
                    elif len(rd) <= 15:
                        for clen in range(2, len(rd) + 1):
                            if any(abs(sum(c) - amt) < 0.5 for c in itertools.combinations(rd, clen)): is_ret = True; break
                stag = ""
                if dn in inv_bal:
                    inv_bal[dn] -= amt
                    stag = " [CLEARED]" if inv_bal[dn] <= 0.5 else f" [Pend: {int(inv_bal[dn]):,}]"
                if dn and dn.startswith('000'):
                    r["Type"] = ">> Advance Carried Forward"
                    r["Remarks"] = f"Advance Carried Forward: {dn} | Amt: {int(amt):,}"
                else:
                    r["Type"] = ">> Bill Settled (Reconciliation)"
                    oamt = doc_det.get(dn, {}).get("Amt")
                    ostr = f" | Amt: {int(oamt):,}" if oamt else ""
                    dtp = "Cr Note" if dn.startswith(('3','4','9')) else "Dr Note" if dn.startswith('5') else "TCS Dr Note" if dn.startswith('8') else "Inv"
                    r["Remarks"] = f"Bill Settled Against Return | Inv: {dn}{ostr} | Adj: {int(amt):,}{stag}" if is_ret else f"{dtp}: {dn}{ostr} | Adj: {int(amt):,}{stag}"
            grp.append(r)
        else:
            if "Goods Return Invoice" in r["Type"]: r["Remarks"] = "Material Returned by Customer"
            elif "Credit Note(Brakage Expiry)" in r["Type"]: r["Remarks"] = "Credit Note Issued for Breakage/Expiry"
            elif "Credit Note(Others)" in r["Type"]: r["Remarks"] = "Credit Note Issued (Others)"
            grp.append(r)

    final_data = grp
    re_sort = []
    cdb = []
    cdt = None
    for r in final_data:
        if str(r.get("Type", "")) in ["OPENING BAL", "CLOSING BAL", "-> BATCH SETTLEMENT SUMMARY"]:
            if cdb: cdb.sort(key=rsk); re_sort.extend(cdb); cdb = []
            re_sort.append(r)
            continue
        if r["Date"] != cdt:
            if cdb: cdb.sort(key=rsk); re_sort.extend(cdb)
            cdt = r["Date"]
            cdb = [r]
        else: cdb.append(r)
    if cdb: cdb.sort(key=rsk); re_sort.extend(cdb)
    final_data = re_sort
    
    inj = []
    for dt_k, g in itertools.groupby(final_data, key=lambda x: x["Date"]):
        gl = list(g)
        rr = [x for x in gl if any(k in str(x["Type"]) for k in ["Reconciliation", "Advance Carried Forward", "Advance Utilized"])]
        inj.extend(gl)
        if rr and dt_k:
            ta = sum(float(x["Debit"] or 0) for x in rr if "Advance Utilized" in str(x["Type"]))
            tb = sum(float(x["Credit"] or 0) for x in rr if "Bill Settled" in str(x["Type"]))
            tcf = sum(float(x["Credit"] or 0) for x in rr if "Advance Carried Forward" in str(x["Type"]))
            if ta > 0 or tb > 0:
                rs = f"Total Adv Used: {int(ta):,} | Total Bills Paid: {int(tb):,}" + (f" | New Adv Created: {int(tcf):,}" if tcf > 0 else "")
                inj.append({"Date": "", "Type": "-> BATCH SETTLEMENT SUMMARY", "Doc No": "", "Chq No": "", "Debit": "", "Credit": "", "Balance": "", "Remarks": rs})
    
    inj.append({"Date": "", "Type": "CLOSING BAL", "Doc No": "", "Chq No": "", "Debit": "", "Credit": "", "Balance": round(running_balance, 2), "Remarks": ""})
    final_data = inj
    
    pinv = []
    for doc, pamt in inv_bal.items():
        if pamt > 0.5:
            d = doc_det.get(doc, {"Date": "01/01/00", "Type": "Invoice", "Amt": pamt})
            is_cr = doc.startswith(('3', '4', '9')) or "Credit Note" in d["Type"] or "Goods Return" in d["Type"]
            pinv.append({"Date": d["Date"], "Doc No": doc, "Type": d["Type"], "Billed Amt": -d["Amt"] if is_cr else d["Amt"], "Pending Amt": round(-pamt if is_cr else pamt, 2)})

    diff = running_balance - sum(p["Pending Amt"] for p in pinv)
    if abs(diff) > 0.5: pinv.append({"Date": "-", "Doc No": "SYSTEM BAL", "Type": "Unadjusted Advance" if diff < 0 else "System Debit", "Billed Amt": round(diff, 2), "Pending Amt": round(diff, 2)})

    hinf = {"Time": run_datetime, "Period": period, "CustomerNo": customer_no, "CustomerName": customer_name}
    sinf = [("OPENING BAL", opening_balance), ("Sales Invoice", s_inv), ("PAYMENT", pay), ("Reconciliation", reco), ("Credit Note(Others)", c_oth), ("Credit Note(Brakage Expiry)", c_brk), ("Goods Return Invoice", g_ret), ("Debit Note", d_not), ("TCS Debit Note", tcs), ("TDS Credit Note", tds), ("TECHNICAL BOUNCED", tech_b), ("NON TECHNICAL BOUNCED", n_tech_b), ("CLOSING BAL", running_balance)]
    dinf = {"Opening Bal": opening_balance, "Billed (Dr)": total_billed_dr, "Paid / Adj (Cr)": total_paid_cr, "Closing Bal": running_balance, "Total Pending": running_balance}
    return final_data, hinf, sinf, dinf, pinv, None

def get_pdf_download_fpdf(f_data, h_info, s_info, d_info, p_inv, m_name="", c_name=""):
    from fpdf import FPDF
    class PDF(FPDF):
        def header(self):
            s_s = lambda x: str(x).encode('latin1', 'ignore').decode('latin1')
            lf = next((f for f in ["logo.png", "Logo.png", "logo.jpg"] if os.path.exists(f)), None)
            if lf: self.image(lf, 85, 5, 40); self.ln(15)
            else: self.ln(5)
            hi = getattr(self, 'h_info', {})
            cnm = getattr(self, 'm_name', '') or hi.get('CustomerName','')
            cfa = f" | CFA Name: {getattr(self, 'c_name', '')}" if getattr(self, 'c_name', '') else ""
            self.set_font("Arial", 'B', 13)
            self.cell(190, 6, "VIRBAC - STATEMENT OF ACCOUNT", ln=True, align='C')
            self.set_font("Arial", size=10)
            self.cell(190, 6, s_s(f"Time: {hi.get('Time','')} | Period: {hi.get('Period','')}"), ln=True, align='C')
            self.cell(190, 6, s_s(f"Customer No: {hi.get('CustomerNo','')} | Customer Name: {cnm}{cfa}"), ln=True, align='C')
            self.ln(5)
        def footer(self):
            self.set_y(-15); self.set_font('Arial', 'I', 9); self.cell(0, 10, f'Page {self.page_no()} of {{nb}}', 0, 0, 'C')

    pdf = PDF(); pdf.h_info, pdf.m_name, pdf.c_name = h_info, m_name, c_name
    pdf.alias_nb_pages(); pdf.set_auto_page_break(False); pdf.add_page()
    ss = lambda x: str(x).encode('latin1', 'ignore').decode('latin1')
    
    dw = 190 / 5
    pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(220, 235, 255)
    for h in ["Opening Bal", "Billed (Dr) +", "Paid/Adj(Cr)-", "Closing Bal", "Total Pending"]: pdf.cell(dw, 6, ss(h), 1, 0, 'C', True)
    pdf.ln(); pdf.set_font("Arial", 'B', 9)
    for i, v in enumerate([d_info['Opening Bal'], d_info['Billed (Dr)'], d_info['Paid / Adj (Cr)'], d_info['Closing Bal'], d_info['Total Pending']]):
        if i == 4: pdf.set_text_color(200, 0, 0)
        pdf.cell(dw, 8, f"{int(v):,}", 1, 0, 'C')
        pdf.set_text_color(0, 0, 0)
    pdf.ln(8); pdf.cell(100, 6, "Transaction Type", 1, 0, 'L'); pdf.cell(40, 6, "Amount (INR)", 1, 1, 'R')
    pdf.set_font("Arial", size=9)
    for r in s_info: pdf.cell(100, 6, ss(r[0]), 1, 0, 'L'); pdf.cell(40, 6, f"{int(float(r[1])):,}", 1, 1, 'R')
    pdf.ln(5)
    
    cw = [13, 35, 16, 31, 15, 15, 17, 48] 
    headers = ["Date", "Type", "Doc No", "Chq/NEFT", "Billed", "Paid", "Balance", "Remarks"]
    pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(240, 240, 240)
    for i in range(8): pdf.cell(cw[i], 6, ss(headers[i]), 1, 0, 'C', True)
    pdf.ln()
    
    for r in f_data:
        wrem = textwrap.wrap(ss(r['Remarks']), 36) or [""]
        rh = len(wrem) * 6
        if pdf.get_y() + rh > 275:
            pdf.add_page(); pdf.set_font("Arial", 'B', 8)
            for i in range(8): pdf.cell(cw[i], 6, ss(headers[i]), 1, 0, 'C', True)
            pdf.ln()
        
        ys, tx = pdf.get_y(), 10
        ts = str(r['Type']).upper()
        fill = "SUMMARY" in ts or "CLOSING BAL" in ts
        if fill: pdf.set_fill_color(240, 245, 250); pdf.set_font("Arial", 'B', 7); pdf.set_text_color(0, 0, 0)
        elif "BOUNCED" in ts and "SUMMARY" not in ts: pdf.set_text_color(200, 0, 0); pdf.set_font("Arial", 'B', 7)
        else: pdf.set_font("Arial", size=7); pdf.set_text_color(0, 0, 0)
        st = 'DF' if fill else 'D'
        
        if "SUMMARY" in ts or "CLOSING BAL" in ts:
            mw = sum(cw[2:7])
            pdf.rect(tx, ys, cw[0], rh, st); pdf.rect(tx+cw[0], ys, cw[1], rh, st); pdf.rect(tx+cw[0]+cw[1], ys, mw, rh, st); pdf.rect(tx+cw[0]+cw[1]+mw, ys, cw[7], rh, st)
            pdf.set_xy(tx, ys); pdf.cell(cw[0], 6, ss(r['Date']), 0, 0, 'C')
            pdf.set_xy(tx+cw[0], ys); pdf.cell(cw[1], 6, ss(r['Type'])[:35], 0, 0, 'L')
            pdf.set_xy(tx+cw[0]+cw[1], ys); pdf.cell(mw, 6, ss(r['Chq No']), 0, 0, 'C')
            for i, l in enumerate(wrem): pdf.set_xy(tx+cw[0]+cw[1]+mw, ys+(i*6)); pdf.cell(cw[7], 6, l, 0, 0, 'L')
        else:
            cx = tx
            for w in cw: pdf.rect(cx, ys, w, rh, st); cx += w
            pdf.set_xy(tx, ys); pdf.cell(cw[0], 6, ss(r['Date']), 0, 0, 'C'); tx+=cw[0]
            pdf.set_xy(tx, ys); pdf.cell(cw[1], 6, ss(r['Type'])[:35], 0, 0, 'L'); tx+=cw[1]
            pdf.set_xy(tx, ys); pdf.cell(cw[2], 6, ss(r['Doc No']), 0, 0, 'C'); tx+=cw[2]
            pdf.set_xy(tx, ys); pdf.cell(cw[3], 6, ss(r['Chq No'])[:30], 0, 0, 'C'); tx+=cw[3]
            pdf.set_xy(tx, ys); pdf.cell(cw[4], 6, f"{int(float(r['Debit'])):,}" if r['Debit']!="" else "", 0, 0, 'R'); tx+=cw[4]
            pdf.set_xy(tx, ys); pdf.cell(cw[5], 6, f"{int(float(r['Credit'])):,}" if r['Credit']!="" else "", 0, 0, 'R'); tx+=cw[5]
            pdf.set_xy(tx, ys); pdf.cell(cw[6], 6, f"{int(float(r['Balance'])):,}" if r['Balance']!="" else "", 0, 0, 'R'); tx+=cw[6]
            for i, l in enumerate(wrem): pdf.set_xy(tx, ys+(i*6)); pdf.cell(cw[7], 6, l, 0, 0, 'L')
        pdf.set_y(ys + rh); pdf.set_text_color(0, 0, 0)
        
    if p_inv:
        pdf.ln(5); pdf.set_font("Arial", 'B', 9); pdf.set_fill_color(255, 204, 204) 
        pdf.cell(190, 6, "OUTSTANDING / PENDING BILLS SUMMARY", 1, 1, 'C', True)
        pw, ph = [25, 40, 55, 35, 35], ["Date", "Doc No", "Type", "Billed (INR)", "Pending (INR)"]
        pdf.set_font("Arial", 'B', 8)
        for i in range(5): pdf.cell(pw[i], 6, ss(ph[i]), 1, 0, 'C', True)
        pdf.ln(); pdf.set_font("Arial", size=8); tp = 0.0
        for p in p_inv:
            if pdf.get_y() > 260: 
                pdf.add_page(); pdf.set_font("Arial", 'B', 8); pdf.set_fill_color(255, 204, 204)
                for i in range(5): pdf.cell(pw[i], 6, ss(ph[i]), 1, 0, 'C', True)
                pdf.ln(); pdf.set_font("Arial", size=8)
            pdf.cell(pw[0], 6, ss(p['Date']), 1, 0, 'C'); pdf.cell(pw[1], 6, ss(p['Doc No']), 1, 0, 'C'); pdf.cell(pw[2], 6, ss(p['Type'])[:25], 1, 0, 'C')
            pdf.cell(pw[3], 6, f"{int(float(p['Billed Amt'])):,}" if p['Billed Amt']!="" else "", 1, 0, 'R'); pdf.cell(pw[4], 6, f"{int(float(p['Pending Amt'])):,}" if p['Pending Amt']!="" else "", 1, 1, 'R')
            tp += p['Pending Amt']
        pdf.set_font("Arial", 'B', 8); pdf.cell(sum(pw[:4]), 6, "Total Outstanding:", 1, 0, 'R'); pdf.cell(pw[4], 6, f"{int(tp):,}", 1, 1, 'R')

    if pdf.get_y() > 240: pdf.add_page()
    pdf.ln(5); pdf.set_font("Arial", 'I', 8); pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(190, 5, ss("Note: This is a computer generated statement. If you have any queries, please contact our CFA / Accounts team within 7 days. Thank you for your business!"), 0, 'L')
    pdf.ln(5); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", 'B', 11); pdf.cell(190, 6, "For Virbac Animal Health India Pvt Ltd", ln=True); pdf.ln(5)
    if c_name.strip(): pdf.cell(190, 6, ss(f"Authorized Signatory: {c_name}"), ln=True); pdf.ln(5)
    pdf.cell(190, 6, "Signature                  Place: ___________      Date: ___________", ln=True)
    return bytes(pdf.output(dest='S').encode('latin1', 'ignore'))

def get_excel_download(f_data, h_info, s_info, d_info, p_inv, m_name="", c_name=""):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        pd.DataFrame([["Report Time:", h_info["Time"]], ["Heading:", "STATEMENT OF ACCOUNT"], ["Period:", h_info["Period"]], ["Customer No:", h_info["CustomerNo"]], ["Customer Name:", m_name or h_info['CustomerName']], ["CFA Name:", c_name]]).to_excel(w, sheet_name='Statement', index=False, header=False)
        pd.DataFrame([["Opening Bal", "Billed (Dr)", "Paid / Adj (Cr)", "Closing Bal", "Total Pending"], [d_info['Opening Bal'], d_info['Billed (Dr)'], d_info['Paid / Adj (Cr)'], d_info['Closing Bal'], d_info['Total Pending']]]).to_excel(w, sheet_name='Statement', index=False, header=False, startrow=8)
        pd.DataFrame(s_info, columns=["Transaction Type", "Amount (INR)"]).to_excel(w, sheet_name='Statement', index=False, startrow=12)
        sm = 12 + len(s_info) + 3
        pd.DataFrame([{"Date": r["Date"], "Type": r["Type"], "Doc No": r["Doc No"], "Chq/NEFT": r["Chq No"], "Billed(Dr)": r["Debit"], "Paid(Cr)": r["Credit"], "Balance": r["Balance"], "Remarks": r["Remarks"]} for r in f_data]).to_excel(w, sheet_name='Statement', index=False, startrow=sm)
        if p_inv:
            sp = sm + len(f_data) + 3
            pd.DataFrame([["OUTSTANDING / PENDING BILLS SUMMARY"]]).to_excel(w, sheet_name='Statement', index=False, header=False, startrow=sp)
            pd.DataFrame(p_inv).to_excel(w, sheet_name='Statement', index=False, startrow=sp + 1)
    return out.getvalue()

if uploaded_files:
    for f in uploaded_files:
        with st.spinner(f"⏳ {f.name} प्रोसेस होत आहे..."):
            try:
                data, h_info, s_info, d_info, p_inv, err = process_pdf_logic(f)
                if err: st.error(f"❌ त्रुटी: {err}")
                else:
                    st.success(f"✅ {f.name} processed successfully!")
                    c1, c2, c3, c4, c5 = st.columns(5)
                    c1.metric("Opening Bal", f"₹ {int(d_info.get('Opening Bal',0)):,}")
                    c2.metric("Billed (Dr)", f"₹ {int(d_info.get('Billed (Dr)',0)):,}")
                    c3.metric("Paid (Cr)", f"₹ {int(d_info.get('Paid / Adj (Cr)',0)):,}")
                    c4.metric("Closing Bal", f"₹ {int(d_info.get('Closing Bal',0)):,}")
                    c5.metric("Total Pending", f"₹ {int(d_info.get('Total Pending',0)):,}", delta_color="inverse")
                    st.write("---")
                    ci1, ci2 = st.columns(2)
                    with ci1: mn = st.text_input("Customer Name:", key=f"c_{f.name}")
                    with ci2: cn = st.text_input("CFA Name:", key=f"cf_{f.name}")
                    if st.button("✅ Generate Statement", key=f"b_{f.name}"): st.session_state[f"r_{f.name}"] = True
                    if st.session_state.get(f"r_{f.name}", False):
                        st.write("---")
                        cl1, cl2 = st.columns(2)
                        with cl1: st.download_button("📥 Excel Download", get_excel_download(data, h_info, s_info, d_info, p_inv, mn, cn), f"{f.name}.xlsx", key=f"xl_{f.name}")
                        with cl2: st.download_button("📥 PDF Download", get_pdf_download_fpdf(data, h_info, s_info, d_info, p_inv, mn, cn), f"{f.name}.pdf", key=f"pd_{f.name}")
            except Exception as e:
                st.error(f"❌ अरेरे! त्रुटी: {str(e)}")
                st.code(traceback.format_exc())
