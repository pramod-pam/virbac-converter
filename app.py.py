import streamlit as st
import pdfplumber
import pandas as pd
import re
import datetime
import io
import os

# Web app design
st.set_page_config(page_title="Virbac Statement Converter", page_icon="📄", layout="centered")

st.title("📄 Virbac Account Statement Converter")
st.markdown("CFA Team sathi: PDF upload kara ani **Excel + PDF** donhi format milva.")

uploaded_files = st.file_uploader("Yethe PDF file upload kara", type="pdf", accept_multiple_files=True)

def process_pdf_logic(uploaded_file):
    run_datetime = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
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
        return None, None, None, f"PDF vachtana error: {e}"

    extracted_rows = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            p_text = page.extract_text()
            if p_text:
                for line in p_text.split('\n'): extracted_rows.append(line.strip())

    final_data, running_balance, opening_balance, found_opening = [], 0.0, 0.0, False
    s_inv, pay, reco, c_oth, c_brk, g_ret, d_not, tcs, tds, tech_b, n_tech_b = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0

    for row_text in extracted_rows:
        row_upper = row_text.upper()
        if "OPENING BALANCE" in row_upper and not found_opening:
            amounts = re.findall(r'\(?[\d,]+\.\d{2}\)?', row_text)
            if amounts:
                amount_str = amounts[-1]
                is_cr = 'CR' in row_upper or '(' in amount_str
                val = float(re.sub(r'[^\d.]', '', amount_str))
                opening_balance = val
                running_balance = -val if is_cr else val
                found_opening = True
                final_data.append({"Date": "", "Type": "OPENING BAL", "Doc No": "", "Chq No": "", "Debit": val if not is_cr else "", "Credit": val if is_cr else "", "Balance": round(running_balance, 2)})
            continue

        date_match = re.search(r'\b(\d{2}/\d{2}/\d{2})\b', row_text)
        if date_match:
            if any(x in row_upper for x in ["DATE:", "ACCOUNTING DATE", "TIME:", "PAGE"]): continue
            date, amounts = date_match.group(1), re.findall(r'\(?[\d,]+\.\d{2}\)?', row_text)
            if not amounts: continue
            amount_str = amounts[-1]
            is_cr, val = ('CR' in row_upper or '(' in amount_str), float(re.sub(r'[^\d.]', '', amounts[-1]))
            if val == 0.0: continue
            doc_no = (re.search(r'\b\d{9}\b', row_text)).group(0) if re.search(r'\b\d{9}\b', row_text) else ""
            
            t_type, s_type = "Other", "Other"
            if "TDSRECO" in row_upper or "TDS CREDIT NOTE" in row_upper or (doc_no.startswith('000') and is_cr): t_type, s_type = "TDS Credit Note", "TDS Credit Note"
            elif "CBOU199" in row_upper: t_type, s_type = "TECHNICAL BOUNCED", "TECHNICAL BOUNCED"
            elif any(code in row_upper for code in ["CBOU101", "CBOU102", "CBOU110"]): t_type, s_type = "NON TECHNICAL BOUNCED", "NON TECHNICAL BOUNCED"
            elif any(x in row_upper for x in ["CHQ", "PAYMENT", "DD-NEFT", "NEFT"]): t_type, s_type = "PAYMENT", "PAYMENT"
            elif "RECONC" in row_upper: t_type, s_type = "RECONCILIATION", "RECONCILIATION"
            elif "INVOICE" in row_upper:
                if doc_no.startswith('3'): t_type, s_type = "Credit Note(Brakage Expiry)", "Credit Note(Brakage Expiry)"
                elif doc_no.startswith('4'): t_type, s_type = "Credit Note(Others)", "Credit Note(Others)"
                elif doc_no.startswith('5'): t_type, s_type = "Debit Note", "Debit Note"
                elif doc_no.startswith('8'): t_type, s_type = "TCS Debit Note", "TCS Debit Note"
                else: t_type, s_type = ("Goods Return Invoice", "Goods Return Invoice") if is_cr else ("Sales Invoice", "Sales Invoice")

            chq_no = (re.search(r'\b\d{6}\b', row_text).group(0) if re.search(r'\b\d{6}\b', row_text) and re.search(r'\b\d{6}\b', row_text).group(0) != doc_no else "") if s_type == "PAYMENT" else ""
            debit, credit = (val, 0.0) if not is_cr else (0.0, val)
            running_balance += (debit - credit)

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
            final_data.append({"Date": date, "Type": t_type, "Doc No": doc_no, "Chq No": chq_no, "Debit": debit if debit > 0 else "", "Credit": credit if credit > 0 else "", "Balance": round(running_balance, 2)})

    if final_data:
        final_data.append({"Date": "", "Type": "CLOSING BAL", "Doc No": "", "Chq No": "", "Debit": "", "Credit": "", "Balance": round(running_balance, 2)})
        header_info = {"Time": run_datetime, "Period": period, "CustomerNo": customer_no, "CustomerName": customer_name}
        summary_info = [
            ("OPENING BAL", opening_balance), ("Sales Invoice", s_inv), ("PAYMENT", pay), ("RECONCILIATION", reco),
            ("Credit Note(Others)", c_oth), ("Credit Note(Brakage Expiry)", c_brk), ("Goods Return Invoice", g_ret),
            ("Debit Note", d_not), ("TCS Debit Note", tcs), ("TDS Credit Note", tds),
            ("TECHNICAL BOUNCED", tech_b), ("NON TECHNICAL BOUNCED", n_tech_b), ("CLOSING BAL", running_balance)
        ]
        return final_data, header_info, summary_info, None
    return None, None, None, "Data sapadla nahi."

# --- PDF Function with Manual Name & CFA Name ---
def get_pdf_download_fpdf(final_data, header_info, summary_info, manual_name="", cfa_name=""):
    from fpdf import FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    
    logo_file = next((f for f in ["logo.png", "Logo.png", "logo.jpg"] if os.path.exists(f)), None)
    if logo_file:
        pdf.image(logo_file, x=85, y=5, w=40)
        pdf.ln(15)
    else: pdf.ln(5)
    
    cust_name = manual_name if manual_name.strip() else header_info['CustomerName']
    
    # CFA Name chi line tayar karne
    cfa_text = f" | CFA Name: {cfa_name}" if cfa_name.strip() else ""

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 6, txt="VIRBAC - STATEMENT OF ACCOUNT", ln=True, align='C')
    pdf.set_font("Arial", size=9)
    pdf.cell(190, 6, txt=f"Time: {header_info['Time']} | Period: {header_info['Period']}", ln=True, align='C')
    
    if cfa_name.strip():
        pdf.cell(190, 6, txt=f"Customer No: {header_info['CustomerNo']} | Customer Name: {cust_name}{cfa_text}", ln=True, align='C')
    else:
        pdf.cell(190, 6, txt=f"Customer No: {header_info['CustomerNo']} | Customer Name: {cust_name}", ln=True, align='C')
        
    pdf.ln(5)
    
    # Summary Table
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(100, 6, "Transaction Type", border=1, align='L')
    pdf.cell(40, 6, "Amount (INR)", border=1, ln=True, align='R')
    pdf.set_font("Arial", size=9)
    for row in summary_info:
        pdf.cell(100, 6, str(row[0]), border=1, align='L')
        pdf.cell(40, 6, f"{int(float(row[1])):,}", border=1, ln=True, align='R')
    pdf.ln(5)
    
    # Main Table
    col_widths = [18, 55, 20, 18, 22, 22, 25]
    headers = ["Date", "Type", "Doc No", "Chq No", "Debit", "Credit", "Balance"]
    pdf.set_font("Arial", 'B', 8)
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 6, headers[i], border=1, align='C')
    pdf.ln()
    pdf.set_font("Arial", size=8)

    for r in final_data:
        if pdf.get_y() > 275: pdf.add_page()
        pdf.cell(col_widths[0], 6, str(r['Date']), border=1, align='C')
        pdf.cell(col_widths[1], 6, str(r['Type'])[:30], border=1, align='L')
        pdf.cell(col_widths[2], 6, str(r['Doc No']), border=1, align='C')
        pdf.cell(col_widths[3], 6, str(r['Chq No']), border=1, align='C')
        pdf.cell(col_widths[4], 6, f"{int(float(r['Debit'])):,}" if r['Debit']!="" else "", border=1, align='R')
        pdf.cell(col_widths[5], 6, f"{int(float(r['Credit'])):,}" if r['Credit']!="" else "", border=1, align='R')
        pdf.cell(col_widths[6], 6, f"{int(float(r['Balance'])):,}" if r['Balance']!="" else "", border=1, align='R')
        pdf.ln()
        
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(190, 6, "For Virbac Animal Health India Pvt Ltd", ln=True)
    pdf.ln(5)
    if cfa_name.strip():
        pdf.cell(190, 6, f"Authorized Signatory: {cfa_name}", ln=True)
    pdf.ln(5)
    pdf.cell(190, 6, "Signature                  Place: ___________      Date: ___________", ln=True)
    
    return bytes(pdf.output(dest='S').encode('latin1'))

# Excel logic
def get_excel_download(final_data, header_info, summary_info, manual_name="", cfa_name=""):
    output = io.BytesIO()
    cust_name = manual_name if manual_name.strip() else header_info['CustomerName']
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([
            ["Report Time:", header_info["Time"]], 
            ["Heading:", "STATEMENT OF ACCOUNT"], 
            ["Period:", header_info["Period"]], 
            ["Customer No:", header_info["CustomerNo"]],
            ["Customer Name:", cust_name],
            ["CFA Name:", cfa_name]
        ]).to_excel(writer, sheet_name='Statement', index=False, header=False)
        pd.DataFrame(summary_info, columns=["TYPE", "AMOUNT"]).to_excel(writer, sheet_name='Statement', index=False, startrow=8)
        pd.DataFrame(final_data).to_excel(writer, sheet_name='Statement', index=False, startrow=24)
    return output.getvalue()

if uploaded_files:
    for file in uploaded_files:
        data, h_info, s_info, err = process_pdf_logic(file)
        if err: st.error(err)
        else:
            st.success(f"✅ {file.name} वाचून तयार आहे!")
            
            st.info(f"PDF मधून आलेले कस्टमरचे नाव: **{h_info['CustomerName']}**")
            col_in1, col_in2 = st.columns(2)
            with col_in1:
                manual_name = st.text_input("Customer Name (जर नाव चुकीचे असेल तर इथे बदला):", key=f"cust_{file.name}")
            with col_in2:
                cfa_name = st.text_input("CFA Name (इथे CFA चे नाव टाका):", key=f"cfa_{file.name}")
            
            # --- नवीन ऑटोमॅटिक सिस्टीम ---
            # हे बटण दाबल्यावर माहिती आपसुक सेव्ह होईल आणि डाऊनलोडची बटणे दिसतील.
            if st.button("✅ फाईल तयार करा (Prepare Files)", key=f"btn_{file.name}"):
                st.session_state[f"ready_{file.name}"] = True
                
            if st.session_state.get(f"ready_{file.name}", False):
                st.write("---")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("📥 Excel डाऊनलोड करा", get_excel_download(data, h_info, s_info, manual_name, cfa_name), f"{file.name}.xlsx", key=f"dl_xl_{file.name}")
                with col2:
                    st.download_button("📥 PDF डाऊनलोड करा", get_pdf_download_fpdf(data, h_info, s_info, manual_name, cfa_name), f"{file.name}.pdf", key=f"dl_pdf_{file.name}")
