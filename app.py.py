import streamlit as st
import pdfplumber
import pandas as pd
import re
import datetime
import io

# वेब ॲपचे डिझाईन
st.set_page_config(page_title="Virbac Statement Converter", page_icon="📄", layout="centered")

st.title("📄 Virbac Account Statement Converter")
st.markdown("CFA टीमसाठी: PDF अपलोड करा आणि **Excel + PDF** दोन्ही फॉरमॅट मिळवा.")

uploaded_files = st.file_uploader("येथे PDF फाईल अपलोड करा", type="pdf", accept_multiple_files=True)

def process_pdf_logic(uploaded_file):
    run_datetime = datetime.datetime.now().strftime("%d/%m/%Y %I:%M %p")
    period, customer_no, customer_name = "", "", ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            first_page = pdf.pages[0]
            text_layout = first_page.extract_text(layout=True) or ""
            clean_text = " ".join(text_layout.split())
            
            # Period
            p_match = re.search(r'(?:Account.*?date|period from).*?(\d{2}/\d{2}/\d{2,4})\s*(?:to|-)\s*(\d{2}/\d{2}/\d{2,4})', clean_text, re.IGNORECASE)
            if p_match: period = f"from {p_match.group(1)} To {p_match.group(2)}"
            
            # Customer No 
            c_match = re.search(r'(?:Payer|Customer No).*?(\d{6})', clean_text, re.IGNORECASE)
            if c_match: 
                customer_no = c_match.group(1).strip()
                
                # --- Customer Name शोधण्याचे नवीन हमखास लॉजिक (खालची ओळ वाचणे) ---
                raw_text = first_page.extract_text() # नॉर्मल टेक्स्ट (Layout शिवाय)
                if raw_text:
                    # 'Customer No 606149' च्या बरोबर खालची ओळ (Next Line) उचलणे
                    pattern = r'Customer No\s*' + re.escape(customer_no) + r'\s*\n+([^\n]+)'
                    name_match = re.search(pattern, raw_text, re.IGNORECASE)
                    if name_match:
                        customer_name = name_match.group(1).strip()
                    else:
                        # जर काही कारणाने वरचे चालले नाही, तर दुसरी पद्धत
                        lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
                        for i, line in enumerate(lines):
                            if customer_no in line and i + 1 < len(lines):
                                customer_name = lines[i+1].strip()
                                break
                # -----------------------------------------------------------

            extracted_rows = []
            for page in pdf.pages:
                p_text = page.extract_text()
                if p_text:
                    for line in p_text.split('\n'): extracted_rows.append(line.strip())
    except Exception as e:
        return None, None, None, f"PDF वाचताना एरर: {e}"

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
    return None, None, None, "डेटा सापडला नाही."

def get_excel_download(final_data, header_info, summary_info):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame([
            ["Report Generation Time:", header_info["Time"]], 
            ["Heading:", "STATEMENT OF ACCOUNT"], 
            ["Period:", header_info["Period"]], 
            ["Customer No:", header_info["CustomerNo"]],
            ["Customer Name:", header_info["CustomerName"]]
        ]).to_excel(writer, sheet_name='Statement', index=False, header=False, startrow=0)
        pd.DataFrame(summary_info, columns=["TRANSACTION TYPE", "AMOUNT (INR)"]).to_excel(writer, sheet_name='Statement', index=False, startrow=7)
        pd.DataFrame(final_data).to_excel(writer, sheet_name='Statement', index=False, startrow=23)
    return output.getvalue()

def get_pdf_download_fpdf(final_data, header_info, summary_info):
    from fpdf import FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 6, txt="VIRBAC - STATEMENT OF ACCOUNT", ln=True, align='C')
    pdf.set_font("Arial", size=9)
    pdf.cell(190, 6, txt=f"Time: {header_info['Time']} | Period: {header_info['Period']}", ln=True, align='C')
    pdf.cell(190, 6, txt=f"Customer No: {header_info['CustomerNo']} | Customer Name: {header_info['CustomerName']}", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(100, 6, "Transaction Type", border=1, align='L')
    pdf.cell(40, 6, "Amount (INR)", border=1, ln=True, align='R')
    pdf.set_font("Arial", size=9)
    for row in summary_info:
        pdf.cell(100, 6, str(row[0]), border=1, align='L')
        pdf.cell(40, 6, f"{int(float(row[1])):,}", border=1, ln=True, align='R')
    pdf.ln(5)
    
    col_widths = [18, 55, 20, 18, 22, 22, 25]
    headers = ["Date", "Type", "Doc No", "Chq No", "Debit", "Credit", "Balance"]
    
    def print_headers():
        pdf.set_font("Arial", 'B', 8)
        for i in range(len(headers)):
            align = 'R' if headers[i] in ["Debit", "Credit", "Balance"] else 'C'
            pdf.cell(col_widths[i], 6, headers[i], border=1, align=align)
        pdf.ln()
        pdf.set_font("Arial", size=8)

    print_headers()
    
    for r in final_data:
        if pdf.get_y() > 275:
            pdf.add_page()
            print_headers()
            
        debit_val = f"{int(float(r['Debit'])):,}" if r['Debit'] != "" else ""
        credit_val = f"{int(float(r['Credit'])):,}" if r['Credit'] != "" else ""
        balance_val = f"{int(float(r['Balance'])):,}" if r['Balance'] != "" else ""

        pdf.cell(col_widths[0], 6, str(r['Date']), border=1, align='C')
        pdf.cell(col_widths[1], 6, str(r['Type'])[:30], border=1, align='L')
        pdf.cell(col_widths[2], 6, str(r['Doc No']), border=1, align='C')
        pdf.cell(col_widths[3], 6, str(r['Chq No']), border=1, align='C')
        pdf.cell(col_widths[4], 6, debit_val, border=1, align='R')
        pdf.cell(col_widths[5], 6, credit_val, border=1, align='R')
        pdf.cell(col_widths[6], 6, balance_val, border=1, align='R')
        pdf.ln()
        
    if pdf.get_y() > 260:
        pdf.add_page()
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(190, 6, "For Virbac Animal Health India Pvt Ltd", ln=True)
    pdf.ln(8)
    pdf.cell(190, 6, "Signature                  Place: ___________      Date: ___________", ln=True)
    
    return bytes(pdf.output(dest='S').encode('latin1'))

if uploaded_files:
    for file in uploaded_files:
        data, h_info, s_info, err = process_pdf_logic(file)
        if err: st.error(err)
        else:
            st.success(f"✅ {file.name} यशस्वीरीत्या कनवर्ट झाली!")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(f"📥 Excel डाऊनलोड करा", get_excel_download(data, h_info, s_info), f"{file.name}.xlsx")
            with col2:
                try:
                    import fpdf
                    st.download_button(f"📥 PDF डाऊनलोड करा", get_pdf_download_fpdf(data, h_info, s_info), f"{file.name}.pdf")
                except:
                    st.warning("PDF साठी आधी 'pip install fpdf' करा.")