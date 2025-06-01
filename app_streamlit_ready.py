
import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
import num2words
from io import BytesIO
import zipfile

st.set_page_config(page_title="NGO Donation Receipt Generator", layout="wide")

class ReceiptPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.set_text_color(200, 0, 0)
        self.cell(0, 8, "Jeev Sewa Foundation", ln=True, align='C')
        self.set_font("Arial", "", 9)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 5,
            "Reg Office:- 1/4230, Gali No.8, Ram Nagar Extension,\nShahdara North East Delhi-110032",
            align='C'
        )
        self.ln(1)
        self.set_font("Arial", "", 8.5)
        self.cell(0, 5, "Reg. No.: 505/2019-20/4-909 | PAN: AADTJ3477H | 80G No: AADTJ3477H24DL02", ln=True, align='C')
        self.ln(2)

def clean_name(name):
    prefixes = ["Dr.", "Mr.", "Mrs.", "Ms.", "Prof.", "Miss", "Sir", "Madam"]
    name = name.strip()
    for prefix in prefixes:
        if name.lower().startswith(prefix.lower()):
            name = name[len(prefix):].strip()
            break
    return name

def generate_pdf(row):
    particulars_raw = str(row['Particulars']).strip()
    donor_name_clean = particulars_raw.split('(')[0].strip()
    donor_name_clean = clean_name(donor_name_clean)

    voucher_no = str(row['Voucher No.']).replace("/", "-").replace("\\", "-")
    t_no = str(row['Voucher No.'])
    ref_no = str(row['Narration']).strip()
    donor_name = donor_name_clean.replace(" ", "_").replace("/", "-").replace("\\", "-")
    serial_no = f"{voucher_no}_{donor_name}"
    donation_date = pd.to_datetime(row['Date']).strftime("%d-%m-%Y")

    amount = int(row['Donation'])
    amount_words = num2words.num2words(amount, to='cardinal', lang='en').title() + " Rupees Only"

    buffer = BytesIO()
    pdf = ReceiptPDF(orientation='P', unit='mm', format='A4')
    pdf.set_margins(left=10, top=5, right=10)
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    pdf.set_font("Arial", "", 10)

    pdf.set_font("Arial", "B", 13)
    pdf.cell(0, 8, "Tax Exempt Receipt", ln=True, align='C')
    pdf.ln(2)

    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Receipt No.  {t_no}", ln=False)
    pdf.cell(0, 6, f"Date: {donation_date}", ln=True, align='R')
    pdf.ln(2)

    pdf.set_font("Arial", "", 10)
    thank_you_lines = [
        f"Received with thanks from {donor_name_clean},",
        f"We have received your donation of Rs. {amount:,.2f}. Thank you for your generosity.",
        "The details of the donation are mentioned below:"
    ]
    for line in thank_you_lines:
        pdf.cell(0, 6, line, ln=True, align='C')
    pdf.ln(3)

    def add_detail(label, value):
        pdf.set_font("Arial", "B", 10)
        pdf.cell(48, 6, f"{label}", ln=False)
        pdf.set_font("Arial", "", 10)
        pdf.multi_cell(0, 6, f"{value}")

    add_detail("Donor Name:", donor_name_clean)
    add_detail("PAN No:", row["PAN No."])
    add_detail("Address:", row["Consignee/Party Address"])
    add_detail("Transaction ID:", ref_no)
    add_detail("Amount in words:", amount_words)

    pdf.ln(3)
    pdf.set_font("Arial", "", 9.5)
    pdf.multi_cell(0, 6,
        "Donations towards Jeev Sewa Foundation, registered under Section 80G of India's Income Tax Act, 1961, are tax-deductible."
    )

    pdf.ln(2)
    pdf.set_font("Arial", "I", 7.5)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(0, 3.8,
        "*This is a computer-generated receipt and does not require a signature.\n"
        "*This e-receipt is invalid in case of non-realization of payment instrument, reversal of credit card charge and/or reversal of amount for any reason.\n"
        "*No goods or services were provided to the donor by the organization in return for the contribution."
    )

    pdf.output(buffer)
    buffer.seek(0)
    return buffer, f"{serial_no}.pdf"

st.title("🧾 NGO Donation Receipt Generator")

uploaded_file = st.file_uploader("📤 Upload Excel file (.xls or .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Donation", header=10, engine='openpyxl')

        required_columns = ['Date', 'Particulars', 'Consignee/Party Address', 'Voucher Type',
                            'Voucher No.', 'PAN No.', 'Narration', 'Gross Total', 'Donation']
        if not set(required_columns).issubset(df.columns):
            st.error(f"❌ Uploaded sheet must contain the following columns:\n{required_columns}")
        else:
            df = df.reset_index(drop=True)
            st.success(f"✅ Loaded {len(df)} entries.")

            search_term = st.text_input("Filter by Consignee/Party Address (optional)").strip().lower()
            if search_term:
                df = df[df['Consignee/Party Address'].str.lower().str.contains(search_term)]
                st.info(f"Found {len(df)} matching entries.")

            if df.empty:
                st.warning("⚠️ No data to display.")
            else:
                st.dataframe(df[['Date', 'Consignee/Party Address', 'Donation', 'PAN No.', 'Voucher No.']])

                st.subheader("📌 Select Range to Generate")
                max_idx = len(df) - 1
                start_idx = st.number_input("Start index", 0, max_idx, 0)
                end_idx = st.number_input("End index", start_idx, max_idx, start_idx)

                if st.button("🖨️ Generate and Download Receipts (ZIP)"):
                    to_generate = df.iloc[start_idx:end_idx + 1]
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for i, row in to_generate.iterrows():
                            buffer, filename = generate_pdf(row)
                            zip_file.writestr(filename, buffer.getvalue())
                    zip_buffer.seek(0)
                    st.success(f"✅ Generated {len(to_generate)} receipts.")
                    st.download_button(
                        label="📥 Download All Receipts (ZIP)",
                        data=zip_buffer,
                        file_name="receipts.zip",
                        mime="application/zip"
                    )
    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
