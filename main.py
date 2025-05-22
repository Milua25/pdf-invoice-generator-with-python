from pathlib import Path
import pandas as pd
from glob import glob
from fpdf import  FPDF

invoice_list = glob("invoices/*.xlsx")

for invoice in invoice_list:
    df = pd.read_excel(invoice, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()
    filename, date = Path(invoice).stem.split("-")
    pdf.set_font("Times", size=12, style="B")
    pdf.cell(50, 8, txt="Invoice nr." + filename, ln=1)

    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}")


    pdf.output("./PDFs/PDF_" + filename + ".pdf")