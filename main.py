import os
import pandas as pd
from glob import glob
from fpdf import  FPDF

invoice_list = glob("invoices/*.xlsx")

for invoice in invoice_list:
    df = pd.read_excel(invoice, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()
    filename = (os.path.basename(invoice)).split("-")[0]
    print(filename)
    pdf.set_font("Times", size=12, style="B")
    pdf.cell(200, 10, txt="Invoice nr." + filename, align="C")
    pdf.output("./PDFs/PDF_" + filename + ".pdf")