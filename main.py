from pathlib import Path
import pandas as pd
from glob import glob
from fpdf import  FPDF

invoice_list = glob("invoices/*.xlsx")

for invoice in invoice_list:
    pdf = FPDF(orientation="P", format="A4", unit="mm")
    pdf.add_page()
    filename, date = Path(invoice).stem.split("-")
    pdf.set_font("Times", size=12, style="B")
    pdf.cell(50, 8, txt="Invoice nr." + filename, ln=1)

    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}" ,ln=1)

    # Add header to the table
    df = pd.read_excel(invoice, sheet_name="Sheet 1")
    columns = [ item.replace("_", " ").title()  for item in df.columns.tolist()]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for _, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80 , 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=40, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=str(sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=0, ln=1)
    pdf.cell(w=30, h=8, txt=f"The total price is {sum}", ln=1)


    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt="PythonOrg")
    pdf.image("images/download.jpeg", w=10, h=10, link="https://google.com")

    pdf.output("./PDFs/PDF_" + filename + ".pdf")