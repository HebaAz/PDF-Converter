from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path
import time

filepaths = glob.glob(f"App4/Invoices/*xlsx")

for filepath in filepaths:
    dataframe = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit = "mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Arial", style="B", size=16)
    pdf.set_text_color(100, 100, 100)

    #invoice ID
    filename = Path(filepath).stem
    invoiceNumber = filename[:5]
    pdf.cell(w=50, h=8, txt = f"Invoice nr. {invoiceNumber}", ln=1)

    #date
    dateYMD = time.strftime("%Y.%m.%d")
    pdf.cell(w=50, h=8, txt = f"Date {dateYMD}", ln=1)

    pdf.ln(20)

    #Title cells
    pdf.set_font(family="Arial", style="B", size=10)
    pdf.cell(w=30, h=8, txt = "Product ID", border=1, ln=0)
    pdf.cell(w=50, h=8, txt = "Product Name", border=1, ln=0)
    pdf.cell(w=30, h=8, txt = "Amount", border=1, ln=0)
    pdf.cell(w=40, h=8, txt = "Price per Unit", border=1, ln=0)
    pdf.cell(w=30, h=8, txt = "Total Price", border=1, ln=1)

    total_due = 0

    pdf.set_font(family="Arial", style="", size=10)
    for index, row in dataframe.iterrows():
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1, ln=0)
        pdf.cell(w=50, h=8, txt=row['product_name'], border=1, ln=0)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1, ln=0)
        pdf.cell(w=40, h=8, txt=str(row['price_per_unit']), border=1, ln=0)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

        total_due = total_due + row["total_price"]

    pdf.ln(50)

    pdf.set_font(family="Arial", style="", size=15)
    pdf.cell(w=50, h=15, txt=f"The total due amount is {total_due} dollars", ln=1)
    pdf.cell(w=50, h=15, txt="Heba Azeef", ln=1)

    # Output the PDF file
    pdf.output(f"App4/PDFs/Invoice_{invoiceNumber}.pdf")