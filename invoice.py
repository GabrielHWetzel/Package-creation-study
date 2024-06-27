import os
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


def generate(invoices_path, pdfs_path, image_path, product_id, product_name,
             amount_purchased, price_per_unit, total_price):
    invoices = glob.glob(f"{invoices_path}/*.xlsx")

    for invoice in invoices:

        # Initiate PDF
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.add_page()

        # Filename as list.
        filename = Path(invoice).stem
        file_number, file_date = filename.split('-')

        # Layout
        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, txt=f"Invoice nr. {file_number}", ln=1)
        pdf.ln(h=2)
        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=50, h=8, txt=f"Date: {file_date}", ln=1)
        pdf.ln(h=5)
        # Excel file
        df = pd.read_excel(invoice, sheet_name="Sheet 1")

        # Headers
        columns = df.columns
        columns = [column.replace("_", " ").title() for column in columns]
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=30, h=8, txt=columns[0], border=1)
        pdf.cell(w=60, h=8, txt=columns[1], border=1)
        pdf.cell(w=35, h=8, txt=columns[2], border=1)
        pdf.cell(w=35, h=8, txt=columns[3], border=1)
        pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

        # Tables
        for index, row in df.iterrows():
            pdf.set_font(family="Times", size=10)
            pdf.cell(w=30, h=8, txt=str(row[product_id]), border=1)
            pdf.cell(w=60, h=8, txt=str(row[product_name]), border=1)
            pdf.cell(w=35, h=8, txt=str(row[amount_purchased]), border=1)
            pdf.cell(w=35, h=8, txt=str(row[price_per_unit]), border=1)
            pdf.cell(w=30, h=8, txt=str(row[total_price]), border=1, ln=1)

        total_price = df["total_price"].sum()
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, border=1)
        pdf.cell(w=60, h=8, border=1)
        pdf.cell(w=35, h=8, border=1)
        pdf.cell(w=35, h=8, border=1)
        pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)
        pdf.ln(h=5)

        # Bottom
        pdf.set_font(family="Times", size=12, style="B")
        pdf.cell(w=30, h=10, txt=f"The total price is {total_price}", ln=1)

        pdf.set_font(family="Times", size=16, style="B")
        pdf.cell(w=35, h=10, txt="PythonHow")
        pdf.image(image_path, w=10)

        os.makedirs(pdfs_path)
        pdf.output(f"{pdfs_path}/{filename}.pdf")
