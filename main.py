import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", 
             align="L", new_x="LMARGIN", new_y="NEXT")
    
    pdf.cell(w=50, h=8, txt=f"Date: {date}", 
             align="L", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # add table header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1,
              new_x="LMARGIN", new_y="NEXT")
    
    # Add table rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1,
                  new_x="LMARGIN", new_y="NEXT")

    # Add total price
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1,
              new_x="LMARGIN", new_y="NEXT")

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total prices is {total_sum}", 
             new_x="LMARGIN", new_y="NEXT")
    
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=26, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)
    
    pdf.output(f"pdfs/{filename}.pdf")