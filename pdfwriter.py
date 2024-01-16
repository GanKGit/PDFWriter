from fpdf import FPDF
import pandas as pd
import glob
import datetime
import os

filepaths = glob.glob("Data/*.xlsx")
pdf = FPDF(orientation="P", unit="mm", format="A4")
for file in filepaths:
    print("Reading " + file + "...")
    df = pd.read_excel(file, sheet_name = "Sheet1")
    print("Adding page...")
    pdf.add_page()
    mtime=os.path.getmtime(file)
    dt_m=datetime.datetime.fromtimestamp(mtime)
    date=(dt_m.strftime("%m/%d/%Y"))
    pdf.set_font(family="Helvetica", style="B", size=14)
    Invoice=file.split("\\")
    pdf.cell(w=50, h=8,txt=f"Invoice : {Invoice[1]}", ln=1 )
    pdf.set_font(family="Courier", size=9)
    pdf.cell(w=50, h=6,txt=f"date : {date}", ln=2 )
    pdf.line(10,26,200,26)

    pdf.ln()
    pdf.set_font(family="Courier", style="B", size=11)
    pdf.cell(w=35, h=8, txt="Invoice No.", border=1)
    pdf.cell(w=25, h=8, txt="Date", border=1)
    pdf.cell(w=40, h=8, txt="Description", border=1)
    pdf.cell(w=30, h=8, txt="Quantity", border=1)
    pdf.cell(w=30, h=8, txt="Unit Price", border=1)
    pdf.cell(w=30, h=8, txt="Total Amount", border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Courier", size=10)
        pdf.cell(w=35,h=8, txt=str(row["Invoice Number"]),border=1)
        just_date=str(row["Date"]).split(" ")[0]
        pdf.cell(w=25,h=8, txt=just_date,border=1)
        pdf.cell(w=40, h=8, txt=str(row["Description"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["Quantity"]), border=1)
        pdf.cell(w=30,h=8, txt=str(row["Unit Price"]),border=1)
        pdf.cell(w=30,h=8, txt=str(row["Total Amount"]),border=1, ln=1)

pdf.output("C:\GAN\Python Projects\PDFWriter\output.pdf")