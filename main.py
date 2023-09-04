import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
# print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # print(df)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    print(filename)
    invoice_nr = filename.split("-")[0]
    print(invoice_nr)
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(h=10, w=0, txt=f"Invoice_no: {invoice_nr}", border=1)
    date = filename.split("-")[1]
    # pdf.line()
    pdf.cell(h=10, w=0, txt=f"Date: {date}", border=1)
    pdf.output(f"pdf/{filename}.pdf")




# # df = pd.read_excel("invoices/10001-2023.1.18.xlsx")
#
# print(df)
#
# pdf = FPDF(orientation="P", format="A4")
#
# pdf.add_page()
# # Set the header
# pdf.set_font(family="Times", style="B", size=24) # this line is a must before adding cell type.Without it, 'unifontsubset' error shows up
# # pdf.set_text_color(100, 100, 100)
# pdf.cell(w=0, h=12, txt="Topic", align="L", ln=1)
#

#
#
#
#
# for item in df.iterrows():
#     print(item)
#     for i in range(len(item)):
#         print(item[1][i])
