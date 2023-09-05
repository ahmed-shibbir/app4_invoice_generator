import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
# print(filepaths)
# pdf = FPDF(orientation="P", unit="mm", format="A4")
for filepath in filepaths:

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    print(filename)

    # invoice_nr = filename.split("-")[0]
    invoice_nr, date = filename.split("-")
    print(invoice_nr)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(h=10, w=0, txt=f"Invoice_no: {invoice_nr}",  ln=1)

    # date = filename.split("-")[1]
    # pdf.line()
    pdf.cell(h=10, w=0, txt=f"Date: {date}", ln=3)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    cols = list(df.columns)
    cols = [col.replace("_", " ").title() for col in cols]
    print(cols)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=30, h=8, txt=cols[0], border=1)
    pdf.cell(w=70, h=8, txt=cols[1], border=1)
    pdf.cell(w=50, h=8, txt=cols[2],border=1)
    pdf.cell(w=50, h=8, txt=cols[3],border=1)
    pdf.cell(w=50, h=8, txt=cols[4],border=1,ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="", size=14)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=50, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=50, h=8, txt=str(row["total_price"]),border=1,ln=1)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=30, h=8, txt="Total", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=8, txt="The total due amount is:", ln=1)





    # pdf.text(x=10, y=10, txt="The total due amount is: ")







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
