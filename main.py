import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Extracting file paths in a list from invoices folder
filepaths = glob.glob("invoices/*.xlsx")
# print(filepaths)
# pdf = FPDF(orientation="P", unit="mm", format="A4")

# Creating separate pdf files for each invoice
for filepath in filepaths:

    # Creating a pdf instance and adding page
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    # Extracting filename portion from the file path
    filename = Path(filepath).stem
    print(filename)

    # Unpacking data/variables from the filename
    # invoice_nr = filename.split("-")[0]
    invoice_nr, date = filename.split("-")
    print(invoice_nr)

    # Adding a pdf cell for showing invoice number
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(h=10, w=0, txt=f"Invoice_no: {invoice_nr}",  ln=1)

    # date = filename.split("-")[1]

    # Adding a pdf cell for showing the date
    pdf.cell(h=10, w=0, txt=f"Date: {date}", ln=3)

    # Reading data from Excel file to create a dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)

    # Replace underscores from column header with first letters capitalised
    cols = list(df.columns)
    cols = [col.replace("_", " ").title() for col in cols]
    print(cols)

    # Add column title
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=30, h=8, txt=cols[0], border=1)
    pdf.cell(w=70, h=8, txt=cols[1], border=1)
    pdf.cell(w=50, h=8, txt=cols[2],border=1)
    pdf.cell(w=50, h=8, txt=cols[3],border=1)
    pdf.cell(w=50, h=8, txt=cols[4],border=1,ln=1)

    # Populate table with data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="", size=14)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=50, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=50, h=8, txt=str(row["total_price"]),border=1,ln=1)

    # Calculate total price to be shown in the table
    sum_price = sum(df["total_price"])

    # Add rows to the table
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=30, h=8, txt="Total", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt=f"{sum_price}", border=1, ln=1)  # str(sum_price)

    # Add a sentence showing total price.
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=8, txt=f"The total due amount is: {sum_price}", ln=1)

    # Add log and image
    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=47, h=8, txt=f"Innovative Co.")
    pdf.image("image.jpg", w=15)

    pdf.output(f"pdf/{filename}.pdf")

