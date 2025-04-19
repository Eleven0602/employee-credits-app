import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

# Load Excel file
df = pd.read_excel("employees.xlsx")

# Loop through each employee
for index, row in df.iterrows():
    # Load Word template
    doc = DocxTemplate("template.docx")

    # Fill in the placeholders
    context = {
        "employeeName": row["employeeName"],
        "Credits": row["Credits"]
    }
    doc.render(context)

    # Create output folder if it doesn't exist
    os.makedirs("output", exist_ok=True)

    # Save as new Word file
    doc_path = f"output/{row['employeeName']}.docx"
    doc.save(doc_path)

    # Convert to PDF
    convert(doc_path)  # This will save PDF in same folder
    print(f"Generated PDF for {row['employeeName']}")
