import os
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import _functions as cfx


def file_to_pdf(input_file, output_pdf):
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension in ['.xls', '.xlsx', '.xlsb']:
        df = pd.read_excel(input_file, sheet_name=None) # Read all sheets into a dictionary
    elif file_extension == '.csv':
        df = {"Sheet1": pd.read_csv(input_file)}

    # width_in_inches = 300
    # height_in_inches = 150
    # custom_size = (width_in_inches * 72, height_in_inches * 72)  # Convert inches to points
    pdf = SimpleDocTemplate(output_pdf, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    max_column_width = 500
    for sheet_name, data in df.items():
        if data.empty:
            elements.append(Paragraph(f"Sheet: {sheet_name} (Empty)", styles["Title"]))
            elements.append(Spacer(1, 12))
            continue

        elements.append(Paragraph(f"Sheet: {sheet_name}", styles["Title"]))
        elements.append(Spacer(1, 12))

        table_data = [data.columns.tolist()] + data.fillna("").values.tolist()
        col_widths = [min(max(len(str(item)) for item in data[column]) * 10, max_column_width) for column in data.columns]

        table = Table(table_data, colWidths=col_widths)
        style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 12),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ("BACKGROUND", (0, 1), (-1, -1), colors.white),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ])
        table.setStyle(style)
        elements.append(table)
        elements.append(Spacer(1, 24))
    pdf.build(elements)
    print(f"PDF saved as {output_pdf}")


input_file = cfx.ifile()
output_pdf = os.path.splitext(input_file)[0] + ".pdf"
file_to_pdf(input_file, output_pdf)