import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from io import BytesIO

st.title("Extraction simple des tableaux PDF vers Excel")

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

def clean(v):
    if v is None:
        return ""
    return str(v).replace("\n", " ").strip()

def style(ws):
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    header_fill = PatternFill("solid", fgColor="9DC3E6")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)

if pdf_file:
    wb = Workbook()
    wb.remove(wb.active)

    compteur = 1

    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 5,
        "text_tolerance": 2,
    }

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):

            tables = page.extract_tables(table_settings)

            for table in tables:
                if table and len(table) > 1:
                    ws = wb.create_sheet(f"Tableau_{compteur}")

                    for i, row in enumerate(table, start=1):
                        for j, val in enumerate(row, start=1):
                            ws.cell(row=i, column=j).value = clean(val)

                    style(ws)
                    compteur += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"{compteur - 1} tableaux extraits.")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="tableaux_copie_colle.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
