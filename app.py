import streamlit as st
import pdfplumber
from openpyxl import Workbook
from io import BytesIO

st.title("Extraction tableaux PDF géotechnique")

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

if pdf_file:
    wb = Workbook()
    wb.remove(wb.active)

    compteur = 1

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tableaux = page.extract_tables()

            for tableau in tableaux:
                if tableau and len(tableau) > 1:
                    ws = wb.create_sheet(title=f"Tableau_{compteur}")
                    ws.cell(row=1, column=1).value = f"Page PDF : {page_num}"

                    for i, ligne in enumerate(tableau, start=3):
                        for j, valeur in enumerate(ligne, start=1):
                            ws.cell(row=i, column=j).value = valeur

                    compteur += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"{compteur - 1} tableaux extraits")

    st.download_button(
        label="Télécharger Excel",
        data=output,
        file_name="tableaux_extraits.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
