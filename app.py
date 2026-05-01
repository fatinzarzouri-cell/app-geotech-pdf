import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from io import BytesIO

st.title("Extraction tableaux PDF géotechnique")

def detect_type(tableau):
    texte = " ".join(
        str(cell).lower()
        for ligne in tableau[:3]
        for cell in ligne if cell
    )

    if ("x" in texte and "y" in texte) or "coord" in texte:
        return "Coordonnees"
    elif "litho" in texte:
        return "Lithologie"
    elif "profondeur" in texte:
        return "Couches"
    elif "pressio" in texte or "pl" in texte:
        return "Pressiometrique"
    else:
        return "Autre"

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

if pdf_file:
    wb = Workbook()
    wb.remove(wb.active)

    data = {}  # stockage par type

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for t in tables:
                if t and len(t) > 1:
                    typ = detect_type(t)

                    if typ not in data:
                        data[typ] = []

                    data[typ].extend(t)

    # écrire dans Excel
    thin = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for typ, lignes in data.items():
        ws = wb.create_sheet(title=typ)

        for i, ligne in enumerate(lignes, start=1):
            for j, val in enumerate(ligne, start=1):
                cell = ws.cell(row=i, column=j)
                cell.value = val
                cell.border = thin

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("Tableaux regroupés et organisés")

    st.download_button(
        label="Télécharger Excel",
        data=output,
        file_name="tables_geotech_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
