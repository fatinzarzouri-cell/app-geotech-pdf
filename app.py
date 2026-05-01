import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font
from io import BytesIO
import re

st.title("Extraction tableaux PDF géotechnique")

def detect_type(tableau):
    texte = " ".join(str(cell).lower() for ligne in tableau[:3] for cell in ligne if cell)

    if "coord" in texte or ("x" in texte and "y" in texte):
        return "Coordonnees"
    elif "litho" in texte:
        return "Lithologie"
    elif "profondeur" in texte:
        return "Couches"
    elif "pressio" in texte or "pl" in texte or "em" in texte:
        return "Pressiometrique"
    else:
        return "Autre"

def extraire_xy(texte):
    if not texte:
        return "", ""

    texte = str(texte).replace("\n", " ")

    x = re.search(r"X\s*[:=]\s*([0-9\s]+[,.]?[0-9]*)", texte, re.IGNORECASE)
    y = re.search(r"Y\s*[:=]\s*([0-9\s]+[,.]?[0-9]*)", texte, re.IGNORECASE)

    x_val = x.group(1).strip() if x else ""
    y_val = y.group(1).strip() if y else ""

    return x_val, y_val

def traiter_coordonnees(tableau):
    lignes = []
    lignes.append(["Sondage", "X", "Y"])

    for ligne in tableau[1:]:
        if len(ligne) >= 2:
            sondage = ligne[0]
            coord = ligne[1]

            x, y = extraire_xy(coord)

            if sondage and (x or y):
                lignes.append([sondage, x, y])

    return lignes

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

if pdf_file:
    wb = Workbook()
    wb.remove(wb.active)

    data = {}

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tableaux = page.extract_tables()

            for tableau in tableaux:
                if tableau and len(tableau) > 1:
                    typ = detect_type(tableau)

                    if typ == "Coordonnees":
                        lignes = traiter_coordonnees(tableau)
                    else:
                        lignes = tableau

                    if typ not in data:
                        data[typ] = []

                    if not data[typ]:
                        data[typ].extend(lignes)
                    else:
                        data[typ].extend(lignes[1:])

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for typ, lignes in data.items():
        ws = wb.create_sheet(title=typ)

        for i, ligne in enumerate(lignes, start=1):
            for j, val in enumerate(ligne, start=1):
                cell = ws.cell(row=i, column=j)
                cell.value = val
                cell.border = thin

                if i == 1:
                    cell.font = Font(bold=True)

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_len + 3

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("Extraction terminée avec X et Y séparés")

    st.download_button(
        label="Télécharger Excel",
        data=output,
        file_name="tables_geotech_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
