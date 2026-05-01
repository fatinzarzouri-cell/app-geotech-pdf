import streamlit as st
import pdfplumber
from openpyxl import Workbook
from io import BytesIO

st.title("Extraction tableaux PDF géotechnique")

def detect_nom_tableau(tableau):
    texte = " ".join(
        str(cell).lower()
        for ligne in tableau[:3]
        for cell in ligne
        if cell
    )

    if ("x" in texte and "y" in texte) or "coord" in texte:
        return "Coordonnees"
    elif "litho" in texte or "nature" in texte:
        return "Lithologie"
    elif "profondeur" in texte or "depth" in texte:
        return "Couches"
    elif "pressio" in texte or "pl" in texte or "em" in texte:
        return "Pressiometrique"
    elif "granulo" in texte or "tamis" in texte:
        return "Granulometrie"
    elif "atterberg" in texte or "wl" in texte or "ip" in texte:
        return "Atterberg"
    else:
        return "Tableau"

def nom_unique(nom_base, compteurs):
    if nom_base not in compteurs:
        compteurs[nom_base] = 1
    else:
        compteurs[nom_base] += 1

    return f"{nom_base}_{compteurs[nom_base]}"

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

if pdf_file:
    wb = Workbook()
    wb.remove(wb.active)

    compteurs = {}
    total = 0

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tableaux = page.extract_tables()

            for tableau in tableaux:
                if tableau and len(tableau) > 1:
                    nom_base = detect_nom_tableau(tableau)
                    nom_feuille = nom_unique(nom_base, compteurs)

                    ws = wb.create_sheet(title=nom_feuille)

                    ws.cell(row=1, column=1).value = f"Page PDF : {page_num}"
                    ws.cell(row=2, column=1).value = f"Type tableau : {nom_base}"

                    for i, ligne in enumerate(tableau, start=4):
                        for j, valeur in enumerate(ligne, start=1):
                            ws.cell(row=i, column=j).value = valeur

                    total += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"{total} tableaux extraits et organisés")

    st.download_button(
        label="Télécharger Excel",
        data=output,
        file_name="tableaux_geotech_organises.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
