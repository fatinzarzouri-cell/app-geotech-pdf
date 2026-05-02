import streamlit as st
import camelot
import pandas as pd
from io import BytesIO
import tempfile
import os
import re

st.title("Extraction tableaux PDF géotechnique")

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

def clean_cell(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\n", " ").strip()

def clean_df(df):
    df = df.applymap(clean_cell)
    df = df.dropna(how="all")
    df = df.loc[:, ~(df == "").all()]
    return df.reset_index(drop=True)

def get_header(df):
    # cherche la première ligne qui ressemble à un header
    for i in range(min(4, len(df))):
        txt = " ".join(df.iloc[i].astype(str)).lower()
        if any(w in txt for w in ["sondage", "profondeur", "pression", "module", "résistance", "resistance", "coord"]):
            return i
    return 0

def normalize_header(header):
    txt = " ".join([str(x).lower().strip() for x in header])
    txt = re.sub(r"\s+", " ", txt)
    return txt

def safe_sheet_name(name):
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
    return name[:31]

def detect_name(header_text):
    h = header_text.lower()

    if "pression" in h or "pressio" in h or "module pressiom" in h:
        return "Pressiometrique"
    if "compression" in h or "résistance" in h or "resistance" in h:
        return "Compression_simple"
    if "coord" in h or (" x " in h and " y " in h):
        return "Coordonnees"
    if "lithologie" in h or "formation" in h:
        return "Lithologie"

    return "Tableau"

if pdf_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name

    st.info("Extraction en cours...")

    tables = camelot.read_pdf(
        pdf_path,
        pages="all",
        flavor="lattice"
    )

    groupes = {}

    for table in tables:
        df = clean_df(table.df)

        if df.empty or len(df) < 2:
            continue

        header_index = get_header(df)
        header = list(df.iloc[header_index])
        data = df.iloc[header_index + 1:].copy()

        # supprimer les répétitions du header dans les pages suivantes
        header_norm = normalize_header(header)
        rows_clean = []

        for _, row in data.iterrows():
            row_list = list(row)
            row_norm = normalize_header(row_list)

            if row_norm == header_norm:
                continue

            if all(str(x).strip() == "" for x in row_list):
                continue

            rows_clean.append(row_list)

        if not rows_clean:
            continue

        key = f"{len(header)}__{header_norm}"
        nom_base = detect_name(header_norm)

        if key not in groupes:
            groupes[key] = {
                "nom": nom_base,
                "header": header,
                "rows": []
            }

        groupes[key]["rows"]..extend(rows_clean)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used_names = {}

        for key, g in groupes.items():
            nom = g["nom"]

            if nom in used_names:
                used_names[nom] += 1
                nom_sheet = f"{nom}_{used_names[nom]}"
            else:
                used_names[nom] = 1
                nom_sheet = nom

            nom_sheet = safe_sheet_name(nom_sheet)

            final_df = pd.DataFrame(g["rows"], columns=g["header"])
            final_df.to_excel(writer, sheet_name=nom_sheet, index=False)

            ws = writer.book[nom_sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)

    output.seek(0)

    st.success(f"{len(groupes)} tableaux homogènes fusionnés.")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="tableaux_fusionnes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    os.remove(pdf_path)
