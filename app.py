import streamlit as st
import camelot
import pandas as pd
from io import BytesIO
import tempfile
import os
import re

st.title("Extraction intelligente des tableaux PDF")

# تنظيف cellule
def clean_cell(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\n", " ").strip()

# تنظيف dataframe
def clean_df(df):
    df = df.map(clean_cell)
    df = df.dropna(how="all")
    df = df.loc[:, ~(df == "").all()]
    return df.reset_index(drop=True)

# تحديد header
def get_header(df):
    for i in range(min(4, len(df))):
        txt = " ".join(df.iloc[i].astype(str)).lower()
        if any(w in txt for w in [
            "sondage", "profondeur", "pression",
            "module", "résistance", "resistance", "coord"
        ]):
            return i
    return 0

# normalize header
def normalize_header(header):
    txt = " ".join([str(x).lower().strip() for x in header])
    txt = re.sub(r"\s+", " ", txt)
    return txt

# اسم feuille
def safe_name(name):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31]

# detect type
def detect_name(h):
    h = h.lower()

    if "pression" in h or "pressio" in h:
        return "Pressiometrique"
    if "compression" in h or "résistance" in h:
        return "Compression_simple"
    if "coord" in h or (" x " in h and " y " in h):
        return "Coordonnees"
    if "litho" in h or "formation" in h:
        return "Lithologie"

    return "Tableau"

# upload
pdf_file = st.file_uploader("Importer PDF", type=["pdf"])

if pdf_file:

    # save temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name

    st.info("Extraction en cours...")

    # extraction camelot
    tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")

    groupes = {}

    for table in tables:
        df = clean_df(table.df)

        if df.empty or len(df) < 2:
            continue

        header_index = get_header(df)
        header = list(df.iloc[header_index])
        data = df.iloc[header_index + 1:]

        header_norm = normalize_header(header)

        rows_clean = []

        for _, row in data.iterrows():
            row_list = list(row)
            row_norm = normalize_header(row_list)

            # skip header duplicate
            if row_norm == header_norm:
                continue

            if all(str(x).strip() == "" for x in row_list):
                continue

            rows_clean.append(row_list)

        if not rows_clean:
            continue

        key = f"{len(header)}__{header_norm}"
        nom = detect_name(header_norm)

        if key not in groupes:
            groupes[key] = {
                "nom": nom,
                "header": header,
                "rows": []
            }

        groupes[key]["rows"].extend(rows_clean)

    # export Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        used = {}

        for g in groupes.values():

            nom = g["nom"]

            if nom in used:
                used[nom] += 1
                sheet = f"{nom}_{used[nom]}"
            else:
                used[nom] = 1
                sheet = nom

            sheet = safe_name(sheet)

            df_final = pd.DataFrame(g["rows"], columns=g["header"])
            df_final.to_excel(writer, sheet_name=sheet, index=False)

            ws = writer.book[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)

    output.seek(0)

    st.success(f"{len(groupes)} tableaux fusionnés")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="tableaux_fusionnes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    os.remove(pdf_path)
