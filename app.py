import streamlit as st
import camelot
import pandas as pd
from io import BytesIO
import tempfile
import os
import re

st.title("Extraction tableaux PDF → Excel")

pdf_file = st.file_uploader("Importer PDF", type=["pdf"])

def clean(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()

def clean_df(df):
    try:
        df = df.map(clean)
    except:
        df = df.applymap(clean)

    df = df.loc[:, ~(df == "").all()]
    df = df[(df != "").any(axis=1)]
    return df.reset_index(drop=True)

def is_header(row):
    txt = " ".join(row).lower()
    letters = len(re.findall(r"[a-zA-Zéèêàç]", txt))
    digits = len(re.findall(r"\d", txt))
    return letters > digits

def signature(header):
    return " | ".join([clean(x).lower() for x in header])

def safe_name(name):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31] if name else "Tableau"

if pdf_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name

    st.info("Extraction en cours...")

    all_tables = []

    for flavor in ["lattice", "stream"]:
        try:
            tables = camelot.read_pdf(pdf_path, pages="all", flavor=flavor)
            for t in tables:
                df = clean_df(t.df)
                if not df.empty and len(df) > 1:
                    all_tables.append(df)
        except:
            pass

    groupes = {}
    last_key_by_cols = {}

    for df in all_tables:
        first_row = df.iloc[0].tolist()
        nb_cols = len(first_row)

        if is_header(first_row):
            header = first_row
            data = df.iloc[1:].values.tolist()
            key = f"{nb_cols}_{signature(header)}"

            if key not in groupes:
                groupes[key] = {
                    "header": header,
                    "rows": []
                }

            groupes[key]["rows"].extend(data)
            last_key_by_cols[nb_cols] = key

        else:
            if nb_cols in last_key_by_cols:
                key = last_key_by_cols[nb_cols]
                groupes[key]["rows"].extend(df.values.tolist())
            else:
                header = [f"Colonne_{i+1}" for i in range(nb_cols)]
                key = f"{nb_cols}_sans_titre_{len(groupes)+1}"
                groupes[key] = {
                    "header": header,
                    "rows": df.values.tolist()
                }
                last_key_by_cols[nb_cols] = key

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used = set()
        compteur = 1

        for key, g in groupes.items():
            sheet_name = f"Tableau_{compteur}"
            sheet_name = safe_name(sheet_name)

            while sheet_name in used:
                compteur += 1
                sheet_name = safe_name(f"Tableau_{compteur}")

            used.add(sheet_name)

            final_df = pd.DataFrame(g["rows"], columns=g["header"])
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 45)

            compteur += 1

    output.seek(0)

    st.success(f"{len(groupes)} tableaux fusionnés.")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="tableaux_fusionnes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    os.remove(pdf_path)
