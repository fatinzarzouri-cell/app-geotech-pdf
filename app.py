import streamlit as st
import camelot
import pandas as pd
from io import BytesIO
import tempfile, os, re

st.title("Extraction tableaux PDF → Excel")

pdf_file = st.file_uploader("Importer PDF", type=["pdf"])

def clean(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()

def clean_df(df):
    df = df.map(clean)
    df = df.loc[:, ~(df == "").all()]
    df = df[(df != "").any(axis=1)]
    return df.reset_index(drop=True)

def is_fake_table(df):
    text = " ".join(df.astype(str).values.flatten()).lower()
    bad_words = ["maitre d", "marché", "marche", "date", "référence", "reference", "page", "rapport"]
    good_words = ["sondage", "profondeur", "pression", "module", "résistance", "resistance", "formation"]
    return sum(w in text for w in bad_words) >= 3 and sum(w in text for w in good_words) == 0

def find_header_rows(df):
    rows = []
    for i in range(min(5, len(df))):
        txt = " ".join(df.iloc[i].tolist()).lower()
        letters = len(re.findall(r"[a-zA-Zéèêàç]", txt))
        digits = len(re.findall(r"\d", txt))
        if letters > digits:
            rows.append(i)
    return rows if rows else [0]

def make_header(df, header_rows):
    ncols = df.shape[1]
    header = []
    for c in range(ncols):
        parts = []
        for r in header_rows:
            val = clean(df.iat[r, c])
            if val:
                parts.append(val)
        header.append(" ".join(parts).strip() or f"Colonne_{c+1}")
    return header

def norm_header(header):
    txt = " | ".join(header).lower()
    txt = txt.replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c")
    txt = re.sub(r"\s+", " ", txt)
    txt = re.sub(r"[^a-z0-9|*/() ]", "", txt)
    return txt.strip()

def same_table(h1, h2):
    return len(h1) == len(h2) and norm_header(h1) == norm_header(h2)

def table_name(header, idx):
    txt = norm_header(header)
    if "pression" in txt or "pressio" in txt or ("pl" in txt and "em" in txt):
        return "Pressiometrique"
    if "compression" in txt or "resistance" in txt:
        return "Compression_simple"
    if "coord" in txt or (" x " in txt and " y " in txt):
        return "Coordonnees"
    if "lithologie" in txt or "formation" in txt:
        return "Lithologie"
    return f"Tableau_{idx}"

def safe_sheet(name, used):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)[:31]
    base = name
    i = 1
    while name in used:
        i += 1
        name = f"{base}_{i}"[:31]
    used.add(name)
    return name

if pdf_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name

    st.info("Extraction en cours...")

    tables = camelot.read_pdf(
        pdf_path,
        pages="all",
        flavor="lattice",
        line_scale=40
    )

    groups = []

    for t in tables:
        df = clean_df(t.df)

        if df.empty or len(df) < 2:
            continue

        if is_fake_table(df):
            continue

        header_rows = find_header_rows(df)
        header = make_header(df, header_rows)
        data_start = max(header_rows) + 1
        data = df.iloc[data_start:].values.tolist()

        data = [
            row for row in data
            if any(clean(x) for x in row)
        ]

        if not data:
            continue

        found = None
        for g in groups:
            if same_table(g["header"], header):
                found = g
                break

        if found is None:
            groups.append({
                "header": header,
                "rows": data
            })
        else:
            found["rows"].extend(data)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used = set()

        for i, g in enumerate(groups, start=1):
            name = safe_sheet(table_name(g["header"], i), used)

            final_df = pd.DataFrame(g["rows"], columns=g["header"])
            final_df.to_excel(writer, sheet_name=name, index=False)

            ws = writer.book[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 45)

    output.seek(0)

    st.success(f"{len(groups)} tableaux réels fusionnés.")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="tableaux_reels_fusionnes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    os.remove(pdf_path)
