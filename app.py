import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
import tempfile
import os

try:
    import camelot
    CAMELOT_OK = True
except Exception:
    CAMELOT_OK = False

st.title("Extraction générale des tableaux PDF vers Excel")

pdf_file = st.file_uploader("Importer un PDF", type=["pdf"])

def clean(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()

def norm(x):
    x = clean(x).lower()
    accents = {
        "é": "e", "è": "e", "ê": "e", "ë": "e",
        "à": "a", "â": "a",
        "ç": "c",
        "ù": "u", "û": "u",
        "î": "i", "ï": "i",
        "ô": "o"
    }
    for a, b in accents.items():
        x = x.replace(a, b)
    x = re.sub(r"\s+", " ", x)
    return x.strip()

def clean_df(df):
    df = df.map(clean)
    df = df.loc[:, ~(df == "").all()]
    df = df[(df != "").any(axis=1)]
    return df.reset_index(drop=True)

def detect_header_index(df):
    best_i = 0
    best_score = -1

    for i in range(min(5, len(df))):
        row = [norm(x) for x in df.iloc[i].tolist()]
        text = " ".join(row)

        score = 0
        score += sum(1 for x in row if x and not re.fullmatch(r"[\d\.,\-]+", x))
        score += len(re.findall(r"[a-zA-Z]", text))
        score -= len(re.findall(r"\d", text))

        if score > best_score:
            best_score = score
            best_i = i

    return best_i

def header_signature(header):
    cleaned = [norm(x) for x in header]
    cleaned = [x for x in cleaned if x]
    return " | ".join(cleaned)

def similar_header(h1, h2):
    s1 = set(header_signature(h1).split())
    s2 = set(header_signature(h2).split())

    if not s1 or not s2:
        return False

    common = len(s1 & s2)
    ratio = common / max(len(s1), len(s2))

    return ratio >= 0.65 and len(h1) == len(h2)

def table_name_from_header(header):
    text = header_signature(header)

    if not text:
        return "Tableau"

    words = re.findall(r"[a-zA-Z]{3,}", text)
    words = [w.capitalize() for w in words[:3]]

    if not words:
        return "Tableau"

    name = "_".join(words)
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:25]

def add_table(groups, df):
    df = clean_df(df)

    if df.empty or len(df) < 2:
        return

    header_i = detect_header_index(df)
    header = df.iloc[header_i].tolist()
    data = df.iloc[header_i + 1:].copy()

    if not header or all(clean(x) == "" for x in header):
        return

    rows = []

    for _, row in data.iterrows():
        row_list = row.tolist()

        if not any(clean(x) for x in row_list):
            continue

        if similar_header(row_list, header):
            continue

        rows.append(row_list)

    if not rows:
        return

    found = None

    for key, g in groups.items():
        if similar_header(g["header"], header):
            found = key
            break

    if found is None:
        name = table_name_from_header(header)
        found = f"group_{len(groups)+1}"
        groups[found] = {
            "name": name,
            "header": header,
            "rows": []
        }

    groups[found]["rows"].extend(rows)

def extract_camelot(pdf_path, groups):
    if not CAMELOT_OK:
        return

    try:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
        for table in tables:
            add_table(groups, table.df)
    except Exception:
        pass

    try:
        tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
        for table in tables:
            add_table(groups, table.df)
    except Exception:
        pass

def extract_pdfplumber(pdf_path, groups):
    settings_list = [
        {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "intersection_tolerance": 5,
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "text_tolerance": 3,
        }
    ]

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for settings in settings_list:
                try:
                    tables = page.extract_tables(settings)
                    for table in tables:
                        if table and len(table) > 1:
                            df = pd.DataFrame(table)
                            add_table(groups, df)
                except Exception:
                    continue

def safe_sheet_name(name, used):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    name = name[:25] if name else "Tableau"

    final = name
    i = 1

    while final in used:
        i += 1
        final = f"{name}_{i}"[:31]

    used.add(final)
    return final

if pdf_file:
    groups = {}

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name

    st.info("Extraction en cours...")

    extract_camelot(pdf_path, groups)
    extract_pdfplumber(pdf_path, groups)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used_names = set()

        for g in groups.values():
            if not g["rows"]:
                continue

            sheet_name = safe_sheet_name(g["name"], used_names)

            df_final = pd.DataFrame(g["rows"], columns=g["header"])
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 45)

    output.seek(0)

    st.success(f"{len(groups)} tableaux homogènes détectés et fusionnés.")

    st.download_button(
        "Télécharger Excel",
        data=output,
        file_name="extraction_tableaux_generale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    os.remove(pdf_path)
