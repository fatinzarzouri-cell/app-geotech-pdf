import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from io import BytesIO
import re

st.title("Extraction tableaux PDF géotechnique")

HEAD_PRESSIO = [
    "Sondage", "Profondeur (m)", "PF (MPa)", "PL (MPa)", "EM (MPa)",
    "Sigma HS (MPa)", "PL* (MPa)", "PF* (MPa)", "PL*/PF*", "EM/PL*"
]

HEAD_COMP = [
    "Sondage", "Niveau d'échantillon (m)", "Formation",
    "Résistance à la compression (MPa)"
]

def clean(v):
    if v is None:
        return ""
    return str(v).replace("\n", " ").strip()

def is_number(v):
    v = clean(v).replace(",", ".")
    return bool(re.fullmatch(r"\d+(\.\d+)?", v))

def style_sheet(ws):
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    header_fill = PatternFill("solid", fgColor="9DC3E6")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)

def extraire_pressio(pdf):
    rows = []
    sondage = ""

    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.splitlines()

        for line in lines:
            line = clean(line)

            m = re.search(r"Sondage\s+(SP_[A-Za-z0-9]+)", line)
            if m:
                sondage = m.group(1)
                continue

            parts = line.split()
            if sondage and len(parts) >= 9 and is_number(parts[0]):
                nums = parts[:9]
                rows.append([sondage] + nums)

    return rows

def extraire_compression(pdf):
    rows = []
    in_section = False
    sondage = ""

    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.splitlines()

        for line in lines:
            line = clean(line)

            if "compression simple" in line.lower():
                in_section = True
                continue

            if in_section and ("conclusion" in line.lower() or "le directeur" in line.lower()):
                in_section = False

            if not in_section:
                continue

            m_sondage = re.match(r"^(SC_[A-Za-z0-9_]+)\s*(.*)", line)
            if m_sondage:
                sondage = m_sondage.group(1)
                line = m_sondage.group(2).strip()

            m = re.match(r"^(\d+[\.,]\d+\s*-\s*\d+[\.,]\d+)\s+(.+?)\s+(\d+[\.,]?\d*)$", line)
            if sondage and m:
                niveau = m.group(1)
                formation = m.group(2)
                resistance = m.group(3)
                rows.append([sondage, niveau, formation, resistance])

    return rows

pdf_file = st.file_uploader("Importer le PDF", type=["pdf"])

if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        pressio_rows = extraire_pressio(pdf)
        comp_rows = extraire_compression(pdf)

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet("Pressiometrique")
    ws1.append(HEAD_PRESSIO)
    for r in pressio_rows:
        ws1.append(r)
    style_sheet(ws1)

    ws2 = wb.create_sheet("Compression_simple")
    ws2.append(HEAD_COMP)
    for r in comp_rows:
        ws2.append(r)
    style_sheet(ws2)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"Extraction terminée : {len(pressio_rows)} lignes pressiométriques et {len(comp_rows)} lignes compression simple.")

    st.download_button(
        label="Télécharger Excel organisé",
        data=output,
        file_name="extraction_geotech_propre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
