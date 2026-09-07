import os
import io
import re
import html
import textwrap
import streamlit as st
from crewai import Agent, Crew, Process, Task, LLM
from docx import Document
from docx.shared import Pt, RGBColor, Inches
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pypdf

def extract_text_from_file(uploaded_file):
    """Extrahiert Text aus PDF-, DOCX- und TXT-Dateien."""
    file_type = uploaded_file.name.split('.')[-1].lower()
    extracted_text = ""
    
    try:
        if file_type == "txt":
            extracted_text = uploaded_file.read().decode("utf-8")
        elif file_type == "docx":
            doc = Document(uploaded_file)
            extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        elif file_type == "pdf":
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
    except Exception as e:
        st.error(f"Fehler beim Einlesen der Datei: {e}")
        
    return extracted_text.strip()

# Seiten-Konfiguration
st.set_page_config(page_title="Case Structuring Agent", page_icon="📊", layout="wide")

# EXECUTIVE C-LEVEL STYLING (CSS INJECTION)
EXECUTIVE_CSS = """
<style>
    /* Typografie & Globale Schriftart */
    html, body, [class*="css"] {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif !important;
    }

    /* Primary Buttons (Haupt-Aktionen) */
    div.stButton > button {
        background-color: #0F2C59 !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        border: 1px solid #0F2C59 !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.25rem !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 2px 4px rgba(15, 44, 89, 0.1) !important;
    }

    div.stButton > button:hover {
        background-color: #1E40AF !important;
        border-color: #1E40AF !important;
        box-shadow: 0 4px 12px rgba(15, 44, 89, 0.25) !important;
        transform: translateY(-1px) !important;
    }

    /* Header Card Refinement */
    .executive-header {
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
        padding: 22px 26px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #0F2C59;
        box-shadow: 0 4px 12px rgba(15, 44, 89, 0.03);
        margin-bottom: 24px;
    }

    .tech-pill {
        display: inline-block;
        background-color: #EFF6FF;
        color: #1E40AF;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 12px;
        border: 1px solid #BFDBFE;
        margin-right: 6px;
    }

    /* Container Card Styling für Tab-Inhalte */
    [data-testid="stTabPanel"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 24px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02) !important;
        margin-top: 12px !important;
    }

    /* KPI Metric Cards Redesign */
    [data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #0F2C59 !important;
        padding: 16px 20px !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(15, 44, 89, 0.04) !important;
    }

    [data-testid="stMetricLabel"] {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #64748B !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }

    [data-testid="stMetricValue"] {
        color: #0F2C59 !important;
        font-weight: 700 !important;
        font-size: 20px !important;
    }

    /* Tabs Styling (C-Level Navigation) */
    button[data-baseweb="tab"] {
        font-weight: 600 !important;
        color: #64748B !important;
        padding: 12px 20px !important;
        border-radius: 4px 4px 0 0 !important;
    }

    button[aria-selected="true"] {
        color: #0F2C59 !important;
        border-bottom-color: #0F2C59 !important;
        border-bottom-width: 3px !important;
        background-color: #F8FAFC !important;
    }

    /* Seitenleiste (Sidebar Refinement) */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }

    /* Formular-Elemente & Textareas */
    textarea {
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }
    textarea:focus {
        border-color: #0F2C59 !important;
        box-shadow: 0 0 0 1px #0F2C59 !important;
    }
</style>
"""
st.markdown(EXECUTIVE_CSS, unsafe_allow_html=True)

# 1. API Key prüfen & abfangen
api_key = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not api_key:
    api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    st.error("⚠️ GEMINI_API_KEY wurde nicht gefunden. Bitte tragen Sie Ihren Key in `.streamlit/secrets.toml` oder in den Streamlit Cloud Secrets ein.")
    st.stop()

os.environ["GEMINI_API_KEY"] = api_key

# 2. Modell initialisieren mit Request-Timeout und Auto-Retries bei Server-Spikes
gemini_llm = LLM(
    model="gemini-3.6-flash",
    api_key=api_key,
    request_timeout=90,
    max_retries=3
)

# 3. Hilfsfunktionen für Graphviz & Exporte
def generate_mece_dot_string(mece_text):
    """Erzeugt einen sauberen DOT-String für st.graphviz_chart mit dynamischem Zeilenumbruch (ohne Abschneiden)."""
    lines = [line.strip() for line in mece_text.split('\n') if line.strip()]
    dot_lines = [
        'digraph MECETree {',
        '    graph [rankdir=LR, bgcolor="transparent", nodesep=0.35, ranksep=0.55];',
        '    node [shape=box, style="filled,rounded", fontname="Segoe UI, sans-serif", fontsize=10, margin="0.2,0.12"];',
        '    edge [color="#94A3B8", penwidth=1.2, arrowsize=0.8];',
        '    root [label="Case Problem Breakdown", fillcolor="#0F2C59", fontcolor="#FFFFFF", color="#0F2C59", fontsize=11];'
    ]

    current_l1 = None
    node_count = 0

    for line in lines:
        clean_line = re.sub(r'[*#]', '', line).strip()
        m_l2 = re.match(r'^(\d+\.\d+)\s*(.*)', clean_line)
        m_l1 = re.match(r'^(\d+)\.\s*(.*)', clean_line)

        if m_l1 and not m_l2:
            num, txt = m_l1.groups()
            current_l1 = f"l1_{num.replace('.', '_')}"
            safe_txt = txt.replace('"', '\\"').strip()
            
            # Dynamischer Zeilenumbruch nach ca. 28 Zeichen
            wrapped_txt = "\\n".join(textwrap.wrap(safe_txt, width=28))
            full_label = f"{num}. {wrapped_txt}"
            
            dot_lines.append(f'    {current_l1} [label="{full_label}", fillcolor="#EFF6FF", color="#BFDBFE", fontcolor="#1E40AF"];')
            dot_lines.append(f'    root -> {current_l1};')

        elif m_l2:
            num, txt = m_l2.groups()
            node_count += 1
            l2_id = f"l2_{node_count}"
            safe_txt = txt.replace('"', '\\"').strip()
            
            # Dynamischer Zeilenumbruch nach ca. 32 Zeichen
            wrapped_txt = "\\n".join(textwrap.wrap(safe_txt, width=32))
            full_label = f"{num} {wrapped_txt}"
            
            dot_lines.append(f'    {l2_id} [label="{full_label}", fillcolor="#FFFFFF", color="#CBD5E1", fontcolor="#334155"];')
            if current_l1:
                dot_lines.append(f'    {current_l1} -> {l2_id};')
            else:
                dot_lines.append(f'    root -> {l2_id};')

    dot_lines.append('}')
    return '\n'.join(dot_lines)

def create_html_report(title, framework, analysis, mece, hypothesis, sub_text, fw_label, sec1, sec2, sec3, footer_text):
    """Erstellt ein professionelles Executive HTML Dashboard mit korrekten A4-PDF-Druckrändern."""
    def md_to_html(md_text):
        lines = md_text.strip().split('\n')
        html_out = []
        in_list = False
        table_lines = []

        def format_text(txt):
            txt = html.escape(txt.strip())
            txt = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', txt)
            txt = re.sub(r'\*(.*?)\*', r'<em>\1</em>', txt)
            return txt

        def render_table(t_lines):
            if not t_lines:
                return ""
            rows = []
            for line in t_lines:
                cleaned = line.strip().strip('|')
                cells = [c.strip() for c in cleaned.split('|')]
                if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                    continue
                rows.append(cells)

            if not rows:
                return ""

            out = ['<table class="executive-table">']
            out.append('<colgroup><col style="width: 20%;"><col style="width: 38%;"><col style="width: 22%;"><col style="width: 20%;"></colgroup>')
            out.append('<thead><tr>')
            for cell in rows[0]:
                out.append(f'<th>{format_text(cell)}</th>')
            out.append('</tr></thead><tbody>')

            for row in rows[1:]:
                out.append('<tr>')
                for cell in row:
                    out.append(f'<td>{format_text(cell)}</td>')
                out.append('</tr>')
            out.append('</tbody></table>')
            return '\n'.join(out)

        for line in lines:
            line_str = line.strip()

            if '|' in line_str and line_str.count('|') >= 2:
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                table_lines.append(line_str)
                continue
            else:
                if table_lines:
                    html_out.append(render_table(table_lines))
                    table_lines = []

            if not line_str:
                continue

            if re.match(r'^\d+\.\d+', line_str) or re.match(r'^[*\-]\s*\d+\.\d+', line_str):
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                html_out.append(f'<p style="margin-left: 24px; margin-top: 3px; margin-bottom: 5px;">{format_text(clean_txt)}</p>')
            elif re.match(r'^\d+\.\s', line_str) or re.match(r'^[*\-]\s*\d+\.\s', line_str):
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                html_out.append(f'<p style="margin-top: 14px; margin-bottom: 4px; font-weight: 600; color: #0F2C59;">{format_text(clean_txt)}</p>')
            elif line_str.startswith('### '):
                html_out.append(f'<h3>{format_text(line_str[4:])}</h3>')
            elif line_str.startswith('## '):
                html_out.append(f'<h2>{format_text(line_str[3:])}</h2>')
            elif line_str.startswith('# '):
                html_out.append(f'<h2>{format_text(line_str[2:])}</h2>')
            elif line_str.startswith('- ') or line_str.startswith('* '):
                if not in_list:
                    html_out.append('<ul class="executive-list">')
                    in_list = True
                html_out.append(f'<li>{format_text(line_str[2:])}</li>')
            else:
                if in_list:
                    html_out.append('</ul>')
                    in_list = False
                html_out.append(f'<p>{format_text(line_str)}</p>')

        if table_lines:
            html_out.append(render_table(table_lines))
        if in_list:
            html_out.append('</ul>')

        return '\n'.join(html_out)

    html_content = f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        @page {{ 
            size: A4 portrait; 
            margin: 18mm 18mm 18mm 18mm; 
        }}
        body {{ 
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; 
            line-height: 1.5; 
            color: #1E293B; 
            background-color: #F8FAFC; 
            margin: 0; 
            padding: 24px; 
        }}
        .container {{ 
            max-width: 850px; 
            margin: 0 auto; 
            background: #FFFFFF; 
            padding: 32px 36px; 
            border-radius: 8px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.05); 
            box-sizing: border-box; 
        }}
        .header {{ border-bottom: 2px solid #0F2C59; padding-bottom: 12px; margin-bottom: 22px; }}
        .badge {{ display: inline-block; background: #0F2C59; color: #FFFFFF; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; }}
        h1 {{ color: #0F2C59; font-size: 22px; margin: 4px 0; font-weight: 700; }}
        h2 {{ color: #0F2C59; font-size: 15px; border-bottom: 1px solid #E2E8F0; padding-bottom: 5px; margin-top: 22px; margin-bottom: 10px; font-weight: 600; page-break-after: avoid; }}
        h3 {{ color: #334155; font-size: 13px; margin-top: 14px; margin-bottom: 5px; font-weight: 600; page-break-after: avoid; }}
        p, li {{ font-size: 11.5px; color: #334155; margin-bottom: 5px; }}
        ul.executive-list {{ padding-left: 20px; margin: 6px 0; }}
        ul.executive-list li {{ margin-bottom: 3px; }}

        .executive-table {{ width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 10.5px; page-break-inside: avoid; table-layout: fixed; }}
        .executive-table th {{ background-color: #0F2C59; color: #FFFFFF; font-weight: bold; text-align: left; padding: 7px 8px; border: 1px solid #0F2C59; white-space: nowrap; }}
        .executive-table td {{ border: 1px solid #CBD5E1; padding: 7px 8px; vertical-align: top; word-wrap: break-word; }}
        .executive-table tr:nth-child(even) {{ background-color: #F8FAFC; }}

        .section-block {{ page-break-inside: avoid; }}
        .footer {{ margin-top: 30px; padding-top: 10px; border-top: 1px solid #E2E8F0; font-size: 10px; color: #94A3B8; text-align: center; }}

        @media print {{
            @page {{
                size: A4 portrait;
                margin: 18mm 18mm 18mm 18mm !important;
            }}
            html, body {{ 
                background: #FFFFFF !important; 
                margin: 0 !important; 
                padding: 0 !important; 
                width: 100% !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            .container {{ 
                box-shadow: none !important; 
                border: none !important;
                padding: 0 !important; 
                margin: 0 !important; 
                width: 100% !important; 
                max-width: 100% !important; 
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <span class="badge">{framework}</span>
            <h1>{title}</h1>
            <p style="color: #64748B; margin: 0; font-size: 11px;">{sub_text}</p>
        </div>

        <div class="section-block">
            <h2>{sec1}</h2>
            {md_to_html(analysis)}
        </div>

        <div class="section-block">
            <h2>{sec2}</h2>
            {md_to_html(mece)}
        </div>

        <div class="section-block">
            <h2>{sec3}</h2>
            {md_to_html(hypothesis)}
        </div>

        <div class="footer">
            {footer_text}
        </div>
    </div>
</body>
</html>"""
    return html_content

def create_docx_report(title, framework, analysis, mece, hypothesis, fw_label, sec1, sec2, sec3, footer_text):
    """Erstellt ein professionelles Word-Dokument."""
    doc = Document()

    heading = doc.add_heading(title, level=0)
    heading.style.font.color.rgb = RGBColor(15, 44, 89)

    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(f"{fw_label}: {framework}")
    run_sub.bold = True
    p_sub.paragraph_format.space_after = Pt(18)

    def add_formatted_text(paragraph, text):
        parts = re.split(r'(\*\*.*?\*\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            else:
                if part:
                    paragraph.add_run(part)

    def render_docx_table(t_lines):
        rows_data = []
        for line in t_lines:
            cleaned = line.strip().strip('|')
            cells = [c.strip() for c in cleaned.split('|')]
            if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                continue
            rows_data.append(cells)

        if not rows_data:
            return

        col_count = max(len(r) for r in rows_data)
        table = doc.add_table(rows=len(rows_data), cols=col_count)
        table.style = 'Table Grid'

        for r_idx, row in enumerate(rows_data):
            for c_idx, cell_text in enumerate(row):
                if c_idx < col_count:
                    cell = table.cell(r_idx, c_idx)
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
                    add_formatted_text(p, cell_text)
                    if r_idx == 0:
                        for run in p.runs:
                            run.bold = True

    def add_md_section(section_title, md_text):
        h = doc.add_heading(section_title, level=1)
        h.style.font.color.rgb = RGBColor(15, 44, 89)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)

        lines = md_text.strip().split('\n')
        table_lines = []

        for line in lines:
            line_str = line.strip()

            if '|' in line_str and line_str.count('|') >= 2:
                table_lines.append(line_str)
                continue
            else:
                if table_lines:
                    render_docx_table(table_lines)
                    table_lines = []

            if not line_str:
                continue

            if re.match(r'^\d+\.\d+', line_str) or re.match(r'^[*\-]\s*\d+\.\d+', line_str):
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_after = Pt(3)
                add_formatted_text(p, clean_txt)
            elif re.match(r'^\d+\.\s', line_str) or re.match(r'^[*\-]\s*\d+\.\s', line_str):
                clean_txt = re.sub(r'^[*\-]\s*', '', line_str)
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(3)
                add_formatted_text(p, clean_txt)
            elif line_str.startswith('### '):
                p = doc.add_heading(level=2)
                add_formatted_text(p, line_str[4:])
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(2)
            elif line_str.startswith('## '):
                p = doc.add_heading(level=2)
                add_formatted_text(p, line_str[3:])
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
            elif line_str.startswith('- ') or line_str.startswith('* '):
                p = doc.add_paragraph(style='List Bullet')
                add_formatted_text(p, line_str[2:])
                p.paragraph_format.space_after = Pt(2)
            else:
                p = doc.add_paragraph()
                add_formatted_text(p, line_str)
                p.paragraph_format.space_after = Pt(4)

        if table_lines:
            render_docx_table(table_lines)

    add_md_section(sec1, analysis)
    add_md_section(sec2, mece)
    add_md_section(sec3, hypothesis)

    p_footer = doc.add_paragraph()
    p_footer.paragraph_format.space_before = Pt(30)
    p_footer.alignment = 1
    run_footer = p_footer.add_run(footer_text)
    run_footer.font.size = Pt(9)
    run_footer.font.color.rgb = RGBColor(148, 163, 184)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def create_excel_report(title, framework, analysis, mece, hypothesis):
    """Erstellt ein professionell strukturiertes Excel-Workbook (.xlsx) mit C-Level Styling ohne Roh-Markdown-Artefakte."""
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="0F2C59")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F2C59")
    regular_font = Font(name="Calibri", size=11)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    def clean_md(text):
        if not text:
            return ""
        txt = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        txt = re.sub(r'\*(.*?)\*', r'\1', txt)
        txt = re.sub(r'^[#*:\s]+', '', txt).strip()
        txt = txt.replace('**', '').replace('*', '').replace('#', '').strip()
        return txt

    # 1. TAB: Executive Summary & SCR
    ws1 = wb.active
    ws1.title = "Executive Summary"
    ws1.views.sheetView[0].showGridLines = True

    ws1['A1'] = title
    ws1['A1'].font = title_font
    ws1['A2'] = f"Framework Focus: {framework}"
    ws1['A2'].font = Font(name="Calibri", size=11, italic=True, color="64748B")

    row_idx = 4
    for line in analysis.split('\n'):
        line_str = line.strip()
        if not line_str:
            continue
        
        is_heading = line_str.startswith('#') or line_str.startswith('**')
        cleaned_line = clean_md(line_str)
        if not cleaned_line:
            continue

        cell = ws1.cell(row=row_idx, column=1, value=cleaned_line)
        cell.font = bold_font if is_heading else regular_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        row_idx += 1

    ws1.column_dimensions['A'].width = 110

    # 2. TAB: MECE Structure
    ws2 = wb.create_sheet(title="MECE Structure")
    ws2.views.sheetView[0].showGridLines = True
    ws2['A1'] = "MECE Problem Breakdown"
    ws2['A1'].font = title_font

    row_idx = 3
    for line in mece.split('\n'):
        line_str = line.strip()
        if not line_str:
            continue

        cleaned_line = clean_md(line_str)
        if not cleaned_line:
            continue

        if re.match(r'^\d+\.\s', cleaned_line):
            cell = ws2.cell(row=row_idx, column=1, value=cleaned_line)
            cell.font = bold_font
        elif re.match(r'^\d+\.\d+', cleaned_line):
            cell = ws2.cell(row=row_idx, column=2, value=cleaned_line)
            cell.font = regular_font
        else:
            cell = ws2.cell(row=row_idx, column=1, value=cleaned_line)
            cell.font = regular_font

        cell.alignment = Alignment(wrap_text=True, vertical="top")
        row_idx += 1

    ws2.column_dimensions['A'].width = 35
    ws2.column_dimensions['B'].width = 80

    # 3. TAB: KPI & Hypothesen Matrix
    ws3 = wb.create_sheet(title="KPI & Hypotheses")
    ws3.views.sheetView[0].showGridLines = True
    ws3['A1'] = "Hypotheses & Quantified KPI Matrix"
    ws3['A1'].font = title_font

    table_data = []
    for line in hypothesis.split('\n'):
        if '|' in line and line.count('|') >= 2:
            cleaned = line.strip().strip('|')
            cells = [clean_md(c) for c in cleaned.split('|')]
            if not all(re.match(r'^:?-+:?$', c) for c in cells if c):
                table_data.append(cells)

    if table_data:
        start_row = 3
        for r_idx, row_values in enumerate(table_data):
            current_row = start_row + r_idx
            for c_idx, val in enumerate(row_values):
                cell = ws3.cell(row=current_row, column=c_idx + 1, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center" if c_idx == 0 or r_idx == 0 else "left")

                if r_idx == 0:
                    cell.fill = header_fill
                    cell.font = header_font
                else:
                    cell.font = regular_font

        for col in ws3.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws3.column_dimensions[col_letter].width = min(max(max_len + 4, 18), 50)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# 4. Seitenleisten-Initialisierung (Sprachauswahl)
with st.sidebar:
    st.markdown("### ⚙️ Executive Settings")
    language = st.selectbox("Language / Sprache", ["Deutsch", "English"])

# 5. DYNAMISCHE UI- UND REPORT-TEXTE (GLOBAL DEFINIERT)
rep_sub = "Automatisierter KI-Fallanalysebericht"
rep_lbl_fw = "Fokus-Framework"
rep_sec1 = "1. Executive Summary & SCR"
rep_sec2 = "2. MECE-Problemstruktur"
rep_sec3 = "3. Hypothesen & KPI-Matrix"
rep_footer = "Erstellt durch KI Consulting & Case Structuring Agent | Vertrauliches Analyse-Tool"

if language == "English":
    ui_title = "📊 Consulting Case Structuring Agent"
    ui_subtitle = "Structured case analysis, MECE issue trees, and data-driven hypothesis development."
    ui_framework_label = "Select Framework Focus"
    ui_input_label = "Enter Case Briefing here:"
    ui_input_placeholder = "Paste the client's problem description here or load a demo case above..."
    ui_privacy_notice = "⚠️ **Data Privacy Notice:** Do not enter confidential client data, proprietary metrics, or PII. Use anonymized or public data only."
    ui_button = "🚀 Analyze Case"
    ui_demo_btn = "💡 Load Demo Case"
    ui_warning = "Please enter a case briefing first."
    ui_status_start = "🔍 AI Crew is analyzing the case briefing..."
    ui_status_done = "✅ Analysis completed successfully!"
    ui_kpi_hdr = "📈 Executive Metrics & Scope"
    ui_m1_label = "Framework Focus"
    ui_m2_label = "MECE Coverage"
    ui_m2_val = "100%"
    ui_m2_delta = "Mutually Exclusive"
    ui_m3_label = "Hypotheses"
    ui_m3_val = "3 Formulated"
    ui_m3_delta = "KPI Validated"
    ui_tab1 = "📌 Executive Summary & SCR"
    ui_tab2 = "🌳 MECE Issue Tree"
    ui_tab3 = "💡 Hypotheses & KPI Matrix"
    ui_report_hdr = "Client Analysis Report"
    ui_export_hdr = "📥 Export Deliverable"
    ui_format_label = "Select Export Format:"
    ui_key_hdr = "🔑 API Key Configuration"
    ui_key_label = "Custom Gemini API Key (Optional)"
    ui_key_help = "If the global demo quota is exhausted, a personal free API key from Google AI Studio can be entered here."
    ui_sec_note = "🔒 *Input is isolated and processed strictly in-memory per session.*"
    ui_demo_info = "⚡ **Demo Preview Active:** To prevent API rate limits and ensure instant response times, a pre-validated agent result is loaded for this standard case. Manually editing the briefing will automatically trigger live AI orchestration."
    ui_view_mode = "MECE View Mode"

    rep_sub = "Automated AI Case Analysis Report"
    rep_lbl_fw = "Framework Focus"
    rep_sec1 = "1. Executive Summary & SCR"
    rep_sec2 = "2. MECE Issue Tree"
    rep_sec3 = "3. Hypotheses & KPI Matrix"
    rep_footer = "Generated by KI Consulting & Case Structuring Agent | Confidential & Professional Support Tool"
else:
    ui_title = "📊 Consulting Case Structuring Agent"
    ui_subtitle = "Strukturierte Case-Analyse, MECE-Problembäume und datengestützte Hypothesen-Entwicklung."
    ui_framework_label = "Fokus-Framework wählen"
    ui_input_label = "Case Briefing hier eingeben:"
    ui_input_placeholder = "Fügen Sie hier die Problemstellung ein oder laden Sie oben einen Demo-Case..."
    ui_privacy_notice = "⚠️ **Datenschutz- & NDA-Hinweis:** Bitte keine echten Mandantennamen, vertraulichen Unternehmenskennzahlen oder personenbezogenen Daten eingeben. Nutzen Sie ausschließlich anonymisierte Case-Informationen."
    ui_button = "🚀 Case Analysieren"
    ui_demo_btn = "💡 Demo-Case laden"
    ui_warning = "Bitte geben Sie zuerst ein Case-Briefing ein."
    ui_status_start = "🔍 Agenten-Crew analysiert die Problemstellung..."
    ui_status_done = "✅ Analyse erfolgreich abgeschlossen!"
    ui_kpi_hdr = "📈 Executive Metrics & Analyse-Fokus"
    ui_m1_label = "Fokus-Framework"
    ui_m2_label = "MECE-Abdeckung"
    ui_m2_val = "100 %"
    ui_m2_delta = "Überschneidungsfrei"
    ui_m3_label = "Arbeitshypothesen"
    ui_m3_val = "3 Hypothesen"
    ui_m3_delta = "KPI-validiert"
    ui_tab1 = "📌 Executive Summary & SCR"
    ui_tab2 = "🌳 MECE-Problemstruktur"
    ui_tab3 = "💡 Hypothesen & KPI-Matrix"
    ui_report_hdr = "Mandanten-Analysebericht"
    ui_export_hdr = "📥 Bericht Exportieren"
    ui_format_label = "Export-Format wählen:"
    ui_key_hdr = "🔑 API-Key Konfiguration"
    ui_key_label = "Eigener Gemini API-Key (Optional)"
    ui_key_help = "Falls das globale Test-Kontingent erschöpft ist, kann hier ein eigener kostenloser Key eingetragen werden."
    ui_sec_note = "🔒 *Verarbeitung erfolgt ausschließlich im flüchtigen Arbeitsspeicher.*"
    ui_demo_info = "⚡ **Demo-Vorschau aktiv:** Zur Vermeidung von API-Rate-Limits und zur Gewährleistung unmittelbarer Antwortzeiten wird für diesen Standard-Case ein vorvalidiertes Agenten-Ergebnis geladen. Bei manueller Anpassung des Briefings wird automatisch die Live-Orchestrierung gestartet."
    ui_view_mode = "MECE Ansichtsmodus"

    rep_sub = "Automatisierter KI-Fallanalysebericht"
    rep_lbl_fw = "Fokus-Framework"
    rep_sec1 = "1. Executive Summary & SCR"
    rep_sec2 = "2. MECE-Problemstruktur"
    rep_sec3 = "3. Hypothesen & KPI-Matrix"
    rep_footer = "Erstellt durch KI Consulting & Case Structuring Agent | Vertrauliches Analyse-Tool"

# 6. Ergänzung der Seitenleiste
with st.sidebar:
    st.markdown("---")
    st.markdown(f"**{ui_framework_label}**")
    framework_focus = st.selectbox(
        "",
        ["General Profitability", "Cost Reduction", "M&A Due Diligence", "Market Entry"],
        label_visibility="collapsed"
    )

    st.markdown("---")

    st.markdown(f"""
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; padding: 14px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
            <span style="font-weight: 700; font-size: 13px; color: #0F2C59;">{ui_key_hdr}</span>
            <p style="font-size: 11px; color: #64748B; margin-top: 4px; margin-bottom: 0px;">{ui_sec_note}</p>
        </div>
    """, unsafe_allow_html=True)

    user_key = st.text_input(
        ui_key_label, 
        type="password", 
        help=ui_key_help,
        label_visibility="collapsed"
    )

    if user_key.strip():
        os.environ["GEMINI_API_KEY"] = user_key.strip()
        gemini_llm = LLM(
            model="gemini-3.6-flash", 
            api_key=user_key.strip(),
            request_timeout=90,
            max_retries=3
        )

    st.markdown("---")
    with st.expander("ℹ️ System Architecture & Workflow"):
        st.markdown("""
        **Sequential Multi-Agent Crew (CrewAI):**
        1. **Senior Strategy Consultant:** Formulates high-level SCR synthesis.
        2. **MECE Framework Architect:** Deconstructs problem into a 100% MECE tree.
        3. **Strategy & Hypothesis Lead:** Derives quantified KPI validation matrix.

        *Built with Streamlit, CrewAI & Gemini 3.6-flash.*
        """)

# HAUPTBEREICH: C-Level Header Banner mit Tech-Badges
st.markdown(f"""
    <div class="executive-header">
        <div style="margin-bottom: 10px;">
            <span class="tech-pill">CrewAI Multi-Agent</span>
            <span class="tech-pill">Gemini 3.6-flash</span>
            <span class="tech-pill">MECE Standard</span>
        </div>
        <h1 style="color: #0F2C59; font-size: 25px; font-weight: 700; margin: 0 0 4px 0;">{ui_title}</h1>
        <p style="color: #64748B; font-size: 13px; margin: 0;">{ui_subtitle}</p>
    </div>
""", unsafe_allow_html=True)

# 7. Demo Cases (Bilingual)
DEMO_CASES_DE = {
    "General Profitability": "Mandant: Mittelständisches Industrieunternehmen (Umsatz: 45 Mio. €).\nProblemstellung: Die EBIT-Marge ist innerhalb der letzten 18 Monate von 11,5 % auf 3,2 % gesunken, obwohl der Umsatz stabil geblieben ist.\nZiel: Identifikation der Hauptursachen für den Margenverfall und Entwicklung konkreter Gegenmaßnahmen zur Erreichung einer Ziel-Marge von > 8,0 %.",
    "Cost Reduction": "Mandant: Internationaler Logistikdienstleister.\nProblemstellung: Stark steigende Opex-Kosten in der Flotte und im Lagerbetrieb schmälern das Gesamtergebnis um 4,5 Mio. € im Vergleich zum Vorjahr.\nZiel: Systematische Kostenstrukturanalyse zur Identifikation von Einsparpotenzialen von mindestens 15 % ohne Qualitätsverlust im Kerngeschäft.",
    "M&A Due Diligence": "Mandant: Finanzinvestor / Private Equity.\nProblemstellung: Bewertung eines potenziellen Akquisitionsziels im Bereich B2B-Software vor Beginn der detaillierten Commercial Due Diligence.\nZiel: Validierung des nachhaltigen EBITDA-Aussagewerts, Identifikation wesentlicher Geschäftsrisiken und Prüfung der Run-Rate im Hinblick auf das Synergiepotenzial.",
    "Market Entry": "Mandant: E-Commerce-Händler für Premium-Konsumgüter.\nProblemstellung: Geplante Expansion in zwei neue europäische Märkte bei einem Investitionsbudget von 2,0 Mio. €.\nZiel: Evaluierung von Markteintrittsbarrieren, Kundenakquisitionskosten (CAC) und der erwarteten Amortisationsdauer (Payback Period)."
}

DEMO_CASES_EN = {
    "General Profitability": "Client: Mid-sized industrial manufacturing company (Revenue: €45M).\nProblem Statement: EBIT margin collapsed from 11.5% to 3.2% over the last 18 months despite stable revenues.\nGoal: Identify core drivers of margin erosion and develop actionable countermeasures to restore target EBIT margin > 8.0%.",
    "Cost Reduction": "Client: International logistics provider.\nProblem Statement: Rising fleet and warehousing OPEX eroded operating result by €4.5M year-over-year.\nGoal: Systematic cost structure analysis to identify at least 15% savings without impacting service quality.",
    "M&A Due Diligence": "Client: Private Equity investor.\nProblem Statement: Evaluation of a target B2B software firm prior to commercial due diligence.\nGoal: Validate sustainable EBITDA run-rate, identify operational risks, and quantify synergy potentials.",
    "Market Entry": "Client: Premium e-commerce consumer goods retailer.\nProblem Statement: Planned expansion into two European growth markets with a €2.0M investment budget.\nGoal: Assess market entry barriers, customer acquisition costs (CAC), and expected payback period."
}

DEMO_CASES = DEMO_CASES_EN if language == "English" else DEMO_CASES_DE

# 7.1 Session State Initialisierung
if "case_text" not in st.session_state:
    st.session_state["case_text"] = ""
if "last_uploaded_file" not in st.session_state:
    st.session_state["last_uploaded_file"] = None

# 7.2 Eingabe-Header mit File-Uploader & Demo-Button
col_label, col_btn = st.columns([3, 1])
with col_label:
    st.markdown(f"**{ui_input_label}**")
with col_btn:
    if st.button(ui_demo_btn, use_container_width=True):
        st.session_state["case_text"] = DEMO_CASES.get(framework_focus, DEMO_CASES["General Profitability"])

# Datenschutz- & Compliance-Hinweis direkt vor der Eingabe
st.caption(ui_privacy_notice)

# File Uploader Baustein (PDF, DOCX, TXT)
uploaded_file = st.file_uploader(
    "📄 Datei hochladen (PDF, DOCX, TXT)", 
    type=["pdf", "docx", "txt"],
    help="Laden Sie ein bestehendes Mandanten-Dokument oder ein Case-Briefing hoch."
)

if uploaded_file is not None:
    file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state["last_uploaded_file"] != file_key:
        file_text = extract_text_from_file(uploaded_file)
        if file_text:
            st.session_state["case_text"] = file_text
            st.session_state["last_uploaded_file"] = file_key
            st.success(f"✅ Text aus '{uploaded_file.name}' erfolgreich geladen!")

# 7.3 Case Briefing Textarea Eingabefeld
case_input = st.text_area(
    "",
    value=st.session_state["case_text"],
    height=160,
    placeholder=ui_input_placeholder,
    label_visibility="collapsed"
)

st.markdown("<div style='margin-top: -10px;'></div>", unsafe_allow_html=True)
run_analysis = st.button(ui_button, use_container_width=True)

# 7.4 Formatierungsregeln für Agenten
FORMATTING_RULES = f"""
STRIKTE FORMATIERUNGS-REGELN (STRIKT EINHALTEN):
1. KEINE ASCII-Boxen oder Rahmenelemente (+---+, |---|, etc.) verwenden.
2. Für MECE-Strukturen und Baumdarstellungen AUSSCHLIESSLICH Standard-Markdown-Listen mit Einrückungen verwenden.
3. KEINE H1-Überschriften (`#`) generieren. Nutze ausschließlich Unterüberschriften ab Ebene 2 (`##`).
4. Für Tabellen ausschließlich sauberes Markdown-Tabellenformat nutzen (`| Spalte 1 | Spalte 2 |`).
5. Antworte extrem präzise, auf den Punkt fokussiert und ohne Floskeln, um die Verarbeitungszeit kurz zu halten.
6. Gesamtsprache der Ausgabe: Strikt auf {language}.
"""

# 8. PRECACHED DEMO ERGEBNISSE FÜR ALLE 4 FRAMEWORKS (DEUTSCH & ENGLISCH)
PRECACHED_DE = {
    "General Profitability": {
        "analysis": """### 1. Situation (S)
Der Mandant ist ein mittelständisches Industrieunternehmen mit einem stabilen Jahresumsatz von **45,0 Mio. €**. Die historische EBIT-Marge lag bei gesunden **11,5 %** (ca. 5,18 Mio. € EBIT).

### 2. Complication (C)
Innerhalb der letzten 18 Monate ist die EBIT-Marge drastisch um **8,3 Prozentpunkte auf 3,2 %** (ca. 1,44 Mio. € EBIT) eingebrochen. Da der Umsatz konstant geblieben ist, resultiert der operative Ergebnisverlust von **~3,74 Mio. €** vollständig aus einer Verschlechterung der Kostenstruktur und ungünstigen Preis-/Mix-Effekten.

### 3. Key Question (KQ)
Welche spezifischen Kosten- und Mix-Treiber haben die Marge erodiert, und mit welchem Maßnahmenpaket kann die EBIT-Marge nachhaltig auf die Zielmarke von **> 8,0 %** (> 3,60 Mio. € EBIT) angehoben werden?

### 4. Resolution & Strategic Approach (R)
Zur Erreichung der Ziel-Marge ist eine EBIT-Steigerung um mindestens **2,16 Mio. €** erforderlich. Dies erfordert eine Ursachenanalyse entlang des Profitabilitäts-Baums sowie die Implementierung eines zweiphasigen Optimierungsprogramms.""",
        "mece": """**1. Erlösqualität & Preisdurchsetzung (Umsatz- & Mix-Hebel)**
1.1 **Preisanpassung & Indexierung:** Unzureichende Weitergabe gestiegener Inputkosten an Endkunden.
1.2 **Portfolio-Mix-Verschiebung:** Shift von hochmargigen Spezialprodukten zu margenschwachen Standardprodukten.
1.3 **Konditionen-Management:** Hohe Rabatte und ungünstige Frachtkonditionen bei A-Kunden.

**2. Variable Herstellungskosten (COGS / Direct Costs)**
2.1 **Einkauf & Material:** Preisanstiege bei Rohstoffen ohne adäquates Sourcing-Gegenhalten.
2.2 **Fertigungseffizienz:** Sinkende OEE-Raten, erhöhte Ausschussquoten und Überstunden.
2.3 **Logistik & Energie:** Gestiegene Fracht- und Energiekosten pro Produktionseinheit.

**3. Operative Fixkosten & Overhead (OPEX / Indirect Costs)**
3.1 **SG&A-Kosten:** Ungesteuerter Anstieg der Verwaltungs- und Vertriebskosten (Fixed Cost Creep).
3.2 **Instandhaltung & F&E:** Erhöhte Wartungsaufwände veralteter Anlagen und uneffiziente Projekte.""",
        "hypothesis": """| Bereich | Primäre Hypothese | KPI / Ziel-Benchmark | Erwarteter EBIT-Hebel |
| :--- | :--- | :--- | :--- |
| **Pricing & Mix** | Selektive Preiserhöhungen (3,5 %) und Indexierung von Rohstoffklauseln stabilisieren den Deckungsbeitrag. | **Price Realization Rate > 85 %** | **+0,90 Mio. €** |
| **COGS & Sourcing** | Neuausschreibung der Top-20 Lieferanten und Reduktion der Ausschussquote senken variable Stückkosten. | **Wareneinsatzquote < 56,1 %** | **+0,85 Mio. €** |
| **SG&A / Overhead** | Einfrieren nicht-kritischer Sachkosten (Discretionary Spending Freeze) stoppt Fixed Cost Creep. | **SG&A-Quote < 17,6 %** | **+0,41 Mio. €** |"""
    },
    "Cost Reduction": {
        "analysis": """### 1. Situation (S)
Der Mandant ist ein international tätiger Logistikdienstleister mit Flotten- und Lagerstandorten in ganz Europa.

### 2. Complication (C)
Stark gestiegene Treibstoff- und Personalkosten sowie Ineffizienzen in der Lagerabwicklung schmälerte das operative Jahresergebnis um **4,5 Mio. €** gegenüber dem Vorjahr.

### 3. Key Question (KQ)
Über welche Hebel lassen sich die OPEX-Strukturen nachhaltig um mindestens **15 %** senken, ohne die Qualität und Termintreue im Kerngeschäft zu beeinträchtigen?

### 4. Resolution & Strategic Approach (R)
Implementierung eines Drei-Säulen-Kostenreduzierungsprogramms mit Fokus auf Tourenoptimierung, Automatisierung im Lager und Reduktion der indirekten Sachkosten.""",
        "mece": """**1. Flotten- & Treibstoffeffizienz (Direkte OPEX)**
1.1 **Routenoptimierung:** Reduktion von Leerfahrten durch KI-gestützte Disposition und Telematik.
1.2 **Fuel Management:** Nachverhandlung von Tankkarten-Konditionen und Fahrerschulungen für eﬃziente Fahrweise.
1.3 **Instandhaltung:** Optimierung der Wartungsintervalle und Flottenverjüngung zur Reduktion von Reparaturaufwänden.

**2. Lager- & Standortlogistik (Infrastruktur & Personal)**
2.1 **Prozessautomatisierung:** Einführung von Barcode-Scanning und Optimierung der Pick-&-Pack-Pfade.
2.2 **Flächennutzung:** Konsolidierung schwach ausgelasteter Lagerstandorte zur Fixkostenreduktion.
2.3 **Schichtplanung:** Flexible Personaleinsatzplanung zur Minimierung teurer Überstunden.

**3. Indirekter Einkauf & Overhead (Indirect OPEX)**
3.1 **Lieferantenbündelung:** Neuverhandlung der Top-15 Sachkostenverträge (Verpackung, IT, Reinigung).
3.2 **Verwaltungsprozesse:** Digitalisierung der Frachtdokumentation zur Senkung der SG&A-Quote.""",
        "hypothesis": """| Bereich | Primäre Hypothese | KPI / Ziel-Benchmark | Erwarteter EBIT-Hebel |
| :--- | :--- | :--- | :--- |
| **Flottenoptimierung** | KI-Disposition und Fahrertrainings senken den Treibstoffverbrauch pro tkm um 8 %. | **Fuel Efficiency +8 %** | **+1,80 Mio. €** |
| **Lagerlogistik** | Verdichtung der Lagerflächen ermöglicht die Schließung eines Nebenstandorts. | **Flächenproduktivität +12 %** | **+1,50 Mio. €** |
| **Procurement & Overhead** | Bündelung des Sachkosten-Einkaufs senkt Dienstleisterkonditionen nachhaltig. | **Procurement Savings > 10 %** | **+1,20 Mio. €** |"""
    },
    "M&A Due Diligence": {
        "analysis": """### 1. Situation (S)
Ein Private Equity Investor prüft den Mehrheitserwerb an einem schnell wachsenden B2B-Softwareunternehmen (SaaS).

### 2. Complication (C)
Vor Beginn der detaillierten Commercial Due Diligence bestehen Unsicherheiten bezüglich der Belastbarkeit der ARR-Run-Rate, der Kunden-Retention und des realisierbaren Synergiepotenzials.

### 3. Key Question (KQ)
Wie nachhaltig ist das organische Umsatzwachstum und welche operativen Wertsteigerungshebel rechtfertigen den geforderten Kaufpreis-Multiple?

### 4. Resolution & Strategic Approach (R)
Durchführung einer Commercial & Operational Due Diligence mit Fokus auf Kohortenanalyse (NRR/GRR), Unit Economics und Post-Merger-Synergien.""",
        "mece": """**1. Umsatzqualität & Kundenbasis (Top-Line Resilience)**
1.1 **ARR & Cohort Health:** Analyse der Net Retention Rate (NRR > 105 %) und Churn-Raten nach Kundensegmenten.
1.2 **Kundenkonzentration:** Klumpenrisiken durch Abhängigkeit von einzelnen Großmandanten.
1.3 **Pricing Power:** Potenzial für künftige Preiserhöhungen bei Vertragsverlängerungen.

**2. Unit Economics & Profitabilität (EBITDA Quality)**
2.1 **CAC-Payback:** Verhältnis von Customer Lifetime Value (LTV) zu Kundenakquisitionskosten (CAC).
2.2 **R&D Capitalization:** Prüfung aktivierter Eigenleistungen auf kosmetische EBITDA-Bereinigungen.
2.3 **Gross Margin:** Stabilität der Hosting- und Customer-Support-Kosten bei Skalierung.

**3. Synergiepotenziale & Wertsteigerung (Value Creation)**
3.1 **Cross-Selling:** Vertrieb der Software über das bestehende Portfolio-Netzwerk des Investors.
3.2 **SG&A-Synergien:** Zusammenlegung von Holding-, Finance- und Legal-Funktionen.""",
        "hypothesis": """| Bereich | Primäre Hypothese | KPI / Ziel-Benchmark | Erwarteter EBIT-Hebel |
| :--- | :--- | :--- | :--- |
| **ARR-Qualität** | Hohe Net Retention Rate (> 110 %) bestätigt starke Preissetzungsmacht bei Bestandskunden. | **NRR > 110 %** | **Valuation Safety** |
| **Churn-Risiko** | Gezieltes Onboarding reduziert Logo Churn im KMU-Segment deutlich. | **Logo Churn < 5 % p.a.** | **+0,80 Mio. € ARR** |
| **G&A-Synergien** | Konsolidierung von Holdingstrukturen hebt unmittelbare Synergien. | **G&A-Quote < 12 %** | **+1,10 Mio. € EBITDA** |"""
    },
    "Market Entry": {
        "analysis": """### 1. Situation (S)
Ein führender E-Commerce-Händler für Premium-Konsumgüter plant die geografische Expansion in zwei neue europäische Kernmärkte.

### 2. Complication (C)
Ein Investitionsbudget von 2,0 Mio. € steht zur Verfügung, jedoch bergen hohe lokale Kundenakquisitionskosten (CAC) und etablierte Lokalwettbewerber das Risiko verlängerter Amortisationszeiten.

### 3. Key Question (KQ)
Über welche Go-to-Market Strategie und Vertriebskanäle kann der Markteintritt mit einer Amortisationsdauer (Payback Period) von unter 12 Monaten realisiert werden?

### 4. Resolution & Strategic Approach (R)
Evaluierung von Markteintrittsbarrieren, Testen von Performance-Marketing-Kanälen und Lokalisierung der Logistik- und Checkout-Prozesse.""",
        "mece": """**1. Marktattraktivität & Wettbewerbsumfeld (Market Attractiveness)**
1.1 **Marktvolumen & Wachstum:** Zielgruppenpotenzial im Premium-Konsumgütersegment.
1.2 **Wettbewerbsintensität:** Preispunkte, Markenloyalität und Marktanteile etablierter lokaler Player.
1.3 **Regulatorik & Steuern:** Lokale Verbraucherschutzgesetze, VAT-Registrierung und Kennzeichnungspflichten.

**2. Go-To-Market & Kundenakquisition (Commercial Strategy)**
2.1 **Marketing-Effizienz:** Erwartete Customer Acquisition Costs (CAC) nach Kanälen (Social, Search, Influencer).
2.2 **Lokalisierung:** Übersetzung, lokale Währung/Payment-Methoden und Vertrauenssiegel im Checkout.
2.3 **Sortimentsstrategie:** Anpassung des Produkt-Portfolios an lokale Kundenpräferenzen.

**3. Operative Abwicklung & Fulfillment (Operations)**
3.1 **Logistik & Versand:** Anbindung lokaler Carrier für schnelle Lieferzeiten (< 48 Stunden).
3.2 **Retourenmanagement:** Einrichtung lokaler Retouren-Hubs zur Reduktion der Rücksendekosten.""",
        "hypothesis": """| Bereich | Primäre Hypothese | KPI / Ziel-Benchmark | Erwarteter EBIT-Hebel |
| :--- | :--- | :--- | :--- |
| **Marketing-CAC** | Lokales Influencer- & Performance-Marketing hält CAC unter der Profitabilitätsschwelle. | **CAC < 35 € / Neukunde** | **Payback < 9 Monate** |
| **Checkout-Conversion** | Integration lokaler Zahlungsarten steigert die Checkout-Conversion-Rate um 18 %. | **Conversion Rate > 3,2 %** | **+0,30 Mio. €** |
| **Retouren-Management** | Lokales Retourenlager senkt Logistik-Rückabwicklungskosten spürbar. | **Retourenkosten -25 %** | **+0,35 Mio. €** |"""
    }
}

PRECACHED_EN = {
    "General Profitability": {
        "analysis": """### 1. Situation (S)
The client is a mid-sized industrial manufacturing company with stable annual revenues of **€45.0M**. Historically, the business achieved a healthy EBIT margin of **11.5%** (~€5.18M EBIT).

### 2. Complication (C)
Over the past 18 months, the EBIT margin experienced a severe drop of **8.3 percentage points to 3.2%** (~€1.44M EBIT). Because top-line revenue remained constant, the total operational loss of **~€3.74M** stems entirely from cost inflation and adverse price/product mix shifts.

### 3. Key Question (KQ)
Which specific cost and mix drivers eroded profitability, and what strategic action plan will sustainably elevate the EBIT margin back above the target benchmark of **> 8.0%** (> €3.60M EBIT)?

### 4. Resolution & Strategic Approach (R)
Reaching the target margin requires a minimum EBIT expansion of **€2.16M**. This necessitates a root-cause decomposition along the profitability tree and the implementation of a two-phased performance improvement program.""",
        "mece": """**1. Revenue Quality & Price Realization (Top-Line & Mix Levers)**
1.1 **Pricing & Indexation:** Inadequate pass-through of inflated input costs to end customers.
1.2 **Portfolio Mix Shift:** Unfavorable volume migration from high-margin specialty items to low-margin standard products.
1.3 **Commercial Terms:** Excessive discounting structures and unfavorable freight allowances across Key Accounts.

**2. Variable Cost of Goods Sold (COGS / Direct Costs)**
2.1 **Procurement & Raw Materials:** Raw material price surges without structured strategic sourcing countermeasures.
2.2 **Manufacturing Efficiency:** Declining OEE metrics, rising scrap rates, and unoptimized overtime shifts.
2.3 **Logistics & Energy:** Escalating outbound freight rates and energy intensity per production unit.

**3. Indirect OPEX & Overhead (Indirect Costs)**
3.1 **SG&A Creep:** Uncontrolled expansion of administrative and commercial fixed overheads.
3.2 **Maintenance & R&D:** Escalating repair expenses for aging assets and non-prioritized development projects.""",
        "hypothesis": """| Focus Area | Primary Working Hypothesis | KPI / Target Benchmark | Expected EBIT Impact |
| :--- | :--- | :--- | :--- |
| **Pricing & Mix** | Targeted price adjustments (+3.5%) and raw material indexing clauses stabilize gross margins. | **Price Realization Rate > 85%** | **+€0.90M** |
| **COGS & Sourcing** | Re-tendering Top-20 supplier contracts and scrap reduction lower direct variable costs. | **Material Cost Ratio < 56.1%** | **+€0.85M** |
| **SG&A / Overhead** | Discretionary spending freeze and indirect cost containment halt fixed cost creep. | **SG&A Ratio < 17.6%** | **+€0.41M** |"""
    },
    "Cost Reduction": {
        "analysis": """### 1. Situation (S)
The client is an international logistics service provider operating extensive transportation fleets and warehousing hubs across Europe.

### 2. Complication (C)
Surging fleet fuel costs, wage inflation, and operational bottlenecks reduced annual operating profit by **€4.5M** year-over-year.

### 3. Key Question (KQ)
Through which operational levers can total OPEX be sustainably reduced by at least **15%** without compromising service quality and delivery SLA compliance?

### 4. Resolution & Strategic Approach (R)
Execution of a comprehensive cost reduction program focused on fleet telematics, warehouse automation, and indirect procurement optimization.""",
        "mece": """**1. Fleet & Fuel Efficiency (Direct OPEX)**
1.1 **Route Optimization:** Reduction of empty mileage using AI-assisted dispatching and telematics.
1.2 **Fuel Management:** Renegotiating fuel card terms and driver training programs for eco-driving.
1.3 **Maintenance:** Optimization of service cycles and fleet modernization to cut repair expenses.

**2. Warehouse & Infrastructure Logistics (Facility & Labor)**
2.1 **Process Automation:** Implementation of barcode scanning and picking route optimization.
2.2 **Footprint Rationalization:** Consolidation of underutilized warehouse space to slash fixed facility overhead.
2.3 **Shift Scheduling:** Flexible labor scheduling to eliminate costly overtime premiums.

**3. Indirect Procurement & Overhead (Indirect OPEX)**
3.1 **Supplier Consolidation:** Re-tendering Top-15 indirect vendor contracts (packaging, IT, cleaning).
3.2 **Administrative Digitization:** Automating freight documentation processing to reduce SG&A ratio.""",
        "hypothesis": """| Focus Area | Primary Working Hypothesis | KPI / Target Benchmark | Expected EBIT Impact |
| :--- | :--- | :--- | :--- |
| **Fleet Optimization** | AI dispatching and driver training lower fuel consumption per ton-km by 8%. | **Fuel Efficiency +8%** | **+€1.80M** |
| **Warehouse Operations** | Space consolidation enables closure of one redundant satellite facility. | **Facility Productivity +12%** | **+€1.50M** |
| **Procurement & Overhead** | Aggregating indirect spend drives vendor discount renegotiations. | **Procurement Savings > 10%** | **+€1.20M** |"""
    },
    "M&A Due Diligence": {
        "analysis": """### 1. Situation (S)
A Private Equity investor is evaluating the majority acquisition of a high-growth B2B SaaS company.

### 2. Complication (C)
Prior to launching detailed Commercial Due Diligence, key uncertainties remain regarding ARR run-rate durability, customer retention health, and achievable post-merger synergy potential.

### 3. Key Question (KQ)
How resilient is the target's organic revenue growth, and which operational value creation levers justify the requested valuation multiple?

### 4. Resolution & Strategic Approach (R)
Execution of a Commercial & Operational Due Diligence framework focused on cohort analysis (NRR/GRR), unit economics, and synergy quantification.""",
        "mece": """**1. Revenue Quality & Customer Base (Top-Line Resilience)**
1.1 **ARR & Cohort Health:** Evaluation of Net Retention Rate (NRR > 105%) and churn dynamics across customer tiers.
1.2 **Customer Concentration:** Concentration risk regarding key account dependencies.
1.3 **Pricing Power:** Potential for contract price uplifts upon upcoming renewals.

**2. Unit Economics & Profitability (EBITDA Quality)**
2.1 **CAC Payback & LTV:** Ratio of Customer Lifetime Value (LTV) to Customer Acquisition Cost (CAC).
2.2 **R&D Capitalization:** Scrutiny of capitalized software development costs to ensure EBITDA validity.
2.3 **Gross Margin Stabilities:** Stabilities of cloud hosting and customer support expense scaling.

**3. Synergy Potential & Value Creation (Post-Merger Value)**
3.1 **Cross-Selling:** Distribution of software products across the investor's portfolio network.
3.2 **SG&A Synergies:** Consolidation of holding, finance, and legal overhead functions.""",
        "hypothesis": """| Focus Area | Primary Working Hypothesis | KPI / Target Benchmark | Expected EBIT Impact |
| :--- | :--- | :--- | :--- |
| **ARR Resilience** | High Net Retention Rate (> 110%) confirms strong pricing power among enterprise clients. | **NRR > 110%** | **Valuation Safety** |
| **Churn Mitigation** | Enhanced onboarding processes significantly reduce SME logo churn. | **Logo Churn < 5% p.a.** | **+€0.80M ARR** |
| **G&A Synergies** | Holding structure consolidation unlocks immediate operational synergies. | **G&A Ratio < 12%** | **+€1.10M EBITDA** |"""
    },
    "Market Entry": {
        "analysis": """### 1. Situation (S)
A leading direct-to-consumer premium consumer goods e-commerce retailer plans geographic expansion into two new European core markets.

### 2. Complication (C)
An investment budget of €2.0M is allocated; however, elevated local Customer Acquisition Costs (CAC) and established incumbents pose risks of extended payback timelines.

### 3. Key Question (KQ)
Which Go-To-Market strategy and commercial channel mix will achieve market entry with a Customer CAC Payback Period of under 12 months?

### 4. Resolution & Strategic Approach (R)
Assessment of market entry barriers, performance channel testing, and localization of fulfillment and checkout workflows.""",
        "mece": """**1. Market Attractiveness & Competitive Landscape (Market Attractiveness)**
1.1 **Market Size & Growth:** Target demographic size within the premium consumer segment.
1.2 **Competitive Intensity:** Price positioning, brand loyalty, and market share of local incumbents.
1.3 **Regulatory & Tax:** Local consumer protection compliance, VAT registration, and labeling mandates.

**2. Go-To-Market & Commercial Strategy (Commercial Strategy)**
2.1 **Marketing Efficiency:** Expected Customer Acquisition Costs (CAC) across channels (Social, Search, Influencers).
2.2 **Localization:** Translation, localized payment methods, and trust badges at checkout.
2.3 **Sortimentsstrategie:** Tailoring product catalog bundles to local consumer preferences.

**3. Operational Execution & Fulfillment (Operations)**
3.1 **Logistics & Delivery:** Integration of local carriers for rapid delivery SLAs (< 48 hours).
3.2 **Returns Management:** Establishing local return processing hubs to optimize logistics reverse-processing expenses.""",
        "hypothesis": """| Focus Area | Primary Working Hypothesis | KPI / Target Benchmark | Expected EBIT Impact |
| :--- | :--- | :--- | :--- |
| **Marketing CAC** | Localized influencer and search campaigns maintain CAC below profitability thresholds. | **CAC < €35 / New Customer** | **Payback < 9 Months** |
| **Checkout Conversion** | Integrating local payment methods increases checkout conversion rate by 18%. | **Conversion Rate > 3.2%** | **+€0.30M** |
| **Returns Efficiency** | Local return hub establishment lowers logistics reverse-processing expenses. | **Return Costs -25%** | **+€0.35M** |"""
    }
}

if run_analysis:
    if not case_input.strip():
        st.warning(ui_warning)
    else:
        def normalize_text(txt):
            if not txt:
                return ""
            return re.sub(r'\r\n', '\n', txt).strip()

        clean_case_input = normalize_text(case_input)

        all_demo_map = {}
        for fw, txt in DEMO_CASES_DE.items():
            all_demo_map[normalize_text(txt)] = fw
        for fw, txt in DEMO_CASES_EN.items():
            all_demo_map[normalize_text(txt)] = fw

        is_default_demo = clean_case_input in all_demo_map
        matched_framework = all_demo_map.get(clean_case_input, framework_focus)

        if is_default_demo and not user_key.strip():
            precached_dict = PRECACHED_EN if language == "English" else PRECACHED_DE
            precached = precached_dict.get(matched_framework, precached_dict["General Profitability"])

            st.session_state["out_analysis"] = precached["analysis"]
            st.session_state["out_mece"] = precached["mece"]
            st.session_state["out_hypothesis"] = precached["hypothesis"]
            st.session_state["has_analysis"] = True
            st.info(ui_demo_info)
        else:
            with st.status(ui_status_start, expanded=True) as status:
                analyzer = Agent(
                    role="Senior Strategy Consultant",
                    goal=f"Erstelle eine präzise Executive Summary und Situation-Complication-Resolution (SCR) Analyse mit Fokus auf {framework_focus}.",
                    backstory="Erfahrener Strategy Consultant mit Spezialisierung auf prägnante Problemsynthesen und strukturierte Analysen.",
                    llm=gemini_llm,
                    max_iter=1,
                    verbose=False
                )

                structurer = Agent(
                    role="MECE Framework Architect",
                    goal="Erstelle eine 100% überschneidungsfreie und vollständige Problemstruktur (MECE Issue Tree) in sauberem Markdown.",
                    backstory="Spezialist für logische Problemzerlegung. Achtet strikt auf Vollständigkeit und Überschneidungsfreiheit.",
                    llm=gemini_llm,
                    max_iter=1,
                    verbose=False
                )

                hypothesis_builder = Agent(
                    role="Strategy & Hypothesis Lead",
                    goal="Entwickle 3 quantifizierbare Arbeitshypothesen inklusive einer strukturierten KPI-Validierungsmatrix.",
                    backstory="Experte für datengestützte Unternehmensanalysen, Performance Improvement und KPI-Konzepte.",
                    llm=gemini_llm,
                    max_iter=1,
                    verbose=False
                )

                t1 = Task(
                    description=f"Analysiere folgendes Case-Briefing unter Berücksichtigung von '{framework_focus}':\n\n{case_input}\n\nErstelle eine strukturierte Executive Summary im SCR-Format.\n\n{FORMATTING_RULES}",
                    expected_output=f"Strukturierte SCR-Analyse auf {language}.",
                    agent=analyzer
                )

                t2 = Task(
                    description=f"Basierend auf Task 1: Erstelle einen vollständigen MECE Issue Tree auf {language}.\n\n{FORMATTING_RULES}",
                    expected_output=f"Ein übersichtlicher MECE Issue Tree auf {language}.",
                    agent=structurer
                )

                t3 = Task(
                    description=f"Basierend auf Task 2: Formuliere genau 3 priorisierte Arbeitshypothesen auf {language} inklusive KPI-Matrix.\n\n{FORMATTING_RULES}",
                    expected_output=f"3 Hypothesen mit KPI-Validierungsmatrix als Markdown-Tabelle auf {language}.",
                    agent=hypothesis_builder
                )

                crew = Crew(
                    agents=[analyzer, structurer, hypothesis_builder],
                    tasks=[t1, t2, t3],
                    process=Process.sequential
                )

                try:
                    result = crew.kickoff()
                    status.update(label=ui_status_done, state="complete", expanded=False)
                    st.session_state["out_analysis"] = result.tasks_output[0].raw
                    st.session_state["out_mece"] = result.tasks_output[1].raw
                    st.session_state["out_hypothesis"] = result.tasks_output[2].raw
                    st.session_state["has_analysis"] = True
                except Exception as e:
                    err_str = str(e)
                    status.update(label="⚠️ Schnittstellen-Fehler", state="error", expanded=False)
                    
                    if "503" in err_str or "UNAVAILABLE" in err_str or "high demand" in err_str:
                        st.warning("⚠️ **Google AI Server-Engpass (503):** Die Google Gemini-Server sind derzeit weltweit stark ausgelastet. Bitte versuchen Sie es in wenigen Sekunden erneut oder laden Sie den integrierten Demo-Case für eine Vorschau.")
                    elif "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                        st.error("⚠️ Das kostenlose Tageskontingent der API ist vorübergehend erschöpft. Bitte tragen Sie in der linken Seitenleiste einen eigenen kostenlosen Gemini API-Key ein oder versuchen Sie es in wenigen Minuten erneut.")
                    else:
                        st.error(f"Fehler bei der Analyse: {e}")
                    st.stop()

# 9. Ergebnisanzeige & Export
if st.session_state.get("has_analysis", False):
    out_analysis = st.session_state["out_analysis"]
    out_mece = st.session_state["out_mece"]
    out_hypothesis = st.session_state["out_hypothesis"]

    st.subheader(ui_kpi_hdr)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label=ui_m1_label, value=framework_focus)
    with col2:
        st.metric(label=ui_m2_label, value=ui_m2_val, delta=ui_m2_delta)
    with col3:
        st.metric(label=ui_m3_label, value=ui_m3_val, delta=ui_m3_delta)

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([ui_tab1, ui_tab2, ui_tab3])

    with tab1:
        st.markdown(out_analysis)

    with tab2:
        col_mece_head, col_mece_toggle = st.columns([3, 1])
        with col_mece_toggle:
            view_type = st.radio(ui_view_mode, ["List", "Diagram"], horizontal=True, label_visibility="collapsed")
        
        if view_type == "Diagram":
            try:
                dot_str = generate_mece_dot_string(out_mece)
                st.graphviz_chart(dot_str, use_container_width=True)
            except Exception:
                st.markdown(out_mece)
        else:
            st.markdown(out_mece)

    with tab3:
        st.markdown(out_hypothesis)

    st.markdown("---")

    # Multi-Format Export Sektion
    st.subheader(ui_export_hdr)
    col_fmt, col_btn = st.columns([2, 2])

    with col_fmt:
        export_choice = st.selectbox(
            ui_format_label,
            [
                "Microsoft Excel Matrix (.xlsx)", 
                "HTML Executive Report (.html)", 
                "Microsoft Word (.docx)", 
                "Markdown Raw (.md)"
            ]
        )

    report_title = f"{ui_report_hdr}: {framework_focus}"
    file_base = f"case_report_{framework_focus.lower().replace(' ', '_')}"

    with col_btn:
        st.write(" ")
        st.write(" ")
        if export_choice == "Microsoft Excel Matrix (.xlsx)":
            excel_buffer = create_excel_report(
                report_title, framework_focus, out_analysis, out_mece, out_hypothesis
            )
            st.download_button(
                label="📊 Download Excel Matrix (.xlsx)",
                data=excel_buffer,
                file_name=f"{file_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif export_choice == "HTML Executive Report (.html)":
            html_data = create_html_report(
                report_title, framework_focus, out_analysis, out_mece, out_hypothesis,
                rep_sub, rep_lbl_fw, rep_sec1, rep_sec2, rep_sec3, rep_footer
            )
            st.download_button(
                label="📄 Download HTML Report",
                data=html_data,
                file_name=f"{file_base}.html",
                mime="text/html"
            )
        elif export_choice == "Microsoft Word (.docx)":
            docx_buffer = create_docx_report(
                report_title, framework_focus, out_analysis, out_mece, out_hypothesis,
                rep_lbl_fw, rep_sec1, rep_sec2, rep_sec3, rep_footer
            )
            st.download_button(
                label="📄 Download Word Document",
                data=docx_buffer,
                file_name=f"{file_base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            full_md = f"# {report_title}\n\n## 1. SCR\n{out_analysis}\n\n## 2. MECE\n{out_mece}\n\n## 3. Hypothesen\n{out_hypothesis}"
            st.download_button(
                label="📄 Download Markdown (.md)",
                data=full_md,
                file_name=f"{file_base}.md",
                mime="text/markdown"
            )