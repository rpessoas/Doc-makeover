"""
Padronizador de Documentos — Amélie & Juliette
Backend FastAPI

Instalar dependências:
  pip install fastapi uvicorn python-multipart python-docx pdfplumber anthropic python-pptx openpyxl reportlab

Rodar localmente:
  uvicorn main:app --reload --port 8000

Deploy Railway/Render:
  - Aponte para este arquivo como entry point
  - Adicione a variável de ambiente: ANTHROPIC_API_KEY=sk-...
"""

import os, io, json, re, tempfile
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import anthropic
import docx
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib import colors
import pptx
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
import openpyxl

app = FastAPI(title="Padronizador Amélie & Juliette")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── Configurações de marca e público ────────────────────────────────────────

BRAND_CONFIG = {
    "amelie": {
        "name": "Amélie",
        "persona": "Você representa a marca Amélie. Identidade sofisticada e acolhedora. Linguagem elegante sem ser distante, acessível sem ser informal.",
        "primary_color": (153, 53, 86),       # #993556
        "heading_color": (114, 36, 62),        # #72243E
        "accent_hex": "993556",
    },
    "juliette": {
        "name": "Juliette",
        "persona": "Você representa a marca Juliette. Identidade moderna e orientada a resultado. Linguagem clara, eficiente e profissional.",
        "primary_color": (24, 95, 165),        # #185FA5
        "heading_color": (12, 68, 124),        # #0C447C
        "accent_hex": "185FA5",
    },
}

AUD_CONFIG = {
    "franqueados": {
        "label": "Franqueados",
        "instrucoes": """PÚBLICO: Franqueados (donos de unidade)
Ao organizar a estrutura do documento, priorize o que o franqueado precisa saber para operar bem e tem direito de receber conforme a relação contratual. Não exponha informações estratégicas internas ou qualquer conteúdo que possa gerar interpretações jurídicas indevidas. Use linguagem contratualmente segura: prefira "previsto em contrato", "conforme manual operacional", "sujeito a aprovação". O franqueado deve terminar a leitura sabendo exatamente o que fazer.""",
    },
    "time": {
        "label": "Time interno",
        "instrucoes": """PÚBLICO: Time interno da franqueadora/marca
Ao organizar a estrutura do documento, garanta que o quadro completo fique legível: contexto, impactos por área, dependências, responsáveis, prazos e pendências. A estrutura deve facilitar a visão 360 — nada relevante deve ficar enterrado no texto.""",
    },
    "liderancas": {
        "label": "Lideranças",
        "instrucoes": """PÚBLICO: Lideranças (diretores, gerentes sênior, tomadores de decisão)
Ao organizar a estrutura do documento, aplique Pareto: o que mais impacta deve aparecer primeiro. Priorize: headline de performance ou problema, diagnóstico, riscos, ações e decisões pendentes. O que não for acionável deve ser comprimido ou movido para o final.""",
    },
}

DOC_LABELS = {
    "ficha": "ficha técnica",
    "relatorio": "relatório",
    "checklist": "checklist",
    "apresentacao": "apresentação",
}

INTERVENCAO = """ESCOPO DE ATUAÇÃO — REGRA FUNDAMENTAL:
Você NÃO deve alterar, resumir, expandir ou reescrever o conteúdo do documento. O conteúdo pertence ao autor e deve ser preservado integralmente.

Sua atuação se limita a três frentes:
1. FORMATAÇÃO E DESIGN: aplique a estrutura adequada ao tipo de documento e à identidade da marca (hierarquia de títulos, uso correto de listas, tabelas, destaques e espaçamento).
2. ORGANIZAÇÃO ESTRUTURAL: reorganize seções ou blocos apenas se a ordem atual prejudicar a leitura pelo público-alvo.
3. REVISÃO ORTOGRÁFICA E GRAMATICAL: corrija erros de ortografia, concordância, pontuação e gramática. Não altere o vocabulário ou o estilo do autor além do necessário para a correção."""

# ── Extratores de texto ──────────────────────────────────────────────────────

def extract_text_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    lines = []
    for para in doc.paragraphs:
        style = para.style.name
        text = para.text.strip()
        if not text:
            lines.append("")
            continue
        if "Heading 1" in style:
            lines.append(f"# {text}")
        elif "Heading 2" in style:
            lines.append(f"## {text}")
        elif "Heading 3" in style:
            lines.append(f"### {text}")
        elif para.style.name.startswith("List"):
            lines.append(f"- {text}")
        else:
            lines.append(text)
    return "\n".join(lines)

def extract_text_pdf(file_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
    return "\n\n".join(text_parts)

def extract_text_pptx(file_bytes: bytes) -> str:
    prs = Presentation(io.BytesIO(file_bytes))
    lines = []
    for i, slide in enumerate(prs.slides, 1):
        lines.append(f"## Slide {i}")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                lines.append(shape.text.strip())
    return "\n\n".join(lines)

def extract_text_xlsx(file_bytes: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    lines = []
    for sheet in wb.worksheets:
        lines.append(f"## Aba: {sheet.title}")
        for row in sheet.iter_rows(values_only=True):
            row_vals = [str(c) if c is not None else "" for c in row]
            if any(v.strip() for v in row_vals):
                lines.append(" | ".join(row_vals))
    return "\n".join(lines)

# ── Gerador de documentos formatados ────────────────────────────────────────

def build_docx(markdown_text: str, brand: str) -> bytes:
    cfg = BRAND_CONFIG[brand]
    doc = Document()

    # Estilos base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for h_style, size, bold in [("Heading 1", 16, True), ("Heading 2", 13, True), ("Heading 3", 12, False)]:
        hs = doc.styles[h_style]
        hs.font.name = "Calibri"
        hs.font.size = Pt(size)
        hs.font.bold = bold
        r, g, b = cfg["heading_color"]
        hs.font.color.rgb = RGBColor(r, g, b)

    lines = markdown_text.split("\n")
    for line in lines:
        line = line.rstrip()
        if line.startswith("### "):
            p = doc.add_heading(line[4:], level=3)
        elif line.startswith("## "):
            p = doc.add_heading(line[3:], level=2)
        elif line.startswith("# "):
            p = doc.add_heading(line[2:], level=1)
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(line[2:], style="List Bullet")
        elif re.match(r"^\d+\. ", line):
            p = doc.add_paragraph(re.sub(r"^\d+\. ", "", line), style="List Number")
        elif line.strip() == "" or line.strip() == "---":
            doc.add_paragraph()
        else:
            p = doc.add_paragraph()
            # Negrito inline **texto**
            parts = re.split(r"\*\*(.*?)\*\*", line)
            for i, part in enumerate(parts):
                run = p.add_run(part)
                run.bold = (i % 2 == 1)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def build_pdf(markdown_text: str, brand: str) -> bytes:
    cfg = BRAND_CONFIG[brand]
    r, g, b = cfg["primary_color"]
    brand_color = colors.Color(r/255, g/255, b/255)

    buf = io.BytesIO()
    doc_pdf = SimpleDocTemplate(buf, pagesize=A4,
                                 leftMargin=2.5*cm, rightMargin=2.5*cm,
                                 topMargin=2.5*cm, bottomMargin=2.5*cm)
    styles = getSampleStyleSheet()
    story = []

    h1_style = ParagraphStyle("H1Brand", parent=styles["Heading1"],
                               textColor=brand_color, fontSize=16, spaceAfter=8)
    h2_style = ParagraphStyle("H2Brand", parent=styles["Heading2"],
                               textColor=brand_color, fontSize=13, spaceAfter=6)
    h3_style = ParagraphStyle("H3Brand", parent=styles["Heading3"],
                               textColor=brand_color, fontSize=11, spaceAfter=4)
    body_style = ParagraphStyle("Body", parent=styles["Normal"], fontSize=11, leading=16)
    bullet_style = ParagraphStyle("Bullet", parent=styles["Normal"],
                                   fontSize=11, leading=16, leftIndent=20, bulletIndent=10)

    for line in markdown_text.split("\n"):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
        elif line == "---":
            story.append(HRFlowable(width="100%", thickness=0.5, color=brand_color, spaceAfter=6))
        elif line.startswith("# "):
            story.append(Paragraph(line[2:], h1_style))
        elif line.startswith("## "):
            story.append(Paragraph(line[3:], h2_style))
        elif line.startswith("### "):
            story.append(Paragraph(line[4:], h3_style))
        elif line.startswith("- ") or line.startswith("* "):
            story.append(Paragraph(f"• {line[2:]}", bullet_style))
        elif re.match(r"^\d+\. ", line):
            story.append(Paragraph(line, bullet_style))
        else:
            line_html = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", line)
            story.append(Paragraph(line_html, body_style))

    doc_pdf.build(story)
    buf.seek(0)
    return buf.read()

def build_pptx(markdown_text: str, brand: str) -> bytes:
    cfg = BRAND_CONFIG[brand]
    r, g, b = cfg["primary_color"]
    prs = Presentation()
    prs.slide_width = PptInches(13.33)
    prs.slide_height = PptInches(7.5)

    slide_layout = prs.slide_layouts[1]  # title + content
    current_title = None
    current_body = []

    def flush_slide():
        if current_title is None:
            return
        slide = prs.slides.add_slide(slide_layout)
        title_shape = slide.shapes.title
        title_shape.text = current_title
        tf = title_shape.text_frame.paragraphs[0].runs
        if tf:
            tf[0].font.color.rgb = pptx.util.RGBColor(r, g, b)

        body_shape = slide.placeholders[1]
        body_shape.text = ""
        tf2 = body_shape.text_frame
        tf2.clear()
        for i, line in enumerate(current_body):
            p = tf2.add_paragraph() if i > 0 else tf2.paragraphs[0]
            p.text = line
            p.font.size = PptPt(18)

    lines = markdown_text.split("\n")
    for line in lines:
        line = line.strip()
        if line.startswith("## ") or line.startswith("# "):
            flush_slide()
            current_title = line.lstrip("# ")
            current_body = []
        elif line and not line.startswith("---"):
            current_body.append(line.lstrip("- *").lstrip("0123456789. "))
    flush_slide()

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

def build_xlsx(markdown_text: str, brand: str) -> bytes:
    cfg = BRAND_CONFIG[brand]
    r, g, b = cfg["primary_color"]
    hex_color = cfg["accent_hex"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documento"

    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill("solid", fgColor=hex_color)
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=12)
    body_font = Font(name="Calibri", size=11)

    row = 1
    for line in markdown_text.split("\n"):
        line = line.strip()
        if not line or line == "---":
            row += 1
            continue
        clean = re.sub(r"\*\*(.*?)\*\*", r"\1", line)
        if line.startswith("# ") or line.startswith("## ") or line.startswith("### "):
            cell = ws.cell(row=row, column=1, value=clean.lstrip("# "))
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(wrap_text=True)
        elif line.startswith("- ") or line.startswith("* "):
            cell = ws.cell(row=row, column=1, value=f"• {clean.lstrip('- *')}")
            cell.font = body_font
        else:
            cell = ws.cell(row=row, column=1, value=clean)
            cell.font = body_font
        row += 1

    ws.column_dimensions["A"].width = 80

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Chamada à API do Claude ──────────────────────────────────────────────────

def padronizar_com_claude(text: str, brand: str, audience: str, doc_type: str, reference: str = "") -> dict:
    cfg = BRAND_CONFIG[brand]
    aud_cfg = AUD_CONFIG[audience]
    ref_block = f"\n\nEXEMPLO DE REFERÊNCIA:\n---\n{reference}\n---" if reference.strip() else ""

    prompt = f"""Você é um especialista em padronização de documentos corporativos.

MARCA: {cfg['name']}
{cfg['persona']}

{aud_cfg['instrucoes']}

TIPO DE DOCUMENTO: {DOC_LABELS.get(doc_type, doc_type)}

{INTERVENCAO}{ref_block}

DOCUMENTO PARA PADRONIZAR:
---
{text}
---

Responda SOMENTE com um JSON válido, sem markdown, sem explicações fora do JSON:
{{
  "documento_padronizado": "o documento completo em Markdown com formatação aplicada",
  "alteracoes": "lista do que foi ajustado em formatação, estrutura e ortografia (máximo 5 pontos, cada um em nova linha com '-')"
}}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = response.content[0].text
    clean = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(clean)

# ── Endpoint principal ───────────────────────────────────────────────────────

@app.post("/padronizar")
async def padronizar(
    file: UploadFile = File(...),
    brand: str = Form(...),
    audience: str = Form(...),
    doc_type: str = Form(...),
    reference: str = Form(""),
):
    file_bytes = await file.read()
    filename = file.filename or ""
    ext = filename.rsplit(".", 1)[-1].lower()

    # 1. Extrai texto
    try:
        if ext == "docx":
            text = extract_text_docx(file_bytes)
        elif ext == "pdf":
            text = extract_text_pdf(file_bytes)
        elif ext == "pptx":
            text = extract_text_pptx(file_bytes)
        elif ext in ("xlsx", "xls"):
            text = extract_text_xlsx(file_bytes)
        elif ext == "txt":
            text = file_bytes.decode("utf-8", errors="replace")
        else:
            raise HTTPException(status_code=400, detail=f"Formato .{ext} não suportado.")
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Erro ao ler o arquivo: {str(e)}")

    if not text.strip():
        raise HTTPException(status_code=422, detail="Não foi possível extrair texto do arquivo.")

    # 2. Padroniza com Claude
    try:
        result = padronizar_com_claude(text, brand, audience, doc_type, reference)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro na API do Claude: {str(e)}")

    markdown_out = result.get("documento_padronizado", "")
    alteracoes = result.get("alteracoes", "")

    # 3. Reconstrói no formato original
    try:
        if ext == "docx":
            output_bytes = build_docx(markdown_out, brand)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            out_filename = filename.replace(".docx", "_padronizado.docx")
        elif ext == "pdf":
            output_bytes = build_pdf(markdown_out, brand)
            media_type = "application/pdf"
            out_filename = filename.replace(".pdf", "_padronizado.pdf")
        elif ext == "pptx":
            output_bytes = build_pptx(markdown_out, brand)
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            out_filename = filename.replace(".pptx", "_padronizado.pptx")
        elif ext in ("xlsx", "xls"):
            output_bytes = build_xlsx(markdown_out, brand)
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            out_filename = filename.replace(f".{ext}", "_padronizado.xlsx")
        else:
            output_bytes = markdown_out.encode("utf-8")
            media_type = "text/plain"
            out_filename = filename.replace(".txt", "_padronizado.txt")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar arquivo de saída: {str(e)}")

    headers = {
        "Content-Disposition": f'attachment; filename="{out_filename}"',
        "X-Alteracoes": alteracoes.replace("\n", " | "),
    }
    return StreamingResponse(io.BytesIO(output_bytes), media_type=media_type, headers=headers)


@app.get("/health")
def health():
    return {"status": "ok"}
