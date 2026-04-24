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

import os, io, json, base64, re, tempfile
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

import anthropic
import docx
from logos import LOGO_AMELIE, LOGO_JULIETTE
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

FRONTEND_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Padronizador — Amélie & Juliette</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:system-ui,sans-serif;background:#f8f8f6;color:#1a1a1a;padding:2rem 1rem;min-height:100vh}
  .app{max-width:680px;margin:0 auto}
  h1{font-size:18px;font-weight:500;margin-bottom:1.5rem}

  label{display:block;font-size:13px;color:#666;margin-bottom:6px}
  .section{margin-bottom:1.25rem}

  /* Marca */
  .brand-switcher{display:flex;border:0.5px solid #ccc;border-radius:8px;overflow:hidden;margin-bottom:1.5rem}
  .brand-btn{flex:1;padding:10px;border:none;background:transparent;font-size:14px;font-weight:500;cursor:pointer;transition:all 0.15s;color:#888}
  .brand-btn:first-child{border-right:0.5px solid #ccc}
  .amelie-active .brand-btn.amelie{background:#FDECEA;color:#8B0A1A}
  .juliette-active .brand-btn.juliette{background:#EEF2F4;color:#2E4550}

  /* Público */
  .audience-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:1.25rem}
  .aud-card{border:0.5px solid #ddd;border-radius:8px;padding:10px 12px;cursor:pointer;transition:all 0.15s;background:#fff}
  .aud-card:hover{border-color:#aaa;background:#f8f8f6}
  .aud-card.sel-amelie{border-color:#C8102E;background:#FDECEA}
  .aud-card.sel-juliette{border-color:#506775;background:#EEF2F4}
  .aud-title{font-size:13px;font-weight:500;margin-bottom:3px}
  .aud-desc{font-size:11px;color:#888;line-height:1.4}

  /* Tipo doc */
  .doc-tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:1.25rem}
  .doc-tab{padding:5px 12px;border-radius:8px;border:0.5px solid #ccc;background:transparent;font-size:13px;color:#888;cursor:pointer;transition:all 0.15s}
  .doc-tab:hover{background:#f0f0ee}
  .doc-tab.act-amelie{background:#FDECEA;border-color:#C8102E;color:#8B0A1A;font-weight:500}
  .doc-tab.act-juliette{background:#EEF2F4;border-color:#506775;color:#2E4550;font-weight:500}

  /* Upload */
  .drop-zone{border:1.5px dashed #ccc;border-radius:12px;padding:2rem;text-align:center;cursor:pointer;transition:all 0.2s;background:#fff;position:relative}
  .drop-zone:hover,.drop-zone.drag-over{border-color:#aaa;background:#f8f8f6}
  .drop-zone.amelie-drop{border-color:#C8102E;background:#FDECEA}
  .drop-zone.juliette-drop{border-color:#506775;background:#EEF2F4}
  .drop-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
  .drop-icon{font-size:28px;margin-bottom:8px;line-height:1}
  .drop-main{font-size:14px;font-weight:500;color:#1a1a1a;margin-bottom:4px}
  .drop-sub{font-size:12px;color:#888}
  .file-chosen{font-size:13px;font-weight:500;margin-top:8px;color:#1a1a1a}

  /* Referência */
  .toggle-ref{font-size:12px;color:#888;cursor:pointer;text-decoration:underline;display:inline-block;margin-bottom:8px}
  .ref-area{display:none}
  .ref-area.open{display:block}
  textarea{width:100%;border-radius:8px;border:0.5px solid #ddd;padding:10px 12px;font-size:14px;font-family:inherit;color:#1a1a1a;background:#fff;resize:vertical;line-height:1.6}
  textarea:focus{outline:none;border-color:#aaa}

  /* Nota */
  .scope-note{font-size:12px;color:#999;padding:8px 12px;border-left:2px solid #ddd;margin-bottom:1.25rem;line-height:1.5}

  /* Botão */
  .btn{width:100%;padding:11px;border-radius:8px;border:0.5px solid #ccc;background:transparent;font-size:14px;font-weight:500;color:#1a1a1a;cursor:pointer;transition:all 0.15s}
  .btn:hover{background:#f0f0ee}
  .btn:active{transform:scale(0.98)}
  .btn:disabled{opacity:0.4;cursor:not-allowed}

  /* Progress */
  .progress-wrap{margin-top:1rem;display:none}
  .progress-bar{height:3px;background:#e0e0e0;border-radius:2px;overflow:hidden}
  .progress-fill{height:100%;width:0;border-radius:2px;transition:width 0.4s}
  .amelie-active .progress-fill{background:#993556}
  .juliette-active .progress-fill{background:#185FA5}
  .progress-label{font-size:12px;color:#888;margin-top:6px;text-align:center}

  /* Resultado */
  .divider{height:0.5px;background:#e0e0e0;margin:1.5rem 0}
  .result-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
  .pill{display:inline-flex;padding:3px 10px;border-radius:8px;font-size:12px;font-weight:500;margin-right:6px}
  .pill-amelie{background:#FDECEA;color:#8B0A1A}
  .pill-juliette{background:#EEF2F4;color:#2E4550}
  .pill-aud{background:#f0f0ee;color:#666}
  .result-label{font-size:13px;font-weight:500;color:#666}
  .download-btn{padding:6px 14px;border-radius:8px;border:0.5px solid #ccc;background:#fff;font-size:13px;font-weight:500;cursor:pointer;text-decoration:none;color:#1a1a1a;display:inline-flex;align-items:center;gap:6px}
  .download-btn:hover{background:#f0f0ee}
  .alteracoes-box{border:0.5px solid #ddd;border-radius:12px;padding:1.25rem;background:#f4f4f2;font-size:13px;line-height:1.7;color:#444;white-space:pre-wrap;min-height:60px}

  /* Config URL */
  .config-bar{background:#fff;border:0.5px solid #ddd;border-radius:8px;padding:10px 12px;margin-bottom:1.5rem;display:flex;gap:8px;align-items:center}
  .config-bar input{flex:1;border:none;font-size:12px;color:#666;outline:none;background:transparent}
  .config-bar label{font-size:11px;color:#aaa;white-space:nowrap;margin:0}
  .mode-badge{font-size:11px;padding:2px 8px;border-radius:6px;white-space:nowrap}
  .mode-api{background:#E6F1FB;color:#185FA5}
  .mode-demo{background:#FBEAF0;color:#993556}

  .status{font-size:13px;color:#888;text-align:center;padding:4px}
</style>
</head>
<body>
<div class="app amelie-active" id="app">
  <h1>Padronizador de documentos</h1>



  <!-- Marca -->
  <div class="section">
    <label>Marca</label>
    <div class="brand-switcher">
      <button class="brand-btn amelie" onclick="setBrand('amelie')">Amélie</button>
      <button class="brand-btn juliette" onclick="setBrand('juliette')">Juliette</button>
    </div>
  </div>

  <!-- Público -->
  <div class="section">
    <label>Público-alvo</label>
    <div class="audience-grid">
      <div class="aud-card sel-amelie" data-aud="franqueados" onclick="setAud(this)">
        <div class="aud-title">Franqueados</div>
        <div class="aud-desc">O que precisam saber e podem receber</div>
      </div>
      <div class="aud-card" data-aud="time" onclick="setAud(this)">
        <div class="aud-title">Time interno</div>
        <div class="aud-desc">Completo, detalhado, visão 360</div>
      </div>
      <div class="aud-card" data-aud="liderancas" onclick="setAud(this)">
        <div class="aud-title">Lideranças</div>
        <div class="aud-desc">Performance, riscos, pareto do que importa</div>
      </div>
    </div>
  </div>

  <!-- Tipo -->
  <div class="section">
    <label>Tipo de documento</label>
    <div class="doc-tabs">
      <button class="doc-tab act-amelie" data-type="ficha" onclick="setDoc(this)">Ficha técnica</button>
      <button class="doc-tab" data-type="relatorio" onclick="setDoc(this)">Relatório</button>
      <button class="doc-tab" data-type="checklist" onclick="setDoc(this)">Checklist</button>
      <button class="doc-tab" data-type="apresentacao" onclick="setDoc(this)">Apresentação</button>
    </div>
  </div>

  <!-- Upload -->
  <div class="section">
    <label>Documento para padronizar</label>
    <div class="drop-zone" id="drop-zone">
      <input type="file" id="file-input" accept=".docx,.pdf,.pptx,.xlsx,.xls,.txt" onchange="onFileChange(this)">
      <div class="drop-icon">📄</div>
      <div class="drop-main">Clique ou arraste o arquivo aqui</div>
      <div class="drop-sub">.docx · .pdf · .pptx · .xlsx · .txt</div>
      <div class="file-chosen" id="file-name" style="display:none"></div>
    </div>
  </div>

  <!-- Referência -->
  <div class="section">
    <span class="toggle-ref" id="toggle-ref" onclick="toggleRef()">+ Adicionar documento de referência de estilo (recomendado)</span>
    <div class="ref-area" id="ref-area">
      <label>Faça upload de um documento bem feito da marca — o Claude vai usar como modelo de estilo e formatação</label>
      <div class="drop-zone" id="ref-drop-zone" style="padding:1rem;margin-top:6px;">
        <input type="file" id="ref-file-input" accept=".docx,.pdf,.txt" onchange="onRefFileChange(this)">
        <div class="drop-main" style="font-size:13px;">Clique ou arraste o documento de referência</div>
        <div class="drop-sub">.docx · .pdf · .txt</div>
        <div class="file-chosen" id="ref-file-name" style="display:none"></div>
      </div>
    </div>
  </div>

  <p class="scope-note">O conteúdo não será alterado. A padronização cobre formatação conforme a marca, estrutura do tipo de documento e revisão ortográfica. O arquivo será devolvido no mesmo formato enviado.</p>

  <button class="btn" id="run-btn" onclick="runPadronizacao()">Padronizar e baixar documento ↗</button>

  <div class="progress-wrap" id="progress-wrap">
    <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
    <div class="progress-label" id="progress-label">Lendo documento...</div>
  </div>

  <div id="result-section" style="display:none">
    <div class="divider"></div>
    <div class="result-header">
      <span class="result-label">
        <span class="pill" id="out-brand-pill">Amélie</span>
        <span class="pill pill-aud" id="out-aud-pill">Franqueados</span>
      </span>
      <a class="download-btn" id="download-link" href="#" download>⬇ Baixar arquivo</a>
    </div>
    <div style="margin-top:10px">
      <div style="font-size:12px;color:#888;margin-bottom:6px">O que foi ajustado</div>
      <div class="alteracoes-box" id="alteracoes-box"></div>
    </div>
  </div>

  <div class="status" id="status"></div>
</div>

<script>
const EXT_MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pdf: 'application/pdf',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  txt: 'text/plain',
};

let brand = 'amelie', aud = 'franqueados', docType = 'ficha';
let selectedFile = null;
let selectedRefFile = null;

// ── Drag & drop ──────────────────────────────────────────────────────────────
const dropZone = document.getElementById('drop-zone');
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) setFile(f);
});

function onFileChange(input) { if (input.files[0]) setFile(input.files[0]); }

function setFile(f) {
  selectedFile = f;
  const nameEl = document.getElementById('file-name');
  nameEl.textContent = f.name;
  nameEl.style.display = 'block';
  const b = brand === 'amelie' ? 'amelie-drop' : 'juliette-drop';
  dropZone.className = `drop-zone ${b}`;
}

// ── Controles ────────────────────────────────────────────────────────────────
function setBrand(b) {
  brand = b;
  document.getElementById('app').className = 'app ' + (b === 'amelie' ? 'amelie-active' : 'juliette-active');
  document.querySelectorAll('.aud-card').forEach(c => {
    c.className = 'aud-card' + (c.dataset.aud === aud ? ` sel-${b}` : '');
  });
  document.querySelectorAll('.doc-tab').forEach(t => {
    t.className = 'doc-tab' + (t.dataset.type === docType ? ` act-${b}` : '');
  });
  if (selectedFile) dropZone.className = `drop-zone ${b}-drop`;
}

function setAud(el) {
  aud = el.dataset.aud;
  document.querySelectorAll('.aud-card').forEach(c => {
    c.className = 'aud-card' + (c.dataset.aud === aud ? ` sel-${brand}` : '');
  });
}

function setDoc(el) {
  docType = el.dataset.type;
  document.querySelectorAll('.doc-tab').forEach(t => {
    t.className = 'doc-tab' + (t.dataset.type === docType ? ` act-${brand}` : '');
  });
}

function onRefFileChange(input) {
  if (input.files[0]) {
    selectedRefFile = input.files[0];
    const nameEl = document.getElementById('ref-file-name');
    nameEl.textContent = input.files[0].name;
    nameEl.style.display = 'block';
  }
}

function toggleRef() {
  const area = document.getElementById('ref-area');
  const open = area.classList.toggle('open');
  document.getElementById('toggle-ref').textContent = open ? '- Ocultar referência' : '+ Adicionar exemplo de referência (recomendado)';
}


// ── Progress ─────────────────────────────────────────────────────────────────
function setProgress(pct, label) {
  document.getElementById('progress-wrap').style.display = 'block';
  document.getElementById('progress-fill').style.width = pct + '%';
  document.getElementById('progress-label').textContent = label;
}

function hideProgress() {
  document.getElementById('progress-wrap').style.display = 'none';
  document.getElementById('progress-fill').style.width = '0';
}

// ── Labels ───────────────────────────────────────────────────────────────────
const audLabels = { franqueados: 'Franqueados', time: 'Time interno', liderancas: 'Lideranças' };
const brandNames = { amelie: 'Amélie', juliette: 'Juliette' };
const pillClasses = { amelie: 'pill pill-amelie', juliette: 'pill pill-juliette' };

// ── Principal ─────────────────────────────────────────────────────────────────
async function runPadronizacao() {
  // Lê referência se tiver arquivo
  let refText = '';
  if (selectedRefFile) {
    try {
      const refExt = selectedRefFile.name.split('.').pop().toLowerCase();
      if (refExt === 'txt') {
        refText = await selectedRefFile.text();
      } else if (refExt === 'docx' && typeof mammoth !== 'undefined') {
        const ab = await selectedRefFile.arrayBuffer();
        const r = await mammoth.extractRawText({ arrayBuffer: ab });
        refText = r.value;
      } else if (refExt === 'pdf' && typeof pdfjsLib !== 'undefined') {
        const ab = await selectedRefFile.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
        const pages = [];
        for (let i = 1; i <= Math.min(pdf.numPages, 3); i++) {
          const page = await pdf.getPage(i);
          const tc = await page.getTextContent();
          pages.push(tc.items.map(s => s.str).join(' '));
        }
        refText = pages.join('\\n\\n');
      }
    } catch(e) { console.warn('ref read error', e); }
  }

  if (!selectedFile) { document.getElementById('status').textContent = 'Selecione um arquivo antes de continuar.'; return; }

  const apiUrl = window.location.origin;
  const btn = document.getElementById('run-btn');
  btn.disabled = true;
  document.getElementById('status').textContent = '';
  document.getElementById('result-section').style.display = 'none';

  if (apiUrl) {
    // ── Modo backend real ──────────────────────────────────────────────────
    setProgress(20, 'Enviando arquivo...');
    try {
      const form = new FormData();
      form.append('file', selectedFile);
      form.append('brand', brand);
      form.append('audience', aud);
      form.append('doc_type', docType);
      form.append('reference', refText);

      setProgress(50, 'Padronizando com Claude...');
      const resp = await fetch('/padronizar', { method: 'POST', body: form });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ detail: resp.statusText }));
        throw new Error(err.detail || 'Erro desconhecido');
      }

      setProgress(85, 'Gerando arquivo...');
      const blob = await resp.blob();
      const alteracoes = decodeURIComponent(resp.headers.get('X-Alteracoes') || '').replace(/ \\| /g, '\\n');
      const ext = selectedFile.name.rsplit ? selectedFile.name.split('.').pop() : selectedFile.name.split('.').pop();
      const outName = selectedFile.name.replace(`.${ext}`, `_padronizado.${ext}`);

      const url = URL.createObjectURL(blob);
      const link = document.getElementById('download-link');
      link.href = url;
      link.download = outName;

      showResult(alteracoes);
      setProgress(100, 'Pronto!');
      setTimeout(hideProgress, 1000);
    } catch (e) {
      hideProgress();
      document.getElementById('status').textContent = 'Erro: ' + e.message;
    }
  } else if (false) {
    setProgress(20, 'Lendo arquivo...');
    let textContent = '';
    const ext = selectedFile.name.split('.').pop().toLowerCase();

    try {
      if (ext === 'txt') {
        textContent = await selectedFile.text();
      } else if (ext === 'docx') {
        // Usa mammoth via CDN (carregado abaixo)
        const ab = await selectedFile.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer: ab });
        textContent = result.value;
      } else if (ext === 'pdf') {
        const ab = await selectedFile.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
        const pages = [];
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const tc = await page.getTextContent();
          pages.push(tc.items.map(s => s.str).join(' '));
        }
        textContent = pages.join('\\n\\n');
      } else {
        throw new Error(`Para .${ext}, conecte o backend para suporte completo.`);
      }
    } catch (e) {
      hideProgress();
      document.getElementById('status').textContent = 'Erro ao ler arquivo: ' + e.message;
      btn.disabled = false;
      return;
    }

    if (!textContent.trim()) {
      hideProgress();
      document.getElementById('status').textContent = 'Não foi possível extrair texto do arquivo.';
      btn.disabled = false;
      return;
    }

    setProgress(50, 'Padronizando com Claude...');

    const audInstrucoes = {
      franqueados: 'Ao organizar a estrutura, priorize o que o franqueado precisa saber para operar. Linguagem contratualmente segura: "previsto em contrato", "conforme manual operacional".',
      time: 'Garanta que o quadro completo fique legível: contexto, impactos, dependências, responsáveis, prazos e pendências. Visão 360.',
      liderancas: 'Aplique Pareto: o que mais impacta aparece primeiro. Headline → diagnóstico → riscos → ações → decisões pendentes.',
    };
    const brandPersona = {
      amelie: 'Marca Amélie: sofisticada e acolhedora. Linguagem elegante sem ser distante.',
      juliette: 'Marca Juliette: moderna e orientada a resultado. Linguagem clara e eficiente.',
    };
    const docLabels = { ficha: 'ficha técnica', relatorio: 'relatório', checklist: 'checklist', apresentacao: 'apresentação' };

    const prompt = `Você é um especialista em padronização de documentos corporativos.
${brandPersona[brand]}
PÚBLICO: ${audLabels[aud]} — ${audInstrucoes[aud]}
TIPO: ${docLabels[docType]}
ESCOPO: preserve o conteúdo integralmente. Atue apenas em formatação/design da marca, organização estrutural e revisão ortográfica.
${document.getElementById('ref-input').value.trim() ? 'REFERÊNCIA:\\n---\\n' + document.getElementById('ref-input').value.trim() + '\\n---\\n' : ''}
DOCUMENTO:
---
${textContent}
---
Responda SOMENTE com JSON válido:
{"documento_padronizado":"markdown completo","alteracoes":"lista de ajustes (máx 5, cada um com '-' em nova linha)"}`;

    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 4000, messages: [{ role: 'user', content: prompt }] })
      });
      const data = await resp.json();
      const raw = data.content?.find(b => b.type === 'text')?.text || '';
      const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());

      setProgress(85, 'Gerando arquivo para download...');

      // Gera .txt com o markdown (no browser, sem backend)
      const outputText = parsed.documento_padronizado;
      const blob = new Blob([outputText], { type: 'text/plain' });
      const outExt = ext === 'txt' ? 'txt' : 'md';
      const outName = selectedFile.name.replace(`.${ext}`, `_padronizado.${outExt}`);
      const url = URL.createObjectURL(blob);
      const link = document.getElementById('download-link');
      link.href = url;
      link.download = outName;
      link.textContent = `⬇ Baixar (${outExt.toUpperCase()})`;

      showResult(parsed.alteracoes);
      setProgress(100, 'Pronto! (modo demo: arquivo em Markdown — conecte o backend para o formato original)');
      setTimeout(hideProgress, 3000);
    } catch (e) {
      hideProgress();
      document.getElementById('status').textContent = 'Erro: ' + e.message;
    }
  }

  btn.disabled = false;
}

function showResult(alteracoes) {
  const pill = document.getElementById('out-brand-pill');
  pill.textContent = brandNames[brand];
  pill.className = pillClasses[brand];
  document.getElementById('out-aud-pill').textContent = audLabels[aud];
  document.getElementById('alteracoes-box').textContent = alteracoes;
  document.getElementById('result-section').style.display = 'block';
}
</script>

<!-- Bibliotecas para modo demo (leitura de docx e pdf no browser) -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
  if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    return HTMLResponse(content=FRONTEND_HTML)

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
        "persona": """Você representa a marca Amélie.

IDENTIDADE: sofisticada, acolhedora, gastronômica. Linguagem elegante sem ser distante, acessível sem ser informal.

PADRÃO VISUAL (extraído de documentos reais da marca):
- Fundo branco/creme limpo
- Cor primária: vermelho #C8102E — usada em títulos, cabeçalhos de tabela, numeração de passos, bordas e rodapé
- Tipografia: sem serifa, títulos em caixa alta e negrito, corpo em peso normal
- Ficha técnica: bloco de métricas rápidas em destaque (rendimento, tempo, armazenamento), tabela com header vermelho (INSUMO | QUANTIDADE | ESPECIFICAÇÃO), passos numerados com círculos vermelhos
- Comunicado: tag de categoria no canto superior direito, linha separadora vermelha, título grande em negrito
- Logo cursiva no topo esquerdo + data no topo direito; logo no rodapé
- Seções em caixa alta e vermelho (ex: "INGREDIENTES", "MODO DE PREPARO")
- Linguagem direta e operacional, com cuidado estético""",
        "primary_color": (200, 16, 46),
        "heading_color": (160, 10, 30),
        "accent_hex": "C8102E",
    },
    "juliette": {
        "name": "Juliette",
        "persona": """Você representa a marca Juliette Bistrô Art Déco.

IDENTIDADE: nostalgia, sofisticação, glamour anos 60, beleza nos detalhes. Linguagem como poesia urbana — calma, sedutora, elegante. Nunca eufórica ou informal.

PADRÃO VISUAL (extraído do manual de identidade oficial da marca):
- Fundo off-white #F5F6F6 como base
- Azul acinzentado #506775 (Pantone 5405 C) — cor principal de texto e títulos
- Dourado #BFA56E (Pantone 4525 C) — bordas, frisos, separadores, acentos
- Tipografia primária: serifada elegante (estilo Playfair Display), pode ser itálica nos destaques
- Tipografia corpo: sem serifa leve (estilo Montserrat Light)
- Estrutura igual à Amélie (métricas, tabelas, passos numerados) mas nas cores azul+dourado
- Separadores: linha fina dourada horizontal
- Seções em caixa alta com espaçamento de letras, em dourado
- Palavras da marca: refúgio, beleza, pausa, clássico, vintage, sofisticação, aconchego, charme
- Evitar: linguagem rústica, gírias, tons eufóricos""",
        "primary_color": (80, 103, 117),
        "heading_color": (60, 80, 95),
        "accent_hex": "BFA56E",
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

def is_slide_pdf(file_bytes: bytes) -> bool:
    """Detecta se o PDF tem layout de slides (landscape + poucas palavras por página)"""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            if not pdf.pages:
                return False
            page = pdf.pages[0]
            is_landscape = (page.width or 0) > (page.height or 0)
            total_words = sum(len((p.extract_text() or "").split()) for p in pdf.pages)
            avg_words = total_words / len(pdf.pages) if pdf.pages else 0
            return is_landscape and avg_words < 100
    except:
        return False

def extract_text_pdf(file_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            t = page.extract_text()
            if t and t.strip():
                text_parts.append(f"## Slide {i}\n\n{t.strip()}")
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

    # Adiciona logo no cabeçalho
    try:
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        logo_b64 = LOGO_AMELIE if brand == "amelie" else LOGO_JULIETTE
        logo_bytes_docx = base64.b64decode(logo_b64)
        section = doc.sections[0]
        header = section.header
        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        hp.clear()
        run = hp.add_run()
        run.add_picture(io.BytesIO(logo_bytes_docx), width=docx.shared.Cm(4))
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Linha separadora no cabeçalho
        pPr = hp._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        r_hex, g_hex, b_hex = cfg["heading_color"]
        color_hex = f'{r_hex:02X}{g_hex:02X}{b_hex:02X}'
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:color'), color_hex)
        pBdr.append(bottom)
        pPr.append(pBdr)
    except Exception as logo_err:
        pass  # se falhar, documento sai sem logo mas não quebra

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def build_pdf(markdown_text: str, brand: str, is_slides: bool = False) -> bytes:
    cfg = BRAND_CONFIG[brand]
    r, g, b = cfg["primary_color"]
    brand_color = colors.Color(r/255, g/255, b/255)

    from reportlab.platypus import Image as RLImage
    from reportlab.lib.pagesizes import A4, landscape

    logo_b64 = LOGO_AMELIE if brand == "amelie" else LOGO_JULIETTE
    logo_bytes = base64.b64decode(logo_b64)

    if is_slides:
        pagesize = landscape(A4)
        lm = rm = 1.5*cm
        tm = bm = 1.5*cm
    else:
        pagesize = A4
        lm = rm = 2.5*cm
        tm = 3.5*cm  # espaço para logo no topo
        bm = 2*cm

    buf = io.BytesIO()

    def add_logo_header(canvas, doc):
        canvas.saveState()
        from PIL import Image as PILImage
        logo_img = PILImage.open(io.BytesIO(logo_bytes))
        w, h = logo_img.size
        logo_h = 1.2*cm
        logo_w = logo_h * (w / h)
        from reportlab.lib.utils import ImageReader
        logo_reader = ImageReader(io.BytesIO(logo_bytes))
        canvas.drawImage(logo_reader, lm, pagesize[1] - tm + 0.3*cm,
                        width=logo_w, height=logo_h,
                        preserveAspectRatio=True, mask='auto')
        # linha separadora
        canvas.setStrokeColor(colors.Color(r/255, g/255, b/255))
        canvas.setLineWidth(0.5)
        canvas.line(lm, pagesize[1] - tm + 0.1*cm,
                   pagesize[0] - rm, pagesize[1] - tm + 0.1*cm)
        canvas.restoreState()

    doc_pdf = SimpleDocTemplate(buf, pagesize=pagesize,
                                leftMargin=lm, rightMargin=rm,
                                topMargin=tm, bottomMargin=bm)
    styles = getSampleStyleSheet()
    story = []

    slide_title_style = ParagraphStyle("SlideTitle", parent=styles["Heading1"],
                               textColor=brand_color, fontSize=28, spaceAfter=20,
                               leading=34, alignment=1)
    slide_body_style = ParagraphStyle("SlideBody", parent=styles["Normal"],
                               fontSize=16, leading=24, spaceAfter=8)
    slide_bullet_style = ParagraphStyle("SlideBullet", parent=styles["Normal"],
                               fontSize=14, leading=22, leftIndent=25, spaceAfter=6)
    h1_style = ParagraphStyle("H1Brand", parent=styles["Heading1"],
                               textColor=brand_color, fontSize=16, spaceAfter=8)
    h2_style = ParagraphStyle("H2Brand", parent=styles["Heading2"],
                               textColor=brand_color, fontSize=13, spaceAfter=6)
    h3_style = ParagraphStyle("H3Brand", parent=styles["Heading3"],
                               textColor=brand_color, fontSize=11, spaceAfter=4)
    body_style = ParagraphStyle("Body", parent=styles["Normal"], fontSize=11, leading=16)
    bullet_style = ParagraphStyle("Bullet", parent=styles["Normal"],
                                   fontSize=11, leading=16, leftIndent=20, bulletIndent=10)

    lines = markdown_text.split("\n")
    first_slide = True
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6 if not is_slides else 12))
        elif line == "---":
            story.append(HRFlowable(width="100%", thickness=0.5, color=brand_color, spaceAfter=6))
        elif line.startswith("# ") or (is_slides and line.startswith("## ")):
            if is_slides and not first_slide:
                from reportlab.platypus import PageBreak
                story.append(PageBreak())
            first_slide = False
            text = line.lstrip("# ")
            story.append(Paragraph(text, slide_title_style if is_slides else h1_style))
        elif line.startswith("## "):
            story.append(Paragraph(line[3:], h2_style))
        elif line.startswith("### "):
            story.append(Paragraph(line[4:], h3_style))
        elif line.startswith("- ") or line.startswith("* "):
            text = f"• {line[2:]}"
            story.append(Paragraph(text, slide_bullet_style if is_slides else bullet_style))
        elif re.match(r"^\d+\. ", line):
            story.append(Paragraph(line, slide_bullet_style if is_slides else bullet_style))
        else:
            line_html = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", line)
            story.append(Paragraph(line_html, slide_body_style if is_slides else body_style))

    doc_pdf.build(story, onFirstPage=add_logo_header, onLaterPages=add_logo_header)
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
        model="claude-sonnet-4-5",
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
        import traceback
        detail = f"Erro na API do Claude: {str(e)}\n{traceback.format_exc()}"
        raise HTTPException(status_code=500, detail=detail)

    markdown_out = result.get("documento_padronizado", "")
    alteracoes = result.get("alteracoes", "")

    # 3. Reconstrói no formato original
    try:
        if ext == "docx":
            output_bytes = build_docx(markdown_out, brand)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            out_filename = filename.replace(".docx", "_padronizado.docx")
        elif ext == "pdf":
            slides = is_slide_pdf(file_bytes)
            output_bytes = build_pdf(markdown_out, brand, is_slides=slides)
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
