# Padronizador — Backend

## Rodar localmente

```bash
pip install -r requirements.txt
ANTHROPIC_API_KEY=sk-... uvicorn main:app --reload --port 8000
```

Acesse: http://localhost:8000/docs

---

## Deploy no Railway (recomendado)

1. Crie uma conta em https://railway.app
2. Clique em "New Project" → "Deploy from GitHub repo"
3. Aponte para este repositório
4. Em "Variables", adicione:
   - `ANTHROPIC_API_KEY` = sua chave da API Anthropic
5. Railway detecta o `requirements.txt` automaticamente
6. Após o deploy, copie a URL gerada (ex: https://padronizador.up.railway.app)
7. Cole essa URL no campo `API_URL` do arquivo `padronizador_app.html`

---

## Deploy no Render

1. Crie conta em https://render.com
2. "New Web Service" → conecte seu repositório
3. Build command: `pip install -r requirements.txt`
4. Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Adicione a env var `ANTHROPIC_API_KEY`
6. Cole a URL no `padronizador_app.html`

---

## Formatos suportados

| Entrada | Saída |
|---------|-------|
| .docx   | .docx com estilos da marca |
| .pdf    | .pdf com cores e tipografia da marca |
| .pptx   | .pptx com slides padronizados |
| .xlsx   | .xlsx com cabeçalhos coloridos |
| .txt    | .txt revisado |

## Endpoint

`POST /padronizar`

| Campo | Tipo | Descrição |
|-------|------|-----------|
| file | File | Documento a padronizar |
| brand | string | `amelie` ou `juliette` |
| audience | string | `franqueados`, `time` ou `liderancas` |
| doc_type | string | `ficha`, `relatorio`, `checklist` ou `apresentacao` |
| reference | string | (opcional) Texto de exemplo de referência |
