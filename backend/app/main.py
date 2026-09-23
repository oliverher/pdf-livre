"""
PPT Livre - API.

Atende os requisitos funcionais RF01-RF27 e nao funcionais aplicaveis ao
servidor. O reconhecimento inteligente por modelo de visao (RF11/RF12 em sua
forma plena) fica atras da variavel de ambiente ANTHROPIC_API_KEY (aceita
tambem PPTLIVRE_VISION_KEY, por compatibilidade); sem ela, o app opera no
modo convencional (OpenCV + Tesseract), como manda o RNF24. Ha um teto de
gasto diario (ver vision_ai.py) que, se estourado, tambem faz o app cair
de volta no modo convencional para aquela requisicao.

Politica de dados (RNF16, RNF18, RNF19): o original e descartado apos o
processamento; o resultado vive num diretorio isolado por identificador
aleatorio, nao listavel, apagado por varredura e no download.
"""

import asyncio
import os
import shutil
import tempfile
import time
import uuid
from collections import defaultdict, deque
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import vision_ai
from .audit import auditar
from .normalize import extensao_aceita, normalizar, FORMATOS_ACEITOS
from .rebuild import montar_pptx
from .unlock import liberar

# ----------------------------------------------------------------- config
MAX_BYTES = int(os.environ.get("PPTLIVRE_MAX_MB", "60")) * 1024 * 1024
TTL_SEGUNDOS = 20 * 60
DPI = 150
PAGINAS_MAX = 200
RATE_JANELA = 60          # s
RATE_MAX = 8              # requisicoes de processamento por janela por IP
VISION_ATIVA = vision_ai.vision_ativa()

BASE = Path(__file__).parent
TRABALHO = Path(tempfile.gettempdir()) / "pptlivre"
TRABALHO.mkdir(exist_ok=True)

app = FastAPI(title="PPT Livre", version="0.2.0")

# jobs em memoria: id -> {"etapa","pct","pronto","erro","resultado"}
JOBS = {}
# limitador de taxa em memoria: ip -> deque de timestamps
_HITS = defaultdict(deque)


# ----------------------------------------------------------------- infra
def _purgar():
    limite = time.time() - TTL_SEGUNDOS
    for d in TRABALHO.iterdir():
        try:
            if d.is_dir() and d.stat().st_mtime < limite:
                shutil.rmtree(d, ignore_errors=True)
                JOBS.pop(d.name, None)
        except OSError:
            pass


@app.on_event("startup")
async def _startup():
    async def laco():
        while True:
            _purgar()
            await asyncio.sleep(300)
    asyncio.create_task(laco())


def _limitar(request: Request):
    ip = request.client.host if request.client else "?"
    agora = time.time()
    fila = _HITS[ip]
    while fila and fila[0] < agora - RATE_JANELA:
        fila.popleft()
    if len(fila) >= RATE_MAX:
        raise HTTPException(429, "Muitas conversões seguidas. Aguarde um minuto "
                                 "e tente de novo.")
    fila.append(agora)


def _validar(nome, dados):
    if not extensao_aceita(nome):
        aceitos = ", ".join(e.lstrip(".").upper() for e in FORMATOS_ACEITOS)
        raise HTTPException(415, f"Formato não suportado. Envie: {aceitos}.")
    if nome.lower().endswith(".ppt"):
        raise HTTPException(415, "Formato .ppt antigo. Salve como .pptx e reenvie.")
    if len(dados) > MAX_BYTES:
        raise HTTPException(413, f"Arquivo acima de {MAX_BYTES // (1024*1024)} MB.")
    if not dados:
        raise HTTPException(400, "Arquivo vazio.")


# ----------------------------------------------------------------- rotas
@app.get("/api/saude")
async def saude():
    return {"ok": True, "versao": app.version, "visao_inteligente": VISION_ATIVA,
            "formatos": [e.lstrip(".") for e in FORMATOS_ACEITOS]}


@app.post("/api/analisar")
async def analisar(request: Request, arquivo: UploadFile = File(...)):
    """Triagem rapida de um .pptx (mede editabilidade). So para PPTX."""
    _limitar(request)
    dados = await arquivo.read()
    nome = arquivo.filename or ""
    if not nome.lower().endswith((".pptx", ".potx")):
        raise HTTPException(415, "A análise de editabilidade vale para PPTX. "
                                 "Para PDF ou imagem, use Reconstruir.")
    _validar(nome, dados)
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "e.pptx"
        p.write_bytes(dados)
        try:
            rel = auditar(str(p))
        except Exception as e:
            raise HTTPException(422, f"Não foi possível ler o arquivo: {e}")
    rel["arquivo"] = nome
    return JSONResponse(rel)


@app.post("/api/liberar")
async def api_liberar(request: Request, arquivo: UploadFile = File(...),
                      destravar: bool = Form(True), desagrupar: bool = Form(True),
                      promover: bool = Form(False)):
    """Trilha A: destrava e desagrupa um PPTX nativo."""
    _limitar(request)
    dados = await arquivo.read()
    nome = arquivo.filename or ""
    if not nome.lower().endswith((".pptx", ".potx")):
        raise HTTPException(415, "Liberar vale para PPTX nativo.")
    _validar(nome, dados)

    ident = uuid.uuid4().hex
    dt = TRABALHO / ident
    dt.mkdir(parents=True, exist_ok=True)
    entrada = dt / "entrada.pptx"
    saida = dt / ((Path(nome).stem or "apresentacao") + "-liberado.pptx")
    entrada.write_bytes(dados)
    try:
        antes = auditar(str(entrada))
        ops = liberar(str(entrada), str(saida), destravar, desagrupar, promover)
        depois = auditar(str(saida))
    except Exception as e:
        shutil.rmtree(dt, ignore_errors=True)
        raise HTTPException(422, f"Falha ao processar: {e}")
    finally:
        entrada.unlink(missing_ok=True)
    return JSONResponse({"download_id": ident, "nome_saida": saida.name,
                         "operacoes": ops, "antes": antes, "depois": depois})


def _tem_texto_vivo(caminho_pdf, minimo=40):
    """Decide a rota: PDF/PPTX com texto vetorial vai para reconstrucao fiel
    (Trilha C); sem texto vivo, e imagem achatada e vai para a visao."""
    import pymupdf
    doc = pymupdf.open(caminho_pdf)
    total = sum(len(p.get_text("text").strip()) for p in doc)
    doc.close()
    return total >= minimo


async def _processar_reconstrucao(ident, caminho, nome, dividir):
    job = JOBS[ident]
    try:
        ext = Path(nome).suffix.lower()
        saida = Path(caminho).parent / ((Path(nome).stem or "arquivo") + "-reconstruido.pptx")

        job.update(etapa="Lendo o arquivo", pct=10)
        await asyncio.sleep(0)

        # Rota vetorial (Trilha C) para PDF/PPTX que ainda tem texto vivo.
        pdf_para_c = None
        if ext == ".pdf":
            pdf_para_c = caminho
        elif ext in (".pptx", ".potx"):
            import tempfile as _tf
            dstmp = _tf.mkdtemp(dir=str(Path(caminho).parent))
            pdf_para_c = await asyncio.to_thread(_pptx_para_pdf_local, caminho, dstmp)

        rota = "visao"
        if pdf_para_c and await asyncio.to_thread(_tem_texto_vivo, pdf_para_c):
            job.update(etapa="Reconstruindo texto e formas (vetorial)", pct=55)
            from .pdf2pptx import converter
            rel = await asyncio.to_thread(converter, pdf_para_c, str(saida), "fiel", PAGINAS_MAX)
            rota = "vetorial"
            formato = ext.lstrip(".")
        else:
            job.update(etapa="Analisando a composição", pct=40)
            imagens, formato = await asyncio.to_thread(
                normalizar, caminho, nome, DPI, PAGINAS_MAX)
            job.update(etapa="Reconstruindo textos, formas e imagens", pct=65)
            rel = await asyncio.to_thread(montar_pptx, imagens, str(saida), DPI, dividir)

        job.update(etapa="Conferindo o resultado", pct=88)
        aud = await asyncio.to_thread(auditar, str(saida))
        rel["indice_editabilidade"] = aud["indice_editabilidade_global"]
        rel["rota"] = rota

        Path(caminho).unlink(missing_ok=True)  # RF24: some com a copia de entrada
        job.update(etapa="Pronto", pct=100, pronto=True,
                   resultado={"download_id": ident, "nome_saida": saida.name,
                              "formato_origem": formato, "relatorio": rel})
    except Exception as e:
        Path(caminho).unlink(missing_ok=True)
        job.update(etapa="Erro", pct=100, pronto=True, erro=str(e))


def _pptx_para_pdf_local(caminho_pptx, dir_saida):
    from .normalize import _pptx_para_pdf
    return _pptx_para_pdf(caminho_pptx, dir_saida)


@app.post("/api/reconstruir")
async def reconstruir(request: Request, arquivo: UploadFile = File(...),
                      dividir_densos: bool = Form(True)):
    """RF01-RF19: reconstroi qualquer formato aceito em PPTX editavel."""
    _limitar(request)
    dados = await arquivo.read()
    nome = arquivo.filename or ""
    _validar(nome, dados)

    ident = uuid.uuid4().hex
    dt = TRABALHO / ident
    dt.mkdir(parents=True, exist_ok=True)
    entrada = dt / ("entrada" + Path(nome).suffix.lower())
    entrada.write_bytes(dados)

    JOBS[ident] = {"etapa": "Na fila", "pct": 0, "pronto": False,
                   "erro": None, "resultado": None}
    asyncio.create_task(
        _processar_reconstrucao(ident, str(entrada), nome, dividir_densos))
    return JSONResponse({"job_id": ident})


@app.get("/api/progresso/{ident}")
async def progresso(ident: str):
    if not ident.isalnum() or ident not in JOBS:
        raise HTTPException(404, "Trabalho não encontrado ou expirado.")
    j = JOBS[ident]
    return {"etapa": j["etapa"], "pct": j["pct"], "pronto": j["pronto"],
            "erro": j["erro"], "resultado": j["resultado"]}


@app.get("/api/download/{ident}")
async def download(ident: str):
    if not ident.isalnum():
        raise HTTPException(400, "Identificador inválido.")
    dt = TRABALHO / ident
    if not dt.is_dir():
        raise HTTPException(404, "Resultado expirado. Envie o arquivo de novo.")
    arqs = [f for f in dt.iterdir() if f.suffix == ".pptx"
            and "entrada" not in f.name]
    if not arqs:
        raise HTTPException(404, "Resultado expirado. Envie o arquivo de novo.")
    alvo = arqs[0]
    return FileResponse(
        alvo, filename=alvo.name,
        media_type="application/vnd.openxmlformats-officedocument."
                   "presentationml.presentation")


app.mount("/", StaticFiles(directory=BASE / "static", html=True), name="static")
