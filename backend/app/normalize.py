"""
Normalizacao de entrada (RF01, RF02, RF03).

Converte qualquer formato aceito numa lista de imagens BGR do OpenCV, uma por
pagina. E o ponto unico onde os formatos divergem; o resto do pipeline so ve
imagens.

Formatos: PPTX, PDF, PNG, JPG, JPEG, WebP.
- PPTX -> PDF (LibreOffice) -> imagens (PyMuPDF)
- PDF  -> imagens (PyMuPDF)
- imagem unica -> lista de uma imagem

Nao modifica o arquivo original (RF24/RNF16): tudo ocorre sobre copias em
diretorio temporario.
"""

import subprocess
import tempfile
from pathlib import Path

import cv2
import numpy as np
import pymupdf

FORMATOS_IMAGEM = (".png", ".jpg", ".jpeg", ".webp")
FORMATOS_DOC = (".pdf", ".pptx", ".potx")
FORMATOS_ACEITOS = FORMATOS_IMAGEM + FORMATOS_DOC

SOFFICE = "/mnt/skills/public/pptx/scripts/office/soffice.py"


def extensao_aceita(nome):
    return (nome or "").lower().endswith(FORMATOS_ACEITOS)


def _pdf_para_imagens(caminho_pdf, dpi, paginas_max):
    doc = pymupdf.open(caminho_pdf)
    if doc.page_count > paginas_max:
        doc.close()
        raise ValueError(f"Documento com {doc.page_count} paginas, acima do "
                         f"limite de {paginas_max}.")
    imagens = []
    zoom = dpi / 72.0
    matriz = pymupdf.Matrix(zoom, zoom)
    for pagina in doc:
        pix = pagina.get_pixmap(matrix=matriz, alpha=False)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n)
        bgr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR) if pix.n >= 3 else \
            cv2.cvtColor(arr, cv2.COLOR_GRAY2BGR)
        imagens.append(bgr)
    doc.close()
    return imagens


def _pptx_para_pdf(caminho_pptx, dir_saida):
    subprocess.run(
        ["python3", SOFFICE, "--headless", "--convert-to", "pdf",
         "--outdir", dir_saida, caminho_pptx],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        timeout=180)
    pdfs = list(Path(dir_saida).glob("*.pdf"))
    if not pdfs:
        raise ValueError("Falha ao converter o PPTX. Arquivo pode estar corrompido.")
    return str(pdfs[0])


def normalizar(caminho, nome_original, dpi=150, paginas_max=200):
    """
    Devolve (imagens, formato). imagens e lista de arrays BGR.
    Nao altera o arquivo em `caminho`.
    """
    ext = Path(nome_original).suffix.lower()
    if ext not in FORMATOS_ACEITOS:
        raise ValueError(f"Formato {ext} nao suportado.")

    if ext in FORMATOS_IMAGEM:
        dados = np.fromfile(caminho, dtype=np.uint8)
        bgr = cv2.imdecode(dados, cv2.IMREAD_COLOR)
        if bgr is None:
            raise ValueError("Nao foi possivel decodificar a imagem.")
        return [bgr], "imagem"

    if ext == ".pdf":
        return _pdf_para_imagens(caminho, dpi, paginas_max), "pdf"

    # pptx / potx
    with tempfile.TemporaryDirectory() as td:
        pdf = _pptx_para_pdf(caminho, td)
        return _pdf_para_imagens(pdf, dpi, paginas_max), "pptx"
