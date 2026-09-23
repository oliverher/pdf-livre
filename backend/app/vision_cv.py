"""
Modo convencional de analise visual (sem modelo de IA).

Recebe UMA imagem (uma pagina achatada, um PNG, um JPG) e devolve uma lista de
regioes com tipo e, quando for texto, o conteudo reconhecido por OCR.

Cobre os requisitos: RF03 (ler imagem), RF04 (identificar areas), RF05/RF08
(separar e classificar em foto/logo/icone/texto/forma), RF06 (unir fragmentos),
RF07/RNF07 (descartar ruido e duplicacao), RF09 (OCR pt-BR), RF10/RNF06
(filtrar baixa confianca), RNF14 (limitar quantidade de objetos).

Este modo NAO vetoriza ilustracao. Regiao que nao for texto nem forma geometrica
simples e classificada como foto ou logo e preservada como recorte de imagem, tal
como a "Limitacao funcional atual" do documento de requisitos preve.

Dependencias: opencv-python, pytesseract (+ idioma 'por'), numpy, Pillow.
"""

from dataclasses import dataclass, field
from typing import Optional

import cv2
import numpy as np
import pytesseract

# --- parametros, todos documentados porque sao decisoes de projeto ---
AREA_MIN_FRAC = 0.0008     # RNF07: regiao menor que isto da area total e ruido
AREA_MAX_FRAC = 0.92       # regiao que cobre quase tudo e o fundo, nao um objeto
CONF_OCR_MIN = 60          # RF10/RNF06: palavra abaixo disto nao e exportada
IOU_DUPLICATA = 0.85       # RNF07: duas caixas com esta sobreposicao sao a mesma
DIST_UNIAO_FRAC = 0.015    # RF06: fragmentos a esta distancia sao unidos
MAX_REGIOES = 60           # RNF14: teto de objetos por imagem


@dataclass
class Regiao:
    x: int
    y: int
    w: int
    h: int
    tipo: str = "indefinido"          # texto | foto | logo | icone | forma
    texto: str = ""
    conf: float = 0.0
    cor_dominante: Optional[tuple] = None

    @property
    def caixa(self):
        return (self.x, self.y, self.x + self.w, self.y + self.h)

    @property
    def area(self):
        return self.w * self.h

    def dict(self):
        return {"x": self.x, "y": self.y, "w": self.w, "h": self.h,
                "tipo": self.tipo, "texto": self.texto,
                "conf": round(self.conf, 1), "cor": self.cor_dominante}


def _iou(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    ua = (ax1 - ax0) * (ay1 - ay0) + (bx1 - bx0) * (by1 - by0) - inter
    return inter / ua if ua else 0.0


def _cor_dominante(bgr):
    """Cor media da regiao, em hex, amostrando o miolo para evitar borda."""
    h, w = bgr.shape[:2]
    my, mx = int(h * 0.2), int(w * 0.2)
    miolo = bgr[my:h - my or h, mx:w - mx or w]
    if miolo.size == 0:
        miolo = bgr
    b, g, r = [int(np.median(miolo[:, :, i])) for i in range(3)]
    return f"#{r:02X}{g:02X}{b:02X}"


def _detectar_regioes(bgr, rel):
    """Segmentacao por MSER + morfologia. Devolve caixas candidatas."""
    h, w = bgr.shape[:2]
    area_total = h * w
    cinza = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)

    # realce de bordas e fechamento para juntar tracos proximos (RF06)
    grad = cv2.morphologyEx(cinza, cv2.MORPH_GRADIENT,
                            cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3)))
    _, bin_ = cv2.threshold(grad, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    kx = max(9, w // 55)
    fechado = cv2.morphologyEx(
        bin_, cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (kx, max(5, h // 90))))

    contornos, _ = cv2.findContours(fechado, cv2.RETR_EXTERNAL,
                                    cv2.CHAIN_APPROX_SIMPLE)
    caixas = []
    for c in contornos:
        x, y, cw, ch = cv2.boundingRect(c)
        a = cw * ch
        if a < area_total * AREA_MIN_FRAC:
            rel["descartados_ruido"] += 1
            continue
        if a > area_total * AREA_MAX_FRAC:
            continue
        # RNF07: descarta lasca (regiao muito fina), que e linha de grade ou artefato
        if cw < 6 or ch < 6 or (cw / max(ch, 1) > 60) or (ch / max(cw, 1) > 60):
            rel["descartados_ruido"] += 1
            continue
        caixas.append((x, y, cw, ch))
    return caixas


def _unir_fragmentos(caixas, w, h, rel):
    """RF06: funde caixas proximas ou sobrepostas até estabilizar."""
    tol = int(min(w, h) * DIST_UNIAO_FRAC)
    mudou = True
    while mudou and len(caixas) > 1:
        mudou = False
        novo = []
        usado = [False] * len(caixas)
        for i in range(len(caixas)):
            if usado[i]:
                continue
            x0, y0, cw0, ch0 = caixas[i]
            ax0, ay0, ax1, ay1 = x0, y0, x0 + cw0, y0 + ch0
            for j in range(i + 1, len(caixas)):
                if usado[j]:
                    continue
                x1, y1, cw1, ch1 = caixas[j]
                bx0, by0, bx1, by1 = x1, y1, x1 + cw1, y1 + ch1
                perto = not (bx0 > ax1 + tol or bx1 < ax0 - tol or
                             by0 > ay1 + tol or by1 < ay0 - tol)
                if perto and _iou((ax0, ay0, ax1, ay1),
                                  (bx0, by0, bx1, by1)) >= 0 and perto:
                    ax0, ay0 = min(ax0, bx0), min(ay0, by0)
                    ax1, ay1 = max(ax1, bx1), max(ay1, by1)
                    usado[j] = True
                    mudou = True
                    rel["fragmentos_unidos"] += 1
            usado[i] = True
            novo.append((ax0, ay0, ax1 - ax0, ay1 - ay0))
        caixas = novo
    return caixas


def _remover_duplicatas(regioes, rel):
    """RNF07: elimina caixas quase identicas, mantendo a maior."""
    regioes = sorted(regioes, key=lambda r: r.area, reverse=True)
    mantidas = []
    for r in regioes:
        if any(_iou(r.caixa, m.caixa) >= IOU_DUPLICATA for m in mantidas):
            rel["duplicatas_removidas"] += 1
            continue
        mantidas.append(r)
    return mantidas


def _classificar(bgr, r):
    """RF08: foto | logo | icone | texto | forma, por heuristica visual."""
    recorte = bgr[r.y:r.y + r.h, r.x:r.x + r.w]
    if recorte.size == 0:
        return "indefinido"
    cinza = cv2.cvtColor(recorte, cv2.COLOR_BGR2GRAY)

    # numero de cores distintas (foto tem muitas; forma/texto, poucas)
    reduz = (recorte // 32).reshape(-1, 3)
    n_cores = len(np.unique(reduz, axis=0))
    # densidade de borda (texto e denso em bordas; forma chapada, nao)
    bordas = cv2.Canny(cinza, 60, 160)
    dens_borda = float(np.count_nonzero(bordas)) / bordas.size
    proporcao = r.w / max(r.h, 1)

    if n_cores > 60 and dens_borda > 0.03:
        return "foto"
    if dens_borda > 0.06 and proporcao > 1.8:
        return "texto"
    if dens_borda > 0.06:
        return "texto" if proporcao > 1.2 else "icone"
    if n_cores <= 12 and dens_borda < 0.04:
        return "forma"
    if 0.6 <= proporcao <= 1.7 and n_cores <= 40:
        return "logo"
    return "foto"


def _ocr(bgr, r):
    """RF09/RF10: OCR pt-BR; devolve (texto, confianca media) ou ('',0)."""
    recorte = bgr[r.y:r.y + r.h, r.x:r.x + r.w]
    if recorte.size == 0:
        return "", 0.0
    cinza = cv2.cvtColor(recorte, cv2.COLOR_BGR2GRAY)
    escala = 2 if max(r.w, r.h) < 400 else 1
    if escala > 1:
        cinza = cv2.resize(cinza, None, fx=escala, fy=escala,
                           interpolation=cv2.INTER_CUBIC)
    base = cv2.threshold(cinza, 0, 255,
                         cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    # texto pode ser escuro-sobre-claro ou claro-sobre-escuro. Tesseract espera
    # texto escuro em fundo claro, entao tentamos a versao normal e a invertida
    # e ficamos com a de maior confianca media.
    melhor_txt, melhor_conf = "", 0.0
    for img in (base, cv2.bitwise_not(base)):
        try:
            dados = pytesseract.image_to_data(
                img, lang="por", config="--psm 6",
                output_type=pytesseract.Output.DICT)
        except Exception:
            continue
        palavras, confs = [], []
        for txt, conf in zip(dados["text"], dados["conf"]):
            try:
                conf = float(conf)
            except (TypeError, ValueError):
                conf = -1
            if txt.strip() and conf >= CONF_OCR_MIN:
                palavras.append(txt.strip())
                confs.append(conf)
        if palavras:
            media = sum(confs) / len(confs)
            if media > melhor_conf:
                melhor_txt, melhor_conf = " ".join(palavras), media
    return melhor_txt, melhor_conf


def _texto_corrompido(txt):
    """RF10: rejeita cadeia com excesso de simbolos sem sentido."""
    if not txt:
        return True
    legiveis = sum(c.isalnum() or c.isspace() or c in ".,;:!?%-/()º°ªR$" for c in txt)
    return legiveis / len(txt) < 0.65


def analisar_imagem(caminho_ou_bgr):
    """
    Analisa uma imagem e devolve (regioes, relatorio).
    Aceita caminho de arquivo ou array BGR do OpenCV.
    """
    rel = {"descartados_ruido": 0, "fragmentos_unidos": 0,
           "duplicatas_removidas": 0, "texto_baixa_conf": 0,
           "regioes_por_tipo": {}}

    if isinstance(caminho_ou_bgr, str):
        bgr = cv2.imread(caminho_ou_bgr)
        if bgr is None:
            raise ValueError("Nao foi possivel abrir a imagem.")
    else:
        bgr = caminho_ou_bgr
    h, w = bgr.shape[:2]

    caixas = _detectar_regioes(bgr, rel)
    caixas = _unir_fragmentos(caixas, w, h, rel)

    regioes = [Regiao(x, y, cw, ch) for (x, y, cw, ch) in caixas]
    regioes = _remover_duplicatas(regioes, rel)

    # classifica e roda OCR onde for texto
    finais = []
    for r in regioes:
        r.tipo = _classificar(bgr, r)
        r.cor_dominante = _cor_dominante(bgr[r.y:r.y + r.h, r.x:r.x + r.w])
        if r.tipo in ("texto", "icone", "logo"):
            txt, conf = _ocr(bgr, r)
            if txt and not _texto_corrompido(txt):
                r.texto, r.conf = txt, conf
                r.tipo = "texto"
            elif r.tipo == "texto":
                # parecia texto mas o OCR nao confirmou: vira recorte de imagem
                rel["texto_baixa_conf"] += 1
                r.tipo = "logo"
        finais.append(r)

    # RNF14: teto de objetos, mantendo os maiores
    if len(finais) > MAX_REGIOES:
        finais = sorted(finais, key=lambda r: r.area, reverse=True)[:MAX_REGIOES]

    for r in finais:
        rel["regioes_por_tipo"][r.tipo] = rel["regioes_por_tipo"].get(r.tipo, 0) + 1

    # ordena em leitura natural: de cima para baixo, esquerda para direita
    finais.sort(key=lambda r: (r.y // max(1, h // 12), r.x))
    return finais, rel
