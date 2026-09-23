"""
Orquestrador da reconstrucao.

Recebe um arquivo ja normalizado em imagens (uma por pagina) e monta o .pptx,
decidindo a rota e cumprindo os requisitos de reconstrucao:

- RF11/RF12/RNF24: se a separacao convencional nao produzir elementos editaveis
  suficientes, marca a pagina como "reconstrucao insuficiente" e cai para o modo
  de referencia (imagem limpa + textos que o OCR conseguiu), sem impedir a
  conversao.
- RF13: cada pagina reconstruida ganha, junto, um slide com a imagem original
  limpa como referencia visual.
- RF14/RF15/RF16: pagina densa (muitas regioes) e dividida em ate quatro secoes,
  cada uma em seu slide, com o recorte da secao como apoio.
- RF17/RF18/RF19: textos viram caixas editaveis; formas viram formas nativas
  moviveis; posicoes, proporcoes e cores do original sao preservadas.

Este modulo nao vetoriza ilustracao. Foto e logo sao preservados como imagem,
conforme a "Limitacao funcional atual" dos requisitos.
"""

import io

import cv2
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

from . import vision_ai
from .vision_cv import analisar_imagem

EMU_POR_PX = None  # definido por pagina conforme DPI

# limiares de decisao, documentados
MIN_EDITAVEIS = 3          # RF11: menos textos editaveis que isto = insuficiente
FRAC_TEXTO_MIN = 0.15      # se area de texto < 15% da util, provavel infografico
DENSO_REGIOES = 22         # RF14: acima disto, divide a pagina
MAX_SECOES = 4             # RF14: no maximo quatro secoes


def _hex_para_rgb(h):
    h = (h or "#808080").lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _px2emu(px, dpi):
    return Emu(int(round(px / dpi * 914400)))


def _cor_texto_para(hex_fundo):
    """Escolhe preto ou branco conforme o fundo, para o texto ficar legivel."""
    h = (hex_fundo or "#FFFFFF").lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    luminancia = 0.299 * r + 0.587 * g + 0.114 * b
    return RGBColor(0, 0, 0) if luminancia > 140 else RGBColor(255, 255, 255)


def _adicionar_imagem(slide, bgr, x, y, w, h, dpi):
    ok, buf = cv2.imencode(".png", bgr)
    if not ok:
        return
    slide.shapes.add_picture(io.BytesIO(buf.tobytes()),
                             _px2emu(x, dpi), _px2emu(y, dpi),
                             _px2emu(w, dpi), _px2emu(h, dpi))


def _adicionar_texto(slide, r, dpi):
    cx = slide.shapes.add_textbox(_px2emu(r.x, dpi), _px2emu(r.y, dpi),
                                  _px2emu(r.w, dpi), _px2emu(r.h, dpi))
    tf = cx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(int(0.03 * 914400))
    tf.margin_top = tf.margin_bottom = Emu(int(0.02 * 914400))
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = r.texto
    corpo = max(9, min(40, int(r.h / dpi * 72 * 0.55)))
    run.font.size = Pt(corpo)
    run.font.color.rgb = _cor_texto_para(r.cor_dominante)


def _adicionar_forma(slide, r, dpi):
    forma = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if min(r.w, r.h) > 40 else MSO_SHAPE.RECTANGLE,
        _px2emu(r.x, dpi), _px2emu(r.y, dpi),
        _px2emu(r.w, dpi), _px2emu(r.h, dpi))
    forma.fill.solid()
    forma.fill.fore_color.rgb = _hex_para_rgb(r.cor_dominante)
    forma.line.fill.background()
    forma.shadow.inherit = False
    return forma


def _emitir_regioes(slide, bgr, regioes, dpi):
    """Coloca formas, depois fotos/logos, depois texto (ordem de leitura visual)."""
    ordem = {"forma": 0, "icone": 1, "foto": 2, "logo": 2, "texto": 3}
    for r in sorted(regioes, key=lambda r: ordem.get(r.tipo, 2)):
        if r.tipo == "forma":
            _adicionar_forma(slide, r, dpi)
        elif r.tipo in ("foto", "logo", "icone"):
            recorte = bgr[r.y:r.y + r.h, r.x:r.x + r.w]
            _adicionar_imagem(slide, recorte, r.x, r.y, r.w, r.h, dpi)
        elif r.tipo == "texto" and r.texto:
            _adicionar_texto(slide, r, dpi)


def _slide_referencia(prs, bgr, dpi, rotulo):
    """RF13: slide com a imagem original limpa."""
    h, w = bgr.shape[:2]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _adicionar_imagem(slide, bgr, 0, 0, w, h, dpi)
    cx = slide.shapes.add_textbox(_px2emu(10, dpi), _px2emu(6, dpi),
                                  _px2emu(w - 20, dpi), _px2emu(28, dpi))
    p = cx.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = rotulo
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(120, 120, 120)


def _dividir_secoes(bgr, regioes, n_alvo):
    """RF14: divide a pagina em faixas horizontais equilibradas por conteudo."""
    h = bgr.shape[0]
    regioes = sorted(regioes, key=lambda r: r.y)
    n = min(n_alvo, MAX_SECOES)
    corte = h / n
    secoes = [[] for _ in range(n)]
    for r in regioes:
        centro = r.y + r.h / 2
        idx = min(n - 1, int(centro // corte))
        secoes[idx].append(r)
    limites = [(int(i * corte), int((i + 1) * corte)) for i in range(n)]
    return list(zip(secoes, limites))


def montar_pptx(paginas, caminho_saida, dpi=150, dividir_densos=True):
    """
    paginas: lista de arrays BGR (uma imagem por pagina de origem).
    Devolve o relatorio agregado (RF20/RF21).
    """
    if not paginas:
        raise ValueError("Nenhuma pagina para montar.")

    prs = Presentation()
    h0, w0 = paginas[0].shape[:2]
    prs.slide_width = _px2emu(w0, dpi)
    prs.slide_height = _px2emu(h0, dpi)
    vazio = prs.slide_layouts[6]

    rel = {"paginas": len(paginas), "slides": 0, "objetos": 0,
           "textos": 0, "formas": 0, "imagens": 0,
           "paginas_reconstruidas": 0, "paginas_referencia": 0,
           "paginas_divididas": 0, "detalhe": []}

    # RF11/RF12 + RNF24: usa o modelo de visao so se a chave estiver ativa
    # E o gasto do dia ainda couber no teto para todas as paginas deste
    # arquivo. Decisao tomada uma vez, no inicio do arquivo, para nao gastar
    # so uma parte e trocar de motor no meio da mesma apresentacao.
    usar_ia = False
    if vision_ai.vision_ativa():
        cabe, gasto_hoje, custo_previsto = vision_ai.orcamento_disponivel(len(paginas))
        usar_ia = cabe
        if cabe:
            vision_ai._registrar_gasto(custo_previsto)

    for i, bgr in enumerate(paginas, 1):
        if usar_ia:
            regioes, det = vision_ai.analisar_pagina_com_ia(bgr)
        else:
            regioes, det = analisar_imagem(bgr)
        editaveis = [r for r in regioes if r.tipo == "texto" and r.texto]
        area_util = bgr.shape[0] * bgr.shape[1]
        area_texto = sum(r.area for r in editaveis)
        frac_texto = area_texto / area_util if area_util else 0

        suficiente = len(editaveis) >= MIN_EDITAVEIS  # RF11

        # RF14/RF15: densa e suficiente -> dividir em secoes
        if dividir_densos and suficiente and len(regioes) >= DENSO_REGIOES:
            n = min(MAX_SECOES, 1 + len(regioes) // DENSO_REGIOES + 1)
            secoes = _dividir_secoes(bgr, regioes, n)
            rel["paginas_divididas"] += 1
            for k, (regs, (y0, y1)) in enumerate(secoes, 1):
                if not regs:
                    continue
                slide = prs.slides.add_slide(vazio)
                # desloca as regioes para a origem da secao
                for r in regs:
                    r.y = max(0, r.y - y0)
                _emitir_regioes(slide, bgr[y0:y1], regs, dpi)
                rel["slides"] += 1
                rel["textos"] += sum(1 for r in regs if r.tipo == "texto" and r.texto)
                rel["formas"] += sum(1 for r in regs if r.tipo == "forma")
                rel["imagens"] += sum(1 for r in regs if r.tipo in ("foto", "logo", "icone"))
            rel["paginas_reconstruidas"] += 1
            _slide_referencia(prs, bgr, dpi, f"Referência — página {i} (original)")
            rel["slides"] += 1
            rel["paginas_referencia"] += 1
            rel["detalhe"].append({"pagina": i, "rota": "reconstruida_dividida",
                                   "secoes": len([s for s, _ in secoes if s]),
                                   "editaveis": len(editaveis)})
            continue

        # rota normal de reconstrucao
        slide = prs.slides.add_slide(vazio)
        if suficiente:
            _emitir_regioes(slide, bgr, regioes, dpi)
            rel["paginas_reconstruidas"] += 1
            rota = "reconstruida"
        else:
            # RF12/RNF24: insuficiente -> imagem de referencia + textos achados
            _adicionar_imagem(slide, bgr, 0, 0, bgr.shape[1], bgr.shape[0], dpi)
            for r in editaveis:
                _adicionar_texto(slide, r, dpi)
            rota = "referencia_com_texto"
        rel["slides"] += 1
        rel["textos"] += sum(1 for r in regioes if r.tipo == "texto" and r.texto)
        rel["formas"] += sum(1 for r in regioes if r.tipo == "forma")
        rel["imagens"] += sum(1 for r in regioes if r.tipo in ("foto", "logo", "icone"))

        if suficiente:
            _slide_referencia(prs, bgr, dpi, f"Referência — página {i} (original)")
            rel["slides"] += 1
            rel["paginas_referencia"] += 1

        rel["detalhe"].append({"pagina": i, "rota": rota,
                               "regioes": len(regioes),
                               "editaveis": len(editaveis),
                               "frac_texto": round(frac_texto, 3)})

    rel["objetos"] = rel["textos"] + rel["formas"] + rel["imagens"]
    prs.save(caminho_saida)
    return rel
