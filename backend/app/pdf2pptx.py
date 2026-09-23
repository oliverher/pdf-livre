"""
Trilha C: reconstrucao de PDF em .pptx nativo.

O PDF guarda cada trecho de texto com posicao, fonte, corpo e cor, e cada
desenho como caminho vetorial. Entao isto e remontagem, nao inferencia: nada
aqui adivinha conteudo. O que nao for reconhecido com seguranca e contado e
relatado, nunca aproximado em silencio.

A ordem de empilhamento vem de page.get_bboxlog(), que lista as operacoes de
desenho na ordem real do fluxo de conteudo. Cada elemento extraido e casado
com sua entrada no log pela caixa delimitadora, e emitido nessa ordem. Sem
isso, uma sombra ou textura desenhada antes de uma forma acabaria por cima
dela.

Modos de remontagem do texto:
  "fiel"     - uma caixa por linha. Posicao exata, edicao linha a linha.
  "editavel" - uma caixa por bloco, linhas unidas em um paragrafo com quebra
               automatica. Melhor de editar, sujeito a requebra diferente.

Limitacoes declaradas:
- Texto rotacionado e emitido na horizontal e contado em "texto_rotacionado".
- Gradientes sao achatados na cor de preenchimento informada pelo PDF.
- Caminhos que nao sejam retangulo, retangulo arredondado, elipse ou linha
  nao sao reconstruidos; entram em "caminhos_nao_reconstruidos".
"""

import io
import re
from collections import Counter

import pymupdf
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

EMU_POR_PONTO = 12700
FLAG_ITALICO = 1 << 1
FLAG_NEGRITO = 1 << 4

CLONES = {
    "carlito": "Calibri", "caladea": "Cambria",
    "liberationsans": "Arial", "liberationserif": "Times New Roman",
    "liberationmono": "Courier New", "dejavusans": "Arial",
    "dejavuserif": "Times New Roman", "dejavusansmono": "Courier New",
    "nimbussans": "Arial", "nimbusroman": "Times New Roman",
}


def pt2emu(v):
    return Emu(int(round(v * EMU_POR_PONTO)))


def _cor_int(v):
    return RGBColor((v >> 16) & 255, (v >> 8) & 255, v & 255)


def _cor_tupla(t):
    if not t:
        return None
    return RGBColor(*[max(0, min(255, int(round(c * 255)))) for c in t[:3]])


def normalizar_fonte(nome):
    n = re.sub(r"^[A-Z]{6}\+", "", nome or "")
    baixo = n.lower()
    negrito = "bold" in baixo or "black" in baixo or "heavy" in baixo
    italico = "italic" in baixo or "oblique" in baixo
    base = re.split(r"[-,]", n)[0]
    chave = re.sub(r"[^a-z]", "", base.lower())
    return CLONES.get(chave, base or "Calibri"), negrito, italico


def _dist(a, b):
    return sum(abs(a[i] - b[i]) for i in range(4))


def _classificar(g):
    tipos = [it[0] for it in g["items"]]
    r = g["rect"]
    if r.width <= 0.4 or r.height <= 0.4:
        return "linha" if tipos else None
    if "re" in tipos and len(tipos) == 1:
        return "retangulo"
    n_curvas, n_linhas = tipos.count("c"), tipos.count("l")
    if n_curvas == 0 and 3 <= n_linhas <= 5:
        return "retangulo"
    if n_curvas >= 4 and n_linhas == 0:
        return "elipse"
    if n_curvas >= 4 and 3 <= n_linhas <= 6:
        return "retangulo_arredondado"
    if n_linhas == 1 and n_curvas == 0:
        return "linha"
    return None


def _emitir_caminho(slide, g, tipo, rel):
    r = g["rect"]
    preenche, traco = g.get("fill"), g.get("color")
    if tipo == "linha":
        forma = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, pt2emu(r.x0), pt2emu(r.y0),
            pt2emu(max(r.width, 0.75)), pt2emu(max(r.height, 0.75)))
        forma.fill.solid()
        forma.fill.fore_color.rgb = (_cor_tupla(traco) or _cor_tupla(preenche)
                                     or RGBColor(0, 0, 0))
        forma.line.fill.background()
        forma.shadow.inherit = False
        rel["linhas"] += 1
        return
    molde = {"retangulo": MSO_SHAPE.RECTANGLE,
             "retangulo_arredondado": MSO_SHAPE.ROUNDED_RECTANGLE,
             "elipse": MSO_SHAPE.OVAL}[tipo]
    forma = slide.shapes.add_shape(molde, pt2emu(r.x0), pt2emu(r.y0),
                                   pt2emu(r.width), pt2emu(r.height))
    if preenche:
        forma.fill.solid()
        forma.fill.fore_color.rgb = _cor_tupla(preenche)
    else:
        forma.fill.background()
    if traco and g.get("width"):
        forma.line.color.rgb = _cor_tupla(traco)
        forma.line.width = pt2emu(g["width"])
    else:
        forma.line.fill.background()
    forma.shadow.inherit = False
    rel["formas"] += 1


def _emitir_imagem(slide, dados, bbox, rel):
    try:
        slide.shapes.add_picture(io.BytesIO(dados), pt2emu(bbox[0]),
                                 pt2emu(bbox[1]), pt2emu(bbox[2] - bbox[0]),
                                 pt2emu(bbox[3] - bbox[1]))
        rel["imagens"] += 1
    except Exception:
        rel["imagens_falhas"] += 1


def _aplicar_span(run, s):
    familia, neg_nome, ital_nome = normalizar_fonte(s.get("font"))
    f = run.font
    f.name = familia
    f.size = Pt(round(s.get("size", 12), 1))
    f.bold = bool(s.get("flags", 0) & FLAG_NEGRITO) or neg_nome
    f.italic = bool(s.get("flags", 0) & FLAG_ITALICO) or ital_nome
    f.color.rgb = _cor_int(s.get("color", 0))


def _nova_caixa(slide, bbox, folga_l=2.0, folga_a=1.5):
    x0, y0, x1, y1 = bbox
    cx = slide.shapes.add_textbox(pt2emu(x0 - folga_l), pt2emu(y0 - folga_a),
                                  pt2emu((x1 - x0) + folga_l * 2),
                                  pt2emu((y1 - y0) + folga_a * 2))
    tf = cx.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    return tf


def _emitir_texto_linha(slide, linha, rel):
    x0, y0, x1, y1 = linha["bbox"]
    tf = _nova_caixa(slide, linha["bbox"])
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = Pt(round(y1 - y0, 1))
    for s in linha["spans"]:
        r = p.add_run()
        r.text = s["text"]
        _aplicar_span(r, s)
    rel["caixas_texto"] += 1


def _emitir_texto_bloco(slide, linhas, bbox, rel):
    tf = _nova_caixa(slide, bbox, folga_l=3.0, folga_a=2.0)
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if len(linhas) > 1:
        passo = linhas[1]["bbox"][1] - linhas[0]["bbox"][1]
        if passo > 0:
            p.line_spacing = Pt(round(passo, 1))
    for i, l in enumerate(linhas):
        for j, s in enumerate(l["spans"]):
            txt = s["text"]
            if i and not j:
                txt = " " + txt.lstrip()
            r = p.add_run()
            r.text = txt
            _aplicar_span(r, s)
    rel["caixas_texto"] += 1


def _coletar(doc, pagina, modo, rel):
    elementos = []
    area = pagina.rect.width * pagina.rect.height
    for g in pagina.get_drawings():
        r = g["rect"]
        if (r.width * r.height) >= area * 0.985 and g.get("fill") == (1.0, 1.0, 1.0):
            rel["fundos_ignorados"] += 1
            continue
        tipo = _classificar(g)
        if tipo is None:
            rel["caminhos_nao_reconstruidos"] += 1
            continue
        elementos.append(("caminho", tuple(r), (g, tipo)))
    for info in pagina.get_image_info(xrefs=True):
        xref, bbox = info.get("xref"), info.get("bbox")
        if not bbox:
            continue
        dados = None
        if xref:
            try:
                ex = doc.extract_image(xref)
                dados = ex["image"] if ex else None
            except Exception:
                dados = None
        if dados is None:
            rel["imagens_falhas"] += 1
            continue
        elementos.append(("imagem", tuple(bbox), dados))
    for b in pagina.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        linhas = [l for l in b.get("lines", []) if l.get("spans")]
        if not linhas:
            continue
        for l in linhas:
            if abs(l.get("dir", (1, 0))[1]) > 1e-6:
                rel["texto_rotacionado"] += 1
        if modo == "editavel" and len(linhas) > 1:
            elementos.append(("bloco", tuple(linhas[0]["bbox"]),
                              (linhas, tuple(b["bbox"]))))
        else:
            for l in linhas:
                elementos.append(("linha", tuple(l["bbox"]), l))
    return elementos


ESPECIE_DO_LOG = {
    "fill-path": "caminho", "stroke-path": "caminho", "clip-path": "caminho",
    "fill-image": "imagem", "image": "imagem",
    "fill-text": "texto", "stroke-text": "texto", "ignore-text": "texto",
}


def _ordenar(pagina, elementos):
    try:
        log = pagina.get_bboxlog()
    except Exception:
        return elementos
    pendentes = {"caminho": [], "imagem": [], "texto": []}
    for idx, (especie, bbox, carga) in enumerate(elementos):
        chave = "texto" if especie in ("linha", "bloco") else especie
        pendentes[chave].append(idx)
    z = {}
    proximo = 0
    for entrada in log:
        especie = ESPECIE_DO_LOG.get(entrada[0])
        if especie is None or not pendentes[especie]:
            continue
        alvo = entrada[1]
        melhor = min(pendentes[especie],
                     key=lambda i: _dist(elementos[i][1], alvo))
        z[melhor] = proximo
        proximo += 1
        pendentes[especie].remove(melhor)
    for i in range(len(elementos)):
        z.setdefault(i, proximo + i)
    return [elementos[i] for i in sorted(range(len(elementos)), key=lambda i: z[i])]


def converter(caminho_pdf, caminho_pptx, modo="fiel", paginas_max=200):
    if modo not in ("fiel", "editavel"):
        raise ValueError("modo deve ser 'fiel' ou 'editavel'")
    doc = pymupdf.open(caminho_pdf)
    if doc.page_count == 0:
        raise ValueError("PDF sem paginas.")
    if doc.page_count > paginas_max:
        raise ValueError(f"PDF com {doc.page_count} paginas, acima do limite "
                         f"de {paginas_max}.")
    prs = Presentation()
    prs.slide_width = pt2emu(doc[0].rect.width)
    prs.slide_height = pt2emu(doc[0].rect.height)
    vazio = prs.slide_layouts[6]
    rel = Counter()
    rel["paginas"] = doc.page_count
    tamanhos = set()
    for pagina in doc:
        tamanhos.add((round(pagina.rect.width, 1), round(pagina.rect.height, 1)))
        slide = prs.slides.add_slide(vazio)
        elementos = _ordenar(pagina, _coletar(doc, pagina, modo, rel))
        for especie, _bbox, carga in elementos:
            if especie == "caminho":
                _emitir_caminho(slide, carga[0], carga[1], rel)
            elif especie == "imagem":
                _emitir_imagem(slide, carga, _bbox, rel)
            elif especie == "linha":
                _emitir_texto_linha(slide, carga, rel)
            else:
                _emitir_texto_bloco(slide, carga[0], carga[1], rel)
    doc.close()
    prs.save(caminho_pptx)
    r = dict(rel)
    r["modo"] = modo
    r["palco_polegadas"] = [round(prs.slide_width / 914400, 3),
                            round(prs.slide_height / 914400, 3)]
    r["paginas_de_tamanhos_diferentes"] = len(tamanhos) > 1
    return r
