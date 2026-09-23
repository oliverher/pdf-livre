"""
Trilha A: liberacao deterministica de um .pptx nativo.

Tres operacoes, todas sem modelo e sem inferencia:

1. destravar   - remove atributos restritivos de spLocks / grpSpLocks / picLocks
2. desagrupar  - achata p:grpSp, reposicionando filhos no espaco do slide
3. promover    - copia a arte nao-placeholder do slideLayout para o slide

Limitacoes conhecidas e declaradas:
- Desagrupar ignora rotacao (a:xfrm/@rot) e espelhamento (@flipH/@flipV) do
  grupo. Grupos rotacionados sao pulados e reportados, nao processados errado.
- Shapes sem a:xfrm proprio herdam posicao do layout; nao ha o que transformar,
  entao sao pulados e reportados.
- Promover duplica a arte no slide. O visual nao muda, o peso do arquivo sobe.
"""

import copy
import re

from pptx import Presentation
from pptx.oxml.ns import qn

ATRIBUTOS_TRAVA = (
    "noSelect", "noMove", "noResize", "noRot", "noEditPoints",
    "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType",
    "noTextEdit", "noUngrp", "noCrop", "noChangeAspect", "noGrp",
)

_TAGS_TRAVA = ("a:spLocks", "a:grpSpLocks", "a:picLocks", "a:graphicFrameLocks",
               "a:cxnSpLocks")


def _xfrm_de(el):
    """Devolve o elemento xfrm de um shape, qualquer que seja seu tipo."""
    tag = el.tag
    if tag == qn("p:grpSp"):
        pr = el.find(qn("p:grpSpPr"))
        return pr.find(qn("a:xfrm")) if pr is not None else None
    if tag == qn("p:graphicFrame"):
        return el.find(qn("p:xfrm"))
    pr = el.find(qn("p:spPr"))
    return pr.find(qn("a:xfrm")) if pr is not None else None


def _off_ext(xfrm):
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return off, ext


def destravar(prs):
    """Remove atributos de trava. Devolve numero de travas removidas."""
    removidas = 0
    for slide in prs.slides:
        for el in slide.shapes._spTree.iter():
            nome = el.tag.split("}")[-1]
            if f"a:{nome}" not in _TAGS_TRAVA:
                continue
            for attr in ATRIBUTOS_TRAVA:
                if attr in el.attrib:
                    del el.attrib[attr]
                    removidas += 1
    return removidas


def _achatar_grupo(spTree, grp, relatorio):
    """Move os filhos de um grpSp para o pai, corrigindo as coordenadas."""
    xfrm = _xfrm_de(grp)
    if xfrm is None:
        relatorio["pulados_sem_xfrm"] += 1
        return False

    if xfrm.get("rot") or xfrm.get("flipH") == "1" or xfrm.get("flipV") == "1":
        relatorio["pulados_rotacionados"] += 1
        return False

    par = _off_ext(xfrm)
    ch_off = xfrm.find(qn("a:chOff"))
    ch_ext = xfrm.find(qn("a:chExt"))
    if par is None or ch_off is None or ch_ext is None:
        relatorio["pulados_sem_xfrm"] += 1
        return False

    off, ext = par
    ox, oy = int(off.get("x")), int(off.get("y"))
    ex, ey = int(ext.get("cx")), int(ext.get("cy"))
    cox, coy = int(ch_off.get("x")), int(ch_off.get("y"))
    cex, cey = int(ch_ext.get("cx")), int(ch_ext.get("cy"))

    if cex == 0 or cey == 0:
        relatorio["pulados_sem_xfrm"] += 1
        return False

    sx, sy = ex / cex, ey / cey

    idx = list(spTree).index(grp)
    filhos = [c for c in grp
              if c.tag not in (qn("p:nvGrpSpPr"), qn("p:grpSpPr"))]

    movidos = 0
    for filho in filhos:
        fx = _xfrm_de(filho)
        if fx is None:
            relatorio["pulados_sem_xfrm"] += 1
            continue
        par_f = _off_ext(fx)
        if par_f is None:
            relatorio["pulados_sem_xfrm"] += 1
            continue
        f_off, f_ext = par_f
        nx = ox + round((int(f_off.get("x")) - cox) * sx)
        ny = oy + round((int(f_off.get("y")) - coy) * sy)
        f_off.set("x", str(nx))
        f_off.set("y", str(ny))
        f_ext.set("cx", str(max(1, round(int(f_ext.get("cx")) * sx))))
        f_ext.set("cy", str(max(1, round(int(f_ext.get("cy")) * sy))))
        spTree.insert(idx, filho)
        idx += 1
        movidos += 1

    spTree.remove(grp)
    relatorio["grupos_achatados"] += 1
    relatorio["shapes_liberados"] += movidos
    return True


def desagrupar(prs, max_passes=12):
    """Achata todos os grupos, inclusive aninhados. Devolve relatorio."""
    rel = {"grupos_achatados": 0, "shapes_liberados": 0,
           "pulados_rotacionados": 0, "pulados_sem_xfrm": 0}
    for slide in prs.slides:
        spTree = slide.shapes._spTree
        for _ in range(max_passes):
            grupos = [c for c in spTree if c.tag == qn("p:grpSp")]
            if not grupos:
                break
            mudou = False
            for g in grupos:
                if _achatar_grupo(spTree, g, rel):
                    mudou = True
            if not mudou:
                break
    return rel


def promover_arte_do_layout(prs):
    """
    Copia para o slide os shapes do slideLayout que nao sao placeholders.
    E a arte decorativa que o usuario ve mas nao consegue selecionar.
    """
    copiados = 0
    for slide in prs.slides:
        layout = slide.slide_layout
        spTree = slide.shapes._spTree
        ids = {int(e.get("id")) for e in spTree.iter(qn("p:cNvPr"))
               if e.get("id") and e.get("id").isdigit()}
        proximo = max(ids) + 1 if ids else 100

        for el in layout.shapes._spTree:
            if el.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
                continue
            if el.find(".//" + qn("p:ph")) is not None:
                continue  # placeholder: o slide ja herda
            novo = copy.deepcopy(el)
            for cnv in novo.iter(qn("p:cNvPr")):
                cnv.set("id", str(proximo))
                cnv.set("name", (cnv.get("name") or "shape") + " (do layout)")
                proximo += 1
            spTree.insert(2, novo)  # apos nvGrpSpPr e grpSpPr
            copiados += 1
    return copiados


def liberar(caminho_entrada, caminho_saida, fazer_destravar=True,
            fazer_desagrupar=True, fazer_promover=False):
    """Executa a Trilha A. Devolve o relatorio das operacoes."""
    prs = Presentation(caminho_entrada)
    rel = {"destravadas": 0, "grupos_achatados": 0, "shapes_liberados": 0,
           "pulados_rotacionados": 0, "pulados_sem_xfrm": 0,
           "promovidos_do_layout": 0}

    if fazer_destravar:
        rel["destravadas"] = destravar(prs)
    if fazer_promover:
        rel["promovidos_do_layout"] = promover_arte_do_layout(prs)
    if fazer_desagrupar:
        rel.update(desagrupar(prs))

    prs.save(caminho_saida)
    return rel
