"""
Reconstrucao por modelo de visao (Claude), opcional e com teto de gasto.

So e chamado quando ANTHROPIC_API_KEY esta definida (ver main.py). Sem a
chave, o pipeline usa vision_cv.py (OpenCV + Tesseract), gratuito. Este
modulo existe para melhorar a qualidade da rota "visao" quando o usuario
decidir pagar por isso -- nunca substitui silenciosamente a rota vetorial,
que continua sendo preferida sempre que ha texto vivo (ver main._tem_texto_vivo).

Teto de gasto (RNF exigido antes de ligar em producao, discutido com o
usuario): sem ele, um link publico sem login e risco direto de fatura, nao
so de privacidade. O teto aqui e simples e local -- um contador diario em
arquivo -- suficiente para uso pessoal/piloto. Para producao institucional,
o teto real deve vir do painel de billing da Anthropic (limite de uso
mensal), que este modulo nao substitui.
"""

import json
import os
import time
from pathlib import Path

ARQUIVO_GASTO = Path(os.environ.get("PPTLIVRE_GASTO_ARQUIVO",
                                     "/tmp/pptlivre_gasto_diario.json"))

# Custo aproximado por imagem de slide analisada, em dolares. Estimativa
# conservadora para Claude com visao (entrada de imagem + resposta
# estruturada) -- ajuste conforme o modelo escolhido e o tamanho real das
# imagens enviadas. Documentado aqui porque e decisao de projeto, nao fato.
CUSTO_ESTIMADO_POR_PAGINA_USD = 0.02

TETO_DIARIO_USD = float(os.environ.get("PPTLIVRE_TETO_DIARIO_USD", "5.0"))

MODELO = os.environ.get("PPTLIVRE_VISION_MODEL", "claude-opus-4-8")


def _chave():
    return os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("PPTLIVRE_VISION_KEY")


def vision_ativa():
    return bool(_chave())


def _ler_gasto_hoje():
    hoje = time.strftime("%Y-%m-%d")
    if not ARQUIVO_GASTO.exists():
        return hoje, 0.0
    try:
        d = json.loads(ARQUIVO_GASTO.read_text())
    except (json.JSONDecodeError, OSError):
        return hoje, 0.0
    if d.get("data") != hoje:
        return hoje, 0.0
    return hoje, float(d.get("gasto_usd", 0.0))


def _registrar_gasto(valor_usd):
    hoje, atual = _ler_gasto_hoje()
    novo = atual + valor_usd
    ARQUIVO_GASTO.write_text(json.dumps({"data": hoje, "gasto_usd": novo}))
    return novo


def orcamento_disponivel(n_paginas):
    """Verifica se processar n_paginas cabe no teto diario. Nao gasta ainda."""
    _, gasto = _ler_gasto_hoje()
    custo_previsto = n_paginas * CUSTO_ESTIMADO_POR_PAGINA_USD
    return (gasto + custo_previsto) <= TETO_DIARIO_USD, gasto, custo_previsto


def analisar_pagina_com_ia(imagem_bgr):
    """
    Envia uma pagina para o modelo de visao e devolve regioes no mesmo
    formato que vision_cv.analisar_imagem produz (lista de Regiao + relatorio),
    para que rebuild.py nao precise saber qual motor rodou.

    Faz fallback silencioso para o motor convencional em qualquer falha,
    conforme RNF24 do documento de requisitos -- a IA nunca impede a
    conversao, so a melhora quando disponivel.
    """
    from .vision_cv import analisar_imagem  # fallback

    chave = _chave()
    if not chave:
        return analisar_imagem(imagem_bgr)

    try:
        import base64
        import cv2
        import anthropic

        ok, buf = cv2.imencode(".png", imagem_bgr)
        if not ok:
            raise RuntimeError("falha ao codificar imagem para a IA")
        b64 = base64.b64encode(buf.tobytes()).decode()

        cliente = anthropic.Anthropic(api_key=chave)
        resposta = cliente.messages.create(
            model=MODELO,
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64",
                     "media_type": "image/png", "data": b64}},
                    {"type": "text", "text": (
                        "Liste as regioes visuais desta pagina de slide/documento "
                        "como JSON: uma lista de objetos com x, y, w, h (pixels), "
                        "tipo (texto|foto|logo|icone|forma), texto (se houver) e "
                        "cor_dominante (hex). So o JSON, sem explicacao.")},
                ],
            }],
        )
        texto = resposta.content[0].text
        dados = json.loads(texto[texto.find("["):texto.rfind("]") + 1])

        from .vision_cv import Regiao
        regioes = [Regiao(x=int(d["x"]), y=int(d["y"]), w=int(d["w"]), h=int(d["h"]),
                          tipo=d.get("tipo", "indefinido"), texto=d.get("texto", ""),
                          conf=95.0, cor_dominante=d.get("cor_dominante"))
                  for d in dados]
        rel = {"motor": "ia", "regioes_por_tipo": {}}
        for r in regioes:
            rel["regioes_por_tipo"][r.tipo] = rel["regioes_por_tipo"].get(r.tipo, 0) + 1
        return regioes, rel

    except Exception as e:
        # Fallback obrigatorio (RNF24): nunca deixar a conversao travar
        # porque a IA falhou (rede, chave invalida, limite de conta, etc).
        regioes, rel = analisar_imagem(imagem_bgr)
        rel["ia_falhou"] = str(e)
        return regioes, rel
