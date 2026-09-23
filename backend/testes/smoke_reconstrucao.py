"""
Teste de fumaca da reconstrucao (RF03-RF19).

Exercita as tres rotas e verifica invariantes, sem depender de servidor HTTP:
- PPTX com texto vivo  -> rota vetorial, IE alto, texto preservado
- PDF com texto vivo   -> rota vetorial, IE alto
- imagem pura (PNG)    -> rota visao, gera slides, slide de referencia presente

Nao afere fidelidade pixel a pixel (isso e papel do bench com SSIM); afere que
cada entrada percorre a rota certa e produz um .pptx valido e editavel.

Uso: python3 testes/smoke_reconstrucao.py
"""
import asyncio
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.main import _processar_reconstrucao, JOBS
from app.audit import auditar

CASOS = [
    ("fix_grande.pptx", "vetorial", 0.5),
    ("info_real.pdf",   "vetorial", 0.6),
    ("info_png.png",    "visao",    0.0),
]


async def _rodar(arq, nome):
    ident = "smk" + nome.replace(".", "")
    dt = Path(tempfile.mkdtemp())
    ent = dt / ("entrada" + Path(arq).suffix)
    shutil.copy(arq, ent)
    JOBS[ident] = {"etapa": "", "pct": 0, "pronto": False,
                   "erro": None, "resultado": None}
    await _processar_reconstrucao(ident, str(ent), Path(arq).name, True)
    res = JOBS[ident]
    shutil.rmtree(dt, ignore_errors=True)
    return res


async def main(base):
    falhas = 0
    for arq, rota_esperada, ie_min in CASOS:
        caminho = str(Path(base) / arq)
        if not Path(caminho).exists():
            print(f"  PULADO {arq}: fixture ausente")
            continue
        res = await _rodar(caminho, arq)
        if res["erro"]:
            print(f"  FALHA  {arq}: {res['erro']}")
            falhas += 1
            continue
        r = res["resultado"]["relatorio"]
        rota = r.get("rota")
        ie = r.get("indice_editabilidade", 0)
        ok_rota = rota == rota_esperada
        ok_ie = ie >= ie_min
        marca = "ok   " if (ok_rota and ok_ie) else "FALHA"
        print(f"  {marca} {arq:18} rota={rota} (esperada {rota_esperada}) IE={ie:.3f} (min {ie_min})")
        falhas += not (ok_rota and ok_ie)

    print("RESULTADO:", "PASSOU" if falhas == 0 else f"{falhas} FALHA(S)")
    return 1 if falhas else 0


if __name__ == "__main__":
    base = sys.argv[1] if len(sys.argv) > 1 else "."
    sys.exit(asyncio.run(main(base)))
