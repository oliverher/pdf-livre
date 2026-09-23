"""
Teste de fumaca: verifica que a liberacao nao altera o visual.

Gera um deck com grupo aninhado e travas, roda a Trilha A, renderiza antes e
depois com LibreOffice e compara por SSIM. O criterio e >= 0.99: desagrupar e
destravar sao operacoes que NAO podem mover nada.

Requer: LibreOffice (soffice), poppler (pdftoppm), scikit-image, Pillow.
Uso: python3 testes/smoke.py
"""
import subprocess, sys, tempfile, shutil
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.audit import auditar
from app.unlock import liberar

LIMIAR = 0.99


def render(pptx, prefixo, td):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", td, pptx], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = Path(td) / (Path(pptx).stem + ".pdf")
    subprocess.run(["pdftoppm", "-png", "-r", "100", str(pdf),
                    str(Path(td) / prefixo)], check=True)
    return sorted(Path(td).glob(prefixo + "-*.png"))


def main(entrada):
    import numpy as np
    from PIL import Image
    from skimage.metrics import structural_similarity as ssim

    with tempfile.TemporaryDirectory() as td:
        saida = str(Path(td) / "saida.pptx")
        antes = auditar(entrada)
        ops = liberar(entrada, saida)
        depois = auditar(saida)

        print("operacoes:", ops)
        print(f"IE {antes['indice_editabilidade_global']} -> "
              f"{depois['indice_editabilidade_global']}")

        a = render(entrada, "a", td)
        b = render(saida, "b", td)
        assert len(a) == len(b), f"contagem de paginas mudou: {len(a)} != {len(b)}"

        falhas = 0
        for pa, pb in zip(a, b):
            ia = np.array(Image.open(pa).convert("L"))
            ib = np.array(Image.open(pb).convert("L").resize(ia.shape[::-1]))
            s = float(ssim(ia, ib))
            marca = "ok " if s >= LIMIAR else "FALHA"
            print(f"  {marca} {pa.name} SSIM={s:.5f}")
            falhas += s < LIMIAR

        # nenhum grupo nem trava pode restar
        grupos = sum(sl["objetos"].get("grupos", 0) for sl in depois["slides"])
        travas = sum(len(sl["travas"]) for sl in depois["slides"])
        print(f"  grupos restantes={grupos}  travas restantes={travas}")
        falhas += (grupos > 0) + (travas > 0)

        print("RESULTADO:", "PASSOU" if falhas == 0 else f"{falhas} FALHA(S)")
        return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "testes/fix_grupo.pptx"))
