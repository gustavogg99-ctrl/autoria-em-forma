"""Teste bloqueante (06, 3.11 e E14): padrão da versão MARCADA.

- toda marcação (w:ins, w:del, movimentos, *Change) com autor "Revisão ABNT";
- todo comentário com autor "Revisão ABNT";
- todo comentário começando por [REGRA], [PENDENTE], [DECIDIR], [DIVERGÊNCIA] ou [FATO].
Também informa a contagem por prefixo (os [PENDENTE] e [DECIDIR] vão para o relatório 08).

Uso:
  python3 -I testes/verificar_marcada.py marcada.docx [--autor "Revisão ABNT"]
Código de saída: 0 = padrão OK; 1 = bloqueia a entrega.
"""
from __future__ import annotations

import argparse
import sys
import unicodedata
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
from docx_seguro import NS, abrir, partes_de_texto, texto_paragrafo, w, xml  # noqa: E402

PREFIXOS = ("[REGRA]", "[PENDENTE]", "[DECIDIR]", "[DIVERGÊNCIA]", "[FATO]")
MARCAS = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
          "tblPrChange", "trPrChange", "tcPrChange", "numberingChange")


def nfc(s: str | None) -> str:
    return unicodedata.normalize("NFC", s or "")


def verificar(caminho: str, autor: str) -> tuple[list[str], Counter, int]:
    partes = abrir(caminho)
    problemas: list[str] = []
    prefixos: Counter = Counter()
    autores_errados: Counter = Counter()
    n_marcas = 0
    for nome in partes_de_texto(partes):
        if nome == "word/comments.xml":
            continue
        raiz = xml(partes[nome])
        for tag in MARCAS:
            for el in raiz.iter(w(tag)):
                if tag in ("ins", "del") and el.getparent() is not None and el.getparent().tag == w("rPr") \
                        and el.getparent().getparent() is not None and el.getparent().getparent().tag != w("pPr"):
                    continue  # marcador de run dentro de rPr, já contado no contêiner
                n_marcas += 1
                if nfc(el.get(w("author"))) != autor:
                    autores_errados[nfc(el.get(w("author"))) or "(sem autor)"] += 1
    for a, n in autores_errados.items():
        problemas.append(f"{n} marcação(ões) com autor {a!r} (esperado {autor!r})")

    if "word/comments.xml" in partes:
        raiz = xml(partes["word/comments.xml"])
        for c in raiz.findall("w:comment", NS):
            cid = c.get(w("id"))
            if nfc(c.get(w("author"))) != autor:
                problemas.append(f"comentário {cid}: autor {nfc(c.get(w('author')))!r} (esperado {autor!r})")
            texto = nfc(" ".join(texto_paragrafo(p) for p in c.findall(".//w:p", NS))).strip()
            prefixo = next((p for p in PREFIXOS if texto.startswith(p)), None)
            if prefixo:
                prefixos[prefixo] += 1
            else:
                problemas.append(f"comentário {cid} sem prefixo: {texto[:70]!r}")
    return problemas, prefixos, n_marcas


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("marcada")
    ap.add_argument("--autor", default="Revisão ABNT")
    a = ap.parse_args()
    problemas, prefixos, n = verificar(a.marcada, nfc(a.autor))
    print(f"Marcações: {n} | Comentários por prefixo: {dict(prefixos) or 'nenhum'}")
    if n == 0:
        print("ATENÇÃO: nenhuma marcação. A versão marcada deveria mostrar o que mudou (06, 3.11).")
    if problemas:
        print("FALHOU — fora do padrão da versão marcada. Entrega bloqueada:")
        for p in problemas:
            print("  -", p)
        return 1
    print("OK — autor e prefixos no padrão.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
