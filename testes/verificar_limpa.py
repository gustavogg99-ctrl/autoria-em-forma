"""Teste bloqueante (E14): a versão LIMPA não pode ter nenhuma marcação nem comentário.

Falha se encontrar, em qualquer parte do .docx: w:ins, w:del, w:moveFrom, w:moveTo,
registros de formatação (*Change), commentRangeStart/End, commentReference, w:comment,
partes comments*.xml, ou w:trackRevisions ligado em settings.xml.

Uso:
  python3 -I testes/verificar_limpa.py limpa.docx
Código de saída: 0 = limpa; 1 = bloqueia a entrega.
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
from docx_seguro import NS, abrir, xml  # noqa: E402

PROIBIDOS = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
             "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange",
             "commentRangeStart", "commentRangeEnd", "commentReference", "comment",
             "cellIns", "cellDel", "cellMerge")


def verificar(caminho: str) -> list[str]:
    partes = abrir(caminho)
    problemas: list[str] = []
    for nome, dados in partes.items():
        if not nome.endswith(".xml") or not nome.startswith("word/"):
            continue
        if re.match(r"^word/comments(Extended|Ids|Extensible)?\.xml$", nome):
            problemas.append(f"{nome}: parte de comentários presente")
            continue
        raiz = xml(dados)
        cont = Counter()
        for tag in PROIBIDOS:
            n = len(raiz.findall(f".//w:{tag}", NS))
            if n:
                cont[tag] = n
        if cont:
            problemas.append(f"{nome}: " + ", ".join(f"w:{k}={v}" for k, v in sorted(cont.items())))
        if nome == "word/settings.xml" and raiz.find("w:trackRevisions", NS) is not None:
            problemas.append("word/settings.xml: w:trackRevisions ligado (edição do cliente sairia marcada)")
    return problemas


def main() -> int:
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    problemas = verificar(sys.argv[1])
    if problemas:
        print("FALHOU — versão limpa com marcas ou comentários (E14). Entrega bloqueada:")
        for p in problemas:
            print("  -", p)
        return 1
    print("OK — versão limpa: zero marcações e zero comentários.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
