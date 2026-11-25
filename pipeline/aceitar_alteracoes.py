"""Etapa 6 do fluxo (06, seção 4): gerar a versão LIMPA a partir da MARCADA, por XML.

Aceita todas as alterações e remove todos os comentários:
- w:ins / w:moveTo: mantém o conteúdo, tira a marca;
- w:del / w:moveFrom: remove com o conteúdo;
- marca de parágrafo excluída (w:pPr/w:rPr/w:del): junta o parágrafo ao seguinte;
- linha de tabela excluída: remove a linha;
- registros de formatação (*PrChange, numberingChange, tblGridChange): removidos (fica a nova);
- comentários: âncoras, referências e as partes comments*.xml, com rels e content types;
- settings.xml: desliga w:trackRevisions (senão a edição do cliente sai marcada).

Depois de rodar: renderizar, recalcular o sumário e rodar testes/verificar_limpa.py (E14).

Uso:
  python3 -I pipeline/aceitar_alteracoes.py marcada.docx limpa.docx
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_seguro import NS, abrir, partes_de_texto, salvar, serializar, w, xml  # noqa: E402

REGISTROS_FORMATACAO = ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange",
                        "tcPrChange", "tblGridChange", "tblPrExChange", "numberingChange")
PARTES_COMENTARIO = re.compile(r"^word/comments(Extended|Ids|Extensible)?\.xml$")


def _local(el) -> str:
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def _desembrulhar(el) -> None:
    pai = el.getparent()
    i = pai.index(el)
    for filho in list(el):
        pai.insert(i, filho)
        i += 1
    _preservar_tail(el)
    pai.remove(el)


def _remover(el) -> None:
    _preservar_tail(el)
    el.getparent().remove(el)


def _preservar_tail(el) -> None:
    # tail em WordprocessingML é só espaço em branco de indentação; nada a preservar.
    el.tail = None


def aceitar(raiz) -> dict[str, int]:
    c = {"ins": 0, "del": 0, "paragrafos_unidos": 0, "linhas_removidas": 0, "formatacao": 0, "comentarios": 0}

    # 1. linhas de tabela excluídas / inseridas
    for tr in list(raiz.iter(w("tr"))):
        trpr = tr.find("w:trPr", NS)
        if trpr is None:
            continue
        if trpr.find("w:del", NS) is not None:
            _remover(tr)
            c["linhas_removidas"] += 1
        for m in trpr.findall("w:ins", NS):
            trpr.remove(m)

    # 2. marcas de parágrafo excluídas: o conteúdo vai para o início do parágrafo seguinte
    for p in list(raiz.iter(w("p"))):
        marca = p.find("w:pPr/w:rPr/w:del", NS)
        if marca is None:
            continue
        marca.getparent().remove(marca)
        prox = p.getnext()
        if prox is None or _local(prox) != "p":
            continue  # último parágrafo do contêiner: só tira a marca
        conteudo = [f for f in p if _local(f) != "pPr"]
        ppr_prox = prox.find("w:pPr", NS)
        pos = 1 if ppr_prox is not None else 0
        for f in conteudo:
            prox.insert(pos, f)
            pos += 1
        _remover(p)
        c["paragrafos_unidos"] += 1
    for m in list(raiz.iterfind(".//w:pPr/w:rPr/w:ins", NS)):
        m.getparent().remove(m)

    # 3. exclusões e movimentos de origem
    for tag in ("del", "moveFrom"):
        for el in list(raiz.iter(w(tag))):
            if el.getparent() is not None and _local(el.getparent()) != "rPr":
                _remover(el)
                c["del"] += 1
    for el in list(raiz.iter(w("rPr"))):
        for m in el.findall("w:del", NS) + el.findall("w:ins", NS):
            el.remove(m)

    # 4. inserções e movimentos de destino
    for tag in ("ins", "moveTo"):
        for el in list(raiz.iter(w(tag))):
            _desembrulhar(el)
            c["ins"] += 1
    for tag in ("moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd"):
        for el in list(raiz.iter(w(tag))):
            _remover(el)

    # 5. registros de formatação
    for tag in REGISTROS_FORMATACAO:
        for el in list(raiz.iter(w(tag))):
            _remover(el)
            c["formatacao"] += 1

    # 6. comentários no corpo
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(raiz.iter(w(tag))):
            _remover(el)
    for ref in list(raiz.iter(w("commentReference"))):
        run = ref.getparent()
        _remover(ref)
        c["comentarios"] += 1
        if run is not None and _local(run) == "r" and all(_local(f) == "rPr" for f in run):
            _remover(run)
    return c


def limpar_pacote(partes: dict[str, bytes]) -> None:
    removidas = [n for n in partes if PARTES_COMENTARIO.match(n) or n == "word/people.xml"]
    for n in removidas:
        del partes[n]
    alvos = {n.split("/", 1)[1] for n in removidas}

    rels = "word/_rels/document.xml.rels"
    if rels in partes:
        raiz = xml(partes[rels])
        for r in list(raiz):
            if r.get("Target", "").lstrip("/").replace("word/", "") in alvos:
                raiz.remove(r)
        partes[rels] = serializar(raiz)

    ct = "[Content_Types].xml"
    raiz = xml(partes[ct])
    for o in list(raiz):
        if o.get("PartName", "").lstrip("/") in removidas:
            raiz.remove(o)
    partes[ct] = serializar(raiz)

    if "word/settings.xml" in partes:
        raiz = xml(partes["word/settings.xml"])
        for el in raiz.findall("w:trackRevisions", NS):
            raiz.remove(el)
        partes["word/settings.xml"] = serializar(raiz)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("marcada")
    ap.add_argument("limpa")
    a = ap.parse_args()
    partes = abrir(a.marcada)
    total: dict[str, int] = {}
    for nome in partes_de_texto(partes):
        if PARTES_COMENTARIO.match(nome):
            continue
        raiz = xml(partes[nome])
        for k, v in aceitar(raiz).items():
            total[k] = total.get(k, 0) + v
        partes[nome] = serializar(raiz)
    limpar_pacote(partes)
    salvar(partes, a.limpa)
    print(f"Versão limpa gravada: {a.limpa}")
    print("Aceitas:", total)
    print("Próximo: renderizar, recalcular o sumário e rodar testes/verificar_limpa.py.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
