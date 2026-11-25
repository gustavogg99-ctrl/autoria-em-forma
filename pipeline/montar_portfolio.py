"""Monta amostra de portfólio anonimizada a partir de um PDF (06, "Amostras de portfólio").

- Só com autorização do autor. Usar o ORIGINAL e a ÚLTIMA versão aprovada, nunca intermediária.
- Redação REAL: o texto sob a tarja é apagado do PDF (não é sobreposição).
- Remove metadados, links, anexos, JavaScript e texto oculto (scrub).
- Confere por busca no texto final que nenhum termo sobrou. Ocorrência dentro de outra
  palavra ("Candida" em "Candidaturas") sai como possível falso positivo: conferir à mão.

Arquivo de termos: um termo por linha (nome, sobrenome, orientador, banca, tema identificável,
cidade, empresa...). Liste cada parte do nome separada também ("Maria", "Silva"): um nome que
quebra de linha no PDF não é encontrado como expressão inteira.

Uso:
  python3 -I pipeline/montar_portfolio.py entrada.pdf termos.txt saida.pdf [--paginas 1-3,12] [--marca "AMOSTRA"]

Código de saída: 0 = limpo; 1 = sobrou termo (não publicar).
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_seguro import normalizar, sem_acento  # noqa: E402


def ler_termos(caminho: str) -> list[str]:
    termos = []
    for linha in Path(caminho).read_text(encoding="utf-8").splitlines():
        t = linha.strip()
        if t and not t.startswith("#"):
            termos.append(t)
    variantes = set(termos)
    for t in termos:
        variantes.add(sem_acento(t))
    return sorted(variantes, key=len, reverse=True)


def faixa_paginas(spec: str | None, total: int) -> list[int]:
    if not spec:
        return list(range(total))
    paginas = []
    for parte in spec.split(","):
        if "-" in parte:
            a, b = parte.split("-")
            paginas.extend(range(int(a) - 1, int(b)))
        else:
            paginas.append(int(parte) - 1)
    return [p for p in paginas if 0 <= p < total]


def redigir(entrada: str, termos: list[str], saida: str, paginas: str | None, marca: str | None) -> None:
    doc = pymupdf.open(entrada)
    doc.select(faixa_paginas(paginas, doc.page_count))
    for pagina in doc:
        for termo in termos:
            for area in pagina.search_for(termo):  # busca sem diferenciar maiúsculas
                pagina.add_redact_annot(area, fill=(0, 0, 0))
        pagina.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS)
        if marca:
            r = pagina.rect
            pagina.insert_text((r.width * 0.12, r.height * 0.6), marca, fontsize=60, rotate=0,
                               color=(0.8, 0.8, 0.8), fill_opacity=0.35, overlay=True)
    doc.scrub()
    doc.set_metadata({})
    doc.del_xml_metadata()
    doc.save(saida, garbage=4, deflate=True, clean=True)


def conferir(saida: str, termos: list[str]) -> tuple[list[str], list[str]]:
    doc = pymupdf.open(saida)
    texto = normalizar(" ".join(p.get_text() for p in doc))
    base = sem_acento(texto).lower()
    sobrou, falsos = [], []
    for termo in {sem_acento(t).lower() for t in termos}:
        for m in re.finditer(re.escape(termo), base):
            ini, fim = m.start(), m.end()
            dentro = (ini > 0 and base[ini - 1].isalnum()) or (fim < len(base) and base[fim].isalnum())
            contexto = texto[max(0, ini - 30):fim + 30]
            (falsos if dentro else sobrou).append(f"{termo!r}: ...{contexto}...")
    return sobrou, falsos


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("entrada")
    ap.add_argument("termos")
    ap.add_argument("saida")
    ap.add_argument("--paginas", help="ex.: 1-3,12 (padrão: todas)")
    ap.add_argument("--marca", help="marca d'água, ex.: AMOSTRA")
    a = ap.parse_args()
    termos = ler_termos(a.termos)
    redigir(a.entrada, termos, a.saida, a.paginas, a.marca)
    sobrou, falsos = conferir(a.saida, termos)
    for f in falsos:
        print("CONFERIR À MÃO (dentro de outra palavra):", f)
    if sobrou:
        for s in sobrou:
            print("SOBROU:", s)
        print(f"FALHOU: {len(sobrou)} ocorrência(s). Não publicar.")
        return 1
    print(f"OK: {a.saida} sem os {len(termos)} termos (com variantes sem acento).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
