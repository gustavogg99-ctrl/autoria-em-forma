"""Teste bloqueante (E2, 06 3.6): número de página no lugar certo, página a página.

Lê o PDF renderizado e procura um número no canto superior direito de cada página:
- antes da página inicial (Introdução): nenhum número;
- da página inicial em diante: número = página do PDF - deslocamento.
Deslocamento padrão 1: só a capa fica fora da contagem (a folha de rosto é a folha 1).

A página inicial pode ser dada (--inicio N) ou detectada: primeira página, fora do sumário,
cujo texto do topo começa por "INTRODUÇÃO" ou "1 INTRODUÇÃO".

Isto testa o PDF. O teste no programa do cliente (Word ou Google Docs) continua obrigatório (06, seção 5).

Uso:
  python3 -I testes/verificar_paginacao.py final.pdf [--inicio 8] [--deslocamento 1]
Código de saída: 0 = OK; 1 = bloqueia a entrega.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pdfplumber

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
from docx_seguro import sem_acento  # noqa: E402


def numero_topo_direito(page) -> str | None:
    w, h = page.width, page.height
    candidatos = [p for p in page.extract_words() if p["top"] < h * 0.12 and p["x0"] > w * 0.6
                  and re.fullmatch(r"\d{1,4}", p["text"])]
    if not candidatos:
        return None
    return max(candidatos, key=lambda p: p["x1"])["text"]


def detectar_inicio(pdf) -> int | None:
    for i, page in enumerate(pdf.pages):
        texto = sem_acento(page.extract_text() or "").upper()
        if "SUMARIO" in texto[:300] or re.search(r"\.{5,}|…{2,}", texto):
            continue
        linhas = [l.strip() for l in texto.splitlines() if l.strip() and not re.fullmatch(r"\d{1,4}", l.strip())]
        if linhas and re.match(r"^(1\s+)?INTRODUCAO\b", linhas[0]):
            return i + 1
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("--inicio", type=int, help="página do PDF (1-based) onde a numeração aparece")
    ap.add_argument("--deslocamento", type=int, default=1)
    a = ap.parse_args()
    erros = []
    with pdfplumber.open(a.pdf) as pdf:
        inicio = a.inicio or detectar_inicio(pdf)
        if not inicio:
            print("FALHOU — não achei a Introdução. Informe --inicio N.")
            return 1
        for i, page in enumerate(pdf.pages, start=1):
            achado = numero_topo_direito(page)
            if i < inicio:
                if achado is not None:
                    erros.append(f"pág. {i} do PDF: mostra {achado}, deveria estar sem número (pré-textual)")
            else:
                esperado = str(i - a.deslocamento)
                if achado != esperado:
                    erros.append(f"pág. {i} do PDF: mostra {achado or 'nada'}, esperado {esperado}")
        total = len(pdf.pages)
    print(f"Páginas: {total} | numeração a partir da pág. {inicio} do PDF, deslocamento {a.deslocamento}")
    if erros:
        print("FALHOU — numeração de página. Entrega bloqueada:")
        for e in erros:
            print("  -", e)
        return 1
    print("OK — número no canto superior direito em todas as páginas textuais e ausente nas pré-textuais.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
