"""Ciclo do sumário (06, seção 4): monta a marcada, gera a limpa, renderiza, lê as páginas dos títulos e repete
até as páginas estabilizarem. Copiar para entrada/ junto da spec.

Uso:  python3 -I entrada/loop_sumario.py spec_ifood05 [pagina_inicial_do_texto_no_pdf=7]
A spec precisa ter montar(paginas=None) devolvendo o Montador. Grava saida/ifood05_MARCADO.docx,
saida/ifood05_LIMPA.docx e saida/ifood05_LIMPA.pdf (nomes fixos do teste; trocar para o trabalho real).
"""
import importlib
import json
import subprocess
import sys
from pathlib import Path

R = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(R / "pipeline"))
sys.path.insert(0, str(R / "entrada"))
from montar_docx import paginas_dos_titulos  # noqa: E402
from renderizar import renderizar  # noqa: E402

spec = importlib.import_module(sys.argv[1])
inicio = int(sys.argv[2]) if len(sys.argv) > 2 else 7
pag: dict = {}
for volta in range(1, 6):
    m = spec.montar(pag or None)
    m.salvar(R / "saida/ifood05_MARCADO.docx")
    subprocess.run([sys.executable, "-I", str(R / "pipeline/aceitar_alteracoes.py"), str(R / "saida/ifood05_MARCADO.docx"),
                    str(R / "saida/ifood05_LIMPA.docx")], check=True, capture_output=True)
    pdf = renderizar(R / "saida/ifood05_LIMPA.docx", R / "saida")
    novo = paginas_dos_titulos(pdf, m.titulos, deslocamento=1, a_partir_da_pagina=inicio)
    print("volta", volta, novo)
    if novo == pag:
        print("ESTÁVEL na volta", volta)
        break
    pag = novo
(R / "entrada/paginas.json").write_text(json.dumps(pag))
