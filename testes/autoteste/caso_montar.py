"""Caso de autoteste do motor pipeline/montar_docx.py, sobre o original fictício das fixtures.

Uso (chamado pelo rodar.py):
  python3 -I testes/autoteste/caso_montar.py pasta_das_fixtures {ok|sem_destino|sem_prefixo|juntar|juntar_errado|excluir}
- ok: monta a marcada com uma correção comentada, um parágrafo novo e sumário; grava marcada_motor.docx
- sem_destino: deixa um parágrafo do original de fora -> tem de falhar (texto seria perdido)
- sem_prefixo: comentário sem prefixo de 3.11 -> tem de falhar
- excluir: exclui uma referência do meio da lista (alteração rastreada) com comentário [REGRA]
- juntar: junta dois parágrafos do original (alteração rastreada); juntar_errado: o seguinte não é o próximo -> tem de falhar
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "pipeline"))
from montar_docx import ErroSpec, Montador  # noqa: E402


def montar(pasta: Path, variante: str) -> Montador:
    m = Montador(str(pasta / "original.docx"), criador="Fulana Exemplar de Teste")
    m.par(origem=0, estilo="Centralizado", comentario="[REGRA] Capa no modelo." if variante != "sem_prefixo"
          else "Capa no modelo.")
    for i in (1, 2, 3):
        m.par(origem=i, estilo="Centralizado")
    m.par(origem=5, estilo="Centralizado", sect=m.sect("rIdH1"))
    m.par(origem=6, estilo="Natureza")
    m.par(texto="Aprovado em: ____/____/______.", estilo="Normal", novo=True)
    m.par(origem=8, estilo="TituloPreTextual")
    m.par(origem=9, estilo="Resumo")
    m.descartar(*range(11, 16))
    m.sumario(sect=m.sect("rIdH1", inicio_pagina=1))
    m.par(origem=16, estilo="Heading1", nivel=1, ppr={"pageBreakBefore": '<w:pageBreakBefore w:val="0"/>'})
    m.par(origem=17, edicoes=[{"de": "Silva e Souza relata", "para": "Silva e Souza relatam",
                               "comentario": "[REGRA] Concordância: dois sujeitos, verbo no plural."}])
    if variante in ("juntar", "juntar_errado"):
        # 18 e 19 são uma frase partida em dois parágrafos na fixture
        m.par(origem=18, juntar_com_proximo=True,
              comentario="[DECIDIR] Parágrafos 18 e 19 unidos: a frase estava partida.")
        m.par(origem=19 if variante == "juntar" else 20)
        if variante == "juntar":
            m.par(origem=20)
    else:
        if variante != "sem_destino":
            m.par(origem=18)
        m.par(origem=19)
        m.par(origem=20)
    m.par(origem=22, estilo="Heading1", nivel=1)
    m.par(origem=23)
    m.par(origem=24)
    m.par(origem=25, estilo="Legenda")
    m.tabela_do_original(26)
    m.par(origem=27, estilo="FonteIlustracao")
    m.par(origem=29, estilo="Heading1", nivel=1, ppr={"jc": '<w:jc w:val="center"/>'})
    for i in range(30, 35):
        if variante == "excluir" and i == 33:
            m.par(origem=i, estilo="Referencias", excluir=True,
                  comentario="[REGRA] Referência removida a pedido do autor (não é citada no texto).")
        else:
            m.par(origem=i, estilo="Referencias")
    m.fechar_documento("rIdH2")
    return m


def main() -> int:
    pasta, variante = Path(sys.argv[1]), sys.argv[2]
    try:
        m = montar(pasta, variante)
        m.salvar(pasta / "marcada_motor.docx")
    except ErroSpec as e:
        print(f"ErroSpec: {e}")
        return 2
    print(f"ok: {len(m.comentarios)} comentários, {len(m.titulos)} títulos")
    return 0


if __name__ == "__main__":
    sys.exit(main())
