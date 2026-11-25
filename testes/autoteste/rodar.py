"""Autoteste do pipeline: roda cada script contra os documentos fictícios e confere
se passa onde deve passar e FALHA onde deve falhar.

Rodar sempre que um script mudar (06, seção 4, passo 7: suíte de regressão).

Uso (da raiz do repositório):
  python3 -I testes/autoteste/rodar.py
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "pipeline"))
from docx_seguro import abrir, corpo, texto_paragrafo  # noqa: E402
PY = [sys.executable, "-I"]
resultados: list[tuple[bool, str]] = []


def rodar(script: str, *args, esperado: int, contem: list[str] | None = None, nome: str = "") -> str:
    proc = subprocess.run(PY + [str(RAIZ / script), *map(str, args)], capture_output=True, text=True, timeout=600)
    saida = proc.stdout + proc.stderr
    ok = proc.returncode == esperado and all(c in saida for c in (contem or []))
    resultados.append((ok, nome or script))
    if not ok:
        print(f"\n--- FALHA: {nome} (código {proc.returncode}, esperado {esperado})\n{saida}")
    return saida


def main() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        d = Path(tmp)
        rodar("testes/autoteste/gerar_fixtures.py", d, esperado=0, nome="gerar fixtures")
        j = d / "trechos.json"

        rodar("pipeline/proteger_trechos.py", d / "original.docx", j, esperado=0,
              contem=["citacao-longa", "citacao-curta", "tabela"], nome="proteger_trechos detecta os 3 tipos")

        j2 = d / "trechos_aspas.json"
        rodar("pipeline/proteger_trechos.py", d / "original.docx", j2, "--todas-aspas", "--comecar-em", "1 INTRODUÇÃO",
              esperado=0, contem=["citacao-longa"], nome="proteger_trechos --todas-aspas/--comecar-em ignora sumário")
        rodar("pipeline/proteger_trechos.py", d / "original.docx", j2, "--comecar-em", "NÃO EXISTE", esperado=1,
              contem=["não encontrado"], nome="proteger_trechos recusa marcador inexistente")

        rodar("testes/autoteste/caso_montar.py", d, "ok", esperado=0, nome="montar_docx monta a marcada do zero")
        rodar("testes/verificar_marcada.py", d / "marcada_motor.docx", esperado=0, contem=["[REGRA]"],
              nome="montar_docx: autor e prefixos no padrão")
        rodar("pipeline/aceitar_alteracoes.py", d / "marcada_motor.docx", d / "limpa_motor.docx", esperado=0,
              nome="montar_docx: aceitar gera a limpa")
        rodar("testes/verificar_limpa.py", d / "limpa_motor.docx", esperado=0, nome="montar_docx: limpa sem marcas")
        rodar("testes/verificar_hashes.py", j, d / "limpa_motor.docx", esperado=0,
              nome="montar_docx: citações e tabela intactas")
        rodar("testes/autoteste/caso_montar.py", d, "sem_destino", esperado=2, contem=["sem destino"],
              nome="montar_docx recusa perder parágrafo do original")
        rodar("testes/autoteste/caso_montar.py", d, "sem_prefixo", esperado=2, contem=["sem prefixo"],
              nome="montar_docx recusa comentário sem prefixo")

        # junção de dois parágrafos como alteração rastreada (F21 do teste IFOOD-05)
        rodar("testes/autoteste/caso_montar.py", d, "juntar", esperado=0, nome="montar_docx junta dois parágrafos (marcada)")
        rodar("testes/verificar_marcada.py", d / "marcada_motor.docx", esperado=0, contem=["[DECIDIR]"],
              nome="montar_docx: marcada com junção no padrão")
        rodar("pipeline/aceitar_alteracoes.py", d / "marcada_motor.docx", d / "limpa_juntar.docx", esperado=0,
              contem=["paragrafos_unidos': 1"], nome="aceitar une os dois parágrafos")
        rodar("testes/verificar_limpa.py", d / "limpa_juntar.docx", esperado=0, nome="junção: limpa sem marcas")
        rodar("testes/verificar_hashes.py", j, d / "limpa_juntar.docx", esperado=0, nome="junção: citações e tabela intactas")
        textos = [texto_paragrafo(e, "original") for e in corpo(abrir(str(d / "limpa_juntar.docx")))]
        resultados.append((any(t.strip() == "Para Souza (2019), a circulação dependia do correio imperial." for t in textos)
                           and not any(t.strip() == "do correio imperial." for t in textos),
                           "junção: um só parágrafo, com um espaço entre os textos"))
        rodar("testes/autoteste/caso_montar.py", d, "juntar_errado", esperado=2, contem=["pede junção"],
              nome="montar_docx recusa junção com parágrafo que não é o próximo")

        # exclusão de parágrafo inteiro como alteração rastreada (condição B do IFOOD-05, SOUZA 2019b)
        rodar("testes/autoteste/caso_montar.py", d, "excluir", esperado=0, nome="montar_docx exclui parágrafo (marcada)")
        rodar("testes/verificar_marcada.py", d / "marcada_motor.docx", esperado=0, contem=["[REGRA]"],
              nome="exclusão: marcada no padrão")
        rodar("pipeline/aceitar_alteracoes.py", d / "marcada_motor.docx", d / "limpa_excluir.docx", esperado=0,
              nome="exclusão: aceitar remove o parágrafo")
        rodar("testes/verificar_limpa.py", d / "limpa_excluir.docx", esperado=0, nome="exclusão: limpa sem marcas")
        orig = [texto_paragrafo(e, "original") for e in corpo(abrir(str(d / "original.docx")))]
        lim_ex = [texto_paragrafo(e, "original") for e in corpo(abrir(str(d / "limpa_excluir.docx")))]
        lim_ok = [texto_paragrafo(e, "original") for e in corpo(abrir(str(d / "limpa_motor.docx")))]
        resultados.append((orig[33].strip() not in [t.strip() for t in lim_ex] and len(lim_ex) == len(lim_ok) - 1
                           and orig[34].strip() in [t.strip() for t in lim_ex],
                           "exclusão: só a referência marcada sai, a seguinte fica"))

        rodar("testes/verificar_marcada.py", d / "marcada_ok.docx", esperado=0,
              contem=["[REGRA]", "[PENDENTE]"], nome="verificar_marcada aprova o padrão")
        rodar("testes/verificar_marcada.py", d / "marcada_ruim.docx", esperado=1,
              contem=["GG – Revisão", "sem prefixo"], nome="verificar_marcada pega autor e prefixo (E14b)")

        rodar("pipeline/aceitar_alteracoes.py", d / "marcada_ok.docx", d / "limpa.docx", esperado=0,
              contem=["paragrafos_unidos': 1"], nome="aceitar_alteracoes gera a limpa")
        rodar("pipeline/aceitar_alteracoes.py", d / "marcada_ruim.docx", d / "limpa_ruim.docx", esperado=0,
              nome="aceitar_alteracoes na marcada ruim")

        rodar("testes/verificar_limpa.py", d / "limpa.docx", esperado=0, nome="verificar_limpa aprova a limpa")
        rodar("testes/verificar_limpa.py", d / "marcada_ok.docx", esperado=1,
              contem=["w:ins", "comments.xml", "trackRevisions"], nome="verificar_limpa pega marcada entregue como limpa (E14a)")

        rodar("testes/verificar_hashes.py", j, d / "limpa.docx", esperado=0, nome="verificar_hashes aprova trechos intactos")
        rodar("testes/verificar_hashes.py", j, d / "limpa_ruim.docx", esperado=1,
              contem=["citacao-longa"], nome="verificar_hashes pega data trocada na citação (E1)")
        rodar("testes/verificar_hashes.py", j, d / "marcada_ok.docx", esperado=0,
              nome="verificar_hashes lê a marcada na visão final")

        saida = rodar("testes/cruzar_citacoes.py", d / "limpa.docx", "--edicao", "2002", esperado=1,
                      contem=["COSTA, 2021", "TAVARES"], nome="cruzar_citacoes acha chamada e referência órfãs")
        resultados.append(("2 item(ns)" in saida, "cruzar_citacoes sem falso positivo (só os 2 órfãos)"))
        rodar("testes/cruzar_citacoes.py", d / "limpa.docx", "--edicao", "2023", esperado=1,
              contem=["padrão 2023", "ponto antes da chamada (2023)"], nome="cruzar_citacoes pega edição misturada (E3)")

        rodar("pipeline/renderizar.py", d / "limpa.docx", d, esperado=0, nome="renderizar limpa")
        rodar("pipeline/renderizar.py", d / "sem_numero.docx", d, esperado=0, nome="renderizar sem_numero")
        rodar("testes/verificar_paginacao.py", d / "limpa.pdf", esperado=0, nome="verificar_paginacao aprova (detecção automática)")
        rodar("testes/verificar_paginacao.py", d / "sem_numero.pdf", esperado=1,
              contem=["mostra nada"], nome="verificar_paginacao pega número ausente (E2)")

        termos = d / "termos.txt"
        termos.write_text("Fulana\nExemplar\nUNIVERSIDADE FICTÍCIA DO PLANALTO\n", encoding="utf-8")
        rodar("pipeline/montar_portfolio.py", d / "limpa.pdf", termos, d / "amostra.pdf", "--marca", "AMOSTRA",
              esperado=0, contem=["OK"], nome="montar_portfolio redige e confere")

        recusado = d / "macro.docm"
        recusado.write_bytes((d / "original.docx").read_bytes())
        rodar("testes/verificar_limpa.py", recusado, esperado=1, contem=["só .docx"], nome="recusa .docm (06, seção 8)")

    falhas = [n for ok, n in resultados if not ok]
    for ok, n in resultados:
        print(("  ok    " if ok else "  FALHA ") + n)
    print(f"\n{len(resultados) - len(falhas)}/{len(resultados)} verificações passaram.")
    return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
