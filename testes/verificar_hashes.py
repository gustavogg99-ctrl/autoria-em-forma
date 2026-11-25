"""Teste bloqueante (regra de ouro 1, E1): trechos protegidos idênticos ao original.

Compara o JSON gerado por pipeline/proteger_trechos.py (sobre o ORIGINAL) com a versão
final (limpa, ou a marcada lida com as alterações aceitas). Cada citação, transcrição e
tabela precisa aparecer com o texto idêntico (normalizado: espaços, aspas tipográficas).
Inserção de "[sic]" é permitida (regra de ouro 1) e não conta como alteração.

Uso:
  python3 -I testes/verificar_hashes.py trechos.json final.docx
Código de saída: 0 = todos intactos; 1 = bloqueia a entrega.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
from docx_seguro import NS, abrir, corpo, normalizar, texto_paragrafo, w  # noqa: E402
from proteger_trechos import texto_tabela  # noqa: E402

SIC = re.compile(r"\s?\[sic\]", re.IGNORECASE)


def tirar_sic(t: str) -> str:
    return normalizar(SIC.sub("", t))


def texto_final(caminho: str) -> tuple[str, list[str]]:
    body = corpo(abrir(caminho))
    paragrafos = [texto_paragrafo(p, "final") for p in body.iter(w("p"))]
    tabelas = [normalizar(_tabela_final(t)) for t in body.iter(w("tbl"))]
    return tirar_sic(" ".join(paragrafos)), [tirar_sic(t) for t in tabelas]


def _tabela_final(tbl) -> str:
    linhas = []
    for tr in tbl.findall("w:tr", NS):
        if tr.find("w:trPr/w:del", NS) is not None:
            continue
        linhas.append(" | ".join(" ".join(texto_paragrafo(p, "final") for p in tc.iter(w("p")))
                                 for tc in tr.findall("w:tc", NS)))
    return "\n".join(linhas)


def main() -> int:
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    dados = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    texto, tabelas = texto_final(sys.argv[2])
    alterados = []
    for t in dados["trechos"]:
        alvo = tirar_sic(t["texto"])
        achou = alvo in tabelas if t["tipo"] == "tabela" else alvo in texto
        if not achou:
            alterados.append(t)
    print(f"{len(dados['trechos']) - len(alterados)}/{len(dados['trechos'])} trechos intactos.")
    if alterados:
        print("FALHOU — trecho protegido alterado ou ausente. Entrega bloqueada:")
        for t in alterados:
            print(f"  - {t['id']} ({t['tipo']}, {t['origem']}): {t['texto'][:90]!r}")
        return 1
    print("OK — nenhuma citação, transcrição ou tabela foi alterada.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
