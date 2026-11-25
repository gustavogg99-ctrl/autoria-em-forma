"""Etapa 4 do fluxo (06, seção 4): proteger trechos que nunca podem mudar.

Extrai do ORIGINAL, antes de qualquer alteração:
- citações longas (recuo esquerdo >= 4 cm ou estilo de citação);
- citações curtas entre aspas seguidas de chamada com página;
- tabelas (dados de pesquisa);
- tudo depois de um marcador (ex.: "APÊNDICE B", transcrições), com --proteger-apos;
- com --todas-aspas: todo trecho entre aspas duplas, com ou sem chamada logo depois
  (útil quando o corpus é citado em pedaços curtos, como manchetes e falas);
- com --comecar-em: ignora tudo antes do parágrafo cujo texto é exatamente o informado
  (evita falso positivo de folha de rosto e sumário digitado com recuo grande).

Grava um JSON com o texto e o hash SHA-256 de cada trecho. O teste
testes/verificar_hashes.py compara esse JSON com a versão final.

Uso:
  python3 -I pipeline/proteger_trechos.py original.docx trechos.json [--proteger-apos "APÊNDICE B"]
      [--todas-aspas] [--comecar-em "1. INTRODUÇÃO"] [--parar-em "9. REFERÊNCIAS"]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_seguro import (NS, RECUO_CITACAO_LONGA, abrir, corpo, mapa_estilos,  # noqa: E402
                         nome_estilo, normalizar, recuo_esquerdo, texto_paragrafo, w)

# "texto entre aspas" seguido, em até 3 caracteres, de chamada com ano e página
CITACAO_CURTA = re.compile(r'"([^"]{10,}?)"\s{0,3}(?=\([^()]*\d{4}[a-z]?[^()]*p\.)')
# qualquer trecho entre aspas duplas (normalizadas), com pelo menos 3 caracteres
ENTRE_ASPAS = re.compile(r'"([^"]{3,}?)"')


def h(texto: str) -> str:
    return hashlib.sha256(normalizar(texto).encode("utf-8")).hexdigest()


def texto_tabela(tbl) -> str:
    linhas = []
    for tr in tbl.findall("w:tr", NS):
        celulas = []
        for tc in tr.findall("w:tc", NS):
            celulas.append(" ".join(texto_paragrafo(p, "original") for p in tc.iter(w("p"))))
        linhas.append(" | ".join(celulas))
    return "\n".join(linhas)


def extrair(caminho: str, proteger_apos: list[str], todas_aspas: bool = False,
            comecar_em: str | None = None, parar_em: str | None = None) -> list[dict]:
    partes = abrir(caminho)
    estilos = mapa_estilos(partes)
    body = corpo(partes)
    trechos: list[dict] = []
    marcadores = [normalizar(m).upper() for m in proteger_apos]
    protegendo = False
    comecou = comecar_em is None
    alvo_inicio = normalizar(comecar_em) if comecar_em else None
    alvo_fim = normalizar(parar_em) if parar_em else None

    def add(tipo: str, texto: str, origem: str) -> None:
        t = normalizar(texto)
        if len(t) >= 3:
            trechos.append({"id": f"T{len(trechos) + 1:03d}", "tipo": tipo, "origem": origem,
                            "texto": t, "sha256": h(t)})

    for i, el in enumerate(body):
        tag = el.tag.split("}")[-1]
        if not comecou:
            if tag == "p" and normalizar(texto_paragrafo(el, "original")) == alvo_inicio:
                comecou = True
            else:
                continue
        if alvo_fim and tag == "p" and normalizar(texto_paragrafo(el, "original")) == alvo_fim:
            break
        if tag == "tbl":
            add("tabela", texto_tabela(el), f"tabela no bloco {i}")
            continue
        if tag != "p":
            continue
        txt = texto_paragrafo(el, "original")
        n = normalizar(txt).upper()
        if marcadores and any(n.startswith(m) for m in marcadores):
            protegendo = True
        if protegendo:
            add("apos-marcador", txt, f"parágrafo {i}")
            continue
        estilo = nome_estilo(el, estilos).lower()
        if recuo_esquerdo(el, estilos) >= RECUO_CITACAO_LONGA or "cita" in estilo or "quote" in estilo:
            add("citacao-longa", txt, f"parágrafo {i}")
            continue
        vistos = set()
        for m in CITACAO_CURTA.finditer(normalizar(txt)):
            add("citacao-curta", m.group(1), f"parágrafo {i}")
            vistos.add(m.group(1))
        if todas_aspas:
            for m in ENTRE_ASPAS.finditer(normalizar(txt)):
                if m.group(1) not in vistos:
                    add("entre-aspas", m.group(1), f"parágrafo {i}")
                    vistos.add(m.group(1))
    if comecar_em and not comecou:
        raise SystemExit(f"Marcador --comecar-em não encontrado: {comecar_em!r}")
    return trechos


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("original")
    ap.add_argument("saida_json")
    ap.add_argument("--proteger-apos", action="append", default=[],
                    help="protege tudo a partir do parágrafo que começa com este texto (repetível)")
    ap.add_argument("--todas-aspas", action="store_true",
                    help="protege todo trecho entre aspas duplas, mesmo sem chamada com página")
    ap.add_argument("--comecar-em", help="ignora o que vem antes do parágrafo com este texto exato")
    ap.add_argument("--parar-em", help="para no parágrafo com este texto exato (ex.: título das referências)")
    a = ap.parse_args()
    trechos = extrair(a.original, a.proteger_apos, a.todas_aspas, a.comecar_em, a.parar_em)
    Path(a.saida_json).write_text(json.dumps({"arquivo": Path(a.original).name, "trechos": trechos},
                                             ensure_ascii=False, indent=2), encoding="utf-8")
    por_tipo: dict[str, int] = {}
    for t in trechos:
        por_tipo[t["tipo"]] = por_tipo.get(t["tipo"], 0) + 1
    print(f"{len(trechos)} trechos protegidos: {por_tipo}")
    print("Confira à mão: citação sem recuo e sem chamada com página não é detectada; acrescente com --proteger-apos.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
