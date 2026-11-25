"""Cruzamento chamada × referência (regra 7) e conferência da edição da NBR 10520.

Lê o .docx (visão final), separa o texto (até o título REFERÊNCIAS) e a lista de
referências (até APÊNDICE/ANEXO ou o fim) e reporta:
1. chamadas sem referência (e ano divergente quando o autor existe com outro ano);
2. referências não citadas;
3. mesmo autor e ano em duas referências sem letra (2024a/2024b);
4. com --edicao: grafia da chamada entre parênteses (2002 = MAIÚSCULAS; 2023 = Maiúsculas e
   minúsculas) e final das citações longas (2002 = "texto. (AUTOR, ano, p. x)."; 2023 =
   "texto (Autor, ano, p. x).") — nunca misturar edições (regra de ouro 5, E3).

É ferramenta de triagem: o cruzamento usa o primeiro sobrenome e o ano, então pode dar
falso positivo (sobrenome composto, entrada por título). Todo item reportado é conferido à
mão antes de virar pendência para o cliente. Nunca completar referência a partir daqui.

Uso:
  python3 -I testes/cruzar_citacoes.py arquivo.docx [--edicao 2002|2023]
Código de saída: 0 = nada a conferir; 1 = há itens para conferir.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
from docx_seguro import (RECUO_CITACAO_LONGA, abrir, corpo, mapa_estilos, normalizar,  # noqa: E402
                         recuo_esquerdo, sem_acento, texto_paragrafo, w)

ANO = r"(?:1[5-9]\d\d|20\d\d)[a-z]?"
SUFIXOS = {"JUNIOR", "JR", "FILHO", "NETO", "SOBRINHO", "SEGUNDO"}
PARTICULAS = {"DE", "DA", "DO", "DAS", "DOS", "E"}
TITULO_REF = re.compile(r"^(\d+\s+)?REFER[EÊ]NCIAS(\s+BIBLIOGR[AÁ]FICAS)?$", re.IGNORECASE)
TITULO_FIM = re.compile(r"^(AP[EÊ]NDICE|ANEXO|GLOSS[AÁ]RIO|[IÍ]NDICE)\b", re.IGNORECASE)
PARENTESES = re.compile(r"\(([^()]*?" + ANO + r"[^()]*)\)")
NARRATIVA = re.compile(
    r"([A-ZÀ-Ý][\wÀ-ÿ'’-]+(?:\s+[A-ZÀ-Ý][\wÀ-ÿ'’-]+)?)"   # sobrenome (+ sufixo)
    r"(?:\s+(?:e|&)\s+[A-ZÀ-Ý][\wÀ-ÿ'’-]+)?(?:\s+et\s+al\.?)?"
    r"\s*\((" + ANO + r")(?:,\s*p\.[^)]*)?\)")


def chave(nome: str) -> str:
    """Primeiro sobrenome em maiúsculas, sem acento e sem pontuação."""
    tokens = [re.sub(r"[^A-Z]", "", sem_acento(t).upper()) for t in nome.replace("...", " ").split()]
    tokens = [t for t in tokens if t and t not in PARTICULAS]
    return tokens[0] if tokens else ""


def separar(caminho: str):
    partes = abrir(caminho)
    estilos = mapa_estilos(partes)
    texto, refs, longas = [], [], []
    estado = "texto"
    for p in corpo(partes).iter(w("p")):
        t = normalizar(texto_paragrafo(p, "final"))
        if not t:
            continue
        if estado == "texto" and TITULO_REF.match(t):
            estado = "refs"
            continue
        if estado == "refs" and TITULO_FIM.match(t):
            estado = "fim"
            continue
        if estado == "texto":
            texto.append(t)
            if recuo_esquerdo(p, estilos) >= RECUO_CITACAO_LONGA:
                longas.append(t)
        elif estado == "refs":
            refs.append(t)
    return texto, refs, longas


def chamadas(paragrafos: list[str]) -> list[dict]:
    achadas = []
    for par in paragrafos:
        for m in PARENTESES.finditer(par):
            conteudo = m.group(1)
            if " apud " in conteudo:
                conteudo = conteudo.split(" apud ", 1)[1]   # só a fonte consultada vai às referências
            grupo: list[str] = []
            for seg in [s.strip() for s in conteudo.split(";")]:
                grupo.append(seg)
                ano = re.search(r"\b(" + ANO + r")\b", seg)
                if ano:
                    autor = grupo[0].split(",")[0]
                    if re.search(r"[A-Za-zÀ-ÿ]", autor) and not re.fullmatch(r"\s*" + ANO + r".*", autor):
                        achadas.append({"autor": autor.strip(), "chave": chave(autor), "ano": ano.group(1),
                                        "forma": "parenteses", "trecho": m.group(0)})
                    grupo = []
        for m in NARRATIVA.finditer(par):
            nome = m.group(1).split()
            if len(nome) == 2 and sem_acento(nome[1]).upper().strip(".") not in SUFIXOS:
                nome = nome[1:]   # "Segundo Silva" -> "Silva"
            achadas.append({"autor": " ".join(nome), "chave": chave(" ".join(nome)), "ano": m.group(2),
                            "forma": "narrativa", "trecho": m.group(0)})
    return [c for c in achadas if c["chave"]]


def referencias(linhas: list[str]) -> list[dict]:
    lista = []
    for r in linhas:
        entrada = re.split(r"[.,]", r, maxsplit=1)[0]
        corpo_ref = re.split(r"Acesso em:", r, flags=re.IGNORECASE)[0]
        anos = set(re.findall(r"\b(" + ANO + r")\b", corpo_ref))
        lista.append({"chave": chave(entrada), "anos": anos, "texto": r})
    return lista


def conferir_edicao(cits: list[dict], longas: list[str], edicao: str) -> list[str]:
    avisos = []
    for c in cits:
        if c["forma"] != "parenteses":
            continue
        letras = re.sub(r"[^A-Za-zÀ-ÿ]", "", c["autor"])
        caixa_alta = letras.isupper()
        if edicao == "2002" and not caixa_alta:
            avisos.append(f"chamada fora do padrão 2002 (MAIÚSCULAS): {c['trecho']}")
        if edicao == "2023" and caixa_alta and len(letras) > 5:
            avisos.append(f"chamada fora do padrão 2023 (Maiúsculas e minúsculas): {c['trecho']}")
    for t in longas:
        fim = re.search(r"(\S)\s*\(([^()]*" + ANO + r"[^()]*)\)\.?$", t)
        if not fim:
            continue
        ponto_antes = fim.group(1) == "."
        if edicao == "2002" and not ponto_antes:
            avisos.append(f"citação longa sem ponto antes da chamada (2002): ...{t[-60:]}")
        if edicao == "2023" and ponto_antes:
            avisos.append(f"citação longa com ponto antes da chamada (2023): ...{t[-60:]}")
    return avisos


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--edicao", choices=["2002", "2023"])
    a = ap.parse_args()
    texto, linhas_ref, longas = separar(a.docx)
    if not linhas_ref:
        print("FALHOU — não achei o título REFERÊNCIAS.")
        return 1
    cits, refs = chamadas(texto), referencias(linhas_ref)
    itens: list[str] = []

    por_chave = defaultdict(list)
    for r in refs:
        por_chave[r["chave"]].append(r)
    citadas = set()
    vistas = set()
    for c in cits:
        alvo = [r for r in por_chave.get(c["chave"], []) if c["ano"] in r["anos"]]
        if alvo:
            citadas.update(id(r) for r in alvo)
            continue
        k = (c["chave"], c["ano"])
        if k in vistas:
            continue
        vistas.add(k)
        outros = sorted({ano for r in por_chave.get(c["chave"], []) for ano in r["anos"]})
        if outros:
            itens.append(f"chamada {c['trecho']}: autor existe nas referências, mas com ano {', '.join(outros)}")
        else:
            itens.append(f"chamada sem referência: {c['trecho']}")
    for r in refs:
        if id(r) not in citadas:
            itens.append(f"referência não citada: {r['texto'][:100]}")
    grupos = defaultdict(list)
    for r in refs:
        for ano in r["anos"]:
            grupos[(r["chave"], ano)].append(r)
    for (k, ano), rs in grupos.items():
        if len(rs) > 1 and not re.search(r"[a-z]$", ano):
            itens.append(f"{k} {ano}: {len(rs)} referências do mesmo autor e ano sem letra (usar {ano}a, {ano}b)")
    avisos = conferir_edicao(cits, longas, a.edicao) if a.edicao else []

    print(f"Chamadas: {len(cits)} | Referências: {len(refs)} | Citações longas: {len(longas)}")
    for i in itens + avisos:
        print("  CONFERIR:", i)
    if itens or avisos:
        print(f"{len(itens) + len(avisos)} item(ns) para conferir à mão antes de virar pendência.")
        return 1
    print("OK — toda chamada tem referência e toda referência é citada.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
