#!/usr/bin/env python3
"""IFOOD-05 — gerador de documento FICTÍCIO para o teste cronometrado (à mão x pipeline).

Alvo: TCC fictício de Jornalismo/IESB com ~31 folhas, exportado "do Google Docs" (Letter, número no
rodapé) e com erros plantados de formatação, citação/referência e texto, mais armadilhas
(trechos protegidos com erro que NÃO podem mudar) e chamarizes (parecem erro, não são).

Uso: python3 -I gerar_fixture_ifood05.py --saida DIR [--seed 5] [--escala 2.0]
Grava DIR/ORIGINAL.docx e, em DIR/.privado/gabarito.json, o gabarito (NÃO mostrar antes do fim da Condição A).
Nenhum dado de cliente: autores, obras e conteúdo são inventados.
"""
from __future__ import annotations

import argparse
import json
import random
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw

GAB: list[dict] = []


def gab(id_, cat, desc, **kw):
    GAB.append({"id": id_, "categoria": cat, "descricao": desc, **kw})


# ----------------------------------------------------------------------------- referências
REFS = {  # chave: (texto da ref, destaque no título?)
    "GOMES": ("GOMES, Helena. {t}Tipografia e poder nas vilas do Império{/t}. Belo Horizonte: Editora Fictícia, 2016.", True),
    "ANDRADE": ("ANDRADE, Paulo; LIMA, Rita. {t}Leitores de província{/t}. Recife: Editora Fictícia, 2012.", True),
    "NOGUEIRA": ("NOGUEIRA, Camila. {t}Papel, tinta e política{/t}. Porto Alegre: Editora Fictícia, 2019.", True),
    "PEREIRA": ("PEREIRA, João. {t}Circulação de jornais no Império{/t}. São Paulo: Editora Fictícia, 2018.", False),
    "SILVA": ("SILVA, Ana; SOUZA, Beto. {t}Jornais de província{/t}. Brasília: Editora Fictícia, 2020.", True),
    "SOUZA1": ("SOUZA, Beto. {t}Correio e imprensa{/t}. Recife: Editora Fictícia, 2019.", True),
    "SOUZA2": ("SOUZA, Beto. {t}Censura e tipografia{/t}. Recife: Editora Fictícia, 2019.", True),
    "MENDES": ("MENDES, Luiza. {t}A opinião pública no século XIX{/t}. Rio de Janeiro: Editora Fictícia, 2008.", True),
    "CARVALHO": ("CARVALHO, Diego; FREITAS, Ana; LOPES, Marcos; REIS, Tânia. {t}Imprensa regional: uma história{/t}. Curitiba: Editora Fictícia, 2014.", True),
    "ARAUJO": ("ARAÚJO, Vera. {t}Oficinas e oficiais{/t}. Salvador: Editora Fictícia, 2005.", True),
    "MARQUES": ("MARQUES, Rui. {t}Vilas, folhas e leitores{/t}. Fortaleza: Editora Fictícia, 2011.", True),
    "FERRAZ": ("FERRAZ, Olavo. {t}Impressos provinciais{/t}. [s.l.]: Editora Fictícia.", True),
    "TAVARES": ("TAVARES, Carla. {t}Referência que ninguém citou{/t}. Natal: Editora Fictícia, 2015.", True),
}
# ordem alfabética correta; MENDES x MARQUES trocados de propósito (C08)
ORDEM_REF = ["ANDRADE", "ARAUJO", "CARVALHO", "FERRAZ", "GOMES", "MENDES", "MARQUES", "NOGUEIRA",
             "PEREIRA", "SILVA", "SOUZA1", "SOUZA2", "TAVARES"]
# chamadas corretas (MAIÚSCULAS, NBR 10520:2002) usadas no preenchimento
CHAMADAS = ["(GOMES, 2016)", "(ANDRADE; LIMA, 2012)", "(NOGUEIRA, 2019)", "(PEREIRA, 2018)", "(SILVA; SOUZA, 2020)",
            "(MENDES, 2008)", "(CARVALHO; FREITAS; LOPES; REIS, 2014)", "(ARAÚJO, 2005)", "(MARQUES, 2011)"]
NARRATIVAS = ["Gomes (2016)", "Andrade e Lima (2012)", "Nogueira (2019)", "Pereira (2018)", "Mendes (2008)",
              "Araújo (2005)", "Marques (2011)"]

SUJ = ["A circulação dos periódicos", "A produção dos jornais de província", "O trabalho dos tipógrafos",
       "A relação entre editores e leitores", "A distribuição das folhas impressas", "A presença da imprensa nas vilas",
       "O comércio de assinaturas", "A cobertura dos fatos políticos locais", "A disputa entre jornais rivais",
       "A escolha dos assuntos nas folhas", "O custo do papel importado", "A leitura coletiva nas praças"]
PRED = ["dependia de rotas de correio irregulares", "variava conforme a época do ano", "era marcada pela escassez de insumos",
        "refletia as tensões da política regional", "ocupava lugar central na vida urbana",
        "exigia negociações constantes com as autoridades locais", "se mostrava limitada pelo número reduzido de leitores",
        "mudou de forma gradual ao longo das décadas", "estava ligada à atuação de pequenos grupos de letrados",
        "dava forma a uma esfera pública ainda restrita", "acompanhava o ritmo das safras e das feiras",
        "revelava desigualdades entre o litoral e o sertão"]
COMPL = ["nas cidades do interior", "entre as décadas de 1830 e 1860", "nas capitais provinciais", "em regiões de difícil acesso",
         "nos períodos de maior instabilidade", "segundo os registros analisados", "na maior parte do período estudado",
         "nas vilas mais afastadas"]
LIGA = ["Além disso,", "Nesse sentido,", "Por outro lado,", "Em contrapartida,", "Do mesmo modo,", "Ainda assim,", "Com isso,"]
OUTRAS = ["esse movimento", "essa dinâmica", "esse arranjo", "tal configuração", "esse quadro"]


def frase(rng: random.Random, com_chamada: bool = True) -> str:
    k = rng.randint(0, 4)
    s, p, c = rng.choice(SUJ), rng.choice(PRED), rng.choice(COMPL)
    if k == 0:
        t = f"{s} {p} {c}"
    elif k == 1:
        t = f"Segundo {rng.choice(NARRATIVAS)}, {s[0].lower() + s[1:]} {p}"
    elif k == 2:
        t = f"{rng.choice(LIGA)} {s[0].lower() + s[1:]} {p} {c}"
    elif k == 3:
        t = f"{s} {p}, e {rng.choice(OUTRAS)} {rng.choice(PRED)}"
    else:
        t = f"{s} {p}"
    if com_chamada and rng.random() < 0.45 and k != 1:
        t += " " + rng.choice(CHAMADAS)
    return t + "."


def paragrafo_texto(rng: random.Random, n: int | None = None, extra: list[str] | None = None) -> str:
    n = n or rng.randint(5, 7)
    fr = [frase(rng) for _ in range(n)]
    for e in extra or []:
        fr.insert(rng.randint(1, len(fr) - 1), e)
    return " ".join(fr)


# ----------------------------------------------------------------------------- helpers docx
def fonte(obj, nome):
    obj.font.name = nome
    rpr = obj.element.get_or_add_rPr() if hasattr(obj, "element") and hasattr(obj.element, "get_or_add_rPr") else None
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        for a in list(rf.attrib):
            del rf.attrib[a]
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            rf.set(qn(f"w:{a}"), nome)


def run_font(run, nome):
    run.font.name = nome
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn(f"w:{a}"), nome)


def add_field(par, instr, valor="1"):
    for tipo, txt in (("begin", None), (None, instr), ("separate", None), (None, valor), ("end", None)):
        r = par.add_run()
        if tipo:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), tipo)
            r._r.append(f)
        elif txt == instr:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = instr
            r._r.append(it)
        else:
            r.text = txt


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2.54)  # F02 (Letter = F01)
        st = self.d.styles
        fonte(st["Normal"], "Arial")
        st["Normal"].font.size = Pt(12)
        for nome in ("Heading 1", "Heading 2"):
            h = st[nome]
            fonte(h, "Arial")
            h.font.size = Pt(12)
            h.font.color.rgb = RGBColor(0, 0, 0)
            h.font.bold = nome == "Heading 1"
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(12)
            h.paragraph_format.keep_with_next = True
        # F04/F05: número de página no RODAPÉ, em todas as páginas (inclusive capa)
        f = s.footer
        f.is_linked_to_previous = False
        p = f.paragraphs[0]
        p.alignment = AL.CENTER
        add_field(p, " PAGE ")
        cp = self.d.core_properties
        cp.author, cp.title, cp.last_modified_by = "Fulana Exemplar de Teste", "TCC fictício", "Fulana Exemplar de Teste"

    def p(self, texto="", *, al=AL.JUSTIFY, ls=1.15, first=1.25, left=None, size=None, bold=False, italic=False,
          before=0, after=6, font=None, style=None, break_before=False, keep=False):
        par = self.d.add_paragraph(style=style)
        pf = par.paragraph_format
        par.alignment = al
        pf.line_spacing = ls
        pf.first_line_indent = Cm(first) if first else None
        if left is not None:
            pf.left_indent = Cm(left)
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.page_break_before = break_before
        pf.keep_with_next = keep
        if texto:
            r = par.add_run(texto)
            r.bold, r.italic = bold, italic
            if size:
                r.font.size = Pt(size)
            if font:
                run_font(r, font)
        return par

    def ref(self, texto, destaque):
        par = self.p("", al=AL.JUSTIFY, ls=1.5, first=0, after=0)  # F11: justificada, 1,5, sem linha em branco
        pre, resto = texto.split("{t}")
        tit, pos = resto.split("{/t}")
        par.add_run(pre)
        r = par.add_run(tit)
        r.bold = destaque  # F24: destaque inconsistente
        par.add_run(pos)

    def br(self):
        self.d.add_page_break()


# ----------------------------------------------------------------------------- conteúdo
Q1 = ("Os tipógrafos da província eram obrigados a compor as páginas à luz de velas, e o papel, importado da Corte, "
      "chegava com meses de atrazo, o que tornava a periodicidade uma promessa raramente cumprida pelos editores.")
Q2 = ("A primeira tipografia da vila foi instalada em 1831, graças ao empenho de um comerciante que financiou a "
      "compra das caixas de tipos e contratou um oficial vindo da capital, de modo que a imprensa local nasceu "
      "ligada aos interesses de quem a sustentava.")
Q3 = "estava en crise todo o sistema de distribuição"
MANCHETES = ["Chegou o vapor com as folhas da Corte", "Pedese a atenção dos senhores assignantes",
             "Grande abalo no commercio da praça", "A typographia mudou de rua"]

ERROS_TEXTO = [  # id, frase errada, frase certa
    ("T01", "Os jornais da província circulava de forma irregular durante o inverno.", "Os jornais da província circulavam de forma irregular durante o inverno."),
    ("T02", "Esse quadro se agravou à partir de 1840, quando o preço do papel subiu.", "Esse quadro se agravou a partir de 1840, quando o preço do papel subiu."),
    ("T03", "A censura foi a única excessão ao regime de liberdade proclamado pelos editores.", "A censura foi a única exceção ao regime de liberdade proclamado pelos editores."),
    ("T04", "Haviam poucos tipógrafos qualificados nas vilas do interior.", "Havia poucos tipógrafos qualificados nas vilas do interior."),
    ("T05", "Os assinantes eram poucos, porisso a tiragem permanecia baixa.", "Os assinantes eram poucos, por isso a tiragem permanecia baixa."),
    ("T06", "A tiragem crescia a medida que a cidade se urbanizava.", "A tiragem crescia à medida que a cidade se urbanizava."),
    ("T07", "Os editores da época, defendiam a liberdade de imprensa com ressalvas.", "Os editores da época defendiam a liberdade de imprensa com ressalvas."),
    ("T08", "O Correio Provincial foi o jornal mas lido da região naquele período.", "O Correio Provincial foi o jornal mais lido da região naquele período."),
    ("T09", "Os dados da pesquisa mostra que a circulação dependia do correio.", "Os dados da pesquisa mostram que a circulação dependia do correio."),
    ("T10", "Os impressos eram enviados afim de alcançar leitores distantes.", "Os impressos eram enviados a fim de alcançar leitores distantes."),
    ("T11", "Fazem dois séculos que esses periódicos foram esquecidos pela historiografia.", "Faz dois séculos que esses periódicos foram esquecidos pela historiografia."),
    ("T12", "Concerteza, a circulação dos jornais variou entre as províncias.", "Com certeza, a circulação dos jornais variou entre as províncias."),
]
TXT = {t[0]: t for t in ERROS_TEXTO}

# ordem das seções. Itens: ('h1', texto, opts) ('h2', texto, opts) ('fill', n) ('sp', chave)
SEC = [
    ("h1", "1 INTRODUÇÃO", {"quebra": True}),
    ("fill", 3), ("sp", "T01"), ("fill", 2), ("sp", "fato1832"), ("sp", "objetivos"), ("fill", 3), ("sp", "T02"), ("fill", 2),
    ("h1", "2 REFERENCIAL TEÓRICO", {"quebra": True}),
    ("h2", "2.1 TIPOGRAFIA E PODER", {}),
    ("fill", 3), ("sp", "Q1"), ("fill", 2), ("sp", "T03"), ("sp", "C01a"), ("fill", 3), ("sp", "T04"), ("fill", 2),
    ("h2", "2.2 Leitores e alfabetização", {"errado": True}),  # F08
    ("fill", 3), ("sp", "Q3"), ("sp", "T05"), ("fill", 2), ("sp", "C01b"), ("sp", "C04"), ("fill_times", 1), ("fill", 3),
    ("h2", "2.3 CIRCULAÇÃO E CORREIO", {}),
    ("fill", 3), ("sp", "Q2"), ("sp", "T06"), ("fill", 2), ("sp", "C09"), ("sp", "C06a"), ("fill_espacos", 1), ("fill", 3),
    ("sp", "T07"), ("sp", "C01c"), ("fill", 2), ("sp", "C02"),
    ("h1", "3 METODOLOGIA", {"quebra": False}),  # F09
    ("h2", "3.1 PERCURSO DA PESQUISA", {}),
    ("fill", 3), ("sp", "T08"), ("fill", 2), ("sp", "C06b"), ("fill_espacos", 1), ("sp", "C07"), ("fill", 3),
    ("h2", "3.2 CORPUS E CRITÉRIOS", {}),
    ("fill", 3), ("sp", "D1"), ("sp", "T09"), ("fill_times", 1), ("fill", 3), ("sp", "C01d"),
    ("h1", "4 ANÁLISE DOS RESULTADOS", {"quebra": True}),
    ("h2", "4.1 PERFIL DOS JORNAIS", {}),
    ("fill", 2), ("sp", "tabela"), ("fill", 3), ("sp", "T10"), ("fill", 2), ("sp", "figura"), ("fill", 3),
    ("h2", "4.2 O CORPUS DE MANCHETES", {}),
    ("fill", 2), ("sp", "manchetes"), ("sp", "T11"), ("fill", 3), ("fill_espacos", 1), ("sp", "C01e"), ("fill", 3),
    ("h2", "4.3 DISCUSSÃO", {}),
    ("fill", 3), ("sp", "T12"), ("fill", 3), ("sp", "D3"), ("fill", 2),
    ("h1", "5. Considerações Finais", {"quebra": False, "errado": True}),  # F23
    ("fill", 4), ("sp", "fim"),
]


def construir(saida: Path, seed: int, escala: float) -> dict:
    rng = random.Random(seed)
    D = Doc()
    d = D.d

    # ---- capa
    D.p("INSTITUTO DE EDUCAÇÃO SUPERIOR DE BRASÍLIA – IESB", al=AL.CENTER, first=0, bold=True, ls=1.15)
    D.p("CURSO DE COMUNICAÇÃO SOCIAL – JORNALISMO", al=AL.CENTER, first=0, bold=True, after=120)
    D.p("FULANA EXEMPLAR DE TESTE", al=AL.CENTER, first=0, bold=True, after=120)
    D.p("A IMPRENSA DE PROVÍNCIA E O LEITOR: um estudo sobre a circulação de jornais", al=AL.CENTER, first=0, bold=True, after=140)  # F17
    D.p("Brasília", al=AL.CENTER, first=0, after=0)
    D.p("2026", al=AL.CENTER, first=0, after=0)
    D.br()
    # ---- folha de rosto (sem local e ano; título diferente da capa; sem folha de aprovação depois)
    D.p("FULANA EXEMPLAR DE TESTE", al=AL.CENTER, first=0, bold=True, after=100)  # F14, F15
    D.p("A IMPRENSA DE PROVÍNCIA E O LEITOR: estudo da circulação de jornais", al=AL.CENTER, first=0, bold=True, after=80)
    D.p("Trabalho de Conclusão de Curso apresentado ao curso de Comunicação Social – Jornalismo do Instituto de "
        "Educação Superior de Brasília – IESB, como requisito parcial para a obtenção do grau de Bacharel em Jornalismo.",
        al=AL.JUSTIFY, first=0, left=8, ls=1.0, size=11, after=30)
    D.p("Orientador: Prof. Beltrano de Teste", al=AL.LEFT, first=0, left=8, ls=1.0, size=11)
    D.br()  # F16: sem folha de aprovação
    # ---- resumo (2 parágrafos = F21; palavras-chave com ; e minúsculas = F13)
    D.p("RESUMO", al=AL.CENTER, first=0, bold=True, after=12)
    D.p("Este trabalho analisa a circulação dos jornais de província no Brasil do século XIX, com atenção às rotas de "
        "correio, ao custo do papel e ao perfil dos leitores. A pesquisa parte de um corpus fictício de periódicos e "
        "manchetes, reunido para fins de teste do pipeline de formatação.", ls=1.15)
    D.p("Os resultados indicam que a distribuição dependia de redes locais frágeis e que a leitura coletiva "
        "compensava o número reduzido de assinantes. Conclui-se que a imprensa de província ocupou papel relevante, "
        "embora restrito, na formação da esfera pública regional.", ls=1.15)
    D.p("Palavras-chave: imprensa; província; circulação; leitores.", first=0, after=12)
    D.br()
    D.p("ABSTRACT", al=AL.CENTER, first=0, bold=True, after=12)
    D.p("This paper analyzes the circulation of provincial newspapers in nineteenth-century Brazil, focusing on mail "
        "routes, paper costs and reader profiles. The corpus is fictional and was assembled to test a formatting "
        "pipeline. Results suggest that distribution relied on fragile local networks and that collective reading "
        "offset the small number of subscribers.", ls=1.15)
    D.p("Keywords: press; province; circulation; readers.", first=0, after=12)
    D.br()
    # ---- sumário digitado à mão (F07): pontilhado misto e páginas inventadas
    D.p("SUMÁRIO", al=AL.CENTER, first=0, bold=True, after=12)
    pg = 6
    for i, it in enumerate(x for x in SEC if x[0] in ("h1", "h2")):
        pg += rng.randint(0, 3)
        pt = "." * rng.randint(30, 55) if i % 4 else "…" * rng.randint(22, 40)
        D.p(f"{it[1]} {pt} {pg}", al=AL.LEFT, first=0, ls=1.15, after=2, left=0.6 if it[0] == "h2" else 0)
    D.p("REFERÊNCIAS " + "." * 40 + f" {pg + 2}", al=AL.LEFT, first=0, ls=1.15, after=2)

    # ---- texto
    def h(item):
        tipo, texto, op = item
        if tipo == "h1":
            par = d.add_paragraph(texto, style="Heading 1")
            par.paragraph_format.page_break_before = op.get("quebra", False)
            if op.get("quebra") is True and texto.startswith("1 "):
                par.paragraph_format.page_break_before = True
            if op.get("errado"):
                for r in par.runs:
                    r.bold = False
        else:
            par = d.add_paragraph(texto, style="Heading 2")
            if op.get("errado"):
                for r in par.runs:
                    r.bold = True
            else:
                for r in par.runs:
                    r.bold = False
        par.alignment = AL.LEFT

    def corpo(txt, **kw):
        kw.setdefault("ls", rng.choice([1.15, 1.15, 1.15, 1.5, 1.0]))  # F03: espaçamento irregular
        return D.p(txt, **kw)

    def t_par(id_):
        _, errada, _certa = TXT[id_]
        txt = paragrafo_texto(rng, rng.randint(4, 5), [errada])
        corpo(txt)

    chamada_c01 = {"C01a": ("A distribuição dependia do correio imperial (Gomes, 2016).", "(GOMES, 2016)"),
                   "C01b": ("O custo do papel pesava no orçamento dos editores (Nogueira, 2019).", "(NOGUEIRA, 2019)"),
                   "C01c": ("Os leitores formavam um grupo restrito (Andrade; Lima, 2012).", "(ANDRADE; LIMA, 2012)"),
                   "C01d": ("O corpus foi organizado por década (Mendes, 2008).", "(MENDES, 2008)"),
                   "C01e": ("A leitura coletiva ampliava o alcance das folhas (Pereira, 2018).", "(PEREIRA, 2018)")}
    outros = {
        "C04": "Os jornais de maior tiragem concentravam-se nas capitais (PEREIRA, 2019).",
        "C09": "A relação entre editores e correio é discutida em Silva e Souza (SILVA e SOUZA, 2020).",
        "C06a": "A censura prévia alterou a rotina das oficinas (SOUZA, 2019).",
        "C06b": "O serviço postal definia o ritmo das edições (SOUZA, 2019).",
        "C02": "A taxa de alfabetização limitava o público de forma decisiva (COSTA, 2021).",
        "C07": "Os impressos provinciais eram distribuídos em lotes pequenos (FERRAZ, 2013).",
        "D1": "A Guerra do Paraguai (1865) alterou a demanda por notícias nas províncias.",
        "D3": "O jornal era lido na praça. O jornal era discutido na praça. O jornal era repassado de mão em mão na praça.",
        "fato1832": "A primeira tipografia da vila foi instalada em 1832, segundo os registros da câmara municipal.",
        "fim": "Em síntese, a imprensa de província, apesar das limitações, deixou marcas duradouras na vida pública regional.",
    }
    for it in SEC:
        k = it[0]
        if k in ("h1", "h2"):
            h(it)
        elif k == "fill":
            for _ in range(max(1, round(it[1] * escala))):
                corpo(paragrafo_texto(rng))
        elif k == "fill_times":
            corpo(paragrafo_texto(rng), font="Times New Roman")
        elif k == "fill_espacos":
            D.p("    " + paragrafo_texto(rng), first=0, ls=1.15)
        elif k == "sp":
            c = it[1]
            if c in TXT:
                t_par(c)
            elif c in chamada_c01 or c in outros:
                frase_esp = (chamada_c01.get(c) or (outros[c],))[0] if c in chamada_c01 else outros[c]
                corpo(paragrafo_texto(rng, 4, [frase_esp]))
            elif c == "objetivos":
                corpo("O objetivo geral deste trabalho é analisar a circulação dos jornais de província no século XIX.")
                corpo("Os objetivos específicos são:", first=0)
                for n, o in enumerate(["identificar os jornais de maior circulação na região",
                                       "descrever as rotas de distribuição das folhas",
                                       "comparar o perfil dos leitores entre as décadas"], 1):
                    D.p(f"{n}. {o}", first=0, left=1.25, al=AL.LEFT, ls=1.15, after=2)  # F12
            elif c == "Q1":
                corpo(paragrafo_texto(rng, 2, ["Ao descrever as oficinas, Araújo observa o seguinte:"]))
                D.p(Q1 + ". (ARAÚJO, 2005, p. 41).", first=0, left=4, ls=1.0, size=10, after=12)  # correta: recuo 4 cm
            elif c == "Q2":
                corpo(paragrafo_texto(rng, 2))
                D.p('Para Marques, "' + Q2 + '" (MARQUES, 2011, p. 88).', first=0, left=2, ls=1.5)  # F10, C05
            elif c == "Q3":
                corpo(paragrafo_texto(rng, 3, [f'Para Mendes, o sistema "{Q3}" (MENDES, 2008, p. 19).']))
            elif c == "tabela":
                tb = d.add_table(rows=1, cols=4)
                tb.style = "Table Grid"
                dados = [("Década", "Jornais", "Tiragem média", "Assinantes"), ("1830", "4", "120", "310"),
                         ("1840", "7", "180", "520"), ("1850", "11", "240", "790"), ("1860", "15", "300", "1.040")]
                for i, linha in enumerate(dados):
                    cells = tb.rows[0].cells if i == 0 else tb.add_row().cells
                    for c_, v in zip(cells, linha):
                        c_.text = ""
                        rr = c_.paragraphs[0].add_run(v)
                        run_font(rr, "Calibri")  # F06
                        c_.paragraphs[0].paragraph_format.first_line_indent = None
                D.p("Tabela 1 – Jornais em circulação por década", al=AL.CENTER, first=0, before=6, size=11)  # F19: abaixo e sem Fonte
            elif c == "figura":
                img = saida / "_fig.png"
                im = Image.new("RGB", (1200, 700), "white")
                dr = ImageDraw.Draw(im)
                for i, v in enumerate([4, 7, 11, 15]):
                    dr.rectangle([150 + i * 250, 600 - v * 30, 330 + i * 250, 600], fill=(40, 90, 160))
                dr.line([100, 600, 1150, 600], fill="black", width=3)
                im.save(img)
                par = d.add_paragraph()
                par.alignment = AL.CENTER
                par.add_run().add_picture(str(img), width=Cm(12))
                D.p("Figura 1 – Periódicos em circulação por década", al=AL.CENTER, first=0, size=11, after=12)  # F18
                img.unlink()
            elif c == "manchetes":
                corpo(paragrafo_texto(rng, 2, ["As manchetes reunidas no corpus incluem as seguintes: "
                      + "; ".join(f'"{m}"' for m in MANCHETES) + "."]))
            else:
                raise SystemExit("sp desconhecido " + c)

    # ---- referências
    D.p("REFERÊNCIAS", al=AL.CENTER, first=0, bold=True, break_before=True, before=0, after=12)
    for k in ORDEM_REF:
        D.ref(*REFS[k])
    d.save(saida / "ORIGINAL.docx")
    return {"seed": seed, "escala": escala}


def gabarito():
    F = [
        ("F01", "Papel Letter (deve ser A4)"), ("F02", "Margens 2,54 cm (devem ser 3/3/2/2)"),
        ("F03", "Espaçamento entre linhas irregular no corpo (1,0/1,15/1,5; deve ser 1,5)"),
        ("F04", "Número de página no rodapé (deve ser no cabeçalho, canto superior direito)"),
        ("F05", "Número de página visível na capa e nos pré-textuais"),
        ("F06", "Mais de uma família tipográfica (Times New Roman em parágrafos, Calibri na tabela)"),
        ("F07", "Sumário digitado à mão (pontilhado misto, páginas erradas)"),
        ("F08", "Título 2.2 com destaque diferente dos demais subtítulos"),
        ("F09", "Seção 3 sem quebra de página (seções primárias devem começar em nova página)"),
        ("F10", "Citação longa de Marques sem recuo de 4 cm, com aspas, corpo 12 e espaçamento 1,5"),
        ("F11", "Referências justificadas, espaçamento 1,5 e sem linha em branco entre elas"),
        ("F12", "Objetivos específicos numerados 1., 2., 3. com inicial minúscula e sem ponto final (decisão do curso: a), b), c), inicial maiúscula, ponto final)"),
        ("F13", "Palavras-chave separadas por ponto e vírgula e em minúsculas (decisão do curso: ponto, inicial maiúscula)"),
        ("F14", "Folha de rosto sem local e ano"), ("F15", "Título da capa diferente do da folha de rosto"),
        ("F16", "Folha de aprovação ausente"), ("F17", "Capa com título e subtítulo na mesma linha"),
        ("F18", "Figura 1 com legenda abaixo e sem Fonte"), ("F19", "Tabela 1 com legenda abaixo e sem Fonte"),
        ("F20", "Parágrafos com espaços digitados no início em vez de recuo (3 ocorrências)"),
        ("F21", "Resumo em dois parágrafos (deve ser parágrafo único)"),
        ("F23", "Título da seção 5 em formato diferente dos demais (5. Considerações Finais)"),
        ("F24", "Destaque tipográfico inconsistente nas referências (PEREIRA sem negrito)"),
    ]
    for i, d_ in F:
        gab(i, "formatacao", d_, corrigir=True)
    C = [
        ("C01a", "Chamada em caixa mista (Gomes, 2016); deve ser (GOMES, 2016)"),
        ("C01b", "Chamada em caixa mista (Nogueira, 2019)"), ("C01c", "Chamada em caixa mista (Andrade; Lima, 2012)"),
        ("C01d", "Chamada em caixa mista (Mendes, 2008)"), ("C01e", "Chamada em caixa mista (Pereira, 2018)"),
        ("C02", "Citação (COSTA, 2021) sem referência correspondente (pendência ao cliente)"),
        ("C03", "Referência TAVARES (2015) não citada no texto (pendência ao cliente)"),
        ("C04", "Chamada (PEREIRA, 2019) × referência PEREIRA 2018 (divergência de ano)"),
        ("C05", "Citação de Marques: sem ponto final antes da chamada (padrão 2002 do curso: ponto do texto citado + chamada + ponto)"),
        ("C06", "Duas obras de SOUZA (2019) sem letras a/b na chamada e na referência"),
        ("C07", "Referência FERRAZ sem ano (texto cita 2013): pendência, não completar de memória"),
        ("C08", "Referências fora de ordem alfabética (MENDES antes de MARQUES)"),
        ("C09", "Chamada (SILVA e SOUZA, 2020); deve ser (SILVA; SOUZA, 2020)"),
    ]
    for i, d_ in C:
        gab(i, "citacao_referencia", d_, corrigir=True)
    for i, errada, certa in ERROS_TEXTO:
        gab(i, "texto", "Erro de revisão textual", errada=errada, certa=certa, corrigir=True)
    gab("P1", "protegido", "Citação longa de Araújo contém 'atrazo' (grafia errada): NÃO alterar", texto=Q1, corrigir=False)
    gab("P2", "protegido", "Citação de Marques diz 1831; o texto diz 1832: NÃO alterar; só comentário [FATO]/[DIVERGÊNCIA]", texto=Q2, corrigir=False)
    gab("P3", "protegido", "Citação curta de Mendes contém 'en' no lugar de 'em': NÃO alterar", texto=Q3, corrigir=False)
    gab("P4", "protegido", "Quatro manchetes de época entre aspas (grafia antiga): NÃO alterar", texto=MANCHETES, corrigir=False)
    gab("D1", "chamariz", "'Guerra do Paraguai (1865)' parece chamada, mas não é", corrigir=False)
    gab("D2", "chamariz", "(CARVALHO; FREITAS; LOPES; REIS, 2014) com todos os autores é aceitável: não exige et al.", corrigir=False)
    gab("D3", "chamariz", "Repetição de 'O jornal era...' é escolha de estilo correta: não reescrever", corrigir=False)
    gab("X1", "risco", "Extensão total abaixo do mínimo do curso (40 a 60 páginas): risco alto, vai no relatório; formatação não resolve", corrigir=False)
    return GAB


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--saida", required=True)
    ap.add_argument("--seed", type=int, default=5)
    ap.add_argument("--escala", type=float, default=2.0)
    a = ap.parse_args()
    out = Path(a.saida)
    out.mkdir(parents=True, exist_ok=True)
    meta = construir(out, a.seed, a.escala)
    priv = out / ".privado"
    priv.mkdir(exist_ok=True)
    g = gabarito()
    (priv / "gabarito.json").write_text(json.dumps({"meta": meta, "itens": g}, ensure_ascii=False, indent=2), encoding="utf-8")
    por = {}
    for i in g:
        if i["corrigir"]:
            por[i["categoria"]] = por.get(i["categoria"], 0) + 1
    print("OK", out / "ORIGINAL.docx", "| erros a corrigir:", sum(por.values()), por,
          "| armadilhas:", sum(1 for i in g if i["categoria"] == "protegido"),
          "| chamarizes:", sum(1 for i in g if i["categoria"] == "chamariz"))


if __name__ == "__main__":
    main()
