"""Gera documentos FICTÍCIOS para o autoteste do pipeline (nenhum dado de cliente).

Estrutura igual à de um TCC: capa, folha de rosto, resumo e sumário sem número; Introdução
em diante com número no canto superior direito; citação longa com recuo de 4 cm; citação
curta; tabela de dados; referências.

Gera em <pasta>:
  original.docx        sem marcações
  marcada_ok.docx      correções marcadas por "Revisão ABNT", comentários com prefixo,
                       uma marca de parágrafo excluída; nenhum trecho protegido alterado
  marcada_ruim.docx    autor fora do padrão, comentário sem prefixo e data trocada DENTRO
                       da citação longa (reproduz o E1 e o E14)
  sem_numero.docx      igual ao original, mas sem número de página na seção do texto (E2)
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
DATA = "2026-10-21T00:00:00Z"


def run(t: str, extra: str = "") -> str:
    return f'<w:r>{extra}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'


def ins(t: str, autor: str, i: int) -> str:
    return f'<w:ins w:id="{i}" w:author="{autor}" w:date="{DATA}">{run(t)}</w:ins>'


def dele(t: str, autor: str, i: int) -> str:
    return (f'<w:del w:id="{i}" w:author="{autor}" w:date="{DATA}"><w:r>'
            f'<w:delText xml:space="preserve">{escape(t)}</w:delText></w:r></w:del>')


def p(conteudo: str, estilo: str | None = None, ppr_extra: str = "", centro: bool = False) -> str:
    ppr = ""
    if estilo or ppr_extra or centro:
        ppr = ("<w:pPr>" + (f'<w:pStyle w:val="{estilo}"/>' if estilo else "") + ppr_extra
               + ('<w:jc w:val="center"/>' if centro else "") + "</w:pPr>")
    return f"<w:p>{ppr}{conteudo}</w:p>"


def quebra() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def sect(header_rid: str, extra: str = "", final: bool = False) -> str:
    s = (f'<w:sectPr><w:headerReference w:type="default" r:id="{header_rid}"/>'
         f'<w:pgSz w:w="11906" w:h="16838"/>'
         f'<w:pgMar w:top="1701" w:right="1134" w:bottom="1134" w:left="1701" w:header="1134" w:footer="709" w:gutter="0"/>'
         f"{extra}</w:sectPr>")
    return s if final else f"<w:p><w:pPr>{s}</w:pPr></w:p>"


CITACAO_LONGA = ("A imprensa brasileira nasceu sob a tutela da Coroa e só ganhou autonomia depois da "
                 "Independência, quando os periódicos passaram a disputar a opinião pública nas províncias e "
                 "na Corte, num processo que se estendeu por várias décadas do século XIX")
CITACAO_LONGA_RUIM = CITACAO_LONGA.replace("século XIX", "século XVIII")


def documento(variante: str) -> str:
    A = {"original": None, "ok": "Revisão ABNT", "ruim": "GG – Revisão", "sem_numero": None}[variante]
    marc = A is not None
    n = iter(range(100, 200))

    capa = [p(run("UNIVERSIDADE FICTÍCIA DO PLANALTO"), centro=True),
            p(run("Fulana Exemplar de Teste"), centro=True),
            p(run("O JORNALISMO DE PROVÍNCIA:"), centro=True),
            p(run("um estudo fictício"), centro=True),
            sect("rIdH1")]
    pre = [p(run("Fulana Exemplar de Teste"), centro=True),
           p(run("Trabalho de Conclusão de Curso apresentado como requisito fictício."), centro=True),
           quebra(),
           p(run("RESUMO"), centro=True),
           p(run("Este resumo é fictício e serve só para testar o pipeline.")),
           quebra(),
           p(run("SUMÁRIO"), centro=True),
           p(run("1 INTRODUÇÃO ........................................................ 4")),
           p(run("2 DESENVOLVIMENTO .............................................. 5")),
           p(run("REFERÊNCIAS ......................................................... 6")),
           sect("rIdH1", '<w:pgNumType w:start="1"/>')]

    # parágrafo com correção marcada: concordância ("relata" -> "relatam")
    if marc:
        corr = run("Silva e Souza ") + dele("relata", A, next(n)) + ins("relatam", A, next(n)) + run(
            " que os jornais de província tinham circulação restrita (SILVA, 2020).")
        corr_com = ('<w:commentRangeStart w:id="1"/>' + corr + '<w:commentRangeEnd w:id="1"/>'
                    '<w:r><w:commentReference w:id="1"/></w:r>')
    else:
        corr_com = run("Silva e Souza relata que os jornais de província tinham circulação restrita (SILVA, 2020).")

    # marca de parágrafo excluída (dois parágrafos que deveriam ser um)
    if marc:
        uniao = (f'<w:p><w:pPr><w:rPr><w:del w:id="{next(n)}" w:author="{A}" w:date="{DATA}"/></w:rPr></w:pPr>'
                 + run("Para Souza (2019), a circulação dependia") + "</w:p>"
                 + p(run(" do correio imperial.")))
    else:
        uniao = p(run("Para Souza (2019), a circulação dependia")) + p(run(" do correio imperial."))

    longa_txt = CITACAO_LONGA_RUIM if variante == "ruim" else CITACAO_LONGA
    if variante == "ruim":
        longa_run = run(CITACAO_LONGA.split("século")[0] + "século ") + dele("XIX", A, next(n)) + ins(
            "XVIII", A, next(n)) + run(". (BARBOSA, 2010, p. 45).")
    else:
        longa_run = run(longa_txt + ". (BARBOSA, 2010, p. 45).")
    longa = p(longa_run, ppr_extra='<w:spacing w:line="240" w:lineRule="auto"/><w:ind w:left="2268"/>')

    if variante == "ruim":
        com2 = ('<w:commentRangeStart w:id="2"/>' + run("Conforme o autor, ")
                + '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>')
    elif marc:
        com2 = ('<w:commentRangeStart w:id="2"/>' + run("Conforme o autor, ")
                + '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>')
    else:
        com2 = run("Conforme o autor, ")
    curta = p(com2 + run('o jornal era "o espelho da vida provincial" (BARBOSA, 2010, p. 12).'))

    tabela = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/>'
              '<w:gridCol w:w="4000"/></w:tblGrid>'
              + "".join(f"<w:tr><w:tc><w:p>{run(a)}</w:p></w:tc><w:tc><w:p>{run(b)}</w:p></w:tc></w:tr>"
                        for a, b in [("Jornal", "Edições"), ("Gazeta A", "22"), ("Correio B", "17")])
              + "</w:tbl>")

    texto = [p(run("1 INTRODUÇÃO")),
             p(corr_com),
             uniao,
             p(run("A pesquisa usa dados de Pereira (2018) e de autor que não está nas referências (COSTA, 2021).")),
             quebra(),
             p(run("2 DESENVOLVIMENTO")),
             curta,
             longa,
             p(run("Tabela 1 – Edições analisadas")),
             tabela,
             p(run("Fonte: dados fictícios da pesquisa.")),
             quebra(),
             p(run("REFERÊNCIAS")),
             p(run("BARBOSA, Marialva. História cultural da imprensa: Brasil 1800-1900. Rio de Janeiro: Mauad X, 2010.")),
             p(run("PEREIRA, João. Circulação de jornais no Império. São Paulo: Editora Fictícia, 2018.")),
             p(run("SILVA, Ana; SOUZA, Beto. Jornais de província. Brasília: Editora Fictícia, 2020.")),
             p(run("SOUZA, Beto. Correio e imprensa. Recife: Editora Fictícia, 2019.")),
             p(run("TAVARES, Carla. Referência que ninguém citou. Natal: Editora Fictícia, 2015."))]
    header_texto = "rIdH1" if variante == "sem_numero" else "rIdH2"
    corpo = "".join(capa + pre + texto) + sect(header_texto, final=True)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W_NS}><w:body>{corpo}</w:body></w:document>'


def comentarios(variante: str) -> str | None:
    if variante in ("original", "sem_numero"):
        return None
    A = "Revisão ABNT" if variante == "ok" else "GG – Revisão"
    t1 = "[REGRA] Concordância: sujeito composto pede verbo no plural."
    t2 = "[PENDENTE] Confirmar a página da citação no livro." if variante == "ok" else "Confirmar a página da citação."
    def c(i, t):
        return (f'<w:comment w:id="{i}" w:author="{A}" w:date="{DATA}" w:initials="RA">'
                f'<w:p>{run(t)}</w:p></w:comment>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {W_NS}>{c(1, t1)}{c(2, t2)}</w:comments>'


ESTILOS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W_NS}>'
           '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
           '<w:sz w:val="24"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault></w:docDefaults>'
           '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
           '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr></w:style>'
           '<w:style w:type="paragraph" w:styleId="Cabealho"><w:name w:val="header"/><w:basedOn w:val="Normal"/>'
           '<w:pPr><w:ind w:firstLine="0"/><w:jc w:val="right"/></w:pPr></w:style>'
           '</w:styles>')

HEADER_VAZIO = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W_NS}>'
                '<w:p><w:pPr><w:pStyle w:val="Cabealho"/></w:pPr></w:p></w:hdr>')
HEADER_PAGE = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W_NS}>'
               '<w:p><w:pPr><w:pStyle w:val="Cabealho"/></w:pPr>'
               '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
               '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
               '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
               '<w:r><w:t>4</w:t></w:r>'
               '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>')


def pacote(variante: str, destino: Path) -> None:
    com = comentarios(variante)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          + ('<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' if com else "")
          + "</Types>")
    rels_raiz = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 '</Relationships>')
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    rels_doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rIdS" Type="{R}styles" Target="styles.xml"/>'
                f'<Relationship Id="rIdT" Type="{R}settings" Target="settings.xml"/>'
                f'<Relationship Id="rIdH1" Type="{R}header" Target="header1.xml"/>'
                f'<Relationship Id="rIdH2" Type="{R}header" Target="header2.xml"/>'
                + (f'<Relationship Id="rIdC" Type="{R}comments" Target="comments.xml"/>' if com else "")
                + "</Relationships>")
    track = "<w:trackRevisions/>" if variante in ("ok", "ruim") else ""
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W_NS}>{track}'
                '<w:updateFields w:val="true"/></w:settings>')
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels_raiz)
        z.writestr("word/_rels/document.xml.rels", rels_doc)
        z.writestr("word/document.xml", documento(variante))
        z.writestr("word/styles.xml", ESTILOS)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/header1.xml", HEADER_VAZIO)
        z.writestr("word/header2.xml", HEADER_PAGE)
        if com:
            z.writestr("word/comments.xml", com)


def main(pasta: str) -> None:
    d = Path(pasta)
    d.mkdir(parents=True, exist_ok=True)
    nomes = {"original": "original.docx", "ok": "marcada_ok.docx", "ruim": "marcada_ruim.docx",
             "sem_numero": "sem_numero.docx"}
    for v, nome in nomes.items():
        pacote(v, d / nome)
    print("Fixtures em", d)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "fixtures")
