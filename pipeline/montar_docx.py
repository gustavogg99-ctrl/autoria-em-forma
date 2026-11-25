"""Etapa 5 do fluxo (06, seção 4): monta a versão MARCADA do zero, a partir do original.

Motor genérico. O que muda de um trabalho para outro fica num arquivo de especificação
(spec) guardado em entrada/, junto do trabalho do cliente (nunca no repositório):
a ordem dos blocos, o papel de cada parágrafo do original, as correções e os comentários.

O motor garante o "como" (06, 3.6, 3.7 e 3.11):
- estrutura reconstruída do zero: A4, margens, estilos, três seções (capa / pré-textuais /
  texto), cabeçalho explícito em todas, campo PAGE montado do zero, sumário automático;
- toda alteração de texto em controle de alterações, com autor neutro ("Revisão ABNT"),
  reduzida ao trecho que de fato muda (prefixo e sufixo comuns ficam de fora);
- parágrafo novo (ex.: folha de aprovação) entra inteiro como inserção;
- comentários com prefixo obrigatório, ancorados no trecho alterado, num trecho dado ou no
  parágrafo inteiro;
- conferência de cobertura: todo parágrafo com texto do original precisa ser usado uma vez
  ou declarado como descartado (ex.: sumário digitado), senão o build falha.

Uso (a spec importa este módulo):
  python3 -I entrada/spec_<cliente>.py
"""
from __future__ import annotations

import copy
import datetime as _dt
import re
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

from docx_seguro import NS, abrir, corpo, salvar, texto_paragrafo, w

PREFIXOS = ("[REGRA]", "[PENDENTE]", "[DECIDIR]", "[DIVERGÊNCIA]", "[FATO]")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSDECL = (f'xmlns:w="{W_NS}" xmlns:r="{R_NS}" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')

# A4 e margens da NBR 14724 (twips)
PAG_L, PAG_A = 11906, 16838
M_SUP, M_ESQ, M_INF, M_DIR = 1701, 1701, 1134, 1134
LARGURA_TEXTO = PAG_L - M_ESQ - M_DIR  # 9071

# ordem dos filhos de w:pPr exigida pelo schema (06, erros técnicos)
ORDEM_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "numPr",
             "pBdr", "shd", "tabs", "spacing", "ind", "jc", "outlineLvl"]


class ErroSpec(Exception):
    pass


def _t(texto: str) -> str:
    return escape(texto, {'"': "&quot;"})


@dataclass
class Comentario:
    id: int
    texto: str


@dataclass
class Montador:
    original: str
    autor: str = "Revisão ABNT"
    fonte: str = "Times New Roman"
    data: str = field(default_factory=lambda: _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:00Z"))
    criador: str = ""

    def __post_init__(self):
        self.partes = abrir(self.original)
        self.body = list(corpo(self.partes))
        self.usados: set[int] = set()
        self.descartados: set[int] = set()
        self.xml: list[str] = []
        self.comentarios: list[Comentario] = []
        self._rev = 100
        self._docpr = 10
        self.titulos: list[tuple[int, str, str]] = []  # (nível, texto final, bookmark)
        self.paginas_sumario: dict[str, str] = {}
        self.imagens: dict[str, bytes] = {}
        self._juntar_origem: int | None = None

    # ------------------------------------------------------------------ utilidades
    def _id(self) -> int:
        self._rev += 1
        return self._rev

    def _attrs(self) -> str:
        return f'w:id="{self._id()}" w:author="{_t(self.autor)}" w:date="{self.data}"'

    def comentario(self, texto: str) -> int:
        texto = texto.strip()
        if not texto.startswith(PREFIXOS):
            raise ErroSpec(f"Comentário sem prefixo de 3.11: {texto[:60]!r}")
        cid = len(self.comentarios)
        self.comentarios.append(Comentario(cid, texto))
        return cid

    def texto_original(self, i: int) -> str:
        return texto_paragrafo(self.body[i], "original")

    def runs_original(self, i: int, formatacao: bool = True) -> list[tuple[str, frozenset]]:
        """Caracteres do parágrafo i com as marcas de negrito/itálico de cada run."""
        el = self.body[i]
        if etree.QName(el).localname != "p":
            raise ErroSpec(f"bloco {i} não é parágrafo")
        chars: list[tuple[str, frozenset]] = []
        for r in el.iter(w("r")):
            props = set()
            rpr = r.find("w:rPr", NS)
            if rpr is not None and formatacao:
                for tag, marca in (("b", "b"), ("i", "i"), ("u", "u")):
                    e = rpr.find(f"w:{tag}", NS)
                    if e is not None and e.get(w("val")) not in ("0", "false", "none"):
                        props.add(marca)
            for t in r:
                nome = etree.QName(t).localname
                if nome == "t":
                    chars += [(c, frozenset(props)) for c in (t.text or "")]
                elif nome == "tab":
                    chars.append(("\t", frozenset(props)))
        # espaços sobrando nas pontas não são texto do autor: saem sem marcação
        while chars and chars[0][0].isspace():
            chars.pop(0)
        while chars and chars[-1][0].isspace():
            chars.pop()
        return chars

    def descartar(self, *indices: int) -> None:
        """Parágrafos do original que não entram (ex.: sumário digitado, substituído pelo automático)."""
        for i in indices:
            self.descartados.add(i)

    # ------------------------------------------------------------------ parágrafo
    def par(self, origem: int | None = None, texto: str | list | None = None, estilo: str = "Normal",
            edicoes: list[dict] = (), ancoras: list[dict] = (), comentario: str | None = None,
            novo: bool = False, ppr: dict | None = None, rpr: str = "", formatacao: bool = True,
            props_novas: frozenset = frozenset(), sect: str | None = None, nivel: int | None = None,
            idioma: str | None = None, juntar_com_proximo: bool = False, excluir: bool = False) -> str:
        """Acrescenta um parágrafo.

        origem   índice do parágrafo no corpo do original (texto e negrito/itálico vêm de lá);
        texto    texto novo (str ou lista de (texto, props)); com novo=True entra como inserção;
        edicoes  [{"de": ..., "para": ..., "n": 1, "comentario": "[REGRA] ..."}] sobre o original;
        ancoras  [{"ancora": ..., "n": 1, "comentario": "[PENDENTE] ..."}] sem alterar o texto;
        comentario  comentário ancorado no parágrafo inteiro;
        nivel    1 ou 2 para títulos que entram no sumário;
        juntar_com_proximo  junta este parágrafo ao seguinte, como alteração rastreada: a marca de parágrafo é
                 excluída (aceitar_alteracoes.py une os dois na limpa) e, se faltar, entra um espaço inserido
                 entre os dois textos. O parágrafo seguinte tem de ser o próximo do original (origem + 1) e
                 vem logo depois nesta spec. A união usa o estilo e a formatação do segundo parágrafo.
        excluir  exclui o parágrafo inteiro do original como alteração rastreada (texto e marca de parágrafo
                 excluídos; o aceitar_alteracoes.py o remove na limpa). Pede origem; combina com comentario
                 (ex.: "[REGRA] Referência removida a pedido do autor"), não com edicoes nem juntar_com_proximo.
                 Não usar no último parágrafo do documento: sem parágrafo seguinte, o aceitar só tira a marca
                 e sobra um parágrafo vazio.
        """
        if self._juntar_origem is not None:
            if origem is None or origem != self._juntar_origem + 1:
                raise ErroSpec(f"parágrafo {self._juntar_origem} pede junção com o próximo "
                               f"({self._juntar_origem + 1}), mas vem {origem!r}")
            self._juntar_origem = None
        if juntar_com_proximo and origem is None:
            raise ErroSpec("juntar_com_proximo só vale para parágrafo do original (origem)")
        if excluir and (origem is None or edicoes or juntar_com_proximo or novo):
            raise ErroSpec("excluir só vale para parágrafo do original (origem), sem edicoes, novo nem junção")
        if origem is not None:
            if origem in self.usados:
                raise ErroSpec(f"parágrafo {origem} usado duas vezes")
            self.usados.add(origem)
            base = self.runs_original(origem, formatacao)
            if not base:
                raise ErroSpec(f"parágrafo {origem} está vazio no original")
        else:
            if isinstance(texto, str):
                base = [(c, frozenset(props_novas)) for c in texto]
            else:
                base = [(c, frozenset(p)) for parte, p in (texto or []) for c in parte]
        txt = "".join(c for c, _ in base)

        # ---- operações de edição (posições no texto original do parágrafo)
        ops = []
        for e in edicoes:
            de, para, n = e["de"], e["para"], e.get("n", 1)
            pos = -1
            for _ in range(n):
                pos = txt.find(de, pos + 1)
                if pos < 0:
                    raise ErroSpec(f"trecho não encontrado (ocorrência {n}): {de!r} em {txt[:60]!r}")
            p = 0
            while p < min(len(de), len(para)) and de[p] == para[p]:
                p += 1
            q = 0
            while q < min(len(de), len(para)) - p and de[len(de) - 1 - q] == para[len(para) - 1 - q]:
                q += 1
            ini, fim, novo_txt = pos + p, pos + len(de) - q, para[p:len(para) - q]
            if ini == fim and not novo_txt:
                raise ErroSpec(f"edição sem efeito: {de!r}")
            ops.append((ini, fim, novo_txt, e.get("comentario")))
        ops.sort(key=lambda o: o[0])
        for a, b in zip(ops, ops[1:]):
            if a[1] > b[0]:
                raise ErroSpec(f"edições sobrepostas em {txt[:50]!r}")

        # ---- sequência final: (tipo, char, props)
        tipo_base = "ins" if novo else "keep"
        seq: list[tuple[str, str, frozenset]] = []
        pos_seq: dict[int, int] = {}
        intervalos: list[tuple[int, int, int]] = []  # (ini, fim, id do comentário) em posições de seq
        i = 0
        for ini, fim, novo_txt, com in ops:
            for k in range(i, ini):
                pos_seq[k] = len(seq)
                seq.append((tipo_base, *base[k]))
            a = len(seq)
            for k in range(ini, fim):
                pos_seq[k] = len(seq)
                seq.append(("del", *base[k]))
            herda = base[ini - 1][1] if ini > 0 else (base[ini][1] if base else frozenset())
            seq += [("ins", c, herda) for c in novo_txt]
            if com:
                intervalos.append((a, len(seq), self.comentario(com)))
            i = fim
        for k in range(i, len(base)):
            pos_seq[k] = len(seq)
            seq.append((tipo_base, *base[k]))
        pos_seq[len(base)] = len(seq)
        if excluir:
            seq = [("del", c, pr) for _, c, pr in seq]
        if juntar_com_proximo:
            ultimo = next((c for t, c, _ in reversed(seq) if t != "del"), "")
            if ultimo and not ultimo.isspace():
                seq.append(("ins", " ", seq[-1][2] if seq else frozenset()))
            self._juntar_origem = origem

        for an in ancoras:
            alvo, n = an["ancora"], an.get("n", 1)
            pos = -1
            for _ in range(n):
                pos = txt.find(alvo, pos + 1)
                if pos < 0:
                    raise ErroSpec(f"âncora não encontrada: {alvo!r} em {txt[:60]!r}")
            intervalos.append((pos_seq[pos], pos_seq[pos + len(alvo) - 1] + 1, self.comentario(an["comentario"])))
        if comentario:
            intervalos.append((0, len(seq), self.comentario(comentario)))

        # ---- runs
        inicios: dict[int, list[int]] = {}
        fins: dict[int, list[int]] = {}
        for a, b, cid in intervalos:
            inicios.setdefault(a, []).append(cid)
            fins.setdefault(b, []).append(cid)
        cortes = sorted(set(inicios) | set(fins) | {0, len(seq)})
        rpr_lang = f'<w:lang w:val="{idioma}"/>' if idioma else ""
        partes_xml: list[str] = []

        def emitir(trecho):
            if not trecho:
                return
            grupos = []
            for tipo, c, props in trecho:
                if grupos and grupos[-1][0] == tipo and grupos[-1][2] == props:
                    grupos[-1][1].append(c)
                else:
                    grupos.append([tipo, [c], props])
            for tipo, cs, props in grupos:
                s = "".join(cs)
                marcas = ("<w:b/><w:bCs/>" if "b" in props else "") + ("<w:i/><w:iCs/>" if "i" in props else "") \
                    + ('<w:u w:val="single"/>' if "u" in props else "")
                rp = f"<w:rPr>{marcas}{rpr}{rpr_lang}</w:rPr>" if (marcas or rpr or rpr_lang) else ""
                pedacos = re.split(r"(\t)", s)
                if tipo == "del":
                    corpo_r = "".join("<w:tab/>" if x == "\t" else f'<w:delText xml:space="preserve">{_t(x)}</w:delText>'
                                      for x in pedacos if x)
                    partes_xml.append(f"<w:del {self._attrs()}><w:r>{rp}{corpo_r}</w:r></w:del>")
                else:
                    corpo_r = "".join("<w:tab/>" if x == "\t" else f'<w:t xml:space="preserve">{_t(x)}</w:t>'
                                      for x in pedacos if x)
                    r = f"<w:r>{rp}{corpo_r}</w:r>"
                    partes_xml.append(f"<w:ins {self._attrs()}>{r}</w:ins>" if tipo == "ins" else r)

        for a, b in zip(cortes, cortes[1:] + [None]):
            for cid in fins.get(a, []):
                partes_xml.append(f'<w:commentRangeEnd w:id="{cid}"/>'
                                  f'<w:r><w:rPr><w:rStyle w:val="Refdecomentario"/></w:rPr>'
                                  f'<w:commentReference w:id="{cid}"/></w:r>')
            for cid in inicios.get(a, []):
                partes_xml.append(f'<w:commentRangeStart w:id="{cid}"/>')
            if b is not None:
                emitir(seq[a:b])

        # ---- pPr
        props_p = dict(ppr or {})
        bookmark = ""
        if nivel:
            final = "".join(c for t, c, _ in seq if t != "del")
            nome_bm = f"_Toc{len(self.titulos) + 1:04d}"
            self.titulos.append((nivel, final, nome_bm))
            bid = self._id()
            bookmark = (f'<w:bookmarkStart w:id="{bid}" w:name="{nome_bm}"/>', f'<w:bookmarkEnd w:id="{bid}"/>')
        filhos = [f'<w:pStyle w:val="{estilo}"/>'] if estilo else []
        for chave in ORDEM_PPR[1:]:
            if chave in props_p and props_p[chave] is not None:
                filhos.append(props_p[chave])
        marca_par = f"<w:rPr><w:ins {self._attrs()}/></w:rPr>" if novo else ""
        if juntar_com_proximo or excluir:
            marca_par = f"<w:rPr><w:del {self._attrs()}/></w:rPr>"
        sect_xml = sect or ""
        ppr_xml = f"<w:pPr>{''.join(filhos)}{marca_par}{sect_xml}</w:pPr>"
        conteudo = "".join(partes_xml)
        if bookmark:
            conteudo = bookmark[0] + conteudo + bookmark[1]
        self.xml.append(f"<w:p>{ppr_xml}{conteudo}</w:p>")
        return "".join(c for t, c, _ in seq if t != "del")

    # ------------------------------------------------------------------ outros blocos
    def imagem_do_original(self, origem: int, estilo: str = "Centralizado", ppr: dict | None = None) -> None:
        """Copia a imagem (ex.: logomarca da capa) de um parágrafo do original."""
        el = self.body[origem]
        dr = el.find(".//w:drawing", NS)
        if dr is None:
            raise ErroSpec(f"parágrafo {origem} sem imagem")
        self.usados.add(origem)
        blip = dr.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
        rid = blip.get(f"{{{R_NS}}}embed")
        rels = etree.fromstring(self.partes["word/_rels/document.xml.rels"])
        alvo = next(r.get("Target") for r in rels if r.get("Id") == rid)
        nome = Path(alvo).name
        self.imagens[nome] = self.partes["word/" + alvo]
        novo = copy.deepcopy(dr)
        novo.find(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}).set(f"{{{R_NS}}}embed", f"rIdImg_{nome}")
        for docpr in novo.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
            self._docpr += 1
            docpr.set("id", str(self._docpr))
        s = etree.tostring(novo, encoding="unicode")
        s = re.sub(r'\sxmlns(:\w+)?="[^"]+"', "", s)
        filhos = [f'<w:pStyle w:val="{estilo}"/>'] + [v for k, v in (ppr or {}).items() if v]
        self.xml.append(f"<w:p><w:pPr>{''.join(filhos)}</w:pPr><w:r>{s}</w:r></w:p>")

    def tabela_do_original(self, origem: int, larguras: list[int] | None = None,
                           repetir_cabecalho: bool = True, comentario: str | None = None) -> None:
        """Reconstrói a tabela/quadro do original: mesmo texto, bordas fechadas, fonte 10, espaço simples."""
        tbl = self.body[origem]
        if etree.QName(tbl).localname != "tbl":
            raise ErroSpec(f"bloco {origem} não é tabela")
        self.usados.add(origem)
        linhas = tbl.findall("w:tr", NS)
        ncol = max(len(tr.findall("w:tc", NS)) for tr in linhas)
        larguras = larguras or [LARGURA_TEXTO // ncol] * ncol
        larguras[-1] += LARGURA_TEXTO - sum(larguras)
        borda = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        bordas = "".join(borda.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
        out = [f'<w:tbl><w:tblPr><w:tblW w:w="{LARGURA_TEXTO}" w:type="dxa"/><w:jc w:val="center"/>'
               f'<w:tblBorders>{bordas}</w:tblBorders><w:tblLayout w:type="fixed"/>'
               '<w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="85" w:type="dxa"/>'
               '<w:bottom w:w="28" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>'
               '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
               'w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>']
        out += [f'<w:gridCol w:w="{x}"/>' for x in larguras]
        out.append("</w:tblGrid>")
        primeira_com = comentario
        for li, tr in enumerate(linhas):
            trpr = "<w:cantSplit/>" + ("<w:tblHeader/>" if li == 0 and repetir_cabecalho else "")
            out.append(f"<w:tr><w:trPr>{trpr}</w:trPr>")
            for ci, tc in enumerate(tr.findall("w:tc", NS)):
                out.append(f'<w:tc><w:tcPr><w:tcW w:w="{larguras[ci]}" w:type="dxa"/></w:tcPr>')
                pars = tc.findall("w:p", NS)
                for p in pars:
                    chars = []
                    for r in p.iter(w("r")):
                        rpr = r.find("w:rPr", NS)
                        b = rpr is not None and rpr.find("w:b", NS) is not None and \
                            rpr.find("w:b", NS).get(w("val")) not in ("0", "false")
                        for t in r.iter(w("t")):
                            chars.append((t.text or "", b))
                    texto = "".join(c for c, _ in chars).strip()
                    negrito = any(b for _, b in chars)
                    rp = "<w:rPr><w:b/><w:bCs/></w:rPr>" if negrito else ""
                    marca = ""
                    if primeira_com:
                        cid = self.comentario(primeira_com)
                        primeira_com = None
                        marca = (f'<w:commentRangeStart w:id="{cid}"/>',
                                 f'<w:commentRangeEnd w:id="{cid}"/><w:r><w:rPr><w:rStyle w:val="Refdecomentario"/>'
                                 f'</w:rPr><w:commentReference w:id="{cid}"/></w:r>')
                    run = f'<w:r>{rp}<w:t xml:space="preserve">{_t(texto)}</w:t></w:r>' if texto else ""
                    if marca:
                        run = marca[0] + run + marca[1]
                    out.append(f'<w:p><w:pPr><w:pStyle w:val="TextoTabela"/></w:pPr>{run}</w:p>')
                if not pars:
                    out.append('<w:p><w:pPr><w:pStyle w:val="TextoTabela"/></w:pPr></w:p>')
                out.append("</w:tc>")
            out.append("</w:tr>")
        out.append("</w:tbl>")
        self.xml.append("".join(out))

    def sumario(self, titulo_estilo: str = "TituloPreTextual", comentario: str | None = None,
                sect: str | None = None) -> None:
        """Marca a posição do sumário automático; as entradas são preenchidas em salvar()."""
        self.par(texto="SUMÁRIO", estilo=titulo_estilo, comentario=comentario)
        self.xml.append("@@SUMARIO@@")
        if sect:
            self.xml.append(f'<w:p><w:pPr><w:pStyle w:val="Normal"/>{sect}</w:pPr></w:p>')

    def _xml_sumario(self) -> str:
        if not self.titulos:
            return ""
        out = []
        for k, (nivel, texto, bm) in enumerate(self.titulos):
            pag = self.paginas_sumario.get(bm, "0")
            ini = ""
            if k == 0:
                ini = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                       '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
                       '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
            fim = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if k == len(self.titulos) - 1 else ""
            corpo_e = (f'<w:hyperlink w:anchor="{bm}" w:history="1">'
                       f'<w:r><w:t xml:space="preserve">{_t(texto)}</w:t></w:r><w:r><w:tab/></w:r>'
                       f'<w:r><w:t>{pag}</w:t></w:r></w:hyperlink>')
            out.append(f'<w:p><w:pPr><w:pStyle w:val="Sumario{nivel}"/></w:pPr>{ini}{corpo_e}{fim}</w:p>')
        return "".join(out)

    # ------------------------------------------------------------------ seções
    @staticmethod
    def sect(header_rid: str, inicio_pagina: int | None = None, ultimo: bool = False) -> str:
        pg = f'<w:pgNumType w:start="{inicio_pagina}"/>' if inicio_pagina else ""
        return (f'<w:sectPr><w:headerReference w:type="default" r:id="{header_rid}"/>'
                f'<w:type w:val="nextPage"/><w:pgSz w:w="{PAG_L}" w:h="{PAG_A}"/>'
                f'<w:pgMar w:top="{M_SUP}" w:right="{M_DIR}" w:bottom="{M_INF}" w:left="{M_ESQ}" '
                f'w:header="1134" w:footer="709" w:gutter="0"/>{pg}<w:cols w:space="708"/>'
                '<w:docGrid w:linePitch="360"/></w:sectPr>')

    def conferir_cobertura(self, a_partir_de: int = 0) -> list[int]:
        faltando = []
        for i, el in enumerate(self.body):
            if i < a_partir_de or i in self.usados or i in self.descartados:
                continue
            nome = etree.QName(el).localname
            if nome == "tbl" or (nome == "p" and (texto_paragrafo(el, "original").strip()
                                                 or el.find(".//w:drawing", NS) is not None)):
                faltando.append(i)
        return faltando

    # ------------------------------------------------------------------ pacote
    def salvar(self, destino: str | Path) -> None:
        faltando = self.conferir_cobertura()
        if faltando:
            amostra = [(i, self.texto_original(i)[:50]) for i in faltando[:8]]
            raise ErroSpec(f"parágrafos do original sem destino (texto seria perdido): {amostra}")
        if self._juntar_origem is not None:
            raise ErroSpec(f"parágrafo {self._juntar_origem} pede junção, mas o seguinte não veio")
        if not self._sect_final:
            raise ErroSpec("chame fechar_documento() antes de salvar (sectPr da última seção)")
        corpo_xml = "".join(self.xml).replace("@@SUMARIO@@", self._xml_sumario())
        doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<w:document {NSDECL}><w:body>{corpo_xml}{self._sect_final}</w:body></w:document>')
        partes = {
            "[Content_Types].xml": self._content_types().encode(),
            "_rels/.rels": _RELS_PKG.encode(),
            "docProps/core.xml": self._core().encode(),
            "docProps/app.xml": _APP.encode(),
            "word/document.xml": doc.encode(),
            "word/_rels/document.xml.rels": self._rels_doc().encode(),
            "word/styles.xml": estilos(self.fonte).encode(),
            "word/settings.xml": _SETTINGS.encode(),
            "word/numbering.xml": _NUMBERING.encode(),
            "word/header1.xml": _cabecalho(False).encode(),
            "word/header2.xml": _cabecalho(True).encode(),
            "word/comments.xml": self._comments().encode(),
        }
        for nome, dados in self.imagens.items():
            partes["word/media/" + nome] = dados
        # valida XML bem formado antes de gravar
        for nome, dados in partes.items():
            if nome.endswith(".xml") or nome.endswith(".rels"):
                etree.fromstring(dados)
        salvar(partes, destino)

    _sect_final = ""

    def fechar_documento(self, header_rid: str = "rIdH2") -> None:
        """O sectPr da última seção vai direto no body (não num parágrafo)."""
        self._sect_final = self.sect(header_rid)

    def _comments(self) -> str:
        cs = []
        for c in self.comentarios:
            cs.append(f'<w:comment w:id="{c.id}" w:author="{_t(self.autor)}" w:date="{self.data}" w:initials="RA">'
                      '<w:p><w:pPr><w:pStyle w:val="Textodecomentario"/></w:pPr>'
                      '<w:r><w:rPr><w:rStyle w:val="Refdecomentario"/></w:rPr><w:annotationRef/></w:r>'
                      f'<w:r><w:t xml:space="preserve">{_t(c.texto)}</w:t></w:r></w:p></w:comment>')
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {NSDECL}>'
                + "".join(cs) + "</w:comments>")

    def _rels_doc(self) -> str:
        base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        rels = [("rIdStyles", "styles", "styles.xml"), ("rIdSettings", "settings", "settings.xml"),
                ("rIdNum", "numbering", "numbering.xml"), ("rIdH1", "header", "header1.xml"),
                ("rIdH2", "header", "header2.xml"), ("rIdComments", "comments", "comments.xml")]
        rels += [(f"rIdImg_{n}", "image", f"media/{n}") for n in self.imagens]
        corpo_r = "".join(f'<Relationship Id="{i}" Type="{base}{t}" Target="{alvo}"/>' for i, t, alvo in rels)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{corpo_r}</Relationships>')

    def _content_types(self) -> str:
        wml = "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ov = [("/word/document.xml", f"{wml}.document.main+xml"), ("/word/styles.xml", f"{wml}.styles+xml"),
              ("/word/settings.xml", f"{wml}.settings+xml"), ("/word/numbering.xml", f"{wml}.numbering+xml"),
              ("/word/header1.xml", f"{wml}.header+xml"), ("/word/header2.xml", f"{wml}.header+xml"),
              ("/word/comments.xml", f"{wml}.comments+xml"),
              ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
              ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")]
        o = "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in ov)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
                f'{o}</Types>')

    def _core(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:creator>{_t(self.criador)}</dc:creator><cp:lastModifiedBy>{_t(self.criador)}</cp:lastModifiedBy>'
                f'<dcterms:created xsi:type="dcterms:W3CDTF">{self.data}</dcterms:created>'
                f'<dcterms:modified xsi:type="dcterms:W3CDTF">{self.data}</dcterms:modified></cp:coreProperties>')


# ---------------------------------------------------------------------- partes fixas
_RELS_PKG = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
             '</Relationships>')
_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        '<Application>Microsoft Office Word</Application></Properties>')
# updateFields antes de compat (06, erros técnicos): o Word atualiza o sumário ao abrir
_SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W_NS}">'
             '<w:zoom w:percent="100"/><w:defaultTabStop w:val="709"/><w:hyphenationZone w:val="425"/>'
             '<w:characterSpacingControl w:val="doNotCompress"/><w:updateFields w:val="true"/>'
             '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
             '<w:themeFontLang w:val="pt-BR"/><w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>')
# alíneas: a), b), c) — recuo na própria numeração (06, erros técnicos)
_NUMBERING = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{W_NS}">'
              '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>'
              '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/>'
              '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="357" w:hanging="357"/></w:pPr></w:lvl></w:abstractNum>'
              '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>')


def _cabecalho(com_numero: bool) -> str:
    """Cabeçalho explícito por seção; PAGE com begin/instr/separate/valor/end em runs separados (E2)."""
    if com_numero:
        campo = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                 '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
    else:
        campo = ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NSDECL}>'
            f'<w:p><w:pPr><w:pStyle w:val="Cabealho"/></w:pPr>{campo}</w:p></w:hdr>')


def estilos(fonte: str) -> str:
    """Estilos ABNT (04, A2). Corpo no Normal (recuo 1,25 cm, 1,5, justificado) para o Google Docs."""
    def st(sid, nome, ppr="", rpr="", base="Normal", tipo="paragraph", extra=""):
        b = f'<w:basedOn w:val="{base}"/>' if base else ""
        return (f'<w:style w:type="{tipo}" w:styleId="{sid}"><w:name w:val="{nome}"/>{b}{extra}'
                f'{f"<w:pPr>{ppr}</w:pPr>" if ppr else ""}{f"<w:rPr>{rpr}</w:rPr>" if rpr else ""}</w:style>')
    sp15 = '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
    sp1 = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    sem_recuo = '<w:ind w:left="0" w:firstLine="0"/>'
    tab_sum = f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{LARGURA_TEXTO}"/></w:tabs>'
    s = [
        st("Normal", "Normal", f'<w:widowControl/>{sp15}<w:ind w:firstLine="709"/><w:jc w:val="both"/>',
           base="", extra='<w:qFormat/>'),
        st("Heading1", "heading 1",
           '<w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:spacing w:before="0" w:after="360" w:line="360" w:lineRule="auto"/>'
           f'{sem_recuo}<w:jc w:val="left"/><w:outlineLvl w:val="0"/>', "<w:b/><w:bCs/>", extra='<w:next w:val="Normal"/><w:qFormat/>'),
        st("Heading2", "heading 2",
           '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="360" w:line="360" w:lineRule="auto"/>'
           f'{sem_recuo}<w:jc w:val="left"/><w:outlineLvl w:val="1"/>', "<w:b/><w:bCs/>", extra='<w:next w:val="Normal"/><w:qFormat/>'),
        st("TituloPreTextual", "Título pré-textual",
           f'<w:keepNext/><w:pageBreakBefore/><w:spacing w:before="0" w:after="360" w:line="360" w:lineRule="auto"/>{sem_recuo}<w:jc w:val="center"/>',
           "<w:b/><w:bCs/>"),
        st("Centralizado", "Centralizado", f'{sp15}{sem_recuo}<w:jc w:val="center"/>'),
        st("Natureza", "Natureza do trabalho", f'{sp1}<w:ind w:left="4536" w:firstLine="0"/><w:jc w:val="both"/>'),
        st("Resumo", "Resumo", f'{sp1}{sem_recuo}<w:jc w:val="both"/>'),
        st("PalavrasChave", "Palavras-chave", f'<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>{sem_recuo}<w:jc w:val="both"/>'),
        st("CitacaoLonga", "Citação longa",
           f'<w:spacing w:before="240" w:after="240" w:line="240" w:lineRule="auto"/><w:ind w:left="2268" w:firstLine="0"/><w:jc w:val="both"/>',
           '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Legenda", "Legenda ABNT",
           f'<w:keepNext/><w:spacing w:before="240" w:after="60" w:line="240" w:lineRule="auto"/>{sem_recuo}<w:jc w:val="left"/>',
           '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("FonteIlustracao", "Fonte da ilustração",
           f'<w:spacing w:before="60" w:after="240" w:line="240" w:lineRule="auto"/>{sem_recuo}<w:jc w:val="left"/>',
           '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("TextoTabela", "Texto de tabela", f'{sp1}{sem_recuo}<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Referencias", "Referências",
           f'<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/>{sem_recuo}<w:jc w:val="left"/>'),
        st("Alinea", "Alínea", f'{sp15}<w:ind w:left="357" w:hanging="357"/><w:jc w:val="both"/>'),
        st("Sumario1", "toc 1", f'{tab_sum}{sp15}{sem_recuo}<w:jc w:val="left"/>', "<w:b/><w:bCs/>"),
        st("Sumario2", "toc 2", f'{tab_sum}{sp15}{sem_recuo}<w:jc w:val="left"/>', "<w:b/><w:bCs/>"),
        st("Cabealho", "header", f'{sp1}{sem_recuo}<w:jc w:val="right"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Textodecomentario", "annotation text", f'{sp1}{sem_recuo}<w:jc w:val="left"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>'),
        st("Refdecomentario", "annotation reference", "", '<w:sz w:val="16"/><w:szCs w:val="16"/>',
           base="", tipo="character"),
    ]
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{fonte}" w:hAnsi="{fonte}" w:eastAsia="{fonte}" '
            f'w:cs="{fonte}"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="pt-BR" w:eastAsia="en-US" '
            'w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
            '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            + "".join(s) + "</w:styles>")


# ---------------------------------------------------------------------- sumário em 2 passadas
def paginas_dos_titulos(pdf: str | Path, titulos: list[tuple[int, str, str]], deslocamento: int = 1,
                        a_partir_da_pagina: int = 1) -> dict[str, str]:
    """Procura cada título no PDF da versão limpa, em ordem, e devolve o número exibido da página."""
    import pdfplumber

    def norm(s: str) -> str:
        return re.sub(r"\s+", " ", s).strip().upper()
    res = {}
    with pdfplumber.open(str(pdf)) as doc:
        textos = [norm(p.extract_text() or "") for p in doc.pages]
    atual = a_partir_da_pagina - 1
    for _, texto, bm in titulos:
        alvo = norm(texto)
        for k in range(atual, len(textos)):
            if alvo in textos[k]:
                res[bm] = str(k + 1 - deslocamento)
                atual = k
                break
        else:
            raise ErroSpec(f"título não achado no PDF: {texto!r}")
    return res
