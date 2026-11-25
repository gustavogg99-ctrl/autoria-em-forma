"""Leitura e escrita segura de .docx, e extração de texto com ou sem as alterações marcadas.

Base comum de todos os scripts do pipeline. Segue o 06, seção 8:
- só .docx: recusa .docm e qualquer arquivo com macro (vbaProject.bin);
- descompacta com limite de tamanho e de taxa de compressão (zip bomb);
- lê XML sem DTD, sem entidades externas e sem rede (XXE).
"""
from __future__ import annotations

import re
import unicodedata
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}


def w(tag: str) -> str:
    """Nome qualificado no namespace w: w('p') -> '{...}p'."""
    return f"{{{W}}}{tag}"


LIMITE_TOTAL = 300 * 1024 * 1024   # 300 MB descompactados
LIMITE_PARTE = 100 * 1024 * 1024   # 100 MB por parte
LIMITE_TAXA = 200                   # compressão acima de 200:1 é suspeita

PARSER = etree.XMLParser(
    resolve_entities=False, no_network=True, load_dtd=False, dtd_validation=False, huge_tree=False
)


class ArquivoRecusado(Exception):
    pass


def abrir(caminho: str | Path) -> dict[str, bytes]:
    """Lê o pacote inteiro em memória, na ordem original. Levanta ArquivoRecusado se inseguro."""
    caminho = Path(caminho)
    if caminho.suffix.lower() != ".docx":
        raise ArquivoRecusado(f"{caminho.name}: só .docx é aceito (06, seção 8).")
    with zipfile.ZipFile(caminho) as z:
        total = 0
        for info in z.infolist():
            if info.file_size > LIMITE_PARTE:
                raise ArquivoRecusado(f"{info.filename}: parte grande demais ({info.file_size} bytes).")
            if info.compress_size and info.file_size / info.compress_size > LIMITE_TAXA and info.file_size > 1_000_000:
                raise ArquivoRecusado(f"{info.filename}: taxa de compressão suspeita (zip bomb).")
            total += info.file_size
            if total > LIMITE_TOTAL:
                raise ArquivoRecusado("Pacote descompactado passa do limite (zip bomb).")
            if info.filename.lower().endswith("vbaproject.bin"):
                raise ArquivoRecusado("Arquivo com macro (vbaProject.bin). Pedir .docx sem macro.")
        partes = {info.filename: z.read(info.filename) for info in z.infolist()}
    ct = partes.get("[Content_Types].xml", b"")
    if b"macroEnabled" in ct:
        raise ArquivoRecusado("Content type de documento com macro. Pedir .docx sem macro.")
    return partes


def xml(dados: bytes) -> etree._Element:
    return etree.fromstring(dados, PARSER)


def serializar(raiz: etree._Element) -> bytes:
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def salvar(partes: dict[str, bytes], destino: str | Path) -> None:
    """Grava o pacote; [Content_Types].xml sempre primeiro."""
    nomes = sorted(partes, key=lambda n: (n != "[Content_Types].xml",))
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for nome in nomes:
            z.writestr(nome, partes[nome])


def partes_de_texto(partes: dict[str, bytes]) -> list[str]:
    """Partes que podem ter texto, marcações ou comentários."""
    padrao = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")
    return [n for n in partes if padrao.match(n)]


# ---------- texto ----------

def texto_paragrafo(p: etree._Element, visao: str = "final") -> str:
    """Texto de um w:p.

    visao='final'    -> como fica com as alterações aceitas (ignora w:del, inclui w:ins)
    visao='original' -> como era antes das alterações (inclui w:delText, ignora w:ins)
    """
    partes: list[str] = []
    for el in p.iter():
        tag = etree.QName(el).localname if isinstance(el.tag, str) else ""
        if tag not in ("t", "delText", "tab", "br", "noBreakHyphen"):
            continue
        dentro_ins = any(etree.QName(a).localname in ("ins", "moveTo") for a in el.iterancestors() if isinstance(a.tag, str))
        dentro_del = any(etree.QName(a).localname in ("del", "moveFrom") for a in el.iterancestors() if isinstance(a.tag, str))
        if visao == "final" and (tag == "delText" or dentro_del):
            continue
        if visao == "original" and dentro_ins:
            continue
        if tag in ("t", "delText"):
            partes.append(el.text or "")
        elif tag == "tab":
            partes.append("\t")
        elif tag == "br":
            partes.append("\n")
        elif tag == "noBreakHyphen":
            partes.append("-")
    return "".join(partes)


def normalizar(texto: str) -> str:
    """Normalização para comparar trechos: NFC, espaços colapsados, aspas e travessões unificados."""
    t = unicodedata.normalize("NFC", texto)
    t = t.replace(" ", " ").replace(" ", " ")
    t = re.sub(r"[“”„«»]", '"', t)
    t = re.sub(r"[‘’‚]", "'", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


# ---------- estilos e recuo ----------

def mapa_estilos(partes: dict[str, bytes]) -> dict[str, etree._Element]:
    if "word/styles.xml" not in partes:
        return {}
    raiz = xml(partes["word/styles.xml"])
    return {s.get(w("styleId")): s for s in raiz.findall("w:style", NS)}


def _ind_estilo(estilos: dict, style_id: str | None, profundidade: int = 0) -> int | None:
    if not style_id or style_id not in estilos or profundidade > 10:
        return None
    s = estilos[style_id]
    ind = s.find("w:pPr/w:ind", NS)
    if ind is not None:
        val = ind.get(w("left")) or ind.get(w("start"))
        if val is not None:
            return int(float(val))
    base = s.find("w:basedOn", NS)
    return _ind_estilo(estilos, base.get(w("val")) if base is not None else None, profundidade + 1)


def recuo_esquerdo(p: etree._Element, estilos: dict) -> int:
    """Recuo esquerdo efetivo do parágrafo, em twips (direto > estilo > Normal)."""
    ind = p.find("w:pPr/w:ind", NS)
    if ind is not None:
        val = ind.get(w("left")) or ind.get(w("start"))
        if val is not None:
            return int(float(val))
    est = p.find("w:pPr/w:pStyle", NS)
    v = _ind_estilo(estilos, est.get(w("val")) if est is not None else "Normal")
    if v is None:
        v = _ind_estilo(estilos, "Normal")
    return v or 0


def nome_estilo(p: etree._Element, estilos: dict) -> str:
    est = p.find("w:pPr/w:pStyle", NS)
    if est is None:
        return ""
    s = estilos.get(est.get(w("val")))
    nome = s.find("w:name", NS) if s is not None else None
    return (nome.get(w("val")) if nome is not None else est.get(w("val"))) or ""


RECUO_CITACAO_LONGA = 2200  # 4 cm = 2268 twips; tolerância para arredondamentos


def corpo(partes: dict[str, bytes]) -> etree._Element:
    return xml(partes["word/document.xml"]).find("w:body", NS)
