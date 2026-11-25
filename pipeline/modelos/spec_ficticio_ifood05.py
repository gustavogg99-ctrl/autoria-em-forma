"""Spec do teste IFOOD-05 (condição B): TCC FICTÍCIO IESB/Jornalismo, gerado por testes/ifood05/gerar_fixture_ifood05.py
(seed 5, --escala 0.6). Exemplo completo de spec sobre o motor montar_docx; nada aqui é de cliente.

Copiar para entrada/spec_ifood05.py (fora do repositório) junto com entrada/ORIGINAL_fulana_teste.docx:
  python3 -I entrada/spec_ifood05.py saida/ifood05_MARCADO.docx [entrada/paginas.json]
Ou rodar o ciclo inteiro (marcada, limpa, PDF, sumário até estabilizar): ver loop_sumario.py.
"""
import json
import re
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "pipeline"))
from montar_docx import Montador  # noqa: E402

ORIGINAL = RAIZ / "entrada" / "ORIGINAL_fulana_teste.docx"
AUTOR_ARQUIVO = "Fulana Exemplar de Teste"


def sp(before=None, after=None, line=None):
    a = []
    if before is not None:
        a.append(f'w:before="{before}"')
    if after is not None:
        a.append(f'w:after="{after}"')
    if line is not None:
        a.append(f'w:line="{line}" w:lineRule="auto"')
    return {"spacing": f"<w:spacing {' '.join(a)}/>"}


QUEBRA = {"pageBreakBefore": "<w:pageBreakBefore/>"}
SEM_QUEBRA = {"pageBreakBefore": '<w:pageBreakBefore w:val="0"/>'}
CENTRO = {"jc": '<w:jc w:val="center"/>'}
ALINEA = {"numPr": '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'}

# ---------------------------------------------------------------- comentários (06, 3.11)
C_CAPA = ("[REGRA] Capa refeita no modelo IESB (Manual IESB 2024, 3.1.1 e Fig. 11): instituição no topo, autor, "
          "título em negrito com o subtítulo na linha de baixo, em minúsculas e sem negrito, e local e ano no pé "
          "da folha. A4, margens 3 cm (superior e esquerda) e 2 cm (inferior e direita), Arial 12 (fonte "
          "permitida pela coordenação; o manual pede Times New Roman).")
C_TITULO_DIV = ("[DECIDIR] O título da capa (\"um estudo sobre a circulação de jornais\") não é igual ao da "
                "folha de rosto (\"estudo da circulação de jornais\"). Capa e folha de rosto devem ter o mesmo "
                "título (Manual IESB 3.1.1 e 3.2.1.1.1). Não alterei nenhum dos dois: me diga qual vale.")
C_ROSTO = ("[REGRA] Folha de rosto no modelo IESB (Manual 3.2.1.1.1 e Fig. 14): natureza do trabalho e "
           "orientador a 8 cm da margem esquerda, em espaço simples (Manual 2.2 b e d). Texto da natureza mantido.")
C_ROSTO_LOCAL = ("[REGRA] A folha de rosto precisa de local e ano (Manual IESB 3.2.1.1.1 g e h). Acrescentei "
                 "iguais aos da capa.")
C_APROVACAO = ("[PENDENTE] A folha de aprovação é obrigatória e não existia (Manual IESB 3.2.1.3). Montei no "
               "modelo (Fig. 17), sem a instituição dos membros da banca, que é a prática dos TCCs aprovados do "
               "curso. Falta o nome e a titulação do membro convidado. Data e assinaturas ficam em branco até a "
               "defesa (Manual 3.2.1.3).")
C_RESUMO = ("[REGRA] Resumo e abstract em folhas próprias, com título centralizado e espaço simples (Manual IESB "
            "3.2.1.7 e 3.2.1.8). Texto mantido.")
C_RESUMO_PAR = ("[DECIDIR] O resumo deve ser um parágrafo único (NBR 6028:2021); está em dois parágrafos. Juntar "
                "é decisão de redação do autor, então não mexi.")
C_RESUMO_TAM = ("[PENDENTE] O resumo tem cerca de 88 palavras e o Manual IESB (3.2.1.7) pede de 150 a 500. O "
                "abstract (cerca de 51 palavras) acompanha o resumo. Ampliar o texto é com o autor; formatação "
                "não resolve.")
C_PCHAVE_SEP = ("[DIVERGÊNCIA] O Manual IESB 2024 (3.2.1.7) pede as palavras-chave \"separadas entre si por "
                "ponto\"; a NBR 6028:2021 usa ponto e vírgula. Vale o manual da instituição, que fica acima da "
                "ABNT. Depois do ponto, cada termo começa com maiúscula. Mesmo ajuste nas keywords. Se o "
                "orientador preferir ponto e vírgula, eu volto.")
C_SUMARIO = ("[REGRA] O sumário digitado virou sumário automático, com pontilhado e número de página, gerado a "
             "partir dos títulos do texto (Manual IESB 3.2.1.13; NBR 6027). Se mexer no trabalho, atualize: "
             "botão direito no sumário → Atualizar campo → Atualizar o índice inteiro. O digitado tinha títulos "
             "em caixa mista, \"5.\" com ponto e números de página que não batiam com o texto.")
C_FORMATO_TEXTO = ("[REGRA] Formatação do texto: Arial 12, espaço 1,5, parágrafo com recuo de 1,25 cm e "
                   "justificado, sem espaço extra entre parágrafos; cada seção primária começa em página nova, e "
                   "os títulos ficam separados do texto por um espaço de 1,5 (Manual IESB 2.1, 2.2 e 2.3.1 c). "
                   "Número de página no canto superior direito, a partir da Introdução, contando desde a folha de "
                   "rosto (Manual 2.4). Ele fica no cabeçalho, não no rodapé. O corpo do texto estava todo em "
                   "negrito e itálico: tirei.")
C_SECAO_SEC = ("[REGRA] Título de seção secundária em MAIÚSCULAS, como as demais do trabalho (2.1 e 2.3 já "
               "estavam assim) (Manual IESB 2.3.1).")
C_SECAO_PRIM = ("[REGRA] Sem ponto depois do número da seção e título primário em MAIÚSCULAS, como os demais "
                "(Manual IESB 2.3.1 a).")
C_ALINEAS = ("[DIVERGÊNCIA] Objetivos específicos em alíneas com letra, a), b), c)... (Manual IESB 2.3.2 a). "
             "Inicial maiúscula e ponto final em cada item seguem a orientação da coordenação do curso, que "
             "prevalece sobre o manual da biblioteca (2.3.2 c pede minúscula e ponto e vírgula).")
C_CITACAO_LONGA = ("[REGRA] Citação com mais de três linhas: recuo de 4 cm, fonte 10, espaço simples e sem "
                   "aspas (Manual IESB 5.1.1). A pontuação final segue a NBR 10520:2002, adotada pelo curso. "
                   "O texto citado não foi alterado.")
C_CIT_ATRAZO = ("[PENDENTE] Citação direta: não alterei, mas \"atrazo\" e o \"..\" antes da chamada parecem erro de "
                "digitação em relação à fonte. Confira com o original de ARAÚJO (2005, p. 41).")
C_CIT_EN = ("[PENDENTE] Citação direta: \"en crise\" parece erro de digitação. Não alterei (citação direta não "
            "se revisa); confira com MENDES (2008, p. 19).")
C_MANCHETES = ("[PENDENTE] Manchetes entre aspas mantidas como estão (\"Pedese\", \"assignantes\", \"commercio\", "
               "\"typographia\"). Se são grafia de época, está certo; se são erro de digitação, confira com a "
               "fonte. Dado de pesquisa não se corrige aqui.")
C_INTRO_CIT = ("[DECIDIR] A frase termina em dois-pontos (\"observa o seguinte:\") mas o que vem depois não é a "
               "citação de Araújo, que está no parágrafo seguinte. A ordem das frases é com o autor; não alterei.")
C_1831 = ("[FATO] O texto diz que a primeira tipografia da vila foi instalada em 1832; a citação direta de "
          "MARQUES (2011, p. 88), mais adiante, diz 1831. Não alterei nenhum dos dois: confira com a fonte.")
C_CIT_BLOCO = ("[REGRA] A citação direta passou de três linhas, então virou bloco recuado, sem aspas, em fonte 10 e "
               "espaço simples (Manual IESB 5.1.1; NBR 10520:2002, 5.3). O texto citado não foi alterado; só a "
               "frase de apresentação passou a terminar em dois-pontos e a chamada ficou depois do ponto final.")
C_QUADRO = ("[REGRA] Título da tabela acima e fonte abaixo, os dois em fonte 10 e espaço simples (Manual IESB "
            "2.1 e, 2.2 b e 2.7 b e c). Conteúdo e dados não foram alterados; o cabeçalho se repete se passar "
            "de página.")
C_FONTE_NOVA = ("[PENDENTE] A tabela/figura não trazia fonte. Coloquei \"Elaborado pelo autor\" porque é o "
                "padrão do manual para material do próprio autor (Manual IESB 2.7 c). Confirme se é isso ou "
                "informe a fonte.")
C_FIG = ("[REGRA] Título da figura acima da imagem e fonte abaixo (Manual IESB 2.7 b e c), fonte 10. A "
         "imagem não foi alterada.")
C_REFS = ("[REGRA] REFERÊNCIAS sem número e centralizado (Manual IESB 2.3.3). Lista alinhada à esquerda, em "
          "espaço simples, com uma linha em branco entre as referências (Manual 6.2 c; o 2.2 c fala em espaço "
          "duplo, contradição interna do manual).")
C_ORDEM = ("[REGRA] Ordem alfabética (NBR 6023): MARQUES vem antes de MENDES. Troquei a posição das duas "
           "referências.")
C_FERRAZ = ("[PENDENTE] Referência sem ano. A chamada no texto traz (FERRAZ, 2013), mas a lista não tem o ano; "
            "não completei. Confirme o ano e os dados de local e editora ([s.l.]).")
C_SOUZA = ("[REGRA] Referência removida a pedido do autor (SOUZA, Beto, 2019, \"Correio e imprensa\"). Com uma só obra de SOUZA em 2019, as chamadas (SOUZA, 2019) e a referência ficam sem letra (NBR 6023: a letra só vale para mais de uma obra do mesmo autor e ano).")
C_TAVARES = ("[PENDENTE] Referência que não é citada em nenhum ponto do texto (regra: toda referência tem "
             "chamada). Remover da lista ou citar no texto é decisão sua.")
C_PEREIRA = ("[PENDENTE] A chamada traz PEREIRA, 2019, mas a lista só tem PEREIRA, 2018 (Circulação de jornais "
             "no Império). Confirme o ano; não alterei.")
C_COSTA = ("[PENDENTE] Chamada sem referência correspondente: não há COSTA na lista. Me envie a obra "
           "(autor, título, local, editora, ano) ou retire a citação; não completei de memória.")
C_FERRAZ_CHAMADA = ("[PENDENTE] Chamada (FERRAZ, 2013): a referência de FERRAZ na lista não tem ano. Ver "
                    "comentário na lista.")
C_CHAMADA = ("[REGRA] Chamada entre parênteses em MAIÚSCULAS (NBR 10520:2002, item 5, adotada pelo curso).")
C_CHAMADA_E = ("[REGRA] Chamada com duas autorias: sobrenomes em MAIÚSCULAS separados por ponto e vírgula "
               "(NBR 10520:2002, 6.1); como o autor já está no texto, basta o ano entre parênteses (item 5).")
C_ETAL = ("[REGRA] Mais de três autores na chamada: primeiro autor seguido de et al. (NBR 10520:2002, 6.3).")
C_DESTAQUE = ("[REGRA] Destaque uniforme nas referências: o título desta obra estava sem negrito e as outras "
              "estavam com (NBR 6023). Coloquei em negrito.")
C_CONC = ("[REGRA] Concordância de gênero: o sujeito é masculino (\"o custo\", \"o trabalho\", \"o comércio\", "
          "\"esse quadro\"...) e pede o particípio no masculino. Corrigi todos os casos iguais, só comentei este.")

# ---------------------------------------------------------------- índices do original
IDX = {
    "capa_inst": (0, 1), "capa_autor": 2, "capa_titulo": 3, "capa_local": 4, "capa_ano": 5,
    "rosto_autor": 7, "rosto_titulo": 8, "natureza": 9, "orientador": 10,
    "resumo_t": 12, "resumo": (13, 14), "pchave": 15,
    "abstract_t": 17, "abstract": 18, "keywords": 19,
    "sumario_digitado": tuple(range(21, 36)),
}

# ---------------------------------------------------------------- edições do texto
# Erros de texto isolados (revisão mínima): índice -> lista de edições.
TEXTO = {
    39: [{"de": "Os jornais da província circulava", "para": "Os jornais da província circulavam",
          "comentario": "[REGRA] Concordância verbal: sujeito no plural (\"os jornais\")."}],
    49: [{"de": "à partir de", "para": "a partir de",
          "comentario": "[REGRA] A locução é \"a partir de\", sem crase."}],
    58: [{"de": "excessão", "para": "exceção", "comentario": "[REGRA] Grafia: \"exceção\"."}],
    62: [{"de": "Haviam poucos", "para": "Havia poucos",
          "comentario": "[REGRA] \"Haver\" no sentido de existir é impessoal: \"havia\"."}],
    68: [{"de": "porisso", "para": "por isso", "comentario": "[REGRA] Grafia: \"por isso\" (duas palavras)."}],
    80: [{"de": "a medida que", "para": "à medida que",
          "comentario": "[REGRA] A locução é \"à medida que\", com crase."}],
    87: [{"de": "Os editores da época, defendiam", "para": "Os editores da época defendiam",
          "comentario": "[REGRA] Vírgula entre sujeito e verbo: retirada."}],
    95: [{"de": "jornal mas lido", "para": "jornal mais lido", "comentario": "[REGRA] \"mais\" (advérbio), não \"mas\"."}],
    106: [{"de": "Os dados da pesquisa mostra", "para": "Os dados da pesquisa mostram",
           "comentario": "[REGRA] Concordância verbal: \"os dados ... mostram\"."}],
    118: [{"de": "afim de", "para": "a fim de", "comentario": "[REGRA] A locução é \"a fim de\" (três palavras)."}],
    127: [{"de": "Fazem dois séculos", "para": "Faz dois séculos",
           "comentario": "[REGRA] \"Fazer\" indicando tempo decorrido é impessoal: \"faz\"."}],
    137: [{"de": "Concerteza", "para": "Com certeza", "comentario": "[REGRA] Grafia: \"com certeza\"."}],
    82: [{"de": "(SILVA e SOUZA, 2020)", "para": "(2020)", "comentario": C_CHAMADA_E}],
}
# Comentários sem alteração de texto: índice -> lista de âncoras.
ANCORAS = {
    41: [{"ancora": "1832", "comentario": C_1831}],
    55: [{"ancora": "Ao descrever as oficinas, Araújo observa o seguinte:", "comentario": C_INTRO_CIT}],
    67: [{"ancora": "en crise", "comentario": C_CIT_EN}],
    71: [{"ancora": "(PEREIRA, 2019)", "comentario": C_PEREIRA}],
    90: [{"ancora": "(COSTA, 2021)", "comentario": C_COSTA}],
    99: [{"ancora": "(FERRAZ, 2013)", "comentario": C_FERRAZ_CHAMADA}],
    126: [{"ancora": "Pedese", "comentario": C_MANCHETES}],
}

# Concordância (defeito não planejado do gerador, mas erro real do texto): o motor recebe uma edição por caso.
MASC = ["custo do papel importado", "trabalho dos tipógrafos", "comércio de assinaturas", "esse quadro",
        "esse movimento", "esse arranjo", "esse sistema"]
FEM = ["tal configuração", "essa dinâmica", "cobertura dos fatos", "relação entre editores",
       "circulação dos periódicos", "presença da imprensa", "escolha dos assuntos", "disputa entre jornais",
       "produção dos jornais", "imprensa local", "distribuição das folhas", "leitura coletiva", "esfera pública"]
_VISTO = set()


def so_primeira(chave, texto):
    """Erro repetitivo: comentário só na primeira ocorrência de cada tipo (checklist 01, item 81)."""
    if chave in _VISTO:
        return None
    _VISTO.add(chave)
    return texto


PRED = re.compile(r"\b(marcada|limitada|ligada)\b")


def concordancia(txt: str) -> list[dict]:
    out = []
    for m in PRED.finditer(txt):
        ini_frase = max(txt.rfind(". ", 0, m.start()) + 2, 0)
        trecho = txt[ini_frase:m.start()].lower()
        melhor = (-1, None)
        for nome, lista in (("m", MASC), ("f", FEM)):
            for s in lista:
                k = trecho.rfind(s)
                if k > melhor[0]:
                    melhor = (k, nome)
        if melhor[1] is None:
            raise SystemExit(f"sujeito não identificado: {txt[ini_frase:m.end()]!r}")
        if melhor[1] == "m":
            antes = txt[max(0, m.start() - 12):m.start()]
            de = antes.split()[-1] + " " + m.group(1) if antes.split() else m.group(1)
            n = txt[:m.end()].count(de)
            out.append({"de": de, "para": de[:-1] + "o", "n": n, "comentario": so_primeira("conc", C_CONC)})
    return out


CAIXA = re.compile(r"\((?![A-ZÁ-Ú]{2})([A-ZÁ-Ú][a-zá-ú]+(?:(?:; | e )[A-ZÁ-Ú][a-zá-ú]+)*, \d{4}[a-z]?(?:, p\. \d+)?)\)")


def chamadas_caixa(txt: str) -> list[dict]:
    out = []
    for m in CAIXA.finditer(txt):
        de = m.group(0)
        sobrenomes, resto = de[1:-1].split(", ", 1)
        para = "(" + sobrenomes.upper() + ", " + resto + ")"
        out.append({"de": de, "para": para, "n": txt[:m.end()].count(de), "comentario": so_primeira("caixa", C_CHAMADA)})
    return out


def edicoes_do(m: Montador, i: int) -> list[dict]:
    txt = m.texto_original(i)
    ed = list(TEXTO.get(i, [])) + chamadas_caixa(txt) + concordancia(txt)
    # edições que se sobrepõem a outra do mesmo parágrafo: a manual vence
    return ed


def corpo(m: Montador, i: int, **kw):
    m.par(origem=i, estilo="Normal", formatacao=False, edicoes=edicoes_do(m, i), ancoras=ANCORAS.get(i, []), **kw)


def montar(paginas=None):
    _VISTO.clear()
    m = Montador(str(ORIGINAL), criador=AUTOR_ARQUIVO, fonte="Arial")
    if paginas:
        m.paginas_sumario = paginas
    # ------------------------------------------------------------- 1) CAPA (seção 1)
    m.par(origem=0, estilo="Centralizado", comentario=C_CAPA, ppr=sp(after=0, line=360))
    m.par(origem=1, estilo="Centralizado", ppr=sp(after=0, line=360))
    m.par(origem=2, estilo="Centralizado", ppr=sp(before=2400, after=0, line=360))
    m.par(origem=3, estilo="Centralizado", ppr=sp(before=2400, after=0, line=360),
          edicoes=[{"de": ": um estudo sobre a circulação de jornais", "para": ":"}], comentario=C_TITULO_DIV)
    m.par(texto="um estudo sobre a circulação de jornais", novo=True, estilo="Centralizado",
          ppr=sp(before=0, after=0, line=360))
    m.par(origem=4, estilo="Centralizado", formatacao=False, ppr=sp(before=6000, after=0, line=360))
    m.par(origem=5, estilo="Centralizado", formatacao=False, ppr=sp(after=0, line=360),
          sect=m.sect("rIdH1"))
    # ------------------------------------------------------------- 2) PRÉ-TEXTUAIS (seção 2)
    m.par(origem=7, estilo="Centralizado", ppr={**sp(after=0, line=360), **SEM_QUEBRA}, comentario=C_ROSTO)
    m.par(origem=8, estilo="Centralizado", ppr=sp(before=2400, after=0, line=360),
          edicoes=[{"de": ": estudo da circulação de jornais", "para": ":"}], comentario=C_TITULO_DIV)
    m.par(texto="estudo da circulação de jornais", novo=True, estilo="Centralizado",
          ppr=sp(before=0, after=0, line=360))
    m.par(origem=9, estilo="Natureza", formatacao=False, ppr=sp(before=2400, after=0, line=240))
    m.par(origem=10, estilo="Natureza", formatacao=False, ppr=sp(before=240, after=0, line=240))
    m.par(texto="Brasília", novo=True, estilo="Centralizado", ppr=sp(before=4400, after=0, line=360),
          comentario=C_ROSTO_LOCAL)
    m.par(texto="2026", novo=True, estilo="Centralizado", ppr=sp(after=0, line=360))
    # folha de aprovação (nova)
    m.par(texto="FULANA EXEMPLAR DE TESTE", novo=True, estilo="Centralizado",
          ppr={**QUEBRA, **sp(after=0, line=360)}, props_novas=frozenset({"b"}), comentario=C_APROVACAO)
    m.par(texto="A IMPRENSA DE PROVÍNCIA E O LEITOR:", novo=True, estilo="Centralizado",
          ppr=sp(before=1200, after=0, line=360), props_novas=frozenset({"b"}))
    m.par(texto="um estudo sobre a circulação de jornais", novo=True, estilo="Centralizado",
          ppr=sp(after=0, line=360))
    m.par(texto=m.texto_original(9) + " Aprovado em: ___/___/______.", novo=True, estilo="Natureza",
          ppr=sp(before=1200, after=0, line=240))
    m.par(texto="BANCA EXAMINADORA", novo=True, estilo="Centralizado", ppr=sp(before=1200, after=0, line=360),
          props_novas=frozenset({"b"}))
    m.par(texto="______________________________________", novo=True, estilo="Centralizado",
          ppr=sp(before=720, after=0, line=240))
    m.par(texto="Prof. Beltrano de Teste – Orientador", novo=True, estilo="Centralizado", ppr=sp(after=0, line=240))
    m.par(texto="______________________________________", novo=True, estilo="Centralizado",
          ppr=sp(before=720, after=0, line=240))
    m.par(texto="Prof. ______________________________ – Membro convidado", novo=True, estilo="Centralizado",
          ppr=sp(after=0, line=240))
    # resumo
    m.par(origem=12, estilo="TituloPreTextual", formatacao=False, comentario=C_RESUMO)
    m.par(origem=13, estilo="Resumo", formatacao=False, comentario=C_RESUMO_PAR)
    m.par(origem=14, estilo="Resumo", formatacao=False, comentario=C_RESUMO_TAM,
          ppr=sp(before=0, after=0, line=240))
    m.par(origem=15, estilo="PalavrasChave", formatacao=False,
          edicoes=[{"de": "imprensa; província; circulação; leitores.", "para": "Imprensa. Província. Circulação. Leitores.",
                    "comentario": C_PCHAVE_SEP}])
    # abstract
    m.par(origem=17, estilo="TituloPreTextual", formatacao=False)
    m.par(origem=18, estilo="Resumo", formatacao=False, idioma="en-US")
    m.par(origem=19, estilo="PalavrasChave", formatacao=False, idioma="en-US",
          edicoes=[{"de": "press; province; circulation; readers.", "para": "Press. Province. Circulation. Readers."}])
    # sumário
    m.descartar(*IDX["sumario_digitado"])
    m.sumario(comentario=C_SUMARIO, sect=m.sect("rIdH1", inicio_pagina=1))
    # ------------------------------------------------------------- 3) TEXTO (seção 3)
    prim = lambda i, **kw: m.par(origem=i, estilo="Heading1", nivel=1, formatacao=False, **kw)
    sec = lambda i, **kw: m.par(origem=i, estilo="Heading2", nivel=2, formatacao=False, **kw)

    prim(36, ppr=SEM_QUEBRA, comentario=C_FORMATO_TEXTO)
    for i in (37, 38, 39, 40, 41):
        corpo(m, i)
    corpo(m, 42)
    corpo(m, 43)
    ali = {44: "Identificar os jornais de maior circulação na região.",
           45: "Descrever as rotas de distribuição das folhas.",
           46: "Comparar o perfil dos leitores entre as décadas."}
    for k, i in enumerate((44, 45, 46)):
        de = m.texto_original(i)
        m.par(origem=i, estilo="Alinea", formatacao=False, ppr=ALINEA,
              edicoes=[{"de": de, "para": ali[i], "comentario": C_ALINEAS if k == 0 else None}])
    for i in range(47, 51):
        corpo(m, i)
    prim(51)
    sec(52)
    # 53-55 ... 56 é citação longa protegida
    for i in (53, 54, 55):
        corpo(m, i)
    m.par(origem=56, estilo="CitacaoLonga", formatacao=False, comentario=C_CITACAO_LONGA,
          ancoras=[{"ancora": "atrazo", "comentario": C_CIT_ATRAZO}])
    for i in range(57, 64):
        corpo(m, i)
    sec(64, edicoes=[{"de": "Leitores e alfabetização", "para": "LEITORES E ALFABETIZAÇÃO", "comentario": C_SECAO_SEC}])
    for i in range(65, 75):
        corpo(m, i)
    sec(75)
    for i in range(76, 79):
        corpo(m, i)
    t79 = m.texto_original(79)
    ini = t79.index('"') + 1
    fim = t79.rindex('"')
    citado = t79[ini:fim]                      # texto da citação, sem aspas, intocado
    m.par(origem=79, estilo="Normal", formatacao=False,
          edicoes=[{"de": t79[len("Para Marques"):], "para": ":", "comentario": C_CIT_BLOCO}])
    m.par(texto=citado + " (MARQUES, 2011, p. 88).", novo=True, estilo="CitacaoLonga")
    for i in range(80, 91):
        corpo(m, i)
    prim(91)
    sec(92)
    for i in range(93, 102):
        corpo(m, i)
    sec(102)
    for i in range(103, 111):
        corpo(m, i)
    prim(111)
    sec(112)
    corpo(m, 113)
    m.par(origem=115, estilo="Legenda", formatacao=False, comentario=C_QUADRO)
    m.tabela_do_original(114)
    m.par(texto="Fonte: Elaborado pelo autor.", novo=True, estilo="FonteIlustracao", comentario=C_FONTE_NOVA)
    for i in range(116, 120):
        corpo(m, i)
    m.par(origem=121, estilo="Legenda", formatacao=False, comentario=C_FIG)
    m.imagem_do_original(120, ppr=CENTRO)
    m.par(texto="Fonte: Elaborado pelo autor.", novo=True, estilo="FonteIlustracao", comentario=C_FONTE_NOVA)
    corpo(m, 122)
    corpo(m, 123)
    sec(124)
    for i in range(125, 134):
        corpo(m, i)
    sec(134)
    for i in range(135, 142):
        corpo(m, i)
    prim(142, edicoes=[{"de": "5. Considerações Finais", "para": "5 CONSIDERAÇÕES FINAIS", "comentario": C_SECAO_PRIM}])
    for i in range(143, 146):
        corpo(m, i)
    # ------------------------------------------------------------- REFERÊNCIAS
    m.par(origem=146, estilo="Heading1", nivel=1, formatacao=False, ppr=CENTRO, comentario=C_REFS)
    ref = lambda i, **kw: m.par(origem=i, estilo="Referencias", **kw)
    for i in (147, 148, 149):
        ref(i)
    ref(150, ancoras=[{"ancora": "[s.l.]: Editora Fictícia.", "comentario": C_FERRAZ}])
    for i in (151,):
        ref(i)
    ref(153, comentario=C_ORDEM)       # MARQUES antes de MENDES
    ref(152, comentario=C_ORDEM)
    ref(154)
    ref(155, comentario=C_DESTAQUE)    # PEREIRA vem antes de SILVA (NBR 6023); a ordem do original já está certa
    _t = "PEREIRA, João. Circulação de jornais no Império. São Paulo: Editora Fictícia, 2018."
    _novo = ('PEREIRA, João. </w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Circulação de jornais no Império'
             '</w:t></w:r><w:r><w:t xml:space="preserve">. São Paulo: Editora Fictícia, 2018.')
    assert _t in m.xml[-1], "PEREIRA não achado no XML"
    m.xml[-1] = m.xml[-1].replace(_t, _novo)
    ref(156)                           # SILVA e SOUZA, depois de PEREIRA
    ref(158)                           # SOUZA 2019 (Censura e tipografia): única obra que fica, sem letra
    ref(157, excluir=True, comentario=C_SOUZA)   # 2ª obra de SOUZA removida a pedido do autor
    ref(159, comentario=C_TAVARES)
    # ------------------------------------------------------------- 4) fim
    m.fechar_documento("rIdH2")
    falt = m.conferir_cobertura()
    if falt:
        raise SystemExit(f"parágrafos do original sem destino: {falt}")
    return m


if __name__ == "__main__":
    destino = Path(sys.argv[1])
    paginas = json.loads(Path(sys.argv[2]).read_text()) if len(sys.argv) > 2 else None
    mm = montar(paginas)
    mm.salvar(destino)
    print(f"ok: {destino} | comentários: {len(mm.comentarios)} | títulos: {len(mm.titulos)}")
