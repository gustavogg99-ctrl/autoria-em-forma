"""Molde de spec para TCC (monografia).

Copiar para entrada/spec_<cliente>.py (fora do repositório), preencher IDX com os índices do
original (dump dos parágrafos: ver README) e escrever as correções de revisão de cada parágrafo.
Os comentários abaixo são textos-padrão das regras do manual de uma instituição do DF já mapeada
(citada como [INSTITUIÇÃO], com os números de item desse manual) e de decisões típicas de curso
(ver docs/09_decisoes_por_curso_exemplo.md). Antes de usar: trocar [INSTITUIÇÃO] e os números de
item pelos do manual do cliente e conferir se cada regra vale para o curso dele.

Uso:
  python3 -I entrada/spec_<cliente>.py saida/<cliente>_MARCADO.docx [entrada/paginas.json]
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]          # entrada/ fica na raiz do repositório
sys.path.insert(0, str(RAIZ / "pipeline"))
from montar_docx import Montador  # noqa: E402

ORIGINAL = RAIZ / "entrada" / "ORIGINAL_cliente.docx"
AUTOR_ARQUIVO = "Nome do Cliente"                     # vai para as propriedades do .docx


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

# ------------------------------------------------ comentários-padrão (06, 3.11) — [INSTITUIÇÃO]
C_CAPA = ("[REGRA] Capa refeita no modelo da [INSTITUIÇÃO] (Manual [INSTITUIÇÃO], 3.1.1 e Fig. 11): instituição no topo, autor, "
          "título em negrito com o subtítulo na linha de baixo, em minúsculas e sem negrito, e local e ano no pé "
          "da folha. A4, margens 3 cm (superior e esquerda) e 2 cm (inferior e direita), Times New Roman 12 "
          "(Manual 2.1 b, c e e).")
C_ROSTO = ("[REGRA] Folha de rosto no modelo da [INSTITUIÇÃO] (Manual 3.2.1.1.1 e Fig. 14): natureza do trabalho e "
           "orientador a 8 cm da margem esquerda, em espaço simples (Manual 2.2 b e d). Texto da natureza mantido.")
C_ROSTO_LOCAL = ("[REGRA] A folha de rosto precisa de local e ano (Manual [INSTITUIÇÃO] 3.2.1.1.1 g e h). Acrescentei "
                 "iguais aos da capa.")
C_APROVACAO = ("[PENDENTE] A folha de aprovação é obrigatória e não existia (Manual [INSTITUIÇÃO] 3.2.1.3). Montei no "
               "modelo (Fig. 17), sem a instituição dos membros da banca, que é a prática dos TCCs aprovados do "
               "curso. Falta o nome e a titulação do membro convidado. Data e assinaturas ficam em branco até a "
               "defesa (Manual 3.2.1.3).")
C_AGRAD = ("[REGRA] Agradecimentos passaram para antes do resumo: a ordem é folha de rosto, folha de "
           "aprovação, dedicatória, agradecimentos, epígrafe, resumo, abstract e sumário (Manual [INSTITUIÇÃO] 3.2.1; "
           "NBR 14724).")
C_RESUMO = ("[REGRA] Resumo e abstract em folhas próprias, com título centralizado, parágrafo único e espaço "
            "simples (Manual [INSTITUIÇÃO] 3.2.1.7 e 3.2.1.8). Texto mantido.")
C_PCHAVE_SEP = ("[DIVERGÊNCIA] O Manual [INSTITUIÇÃO] (3.2.1.7) pede as palavras-chave \"separadas entre si por "
                "ponto\"; a NBR 6028:2021 usa ponto e vírgula. Vale o manual da instituição, que fica acima da "
                "ABNT. Depois do ponto, cada termo começa com maiúscula. Mesmo ajuste nas keywords. Se o "
                "orientador preferir ponto e vírgula, eu volto.")
C_PCHAVE_QTD = ("[DECIDIR] O Manual [INSTITUIÇÃO] (3.2.1.7) pede de 3 a 5 palavras-chave. Escolha qual sai (a mesma nas "
                "keywords). Não tirei nenhuma porque a escolha é sua.")
C_SUMARIO = ("[REGRA] O sumário digitado virou sumário automático, com pontilhado e número de página, gerado a "
             "partir dos títulos do texto (Manual [INSTITUIÇÃO] 3.2.1.13; NBR 6027). Se mexer no trabalho, atualize: "
             "botão direito no sumário → Atualizar campo → Atualizar o índice inteiro.")
C_FORMATO_TEXTO = ("[REGRA] Formatação do texto: Times New Roman 12, espaço 1,5, parágrafo com recuo de "
                   "1,25 cm e justificado, sem espaço extra entre parágrafos; cada seção primária começa em "
                   "página nova, e os títulos ficam separados do texto por um espaço de 1,5 (Manual [INSTITUIÇÃO] 2.1, "
                   "2.2 e 2.3.1 c). Número de página no canto superior direito, a partir daqui, contando desde "
                   "a folha de rosto (Manual 2.4). Ele fica no cabeçalho, não no rodapé.")
C_PONTO_SECAO = ("[REGRA] Sem ponto depois do número da seção (Manual [INSTITUIÇÃO] 2.3.1 a: \"Não deve ser usado nenhum "
                 "sinal (ponto, hífen, etc) para separar o indicativo numérico do título da seção\"). Corrigido "
                 "em todos os títulos.")
C_SIGLA = ("[REGRA] Sigla na primeira vez que aparece: nome por extenso e sigla entre parênteses (Manual [INSTITUIÇÃO] "
           "2.5).")
C_ALINEAS = ("[DIVERGÊNCIA] Objetivos específicos em alíneas com letra, a), b), c)... (Manual [INSTITUIÇÃO] 2.3.2 a). "
             "Inicial maiúscula e ponto final em cada item seguem a orientação da coordenação do curso, que "
             "prevalece sobre o manual da biblioteca (2.3.2 c pede minúscula e ponto e vírgula).")
C_CITACAO_LONGA = ("[REGRA] Citação com mais de três linhas: recuo de 4 cm, fonte 10, espaço simples e sem "
                   "aspas (Manual [INSTITUIÇÃO] 5.1.1). A pontuação final segue a NBR 10520:2002, adotada pelo curso. "
                   "O texto citado não foi alterado.")
C_QUADRO = ("[REGRA] Quadro com título acima e fonte abaixo, os dois em fonte 10 e espaço simples (Manual [INSTITUIÇÃO] "
            "2.1 e, 2.2 b e 2.7 b e c). Conteúdo não foi alterado; o cabeçalho se repete se passar de página.")
C_REFS = ("[REGRA] REFERÊNCIAS sem número e centralizado (Manual [INSTITUIÇÃO] 2.3.3). Lista alinhada à esquerda, em "
          "espaço simples, com uma linha em branco entre as referências (Manual 6.2 c; o 2.2 c fala em espaço "
          "duplo, contradição interna do manual).")
C_DUAS_EDITORAS = "[REGRA] Duas editoras do mesmo local ficam separadas por dois-pontos (Manual [INSTITUIÇÃO] 6.3.5)."
C_CAPITULO_PAGINAS = ("[PENDENTE] Capítulo de livro precisa das páginas inicial e final no fim da referência "
                      "(Manual [INSTITUIÇÃO] 6.4.1.2 a). Me mande as páginas.")
C_TRADUTOR = ("[DECIDIR] Só esta referência traz o tradutor. Elemento complementar tem de aparecer em todas ou "
              "em nenhuma (Manual [INSTITUIÇÃO] 6.2 g). Opções: 1) me mandar o tradutor das outras obras traduzidas; ou "
              "2) tirar a menção daqui. Não completei nada de memória.")
C_PAGINA_CITACAO = ("[PENDENTE] Citação direta sem página. O Manual [INSTITUIÇÃO] (5.1.1) torna a página obrigatória e, "
                    "em documento sem paginação, recomenda indicar o trecho citado.")
C_SEM_FONTE = ("[PENDENTE] Citação direta sem fonte. Toda citação precisa de chamada e de referência (Manual "
               "[INSTITUIÇÃO] 5). Me diga de onde saiu o trecho.")

# ------------------------------------------------ índices do original (preencher a partir do dump)
IDX = {
    "logo": None, "instituicao": [], "autor_capa": None, "titulo_capa": None, "subtitulo_capa": None,
    "local_capa": None, "ano_capa": None, "autor_rosto": None, "titulo_rosto": None, "subtitulo_rosto": None,
    "natureza": None, "orientador": None, "agradecimentos": (None, []), "resumo": (None, None, None),
    "abstract": (None, None, None), "sumario_digitado": range(0), "referencias_titulo": None, "referencias": [],
}


def montar(paginas=None):
    m = Montador(str(ORIGINAL), criador=AUTOR_ARQUIVO)
    if paginas:
        m.paginas_sumario = paginas
    if None in (IDX["autor_capa"], IDX["titulo_capa"]):
        raise SystemExit("Preencha IDX com os índices do original antes de montar.")
    # 1) CAPA — seção 1, cabeçalho vazio. Ajustar os 'before' depois de olhar o PDF.
    #    m.imagem_do_original(IDX["logo"], ppr=sp(after=120, line=240))
    #    m.par(origem=i, estilo="Centralizado", comentario=C_CAPA) ... ; último: sect=m.sect("rIdH1")
    # 2) PRÉ-TEXTUAIS — folha de rosto (C_ROSTO; local/ano novo=True se faltarem), folha de aprovação
    #    (novo=True, primeiro parágrafo com QUEBRA), agradecimentos (TituloPreTextual), resumo
    #    (Resumo + PalavrasChave), abstract (idioma="en-US"), m.descartar(*IDX["sumario_digitado"]),
    #    m.sumario(comentario=C_SUMARIO, sect=m.sect("rIdH1", inicio_pagina=1))
    # 3) TEXTO — títulos: m.par(..., estilo="Heading1", nivel=1, formatacao=False,
    #    edicoes=[{"de": "1. ", "para": "1 ", "comentario": C_PONTO_SECAO}], ppr=SEM_QUEBRA no primeiro);
    #    subtítulos: estilo="Heading2", nivel=2; corpo: estilo="Normal", formatacao=False, com as
    #    edicoes/ancoras da revisão; citação longa: estilo="CitacaoLonga"; alíneas: estilo="Alinea",
    #    ppr=ALINEA; quadro: Legenda + m.tabela_do_original(i) + FonteIlustracao;
    #    REFERÊNCIAS: estilo="Heading1", nivel=1, ppr=CENTRO; entradas: estilo="Referencias".
    # 4) m.fechar_documento("rIdH2")
    raise NotImplementedError("Escreva os blocos do trabalho seguindo os passos acima.")


if __name__ == "__main__":
    destino = Path(sys.argv[1])
    paginas = json.loads(Path(sys.argv[2]).read_text()) if len(sys.argv) > 2 else None
    mm = montar(paginas)
    mm.salvar(destino)
    print(f"ok: {destino} | comentários: {len(mm.comentarios)} | títulos: {len(mm.titulos)}")
