# Guia de Configuração e Uso do Modelo Word ABNT

Objetivo: um modelo mestre (`MODELO_ABNT_MESTRE.dotx`) em que **toda** formatação vem de estilos. Nada de formatação manual: é ela que gera sumário errado, numeração quebrada e retrabalho.

**Relação com o pipeline:** em produção, a formatação é feita pelo script de build descrito no `06_Licoes_Aprendidas_TCC.md` (seção 4). Este guia vale para ajustes manuais, para conferir o resultado do script no Word e para entender o que ele faz. Onde este guia e o 06 divergirem, vale o 06, que reflete o que foi testado em trabalho real.

Referências de menu: Word 365 em português.

Última atualização: 20/10/2026 — cabeçalho explícito por seção (A4), duas versões na entrega (B2, B12, B13), novo sintoma na Parte C.

---

## Parte A — Montar o modelo (uma vez)

### A1. Página
**Layout → Margens → Margens personalizadas**
- Superior 3 cm · Esquerda 3 cm · Inferior 2 cm · Direita 2 cm
- Papel: A4
- Cabeçalho: 2 cm da borda (para o número de página ficar a 2 cm do topo)
- Frente e verso: marcar **Margens espelho** (interna 3 cm, externa 2 cm)

### A2. Estilos de parágrafo
**Página Inicial → Estilos → Gerenciar estilos**. Criar/ajustar:

| Estilo | Configuração |
|---|---|
| **Normal / Texto ABNT** | Fonte 12 (Arial ou Times, conforme manual) · justificado · recuo 1ª linha 1,25 cm · espaço 1,5 · antes/depois 0 pt |
| **Título 1** (seção primária) | 12, MAIÚSCULAS, negrito · alinhado à esquerda · sem recuo · espaço 1,5 · antes 0 / depois 18 pt · **quebra de página antes** · nível de tópico 1 |
| **Título 2** | 12, MAIÚSCULAS, sem negrito · esquerda · antes/depois 18 pt · nível 2 |
| **Título 3** | 12, negrito · esquerda · antes/depois 18 pt · nível 3 |
| **Título 4 / 5** | 12, itálico / normal · esquerda · nível 4 / 5 |
| **Título Pré-textual** | 12, MAIÚSCULAS, negrito, **centralizado** · quebra de página antes · **nível de tópico: corpo de texto** (não entra no sumário) |
| **Título Pós-textual** | igual ao anterior, mas **nível de tópico 1** (entra no sumário sem número: REFERÊNCIAS, APÊNDICE A – …, ANEXO A – …) |
| **Citação Longa** | 10 · justificado · recuo esquerdo 4 cm · sem recuo de 1ª linha · espaço simples · antes/depois 12 pt |
| **Referências** | 12 · **alinhado à esquerda** · sem recuo · espaço simples · depois 12 pt |
| **Legenda** | 10 · centralizado ou à esquerda (conforme manual) · espaço simples |
| **Fonte de Ilustração** | 10 · mesmo alinhamento da legenda · espaço simples · depois 12 pt |
| **Texto de nota de rodapé** | 10 · justificado · espaço simples |
| **Natureza do Trabalho** | 10 ou 12 · justificado · recuo esquerdo 8 cm · espaço simples |
| **Resumo** | 12 · justificado · **sem** recuo de 1ª linha · espaço 1,5 |
| **Tabela** (texto em células) | 10 ou 12 conforme manual · **recuo de 1ª linha 0** · espaço simples (o recuo herdado do Normal desalinha células — 06, erros técnicos) |

Obs.: 18 pt ≈ uma linha em branco em espaço 1,5. Destaques dos níveis de título podem variar — seguir o manual da instituição ou as decisões do curso (05).

### A3. Numeração dos títulos (NBR 6024)
**Página Inicial → Lista de vários níveis → Definir nova lista de vários níveis → Mais**
- Nível 1 vinculado a Título 1: formato `1` (sem ponto depois)
- Nível 2 → Título 2: `1.1` · Nível 3 → Título 3: `1.1.1` · e assim por diante
- "Seguir número com": **Espaço** (não tabulação)
- Posição do número e recuo do texto: 0 cm

### A4. Seções e paginação
A lógica: folhas são **contadas a partir da folha de rosto**, mas o número só **aparece a partir da Introdução**. Cada seção tem **cabeçalho próprio e explícito**; nunca depender de "Vincular ao Anterior" (erro E2 do 06).

1. **Seção 1 — Capa.** Cabeçalho vazio. Fim da capa: **Layout → Quebras → Próxima página**.
2. **Seção 2 — Pré-textuais** (folha de rosto até sumário). Desmarcar "Vincular ao Anterior" e deixar o cabeçalho vazio. **Inserir → Número de Página → Formatar** → *Iniciar em: 1*. Não inserir número visível.
3. **Seção 3 — Textuais e pós-textuais** (Introdução em diante). No cabeçalho, **desmarcar "Vincular ao Anterior"**, inserir número no canto superior direito (fonte 10), formato **"Continuar da seção anterior"**.
4. **Páginas em paisagem** (tabelas largas, apêndices): seção própria, com "Vincular ao Anterior" desmarcado e o número inserido de novo no cabeçalho dessa seção. Na volta para retrato, repetir.
5. Conferir: se a folha de rosto é a 1 e há 6 folhas pré-textuais, a Introdução deve mostrar o número 7.
6. Frente e verso: usar **Página ímpar** como tipo de quebra antes de seções primárias e numerar no canto superior direito (anverso) e esquerdo (verso) com **Diferente em páginas pares e ímpares**.
7. **Testar no programa do cliente.** Se ele usa Google Docs, seguir o teste da seção 5 do 06.

### A5. Sumário
**Referências → Sumário → Sumário personalizado → Opções**
- Mapear: Título 1 = nível 1 · Título 2 = 2 · Título 3 = 3 · Título Pós-textual = 1
- Remover Título Pré-textual da lista
- Mostrar números de página alinhados à direita com pontilhado
- Ajustar os estilos **Sumário 1, 2, 3** para replicar o destaque de cada nível de título

### A6. Legendas e listas
**Referências → Inserir Legenda**
- Criar rótulos: Figura, Quadro, Gráfico, Tabela
- Posição: **Acima do item selecionado**
- Separador entre número e título: **travessão (–)** → resultado "Figura 1 – Título"
- Listas: **Referências → Inserir Índice de Ilustrações**, uma por rótulo

### A7. Notas de rodapé
**Referências → Notas de Rodapé (inicializador)** → numeração 1, 2, 3 contínua no documento.

### A8. Elementos fixos
Montar no modelo, já com estilos aplicados e textos de exemplo em [colchetes]: capa, folha de rosto (com Natureza do Trabalho), folha de aprovação, resumo + palavras-chave, abstract + keywords, sumário, 1 INTRODUÇÃO, REFERÊNCIAS, APÊNDICE A, ANEXO A.

Salvar como **Modelo do Word (.dotx)**.

---

## Parte B — Aplicar o modelo em um trabalho de cliente

1. **Preservar o original:** salvar como `ORIGINAL_[cliente].docx` e não tocar mais nele. Arquivo .docm ou com macro: não abrir com macros habilitadas; pedir o .docx (06, seção 8).
2. **Nome do revisor neutro:** antes de editar, **Arquivo → Opções → Geral → Nome de usuário: "Revisão ABNT"**. É o nome que aparece em cada marcação e comentário.
3. **Novo documento a partir do modelo:** abrir o `.dotx` → salvar como `[cliente]_FORMATADO_v1.docx`.
4. **Transferir o conteúdo por seção:** copiar do original e colar com **Colar especial → Mesclar formatação**. Isso adota os estilos do modelo e mantém itálicos e negritos de ênfase (termos estrangeiros, nomes científicos).
5. **Aplicar estilos:** títulos → Título 1/2/3; citações longas → Citação Longa; referências → Referências. Usar o **Painel de Navegação** (Exibir → Painel de Navegação) para conferir a hierarquia.
6. **Limpeza em massa (Localizar e Substituir, Ctrl+H):**
   - dois espaços → um espaço (repetir até zerar)
   - `^t` (tabulação) no início de parágrafos → nada
   - `^l` (quebra de linha manual) → `^p` ou espaço, conferindo caso a caso
   - `^p^p` (parágrafos vazios) → `^p` no corpo do texto
7. **Ilustrações e tabelas:** reinserir legendas pelo recurso automático; adicionar "Fonte:" abaixo; tabelas com bordas laterais removidas.
8. **Correções de texto com controle de alterações ligado** (Revisão → Controlar Alterações): revisão, citações e referências. Comentários com prefixo ([REGRA], [PENDENTE], [DECIDIR], [DIVERGÊNCIA], [FATO]) conforme 06, 3.11. Nada dentro de citação direta, transcrição ou dado de pesquisa.
9. **Citações e referências:** conferir com o checklist (seções 4 e 6). Para cruzar citações × referências, pedir no projeto: *"compare as citações do texto com a lista de referências e liste divergências"*. **Nunca** pedir para a IA completar referências.
10. **Atualizar tudo:** Ctrl+A → F9 → "Atualizar o índice inteiro".
11. **Revisão visual:** **Exibir → Várias Páginas** e percorrer o documento procurando títulos órfãos no fim de página, figuras cortadas e páginas em branco.
12. **Gerar as duas versões:**
    - **Marcada:** salvar como `[cliente]_MARCADO.docx`, com controle de alterações e comentários preservados.
    - **Limpa:** a partir da marcada, **Revisão → Aceitar todas as alterações** e **Excluir todos os comentários**; salvar como `[cliente]_FINAL.docx`; atualizar de novo o sumário (Ctrl+A → F9), porque a paginação muda; **Arquivo → Informações → Verificar Problemas → Inspecionar Documento** (remover revisões, comentários e propriedades pessoais; autor = cliente).
13. **Rodar o checklist** `01_Checklist_QA_ABNT.md` completo.
14. **Exportar PDFs:**
    - **Final:** a partir da versão limpa, Arquivo → Exportar → PDF → Opções → marcar **Criar indicadores usando: Títulos**. Abrir o PDF e conferir paginação e sumário.
    - **Marcado (prévia):** a partir da versão marcada, com Revisão → "Todas as marcações" ativo, exportar com **Publicar: Documento com marcação**. Conferir se marcas e comentários aparecem.
    - Marca d'água nas prévias: **Design → Marca D'água** antes de exportar (não salvar o .docx com ela).

---

## Parte C — Problemas comuns e correção rápida

| Sintoma | Causa provável | Correção |
|---|---|---|
| Sumário com títulos faltando ou sobrando | Estilo errado aplicado ou nível de tópico errado | Reaplicar estilo; conferir nível de tópico em Parágrafo |
| Numeração de títulos reinicia ou pula | Lista de vários níveis desvinculada | Reaplicar a lista do A3 ao Título 1 |
| Número de página aparece na capa | "Vincular ao Anterior" ativo | Desmarcar na seção 3 e apagar números das seções 1–2 |
| Introdução começa em 1 | Seção 3 com "Iniciar em 1" | Trocar para "Continuar da seção anterior" |
| Número de página some numa parte do documento ou no Google Docs | Seção sem cabeçalho próprio (dependia de "Vincular ao Anterior") ou campo PAGE reaproveitado de arquivo exportado do Google Docs | Cabeçalho explícito em cada seção, inclusive paisagem; inserir o número de novo (06, 3.6 e erro E2) |
| Sumário com páginas erradas na versão final | Sumário atualizado na versão marcada | Atualizar de novo depois de aceitar as alterações (06, 3.7) |
| Lista de figuras vazia | Legendas digitadas manualmente | Reinserir com Inserir Legenda |
| Espaço estranho entre parágrafos | Espaço antes/depois herdado do original | Selecionar tudo → Limpar formatação só no corpo → reaplicar Texto ABNT |
| Citação longa com aspas | Herdado do original | Remover aspas (NBR 10520, tanto 2002 quanto 2023) |
| Marcações aparecem com o nome pessoal do revisor | Nome de usuário do Word não configurado | B2, antes de começar; em arquivo já marcado, refazer a partir do original |
