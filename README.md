# Autoria em Forma — pipeline de formatação ABNT e revisão de TCC

Pipeline em Python que formata e revisa trabalhos acadêmicos (TCC, monografia) seguindo a ABNT e o manual de cada faculdade. Ele entrega uma versão **marcada**, com cada correção no controle de alterações do Word e um comentário que diz a regra, o item e a edição da norma, e uma versão **limpa**, pronta para entregar. Uma bateria de testes automáticos impede a entrega se algo quebrar.

Foi feito para a [Autoria em Forma](https://site.autoriaemforma.workers.dev), serviço de formatação e revisão de TCC em Brasília/DF (atendimento pausado). O código é aberto para estudo e reuso.

**A divisão de trabalho:** a IA decide **o que** corrigir. O código garante **como** a correção é aplicada: sem perder texto do original, sem alterar citação, com sumário e paginação conferidos.

**Resultado medido** (teste controlado, n=1, documento fictício de 29 páginas; método em [`testes/ifood05/README.md`](testes/ifood05/README.md)): 5h14 à mão contra 18 min com o pipeline, dos quais 3min45 de tempo ativo de quem decide.

> **Privacidade:** nenhum trabalho de cliente está ou esteve neste repositório. O `.gitignore` bloqueia `.docx`, `.pdf` e as pastas `entrada/` e `saida/`. Os documentos de teste são gerados na hora pelo autoteste, com conteúdo fictício.

> **Método e regras:** os números entre parênteses no código e neste README, como "(06, seção 4)", apontam para os documentos em [`docs/`](docs/): as instruções do operador (o prompt do projeto), o checklist de QA, a matriz de manuais das faculdades do DF, as lições aprendidas e o template do relatório.

## Como rodar

```bash
git clone https://github.com/gustavogg99-ctrl/autoria-em-forma
cd autoria-em-forma
pip install --break-system-packages -r requirements.txt   # lxml, pdfplumber, pymupdf
python3 -I testes/autoteste/rodar.py                       # tem que dar 41/41
```

Requer LibreOffice (`soffice`) para renderizar. Rodar sempre com `python3 -I` (06, seção 8: o arquivo do cliente é dado, nunca código).

## Ordem no fluxo de um trabalho (06, seção 4)

| Passo do 06 | Comando | Bloqueia a entrega? |
|---|---|---|
| 4. Proteger trechos | `pipeline/proteger_trechos.py original.docx trechos.json [--proteger-apos "APÊNDICE B"] [--todas-aspas] [--comecar-em "1 INTRODUÇÃO"] [--parar-em "REFERÊNCIAS"]` | — |
| 5. Execução | `python3 -I entrada/spec_<cliente>.py saida/<cliente>_MARCADO.docx [paginas.json]` — spec do trabalho (fica em `entrada/`) sobre o motor `pipeline/montar_docx.py` | Falha se algum parágrafo do original ficar sem destino |
| 6. Sumário | `montar_docx.paginas_dos_titulos(limpa.pdf, m.titulos, 1, página_da_introdução)` → json → rodar a spec de novo até as páginas não mudarem | — |
| 6. Versão limpa | `pipeline/aceitar_alteracoes.py marcada.docx limpa.docx` | — |
| 6. Renderizar | `pipeline/renderizar.py limpa.docx saida/` | — |
| 7. Testes | `testes/verificar_limpa.py limpa.docx` | Sim (E14) |
| | `testes/verificar_marcada.py marcada.docx` | Sim (E14, 3.11) |
| | `testes/verificar_hashes.py trechos.json limpa.docx` | Sim (E1) |
| | `testes/verificar_paginacao.py limpa.pdf [--inicio N]` | Sim (E2) |
| | `testes/cruzar_citacoes.py limpa.docx --edicao 2002\|2023` | Não: lista para conferir à mão |
| Portfólio | `pipeline/montar_portfolio.py entrada.pdf termos.txt amostra.pdf --marca AMOSTRA` | Sim, se sobrar termo |

## O que cada arquivo faz

- `pipeline/docx_seguro.py` — base comum: recusa `.docm`/macro, limite contra zip bomb, XML sem entidades externas (XXE), texto na visão "final" ou "original".
- `pipeline/proteger_trechos.py` — hash de citações longas (recuo ≥ 4 cm ou estilo de citação), citações curtas com chamada e página, tabelas e tudo depois de um marcador. `--todas-aspas` protege todo trecho entre aspas (corpus citado em pedaços curtos, sem página); `--comecar-em`/`--parar-em` evitam falso positivo de folha de rosto e sumário digitado com recuo grande.
- `pipeline/montar_docx.py` — motor do build. Monta a versão marcada do zero: A4, margens, estilos ABNT, três seções com cabeçalho explícito e PAGE montado do zero, sumário automático com páginas em cache, quadro/tabela reconstruídos, imagem da capa copiada. Cada correção vira `w:del`/`w:ins` só no trecho que muda, com autor "Revisão ABNT"; comentários exigem prefixo de 3.11; parágrafo novo entra como inserção; `juntar_com_proximo=True` une dois parágrafos do original como alteração rastreada (marca de parágrafo excluída + espaço inserido; o `aceitar_alteracoes.py` une na limpa); `excluir=True` exclui um parágrafo inteiro do original como alteração rastreada (texto e marca de parágrafo excluídos; não usar no último parágrafo do documento); a cobertura impede perder texto do original. A spec de cada trabalho (ordem dos blocos, correções, comentários) fica em `entrada/` e não vai para o repositório — e some com o ambiente da sessão. O molde de spec com os comentários-padrão das regras do Manual IESB é `pipeline/modelos/spec_tcc_iesb.py`: copiar para `entrada/spec_<cliente>.py` e preencher. Exemplo completo de spec (documento fictício, 94 → 53 comentários, ciclo do sumário): `pipeline/modelos/spec_ficticio_ifood05.py` e `loop_sumario.py`; método e resultado do teste IFOOD-05 em `testes/ifood05/README.md`. Para achar os índices dos parágrafos do original: `python3 -I -c "import sys; sys.path.insert(0,'pipeline'); from docx_seguro import *; [print(i, texto_paragrafo(e,'original')[:70]) for i,e in enumerate(corpo(abrir('entrada/ORIGINAL_x.docx')))]"`.
- `pipeline/aceitar_alteracoes.py` — aceita por XML: inserções, exclusões, marca de parágrafo excluída, linha de tabela excluída, registros de formatação; remove comentários e desliga `trackRevisions`.
- `pipeline/renderizar.py` — `.docx` → PDF via LibreOffice com perfil isolado.
- `pipeline/montar_portfolio.py` — redação real (texto apagado), `scrub` de metadados, conferência por busca com variantes sem acento.
- `testes/verificar_*.py` e `testes/cruzar_citacoes.py` — testes do passo 7 do 06.
- `testes/autoteste/` — gera documentos fictícios e confere que cada teste passa e falha onde deve (41 verificações em nov/2026, incluindo `caso_montar.py`, que exercita o motor). Rodar a cada mudança de script; todo erro novo do 06 que der para automatizar vira um caso aqui.

## Limites conhecidos

- `proteger_trechos` não detecta citação sem recuo e sem chamada com página: conferir à mão e usar `--proteger-apos` para transcrições.
- `cruzar_citacoes` casa por primeiro sobrenome + ano: falso positivo com sobrenome composto e entrada por título. Saída é triagem, nunca referência completada.
- `verificar_paginacao` testa o PDF do LibreOffice. O teste no Word ou no Google Docs continua obrigatório (06, seção 5).
- Nome de pessoa quebrado entre linhas no PDF não é achado como expressão inteira: listar cada parte do nome no arquivo de termos.

## Página de apresentação

- Fonte: `pagina/fonte.html` (a mesma publicada como artifact no Claude). `python3 -I pagina/gerar_site.py` gera `site/index.html`, que a Cloudflare publica em https://site.autoriaemforma.workers.dev a cada push (`wrangler.jsonc`: só a pasta `site/`).
- **Preços aparecem em dois lugares da página:** a seção "Preços" (HTML) e a calculadora de orçamento (objeto `TABELA` no script, logo abaixo do comentário "calculadora de orçamento"). Toda mudança no `02_Tabela_Precos_Prazos_Escopo.md` do projeto atualiza os dois.
- O número de WhatsApp no código é fictício (`+55 61 90000-0000`); o real fica só na versão publicada do site.
- A calculadora não guarda o que a pessoa digita: só calcula no navegador e monta a mensagem do WhatsApp. O contador de acessos registra apenas que houve o envio, com serviço, faculdade e faixa de páginas.

## Contador de acessos e painel

- **Painel:** https://site.autoriaemforma.workers.dev/painel (usuário qualquer + senha; nem a senha nem o hash ficam no repositório: o SHA-256 dela é o Secret `PAINEL_HASH` na Cloudflare).
- **Código:** `analytics/worker.js` (coleta em `POST /api/e` e painel) e `analytics/schema.sql` (tabela `eventos` no D1 `autoria-analytics`). O rastreador fica no último `<script>` de `pagina/fonte.html` e só roda no domínio `autoriaemforma.*`.
- **Eventos (1 por sessão cada):** `visita`, `viu_precos`, `calc_inicio` (digitou as páginas), `calc_enviar` (clicou em enviar o orçamento), `whats_direto` (link do WhatsApp), `copiar_numero`, `exemplo` (mexeu na amostra marcada/final).
- **Origem:** o parâmetro `?origem=` do link vence (`whats`, `status`, `indicacao-nome`…); sem ele, o domínio de quem mandou (Google, Instagram…) ou `direto`. O parâmetro some da barra de endereço depois de lido.
- **Seus acessos:** abrir o `/painel` marca o navegador como seu; em outro aparelho, abra o site uma vez com `?eu=1` (`?eu=0` desfaz).
- **Privacidade:** sem cookie, sem IP, sem user agent, sem identificador de pessoa. Eventos com mais de 400 dias são apagados.
- **Teste local:** `npx wrangler d1 execute autoria-analytics --local --file analytics/schema.sql` e `npx wrangler dev`.

## Licença

[MIT](LICENSE). Os textos das normas ABNT (NBR 14724, 6023, 10520 etc.) e dos manuais das faculdades **não** estão incluídos: o código e os comentários só citam o número da norma e o item.
