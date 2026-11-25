# IFOOD-05 — A (à mão) × B (pipeline): registro de tempos e resultados

Documento fictício (29 págs.; IESB/Jornalismo; 48 erros plantados + 39 defeitos de concordância e 1 frase fora de lugar não planejados, do gerador). Só registro; **sem conclusão de "horas economizadas"**.

## 1. Tempos
Formato do cronômetro da B confirmado: min:s (as leituras 3:20:00 etc. são 3 min 20 s).

| | A (à mão) | B (pipeline) |
|---|---|---|
| Total (relógio) | 5:14:00 bruto; 4:56:45 líquido (2 pausas, 17:15) | 18:20 (T) |
| Tempo ativo do usuário | todo o tempo líquido (trabalho manual) | 3:45 (U) |

| Etapa A (ficha) | Tempo A (líquido) | Etapa B (leitura do cronômetro) | T (relógio) | U (ativo) |
|---|---|---|---|---|
| 1 Triagem e abertura | 18:40 | B1 Abertura e autoteste | 3:20 | 1:00 |
| 2 Proteção + Mapa Normativo | 17:35 | B2 Proteção + cruzar citações + Mapa | 1:20 | 0:45 |
| 3 Formatação | 47:20 | B3 Triagem + spec + montagem marcada/limpa + sumário (junta 3–5 da A) | 9:37 | 0:45 |
| 4 Citações/ilustrações/referências | 1:15:40 (retrabalho 8:15) | (dentro da B3) | | |
| 5 Revisão textual + marcada | 1:15:40 | (dentro da B3) | | |
| 6 Paginação e sumário | 19:45 | B6–B7 4 checks + auditoria visual (+ 1 ajuste e novo ciclo) | 1:23 | 0:40 |
| 7 Auditoria visual | 14:25 (retrabalho 3:10) | (dentro de B6–B7) | | |
| 8 Entrega | 27:40 | B8 Checklist 111 + relatório + entrega | 2:40 | 0:35 |
| **Total** | **4:56:45** | | **18:20** | **3:45** |

As fronteiras de etapa não coincidem (B3 junta as etapas 3 a 5 da A); só os totais são comparáveis.

## 2. Os 4 checks bloqueantes
| Check | A | B |
|---|---|---|
| verificar_limpa | OK (docx) | OK |
| verificar_marcada | N/A: não houve versão marcada na A | OK (autor "Revisão ABNT", prefixos) |
| verificar_hashes | FALHOU 5/8 (T001, T003, T004); inconclusivo: o .docx da A é conversão de PDF; no PDF, P1–P4 estão intactos | OK 8/8 (limpa e marcada) |
| verificar_paginacao | FALHOU: PDF da A sem número no canto superior direito (tem cabeçalho "Page x of 24" do navegador) | OK (32 págs.) |

## 3. Checklist de 111 itens
| | A | B |
|---|---|---|
| OK / PENDENTE / N/A | 87 / 6 / 18 | 79 / 11 / 21 |
Observações: a A marcou OK nos itens 78–84 (versão marcada) e 58/60 (tabela IBGE, legenda automática), apesar de não haver versão marcada; a B marcou 58 e 60 como PENDENTE (limite do motor) e teve o item 66 corrigido de OK para PENDENTE na conferência final. Nenhum dos dois testou no programa do cliente por mim: A declara ter conferido no Word; a B não testou no Word/Google Docs durante o teste (testada depois: seção 8).

## 4. Erros plantados (gabarito aberto só depois da B)
| Grupo | A | B |
|---|---|---|
| Texto T01–T12 (12) | 12 | 12 (T11 aparece como falso negativo do avaliador de PDF; confirmado no .docx) |
| Chamadas em caixa mista C01a–e, C09 | 6/6 | 6/6 |
| Citações × referências: C02 (COSTA), C03 (TAVARES), C04 (PEREIRA 2019), C07 (FERRAZ) | tratados por decisão da A: TAVARES removida, PEREIRA 2019→2018, FERRAZ com ano 2013 (do texto), COSTA sem entrada | todos como pendência ao cliente, sem alterar |
| C05 (pontuação da citação de Marques), C08 (ordem MARQUES/MENDES) | feitos | feitos (mas ver seção 8: na B, PEREIRA saiu depois de SILVA nas v1 a v3) |
| C06 (letras 2019a/b) | referências e chamadas | só referências; chamadas ficaram [PENDENTE] |
| Formatação F01–F24 | verificável em PDF: capa, rosto com local/ano, aprovação, F12/F13 **não** aplicados (alíneas em minúscula com ";" e palavras-chave com ";" intactas), F21 feito, F04/F05 não visíveis no PDF (docx inconclusivo) | tudo aplicado, exceto F21 (resumo em 2 parágrafos: motor não junta) e F24 (título de PEREIRA sem negrito: não vi) |
| Protegidos P1–P4 | intactos | intactos; comentários [PENDENTE]/[FATO] sobre "atrazo", "en crise", manchetes e 1831×1832 |
| Chamarizes D1–D3 | D1, D3 mantidos; D2 alterado p/ "et al." | D1, D3 mantidos; D2: **alterei uma** chamada de 4 autores para "et al." e deixei 16 iguais (inconsistência minha) |
| X1 extensão (<40 págs.) | não localizei no material enviado | [risco alto] no relatório |

Defeitos não planejados do gerador (contados à parte): 39 concordâncias de gênero — A e B corrigiram (0 residuais reais nos dois). Frase fora de lugar antes da citação de Araújo: A reverteu a frase; B deixou como está e marcou [DECIDIR].

## 5. O que deu errado ou foi repetido
- B1: autoteste 28/29 por `pymupdf` ausente → instalei e repeti (29/29).
- B3: o build da spec falhou 2 vezes (substantivo "imprensa local" fora da lista; "Concerteza" com maiúscula).
- B6–B7: local e ano da capa e do rosto fora do pé → ajuste e novo ciclo completo (sumário + 4 checks).
- B8: meu primeiro build comentava cada ocorrência de erro repetitivo (viola item 81) → corrigi e refiz o ciclo (94 → 53 comentários).
- Erros meus achados só na conferência final: item 66 marcado OK por engano (PEREIRA sem negrito); "et al." aplicado em 1 de 17 chamadas de 4 autores. Os arquivos entregues não foram alterados depois do cronômetro parar.
- A: retrabalho de 8:15 (SOUZA 2019a/b) e 3:10 (espaçamento após fonte da tabela/figura).

## 6. Limitações
- n=1; documento fictício; o usuário conhece o processo; a A veio primeiro (aprendizado favorece a B).
- Eu escrevi o gerador e vi o gabarito: a B não é cega. A spec foi escrita por mim lendo o documento, mas com esse viés.
- Fronteiras de etapa diferentes; U da A = todo o tempo líquido, U da B = só leitura e decisão do usuário; o T da B inclui meu processamento e o tempo de resposta do usuário.
- B não rodou no Word/Google Docs durante o teste; PDF só pelo LibreOffice (conferido depois no Word e no Google Docs: seção 8). A: .docx entregue é conversão de PDF (checks de .docx inconclusivos) e não houve versão marcada.
- A não passou pelo meu processo de decisão: A tomou decisões que o projeto manda deixar como pendência (TAVARES, PEREIRA, FERRAZ).
- Preferências do projeto (Arial × Times) foram diferentes nas duas condições.
- O gerador foi para o repositório (PR #2); o gabarito (lista de erros plantados) ficou só na sessão, de propósito, para não contaminar execuções futuras.

## 7. Atualização após a conferência (10/11/2026, fora do tempo medido)
- B v2 entregue: PEREIRA com título em negrito; troca isolada para "et al." desfeita (as 17 chamadas de 4 autores ficam iguais, com todos os autores); 4 checks 4/4; checklist 80 OK / 10 PENDENTE / 21 N/A. A v1 cronometrada continua sendo a que vale para o tempo.
- Gerador, spec de exemplo e método no repositório privado (PR #2 de `autoria-em-forma-pipeline`).

## 8. Lições para o 06 (a incorporar na seção 2, "Lições do segundo trabalho" → nova subseção "Teste IFOOD-05")
- **L6 — Erro repetitivo comenta uma vez, também na spec.** O primeiro build comentava cada concordância e cada chamada em caixa mista (94 comentários). Na spec, passar o comentário só na 1ª ocorrência de cada tipo (53 comentários). Item 81 do checklist.
- **L7 — Conferir o destaque das referências obra a obra.** Uma referência sem negrito no título passa despercebida se o motor copia a formatação do original. Checar por script que todas as entradas têm o mesmo destaque antes de marcar o item 66.
- **L8 — Chamada com 4 ou mais autores: tratar todas ou nenhuma.** Trocar uma só por "et al." deixa o texto inconsistente. Escolher uma forma (todos os autores é aceitável; et al. também) e aplicar a todas as ocorrências, com [DECIDIR] se a edição da norma não resolver.
- **L9 — Citação curta que passa de 3 linhas no PDF vira bloco recuado.** O detector de trechos protegidos a classificou como curta; só o PDF renderizado mostra as linhas. Medir no PDF depois do 1º render (o motor não decide isso).
- **L10 — Limites do motor.** Não junta dois parágrafos (resumo em 2 parágrafos vira [DECIDIR]); tabela sai com bordas fechadas (não IBGE) e legendas são digitadas (sem campo SEQ do Word). Itens 58 e 60 do checklist ficam PENDENTE até o motor tratar.
- **L11 — Avaliador de PDF tem falso negativo em frase partida entre páginas.** Confirmar no .docx antes de contar como erro.

## 8. Achados após a corrida (10/11/2026, verificados nos arquivos entregues; não entram na tabela de tempos nem nos checklists acima)
1. **A (à mão), numeração de página.** O .docx feito à mão passa nos checks de hashes (8/8) e de limpeza, mas não tem numeração de página: uma seção só, sem cabeçalho e sem campo PAGE. Por isso a verificação de paginação falha. O arquivo foi mantido intacto, como evidência medida. (A seção 2 registra hashes 5/8, inconclusivo, para o .docx que era conversão de PDF; o 8/8 aqui vem da conferência posterior no .docx feito à mão. Confirmar se são o mesmo arquivo antes de citar os dois números juntos.)
2. **B, referências fora de ordem.** Nas versões 1 a 3 a lista de referências saiu com PEREIRA depois de SILVA. O original já estava em ordem; o erro foi da minha spec de montagem, não do documento de origem. O item 65 do checklist foi marcado OK por engano. Corrigido na v4, fora do tempo medido. Isso se soma aos dois erros da seção 5 (PEREIRA sem negrito, "et al." em 1 de 17 chamadas).
3. **B, Word e Google Docs.** Aberta depois no Word e no Google Docs, a v3 mostrou as mesmas páginas do PDF até a Introdução (p. 6) e uma página a mais em REFERÊNCIAS (32 contra 31). A causa não foi identificada. O cliente precisa atualizar o sumário (Word: F9; Google Docs: ícone de atualizar). Os números gravados no sumário vêm do render do LibreOffice.
4. **Entrega final do teste.** `B_LIMPA_v4.docx`, `B_MARCADA_v4.docx` e `B_LIMPA_v4.pdf`: hashes 8/8 nas duas versões, 32 páginas, limpa e marcada OK. A v1 continua sendo a que vale para o tempo.
5. **Motor.** Para o resumo em 2 parágrafos e para a obra removida de SOUZA o motor ganhou `juntar_com_proximo=True` (PR #3) e `excluir=True` (PR #4). Autoteste: 41/41. A spec de exemplo (`pipeline/modelos/spec_ficticio_ifood05.py`) foi sincronizada com as correções de PEREIRA e SOUZA no PR #5.
6. **O que continua valendo.** Nenhuma conclusão de "horas economizadas". Os achados 1 a 3 pesam contra a leitura favorável do resultado da B e a favor de ler o teste como registro do que falhou em cada condição.
