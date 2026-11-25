# IFOOD-05 — teste cronometrado, à mão × pipeline (10/11/2026)

Teste controlado, **n=1**, com documento **fictício** (nunca TCC de cliente). Aqui ficam só o gerador e o método; os
arquivos de entrega e o gabarito não vão para o repositório. Não é medida de "horas economizadas": só registro de tempos.

## Reproduzir o documento
`python3 -I testes/ifood05/gerar_fixture_ifood05.py --saida DIR --seed 5 --escala 0.6` grava `DIR/ORIGINAL.docx`
(29 páginas, IESB/Jornalismo) e `DIR/.privado/gabarito.json`. O gerador planta 48 erros: 24 de formatação (F),
15 de citação/referência (C), 12 de texto (T), 4 trechos que não podem ser alterados (P), 3 chamarizes (D) e 1 risco (X).
**Não abrir o gabarito antes de terminar a condição A.**

Defeitos **não planejados** do gerador (ficaram no arquivo do teste para manter as cópias idênticas): 39 frases com
sujeito masculino e predicado feminino ("o custo ... era marcada") e uma frase solta antes da citação de Araújo.

## Protocolo
Mesmo arquivo nas duas condições; critério de pronto igual: os 4 checks bloqueantes na saída + checklist de 111 itens.
A: à mão, cronômetro do usuário, sem ajuda. B: pipeline; cronômetro do usuário lido a cada etapa (T = relógio total,
U = tempo ativo de quem decide). Spec de exemplo: `pipeline/modelos/spec_ficticio_ifood05.py` + `loop_sumario.py`.

## Resultado
| | A (à mão) | B (pipeline) |
|---|---|---|
| Relógio total | 5:14:00 (líquido 4:56:45) | 18:20 |
| Tempo ativo do usuário | todo o líquido | 3:45 |
| 4 checks | limpa OK; marcada n/a; hashes 5/8 (inconclusivo: .docx é conversão de PDF); paginação falhou | 4/4 OK |
| Checklist de 111 | 87 OK / 6 PENDENTE / 18 N/A | 79 OK / 11 PENDENTE / 21 N/A (v1) |

Na B, a conferência final contra o gabarito achou 2 erros da própria B (título de PEREIRA sem negrito; "et al." em uma
só de 17 chamadas com 4 autores), corrigidos depois do cronômetro (v2). O resumo em 2 parágrafos ficou
como [DECIDIR]: o motor não junta parágrafos.

## Limitações
n=1; documento fictício; quem cronometrou conhece o processo; a A veio primeiro (favorece a B); a B foi conduzida por
quem escreveu o gerador e viu o gabarito (não é cega); fronteiras de etapa diferentes (a B3 junta as etapas 3 a 5 da A);
o T da B inclui o processamento do modelo e o tempo de resposta do usuário; B testada só no LibreOffice; a A não teve
versão marcada e o .docx entregue é conversão de PDF.
