# Frentes e Milestones — Aceleração Seminovos

Gera um PPT de 6 páginas (uma por frente) a partir do painel HTML
`Lead-to-Sales — Painel de Controle`, com os prazos e entregáveis conferidos
contra o memorando estratégico Lead to Sales (v2, 17/ago). Cada página traz
apenas o **objetivo da frente** e a **tabela de milestones**, replicando o
layout do painel.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `source/dashboard_aceleracao_seminovos.html` | Painel HTML de origem |
| `source/memo_lead_to_sales_v2.pdf` | Memorando estratégico (fonte dos entregáveis) |
| `scripts/extract.js` | Extrai o array `FRENTES` do HTML para `frentes.json` |
| `scripts/frentes.json` | Dados das 6 frentes (objetivo, líderes, time, milestones) |
| `scripts/overrides.json` | Correções do memo + trechos de *selective bold* |
| `scripts/build.py` | Monta o `.pptx` (16:9) com a identidade visual do painel |
| `output/frentes_milestones.pptx` | Deck gerado |

## Como regerar

```bash
pip install python-pptx Pillow
cd scripts
node extract.js          # atualiza frentes.json a partir do HTML
python3 build.py         # gera frentes_milestones.pptx
```

Para editar conteúdo sem mexer no código, altere `overrides.json`: cada
milestone aceita `desc`, `prazo` e `bold` (lista de trechos a destacar); cada
frente aceita `obj_bold` e `nota`.

## Notas de layout

- **Tabelas nativas.** O cabeçalho da frente e a grade de milestones são
  tabelas do PowerPoint, não caixas de texto — dá para editar célula a célula,
  inserir e remover linhas.
- **Menu de frentes** entre o cabeçalho e os milestones, com hiperlinks
  internos: clicar em qualquer frente pula direto para a página dela. A frente
  atual aparece em verde.
- **Selective bold** no objetivo e nos milestones, seguindo os destaques do
  próprio memorando.
- Uma frente por página; a altura das linhas é calculada a partir do texto real
  (medido com métricas Arial/Liberation Sans) e o corpo reduz de 10pt até 7,5pt
  se necessário, garantindo que tudo caiba em uma única página. O espaço que
  sobra é distribuído entre as linhas (teto de 1,1").
- Status seguem o mapeamento do painel: verde = Concluído, amarelo = Em
  progresso, vermelho = Não iniciado.
- As colunas *Progresso recente*, *Próximos passos* e *Pontos a escalar* vêm com
  células já formatadas e o texto `TBD`, prontas para preenchimento.

## Ajustes vindos do memorando

| Milestone | Ajuste |
|---|---|
| 1.1.1 | Texto alinhado ao memo (*definição da necessidade de vendas*) |
| 1.4.1 / 1.5.1 | Prazo S4 set/26 → **S2 out/26 (TBC)** |
| 1.6.1 | Prazo marcado como **(TBC)** |
| 2.2.1 | Prazo S1 set/26 → **S4 set/26** |
| 2.4.3 | Prazo S2 out/26 → **S2 nov/26** |
| 3.5.1 | Prazo set/26 – out/26 → **set/26 – nov/26** |
| Frente 3 | Nota do ciclo completo (~8 meses, antecipável para ~3) |
| 4.2.1 | Prazo S2 set/26 → **S3 set/26** |
| 4.3.1 | Detalhe das classes (13 = nível de EVs, 10 = nível de CVs) |
| 6.8.1 | Inclui *páginas de enriquecimento* |
| 6.10.1 | Inclui a condição de escalonamento por volume, conversão e custo por venda |
