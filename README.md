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
| `source/260828__Acompanhamento_das_frentes_v1.pptx` | Deck já no template Bain (base da v2) |
| `source/memo_lead_to_sales_v2.pdf` | Memorando estratégico (fonte dos entregáveis) |
| `scripts/extract.js` | Extrai o array `FRENTES` do HTML para `frentes.json` |
| `scripts/frentes.json` | Dados das 6 frentes (objetivo, líderes, time, milestones) |
| `scripts/overrides.json` | Correções do memo + trechos de *selective bold* |
| `scripts/deckstyle.py` | Tokens visuais, métricas de texto e helpers de tabela |
| `scripts/build.py` | Monta o `.pptx` (16:9) do zero, com a identidade visual do painel |
| `scripts/content_update.py` | Conteúdo do acompanhamento mensal (frentes 2, 4 e 5) |
| `scripts/update_deck.py` | Atualiza o deck já no template Bain, preservando o branding |
| `output/frentes_milestones.pptx` | Deck gerado do zero |
| `output/260829__Acompanhamento_das_frentes_v2.pptx` | **Entregável atual** — deck Bain atualizado |
| `output/pmo_status_frentes.html` | Painel de status das 4 frentes (simulador de pré-análise, distribuição de fichas, ecossistema de financiadoras, ciclo de crédito) |

## Como regerar

```bash
pip install python-pptx Pillow
cd scripts
node extract.js          # atualiza frentes.json a partir do HTML
python3 build.py         # gera o deck do zero
python3 update_deck.py   # atualiza o deck ja no template Bain
```

`update_deck.py` edita o arquivo em `source/` no lugar de recriá-lo, então
capa, página de objetivos, cronograma do PI Planning, cabeçalhos, rodapés e
todo o branding Bain ficam intactos.

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


## Acompanhamento de agosto/26 (v2)

**Frente 2 dividida em duas páginas**, seguindo as duas subfrentes do próprio
cronograma do PI Planning:

- *Qualificação de leads* — a coleta de informações mínimas virou três
  milestones separados por canal, todos com a Thais Pimenta: site (2.1.1),
  Facebook e WhatsApp (2.1.2) e classificados (2.1.3).
- *Alocação de leads* — as duas trilhas do cronograma: modelo simplificado
  v1.0 (2.4.1 a 2.4.3, com as 100 lojas em 11/09) e modelo robusto
  (2.4.4 a 2.4.6, dos ~5 pilotos até a escala do modelo vencedor), mais o
  piloto na Liza (2.4.7) e a clusterização (2.4.8).

**Frentes 4 e 5** mantiveram os milestones; mudaram status, progresso
recente, próximos passos e pontos a escalar. Dois status subiram de *Não
iniciado* para *Em progresso* (4.2.1 e 5.2.1) e o piloto da Central em lojas
do varejo (4.5.1) voltou de *Concluído* para *Em progresso*, com prazo
S2 set/26 — o piloto só começa em 08/09.

**Correções de formatação**: o menu de frentes estava quebrando em duas
linhas e cobrindo o título "Milestones" (a fonte tinha subido para 10,5pt);
foi refeito em 9pt, numa linha só, ocupando a largura útil. Todo o texto do
deck passou a ser marcado como `pt-BR`, para o PowerPoint tratar a
acentuação como português na edição.

---

# Treinamento — Simulador de Financiamento (PDV Seminovos)

Material de treinamento em HTML (17 slides, 16:9) para os **consultores de vendas
da Localiza Seminovos**, com o objetivo de fazê-los adotar o novo botão
**"Simule com os bancos"** antes de enviar a ficha ao banco.

O storyline segue a lógica *situação → complicação → resolução*, com linguagem
simples e foco no ganho do consultor (tempo, atendimentos e comissão), não no
ganho da empresa:

| Bloco | Slides | Mensagem |
|---|---|---|
| Capa e roteiro | 1–2 | "Pare de mandar ficha no escuro" |
| O problema | 3–4 | O ciclo de ficha negada e o que ele custa do bolso do consultor |
| O que mudou | 5–9 | O botão novo e os 3 passos, com os prints reais das telas |
| Como usar para vender | 10–12 | Vender a partir do crédito aprovado; antes × depois; ganhos |
| O que falar | 13–14 | 3 scripts de fala e 5 objeções com resposta pronta |
| Como começar | 15–17 | 3 regras de ouro, checklist, indicadores e fechamento |

Os slides 4 e 16 têm **espaços em branco (`__`) propositais** para as métricas da
operação (fichas por venda, tempo por tentativa, taxa de aprovação etc.), a serem
preenchidos com a base da loja/região antes da aplicação.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `source/nova_jornada_pdv_simulacao.pdf` | Documentação original da nova jornada (fonte das telas) |
| `scripts/treinamento_template.html` | Template do deck, com marcadores `{{TELA_*}}` |
| `scripts/build_treinamento.py` | Extrai as telas do PDF e embute em base64 |
| `output/treinamento_simulador_pdv.html` | **Entregável** — arquivo único, funciona offline |

## Como regerar

```bash
pip install pymupdf
python3 scripts/build_treinamento.py
```

O HTML gerado é autocontido (imagens em base64, sem fontes ou scripts externos),
então pode ser enviado por e-mail e aberto em qualquer navegador de loja, sem
internet. Navegação por setas ← →, clique nas laterais ou `Ctrl+P` para gerar PDF.

## Painel de status das 4 frentes

`output/pmo_status_frentes.html` é uma página autocontida (um único arquivo,
sem dependências externas) com uma tabela por frente e as colunas *O que
precisava ser feito*, *Prazo*, *Status*, *Progresso recente*, *Próximos passos*
e *Pontos a escalar*.

Convenção de status:

| Marcador | Significado |
|---|---|
| Check verde | Concluído |
| Três pontos verdes | Em andamento — dentro do plano |
| Três pontos vermelhos | Em andamento — fora do plano |
| Três pontos amarelos | Bloqueado — dependência externa, prazo a definir |

O marcador nunca aparece sozinho: cada linha traz o rótulo escrito ao lado do
ícone e uma barra de acento na primeira coluna. A folha de estilo de impressão
gera A4 paisagem (`Ctrl/Cmd+P` → salvar como PDF) para colar em slide.
