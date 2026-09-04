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
| `output/pmo_status_frentes.html` | Página única (16:9) com o status das 4 frentes de F&I |
| `output/260903__Status_frentes_3paginas.pptx` | **Entregável atual** — deck editável de 3 páginas com o status das 4 frentes |

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
| `scripts/extract_treinamento_layout.js` | Mede no Chromium o layout real de cada página do HTML |
| `scripts/build_treinamento_ppt.py` | Monta o `.pptx` a partir do layout medido |
| `output/treinamento_simulador_pdv.html` | **Entregável** — arquivo único, funciona offline |
| `output/treinamento_simulador_pdv.pptx` | **Entregável** — mesmas 17 páginas em PowerPoint editável |

## Como regerar

```bash
pip install pymupdf python-pptx Pillow
python3 scripts/build_treinamento.py          # HTML autocontido

cd scripts                                    # versao PowerPoint
node extract_treinamento_layout.js ../output/treinamento_simulador_pdv.html \
     layout.json telas/
python3 build_treinamento_ppt.py layout.json telas/ \
     ../output/treinamento_simulador_pdv.pptx
```

O HTML gerado é autocontido (imagens em base64, sem fontes ou scripts externos),
então pode ser enviado por e-mail e aberto em qualquer navegador de loja, sem
internet. Navegação por setas ← →, clique nas laterais ou `Ctrl+P` para gerar PDF.

## Versão PowerPoint

As 17 páginas do HTML viraram 17 slides de 1280 × 720 px (16:9 exato, 1 px CSS =
9525 EMU), **sem imagem de fundo**: cada bloco do HTML é uma forma nativa e cada
texto é uma caixa de texto do PowerPoint, com as mesmas coordenadas, cores,
corpos de fonte, entrelinhas e *letter-spacing* do original. A conversão é
medida, não estimada:

- `extract_treinamento_layout.js` abre o HTML no Chromium, torna as 17 páginas
  visíveis, e grava a caixa, o estilo computado e os trechos de texto de cada
  elemento — inclusive os pseudo-elementos (`::before` da régua verde e das
  bolinhas da lista, `::after` da seta do ciclo e do número da página).
- `build_treinamento_ppt.py` traduz esse *display list* para o PPT: fundos e
  bordas assimétricas viram retângulos, a seta do ciclo vira forma livre, a
  tabela "antes × depois" vira tabela nativa (editável célula a célula), as
  telas do produto entram como imagem com `object-fit: contain`, e a capa e o
  fechamento recebem o gradiente linear do CSS.
- **Correção de linha de base**: o CSS distribui meia-entrelinha acima da
  primeira linha, enquanto o PowerPoint mede a primeira linha de base como
  `topo + entrelinha − descendente`. O script compensa a diferença a partir das
  métricas da fonte (`CONTEUDO`), o que zera o desvio vertical dos blocos.
- Cada forma leva o nome da classe do HTML (`card · fundo`, `titulo · texto`),
  para achar o bloco no painel de seleção do PowerPoint.

O deck usa **Segoe UI**, a mesma família declarada no HTML. Como o container de
build não tem essa fonte, a medição roda em Liberation Sans e as métricas da
Segoe UI entram por constante; comparando página a página com a mesma fonte nas
duas pontas, a diferença estrutural contra o HTML fica em 0,4%.

## Página de status das 4 frentes de F&I

`output/pmo_status_frentes.html` é uma página autocontida (arquivo único, sem
dependências externas) desenhada como **um slide 16:9**: um artboard de
1600×900 que se ajusta à janela por `transform: scale()` e imprime em
exatamente uma folha A4 paisagem.

Uma linha por entregável, agrupadas por frente, com as colunas *O que
precisava ser feito* (com o prazo), *Status*, *Progresso recente*, *Próximos
passos* e *Pontos a escalar*.

Convenção de status — a célula recebe fundo tingido, ícone e rótulo escrito,
como no deck de acompanhamento:

| Marcador | Significado |
|---|---|
| Check verde | Concluído |
| Três pontos verdes | Em andamento — dentro do plano |
| Três pontos vermelhos | Em andamento — fora do plano |
| Três pontos amarelos | Bloqueado em terceiros — prazo a definir |

As quatro frentes usam o mesmo *color code* do plano de escalada de F&I
(azul, vinho, ocre e cinza), aplicado na barra e no número à esquerda de cada
grupo. O ocre foi escurecido um passo para o numeral branco atingir contraste
3:1.

---

# Status das quatro frentes (página de PMO)

Uma página de acompanhamento com as quatro frentes de crédito e, para cada
marco, **o que precisava ser feito**, o **prazo**, o **status** (semáforo),
**progresso recente**, **próximos passos** e **pontos a escalar**. Mesma
linguagem visual do deck `Acompanhamento das frentes`: barra verde por frente,
tabela nativa do PowerPoint e corpo que reduz de 9,5pt até 7,5pt para caber em
uma página.

## Semáforo

| Símbolo | Chave | Significado |
|---|---|---|
| ✓ verde | `ok` | Concluído |
| ••• verde | `plano` | Em andamento, dentro do plano |
| ••• amarelo | `risco` | Em andamento, com riscos |
| ••• vermelho | `atraso` | Em andamento, fora do plano |

O check e os pontos são caracteres comuns (`✓` e `•`), não Wingdings, para
renderizar igual em qualquer máquina. A faixa de cada frente traz a contagem
por status à direita, e o título da página é montado a partir dessas contagens,
então os números nunca ficam defasados do conteúdo.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `scripts/pmo_content.py` | **Conteúdo da página** — frentes, marcos, prazos, status e comentários |
| `scripts/pmo_status.py` | Monta a página (16:9), calcula alturas e desenha o semáforo |
| `output/260903__Status_quatro_frentes.pptx` | **Entregável** — página única |

## Como regerar

```bash
pip install python-pptx Pillow
cd scripts
python3 pmo_status.py
```

Para atualizar o acompanhamento basta editar `pmo_content.py`: cada marco é uma
tupla `(o que precisava ser feito, prazo, status, progresso recente, próximos
passos, pontos a escalar)`, e `NA`/`TBD` aparecem em cinza claro como
placeholder. Trechos a destacar em negrito ficam no dicionário `BOLD`.

## Notas de layout

- **Uma tabela nativa só**, editável célula a célula; a faixa de cada frente usa
  células mescladas (título à esquerda, contagem à direita).
- As alturas das linhas são calculadas a partir do texto real (métricas
  Arial/Liberation Sans) com uma folga sobre a altura de linha, porque o
  PowerPoint trata a altura da linha como mínimo e infla a tabela se a
  estimativa for curta.
- Células sem texto têm o parágrafo vazio fixado em 1pt e margens zeradas: um
  parágrafo vazio herda 18pt e forçaria ~0,4" de altura mínima na linha.

---

# Deck de 3 páginas do status das frentes (`pmo_deck.py`)

Junta as duas versões da página de status em um único `.pptx` editável, mais
uma terceira que combina as duas. Tudo em tabelas nativas do PowerPoint, então
dá para editar célula a célula, inserir e remover linhas.

| Página | O que é |
|---|---|
| 1 | Versão da página HTML: identidade por frente na faixa da esquerda, com o *color code* do plano de escalada, responsável, valor gerado e status em pastilha com rótulo escrito |
| 2 | A página do print, gerada por `pmo_status.py` sem alteração |
| 3 | Combinação das duas: color code e identidade da frente da 1; título construído a partir das contagens e contagem por frente da 2; a coluna *Pontos a escalar* sai da tabela e vira a faixa **Decisões pedidas nesta reunião** no pé da página |

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `scripts/pmo_deck_content.py` | **Conteúdo das páginas 1 e 3** — frentes, responsável, valor, marcos e decisões pedidas |
| `scripts/pmo_deck.py` | Monta as três páginas e chama `pmo_status.build()` para a página 2 |
| `scripts/pptx_preview.py` | Preview HTML de um `.pptx`, lido da geometria real do arquivo |

## Como regerar

```bash
pip install python-pptx Pillow
cd scripts
python3 pmo_deck.py

# conferir o layout sem abrir o PowerPoint
python3 pptx_preview.py ../output/260903__Status_frentes_3paginas.pptx /tmp/preview.html
```

## Notas de layout

- **A barra de cor da frente é a primeira coluna da tabela**, mesclada
  verticalmente sobre as linhas do grupo — não é um shape solto. Assim continua
  colada na linha quando o texto for editado no PowerPoint.
- A altura de cada linha é calculada a partir do texto real (métricas
  Arial/Liberation Sans) considerando a coluna mais alta, incluindo a pastilha
  de status e a linha do prazo. O corpo reduz de 9,5pt até 7,5pt para caber em
  uma página; se nem no corpo mínimo couber, o build falha com a medida em vez
  de gerar uma página com sobreposição.
- A faixa de identidade da frente encolhe junto com o corpo (`band_sizes`), e a
  altura do grupo nunca fica menor que o conteúdo da faixa.
- O símbolo do semáforo fica no mesmo parágrafo do rótulo, com tamanho próprio
  de run, para o check e os três pontos ficarem na linha do texto.
- Símbolos são caracteres comuns (`✓` e `•`), não Wingdings.

---

# Página — Piloto de alocação de leads em loja próxima

Recriação da página de acompanhamento do piloto de alocação (pré vs. pós
piloto, por grupo), com o layout original preservado e duas mudanças:

1. **Wordings migrados** de *loja escolhida / loja alocada* para **loja próxima**:
   as categorias passaram a ser `Alocado em loja próxima` e
   `NÃO alocado em loja próxima`, os cabeçalhos viraram
   *"% conversão – leads com loja próxima identificada"* e a linha de base dos
   gráficos virou `% dos leads` (antes `% total`), acompanhando a base dos
   painéis novos.
2. **Gráficos substituídos** pelos painéis atuais (pré-piloto de 103 dias,
   01/04 a 13/07; pós-piloto oficial de 25 dias, 13/07 a 06/08).

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `scripts/build_piloto_alocacao.py` | Monta a página (16:9) do zero, com os dados no topo do arquivo |
| `output/piloto_alocacao_loja_proxima.pptx` | **Entregável** — a página em PPT |

```bash
pip install python-pptx
python3 scripts/build_piloto_alocacao.py
```

## Layout preservado

- Cabeçalhos em chevron (pré em preto, pós em cinza), régua vermelha sob o
  título e tag `/PRELIMINAR`.
- Três painéis tracejados — *Lojas do piloto* (vermelho), *Grupo controle* e
  *Restante do Brasil* (cinza) — com o rótulo do grupo rotacionado à esquerda.
- Barras marimekko (largura proporcional ao `% dos leads`), linha tracejada
  vermelha da média, delta em p.p. entre as barras e a variação pré → pós
  entre os dois gráficos de cada linha.
- Caixa de destaque vermelha na linha do piloto e nota/fonte no rodapé.

## Dados (razão das somas: total de vendas / total de leads)

| Grupo | Pré: leads / vendas / média | Pós: leads / vendas / média | Variação |
|---|---|---|---|
| Lojas do piloto | 7.596 / 44 / 0,58% | 1.626 / 4 / 0,25% | **-~58%** |
| Grupo controle | 74.255 / 380 / 0,51% | 15.213 / 85 / 0,56% | **+~9%** |
| Restante do Brasil | 411.317 / 1.884 / 0,46% | 84.727 / 478 / 0,56% | **+~23%** |

## Ajustes de conteúdo

- **Título e conclusão reescritos.** Os dados novos invertem a mensagem: a
  conversão das lojas do piloto caiu, enquanto controle e restante do Brasil
  subiram. O título antigo (*aumento de ~20% vs. ~10%*) foi substituído pela
  leitura atual, com a ressalva da base pequena do pós piloto.
- **Caixa de conclusão do rodapé removida** (a escala para +70 lojas em 01/09
  não é mais o fecho da página).
- **Destaque da linha do piloto** reescrito: a conversão dos leads alocados em
  loja próxima caiu de 0,55% para 0,18% (3 vendas em 1.623 leads).
- **Sinal do delta na linha *Restante do Brasil*** corrigido para `+0,05 p.p.`
  (pré) e `+0,06 p.p.` (pós). Os painéis de origem imprimem esses dois valores
  com sinal negativo, mas a barra *Alocado em loja próxima* está **acima** da
  *NÃO alocado* nas duas janelas (0,46% vs. 0,41% e 0,57% vs. 0,51%) — as
  magnitudes batem, só o sinal estava invertido. Nas linhas 1 e 2 o sinal dos
  painéis já era coerente e foi mantido.
- **Barra fora de escala.** No pós piloto das lojas do piloto, os 33,33% da
  categoria *NÃO alocado* vêm de 1 venda em 3 leads; a barra é marcada com
  seta vermelha em vez de estourar o eixo.
