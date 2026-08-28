# Frentes e Milestones — Aceleração Seminovos

Gera um PPT de 6 páginas (uma por frente) a partir do painel HTML
`Lead-to-Sales — Painel de Controle`. Cada página traz apenas o **objetivo da
frente** e a **tabela de milestones**, replicando o layout do painel.

## Estrutura

| Caminho | Conteúdo |
|---|---|
| `source/dashboard_aceleracao_seminovos.html` | Painel HTML de origem |
| `scripts/extract.js` | Extrai o array `FRENTES` do HTML para `frentes.json` |
| `scripts/frentes.json` | Dados das 6 frentes (objetivo, líderes, time, milestones) |
| `scripts/build.py` | Monta o `.pptx` (16:9) com a identidade visual do painel |
| `output/frentes_milestones.pptx` | Deck gerado |

## Como regerar

```bash
pip install python-pptx Pillow
cd scripts
node extract.js          # atualiza frentes.json a partir do HTML
python3 build.py         # gera frentes_milestones.pptx
```

## Notas de layout

- Uma frente por página; a altura das linhas é calculada a partir do texto real
  (medido com métricas Arial/Liberation Sans) e o corpo reduz de 8pt até 5,5pt
  se necessário, garantindo que tudo caiba em uma única página.
- O espaço que sobra é distribuído entre as linhas (teto de 1,5") para o
  conteúdo preencher a página.
- Status seguem o mapeamento do painel: verde = Concluído, amarelo = Em
  progresso, vermelho = Não iniciado.
- As colunas *Progresso recente*, *Próximos passos* e *Pontos a escalar* vêm com
  caixas já formatadas e o texto `TBD`, prontas para preenchimento.
