# Orientações para o Claude neste repo

Projeto de materiais da frente **Aceleração Seminovos / Lead-to-Sales**:
scripts Python que geram decks `.pptx` e páginas HTML a partir de dados
versionados. Veja o `README.md` para o mapa completo dos scripts.

## Documento de referência

Quando o usuário subir um documento (memo, deck do cliente, PDF de processo)
ou pedir para conferir algo contra o material oficial:

1. O arquivo vai em `source/referencia/`.
2. Rode `python3 scripts/ler_referencia.py` — converte PDF/PPTX em texto
   pesquisável em `source/referencia/_texto/` (pula o que já está atualizado).
3. Use o agente `referencia` (`.claude/agents/referencia.md`) para o trabalho
   que depende do documento. Ele trata o documento como fonte da verdade,
   cita página/slide em cada afirmação e aponta conflito com os dados do repo
   em vez de decidir sozinho.

Nunca preencha número, prazo ou nome que não esteja no documento — diga que
não está lá.

## Convenções

- Conteúdo (texto de milestone, prazo, destaques em negrito) vive em
  `scripts/overrides.json`, não dentro do `.py`.
- `scripts/frentes.json` é gerado por `node scripts/extract.js` a partir do
  painel HTML — não edite à mão se a origem for o painel.
- `update_deck.py` edita o deck em `source/` no lugar de recriá-lo, para o
  branding Bain ficar intacto. Não troque por "montar do zero".
- Taxas e percentuais são recalculados a partir dos volumes brutos, nunca
  transcritos de valores já arredondados.
- Tabelas nativas do PowerPoint em vez de caixas de texto, tokens visuais de
  `scripts/deckstyle.py`, docstring em português (sem acento) no topo de cada
  script explicando a página que ele monta.
- Windows: os comandos usam `py` no lugar de `python3`.
