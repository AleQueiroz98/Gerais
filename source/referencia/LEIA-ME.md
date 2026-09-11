# Documentos de referência

Solte aqui o documento que deve servir de **fonte da verdade** — memo
estratégico, deck do cliente, PDF de processo, ata. Formatos lidos
automaticamente: `.pdf`, `.pptx`, `.txt`, `.md`, `.csv`, `.html`.

Depois rode, na raiz do projeto:

```bash
python3 scripts/ler_referencia.py    # no Windows: py scripts\ler_referencia.py
```

O script gera um `.md` pesquisável por documento em `_texto/`, com marcação de
página (PDF) ou slide (PPTX), tabelas e notas do apresentador. É esse texto que
o agente `referencia` consulta — ele fica versionado no repo, então qualquer
sessão futura consegue ler o documento sem precisar abrir o arquivo original.

Rodar de novo é barato: documentos já extraídos aparecem como `[ = ]` e são
pulados. Use `--forcar` para refazer tudo e `--listar` para só ver o que há na
pasta.

## Como pedir para o agente usar

Na conversa, chame o agente pelo nome ou simplesmente cite o documento:

> usa o agente referencia: confere os prazos da frente 4 contra o memo v2

> o memo novo está em source/referencia — atualiza os milestones do deck de
> acompanhamento com o que mudou

O agente sempre cita página ou slide de onde tirou cada número, e quando o
documento contradiz o que já está em `scripts/overrides.json` ou nos decks de
`output/`, ele mostra os dois lados antes de mudar qualquer coisa.

## Cuidados

- `.ppt` e `.doc` (formato antigo) não são lidos: abra no Office e salve como
  `.pptx` / `.docx`.
- Página de PDF que é imagem ou scan não tem texto extraível — o script avisa
  quais são. Nesse caso, mande um print da página na conversa.
- O que entra aqui vai para o Git. Não coloque material que não possa ser
  versionado no repositório.
