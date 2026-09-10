# Rodar os scripts Python no seu computador (Windows)

Guia do zero para o notebook corporativo, **sem precisar de direitos de
administrador**. Ao final você consegue regerar os decks deste repositório
com um comando.

Tempo estimado: 15–20 minutos, quase todo de download.

---

## Passo 0 — Ver se o Python já está instalado

Abra o **Terminal** ou o **PowerShell** (tecla Windows → digite `terminal` →
Enter) e rode:

```powershell
py --version
```

- Apareceu algo como `Python 3.12.4` → **pule para o Passo 2**.
- Apareceu `Python 3.8.x` ou anterior → instale uma versão nova (Passo 1).
- Deu erro (`py não é reconhecido...`) → siga o Passo 1.

> **Por que `py` e não `python`?** O Windows vem com um atalho falso chamado
> `python` que só abre a Microsoft Store. O comando `py` é o lançador oficial
> do Python e não tem esse problema. Use `py` em tudo neste guia.

---

## Passo 1 — Instalar o Python

Tente as opções na ordem. A primeira que funcionar já resolve.

### Opção A — Portal corporativo (preferida)

Procure por **Company Portal**, **Software Center** ou o catálogo de
softwares da Bain e instale o **Python 3.12** (ou a versão mais recente
disponível). É o caminho aprovado por TI, não pede senha de administrador e
já vem com as políticas de rede da empresa configuradas.

### Opção B — Microsoft Store

Tecla Windows → **Microsoft Store** → busque por `Python 3.12` → **Obter**.
Instala no perfil do usuário, sem admin, e configura o PATH sozinho.

### Opção C — winget (linha de comando)

No Terminal:

```powershell
winget install Python.Python.3.12 --scope user
```

O `--scope user` é o que evita o pedido de administrador.

### Opção D — Instalador oficial python.org

1. Baixe em <https://www.python.org/downloads/windows/> o
   *Windows installer (64-bit)* da versão 3.12.
2. Execute o arquivo e, **antes de clicar em qualquer coisa**:
   - marque **Add python.exe to PATH** (embaixo da janela);
   - escolha **Install Now** — ela instala só para o seu usuário.
3. Se aparecer o Controle de Conta de Usuário pedindo senha de admin, volte e
   use **Customize installation** → desmarque *Install for all users*.

### Conferir

Feche o Terminal, abra de novo (o PATH só atualiza em janela nova) e rode:

```powershell
py --version
```

Tem que responder `Python 3.12.x` ou similar.

---

## Passo 2 — Baixar este repositório

### Se você tem o Git instalado

```powershell
cd $HOME\Documents
git clone https://github.com/AleQueiroz98/Gerais.git
cd Gerais
```

### Se não tem o Git

Abra <https://github.com/AleQueiroz98/Gerais> → botão verde **Code** →
**Download ZIP** → extraia em `Documentos\Gerais`. Depois, no Terminal:

```powershell
cd $HOME\Documents\Gerais
```

> Todos os comandos deste guia assumem que você está **dentro da pasta
> `Gerais`**. Para conferir onde está, rode `pwd`.

---

## Passo 3 — Instalar as bibliotecas

```powershell
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

Isso instala quatro pacotes, todos no seu perfil de usuário:

| Pacote | Para que serve |
|---|---|
| `python-pptx` | monta e edita os arquivos `.pptx` |
| `Pillow` | mede o texto para calcular a altura das linhas |
| `lxml` | mexe no XML interno do PowerPoint |
| `pymupdf` | lê os PDFs de origem (só o material de treinamento do PDV) |

---

## Passo 4 — Testar o ambiente

```powershell
py scripts\check_ambiente.py
```

O script não instala nada — só verifica e diz o que falta. Quando estiver
tudo certo, a saída termina com **"Tudo pronto."**

---

## Passo 5 — Rodar de verdade

Os scripts usam caminhos relativos, então entre na pasta `scripts` antes:

```powershell
cd scripts
py pmo_deck.py
```

O arquivo gerado aparece em `output\`. Outros comandos úteis:

```powershell
py build.py                    # deck das 6 frentes, do zero
py update_deck.py              # atualiza o deck no template Bain
py pmo_status.py               # pagina HTML de status das 4 frentes
py build_piloto_alocacao.py    # pagina do piloto de alocacao
```

> O README usa `python3` nos exemplos (padrão de Mac/Linux). No Windows,
> troque por `py` — o resto do comando é idêntico.

---

## Problemas comuns

| O que aparece | O que fazer |
|---|---|
| `py não é reconhecido como comando` | Feche e reabra o Terminal. Se persistir, o PATH não foi configurado: reinstale marcando *Add python.exe to PATH* (Passo 1, Opção D). |
| Abre a Microsoft Store ao digitar `python` | É o atalho falso do Windows. Use `py`. Para desligar o atalho: Configurações → Aplicativos → *Aliases de execução de aplicativo* → desative `python.exe`. |
| `No module named pptx` | As bibliotecas não foram instaladas ou foram para outro Python. Rode de novo `py -m pip install -r requirements.txt` (com o `py -m` na frente, que garante o Python certo). |
| `SSL: CERTIFICATE_VERIFY_FAILED` ou timeout no `pip` | O proxy corporativo está bloqueando o PyPI. Peça à TI o endereço do espelho interno de pacotes. **Não** desative a verificação de certificado (`--trusted-host`) — isso deixa o download sem proteção. |
| `Acesso negado` / pede senha de admin | Você está instalando para todos os usuários. Refaça escolhendo a instalação só para o seu usuário (Passo 1, Opções A a D). |
| `FileNotFoundError: frentes.json` | Você rodou de fora da pasta certa. Faça `cd scripts` antes de chamar o script. |
| `node não é reconhecido` | Só os dois scripts `.js` precisam do Node.js. Os `.py` funcionam sem ele. Instale o Node pelo portal corporativo se precisar reextrair dados do painel HTML. |

---

## Opcional — deixar mais confortável

**Editor.** Instale o **VS Code** (também disponível no portal corporativo) e
a extensão *Python* da Microsoft. Aí dá para abrir a pasta `Gerais`, ler o
código com destaque de sintaxe e rodar os scripts pelo botão ▶.

**Ambiente isolado (venv).** Se você for mexer em vários projetos Python e
quiser evitar conflito de versões:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Com o venv ativo, o prompt mostra `(.venv)` e você usa `python` no lugar de
`py`. Para sair, rode `deactivate`. Se o PowerShell recusar o
`Activate.ps1` por política de execução, use
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned` na mesma
janela e tente de novo — vale só para aquela sessão.
