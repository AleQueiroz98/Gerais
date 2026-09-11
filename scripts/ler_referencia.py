# -*- coding: utf-8 -*-
"""Converte os documentos de `source/referencia/` em texto pesquisavel.

Voce joga o arquivo (PDF ou PPTX) na pasta `source/referencia/` e roda:

    python3 scripts/ler_referencia.py      # no Windows: py scripts\\ler_referencia.py

Cada documento vira um `.md` em `source/referencia/_texto/`, com marcacao de
pagina/slide, tabelas e notas do apresentador. E esse `.md` que o agente
`referencia` le -- texto puro, sem dependencia de biblioteca, versionado junto
com o repo, entao qualquer sessao consegue consultar o documento depois.

Opcoes:

    --forcar    reextrai tudo, mesmo o que ja esta atualizado
    --listar    so mostra o que existe na pasta, sem extrair
"""
import argparse
import datetime as dt
import os
import re
import sys
import unicodedata

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA = os.path.join(RAIZ, "source", "referencia")
SAIDA = os.path.join(PASTA, "_texto")

# Extensoes que sabemos abrir. `copia` = texto puro, so transcreve.
SUPORTADAS = {
    ".pdf": "pdf",
    ".pptx": "pptx",
    ".ppt": "aviso_ppt",
    ".doc": "aviso_doc",
    ".txt": "copia",
    ".md": "copia",
    ".csv": "copia",
    ".html": "copia",
}


def slug(nome):
    """`Memo Lead-to-Sales v2.pdf` -> `memo_lead_to_sales_v2`."""
    base = os.path.splitext(nome)[0]
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode()
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_").lower()
    return base or "documento"


def assinatura(caminho):
    """Tamanho + mtime: barato e suficiente para saber se o arquivo mudou."""
    st = os.stat(caminho)
    return "%d-%d" % (st.st_size, int(st.st_mtime))


# ------------------------------------------------------------------ PDF
def extrair_pdf(caminho):
    try:
        import pymupdf
    except ImportError:  # pymupdf < 1.24 so expoe o nome antigo
        import fitz as pymupdf

    doc = pymupdf.open(caminho)
    partes, vazias = [], []
    for i, pagina in enumerate(doc, start=1):
        texto = pagina.get_text("text").strip()
        if not texto:
            vazias.append(i)
            texto = "_(pagina sem texto selecionavel -- provavelmente imagem ou scan)_"
        partes.append("## Pagina %d\n\n%s" % (i, texto))
    doc.close()

    avisos = []
    if vazias:
        avisos.append(
            "%d pagina(s) sem texto selecionavel (%s). Se o conteudo importa, "
            "mande um print ou o arquivo original editavel."
            % (len(vazias), ", ".join(str(p) for p in vazias[:12]))
        )
    return "\n\n".join(partes), {"paginas": len(partes)}, avisos


# ----------------------------------------------------------------- PPTX
def _texto_tabela(tabela):
    """Tabela do PowerPoint -> tabela markdown (primeira linha vira cabecalho)."""
    linhas = []
    for linha in tabela.rows:
        celulas = [c.text.replace("\n", " ").replace("|", "\\|").strip() for c in linha.cells]
        linhas.append("| " + " | ".join(celulas) + " |")
    if linhas:
        n = len(tabela.columns)
        linhas.insert(1, "|" + "|".join([" --- "] * n) + "|")
    return "\n".join(linhas)


def _varrer_shapes(shapes, saida):
    """Percorre shapes recursivamente (grupos incluidos), na ordem do slide."""
    for shape in shapes:
        if shape.shape_type == 6 and hasattr(shape, "shapes"):  # GROUP
            _varrer_shapes(shape.shapes, saida)
            continue
        if getattr(shape, "has_table", False) and shape.has_table:
            saida.append(_texto_tabela(shape.table))
            continue
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            texto = shape.text_frame.text.strip()
            if texto:
                saida.append(texto)


def extrair_pptx(caminho):
    from pptx import Presentation

    prs = Presentation(caminho)
    partes = []
    for i, slide in enumerate(prs.slides, start=1):
        blocos = []
        _varrer_shapes(slide.shapes, blocos)
        if slide.has_notes_slide:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                blocos.append("**Notas do apresentador:** " + notas)
        corpo = "\n\n".join(blocos) if blocos else "_(slide sem texto)_"
        partes.append("## Slide %d\n\n%s" % (i, corpo))
    return "\n\n".join(partes), {"slides": len(partes)}, []


# ----------------------------------------------------------------- outros
def extrair_copia(caminho):
    with open(caminho, "r", encoding="utf-8", errors="replace") as fh:
        texto = fh.read()
    return texto, {"linhas": texto.count("\n") + 1}, []


def extrair_aviso(caminho, formato):
    novo = ".pptx" if formato == "aviso_ppt" else ".docx"
    aviso = (
        "Formato antigo (%s). Abra no Office e salve como %s na mesma pasta -- "
        "ai o extrator consegue ler." % (os.path.splitext(caminho)[1], novo)
    )
    return "_(%s)_" % aviso, {}, [aviso]


EXTRATORES = {
    "pdf": extrair_pdf,
    "pptx": extrair_pptx,
    "copia": extrair_copia,
}


# ------------------------------------------------------------------ main
def documentos():
    """Arquivos soltos na pasta de referencia, ignorando `_texto/` e ocultos."""
    if not os.path.isdir(PASTA):
        return []
    itens = []
    for nome in sorted(os.listdir(PASTA)):
        caminho = os.path.join(PASTA, nome)
        if not os.path.isfile(caminho) or nome.startswith(".") or nome.startswith("_"):
            continue
        if nome.upper().startswith("LEIA-ME"):
            continue
        itens.append((nome, caminho))
    return itens


def extrair(nome, caminho, forcar):
    ext = os.path.splitext(nome)[1].lower()
    formato = SUPORTADAS.get(ext)
    if not formato:
        return None, ["Extensao %s nao suportada -- ignorado." % ext], False

    destino = os.path.join(SAIDA, slug(nome) + ".md")
    sig = assinatura(caminho)
    if not forcar and os.path.exists(destino):
        with open(destino, "r", encoding="utf-8") as fh:
            cabeca = fh.read(600)
        if ("assinatura: " + sig) in cabeca:
            return destino, [], False  # ja atualizado

    if formato.startswith("aviso_"):
        corpo, meta, avisos = extrair_aviso(caminho, formato)
    else:
        corpo, meta, avisos = EXTRATORES[formato](caminho)

    detalhe = ", ".join("%s: %s" % (k, v) for k, v in meta.items())
    cabecalho = (
        "# %s\n\n"
        "> Extraido de `source/referencia/%s`%s.\n"
        "> Gerado por `scripts/ler_referencia.py` em %s.\n"
        "> Nao edite este arquivo a mao -- edite o original e rode o script de novo.\n"
        "> assinatura: %s\n"
        % (
            os.path.splitext(nome)[0],
            nome,
            (" (" + detalhe + ")") if detalhe else "",
            dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            sig,
        )
    )
    os.makedirs(SAIDA, exist_ok=True)
    with open(destino, "w", encoding="utf-8") as fh:
        fh.write(cabecalho + "\n" + corpo.rstrip() + "\n")
    return destino, avisos, True


def escrever_indice(linhas):
    os.makedirs(SAIDA, exist_ok=True)
    with open(os.path.join(SAIDA, "INDICE.md"), "w", encoding="utf-8") as fh:
        fh.write("# Documentos de referencia\n\n")
        fh.write("Gerado por `scripts/ler_referencia.py`. Um `.md` por documento.\n\n")
        fh.write("| Documento | Texto extraido | Tamanho |\n|---|---|---|\n")
        for nome, destino, tamanho in linhas:
            arq = os.path.basename(destino) if destino else "-"
            fh.write("| `%s` | `%s` | %s |\n" % (nome, arq, tamanho))


def humano(n):
    for unidade in ("B", "KB", "MB"):
        if n < 1024 or unidade == "MB":
            return "%.0f %s" % (n, unidade)
        n /= 1024.0


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--forcar", action="store_true", help="reextrai tudo")
    ap.add_argument("--listar", action="store_true", help="so lista a pasta")
    args = ap.parse_args()

    itens = documentos()
    if not itens:
        print("Nada em source/referencia/.")
        print("Jogue o PDF ou o PPTX nessa pasta e rode o script de novo.")
        return 0

    if args.listar:
        print("Documentos em source/referencia/:")
        for nome, caminho in itens:
            print("  %-50s %8s" % (nome, humano(os.path.getsize(caminho))))
        return 0

    os.makedirs(SAIDA, exist_ok=True)
    indice, avisos_gerais, erros = [], [], []
    for nome, caminho in itens:
        try:
            destino, avisos, novo = extrair(nome, caminho, args.forcar)
        except ImportError as exc:
            erros.append("%s: falta uma biblioteca (%s). Rode "
                         "`pip install -r requirements.txt`." % (nome, exc))
            destino, avisos, novo = None, [], False
        except Exception as exc:  # arquivo corrompido, protegido por senha, etc.
            erros.append("%s: nao consegui abrir (%s)." % (nome, exc))
            destino, avisos, novo = None, [], False

        indice.append((nome, destino, humano(os.path.getsize(caminho))))
        if destino:
            marca = "[ OK ]" if novo else "[ = ]"
            print("%s %-45s -> _texto/%s" % (marca, nome, os.path.basename(destino)))
        for aviso in avisos:
            avisos_gerais.append("%s: %s" % (nome, aviso))

    escrever_indice(indice)

    for aviso in avisos_gerais:
        print("[AVISO] " + aviso)
    for erro in erros:
        print("[ERRO ] " + erro)

    print("\nTexto pronto em source/referencia/_texto/.")
    print("Agora e so pedir: 'usa o documento X como referencia para ...'")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
