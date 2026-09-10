"""Verifica se o computador esta pronto para rodar os scripts deste repo.

Uso (a partir da pasta do projeto):

    py scripts/check_ambiente.py

Nao instala nada: so olha o que existe e diz o que falta.
"""

import importlib
import os
import platform
import shutil
import sys

# (nome do import, nome no pip, para que serve)
PACOTES = [
    ("pptx", "python-pptx", "monta e edita os arquivos .pptx"),
    ("PIL", "Pillow", "mede texto e trata imagens"),
    ("lxml", "lxml", "manipula o XML interno do PowerPoint"),
    ("pymupdf", "pymupdf", "le os PDFs de origem (so o treinamento do PDV)"),
]

OK = "[ OK ]"
FALTA = "[FALTA]"
AVISO = "[AVISO]"


def cabecalho(titulo):
    print()
    print(titulo)
    print("-" * len(titulo))


def main():
    problemas = []
    avisos = []

    cabecalho("Python")
    versao = sys.version_info
    print(f"       versao: {platform.python_version()}")
    print(f"       executavel: {sys.executable}")
    if versao < (3, 9):
        problemas.append(
            f"Python {platform.python_version()} e antigo demais; instale a versao 3.11 ou superior."
        )
        print(f"{FALTA} precisa ser 3.9 ou superior (recomendado 3.11+)")
    else:
        print(f"{OK}   versao suficiente")

    cabecalho("Pacotes")
    for modulo, pacote, para_que in PACOTES:
        try:
            importlib.import_module(modulo)
        except ImportError:
            print(f"{FALTA} {pacote:<12} — {para_que}")
            problemas.append(f"Instale o pacote {pacote}.")
        else:
            print(f"{OK}   {pacote:<12} — {para_que}")

    cabecalho("Node.js (opcional)")
    node = shutil.which("node")
    if node:
        print(f"{OK}   encontrado em {node}")
        print("       usado apenas por scripts/extract.js e extract_treinamento_layout.js")
    else:
        print(f"{AVISO} nao encontrado")
        print("       so faz falta se voce for reextrair dados do painel HTML;")
        print("       os scripts .py funcionam sem ele.")
        avisos.append("Node.js ausente (opcional).")

    cabecalho("Pastas do projeto")
    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for pasta in ("scripts", "source", "output"):
        caminho = os.path.join(raiz, pasta)
        if os.path.isdir(caminho):
            print(f"{OK}   {pasta}/")
        else:
            print(f"{FALTA} {pasta}/ nao existe em {raiz}")
            problemas.append(f"A pasta {pasta}/ nao foi encontrada — voce esta na pasta certa do projeto?")

    cabecalho("Resultado")
    if problemas:
        print("Ainda falta resolver:")
        for item in problemas:
            print(f"  - {item}")
        print()
        print("Para instalar os pacotes que faltam, rode na pasta do projeto:")
        print("    py -m pip install -r requirements.txt")
        return 1

    print("Tudo pronto. Voce ja consegue rodar, por exemplo:")
    print("    cd scripts")
    print("    py pmo_deck.py")
    if avisos:
        print()
        print("Observacoes:")
        for item in avisos:
            print(f"  - {item}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
