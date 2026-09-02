#!/usr/bin/env python3
"""Gera o material de treinamento (HTML autocontido) do Simulador de
Financiamento do PDV para os consultores de vendas da Localiza Seminovos.

As telas do produto sao extraidas do PDF em source/ e embutidas em base64,
de modo que o arquivo final funcione offline, em qualquer loja, sem
dependencia de rede ou de pasta de imagens.

    python3 build_treinamento.py
"""
import base64
import pathlib

import pymupdf

ROOT = pathlib.Path(__file__).resolve().parent.parent
PDF = ROOT / "source" / "nova_jornada_pdv_simulacao.pdf"
OUT = ROOT / "output" / "treinamento_simulador_pdv.html"

# (pagina do PDF, indice da imagem na pagina) -> chave usada no template
TELAS = {
    "TELA_BOTAO": (0, 0),      # passo 1 - botao "Simule com os bancos"
    "TELA_DADOS": (1, 0),      # passo 2 - formulario de dados cadastrais
    "TELA_BANCOS": (1, 1),     # passo 2 - consulta aos bancos parceiros
    "TELA_BAIXA": (2, 0),      # passo 3 - sinal amarelo
    "TELA_ALTA": (3, 0),       # passo 3 - sinal verde
}


def extrair_telas() -> dict:
    doc = pymupdf.open(PDF)
    telas = {}
    for chave, (pagina, indice) in TELAS.items():
        xref = doc[pagina].get_images(full=True)[indice][0]
        pix = pymupdf.Pixmap(doc, xref)
        if pix.n > 4:
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        b64 = base64.b64encode(pix.tobytes("png")).decode("ascii")
        telas[chave] = "data:image/png;base64," + b64
    return telas


def main() -> None:
    html = (ROOT / "scripts" / "treinamento_template.html").read_text(encoding="utf-8")
    for chave, data_uri in extrair_telas().items():
        html = html.replace("{{" + chave + "}}", data_uri)
    OUT.write_text(html, encoding="utf-8")
    print(f"{OUT.relative_to(ROOT)} - {OUT.stat().st_size / 1024:.0f} KB")


if __name__ == "__main__":
    main()
