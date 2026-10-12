"""Monitor de estoque: iPhone 18 Pro Max 256GB/512GB (desbloqueado) em lojas Apple perto de Miami.

Roda sozinho no seu computador, sem depender do Claude. So usa a biblioteca padrao do Python.
A cada N minutos consulta a API de retirada em loja da Apple. Quando um modelo monitorado
passa a ficar disponivel, avisa no console (com beep no Windows) e, se informado um topico
ntfy, manda push para o celular (app "ntfy", gratuito: assine o mesmo topico no app).

A API devolve so uma amostra de modelos "parecidos" ao consultado; por isso consultamos
todos os modelos e unimos os resultados.

Uso (Windows: py no lugar de python3):
  py scripts/monitor_iphone.py --ntfy meu-topico-secreto-123
  py scripts/monitor_iphone.py --cep "Miami, FL" --intervalo 5 --modelo "18 Pro Max 256GB" --modelo "18 Pro Max 512GB"
  py scripts/monitor_iphone.py --uma-vez        (uma checagem e sai; bom para testar)

Na nuvem: .github/workflows/monitor-iphone.yml roda --uma-vez --falhar-se-disponivel a cada
5 min; quando acha estoque o job falha e o GitHub avisa por e-mail (nao precisa deixar nada ligado).
"""
import argparse
import json
import sys
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

NOMES = {
    "MJWE4LL/A": "18 Pro Max 1TB Silver",
    "MJWD4LL/A": "18 Pro Max 1TB Black",
    "MJWC4LL/A": "18 Pro Max 512GB Glacier",
    "MJWH4LL/A": "18 Pro Max 2TB Black",
    "MJWG4LL/A": "18 Pro Max 1TB Glacier",
    "MJWF4LL/A": "18 Pro Max 1TB Burgundy",
    "MJWL4LL/A": "18 Pro Max 2TB Glacier",
    "MJWK4LL/A": "18 Pro Max 2TB Burgundy",
    "MJWJ4LL/A": "18 Pro Max 2TB Silver",
    "MJW54LL/A": "18 Pro Max 256GB Silver",
    "MJW44LL/A": "18 Pro Max 256GB Black",
    "MJW94LL/A": "18 Pro Max 512GB Silver",
    "MJW84LL/A": "18 Pro Max 512GB Black",
    "MJW74LL/A": "18 Pro Max 256GB Glacier",
    "MJW64LL/A": "18 Pro Max 256GB Burgundy",
    "MJWA4LL/A": "18 Pro Max 512GB Burgundy",
    "MJQ84LL/A": "18 Pro 512GB Silver",
    "MJQ94LL/A": "18 Pro 512GB Burgundy",
    "MJQ64LL/A": "18 Pro 256GB Glacier",
    "MJQ74LL/A": "18 Pro 512GB Black",
    "MJQA4LL/A": "18 Pro 512GB Glacier",
    "MJQD4LL/A": "18 Pro 1TB Silver",
    "MJQE4LL/A": "18 Pro 1TB Burgundy",
    "MJQC4LL/A": "18 Pro 1TB Black",
    "MJQH4LL/A": "18 Pro 2TB Silver",
    "MJQJ4LL/A": "18 Pro 2TB Burgundy",
    "MJQF4LL/A": "18 Pro 1TB Glacier",
    "MJQG4LL/A": "18 Pro 2TB Black",
    "MJQK4LL/A": "18 Pro 2TB Glacier",
    "MJQ44LL/A": "18 Pro 256GB Silver",
    "MJQ54LL/A": "18 Pro 256GB Burgundy",
    "MJQ34LL/A": "18 Pro 256GB Black"
}

URL = ("https://www.apple.com/shop/pickup-message-recommendations?mts.0=regular&mts.1=compact"
       "&cppart=UNLOCKED/US&searchNearby=true&product={part}&location={loc}")


def consulta(part, cep):
    url = URL.format(part=urllib.parse.quote(part, safe="/"), loc=urllib.parse.quote(cep))
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read())["body"]["PickupMessage"]["stores"]


def disponiveis(cep, filtros):
    with ThreadPoolExecutor(8) as ex:
        resultados = list(ex.map(lambda p: consulta(p, cep), NOMES))
    achados = set()
    for stores in resultados:
        for s in stores:
            for part, info in s["partsAvailability"].items():
                nome = NOMES.get(part, part)
                if info.get("pickupDisplay") == "available" and any(f in nome.lower() for f in filtros):
                    achados.add(f"{nome} @ {s['storeName']}")
    return achados


def avisar(texto, topico):
    print(f"\a[{datetime.now():%H:%M:%S}] DISPONIVEL: {texto}", flush=True)
    try:
        import winsound
        winsound.Beep(1200, 600)
    except ImportError:
        pass
    if topico:
        try:
            req = urllib.request.Request(f"https://ntfy.sh/{topico}", data=texto.encode("utf-8"),
                                         headers={"Title": "iPhone disponivel", "Priority": "high"})
            urllib.request.urlopen(req, timeout=20).read()
        except Exception as e:
            print(f"  (falha ao enviar ntfy: {e})", flush=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cep", default="33101", help="CEP ou cidade (padrao: 33101, Miami)")
    ap.add_argument("--intervalo", type=float, default=5, help="minutos entre checagens (padrao: 5)")
    ap.add_argument("--modelo", action="append",
                    help='trecho do nome a monitorar, repetivel (padrao: "18 Pro Max 256GB" e "18 Pro Max 512GB")')
    ap.add_argument("--ntfy", help="topico ntfy.sh para push no celular")
    ap.add_argument("--uma-vez", action="store_true", help="checa uma vez e sai")
    ap.add_argument("--falhar-se-disponivel", action="store_true",
                    help="com --uma-vez, sai com codigo 1 quando achar estoque (usado pelo GitHub Actions)")
    a = ap.parse_args()
    filtros = [m.lower() for m in (a.modelo or ["18 Pro Max 256GB", "18 Pro Max 512GB"])]

    avisados = set()
    while True:
        try:
            agora = disponiveis(a.cep, filtros)
            novos = agora - avisados
            if novos:
                avisar("; ".join(sorted(novos)), a.ntfy)
            elif not agora:
                print(f"[{datetime.now():%H:%M:%S}] indisponivel", flush=True)
            avisados = agora  # se sumir e voltar, avisa de novo
            if a.uma_vez and a.falhar_se_disponivel and agora:
                sys.exit(1)
        except Exception as e:
            print(f"[{datetime.now():%H:%M:%S}] erro: {e}", file=sys.stderr, flush=True)
        if a.uma_vez:
            return
        time.sleep(a.intervalo * 60)


if __name__ == "__main__":
    main()
