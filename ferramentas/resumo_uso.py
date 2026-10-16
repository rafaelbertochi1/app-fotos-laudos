r"""
Resumo do que a equipe fez com o programa (pasta_publica.py grava):
quantos Excel por dia e por pessoa, versões em uso e os erros recentes.

    python ferramentas\resumo_uso.py            mês atual
    python ferramentas\resumo_uso.py 2026-10    outro mês
    python ferramentas\resumo_uso.py --erros 20 mostra os 20 erros mais recentes (padrão 10)
"""

import argparse
import csv
import glob
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pasta_publica import PASTA_ERROS, PASTA_USO  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mes", nargs="?", default=f"{datetime.now():%Y-%m}", help="aaaa-mm")
    ap.add_argument("--erros", type=int, default=10)
    a = ap.parse_args()

    linhas = []
    for arq in glob.glob(os.path.join(PASTA_USO, f"*_{a.mes}.csv")):
        with open(arq, encoding="utf-8-sig", newline="") as f:
            linhas += list(csv.DictReader(f, delimiter=";"))
    print(f"USO em {a.mes} ({PASTA_USO}): {len(linhas)} Excel gerado(s)")
    if linhas:
        por_dia = Counter(l["data"] for l in linhas)
        print("\nPor dia:")
        for dia in sorted(por_dia, key=lambda d: datetime.strptime(d, "%d/%m/%Y")):
            print(f"  {dia}  {por_dia[dia]:4d}")
        print("\nPor pessoa (Excel / propostas diferentes / tempo médio da busca ao Excel):")
        pessoas = defaultdict(list)
        for l in linhas:
            pessoas[l["usuario"]].append(l)
        for pessoa, ls in sorted(pessoas.items(), key=lambda p: -len(p[1])):
            tempos = [int(l["segundos"]) for l in ls if (l.get("segundos") or "").isdigit()]
            medio = f"{sum(tempos) / len(tempos) / 60:.1f} min" if tempos else "-"
            print(f"  {pessoa:24s} {len(ls):4d}  {len({l['proposta'] for l in ls}):4d}  {medio}")
        print("\nPor banco / modelo:")
        for (banco, modelo), n in Counter((l["banco"], l["modelo"]) for l in linhas).most_common():
            print(f"  {banco:10s} {modelo:34s} {n:4d}")
        print("\nVersões em uso:", ", ".join(f"v{v} ({n})" for v, n in Counter(l["versao"] for l in linhas).most_common()))
        retomados = sum(1 for l in linhas if l.get("retomado") == "sim")
        print(f"Trabalhos retomados: {retomados}   Com fotos do computador: "
              f"{sum(1 for l in linhas if (l.get('fotos_do_pc') or '0') != '0')}")

    erros = sorted(glob.glob(os.path.join(PASTA_ERROS, "*.txt")), reverse=True)
    print(f"\nERROS ({PASTA_ERROS}): {len(erros)} no total; os {min(a.erros, len(erros))} mais recentes:")
    for arq in erros[:a.erros]:
        with open(arq, encoding="utf-8", errors="replace") as f:
            texto = f.read()
        erro = next((l[6:] for l in texto.splitlines() if l.startswith("ERRO: ")), "?")
        print(f"  {os.path.basename(arq)[:-4]}: {erro[:110]}")


if __name__ == "__main__":
    main()
