# -*- coding: utf-8 -*-
"""Audita relatorio_agregado.csv contra PDF(s) de origem em c:\\pdf."""
import csv
import re
from pathlib import Path

from pypdf import PdfReader

BASE = Path(r"c:\pdf")
PDF_V = BASE / "porVendedor.pdf"
PDF_C = BASE / "carga.pdf"
CSV_OUT = BASE / "relatorio_agregado.csv"

LINE_RE = re.compile(
    r"^(\d{6})\s+(\d+)\s+(\d+,\d{2})(\d+,\d{2})(\d+,\d{2})(.+)$"
)


def parse_line_total(line: str):
    line = line.strip()
    if (
        not line
        or "Total do Pedido" in line
        or "Total Geral" in line
        or line.startswith("R$")
    ):
        return None
    m = LINE_RE.match(line)
    if not m:
        return None
    return float(m.group(3).replace(".", "").replace(",", "."))


def scan_pdf(path: Path):
    text = "\n".join((p.extract_text() or "") for p in PdfReader(str(path)).pages)
    lines = text.splitlines()
    totals = []
    for ln in lines:
        t = parse_line_total(ln)
        if t is not None:
            totals.append(t)
    tg = re.findall(
        r"Total\s+Geral[^0-9R$]*R?\$?\s*([\d.,]+)", text.replace("\n", " "), re.I
    )
    rg = re.findall(r"Registros\s+Listados\s*=\s*(\d+)", text, re.I)
    sub_v = re.findall(
        r"R\$\s*([\d.,]+)\s*R\$\s*0,00\s*R\$\s*([\d.,]+)\s*Total\s+Geral",
        text.replace("\n", " "),
        re.I,
    )
    return {
        "path": path.name,
        "itens_regex": len(totals),
        "soma_linhas": sum(totals),
        "registros_listados": rg[0] if rg else None,
        "total_geral_texto": tg[-1] if tg else None,
    }


def ler_csv():
    grupos = 0
    sq = st = 0.0
    peso_linha = None
    with open(CSV_OUT, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f, delimiter=";")
        next(r, None)
        for row in r:
            if len(row) < 2:
                continue
            if row[1] == "__TOTAL_PESO__":
                peso_linha = row[-1] if row else None
                continue
            if row[1] in ("__SUBTOTAL__", "__PEDIDOS__"):
                continue
            if len(row) >= 8:
                grupos += 1
                sq += float(row[6].replace(".", "").replace(",", "."))
                st += float(row[7].replace(".", "").replace(",", "."))
    return grupos, sq, st, peso_linha


def main():
    print("Auditoria: relatório gerado x PDF(s) de origem\n")

    if not PDF_V.exists():
        print(f"ERRO: não encontrado {PDF_V}")
        return

    pv = scan_pdf(PDF_V)
    print(f"=== {pv['path']} (único PDF usado pelo relatorio_unico.py) ===")
    print(f"  Linhas de item capturadas (mesma regex do extrator): {pv['itens_regex']}")
    print(f"  Soma dos totais por linha de item: R$ {pv['soma_linhas']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    if pv["registros_listados"]:
        print(f"  'Registros Listados' no texto do PDF: {pv['registros_listados']}")
    if pv["total_geral_texto"]:
        print(f"  'Total Geral' (trecho extraído do PDF): {pv['total_geral_texto']}")

    if PDF_C.exists():
        pc = scan_pdf(PDF_C)
        print(f"\n=== {pc['path']} (NÃO é usado pelo relatorio_unico.py) ===")
        print(f"  Linhas de item (regex): {pc['itens_regex']}")
        print(f"  Soma dos totais por linha: R$ {pc['soma_linhas']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        if pc["registros_listados"]:
            print(f"  Registros Listados: {pc['registros_listados']}")

    if not CSV_OUT.exists():
        print(f"\nERRO: CSV não encontrado {CSV_OUT}")
        return

    g, sq, st, peso = ler_csv()
    print(f"\n=== {CSV_OUT.name} (saída agregada) ===")
    print(f"  Grupos (linhas de produto): {g}")
    print(f"  Soma quantidades: {sq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print(f"  Soma totais R$: {st:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    if peso:
        print(f"  Peso total (linha TOTAL): {peso} kg")

    diff = abs(pv["soma_linhas"] - st)
    print("\n=== Conclusão ===")
    print(
        f"  Soma financeira PDF (itens) vs CSV agregado: "
        f"{'OK' if diff < 0.05 else 'DIVERGENTE'} (diferença R$ {diff:,.4f})".replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    if pv["registros_listados"] and int(pv["registros_listados"]) != pv["itens_regex"]:
        print(
            f"  Aviso: 'Registros Listados' ({pv['registros_listados']}) ≠ "
            f"linhas parseadas ({pv['itens_regex']}) — possível mudança de layout no PDF."
        )
    print(
        "\n  O executável relatorio_unico só faz merge/extração a partir de "
        "porVendedor.pdf; carga.pdf não entra no cálculo."
    )


if __name__ == "__main__":
    main()
