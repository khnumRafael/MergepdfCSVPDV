# -*- coding: utf-8 -*-
"""Extrai itens do PDF e agrega por vendedor/codigo/barras/descricao/peso/unidade."""
import csv
import re
from collections import defaultdict
from pypdf import PdfReader

PDF_PATH = r"c:\tmp\pedidos pr\porVendedor.pdf"
OUT_CSV = r"c:\tmp\pedidos pr\relatorio_agregado.csv"

PEDIDO_RE = re.compile(r"Pedido:\s*(\d+)", re.I)
LINE_RE = re.compile(r"^(\d{6})\s+(\d+)\s+(\d+,\d{2})(\d+,\d{2})(\d+,\d{2})(.+)$")
QTY_END = re.compile(r"^(.+?)\s+((?:[A-Z]\s*){1,5})\s+(\d+,\d{2})\s*$")
PESO_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kg|g|gr|ml|l)\b", re.I)


def extrair_peso(descricao: str) -> str:
    matches = list(PESO_RE.finditer(descricao or ""))
    if not matches:
        return ""
    m = matches[-1]
    valor = m.group(1).replace(".", ",")
    un = m.group(2).lower()
    mapa_un = {"kg": "kg", "g": "g", "gr": "g", "ml": "ml", "l": "l"}
    return f"{valor} {mapa_un.get(un, un)}"


def peso_para_kg(peso: str) -> float:
    p = (peso or "").strip().lower()
    if not p:
        return 0.0
    m = re.match(r"(\d+(?:[.,]\d+)?)\s*(kg|g|ml|l)\b", p, re.I)
    if not m:
        return 0.0
    valor = float(m.group(1).replace(",", "."))
    un = m.group(2).lower()
    if un == "kg":
        return valor
    if un in ("g", "ml"):
        return valor / 1000.0
    if un == "l":
        return valor
    return 0.0


def parse_line_item(line: str):
    line = line.strip()
    if not line or "Total do Pedido" in line or "Total Geral" in line or line.startswith("R$"):
        return None
    m = LINE_RE.match(line)
    if not m:
        return None

    cod, barras, vt, _disc, _vu, rest = m.groups()
    total = float(vt.replace(".", "").replace(",", "."))

    qm = QTY_END.match(rest.strip())
    if not qm:
        m2 = re.search(r"(.+?)\s+(\d+,\d{2})\s*$", rest.strip())
        if not m2:
            return None
        desc = re.sub(r"\s+", " ", m2.group(1).strip())
        qty = float(m2.group(2).replace(".", "").replace(",", "."))
        unidade = ""
    else:
        desc, unidade, q = qm.groups()
        desc = re.sub(r"\s+", " ", desc.strip())
        unidade = re.sub(r"\s+", " ", unidade.strip())
        qty = float(q.replace(".", "").replace(",", "."))

    return {
        "codigo": cod,
        "barras": barras,
        "descricao": desc,
        "peso": extrair_peso(desc),
        "unidade": unidade,
        "quantidade": qty,
        "total": total,
    }


def main():
    reader = PdfReader(PDF_PATH)
    full_text = "\n".join((p.extract_text() or "") for p in reader.pages)
    lines = full_text.splitlines()

    vendedor = "Iadora"
    pedido_atual = ""
    items = []

    for ln in lines:
        if "Vendedor.:" in ln:
            if "Iadora" in ln:
                vendedor = "Iadora"
            elif "P a t r i c i a" in ln:
                vendedor = "Patricia"

        m_ped = PEDIDO_RE.search(ln)
        if m_ped:
            pedido_atual = m_ped.group(1)

        rec = parse_line_item(ln)
        if rec:
            items.append({**rec, "vendedor": vendedor, "pedido": pedido_atual or ""})

    agg = defaultdict(lambda: {"q": 0.0, "t": 0.0})
    pedidos_por_vendedor = defaultdict(set)

    for it in items:
        key = (
            it["vendedor"],
            it["codigo"],
            it["barras"],
            it["descricao"],
            it["peso"],
            it["unidade"],
        )
        agg[key]["q"] += it["quantidade"]
        agg[key]["t"] += it["total"]
        if it.get("pedido"):
            pedidos_por_vendedor[it["vendedor"]].add(it["pedido"])

    rows = sorted(agg.keys(), key=lambda k: (k[0], k[3]))

    tot_por_vendedor = defaultdict(lambda: {"q": 0.0, "t": 0.0, "peso_kg": 0.0})
    for k in rows:
        v = k[0]
        tot_por_vendedor[v]["q"] += agg[k]["q"]
        tot_por_vendedor[v]["t"] += agg[k]["t"]
        p = k[4]
        tot_por_vendedor[v]["peso_kg"] += peso_para_kg(p) * agg[k]["q"]

    def fmt_br_num(x: float) -> str:
        return f"{x:.2f}".replace(".", ",")

    def escrever_rodape_vendedor(w, nome_v: str) -> None:
        tv = tot_por_vendedor[nome_v]
        w.writerow([
            nome_v,
            "__SUBTOTAL__",
            "",
            "Subtotal do vendedor",
            f"{fmt_br_num(tv['peso_kg'])} kg",
            "",
            fmt_br_num(tv["q"]),
            fmt_br_num(tv["t"]),
        ])
        pids = sorted(pedidos_por_vendedor[nome_v])
        texto = "Pedidos: " + ", ".join(pids) if pids else "Pedidos: (não identificado)"
        w.writerow([nome_v, "__PEDIDOS__", "", texto, "", "", "", ""])

    with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([
            "vendedor",
            "codigo",
            "barras",
            "descricao",
            "peso",
            "unidade",
            "sum_quantidade",
            "sum_total",
        ])

        prev_v = None
        for k in rows:
            v, c, b, d, p, u = k
            if prev_v is not None and v != prev_v:
                escrever_rodape_vendedor(w, prev_v)
            prev_v = v
            w.writerow([v, c, b, d, p, u, fmt_br_num(agg[k]["q"]), fmt_br_num(agg[k]["t"])])

        if prev_v is not None:
            escrever_rodape_vendedor(w, prev_v)

        peso_total_kg = 0.0
        for k in rows:
            _, _, _, _, p, _ = k
            peso_total_kg += peso_para_kg(p) * agg[k]["q"]
        w.writerow(["TOTAL", "__TOTAL_PESO__", "", "Peso total estimado (kg)", "", "", "", fmt_br_num(peso_total_kg)])

    total_geral = sum(agg[k]["t"] for k in rows)
    print(f"Registros de item: {len(items)} | Grupos: {len(rows)} | Total R$ {total_geral:.2f}")
    print(f"Arquivo: {OUT_CSV}")


if __name__ == "__main__":
    main()
