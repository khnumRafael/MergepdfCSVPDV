# -*- coding: utf-8 -*-
"""Executável único: extrai dados do PDF e gera CSV/PDF em c:\\pdf."""
import csv
import re
from collections import OrderedDict, defaultdict
from datetime import date
from pathlib import Path

from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE_DIR = Path(r"c:\pdf")
IN_PDF = BASE_DIR / "porVendedor.pdf"
OUT_CSV = BASE_DIR / "relatorio_agregado.csv"
OUT_PDF = BASE_DIR / "relatorio_agregado.pdf"

PEDIDO_RE = re.compile(r"Pedido:\s*(\d+)", re.I)
LINE_RE = re.compile(r"^(\d{6})\s+(\d+)\s+(\d+,\d{2})(\d+,\d{2})(\d+,\d{2})(.+)$")
QTY_END = re.compile(r"^(.+?)\s+((?:[A-Z]\s*){1,5})\s+(\d+,\d{2})\s*$")
PESO_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kg|g|gr|ml|l)\b", re.I)


def fmt_br(val: float) -> str:
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_br_decimal(s: str) -> float:
    s = (s or "").strip().replace(".", "").replace(",", ".")
    return float(s) if s else 0.0


def esc_xml(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


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


def extrair_peso(descricao: str) -> str:
    matches = list(PESO_RE.finditer(descricao or ""))
    if not matches:
        return ""
    m = matches[-1]
    valor = m.group(1).replace(".", ",")
    un = m.group(2).lower()
    mapa_un = {"kg": "kg", "g": "g", "gr": "g", "ml": "ml", "l": "l"}
    return f"{valor} {mapa_un.get(un, un)}"


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


def localizar_pdf_entrada() -> Path:
    if IN_PDF.exists():
        return IN_PDF
    pdfs = sorted(BASE_DIR.glob("*.pdf"))
    if len(pdfs) == 1:
        return pdfs[0]
    raise FileNotFoundError(
        f"PDF de entrada não encontrado. Coloque 'porVendedor.pdf' em {BASE_DIR}"
    )


def gerar_csv(pdf_path: Path):
    reader = PdfReader(str(pdf_path))
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
        key = (it["vendedor"], it["codigo"], it["barras"], it["descricao"], it["peso"], it["unidade"])
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
        tot_por_vendedor[v]["peso_kg"] += peso_para_kg(k[4]) * agg[k]["q"]

    def escrever_rodape_vendedor(w, nome_v):
        tv = tot_por_vendedor[nome_v]
        w.writerow([nome_v, "__SUBTOTAL__", "", "Subtotal do vendedor", f"{fmt_br(tv['peso_kg'])} kg", "", fmt_br(tv["q"]), fmt_br(tv["t"])])
        pids = sorted(pedidos_por_vendedor[nome_v])
        texto = "Pedidos: " + ", ".join(pids) if pids else "Pedidos: (não identificado)"
        w.writerow([nome_v, "__PEDIDOS__", "", texto, "", "", "", ""])

    with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["vendedor", "codigo", "barras", "descricao", "peso", "unidade", "sum_quantidade", "sum_total"])
        prev_v = None
        for k in rows:
            v, c, b, d, p, u = k
            if prev_v is not None and v != prev_v:
                escrever_rodape_vendedor(w, prev_v)
            prev_v = v
            w.writerow([v, c, b, d, p, u, fmt_br(agg[k]["q"]), fmt_br(agg[k]["t"])])
        if prev_v is not None:
            escrever_rodape_vendedor(w, prev_v)
        peso_total_kg = sum(peso_para_kg(k[4]) * agg[k]["q"] for k in rows)
        w.writerow(["TOTAL", "__TOTAL_PESO__", "", "Peso total estimado (kg)", "", "", "", fmt_br(peso_total_kg)])
    return len(items), len(rows)


def load_rows():
    data_rows = []
    pedidos_por_vendedor = {}
    with open(OUT_CSV, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f, delimiter=";")
        next(r, None)
        for row in r:
            if not row:
                continue
            if len(row) >= 4 and row[1] == "__PEDIDOS__":
                pedidos_por_vendedor[row[0]] = row[3]
            elif len(row) >= 2 and row[1] in ("__SUBTOTAL__", "__TOTAL_PESO__"):
                continue
            elif len(row) >= 8:
                data_rows.append(row)
    return data_rows, pedidos_por_vendedor


def gerar_pdf():
    data_rows, pedidos_por_vendedor = load_rows()
    total_geral = sum(parse_br_decimal(r[7]) for r in data_rows)
    peso_total_kg = sum(peso_para_kg(r[4]) * parse_br_decimal(r[6]) for r in data_rows)
    ordem_vendedores = list(OrderedDict.fromkeys(row[0] for row in data_rows))

    page_size = landscape(A4)
    doc = SimpleDocTemplate(str(OUT_PDF), pagesize=page_size, leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontSize=14, alignment=1, spaceAfter=5)
    sub_style = ParagraphStyle("S", parent=styles["Normal"], fontSize=9, alignment=1, textColor=colors.HexColor("#444444"))
    ped_style = ParagraphStyle("P", parent=styles["Normal"], fontSize=8, leftIndent=3 * mm, spaceBefore=2, spaceAfter=8, textColor=colors.HexColor("#4a5568"))
    subhead_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=10, textColor=colors.HexColor("#2c5282"), spaceBefore=5, spaceAfter=4)

    story = [
        Paragraph("Relação entre relatórios por código do produto", title_style),
        Paragraph(f"Gerado em {date.today().strftime('%d/%m/%Y')} · Agrupamento: vendedor, código, barras, descrição, peso, unidade", sub_style),
        Spacer(1, 6),
    ]

    hdr = ["Código", "Barras", "Descrição", "Peso", "Un.", "Σ Qtd", "Σ Total (R$)"]
    col_widths = [16 * mm, 24 * mm, page_size[0] - 20 * mm - (16 + 24 + 16 + 12 + 17 + 20) * mm, 16 * mm, 12 * mm, 17 * mm, 20 * mm]

    for idx_v, v in enumerate(ordem_vendedores):
        rows_v = [r for r in data_rows if r[0] == v]
        if not rows_v:
            continue
        if idx_v > 0:
            story.append(Spacer(1, 4))
        story.append(Paragraph(f"<b>Vendedor: {esc_xml(v)}</b>", subhead_style))
        table_data = [hdr]
        for row in rows_v:
            _, c, b, d, p, u, q, t = row
            table_data.append([c, b, Paragraph(esc_xml(d), ParagraphStyle("c", parent=styles["Normal"], fontSize=7, leading=9)), p, u, q, t])
        sub_q = sum(parse_br_decimal(r[6]) for r in rows_v)
        sub_t = sum(parse_br_decimal(r[7]) for r in rows_v)
        sub_peso_kg = sum(peso_para_kg(r[4]) * parse_br_decimal(r[6]) for r in rows_v)
        table_data.append(["", "", Paragraph("<b>Subtotal do vendedor</b>", ParagraphStyle("sub", parent=styles["Normal"], fontSize=8)), Paragraph(f"<b>{fmt_br(sub_peso_kg)} kg</b>", ParagraphStyle("sp", parent=styles["Normal"], fontSize=8)), "", Paragraph(f"<b>{fmt_br(sub_q)}</b>", ParagraphStyle("sq", parent=styles["Normal"], fontSize=8)), Paragraph(f"<b>{fmt_br(sub_t)}</b>", ParagraphStyle("st", parent=styles["Normal"], fontSize=8))])
        last_i = len(table_data) - 1
        tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c5282")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("FONTSIZE", (0, 1), (-1, -1), 7),
            ("ALIGN", (5, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, last_i - 1), [colors.white, colors.HexColor("#f7fafc")]),
            ("BACKGROUND", (0, last_i), (-1, last_i), colors.HexColor("#e2e8f0")),
            ("LINEABOVE", (0, last_i), (-1, last_i), 1, colors.HexColor("#2c5282")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e0")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(tbl)
        ped_txt = pedidos_por_vendedor.get(v, "")
        if ped_txt:
            story.append(Paragraph(f"<i>{esc_xml(ped_txt)}</i>", ped_style))

    story.append(Spacer(1, 6))
    story.append(Paragraph(f"<b>Peso total estimado: {fmt_br(peso_total_kg)} kg</b>", ParagraphStyle("tot_peso", parent=styles["Normal"], fontSize=10, alignment=2, textColor=colors.HexColor("#2c5282"))))
    story.append(Paragraph(f"<b>Total geral: R$ {fmt_br(total_geral)}</b>", ParagraphStyle("tot", parent=styles["Normal"], fontSize=10, alignment=2, textColor=colors.HexColor("#2c5282"))))
    doc.build(story)


def main():
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    pdf_in = localizar_pdf_entrada()
    itens, grupos = gerar_csv(pdf_in)
    gerar_pdf()
    print(f"PDF de entrada: {pdf_in}")
    print(f"Itens: {itens} | Grupos: {grupos}")
    print(f"CSV: {OUT_CSV}")
    print(f"PDF: {OUT_PDF}")


if __name__ == "__main__":
    main()
