# -*- coding: utf-8 -*-
"""Gera PDF do relatório agregado a partir do CSV."""
import csv
from collections import OrderedDict
from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

CSV_PATH = Path(r"c:\tmp\pedidos pr\relatorio_agregado.csv")
OUT_PDF = Path(r"c:\tmp\pedidos pr\relatorio_agregado_v2.pdf")


def esc_xml(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse_br_decimal(s: str) -> float:
    s = (s or "").strip().replace(".", "").replace(",", ".")
    return float(s) if s else 0.0


def fmt_br(val: float) -> str:
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def peso_para_kg(peso: str) -> float:
    p = (peso or "").strip().lower()
    if not p:
        return 0.0
    import re
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


def load_rows():
    data_rows = []
    pedidos_por_vendedor = {}
    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f, delimiter=";")
        _header = next(r, None)
        for row in r:
            if not row:
                continue
            if len(row) >= 4 and row[1] == "__PEDIDOS__":
                pedidos_por_vendedor[row[0]] = row[3]
            elif len(row) >= 2 and row[1] == "__SUBTOTAL__":
                continue
            elif len(row) >= 2 and row[1] == "__TOTAL_PESO__":
                continue
            elif len(row) >= 8:
                data_rows.append(row)
    return data_rows, pedidos_por_vendedor


def build_pdf():
    data_rows, pedidos_por_vendedor = load_rows()

    total_geral = sum(parse_br_decimal(r[7]) for r in data_rows)
    peso_total_kg = sum(peso_para_kg(r[4]) * parse_br_decimal(r[6]) for r in data_rows)
    ordem_vendedores = list(OrderedDict.fromkeys(row[0] for row in data_rows))

    page_size = landscape(A4)
    doc = SimpleDocTemplate(
        str(OUT_PDF),
        pagesize=page_size,
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontSize=14, alignment=1, spaceAfter=5)
    sub_style = ParagraphStyle("S", parent=styles["Normal"], fontSize=9, alignment=1, textColor=colors.HexColor("#444444"))
    ped_style = ParagraphStyle("P", parent=styles["Normal"], fontSize=8, leftIndent=3 * mm, spaceBefore=2, spaceAfter=8, textColor=colors.HexColor("#4a5568"))
    subhead_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=10, textColor=colors.HexColor("#2c5282"), spaceBefore=5, spaceAfter=4)

    story = [
        Paragraph("Relação entre relatórios por código do produto", title_style),
        Paragraph(
            f"Gerado em {date.today().strftime('%d/%m/%Y')} · Agrupamento: vendedor, código, barras, descrição, peso, unidade",
            sub_style,
        ),
        Spacer(1, 6),
    ]

    hdr = ["Código", "Barras", "Descrição", "Peso", "Un.", "Σ Qtd", "Σ Total (R$)"]
    col_widths = [
        16 * mm,
        24 * mm,
        page_size[0] - 20 * mm - (16 + 24 + 16 + 12 + 17 + 20) * mm,
        16 * mm,
        12 * mm,
        17 * mm,
        20 * mm,
    ]

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
            table_data.append([
                c,
                b,
                Paragraph(esc_xml(d), ParagraphStyle("c", parent=styles["Normal"], fontSize=7, leading=9)),
                p,
                u,
                q,
                t,
            ])

        sub_q = sum(parse_br_decimal(r[6]) for r in rows_v)
        sub_t = sum(parse_br_decimal(r[7]) for r in rows_v)
        sub_peso_kg = sum(peso_para_kg(r[4]) * parse_br_decimal(r[6]) for r in rows_v)
        table_data.append([
            "",
            "",
            Paragraph("<b>Subtotal do vendedor</b>", ParagraphStyle("sub", parent=styles["Normal"], fontSize=8)),
            Paragraph(f"<b>{fmt_br(sub_peso_kg)} kg</b>", ParagraphStyle("sp", parent=styles["Normal"], fontSize=8)),
            "",
            Paragraph(f"<b>{fmt_br(sub_q)}</b>", ParagraphStyle("sq", parent=styles["Normal"], fontSize=8)),
            Paragraph(f"<b>{fmt_br(sub_t)}</b>", ParagraphStyle("st", parent=styles["Normal"], fontSize=8)),
        ])

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
    story.append(
        Paragraph(
            f"<b>Peso total estimado: {fmt_br(peso_total_kg)} kg</b>",
            ParagraphStyle("tot_peso", parent=styles["Normal"], fontSize=10, alignment=2, textColor=colors.HexColor("#2c5282")),
        )
    )
    story.append(Paragraph(f"<b>Total geral: R$ {fmt_br(total_geral)}</b>", ParagraphStyle("tot", parent=styles["Normal"], fontSize=10, alignment=2, textColor=colors.HexColor("#2c5282"))))

    doc.build(story)
    print(f"PDF gerado: {OUT_PDF}")


if __name__ == "__main__":
    build_pdf()
