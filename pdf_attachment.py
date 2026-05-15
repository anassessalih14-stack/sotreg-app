from __future__ import annotations
import json
import sqlite3
from typing import Any, List, Optional, Tuple

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, NextPageTemplate
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def ensure_pdf_tables_schema(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS pdf_tables (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL,
      entity TEXT NOT NULL,
      provider TEXT NOT NULL,
      table_kind TEXT NOT NULL,
      title TEXT NOT NULL,
      table_json TEXT NOT NULL,
      updated_at TEXT DEFAULT (datetime('now'))
    );
    """)
    cur.execute("""
    CREATE INDEX IF NOT EXISTS idx_pdf_tables_lookup
    ON pdf_tables(period, entity, provider, table_kind);
    """)
    conn.commit()


def upsert_pdf_table(conn: sqlite3.Connection, period: str, entity: str, provider: str,
                     table_kind: str, title: str, columns: List[str], rows: List[List[Any]]) -> None:
    ensure_pdf_tables_schema(conn)
    payload = json.dumps({"columns": columns, "rows": rows}, ensure_ascii=False)
    cur = conn.cursor()
    cur.execute("""
        DELETE FROM pdf_tables
        WHERE period=? AND entity=? AND provider=? AND table_kind=?
    """, (period, entity, provider, table_kind))
    cur.execute("""
        INSERT INTO pdf_tables(period, entity, provider, table_kind, title, table_json)
        VALUES(?,?,?,?,?,?)
    """, (period, entity, provider, table_kind, title, payload))
    conn.commit()


def fetch_pdf_table(conn: sqlite3.Connection, period: str, entity: str, provider: str, table_kind: str
                   ) -> Optional[Tuple[str, List[str], List[List[Any]]]]:
    ensure_pdf_tables_schema(conn)
    cur = conn.cursor()
    cur.execute("""
        SELECT title, table_json
        FROM pdf_tables
        WHERE period=? AND entity=? AND provider=? AND table_kind=?
        LIMIT 1
    """, (period, entity, provider, table_kind))
    row = cur.fetchone()
    if not row:
        return None
    title, table_json = row
    data = json.loads(table_json)
    return title, data["columns"], data["rows"]


def _base_style(font_size: int = 8, header_bg=colors.lightgrey) -> TableStyle:
    return TableStyle([
        ("BACKGROUND", (0,0), (-1,0), header_bg),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), font_size),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.lightyellow]),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])


def rl_table(columns: List[str], rows: List[List[Any]], *, font_size: int = 8, col_widths=None) -> Table:
    data = [columns] + [[("" if v is None else str(v)) for v in r] for r in rows]
    t = Table(data, repeatRows=1, colWidths=col_widths)
    t.setStyle(_base_style(font_size))
    return t


def add_signature_block(entity: str) -> Table:
    left_title = f"Responsable de l'entité ({entity})"
    right_title = "Le Chef d'exploitation"
    left_body = "\n\n\nSignature : ____________________"
    right_body = "\n\n\nSignature : ____________________"
    data = [[left_title, right_title], [left_body, right_body]]
    t = Table(data, colWidths=[260, 260])
    t.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BOX", (0,0), (-1,-1), 0.8, colors.grey),
        ("INNERGRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    return t


def generate_attachment_pdf_from_db(db_path: str, period: str, entity: str, provider: str, out_path: str) -> str:
    styles = getSampleStyleSheet()
    conn = sqlite3.connect(db_path)
    facture = fetch_pdf_table(conn, period, entity, provider, "FACTURE")
    detail = fetch_pdf_table(conn, period, entity, provider, "DETAIL")
    tarifs = fetch_pdf_table(conn, period, entity, provider, "TARIFS")
    conn.close()

    if not facture or not detail:
        missing = []
        if not facture:
            missing.append("FACTURE")
        if not detail:
            missing.append("DETAIL")
        raise FileNotFoundError(f"Table(s) manquante(s) en base: {', '.join(missing)}")


    facture_title, facture_cols, facture_rows = facture
    detail_title, detail_cols, detail_rows = detail

    # BaseDocTemplate with two page templates (portrait then landscape)
    doc = BaseDocTemplate(out_path, pagesize=A4, leftMargin=12, rightMargin=12, topMargin=14, bottomMargin=14)

    frame_portrait = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='F1')
    frame_land = Frame(12, 12, landscape(A4)[0]-24, landscape(A4)[1]-24, id='F2')

    pt_portrait = PageTemplate(id='PORTRAIT', frames=[frame_portrait], pagesize=A4)
    pt_land = PageTemplate(id='LANDSCAPE', frames=[frame_land], pagesize=landscape(A4))

    doc.addPageTemplates([pt_portrait, pt_land])

    elements = []

    # Page 1 (portrait)
    elements.append(Paragraph("Attachement de Transport du Personnel", styles["Title"]))
    elements.append(Paragraph(f"<b>Entité:</b> {entity} &nbsp;&nbsp; <b>Prestataire:</b> {provider} &nbsp;&nbsp; <b>Période:</b> {period}", styles["Normal"]))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(facture_title, styles["Heading2"]))
    elements.append(Spacer(1, 6))

    t1 = rl_table(facture_cols, facture_rows, font_size=9, col_widths=[210, 260, 90])
    nrows = len(facture_rows) + 1
    if nrows >= 5:
        start = nrows - 4
        t1.setStyle(TableStyle([
            ("FONTNAME", (0, start), (-1, nrows-1), "Helvetica-Bold"),
            ("BACKGROUND", (0, start), (-1, nrows-1), colors.beige),
        ]))
    elements.append(t1)
    elements.append(Spacer(1, 14))
    elements.append(add_signature_block(entity))

    # Switch to landscape for page 2
    elements.append(NextPageTemplate('LANDSCAPE'))
    elements.append(PageBreak())

    # Page 2 (landscape)
    elements.append(Paragraph(detail_title, styles["Heading2"]))
    elements.append(Spacer(1, 6))

    t2 = rl_table(detail_cols, detail_rows, font_size=6)
    t2.setStyle(TableStyle([
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING", (0,0), (-1,-1), 1),
        ("BOTTOMPADDING", (0,0), (-1,-1), 1),
    ]))
    elements.append(t2)

    if tarifs:
        tarifs_title, tarifs_cols, tarifs_rows = tarifs
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(tarifs_title, styles["Heading3"]))
        elements.append(Spacer(1, 4))
        elements.append(rl_table(tarifs_cols, tarifs_rows, font_size=7))

    doc.build(elements)
    return out_path
