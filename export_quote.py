"""Exporteert een offerte naar HTML (printen naar PDF via browser)."""

from __future__ import annotations

import html
from pathlib import Path

from db import get_connection
from pricing import quote_line_unit_price


def load_quote(quote_id: int) -> dict:
    with get_connection() as conn:
        q = conn.execute("SELECT * FROM quotes WHERE id = ?", (quote_id,)).fetchone()
        if not q:
            raise ValueError("Offerte niet gevonden")
        lines = conn.execute(
            """
            SELECT ql.*, p.code, p.name, p.unit, p.quote_description
            FROM quote_lines ql
            JOIN products p ON p.id = ql.product_id
            WHERE ql.quote_id = ?
            ORDER BY ql.sort_order, ql.id
            """,
            (quote_id,),
        ).fetchall()
    return {"quote": dict(q), "lines": [dict(r) for r in lines]}


def render_html(quote_id: int) -> str:
    data = load_quote(quote_id)
    q = data["quote"]
    lines = data["lines"]
    vat = float(q["vat_rate"]) / 100.0

    body_rows: list[str] = []
    subtotal = 0.0
    for row in lines:
        wm = row["width_mm"] if "width_mm" in row.keys() else None
        hm = row["height_mm"] if "height_mm" in row.keys() else None
        unit = quote_line_unit_price(
            row["product_id"], row["unit_price_override"], wm, hm
        )
        qty = float(row["qty"])
        disc = float(row["line_discount_pct"]) / 100.0
        line_net = unit * qty * (1.0 - disc)
        subtotal += line_net
        desc = (row["quote_description"] or "").strip()
        desc_html = (
            f'<div class="desc">{html.escape(desc)}</div>' if desc else ""
        )
        maat_note = ""
        if wm is not None and hm is not None:
            maat_note = (
                f'<div class="maat">Maat: {wm:g} × {hm:g} mm (voor maat-afhankelijke onderdelen)</div>'
            )
        body_rows.append(
            f"""
            <tr>
              <td class="num">{qty:g}</td>
              <td class="unit">{html.escape(row["unit"] or "")}</td>
              <td class="name">
                <strong>{html.escape(row["name"])}</strong>
                <span class="code">({html.escape(row["code"])})</span>
                {maat_note}
                {desc_html}
              </td>
              <td class="money">{unit:,.2f}</td>
              <td class="money">{line_net:,.2f}</td>
            </tr>
            """
        )

    vat_amount = subtotal * vat
    total = subtotal + vat_amount

    title = html.escape(q["quote_number"])
    customer_block = f"""
    <div class="customer">
      <strong>{html.escape(q["customer_name"] or "Klant")}</strong><br/>
      {html.escape(q["customer_address"]).replace(chr(10), "<br/>")}<br/>
      {html.escape(q["customer_email"])}
    </div>
    """

    notes = (q["notes"] or "").strip()
    notes_html = (
        f'<div class="notes"><strong>Opmerkingen</strong><br/>{html.escape(notes).replace(chr(10), "<br/>")}</div>'
        if notes
        else ""
    )

    return f"""<!DOCTYPE html>
<html lang="nl">
<head>
  <meta charset="utf-8"/>
  <title>{title}</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 40px; color: #222; }}
    h1 {{ font-size: 22px; margin-bottom: 8px; }}
    .meta {{ color: #555; margin-bottom: 24px; font-size: 14px; }}
    .customer {{ margin-bottom: 28px; line-height: 1.5; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ border-bottom: 1px solid #ddd; padding: 10px 8px; vertical-align: top; }}
    th {{ text-align: left; background: #f5f5f5; }}
    .num {{ width: 70px; text-align: right; }}
    .unit {{ width: 60px; }}
    .money {{ text-align: right; white-space: nowrap; width: 100px; }}
    .code {{ color: #666; font-size: 12px; margin-left: 6px; }}
    .desc {{ margin-top: 6px; color: #444; font-size: 13px; line-height: 1.4; }}
    .maat {{ margin-top: 4px; color: #555; font-size: 12px; }}
    .totals {{ margin-top: 20px; float: right; width: 320px; font-size: 14px; }}
    .totals tr td {{ border: none; padding: 4px 0; }}
    .totals .label {{ text-align: right; padding-right: 16px; color: #555; }}
    .totals .sum {{ font-weight: bold; font-size: 16px; }}
    .notes {{ margin-top: 120px; clear: both; font-size: 13px; color: #444; }}
    @media print {{
      body {{ margin: 12mm; }}
      a {{ display: none; }}
    }}
  </style>
</head>
<body>
  <h1>Offerte {html.escape(q["quote_number"])}</h1>
  <div class="meta">Datum: {html.escape(q["created_at"][:10])}</div>
  {customer_block}
  <table>
    <thead>
      <tr>
        <th class="num">Aantal</th>
        <th class="unit">Eenheid</th>
        <th>Omschrijving</th>
        <th class="money">Prijs / eenheid</th>
        <th class="money">Regel</th>
      </tr>
    </thead>
    <tbody>
      {''.join(body_rows)}
    </tbody>
  </table>
  <table class="totals">
    <tr><td class="label">Subtotaal excl. BTW</td><td class="money">{subtotal:,.2f}</td></tr>
    <tr><td class="label">BTW ({q["vat_rate"]:g}%)</td><td class="money">{vat_amount:,.2f}</td></tr>
    <tr><td class="label sum">Totaal incl. BTW</td><td class="money sum">{total:,.2f}</td></tr>
  </table>
  {notes_html}
  <p style="margin-top:48px;font-size:12px;color:#888;">Afgedrukt vanuit Kozijnen calculator</p>
</body>
</html>
"""


def export_quote_html(quote_id: int, out_path: Path | None = None) -> Path:
    html_str = render_html(quote_id)
    if out_path is None:
        q = load_quote(quote_id)["quote"]
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in q["quote_number"])
        out_path = Path(__file__).resolve().parent / f"offerte_{safe}.html"
    out_path.write_text(html_str, encoding="utf-8")
    return out_path
