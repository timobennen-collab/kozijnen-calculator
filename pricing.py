"""Prijsberekening en BOM-uitrol (zonder circulaire verwijzingen)."""

from __future__ import annotations

from dataclasses import dataclass
from db import get_connection


@dataclass
class ExpandedLine:
    """Eén regel voor interne materiaallijst (uitgerold)."""

    product_id: int
    code: str
    name: str
    unit: str
    qty: float
    unit_price: float
    line_total: float


def _load_products_map():
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM products").fetchall()
    return {r["id"]: dict(r) for r in rows}


def _load_bom_map():
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT parent_id, child_id, qty FROM bom"
        ).fetchall()
    bom: dict[int, list[tuple[int, float]]] = {}
    for r in rows:
        bom.setdefault(r["parent_id"], []).append((r["child_id"], float(r["qty"])))
    return bom


def effective_unit_price(product_id: int, ancestors: set[int] | None = None) -> float:
    """Effectieve verkoopprijs per eenheid: bij stuklijst (BOM) som van onderdelen, anders unit_price."""
    if ancestors is None:
        ancestors = set()
    if product_id in ancestors:
        raise ValueError(f"Cirkel in stuklijst rond product-id {product_id}")

    products = _load_products_map()
    bom = _load_bom_map()
    p = products[product_id]
    if bom.get(product_id):
        next_anc = ancestors | {product_id}
        total = 0.0
        for child_id, qty in bom[product_id]:
            total += qty * effective_unit_price(child_id, next_anc)
        return total
    return float(p["unit_price"])


def _fetch_matching_dimension_rules(parent_id: int, w_mm: float, h_mm: float) -> list[tuple[int, float]]:
    """Alle regels waar (B,H) in het opgegeven bereik valt. Leeg veld = geen grens."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT child_id, qty FROM dimension_bom_rules
            WHERE parent_id = ?
              AND (w_min IS NULL OR w_min <= ?)
              AND (w_max IS NULL OR w_max >= ?)
              AND (h_min IS NULL OR h_min <= ?)
              AND (h_max IS NULL OR h_max >= ?)
            """,
            (parent_id, w_mm, w_mm, h_mm, h_mm),
        ).fetchall()
    return [(int(r["child_id"]), float(r["qty"])) for r in rows]


def dimension_extra_unit_price(parent_id: int, w_mm: float, h_mm: float) -> float:
    """Extra prijs per stuk door maatregels (som van kind × prijs × regel-aantal)."""
    total = 0.0
    for child_id, qty in _fetch_matching_dimension_rules(parent_id, w_mm, h_mm):
        total += qty * effective_unit_price(child_id)
    return total


def quote_line_unit_price(
    product_id: int,
    override: float | None,
    width_mm: float | None = None,
    height_mm: float | None = None,
) -> float:
    if override is not None:
        return float(override)
    base = effective_unit_price(product_id)
    if width_mm is None or height_mm is None:
        return base
    return base + dimension_extra_unit_price(product_id, float(width_mm), float(height_mm))


def expand_bom_for_picking(
    product_id: int,
    multiplier: float = 1.0,
    ancestors: set[int] | None = None,
) -> list[ExpandedLine]:
    """Rolt een samengesteld product uit naar onderdelen voor magazijn/picking."""
    if ancestors is None:
        ancestors = set()
    if product_id in ancestors:
        raise ValueError("Cirkel in stuklijst")

    products = _load_products_map()
    bom = _load_bom_map()
    children = bom.get(product_id)

    if not children:
        p = products[product_id]
        price = effective_unit_price(product_id)
        q = multiplier
        return [
            ExpandedLine(
                product_id=product_id,
                code=p["code"],
                name=p["name"],
                unit=p["unit"],
                qty=q,
                unit_price=price,
                line_total=round(q * price, 2),
            )
        ]

    next_anc = ancestors | {product_id}
    out: list[ExpandedLine] = []
    for child_id, qty in children:
        out.extend(
            expand_bom_for_picking(child_id, multiplier * qty, next_anc)
        )
    return out


def merge_expanded_lines(lines: list[ExpandedLine]) -> list[ExpandedLine]:
    """Voegt regels voor hetzelfde artikel samen (som aantallen)."""
    if not lines:
        return []
    products = _load_products_map()
    acc: dict[int, float] = {}
    for ln in lines:
        acc[ln.product_id] = acc.get(ln.product_id, 0.0) + ln.qty
    out: list[ExpandedLine] = []
    for pid in sorted(acc.keys(), key=lambda i: products[i]["code"]):
        q = acc[pid]
        p = products[pid]
        price = effective_unit_price(pid)
        out.append(
            ExpandedLine(
                product_id=pid,
                code=p["code"],
                name=p["name"],
                unit=p["unit"],
                qty=q,
                unit_price=price,
                line_total=round(q * price, 2),
            )
        )
    return out


def expand_bom_with_dimensions(
    product_id: int,
    multiplier: float = 1.0,
    width_mm: float | None = None,
    height_mm: float | None = None,
    ancestors: set[int] | None = None,
) -> list[ExpandedLine]:
    """Vaste BOM + maatregels (extra onderdelen per B×H), uitgerold naar onderdelen."""
    base = expand_bom_for_picking(product_id, multiplier, ancestors)
    if width_mm is None or height_mm is None:
        return base
    rules = _fetch_matching_dimension_rules(product_id, float(width_mm), float(height_mm))
    extra: list[ExpandedLine] = []
    for child_id, qty in rules:
        extra.extend(expand_bom_for_picking(child_id, multiplier * qty))
    return merge_expanded_lines(base + extra)


def validate_bom_acyclic() -> list[str]:
    """Controleert alle ouders op circulaire BOM. Retourneert foutmeldingen."""
    with get_connection() as conn:
        parents = [r[0] for r in conn.execute("SELECT DISTINCT parent_id FROM bom")]

    errors: list[str] = []
    for pid in parents:
        try:
            effective_unit_price(pid)
        except ValueError as e:
            errors.append(str(e))
    return errors
