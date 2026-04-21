"""CSV-prijslijst per leverancier: netto ophalen op leveranciersartikelcode."""

from __future__ import annotations

import csv
from pathlib import Path


def nk(k: str) -> str:
    return (k or "").strip().lower().replace(" ", "_")


def excel_column_to_index(letters: str) -> int:
    """Excel-kolomletter(s) naar 0-based index (A=0, B=1, …, AA=26)."""
    letters = (letters or "").strip().upper()
    if not letters:
        return 0
    n = 0
    for c in letters:
        if not ("A" <= c <= "Z"):
            raise ValueError(
                f"Ongeldige kolom: {letters!r} — gebruik alleen letters (A–Z, AA, …)."
            )
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n - 1


def parse_euro_cell(raw: str) -> float | None:
    s = (raw or "").strip()
    if not s:
        return None
    s = (
        s.replace("€", "")
        .replace("EUR", "")
        .replace("eur", "")
        .replace("\xa0", "")
        .replace(" ", "")
    )
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        v = float(s)
        return v if v == v else None
    except ValueError:
        return None


def _delimiter_from_sample(sample: str) -> str:
    first = sample.splitlines()[0] if sample.strip() else ""
    if not first:
        return ";"
    tabs = first.count("\t")
    semis = first.count(";")
    commas = first.count(",")
    if tabs >= max(semis, commas) and tabs > 0:
        return "\t"
    if semis >= commas and semis > 0:
        return ";"
    return ","


def read_price_list_csv(path: str | Path) -> list[dict[str, str]]:
    path = str(path)
    rows: list[dict] | None = None
    last_decode: str | None = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
                    delim = dialect.delimiter
                except csv.Error:
                    delim = _delimiter_from_sample(sample)
                rdr = csv.DictReader(f, delimiter=delim)
                rows = list(rdr)
            break
        except UnicodeDecodeError as e:
            last_decode = str(e)
            continue
    if rows is None:
        raise OSError(last_decode or "Kan CSV niet lezen (codering).")
    return rows


def read_price_list_csv_raw(path: str | Path) -> list[list[str]]:
    """Zelfde codering/delimiter als read_price_list_csv, maar als lijsten per rij."""
    path = str(path)
    out: list[list[str]] | None = None
    last_decode: str | None = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=";,\t")
                    delim = dialect.delimiter
                except csv.Error:
                    delim = _delimiter_from_sample(sample)
                rdr = csv.reader(f, delimiter=delim)
                out = list(rdr)
            break
        except UnicodeDecodeError as e:
            last_decode = str(e)
            continue
    if out is None:
        raise OSError(last_decode or "Kan CSV niet lezen (codering).")
    return out


def _extract_netto(lk: dict) -> float | None:
    keys = (
        "netto_prijs",
        "nettoprijs",
        "netto_prijzen",
        "netto",
        "purchase_price",
        "inkoop",
        "inkoopprijs",
        "inkoopsprijs",
        "netto_inkoop",
        "inkoop_prijs",
        "inkoop_netto",
        "prijs",
    )
    for key in keys:
        cell = lk.get(key)
        if cell is None or (isinstance(cell, str) and not cell.strip()):
            continue
        f = parse_euro_cell(str(cell))
        if f is not None:
            return f
    return None


ARTICLE_KEYS = (
    "supplier_article",
    "leverancier_artikel",
    "inkoopnummer",
    "artikelcode",
    "artikel_code",
    "code",
    "artikelnummer",
    "artikel_nr",
    "artikelnr",
    "artnr",
    "artikel",
    "art",
    "sku",
)


def _row_matches_article(lk: dict, want: str) -> bool:
    w = want.strip()
    if not w:
        return False
    for key in ARTICLE_KEYS:
        cell = lk.get(key)
        if cell is None:
            continue
        if str(cell).strip() == w:
            return True
    return False


def lookup_netto_price(rows: list[dict], supplier_article: str) -> float | None:
    want = (supplier_article or "").strip()
    if not want:
        return None
    for raw in rows:
        lk = {nk(k): (v or "").strip() for k, v in raw.items() if k}
        if not _row_matches_article(lk, want):
            continue
        netto = _extract_netto(lk)
        if netto is not None:
            return netto
    return None


def lookup_netto_by_excel_columns(
    rows: list[list[str]],
    supplier_article: str,
    col_article: str,
    col_netto: str,
    col_unit: str | None,
    col_description: str | None,
    skip_header_rows: int,
) -> tuple[float | None, str | None, str | None]:
    """Zoekt op artikelcode; optioneel eenheid en omschrijving offerte."""
    i_art = excel_column_to_index(col_article)
    i_net = excel_column_to_index(col_netto)
    i_unit = excel_column_to_index(col_unit) if (col_unit or "").strip() else None
    i_desc = (
        excel_column_to_index(col_description) if (col_description or "").strip() else None
    )
    want = (supplier_article or "").strip()
    if not want:
        return None, None, None

    skip = max(0, int(skip_header_rows))
    need = max(i_art, i_net)
    if i_unit is not None:
        need = max(need, i_unit)
    if i_desc is not None:
        need = max(need, i_desc)

    for row in rows[skip:]:
        if len(row) <= need:
            continue
        cell_art = str(row[i_art]).strip() if len(row) > i_art else ""
        if cell_art != want:
            continue
        netto = parse_euro_cell(row[i_net] if len(row) > i_net else "")
        unit_str: str | None = None
        if i_unit is not None and len(row) > i_unit:
            u = str(row[i_unit]).strip()
            unit_str = u if u else None
        desc_str: str | None = None
        if i_desc is not None and len(row) > i_desc:
            d = str(row[i_desc]).strip()
            desc_str = d if d else None
        return netto, unit_str, desc_str
    return None, None, None


def lookup_netto_price_from_file(
    path: str | Path,
    supplier_article: str,
    *,
    col_article: str | None = None,
    col_netto: str | None = None,
    col_unit: str | None = None,
    col_description: str | None = None,
    skip_header_rows: int | None = None,
) -> tuple[float | None, str | None, str | None]:
    """
    Leest prijslijst. Als col_article en col_netto zijn ingevuld (Excel-kolommen),
    wordt per positie gelezen; anders fallback op kolomnamen in de eerste regel.
    Retourneert (netto_bedrag, eenheid_uit_csv_of_None, omschrijving_offerte_of_None).
    """
    p = Path(path)
    if not p.is_file():
        return None, None, None

    use_cols = (
        (col_article or "").strip() != ""
        and (col_netto or "").strip() != ""
    )
    if use_cols:
        try:
            skip = 1 if skip_header_rows is None else int(skip_header_rows)
        except (TypeError, ValueError):
            skip = 1
        raw_rows = read_price_list_csv_raw(p)
        netto, unit, desc = lookup_netto_by_excel_columns(
            raw_rows,
            supplier_article,
            (col_article or "A").strip(),
            (col_netto or "E").strip(),
            (col_unit or "").strip() or None,
            (col_description or "").strip() or None,
            skip,
        )
        return netto, unit, desc

    rows = read_price_list_csv(p)
    n = lookup_netto_price(rows, supplier_article)
    return n, None, None
