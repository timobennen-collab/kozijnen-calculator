"""SQLite-database voor producten, stuklijsten en offertes."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "kozijnen_data.sqlite"


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def transaction():
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


DEFAULT_CATEGORIES: tuple[tuple[str, str, int], ...] = (
    ("hout", "Hout", 10),
    ("glas", "Glas", 20),
    ("hang_sluitwerk", "Hang en sluitwerk", 30),
    ("verf", "Verf", 40),
    ("uren_verbindingen", "Uren & verbindingen", 50),
)


def migrate_categories(conn: sqlite3.Connection) -> None:
    """Maakt categorieën aan en koppelt producten (bestaande DB's krijgen een kolom)."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    for code, name, sort_order in DEFAULT_CATEGORIES:
        conn.execute(
            "INSERT OR IGNORE INTO categories (code, name, sort_order) VALUES (?, ?, ?)",
            (code, name, sort_order),
        )
    cols = {row[1] for row in conn.execute("PRAGMA table_info(products)")}
    if "category_id" not in cols:
        conn.execute(
            "ALTER TABLE products ADD COLUMN category_id INTEGER REFERENCES categories (id)"
        )
    conn.execute(
        """
        UPDATE products SET category_id = (
            SELECT id FROM categories WHERE LOWER(code) = 'hout' LIMIT 1
        )
        WHERE category_id IS NULL
        """
    )
    cols_cat = {row[1] for row in conn.execute("PRAGMA table_info(categories)")}
    if "parent_id" not in cols_cat:
        conn.execute("ALTER TABLE categories ADD COLUMN parent_id INTEGER REFERENCES categories (id)")


def init_db() -> None:
    with get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL,
                unit TEXT NOT NULL DEFAULT 'stuk',
                unit_price REAL NOT NULL DEFAULT 0,
                quote_description TEXT DEFAULT '',
                price_source TEXT NOT NULL DEFAULT 'manual'
                    CHECK (price_source IN ('manual', 'bom'))
            );

            CREATE TABLE IF NOT EXISTS bom (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_id INTEGER NOT NULL,
                child_id INTEGER NOT NULL,
                qty REAL NOT NULL DEFAULT 1,
                FOREIGN KEY (parent_id) REFERENCES products (id) ON DELETE CASCADE,
                FOREIGN KEY (child_id) REFERENCES products (id) ON DELETE CASCADE,
                UNIQUE (parent_id, child_id)
            );

            CREATE TABLE IF NOT EXISTS quotes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                quote_number TEXT NOT NULL,
                customer_name TEXT NOT NULL DEFAULT '',
                customer_address TEXT NOT NULL DEFAULT '',
                customer_email TEXT NOT NULL DEFAULT '',
                notes TEXT NOT NULL DEFAULT '',
                vat_rate REAL NOT NULL DEFAULT 21
            );

            CREATE TABLE IF NOT EXISTS quote_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                qty REAL NOT NULL DEFAULT 1,
                unit_price_override REAL,
                line_discount_pct REAL NOT NULL DEFAULT 0,
                sort_order INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (quote_id) REFERENCES quotes (id) ON DELETE CASCADE,
                FOREIGN KEY (product_id) REFERENCES products (id)
            );

            CREATE INDEX IF NOT EXISTS idx_bom_parent ON bom (parent_id);
            CREATE INDEX IF NOT EXISTS idx_quote_lines_quote ON quote_lines (quote_id);
            """
        )
        migrate_categories(conn)
        migrate_dimension_rules(conn)
        migrate_product_extras(conn)
        migrate_category_root_features(conn)
        migrate_suppliers_and_product_extras(conn)


def migrate_product_extras(conn: sqlite3.Connection) -> None:
    """Optionele houtvelden (mm) op product."""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(products)")}
    if "hout_dikte_mm" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN hout_dikte_mm REAL")
    if "hout_breedte_mm" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN hout_breedte_mm REAL")
    if "hout_afkort_verlies_pct" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN hout_afkort_verlies_pct REAL")
    if "purchase_price" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN purchase_price REAL")
    if "margin_pct" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN margin_pct REAL")
    # Oude eenheid «m» (lopende meter) hernoemen naar «m1»
    conn.execute("UPDATE products SET unit = 'm1' WHERE LOWER(TRIM(unit)) = 'm'")


def migrate_dimension_rules(conn: sqlite3.Connection) -> None:
    """Maatregels: extra onderdelen per B×H-bereik (mm) bovenop de vaste stuklijst."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS dimension_bom_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_id INTEGER NOT NULL,
            child_id INTEGER NOT NULL,
            qty REAL NOT NULL DEFAULT 1,
            w_min REAL,
            w_max REAL,
            h_min REAL,
            h_max REAL,
            FOREIGN KEY (parent_id) REFERENCES products (id) ON DELETE CASCADE,
            FOREIGN KEY (child_id) REFERENCES products (id) ON DELETE CASCADE
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_dim_rules_parent ON dimension_bom_rules (parent_id)"
    )
    cols = {row[1] for row in conn.execute("PRAGMA table_info(quote_lines)")}
    if "width_mm" not in cols:
        conn.execute("ALTER TABLE quote_lines ADD COLUMN width_mm REAL")
    if "height_mm" not in cols:
        conn.execute("ALTER TABLE quote_lines ADD COLUMN height_mm REAL")


def category_id_by_code(conn: sqlite3.Connection, code: str) -> int:
    row = conn.execute("SELECT id FROM categories WHERE code = ?", (code,)).fetchone()
    if not row:
        raise ValueError(f"Onbekende categoriecode: {code}")
    return int(row[0])


# Opties per hoofdcategorie (subgroepen erven dit mee).
FEATURE_DIKTE_BREEDTE_MM = "dikte_breedte_mm"
FEATURE_AFKORT_VERLIES = "afkort_verlies"
FEATURE_NAAM_SUB_MAAT = "naam_sub_maat"
FEATURE_OPPERVLAKTE_M2 = "oppervlakte_m2_minmax"
FEATURE_DIKTE_LOS_MM = "dikte_mm_los"
FEATURE_LEVERANCIER = "leverancier_velden"
FEATURE_UREN_MINUTEN = "uren_minuten"

ROOT_FEATURE_LABELS: dict[str, str] = {
    FEATURE_DIKTE_BREEDTE_MM: "Afmetingen dikte × breedte (mm)",
    FEATURE_AFKORT_VERLIES: "Afkortverlies (%) op inkoop (in combinatie met marge)",
    FEATURE_NAAM_SUB_MAAT: "Productnaam automatisch: subcategorie + dikte×breedte",
    FEATURE_OPPERVLAKTE_M2: "Min./max. oppervlakte (m²) op product",
    FEATURE_DIKTE_LOS_MM: "Afmeting dikte alleen (mm), los van dikte×breedte",
    FEATURE_LEVERANCIER: "Leverancier + inkoopartikelnummer",
    FEATURE_UREN_MINUTEN: "Werkminuten totaal (bv. 120 = 2 uur; opgeslagen als decimale uren)",
}


def migrate_category_root_features(conn: sqlite3.Connection) -> None:
    """Tabel met welke extra velden per hoofdcategorie gelden."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS category_root_features (
            root_category_id INTEGER NOT NULL REFERENCES categories(id) ON DELETE CASCADE,
            feature_code TEXT NOT NULL,
            PRIMARY KEY (root_category_id, feature_code)
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_cat_root_feat_code ON category_root_features(feature_code)"
    )
    row = conn.execute(
        "SELECT id FROM categories WHERE LOWER(code) = 'hout' LIMIT 1"
    ).fetchone()
    if row:
        hid = int(row[0])
        for fc in (
            FEATURE_DIKTE_BREEDTE_MM,
            FEATURE_AFKORT_VERLIES,
            FEATURE_NAAM_SUB_MAAT,
        ):
            conn.execute(
                "INSERT OR IGNORE INTO category_root_features (root_category_id, feature_code) VALUES (?, ?)",
                (hid, fc),
            )
    row_hs = conn.execute(
        "SELECT id FROM categories WHERE LOWER(code) = 'hang_sluitwerk' LIMIT 1"
    ).fetchone()
    if row_hs:
        hs_id = int(row_hs[0])
        conn.execute(
            "INSERT OR IGNORE INTO category_root_features (root_category_id, feature_code) VALUES (?, ?)",
            (hs_id, FEATURE_LEVERANCIER),
        )


def get_root_category_id(conn: sqlite3.Connection, category_id: int) -> int | None:
    """Loopt omhoog tot de hoofdcategorie (parent_id IS NULL)."""
    cid = category_id
    for _ in range(10000):
        row = conn.execute(
            "SELECT parent_id FROM categories WHERE id = ?", (cid,)
        ).fetchone()
        if not row:
            return None
        if row["parent_id"] is None:
            return cid
        cid = int(row["parent_id"])
    return None


def root_has_feature(conn: sqlite3.Connection, root_id: int, feature_code: str) -> bool:
    r = conn.execute(
        """
        SELECT 1 FROM category_root_features
        WHERE root_category_id = ? AND feature_code = ?
        """,
        (root_id, feature_code),
    ).fetchone()
    return r is not None


def category_root_has_feature(
    conn: sqlite3.Connection, category_id: int, feature_code: str
) -> bool:
    rid = get_root_category_id(conn, category_id)
    if rid is None:
        return False
    return root_has_feature(conn, rid, feature_code)


def category_ids_in_roots_with_any_feature(
    conn: sqlite3.Connection, *feature_codes: str
) -> set[int]:
    """Alle categorie-id's in subtrees van hoofden die minstens één feature hebben."""
    if not feature_codes:
        return set()
    ph = ",".join("?" * len(feature_codes))
    roots = conn.execute(
        f"""
        SELECT DISTINCT root_category_id FROM category_root_features
        WHERE feature_code IN ({ph})
        """,
        feature_codes,
    ).fetchall()
    out: set[int] = set()
    for r in roots:
        out.update(category_subtree_ids(conn, int(r[0])))
    return out


def is_direct_sub_of_root_with_feature(
    conn: sqlite3.Connection, category_id: int, feature_code: str
) -> bool:
    """True als category_id een directe subcategorie is van een hoofdcategorie met feature."""
    row = conn.execute(
        "SELECT parent_id FROM categories WHERE id = ?", (category_id,)
    ).fetchone()
    if not row or row["parent_id"] is None:
        return False
    root_id = int(row["parent_id"])
    pro = conn.execute(
        "SELECT parent_id FROM categories WHERE id = ?", (root_id,)
    ).fetchone()
    if not pro or pro["parent_id"] is not None:
        return False
    return root_has_feature(conn, root_id, feature_code)


def category_subtree_ids(conn: sqlite3.Connection, root_id: int) -> list[int]:
    """Alle categorie-id's: root zelf plus alle nakomelingen (recursief)."""
    rows = conn.execute(
        """
        WITH RECURSIVE sub(id) AS (
            SELECT id FROM categories WHERE id = ?
            UNION ALL
            SELECT c.id FROM categories c JOIN sub ON c.parent_id = sub.id
        )
        SELECT id FROM sub
        """,
        (root_id,),
    ).fetchall()
    return [int(r[0]) for r in rows]


def hout_category_ids(conn: sqlite3.Connection) -> set[int]:
    """Alle categorie-id's onder hoofdcategorieën met minstens één product-extra-feature."""
    return category_ids_in_roots_with_any_feature(
        conn,
        FEATURE_DIKTE_BREEDTE_MM,
        FEATURE_AFKORT_VERLIES,
        FEATURE_NAAM_SUB_MAAT,
        FEATURE_OPPERVLAKTE_M2,
        FEATURE_DIKTE_LOS_MM,
        FEATURE_LEVERANCIER,
        FEATURE_UREN_MINUTEN,
    )


def migrate_suppliers_and_product_extras(conn: sqlite3.Connection) -> None:
    """Leveranciers + extra productkolommen."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT NOT NULL,
            notes TEXT NOT NULL DEFAULT '',
            price_list_path TEXT,
            csv_col_article TEXT DEFAULT 'A',
            csv_col_netto TEXT DEFAULT 'E',
            csv_col_unit TEXT DEFAULT '',
            csv_col_description TEXT DEFAULT '',
            csv_skip_header_rows INTEGER DEFAULT 1
        )
        """
    )
    scols = {row[1] for row in conn.execute("PRAGMA table_info(suppliers)")}
    if "price_list_path" not in scols:
        conn.execute("ALTER TABLE suppliers ADD COLUMN price_list_path TEXT")
    if "csv_col_article" not in scols:
        conn.execute(
            "ALTER TABLE suppliers ADD COLUMN csv_col_article TEXT DEFAULT 'A'"
        )
    if "csv_col_netto" not in scols:
        conn.execute("ALTER TABLE suppliers ADD COLUMN csv_col_netto TEXT DEFAULT 'E'")
    if "csv_col_unit" not in scols:
        conn.execute("ALTER TABLE suppliers ADD COLUMN csv_col_unit TEXT DEFAULT ''")
    if "csv_skip_header_rows" not in scols:
        conn.execute(
            "ALTER TABLE suppliers ADD COLUMN csv_skip_header_rows INTEGER DEFAULT 1"
        )
    if "csv_col_description" not in scols:
        conn.execute(
            "ALTER TABLE suppliers ADD COLUMN csv_col_description TEXT DEFAULT ''"
        )
    cols = {row[1] for row in conn.execute("PRAGMA table_info(products)")}
    if "min_oppervlakte_m2" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN min_oppervlakte_m2 REAL")
    if "max_oppervlakte_m2" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN max_oppervlakte_m2 REAL")
    if "product_dikte_mm" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN product_dikte_mm REAL")
    if "supplier_id" not in cols:
        conn.execute(
            """
            ALTER TABLE products ADD COLUMN supplier_id INTEGER
                REFERENCES suppliers(id) ON DELETE SET NULL
            """
        )
    if "supplier_article" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN supplier_article TEXT")
    if "work_hours" not in cols:
        conn.execute("ALTER TABLE products ADD COLUMN work_hours REAL")


def seed_demo_data() -> None:
    """Voorbeeldartikelen als de database leeg is."""
    with get_connection() as conn:
        n = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        if n > 0:
            return

    # Losse onderdelen (categorie per artikel)
    parts = [
        ("SLPL-01", "Sluitplaat standaard", "stuk", 12.50, "Sluitplaat voor meerpuntsluiting", "hang_sluitwerk"),
        ("KOM-01", "Kom standaard", "stuk", 8.00, "Kom voor meerpuntsluiting", "hang_sluitwerk"),
        ("GLAS-1", "Glas isolatie 4-16-4", "m²", 85.00, "HR++ glas per m²", "glas"),
    ]
    with transaction() as conn:
        for code, name, unit, price, desc, cat in parts:
            cid = category_id_by_code(conn, cat)
            conn.execute(
                """INSERT INTO products (code, name, unit, unit_price, quote_description, price_source, category_id)
                   VALUES (?, ?, ?, ?, ?, 'manual', ?)""",
                (code, name, unit, price, desc, cid),
            )

        cid_hs = category_id_by_code(conn, "hang_sluitwerk")
        conn.execute(
            """INSERT INTO products (code, name, unit, unit_price, quote_description, price_source, category_id)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                "MP-STD",
                "Meerpuntsluiting compleet",
                "stuk",
                0,
                "Complete meerpuntsluiting inclusief sluitplaten en kommen, montage volgens fabrieksvoorschrift.",
                "bom",
                cid_hs,
            ),
        )
        parent = conn.execute(
            "SELECT id FROM products WHERE code = ?", ("MP-STD",)
        ).fetchone()[0]
        c1 = conn.execute(
            "SELECT id FROM products WHERE code = ?", ("SLPL-01",)
        ).fetchone()[0]
        c2 = conn.execute(
            "SELECT id FROM products WHERE code = ?", ("KOM-01",)
        ).fetchone()[0]
        conn.execute(
            "INSERT INTO bom (parent_id, child_id, qty) VALUES (?, ?, ?)",
            (parent, c1, 2),
        )
        conn.execute(
            "INSERT INTO bom (parent_id, child_id, qty) VALUES (?, ?, ?)",
            (parent, c2, 2),
        )


def next_quote_number() -> str:
    """Genereert OFF-YYYYMMDD-001 stijl nummer."""
    today = datetime.now().strftime("%Y%m%d")
    prefix = f"OFF-{today}-"
    with get_connection() as conn:
        row = conn.execute(
            "SELECT quote_number FROM quotes WHERE quote_number LIKE ? ORDER BY id DESC LIMIT 1",
            (prefix + "%",),
        ).fetchone()
    if not row:
        return prefix + "001"
    last = row["quote_number"]
    try:
        n = int(last.split("-")[-1]) + 1
    except ValueError:
        n = 1
    return prefix + f"{n:03d}"
