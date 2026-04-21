"""Tkinter-interface: producten, stuklijst, offerte + HTML-export."""

from __future__ import annotations

import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from db import (
    FEATURE_AFKORT_VERLIES,
    FEATURE_DIKTE_BREEDTE_MM,
    FEATURE_DIKTE_LOS_MM,
    FEATURE_LEVERANCIER,
    FEATURE_NAAM_SUB_MAAT,
    FEATURE_OPPERVLAKTE_M2,
    FEATURE_UREN_MINUTEN,
    ROOT_FEATURE_LABELS,
    category_root_has_feature,
    category_subtree_ids,
    get_connection,
    get_root_category_id,
    init_db,
    is_direct_sub_of_root_with_feature,
    next_quote_number,
    seed_demo_data,
    transaction,
)
from export_quote import export_quote_html
from supplier_pricelist import excel_column_to_index, lookup_netto_price_from_file
from pricing import (
    effective_unit_price,
    expand_bom_for_picking,
    expand_bom_with_dimensions,
    quote_line_unit_price,
    validate_bom_acyclic,
)


class KozijnenApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Kozijnen calculator — producten & offertes")
        self.geometry("1100x720")
        self.minsize(900, 600)

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self._frm_categories = ttk.Frame(nb)
        self._frm_suppliers = ttk.Frame(nb)
        self._frm_products = ttk.Frame(nb)
        self._frm_bom = ttk.Frame(nb)
        self._frm_dim = ttk.Frame(nb)
        self._frm_quotes = ttk.Frame(nb)
        self._frm_pick = ttk.Frame(nb)

        nb.add(self._frm_categories, text="Categorieën")
        nb.add(self._frm_suppliers, text="Leveranciers")
        nb.add(self._frm_products, text="Producten")
        nb.add(self._frm_bom, text="Stuklijst (BOM)")
        nb.add(self._frm_dim, text="Maatregels (B×H)")
        nb.add(self._frm_quotes, text="Offertes")
        nb.add(self._frm_pick, text="Materiaallijst")
        self._notebook = nb

        self._build_categories_tab()
        self._build_suppliers_tab()
        self._build_products_tab()
        self._build_bom_tab()
        self._build_dimension_rules_tab()
        self._build_quotes_tab()
        self._build_pick_tab()

        self.after(100, self._load_initial_data)
        self.after(200, self._refresh_quotes)

    def _load_initial_data(self) -> None:
        self._editing_category_id = None
        self._refresh_cat_parent_combo()
        self.combo_cat_parent.set("(Hoofdcategorie — geen ouder)")
        self._refresh_category_combos()
        self._refresh_categories_tree()
        self._refresh_products_tree()
        self._refresh_bom_parent_combo()
        self._refresh_dim_parent_combo()
        if self.combo_dim_parent["values"]:
            self.combo_dim_parent.current(0)
            self._on_dim_parent_change()
        self._refresh_suppliers_tree()
        self._refresh_supplier_combo()

    # --- Leveranciers ---
    def _build_suppliers_tab(self) -> None:
        hint = ttk.Label(
            self._frm_suppliers,
            text="Beheer leveranciers. Stel per leverancier het CSV-pad en Excel-kolommen in (artikelcode, netto, "
            "optioneel eenheid en omschrijving offerte). Op het product: «Netto ophalen» vult inkoop en eventueel "
            "eenheid en omschrijving op de offerte.",
            wraplength=900,
            justify=tk.LEFT,
        )
        hint.pack(anchor=tk.W, padx=8, pady=(8, 0))

        top = ttk.Frame(self._frm_suppliers)
        top.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree_suppliers = ttk.Treeview(
            top,
            columns=("id", "code", "name"),
            show="headings",
            height=14,
            selectmode="browse",
        )
        for c, t, w in [
            ("id", "ID", 50),
            ("code", "Code", 120),
            ("name", "Naam", 280),
        ]:
            self.tree_suppliers.heading(c, text=t)
            self.tree_suppliers.column(c, width=w)
        sb = ttk.Scrollbar(top, orient=tk.VERTICAL, command=self.tree_suppliers.yview)
        self.tree_suppliers.configure(yscrollcommand=sb.set)
        self.tree_suppliers.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_suppliers.bind("<<TreeviewSelect>>", self._on_supplier_select)

        form = ttk.LabelFrame(self._frm_suppliers, text="Leverancier bewerken")
        form.pack(fill=tk.X, padx=8, pady=(0, 8))

        r = 0
        ttk.Label(form, text="Code").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_sup_code = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_sup_code, width=24).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="Naam").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_sup_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_sup_name, width=48).grid(
            row=r, column=1, columnspan=2, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="Notities").grid(row=r, column=0, sticky=tk.NW, padx=4, pady=2)
        self.txt_sup_notes = tk.Text(form, width=60, height=3, wrap=tk.WORD)
        self.txt_sup_notes.grid(row=r, column=1, columnspan=2, sticky=tk.W, padx=4, pady=2)
        r += 1
        ttk.Label(form, text="Prijslijst (CSV)").grid(
            row=r, column=0, sticky=tk.NW, padx=4, pady=2
        )
        self.var_sup_price_list = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_sup_price_list, width=52).grid(
            row=r, column=1, columnspan=2, sticky=tk.W, padx=4, pady=2
        )
        ttk.Button(form, text="Bladeren…", command=self._supplier_browse_pricelist).grid(
            row=r, column=3, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="CSV-kolom artikelcode").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_sup_csv_col_art = tk.StringVar(value="A")
        ttk.Entry(form, textvariable=self.var_sup_csv_col_art, width=6).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        ttk.Label(form, text="CSV-kolom netto inkoop").grid(
            row=r, column=2, sticky=tk.W, padx=12, pady=2
        )
        self.var_sup_csv_col_net = tk.StringVar(value="E")
        ttk.Entry(form, textvariable=self.var_sup_csv_col_net, width=6).grid(
            row=r, column=3, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="CSV-kolom eenheid (optioneel)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_sup_csv_col_unit = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_sup_csv_col_unit, width=6).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        ttk.Label(form, text="Kopregels overslaan (aantal)").grid(
            row=r, column=2, sticky=tk.W, padx=12, pady=2
        )
        self.var_sup_csv_skip = tk.StringVar(value="1")
        ttk.Entry(form, textvariable=self.var_sup_csv_skip, width=6).grid(
            row=r, column=3, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="CSV-kolom omschrijving offerte (optioneel)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_sup_csv_col_desc = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_sup_csv_col_desc, width=6).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        ttk.Label(
            form,
            text="(tekst voor «Omschrijving op offerte» op het product)",
            foreground="#555",
        ).grid(row=r, column=2, columnspan=2, sticky=tk.W, padx=12, pady=2)
        r += 1

        btnf = ttk.Frame(form)
        btnf.grid(row=r, column=0, columnspan=4, pady=10, sticky=tk.W)
        ttk.Button(btnf, text="Nieuw", command=self._supplier_new).pack(side=tk.LEFT, padx=4)
        ttk.Button(btnf, text="Opslaan", command=self._supplier_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btnf, text="Verwijderen", command=self._supplier_delete).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btnf, text="Vernieuwen", command=self._refresh_suppliers_tree).pack(
            side=tk.LEFT, padx=4
        )

        self._editing_supplier_id: int | None = None

    def _refresh_suppliers_tree(self) -> None:
        for i in self.tree_suppliers.get_children():
            self.tree_suppliers.delete(i)
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM suppliers ORDER BY name"
            ).fetchall()
        for row in rows:
            self.tree_suppliers.insert(
                "",
                tk.END,
                values=(row["id"], row["code"] or "", row["name"]),
            )

    def _on_supplier_select(self, _evt=None) -> None:
        sel = self.tree_suppliers.selection()
        if not sel:
            return
        vals = self.tree_suppliers.item(sel[0], "values")
        sid = int(vals[0])
        self._editing_supplier_id = sid
        with get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM suppliers WHERE id = ?", (sid,)
            ).fetchone()
        if not row:
            return
        self.var_sup_code.set(row["code"] or "")
        self.var_sup_name.set(row["name"])
        self.txt_sup_notes.delete("1.0", tk.END)
        self.txt_sup_notes.insert("1.0", row["notes"] or "")
        if "price_list_path" in row.keys() and row["price_list_path"]:
            self.var_sup_price_list.set(row["price_list_path"])
        else:
            self.var_sup_price_list.set("")
        if "csv_col_article" in row.keys():
            self.var_sup_csv_col_art.set(row["csv_col_article"] or "A")
        else:
            self.var_sup_csv_col_art.set("A")
        if "csv_col_netto" in row.keys():
            self.var_sup_csv_col_net.set(row["csv_col_netto"] or "E")
        else:
            self.var_sup_csv_col_net.set("E")
        if "csv_col_unit" in row.keys():
            self.var_sup_csv_col_unit.set(row["csv_col_unit"] or "")
        else:
            self.var_sup_csv_col_unit.set("")
        if "csv_skip_header_rows" in row.keys() and row["csv_skip_header_rows"] is not None:
            self.var_sup_csv_skip.set(str(int(row["csv_skip_header_rows"])))
        else:
            self.var_sup_csv_skip.set("1")
        if "csv_col_description" in row.keys():
            self.var_sup_csv_col_desc.set(row["csv_col_description"] or "")
        else:
            self.var_sup_csv_col_desc.set("")

    def _supplier_browse_pricelist(self) -> None:
        path = filedialog.askopenfilename(
            title="Prijslijst (CSV)",
            filetypes=[("CSV", "*.csv *.CSV"), ("Alle bestanden", "*.*")],
        )
        if path:
            self.var_sup_price_list.set(path)

    def _supplier_new(self) -> None:
        self._editing_supplier_id = None
        self.var_sup_code.set("")
        self.var_sup_name.set("")
        self.txt_sup_notes.delete("1.0", tk.END)
        self.var_sup_price_list.set("")
        self.var_sup_csv_col_art.set("A")
        self.var_sup_csv_col_net.set("E")
        self.var_sup_csv_col_unit.set("")
        self.var_sup_csv_skip.set("1")
        self.var_sup_csv_col_desc.set("")

    def _supplier_save(self) -> None:
        code = self.var_sup_code.get().strip() or None
        name = self.var_sup_name.get().strip()
        if not name:
            messagebox.showwarning("Ontbrekend", "Vul minimaal de naam in.")
            return
        notes = self.txt_sup_notes.get("1.0", tk.END).strip()
        plist = self.var_sup_price_list.get().strip() or None
        ca = self.var_sup_csv_col_art.get().strip()
        cn = self.var_sup_csv_col_net.get().strip()
        cu = self.var_sup_csv_col_unit.get().strip()
        cd = self.var_sup_csv_col_desc.get().strip()
        try:
            if (ca and not cn) or (cn and not ca):
                messagebox.showwarning(
                    "CSV-kolommen",
                    "Vul zowel artikelcode- als netto-kolom in (bijv. A en E), "
                    "of laat beide leeg voor oude modus (kolomnamen in bestand).",
                )
                return
            if ca and cn:
                excel_column_to_index(ca)
                excel_column_to_index(cn)
            if cu:
                excel_column_to_index(cu)
            if cd:
                excel_column_to_index(cd)
            skip = int(self.var_sup_csv_skip.get().strip() or "0")
            if skip < 0:
                raise ValueError("negatief")
        except ValueError as e:
            messagebox.showerror(
                "CSV-kolommen",
                "Vul geldige Excel-kolommen in (bijv. A, E, AA) en een niet-negatief "
                "aantal kopregels.\n"
                f"({e})",
            )
            return
        try:
            with transaction() as conn:
                if self._editing_supplier_id is None:
                    conn.execute(
                        """
                        INSERT INTO suppliers (
                            code, name, notes, price_list_path,
                            csv_col_article, csv_col_netto, csv_col_unit, csv_col_description,
                            csv_skip_header_rows
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            code,
                            name,
                            notes,
                            plist,
                            ca or None,
                            cn or None,
                            cu or None,
                            cd or None,
                            skip,
                        ),
                    )
                else:
                    conn.execute(
                        """
                        UPDATE suppliers SET
                            code = ?, name = ?, notes = ?, price_list_path = ?,
                            csv_col_article = ?, csv_col_netto = ?, csv_col_unit = ?,
                            csv_col_description = ?, csv_skip_header_rows = ?
                        WHERE id = ?
                        """,
                        (
                            code,
                            name,
                            notes,
                            plist,
                            ca or None,
                            cn or None,
                            cu or None,
                            cd or None,
                            skip,
                            self._editing_supplier_id,
                        ),
                    )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return
        self._refresh_suppliers_tree()
        self._refresh_supplier_combo()
        messagebox.showinfo("Opgeslagen", "Leverancier opgeslagen.")

    def _supplier_delete(self) -> None:
        if self._editing_supplier_id is None:
            messagebox.showinfo("Geen selectie", "Selecteer eerst een leverancier.")
            return
        if not messagebox.askyesno("Verwijderen", "Leverancier verwijderen?"):
            return
        try:
            with transaction() as conn:
                conn.execute(
                    "DELETE FROM suppliers WHERE id = ?", (self._editing_supplier_id,)
                )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return
        self._supplier_new()
        self._refresh_suppliers_tree()
        self._refresh_supplier_combo()

    def _refresh_supplier_combo(self, reset_selection: bool = True) -> None:
        if not hasattr(self, "combo_supplier"):
            return
        cur = self.combo_supplier.get()
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM suppliers ORDER BY name"
            ).fetchall()
        labels = ["(geen leverancier)"]
        self._supplier_label_to_id = {labels[0]: None}
        for row in rows:
            lab = f"{row['code'] + ' — ' if row['code'] else ''}{row['name']}"
            labels.append(lab)
            self._supplier_label_to_id[lab] = int(row["id"])
        self.combo_supplier["values"] = labels
        if reset_selection or cur not in labels:
            self.combo_supplier.set(labels[0])
        else:
            self.combo_supplier.set(cur)

    def _set_supplier_combo_from_id(self, sid: int | None) -> None:
        if not hasattr(self, "combo_supplier"):
            return
        if sid is None:
            self.combo_supplier.set("(geen leverancier)")
            return
        self._refresh_supplier_combo(reset_selection=False)
        with get_connection() as conn:
            row = conn.execute(
                "SELECT id, code, name FROM suppliers WHERE id = ?", (sid,)
            ).fetchone()
        if not row:
            self.combo_supplier.set("(geen leverancier)")
            return
        lab = f"{row['code'] + ' — ' if row['code'] else ''}{row['name']}"
        vals = list(self.combo_supplier["values"])
        if lab in vals:
            self.combo_supplier.set(lab)
        else:
            self.combo_supplier.set("(geen leverancier)")

    # --- Categorieën ---
    def _build_categories_tab(self) -> None:
        hint = ttk.Label(
            self._frm_categories,
            text="Hoofdcategorieën en subcategorieën. Gebruik «Hoofdcategorie aanmaken» om een hoofdgroep met opties (afmetingen, afkortverlies, …) te maken. "
            "Bij een hoofdcategorie kun je ook hieronder aanvinken welke extra velden gelden. "
            "Dubbelklik op een regel opent het tabblad Producten met filter op die categorie.",
            wraplength=920,
            justify=tk.LEFT,
        )
        hint.pack(anchor=tk.W, padx=8, pady=(8, 0))

        top = ttk.Frame(self._frm_categories)
        top.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree_categories = ttk.Treeview(
            top,
            columns=("code", "sort_order", "n_products"),
            show="tree headings",
            height=14,
            selectmode="browse",
        )
        self.tree_categories.heading("#0", text="Naam")
        self.tree_categories.column("#0", width=260)
        for c, t, w in [
            ("code", "Code (intern)", 130),
            ("sort_order", "Volgorde", 70),
            ("n_products", "Producten", 90),
        ]:
            self.tree_categories.heading(c, text=t)
            self.tree_categories.column(c, width=w)
        sb = ttk.Scrollbar(top, orient=tk.VERTICAL, command=self.tree_categories.yview)
        self.tree_categories.configure(yscrollcommand=sb.set)
        self.tree_categories.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_categories.bind("<<TreeviewSelect>>", self._on_category_select)
        self.tree_categories.bind("<Double-1>", self._on_category_double_click)

        form = ttk.LabelFrame(self._frm_categories, text="Categorie bewerken")
        form.pack(fill=tk.X, padx=8, pady=(0, 8))

        r = 0
        ttk.Label(form, text="Onderdeel van (hoofdcategorie)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.combo_cat_parent = ttk.Combobox(form, width=38, state="readonly")
        self.combo_cat_parent.grid(row=r, column=1, columnspan=2, sticky=tk.W, padx=4, pady=2)
        ttk.Label(
            form,
            text="Kies een hoofdcategorie voor een subcategorie; «Hoofdcategorie» = topniveau.",
            foreground="#666",
            font=("", 9),
        ).grid(row=r, column=3, sticky=tk.W, padx=8)
        r += 1
        ttk.Label(form, text="Code (intern, uniek)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_cat_code = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_cat_code, width=32).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        ttk.Label(
            form,
            text="Letters (hoofd/klein), cijfers, _ en -; spaties worden _",
            foreground="#666",
            font=("", 9),
        ).grid(row=r, column=2, sticky=tk.W, padx=8)
        r += 1
        ttk.Label(form, text="Weergavenaam").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_cat_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_cat_name, width=40).grid(
            row=r, column=1, columnspan=2, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="Sorteervolgorde").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_cat_sort = tk.StringVar(value="100")
        ttk.Entry(form, textvariable=self.var_cat_sort, width=12).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        ttk.Label(
            form,
            text="Lager = eerder in lijsten",
            foreground="#666",
            font=("", 9),
        ).grid(row=r, column=2, sticky=tk.W, padx=8)
        r += 1

        self.frm_cat_root_features = ttk.LabelFrame(
            form, text="Extra velden voor producten in deze hoofdcategorie (en subgroepen)"
        )
        self.frm_cat_root_features.grid(
            row=r, column=0, columnspan=4, sticky=tk.EW, padx=4, pady=4
        )
        self._root_feature_vars: dict[str, tk.BooleanVar] = {}
        frf = ttk.Frame(self.frm_cat_root_features)
        frf.pack(fill=tk.X, padx=8, pady=6)
        for code, label in ROOT_FEATURE_LABELS.items():
            v = tk.BooleanVar(value=False)
            self._root_feature_vars[code] = v
            ttk.Checkbutton(frf, text=label, variable=v).pack(anchor=tk.W, pady=1)
        self.frm_cat_root_features.grid_remove()
        r += 1

        btnf = ttk.Frame(form)
        btnf.grid(row=r, column=0, columnspan=3, pady=10, sticky=tk.W)
        ttk.Button(btnf, text="Nieuwe categorie", command=self._category_new).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(
            btnf,
            text="Hoofdcategorie aanmaken…",
            command=self._open_create_root_category_dialog,
        ).pack(side=tk.LEFT, padx=4)
        ttk.Button(btnf, text="Opslaan", command=self._category_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btnf, text="Verwijderen", command=self._category_delete).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btnf, text="Vernieuwen", command=self._refresh_categories_tree).pack(
            side=tk.LEFT, padx=4
        )

        self._editing_category_id: int | None = None

    def _refresh_cat_parent_combo(self) -> None:
        """Vul keuzelijst voor ouder: alleen hoofdcategorieën (of zichzelf uitsluiten bij bewerken)."""
        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT id, name FROM categories
                WHERE parent_id IS NULL
                ORDER BY sort_order, name
                """
            ).fetchall()
        labels = ["(Hoofdcategorie — geen ouder)"]
        self._cat_parent_label_to_id: dict[str, int | None] = {labels[0]: None}
        for r in rows:
            if self._editing_category_id is not None and int(r["id"]) == self._editing_category_id:
                continue
            labels.append(r["name"])
            self._cat_parent_label_to_id[r["name"]] = int(r["id"])
        self.combo_cat_parent["values"] = labels

    def _refresh_categories_tree(self) -> None:
        for i in self.tree_categories.get_children():
            self.tree_categories.delete(i)
        with get_connection() as conn:
            roots = conn.execute(
                """
                SELECT id, code, name, sort_order FROM categories
                WHERE parent_id IS NULL
                ORDER BY sort_order, name
                """
            ).fetchall()
            for root in roots:
                n_p = conn.execute(
                    "SELECT COUNT(*) FROM products WHERE category_id = ?",
                    (root["id"],),
                ).fetchone()[0]
                rid = str(root["id"])
                self.tree_categories.insert(
                    "",
                    tk.END,
                    iid=rid,
                    text=root["name"],
                    values=(root["code"], root["sort_order"], n_p),
                )
                children = conn.execute(
                    """
                    SELECT id, code, name, sort_order FROM categories
                    WHERE parent_id = ?
                    ORDER BY sort_order, name
                    """,
                    (root["id"],),
                ).fetchall()
                for ch in children:
                    n_c = conn.execute(
                        "SELECT COUNT(*) FROM products WHERE category_id = ?",
                        (ch["id"],),
                    ).fetchone()[0]
                    self.tree_categories.insert(
                        rid,
                        tk.END,
                        iid=str(ch["id"]),
                        text=ch["name"],
                        values=(ch["code"], ch["sort_order"], n_c),
                    )

    def _on_category_double_click(self, _evt=None) -> None:
        sel = self.tree_categories.selection()
        if not sel:
            return
        cid = int(sel[0])
        self._go_to_products_for_category(cid)

    def _go_to_products_for_category(self, category_id: int) -> None:
        """Spring naar Producten-tab en zet filter op deze categorie (subboom of exact)."""
        with get_connection() as conn:
            row = conn.execute(
                "SELECT parent_id FROM categories WHERE id = ?", (category_id,)
            ).fetchone()
        if not row:
            return
        if row["parent_id"] is None:
            self._filter_main_id = category_id
            self._filter_sub_id = None
        else:
            self._filter_main_id = int(row["parent_id"])
            self._filter_sub_id = category_id
        self._notebook.select(self._frm_products)
        self._refresh_filter_combos_from_state()
        self._refresh_products_tree()

    def _on_category_select(self, _evt=None) -> None:
        sel = self.tree_categories.selection()
        if not sel:
            return
        cid = int(sel[0])
        self._editing_category_id = cid
        self._refresh_cat_parent_combo()
        with get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM categories WHERE id = ?", (cid,)
            ).fetchone()
            if not row:
                return
            pid = row["parent_id"] if "parent_id" in row.keys() else None
            prow = None
            if pid is not None:
                prow = conn.execute(
                    "SELECT name FROM categories WHERE id = ?", (pid,)
                ).fetchone()
        self.var_cat_code.set(row["code"])
        self.var_cat_name.set(row["name"])
        self.var_cat_sort.set(str(row["sort_order"]))
        if pid is None:
            self.combo_cat_parent.set("(Hoofdcategorie — geen ouder)")
            self.frm_cat_root_features.grid()
            with get_connection() as conn:
                feats = conn.execute(
                    """
                    SELECT feature_code FROM category_root_features
                    WHERE root_category_id = ?
                    """,
                    (cid,),
                ).fetchall()
            have = {f["feature_code"] for f in feats}
            for fc, var in self._root_feature_vars.items():
                var.set(fc in have)
        elif prow and prow["name"] in self.combo_cat_parent["values"]:
            self.combo_cat_parent.set(prow["name"])
            self.frm_cat_root_features.grid_remove()
        else:
            self.combo_cat_parent.set("(Hoofdcategorie — geen ouder)")
            self.frm_cat_root_features.grid_remove()

    def _category_new(self) -> None:
        self._editing_category_id = None
        self._refresh_cat_parent_combo()
        self.var_cat_code.set("")
        self.var_cat_name.set("")
        self.var_cat_sort.set("100")
        self.combo_cat_parent.set("(Hoofdcategorie — geen ouder)")
        for var in self._root_feature_vars.values():
            var.set(False)
        self.frm_cat_root_features.grid_remove()

    def _open_create_root_category_dialog(self) -> None:
        dlg = tk.Toplevel(self)
        dlg.title("Nieuwe hoofdcategorie")
        dlg.transient(self)
        dlg.grab_set()
        f = ttk.Frame(dlg, padding=12)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="Code (intern)").grid(row=0, column=0, sticky=tk.W, pady=2)
        var_code = tk.StringVar()
        ttk.Entry(f, textvariable=var_code, width=32).grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Label(f, text="Weergavenaam").grid(row=1, column=0, sticky=tk.W, pady=2)
        var_name = tk.StringVar()
        ttk.Entry(f, textvariable=var_name, width=40).grid(row=1, column=1, sticky=tk.W, pady=2)
        ttk.Label(f, text="Sorteervolgorde").grid(row=2, column=0, sticky=tk.W, pady=2)
        var_sort = tk.StringVar(value="100")
        ttk.Entry(f, textvariable=var_sort, width=12).grid(row=2, column=1, sticky=tk.W, pady=2)
        ttk.Label(f, text="Welke extra velden op producten?").grid(
            row=3, column=0, sticky=tk.NW, pady=(12, 4)
        )
        cb_frame = ttk.Frame(f)
        cb_frame.grid(row=3, column=1, sticky=tk.W, pady=(12, 4))
        dlg_vars: dict[str, tk.BooleanVar] = {}
        for code, label in ROOT_FEATURE_LABELS.items():
            v = tk.BooleanVar(value=False)
            dlg_vars[code] = v
            ttk.Checkbutton(cb_frame, text=label, variable=v).pack(anchor=tk.W, pady=1)

        def do_ok() -> None:
            code = self._normalize_category_code(var_code.get())
            name = var_name.get().strip()
            if not code or not name:
                messagebox.showwarning(
                    "Ontbrekend", "Vul code en weergavenaam in.", parent=dlg
                )
                return
            if not all(ch.isalnum() or ch in "_-" for ch in code):
                messagebox.showerror(
                    "Code",
                    "Code mag alleen letters, cijfers, _ en - bevatten.",
                    parent=dlg,
                )
                return
            try:
                sort_order = int(var_sort.get().strip())
            except ValueError:
                messagebox.showerror(
                    "Fout",
                    "Sorteervolgorde moet een geheel getal zijn.",
                    parent=dlg,
                )
                return
            try:
                with transaction() as conn:
                    conn.execute(
                        """
                        INSERT INTO categories (code, name, sort_order, parent_id)
                        VALUES (?, ?, ?, NULL)
                        """,
                        (code, name, sort_order),
                    )
                    rid = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
                    for fc, var in dlg_vars.items():
                        if var.get():
                            conn.execute(
                                """
                                INSERT INTO category_root_features (root_category_id, feature_code)
                                VALUES (?, ?)
                                """,
                                (rid, fc),
                            )
            except Exception as e:
                messagebox.showerror("Database", str(e), parent=dlg)
                return
            dlg.destroy()
            self._refresh_categories_tree()
            self._refresh_category_combos()
            self._refresh_products_tree()
            messagebox.showinfo("Opgeslagen", "Hoofdcategorie aangemaakt.")

        bf = ttk.Frame(f)
        bf.grid(row=4, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="Aanmaken", command=do_ok).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="Annuleren", command=dlg.destroy).pack(side=tk.LEFT, padx=4)

    @staticmethod
    def _normalize_category_code(raw: str) -> str:
        """Spaties naar _, verder hoofdletters en kleine letters behouden zoals ingevoerd."""
        return raw.strip().replace(" ", "_")

    def _category_save(self) -> None:
        code = self._normalize_category_code(self.var_cat_code.get())
        name = self.var_cat_name.get().strip()
        if not code or not name:
            messagebox.showwarning("Ontbrekend", "Vul code en weergavenaam in.")
            return
        if not all(ch.isalnum() or ch in "_-" for ch in code):
            messagebox.showerror(
                "Code",
                "Code mag alleen letters, cijfers, _ en - bevatten (geen spaties).",
            )
            return
        try:
            sort_order = int(self.var_cat_sort.get().strip())
        except ValueError:
            messagebox.showerror("Fout", "Sorteervolgorde moet een geheel getal zijn.")
            return

        if not getattr(self, "_cat_parent_label_to_id", None):
            self._refresh_cat_parent_combo()
        plabel = self.combo_cat_parent.get()
        parent_id = self._cat_parent_label_to_id.get(plabel)
        if parent_id is not None and self._editing_category_id is not None:
            if parent_id == self._editing_category_id:
                messagebox.showerror("Fout", "Een categorie kan niet onder zichzelf hangen.")
                return

        try:
            with transaction() as conn:
                if parent_id is not None:
                    pr = conn.execute(
                        "SELECT parent_id FROM categories WHERE id = ?", (parent_id,)
                    ).fetchone()
                    if pr is None or pr["parent_id"] is not None:
                        messagebox.showerror(
                            "Fout",
                            "Subcategorieën kunnen alleen direct onder een hoofdcategorie hangen.",
                        )
                        return
                if self._editing_category_id is None:
                    conn.execute(
                        "INSERT INTO categories (code, name, sort_order, parent_id) VALUES (?, ?, ?, ?)",
                        (code, name, sort_order, parent_id),
                    )
                    cat_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
                else:
                    cat_id = self._editing_category_id
                    conn.execute(
                        """UPDATE categories SET code = ?, name = ?, sort_order = ?, parent_id = ?
                           WHERE id = ?""",
                        (code, name, sort_order, parent_id, cat_id),
                    )
                prow = conn.execute(
                    "SELECT parent_id FROM categories WHERE id = ?", (cat_id,)
                ).fetchone()
                conn.execute(
                    "DELETE FROM category_root_features WHERE root_category_id = ?",
                    (cat_id,),
                )
                if prow and prow["parent_id"] is None:
                    for fc, var in self._root_feature_vars.items():
                        if var.get():
                            conn.execute(
                                """
                                INSERT INTO category_root_features (root_category_id, feature_code)
                                VALUES (?, ?)
                                """,
                                (cat_id, fc),
                            )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return

        self._refresh_categories_tree()
        self._refresh_category_combos()
        self._refresh_products_tree()
        messagebox.showinfo("Opgeslagen", "Categorie opgeslagen.")

    def _category_delete(self) -> None:
        if self._editing_category_id is None:
            messagebox.showinfo("Geen selectie", "Selecteer eerst een categorie.")
            return
        cid = self._editing_category_id
        with get_connection() as conn:
            n_sub = conn.execute(
                "SELECT COUNT(*) FROM categories WHERE parent_id = ?", (cid,)
            ).fetchone()[0]
            if n_sub > 0:
                messagebox.showwarning(
                    "Subcategorieën",
                    "Verwijder of verplaats eerst alle subcategorieën onder deze categorie.",
                )
                return
            n = conn.execute(
                "SELECT COUNT(*) FROM products WHERE category_id = ?", (cid,)
            ).fetchone()[0]
            cat_name = conn.execute(
                "SELECT name FROM categories WHERE id = ?", (cid,)
            ).fetchone()[0]

        if n > 0:
            others = self._other_category_options(exclude_id=cid)
            if not others:
                messagebox.showerror(
                    "Niet mogelijk",
                    "Er zijn producten in deze categorie en er is geen andere categorie om naar te verplaatsen.",
                )
                return
            dlg = tk.Toplevel(self)
            dlg.title("Producten verplaatsen")
            dlg.transient(self)
            dlg.grab_set()
            ttk.Label(
                dlg,
                text=f"Er staan {n} product(en) in «{cat_name}».\n"
                "Kies een categorie om ze naartoe te verplaatsen:",
                justify=tk.LEFT,
            ).pack(padx=12, pady=12)
            var_move = tk.StringVar()
            cb = ttk.Combobox(dlg, textvariable=var_move, values=list(others.keys()), width=40, state="readonly")
            cb.pack(padx=12, pady=4)
            cb.set(list(others.keys())[0])

            def do_move() -> None:
                label = var_move.get()
                target_id = others.get(label)
                if target_id is None:
                    return
                try:
                    with transaction() as conn:
                        conn.execute(
                            "UPDATE products SET category_id = ? WHERE category_id = ?",
                            (target_id, cid),
                        )
                        conn.execute("DELETE FROM categories WHERE id = ?", (cid,))
                except Exception as e:
                    messagebox.showerror("Database", str(e), parent=dlg)
                    return
                dlg.destroy()
                self._category_new()
                self._refresh_categories_tree()
                self._refresh_category_combos()
                self._refresh_products_tree()
                messagebox.showinfo("Verwijderd", "Producten verplaatst en categorie verwijderd.")

            def cancel() -> None:
                dlg.destroy()

            bf = ttk.Frame(dlg)
            bf.pack(pady=12)
            ttk.Button(bf, text="Verplaatsen en verwijderen", command=do_move).pack(
                side=tk.LEFT, padx=6
            )
            ttk.Button(bf, text="Annuleren", command=cancel).pack(side=tk.LEFT, padx=6)
            return

        if not messagebox.askyesno(
            "Verwijderen", f"Categorie «{cat_name}» verwijderen? Er zijn geen gekoppelde producten."
        ):
            return
        try:
            with transaction() as conn:
                conn.execute("DELETE FROM categories WHERE id = ?", (cid,))
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return
        self._category_new()
        self._refresh_categories_tree()
        self._refresh_category_combos()
        self._refresh_products_tree()
        messagebox.showinfo("Verwijderd", "Categorie verwijderd.")

    def _other_category_options(self, exclude_id: int) -> dict[str, int]:
        """Label -> id voor alle categorieën behalve exclude_id."""
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, name FROM categories WHERE id != ? ORDER BY sort_order, name",
                (exclude_id,),
            ).fetchall()
        return {r["name"]: int(r["id"]) for r in rows}

    # --- Producten ---
    def _build_products_tab(self) -> None:
        filt = ttk.Frame(self._frm_products)
        filt.pack(fill=tk.X, padx=4, pady=(4, 0))
        ttk.Label(filt, text="Filter hoofdcategorie:").pack(side=tk.LEFT, padx=(0, 8))
        self.combo_filter_main = ttk.Combobox(filt, width=22, state="readonly")
        self.combo_filter_main.pack(side=tk.LEFT, padx=(0, 10))
        self.combo_filter_main.bind("<<ComboboxSelected>>", self._on_filter_main_change)
        ttk.Label(filt, text="Sub:").pack(side=tk.LEFT, padx=(0, 6))
        self.combo_filter_sub = ttk.Combobox(filt, width=30, state="readonly")
        self.combo_filter_sub.pack(side=tk.LEFT)
        self.combo_filter_sub.bind("<<ComboboxSelected>>", self._on_filter_sub_change)

        top = ttk.Frame(self._frm_products)
        top.pack(fill=tk.X, padx=4, pady=4)

        cols = ("id", "code", "name", "category", "unit", "unit_price", "price_source")
        self.tree_products = ttk.Treeview(
            top, columns=cols, show="headings", height=14, selectmode="browse"
        )
        for c, t, w in [
            ("id", "ID", 50),
            ("code", "Code", 100),
            ("name", "Naam", 180),
            ("category", "Categorie", 130),
            ("unit", "Ehd", 60),
            ("unit_price", "Prijs/st", 80),
            ("price_source", "Bron", 90),
        ]:
            self.tree_products.heading(c, text=t)
            self.tree_products.column(c, width=w)
        sb = ttk.Scrollbar(top, orient=tk.VERTICAL, command=self.tree_products.yview)
        self.tree_products.configure(yscrollcommand=sb.set)
        self.tree_products.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_products.bind("<<TreeviewSelect>>", self._on_product_select)

        form = ttk.LabelFrame(self._frm_products, text="Product bewerken")
        form.pack(fill=tk.BOTH, expand=True, padx=4, pady=8)

        r = 0
        ttk.Label(form, text="Code").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_code = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_code, width=24).grid(
            row=r, column=1, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        ttk.Label(form, text="Naam").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_name = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_name, width=48).grid(
            row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=2
        )
        r += 1
        self.lbl_product_name_hint = ttk.Label(
            form, text="", foreground="#666", font=("", 9), wraplength=720
        )
        self.lbl_product_name_hint.grid(
            row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=(0, 2)
        )
        r += 1
        ttk.Label(form, text="Hoofdcategorie").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.combo_product_main_cat = ttk.Combobox(form, width=36, state="readonly")
        self.combo_product_main_cat.grid(row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=2)
        self.combo_product_main_cat.bind("<<ComboboxSelected>>", self._on_product_main_change)
        r += 1
        ttk.Label(form, text="Subcategorie").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.combo_product_sub_cat = ttk.Combobox(form, width=36, state="readonly")
        self.combo_product_sub_cat.grid(row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=2)
        self.combo_product_sub_cat.bind("<<ComboboxSelected>>", self._on_product_sub_change)
        r += 1

        self.frm_hout = ttk.LabelFrame(
            form,
            text="Extra velden (afhankelijk van hoofdcategorie)",
        )
        self.frm_hout.grid(row=r, column=0, columnspan=4, sticky=tk.EW, padx=4, pady=4)
        self._frm_mm_dims = ttk.Frame(self.frm_hout)
        self._frm_mm_dims.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(self._frm_mm_dims, text="Dikte").pack(side=tk.LEFT)
        self.var_hout_dikte = tk.StringVar()
        ttk.Entry(self._frm_mm_dims, textvariable=self.var_hout_dikte, width=10).pack(
            side=tk.LEFT, padx=(4, 16)
        )
        ttk.Label(self._frm_mm_dims, text="Breedte").pack(side=tk.LEFT)
        self.var_hout_breedte = tk.StringVar()
        ttk.Entry(self._frm_mm_dims, textvariable=self.var_hout_breedte, width=10).pack(
            side=tk.LEFT, padx=4
        )
        self._frm_afkort = ttk.Frame(self.frm_hout)
        self._frm_afkort.pack(fill=tk.X, padx=6, pady=(0, 4))
        ttk.Label(self._frm_afkort, text="Afkortverlies (%)").pack(side=tk.LEFT)
        self.var_hout_afkort_verlies = tk.StringVar()
        ttk.Entry(self._frm_afkort, textvariable=self.var_hout_afkort_verlies, width=10).pack(
            side=tk.LEFT, padx=(4, 0)
        )
        r += 1

        self.var_supplier_article = tk.StringVar()
        self.var_unit_price = tk.StringVar(value="0")
        ttk.Label(form, text="Eenheid").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_unit = tk.StringVar(value="stuk")
        self.combo_unit = ttk.Combobox(
            form,
            textvariable=self.var_unit,
            width=10,
            state="readonly",
            values=("stuk", "m1", "m²"),
        )
        self.combo_unit.grid(row=r, column=1, sticky=tk.W, padx=4, pady=2)
        self._frm_unit_price_row = ttk.Frame(form)
        self._frm_unit_price_row_grid = {
            "row": r,
            "column": 2,
            "columnspan": 2,
            "sticky": tk.W,
            "padx": 12,
        }
        self._frm_unit_price_row.grid(**self._frm_unit_price_row_grid)
        self.frm_supplier_csv = ttk.Frame(self._frm_unit_price_row)
        ttk.Label(
            self.frm_supplier_csv,
            text="Artikelcode leverancier (CSV-prijslijst)",
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(
            self.frm_supplier_csv,
            textvariable=self.var_supplier_article,
            width=22,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            self.frm_supplier_csv,
            text="Netto ophalen",
            command=self._fetch_netto_from_supplier_pricelist,
        ).pack(side=tk.LEFT)
        self.frm_supplier_csv.pack(side=tk.LEFT)
        r += 1
        ttk.Label(form, text="Inkoopprijs netto (€)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_purchase_price = tk.StringVar()
        e_pur = ttk.Entry(form, textvariable=self.var_purchase_price, width=12)
        e_pur.grid(row=r, column=1, sticky=tk.W, padx=4, pady=2)
        ttk.Label(form, text="Winstmarge (%)").grid(
            row=r, column=2, sticky=tk.W, padx=12, pady=2
        )
        self.var_margin_pct = tk.StringVar()
        e_mar = ttk.Entry(form, textvariable=self.var_margin_pct, width=8)
        e_mar.grid(row=r, column=3, sticky=tk.W, padx=4, pady=2)
        r += 1
        self.lbl_price_suggest = ttk.Label(
            form, text="", foreground="#333", font=("", 10)
        )
        self.lbl_price_suggest.grid(
            row=r, column=0, columnspan=4, sticky=tk.W, padx=4, pady=(0, 6)
        )
        self.var_purchase_price.trace_add("write", self._update_purchase_suggest_label)
        self.var_margin_pct.trace_add("write", self._update_purchase_suggest_label)
        self.var_hout_afkort_verlies.trace_add("write", self._update_purchase_suggest_label)
        self.var_hout_dikte.trace_add("write", self._sync_hout_product_name)
        self.var_hout_breedte.trace_add("write", self._sync_hout_product_name)
        r += 1
        self.frm_product_extra = ttk.LabelFrame(
            form, text="Extra velden (hoofdcategorie — uitbreiding)"
        )
        self.frm_product_extra.grid(row=r, column=0, columnspan=4, sticky=tk.EW, padx=4, pady=4)
        self._frm_opp = ttk.Frame(self.frm_product_extra)
        self._frm_opp.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(self._frm_opp, text="Min. oppervlakte (m²)").pack(side=tk.LEFT)
        self.var_min_opp = tk.StringVar()
        ttk.Entry(self._frm_opp, textvariable=self.var_min_opp, width=10).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Label(self._frm_opp, text="Max. oppervlakte (m²)").pack(side=tk.LEFT, padx=(12, 0))
        self.var_max_opp = tk.StringVar()
        ttk.Entry(self._frm_opp, textvariable=self.var_max_opp, width=10).pack(
            side=tk.LEFT, padx=4
        )
        self._frm_dikte_los = ttk.Frame(self.frm_product_extra)
        self._frm_dikte_los.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(self._frm_dikte_los, text="Dikte (mm) — los").pack(side=tk.LEFT)
        self.var_product_dikte_mm = tk.StringVar()
        ttk.Entry(self._frm_dikte_los, textvariable=self.var_product_dikte_mm, width=12).pack(
            side=tk.LEFT, padx=4
        )
        self._frm_lev = ttk.Frame(self.frm_product_extra)
        self._frm_lev.pack(fill=tk.X, padx=6, pady=4)
        self._frm_lev_top = ttk.Frame(self._frm_lev)
        self._frm_lev_top.pack(fill=tk.X)
        ttk.Label(self._frm_lev_top, text="Leverancier").pack(side=tk.LEFT)
        self.combo_supplier = ttk.Combobox(self._frm_lev_top, width=42, state="readonly")
        self.combo_supplier.pack(side=tk.LEFT, padx=(4, 12))
        self._frm_lev_article = ttk.Frame(self._frm_lev)
        ttk.Label(self._frm_lev_article, text="Inkoopartikel nr.").pack(side=tk.LEFT)
        ttk.Entry(self._frm_lev_article, textvariable=self.var_supplier_article, width=32).pack(
            side=tk.LEFT, padx=4
        )
        self._frm_uren = ttk.Frame(self.frm_product_extra)
        self._frm_uren.pack(fill=tk.X, padx=6, pady=(0, 4))
        ttk.Label(self._frm_uren, text="Minuten (totaal)").pack(side=tk.LEFT)
        self.var_work_minutes = tk.StringVar(value="")
        ttk.Entry(self._frm_uren, textvariable=self.var_work_minutes, width=10).pack(
            side=tk.LEFT, padx=4
        )
        self.lbl_work_total = ttk.Label(self._frm_uren, text="", foreground="#444")
        self.lbl_work_total.pack(side=tk.LEFT, padx=(12, 0))
        self.var_work_minutes.trace_add("write", self._update_work_minutes_label)
        self.frm_product_extra.grid_remove()
        r += 1
        ttk.Label(form, text="Omschrijving op offerte").grid(
            row=r, column=0, sticky=tk.NW, padx=4, pady=4
        )
        self.txt_quote_desc = tk.Text(form, width=72, height=5, wrap=tk.WORD)
        self.txt_quote_desc.grid(row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=4)
        r += 1

        self.lbl_effective = ttk.Label(form, text="")
        self.lbl_effective.grid(row=r, column=1, columnspan=3, sticky=tk.W, padx=4)

        btnf = ttk.Frame(form)
        btnf.grid(row=r + 1, column=0, columnspan=4, pady=8)
        ttk.Button(btnf, text="Nieuw product", command=self._product_new).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btnf, text="Opslaan", command=self._product_save).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btnf, text="Verwijderen", command=self._product_delete).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(btnf, text="Vernieuwen", command=self._refresh_products_tree).pack(
            side=tk.LEFT, padx=4
        )

        self._editing_product_id: int | None = None
        self._filter_main_id: int | None = None
        self._filter_sub_id: int | None = None
        self._supplier_label_to_id = {"(geen leverancier)": None}
        self._refresh_supplier_combo()

    def _update_work_minutes_label(self, *_args) -> None:
        """Toont het equivalent in uren; invoer is altijd in totale minuten."""
        try:
            s = self.var_work_minutes.get().strip().replace(",", ".")
            if not s:
                self.lbl_work_total.config(text="")
                return
            total_min = float(s)
            hours = total_min / 60.0
            self.lbl_work_total.config(text=f"= {hours:.2f} uur")
        except (ValueError, tk.TclError):
            self.lbl_work_total.config(text="")

    def _clear_product_extra_vars(self) -> None:
        self.var_min_opp.set("")
        self.var_max_opp.set("")
        self.var_product_dikte_mm.set("")
        self.var_supplier_article.set("")
        self.var_work_minutes.set("")
        self.lbl_work_total.config(text="")
        if hasattr(self, "combo_supplier"):
            self.combo_supplier.set("(geen leverancier)")

    def _load_product_extra_from_row(self, row) -> None:
        """Vult oppervlakte-, dikte-, leverancier- en urenvelden vanuit een products-query."""
        rk = row.keys()
        if "min_oppervlakte_m2" in rk and row["min_oppervlakte_m2"] is not None:
            self.var_min_opp.set(str(row["min_oppervlakte_m2"]))
        else:
            self.var_min_opp.set("")
        if "max_oppervlakte_m2" in rk and row["max_oppervlakte_m2"] is not None:
            self.var_max_opp.set(str(row["max_oppervlakte_m2"]))
        else:
            self.var_max_opp.set("")
        if "product_dikte_mm" in rk and row["product_dikte_mm"] is not None:
            self.var_product_dikte_mm.set(str(row["product_dikte_mm"]))
        else:
            self.var_product_dikte_mm.set("")
        if "supplier_article" in rk and row["supplier_article"]:
            self.var_supplier_article.set(str(row["supplier_article"]))
        else:
            self.var_supplier_article.set("")
        sid = None
        if "supplier_id" in rk and row["supplier_id"] is not None:
            sid = int(row["supplier_id"])
        self._set_supplier_combo_from_id(sid)
        wh = row["work_hours"] if "work_hours" in rk else None
        if wh is not None:
            total_min = float(wh) * 60.0
            if abs(total_min - round(total_min)) < 1e-9:
                self.var_work_minutes.set(str(int(round(total_min))))
            else:
                s = f"{total_min:.6f}".rstrip("0").rstrip(".")
                self.var_work_minutes.set(s.replace(".", ","))
        else:
            self.var_work_minutes.set("")
        self._update_work_minutes_label()

    def _update_product_extra_visibility(self) -> None:
        cid = self._product_resolve_category_id()
        with get_connection() as conn:
            if cid is None:
                self.frm_product_extra.grid_remove()
                self._clear_product_extra_vars()
                self._update_price_row_mode()
                return
            show_opp = category_root_has_feature(conn, cid, FEATURE_OPPERVLAKTE_M2)
            show_dl = category_root_has_feature(conn, cid, FEATURE_DIKTE_LOS_MM)
            show_lev = category_root_has_feature(conn, cid, FEATURE_LEVERANCIER)
            show_uren = category_root_has_feature(conn, cid, FEATURE_UREN_MINUTEN)
            if not (show_opp or show_dl or show_lev or show_uren):
                self.frm_product_extra.grid_remove()
                self._clear_product_extra_vars()
                self._update_price_row_mode()
                return
            root_id = get_root_category_id(conn, cid)
            rname = (
                conn.execute(
                    "SELECT name FROM categories WHERE id = ?", (root_id,)
                ).fetchone()
                if root_id is not None
                else None
            )
            rn = rname["name"] if rname else "categorie"
            self.frm_product_extra.config(text=f"Extra voor «{rn}» (uitbreiding)")
            self._frm_opp.pack_forget()
            self._frm_dikte_los.pack_forget()
            self._frm_lev.pack_forget()
            self._frm_uren.pack_forget()
            if show_opp:
                self._frm_opp.pack(fill=tk.X, padx=6, pady=4)
            else:
                self.var_min_opp.set("")
                self.var_max_opp.set("")
            if show_dl:
                self._frm_dikte_los.pack(fill=tk.X, padx=6, pady=4)
            else:
                self.var_product_dikte_mm.set("")
            if show_lev:
                self._frm_lev.pack(fill=tk.X, padx=6, pady=4)
                self._refresh_supplier_combo(reset_selection=False)
                self._frm_lev_article.pack_forget()
            else:
                self.var_supplier_article.set("")
                self.combo_supplier.set("(geen leverancier)")
                self._frm_lev_article.pack_forget()
            if show_uren:
                self._frm_uren.pack(fill=tk.X, padx=6, pady=(0, 4))
                self._update_work_minutes_label()
            else:
                self.var_work_minutes.set("")
                self.lbl_work_total.config(text="")
        self.frm_product_extra.grid()
        self._update_price_row_mode()

    def _refresh_category_combos(self) -> None:
        self._refresh_filter_combos()
        self._refresh_product_category_combos()
        self._update_hout_panel_visibility()
        self._sync_hout_product_name()
        self._update_product_extra_visibility()

    def _refresh_filter_combos(self) -> None:
        with get_connection() as conn:
            roots = conn.execute(
                """
                SELECT id, name FROM categories
                WHERE parent_id IS NULL
                ORDER BY sort_order, name
                """
            ).fetchall()
        main_labels = ["(Alle categorieën)"]
        self._filter_main_label_to_id: dict[str, int | None] = {main_labels[0]: None}
        for r in roots:
            main_labels.append(r["name"])
            self._filter_main_label_to_id[r["name"]] = int(r["id"])
        self.combo_filter_main["values"] = main_labels
        self.combo_filter_main.set(main_labels[0])
        self._filter_main_id = None
        self._filter_sub_id = None
        self.combo_filter_sub["values"] = []
        self.combo_filter_sub.set("")
        self.combo_filter_sub.configure(state="disabled")

    def _refresh_filter_combos_from_state(self) -> None:
        """Zet filter-comboboxen volgens _filter_main_id / _filter_sub_id (na dubbelklik Categorieën)."""
        with get_connection() as conn:
            roots = conn.execute(
                """
                SELECT id, name FROM categories
                WHERE parent_id IS NULL
                ORDER BY sort_order, name
                """
            ).fetchall()
            main_labels = ["(Alle categorieën)"]
            self._filter_main_label_to_id = {main_labels[0]: None}
            for r in roots:
                main_labels.append(r["name"])
                self._filter_main_label_to_id[r["name"]] = int(r["id"])
            self.combo_filter_main["values"] = main_labels
            if self._filter_main_id is None:
                self.combo_filter_main.set("(Alle categorieën)")
                self.combo_filter_sub["values"] = []
                self.combo_filter_sub.set("")
                self.combo_filter_sub.configure(state="disabled")
                return
            mrow = conn.execute(
                "SELECT name FROM categories WHERE id = ?", (self._filter_main_id,)
            ).fetchone()
            if mrow and mrow["name"] in main_labels:
                self.combo_filter_main.set(mrow["name"])
            self._populate_filter_sub_combo(conn)
            if self._filter_sub_id is None:
                self.combo_filter_sub.set("(alle binnen deze hoofd)")
            else:
                srow = conn.execute(
                    "SELECT name FROM categories WHERE id = ?", (self._filter_sub_id,)
                ).fetchone()
                if srow and srow["name"] in self.combo_filter_sub["values"]:
                    self.combo_filter_sub.set(srow["name"])

    def _populate_filter_sub_combo(self, conn) -> None:
        if self._filter_main_id is None:
            self.combo_filter_sub["values"] = []
            self.combo_filter_sub.set("")
            self.combo_filter_sub.configure(state="disabled")
            self._filter_sub_id = None
            return
        subs = conn.execute(
            """
            SELECT id, name FROM categories
            WHERE parent_id = ?
            ORDER BY sort_order, name
            """,
            (self._filter_main_id,),
        ).fetchall()
        sub_labels = ["(alle binnen deze hoofd)"]
        self._filter_sub_label_to_id: dict[str, int | None] = {sub_labels[0]: None}
        for s in subs:
            sub_labels.append(s["name"])
            self._filter_sub_label_to_id[s["name"]] = int(s["id"])
        self.combo_filter_sub["values"] = sub_labels
        self.combo_filter_sub.configure(state="readonly")

    def _on_filter_main_change(self, _evt=None) -> None:
        label = self.combo_filter_main.get()
        self._filter_main_id = self._filter_main_label_to_id.get(label)
        self._filter_sub_id = None
        with get_connection() as conn:
            self._populate_filter_sub_combo(conn)
            if self._filter_main_id is not None:
                self.combo_filter_sub.set("(alle binnen deze hoofd)")
        self._refresh_products_tree()

    def _on_filter_sub_change(self, _evt=None) -> None:
        if self._filter_main_id is None:
            return
        sl = self.combo_filter_sub.get()
        self._filter_sub_id = self._filter_sub_label_to_id.get(sl)
        self._refresh_products_tree()

    def _refresh_product_category_combos(self) -> None:
        with get_connection() as conn:
            roots = conn.execute(
                """
                SELECT id, name FROM categories
                WHERE parent_id IS NULL
                ORDER BY sort_order, name
                """
            ).fetchall()
            mlabels = [r["name"] for r in roots]
            self._product_main_name_to_id = {r["name"]: int(r["id"]) for r in roots}
            self.combo_product_main_cat["values"] = mlabels
            if not mlabels:
                self.combo_product_sub_cat["values"] = []
                self._update_hout_panel_visibility()
                return
            if self.combo_product_main_cat.get() not in mlabels:
                self.combo_product_main_cat.set(mlabels[0])
            self._populate_product_sub_combo(conn)

    def _populate_product_sub_combo(self, conn) -> None:
        mname = self.combo_product_main_cat.get()
        mid = self._product_main_name_to_id.get(mname)
        if mid is None:
            self.combo_product_sub_cat["values"] = []
            self.combo_product_sub_cat.configure(state="disabled")
            self._update_hout_panel_visibility()
            return
        subs = conn.execute(
            """
            SELECT id, name FROM categories
            WHERE parent_id = ?
            ORDER BY sort_order, name
            """,
            (mid,),
        ).fetchall()
        sub_labels = ["(geen — product op hoofdcategorie)"]
        self._product_sub_label_to_id: dict[str, int | None] = {sub_labels[0]: None}
        for s in subs:
            sub_labels.append(s["name"])
            self._product_sub_label_to_id[s["name"]] = int(s["id"])
        self.combo_product_sub_cat["values"] = sub_labels
        if len(sub_labels) == 1:
            self.combo_product_sub_cat.set(sub_labels[0])
            self.combo_product_sub_cat.configure(state="disabled")
        else:
            self.combo_product_sub_cat.configure(state="readonly")
            self.combo_product_sub_cat.set(sub_labels[0])

    def _on_product_main_change(self, _evt=None) -> None:
        with get_connection() as conn:
            self._populate_product_sub_combo(conn)
        self._update_hout_panel_visibility()
        self._sync_hout_product_name()
        self._update_product_extra_visibility()

    def _on_product_sub_change(self, _evt=None) -> None:
        self._update_hout_panel_visibility()
        self._sync_hout_product_name()
        self._update_product_extra_visibility()

    @staticmethod
    def _format_mm_for_name(v: float) -> str:
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")

    def _sync_hout_product_name(self, *_args) -> None:
        """Subgroep met optie «naam_sub_maat»: naam = subgroep; met maten: naam + dikte×breedte, code = subcode_maat."""
        cid = self._product_resolve_category_id()
        with get_connection() as conn:
            if cid is None:
                self.lbl_product_name_hint.config(text="")
                return
            if not is_direct_sub_of_root_with_feature(
                conn, cid, FEATURE_NAAM_SUB_MAAT
            ) or not category_root_has_feature(conn, cid, FEATURE_DIKTE_BREEDTE_MM):
                self.lbl_product_name_hint.config(text="")
                return
            subrow = conn.execute(
                "SELECT name, code FROM categories WHERE id = ?", (cid,)
            ).fetchone()
        subname = (subrow["name"] if subrow else "").strip()
        cat_code = (subrow["code"] or "").strip() if subrow else ""
        if not subname:
            self.lbl_product_name_hint.config(text="")
            return
        try:
            ds = self.var_hout_dikte.get().strip()
            bs = self.var_hout_breedte.get().strip()
            if not ds or not bs:
                if self._editing_product_id is None:
                    self.var_name.set(subname)
                    if cat_code:
                        self.var_code.set(cat_code)
                self.lbl_product_name_hint.config(
                    text=(
                        "Subgroep: naam en code van de subcategorie; vul dikte en breedte (mm) — "
                        "dan naam + maat en unieke code (subcode_maat)."
                    )
                )
                return
            hd = float(ds.replace(",", "."))
            hb = float(bs.replace(",", "."))
        except (ValueError, tk.TclError):
            self.lbl_product_name_hint.config(
                text="Ongeldige dikte of breedte — corrigeer de mm of pas de naam handmatig aan."
            )
            return
        fd = self._format_mm_for_name(hd)
        fb = self._format_mm_for_name(hb)
        self.var_name.set(f"{subname} {fd}×{fb}")
        if cat_code and self._editing_product_id is None:
            fdc = fd.replace(",", "_")
            fbc = fb.replace(",", "_")
            self.var_code.set(f"{cat_code}_{fdc}x{fbc}")
        self.lbl_product_name_hint.config(
            text="Naam: subgroep + dikte × breedte. Code: subgroepcode + maat (aanpasbaar)."
        )

    @staticmethod
    def _normalize_unit(raw: str) -> str:
        u = (raw or "").strip().lower()
        if u in ("", "stuk", "stuks", "st"):
            return "stuk"
        if u in ("m", "m1", "lm", "lm1"):
            return "m1"
        if u in ("m2", "m²", "m^2"):
            return "m²"
        return "stuk"

    def _update_hout_panel_visibility(self) -> None:
        """Toont mm- en/of afkortvelden volgens ingestelde opties op de hoofdcategorie."""
        cid = self._product_resolve_category_id()
        with get_connection() as conn:
            if cid is None:
                self.frm_hout.grid_remove()
                self.var_hout_dikte.set("")
                self.var_hout_breedte.set("")
                self.var_hout_afkort_verlies.set("")
                return
            show_mm = category_root_has_feature(conn, cid, FEATURE_DIKTE_BREEDTE_MM)
            show_af = category_root_has_feature(conn, cid, FEATURE_AFKORT_VERLIES)
            if not show_mm and not show_af:
                self.frm_hout.grid_remove()
                self.var_hout_dikte.set("")
                self.var_hout_breedte.set("")
                self.var_hout_afkort_verlies.set("")
                return
            root_id = get_root_category_id(conn, cid)
            rname = (
                conn.execute(
                    "SELECT name FROM categories WHERE id = ?", (root_id,)
                ).fetchone()
                if root_id is not None
                else None
            )
            rn = rname["name"] if rname else "categorie"
            parts: list[str] = []
            if show_mm:
                parts.append("dikte × breedte (mm)")
            if show_af:
                parts.append("afkortverlies (%)")
            self.frm_hout.config(
                text=f"Extra voor «{rn}»: {', '.join(parts)}",
            )
            self.frm_hout.grid()
            self._frm_mm_dims.pack_forget()
            self._frm_afkort.pack_forget()
            if show_mm:
                self._frm_mm_dims.pack(fill=tk.X, padx=6, pady=4)
            else:
                self._frm_mm_dims.pack_forget()
                self.var_hout_dikte.set("")
                self.var_hout_breedte.set("")
            if show_af:
                self._frm_afkort.pack(fill=tk.X, padx=6, pady=(0, 4))
            else:
                self._frm_afkort.pack_forget()
                self.var_hout_afkort_verlies.set("")

    def _product_resolve_category_id(self) -> int | None:
        mname = self.combo_product_main_cat.get()
        mid = self._product_main_name_to_id.get(mname)
        if mid is None:
            return None
        sname = self.combo_product_sub_cat.get()
        sid = self._product_sub_label_to_id.get(sname)
        if sid is not None:
            return sid
        return mid

    def _refresh_products_tree(self) -> None:
        for i in self.tree_products.get_children():
            self.tree_products.delete(i)
        sql_base = """
            SELECT p.id, p.code, p.name, p.unit, p.price_source, p.unit_price,
                   CASE
                     WHEN c.id IS NULL THEN ''
                     WHEN c.parent_id IS NULL THEN c.name
                     ELSE (SELECT x.name FROM categories x WHERE x.id = c.parent_id) || ' › ' || c.name
                   END AS category_display
            FROM products p
            LEFT JOIN categories c ON c.id = p.category_id
        """
        with get_connection() as conn:
            if self._filter_main_id is None:
                rows = conn.execute(
                    sql_base + " ORDER BY category_display, p.code"
                ).fetchall()
            elif self._filter_sub_id is None:
                ids = category_subtree_ids(conn, self._filter_main_id)
                if not ids:
                    rows = []
                else:
                    ph = ",".join("?" * len(ids))
                    rows = conn.execute(
                        sql_base + f" WHERE p.category_id IN ({ph}) ORDER BY p.code",
                        ids,
                    ).fetchall()
            else:
                rows = conn.execute(
                    sql_base + " WHERE p.category_id = ? ORDER BY p.code",
                    (self._filter_sub_id,),
                ).fetchall()
        for row in rows:
            ps = row["price_source"]
            ps_label = "handmatig" if ps == "manual" else "uit BOM"
            self.tree_products.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["code"],
                    row["name"],
                    row["category_display"] or "—",
                    row["unit"],
                    f"{row['unit_price']:.2f}",
                    ps_label,
                ),
            )
        self._refresh_bom_parent_combo()
        self._refresh_quote_product_combo()

    def _on_product_select(self, _evt=None) -> None:
        sel = self.tree_products.selection()
        if not sel:
            return
        vals = self.tree_products.item(sel[0], "values")
        pid = int(vals[0])
        self._editing_product_id = pid
        with get_connection() as conn:
            row = conn.execute(
                """
                SELECT p.*, c.name AS category_name, c.parent_id AS cat_parent_id
                FROM products p
                LEFT JOIN categories c ON c.id = p.category_id
                WHERE p.id = ?
                """,
                (pid,),
            ).fetchone()
        if not row:
            return
        self.var_code.set(row["code"])
        self.var_name.set(row["name"])
        self.var_unit.set(self._normalize_unit(row["unit"] if "unit" in row.keys() else "stuk"))
        self.var_unit_price.set(str(row["unit_price"]))
        self._refresh_product_category_combos()
        cid = row["category_id"]
        if cid is None:
            if self.combo_product_main_cat["values"]:
                self.combo_product_main_cat.current(0)
                with get_connection() as conn:
                    self._populate_product_sub_combo(conn)
        else:
            with get_connection() as conn:
                cat = conn.execute(
                    "SELECT id, name, parent_id FROM categories WHERE id = ?", (cid,)
                ).fetchone()
                if not cat:
                    pass
                elif cat["parent_id"] is None:
                    if cat["name"] in self.combo_product_main_cat["values"]:
                        self.combo_product_main_cat.set(cat["name"])
                    self._populate_product_sub_combo(conn)
                    self.combo_product_sub_cat.set("(geen — product op hoofdcategorie)")
                else:
                    pr = conn.execute(
                        "SELECT name FROM categories WHERE id = ?", (cat["parent_id"],)
                    ).fetchone()
                    if pr and pr["name"] in self.combo_product_main_cat["values"]:
                        self.combo_product_main_cat.set(pr["name"])
                    self._populate_product_sub_combo(conn)
                    if cat["name"] in self.combo_product_sub_cat["values"]:
                        self.combo_product_sub_cat.set(cat["name"])
        self._update_hout_panel_visibility()
        if "hout_dikte_mm" in row.keys() and row["hout_dikte_mm"] is not None:
            self.var_hout_dikte.set(str(row["hout_dikte_mm"]))
        else:
            self.var_hout_dikte.set("")
        if "hout_breedte_mm" in row.keys() and row["hout_breedte_mm"] is not None:
            self.var_hout_breedte.set(str(row["hout_breedte_mm"]))
        else:
            self.var_hout_breedte.set("")
        if "hout_afkort_verlies_pct" in row.keys() and row["hout_afkort_verlies_pct"] is not None:
            self.var_hout_afkort_verlies.set(str(row["hout_afkort_verlies_pct"]))
        else:
            self.var_hout_afkort_verlies.set("")
        if "purchase_price" in row.keys() and row["purchase_price"] is not None:
            self.var_purchase_price.set(str(row["purchase_price"]))
        else:
            self.var_purchase_price.set("")
        if "margin_pct" in row.keys() and row["margin_pct"] is not None:
            self.var_margin_pct.set(str(row["margin_pct"]))
        else:
            self.var_margin_pct.set("")
        self._update_purchase_suggest_label()
        self._sync_hout_product_name()
        self.txt_quote_desc.delete("1.0", tk.END)
        self.txt_quote_desc.insert("1.0", row["quote_description"] or "")
        self._load_product_extra_from_row(row)
        self._update_product_extra_visibility()
        self.after(50, self._update_effective_label)

    def _update_effective_label(self) -> None:
        if self._editing_product_id is None:
            self.lbl_effective.config(text="")
            return
        try:
            p = effective_unit_price(self._editing_product_id)
            self.lbl_effective.config(
                text=f"Effectieve verkoopprijs (voor berekening): € {p:.2f}"
            )
        except ValueError as e:
            self.lbl_effective.config(text=str(e))

    def _update_purchase_suggest_label(self, *_args) -> None:
        try:
            ps = self.var_purchase_price.get().strip()
            ms = self.var_margin_pct.get().strip()
            if not ps or not ms:
                self.lbl_price_suggest.config(text="")
                return
            pp = float(ps.replace(",", "."))
            mp = float(ms.replace(",", "."))
            avs = self.var_hout_afkort_verlies.get().strip()
            av = float(avs.replace(",", ".")) if avs else 0.0
            cost_na_afkort = pp * (1.0 + av / 100.0)
            sug = cost_na_afkort * (1.0 + mp / 100.0)
            if av != 0.0:
                uitleg = f"inkoop × (1+{av:g}% afkort) × (1+{mp:g}% marge)"
            else:
                uitleg = f"inkoop × (1+{mp:g}% marge)"
            self.lbl_price_suggest.config(text=f"Totaalprijs: € {sug:.2f}  ({uitleg})")
        except (ValueError, tk.TclError):
            self.lbl_price_suggest.config(text="")

    def _product_use_supplier_pricelist(self) -> bool:
        cid = self._product_resolve_category_id()
        if cid is None:
            return False
        with get_connection() as conn:
            return category_root_has_feature(conn, cid, FEATURE_LEVERANCIER)

    def _update_price_row_mode(self) -> None:
        if self._product_use_supplier_pricelist():
            self._frm_unit_price_row.grid(**self._frm_unit_price_row_grid)
        else:
            self._frm_unit_price_row.grid_remove()

    def _computed_sale_price_from_margin(self) -> float | None:
        try:
            ps = self.var_purchase_price.get().strip()
            ms = self.var_margin_pct.get().strip()
            if not ps or not ms:
                return None
            pp = float(ps.replace(",", "."))
            mp = float(ms.replace(",", "."))
            avs = self.var_hout_afkort_verlies.get().strip()
            av = float(avs.replace(",", ".")) if avs else 0.0
            cost = pp * (1.0 + av / 100.0)
            return cost * (1.0 + mp / 100.0)
        except (ValueError, tk.TclError):
            return None

    def _fetch_netto_from_supplier_pricelist(self) -> None:
        sid = self._supplier_label_to_id.get(self.combo_supplier.get())
        art = self.var_supplier_article.get().strip()
        if not sid:
            messagebox.showwarning("Leverancier", "Kies een leverancier.")
            return
        if not art:
            messagebox.showwarning(
                "Artikelcode leverancier",
                "Vul de artikelcode in zoals in de prijslijst van de leverancier.",
            )
            return
        with get_connection() as conn:
            row = conn.execute(
                """
                SELECT price_list_path, csv_col_article, csv_col_netto, csv_col_unit,
                       csv_col_description, csv_skip_header_rows
                FROM suppliers WHERE id = ?
                """,
                (sid,),
            ).fetchone()
        if not row or not row["price_list_path"]:
            messagebox.showerror(
                "Prijslijst",
                "Stel bij deze leverancier (tab Leveranciers) het pad naar CSV in.",
            )
            return
        path = row["price_list_path"]
        if not Path(path).is_file():
            messagebox.showerror(
                "Prijslijst",
                f"Bestand niet gevonden of niet bereikbaar:\n{path}",
            )
            return
        ca = row["csv_col_article"] if "csv_col_article" in row.keys() else None
        cn = row["csv_col_netto"] if "csv_col_netto" in row.keys() else None
        cu = row["csv_col_unit"] if "csv_col_unit" in row.keys() else None
        cdesc = row["csv_col_description"] if "csv_col_description" in row.keys() else None
        sk = row["csv_skip_header_rows"] if "csv_skip_header_rows" in row.keys() else 1
        try:
            netto, unit_csv, desc_csv = lookup_netto_price_from_file(
                path,
                art,
                col_article=ca,
                col_netto=cn,
                col_unit=(cu or None),
                col_description=(cdesc or None),
                skip_header_rows=int(sk) if sk is not None else 1,
            )
        except ValueError as e:
            messagebox.showerror("Prijslijst", str(e))
            return
        except OSError as e:
            messagebox.showerror("Prijslijst", str(e))
            return
        if netto is None:
            messagebox.showwarning(
                "Prijslijst",
                f"Geen netto-prijs gevonden voor artikel «{art}» in het CSV-bestand.",
            )
            return
        self.var_purchase_price.set(f"{netto:.2f}".replace(".", ","))
        if unit_csv:
            self.var_unit.set(self._normalize_unit(unit_csv))
        if desc_csv:
            self.txt_quote_desc.delete("1.0", tk.END)
            self.txt_quote_desc.insert("1.0", desc_csv)
        self._update_purchase_suggest_label()
        sug = self._computed_sale_price_from_margin()
        if sug is not None:
            self.var_unit_price.set(f"{sug:.2f}")

    def _product_new(self) -> None:
        self._editing_product_id = None
        self.var_code.set("")
        self.var_name.set("")
        self.var_unit.set("stuk")
        self.var_unit_price.set("0")
        self.var_hout_dikte.set("")
        self.var_hout_breedte.set("")
        self.var_hout_afkort_verlies.set("")
        self.var_purchase_price.set("")
        self.var_margin_pct.set("")
        self.lbl_price_suggest.config(text="")
        self.txt_quote_desc.delete("1.0", tk.END)
        self.lbl_effective.config(text="")
        self._refresh_product_category_combos()

        use_hout_filter_sub = False
        with get_connection() as conn:
            if self._filter_sub_id is not None and is_direct_sub_of_root_with_feature(
                conn, self._filter_sub_id, FEATURE_NAAM_SUB_MAAT
            ):
                row = conn.execute(
                    """
                    SELECT c.code, c.name, p.name AS parent_name
                    FROM categories c
                    JOIN categories p ON p.id = c.parent_id
                    WHERE c.id = ?
                    """,
                    (self._filter_sub_id,),
                ).fetchone()
                vals = self.combo_product_main_cat["values"]
                if (
                    row
                    and vals
                    and row["parent_name"] in vals
                ):
                    use_hout_filter_sub = True
                    self.combo_product_main_cat.set(row["parent_name"])
                    self._populate_product_sub_combo(conn)
                    if row["name"] in self.combo_product_sub_cat["values"]:
                        self.combo_product_sub_cat.set(row["name"])
                    self.var_code.set(row["code"])
                    self.var_name.set(row["name"])

        if not use_hout_filter_sub:
            if self.combo_product_main_cat["values"]:
                if "Hout" in self.combo_product_main_cat["values"]:
                    self.combo_product_main_cat.set("Hout")
                else:
                    self.combo_product_main_cat.current(0)
                with get_connection() as conn:
                    self._populate_product_sub_combo(conn)

        self._update_hout_panel_visibility()
        self._clear_product_extra_vars()
        self._update_product_extra_visibility()
        self._sync_hout_product_name()

    def _product_save(self) -> None:
        code = self.var_code.get().strip()
        self._sync_hout_product_name()
        name = self.var_name.get().strip()
        if not code or not name:
            messagebox.showwarning("Ontbrekend", "Vul minimaal code en naam in.")
            return
        unit = self._normalize_unit(self.var_unit.get())
        qd = self.txt_quote_desc.get("1.0", tk.END).strip()
        cat_id = self._product_resolve_category_id()
        if cat_id is None:
            messagebox.showwarning("Categorie", "Kies een hoofdcategorie.")
            return

        bom_count = 0
        if self._editing_product_id is not None:
            with get_connection() as conn:
                bom_count = int(
                    conn.execute(
                        "SELECT COUNT(*) FROM bom WHERE parent_id = ?",
                        (self._editing_product_id,),
                    ).fetchone()[0]
                )
        src = "bom" if bom_count > 0 else "manual"

        if src == "bom":
            price = 0.0
        else:
            sug = self._computed_sale_price_from_margin()
            if sug is not None:
                price = sug
            else:
                messagebox.showerror(
                    "Fout",
                    "Vul inkoopprijs en winstmarge in, of gebruik «Netto ophalen» (CSV).",
                )
                return

        try:
            with transaction() as conn:
                needs_mm_af = (
                    category_root_has_feature(conn, cat_id, FEATURE_DIKTE_BREEDTE_MM)
                    or category_root_has_feature(conn, cat_id, FEATURE_AFKORT_VERLIES)
                )
                if needs_mm_af:
                    try:
                        hd = self._parse_float_opt(self.var_hout_dikte.get())
                        hb = self._parse_float_opt(self.var_hout_breedte.get())
                        hav = self._parse_float_opt(self.var_hout_afkort_verlies.get())
                    except ValueError:
                        messagebox.showerror(
                            "Fout",
                            "Dikte, breedte en afkortverlies moeten leeg zijn of geldige getallen.",
                        )
                        return
                else:
                    hd = hb = hav = None

                if (
                    is_direct_sub_of_root_with_feature(
                        conn, cat_id, FEATURE_NAAM_SUB_MAAT
                    )
                    and category_root_has_feature(conn, cat_id, FEATURE_DIKTE_BREEDTE_MM)
                    and (hd is None or hb is None)
                ):
                    messagebox.showwarning(
                        "Subcategorie",
                        "Voor deze subgroep vul je dikte en breedte in (mm) — "
                        "de productnaam wordt daaruit samengesteld.",
                    )
                    return

                try:
                    pur = self._parse_float_opt(self.var_purchase_price.get())
                    mar = self._parse_float_opt(self.var_margin_pct.get())
                except ValueError:
                    messagebox.showerror(
                        "Fout",
                        "Inkoopprijs en winstmarge moeten leeg zijn of geldige getallen zijn.",
                    )
                    return

                min_o = max_o = d_los = wh = None
                sid: int | None = None
                s_art: str | None = None
                if category_root_has_feature(conn, cat_id, FEATURE_OPPERVLAKTE_M2):
                    try:
                        min_o = self._parse_float_opt(self.var_min_opp.get())
                        max_o = self._parse_float_opt(self.var_max_opp.get())
                    except ValueError:
                        messagebox.showerror(
                            "Fout",
                            "Min./max. oppervlakte moeten leeg zijn of geldige getallen.",
                        )
                        return
                if category_root_has_feature(conn, cat_id, FEATURE_DIKTE_LOS_MM):
                    try:
                        d_los = self._parse_float_opt(self.var_product_dikte_mm.get())
                    except ValueError:
                        messagebox.showerror(
                            "Fout",
                            "Dikte (mm) moet leeg zijn of een geldig getal.",
                        )
                        return
                if category_root_has_feature(conn, cat_id, FEATURE_LEVERANCIER):
                    lab = self.combo_supplier.get()
                    sid = self._supplier_label_to_id.get(lab)
                    s_art = self.var_supplier_article.get().strip() or None
                if category_root_has_feature(conn, cat_id, FEATURE_UREN_MINUTEN):
                    try:
                        s = self.var_work_minutes.get().strip().replace(",", ".")
                        total_m = float(s) if s else 0.0
                        if total_m < 0.0:
                            messagebox.showwarning(
                                "Minuten",
                                "Het aantal minuten mag niet negatief zijn.",
                            )
                            return
                        wh = total_m / 60.0
                    except ValueError:
                        messagebox.showerror(
                            "Fout",
                            "Vul een geldig getal in voor minuten (leeg = 0).",
                        )
                        return

                if self._editing_product_id is None:
                    conn.execute(
                        """
                        INSERT INTO products (
                            code, name, unit, unit_price, quote_description, price_source, category_id,
                            hout_dikte_mm, hout_breedte_mm, hout_afkort_verlies_pct, purchase_price, margin_pct,
                            min_oppervlakte_m2, max_oppervlakte_m2, product_dikte_mm,
                            supplier_id, supplier_article, work_hours
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            code,
                            name,
                            unit,
                            price,
                            qd,
                            src,
                            cat_id,
                            hd,
                            hb,
                            hav,
                            pur,
                            mar,
                            min_o,
                            max_o,
                            d_los,
                            sid,
                            s_art,
                            wh,
                        ),
                    )
                else:
                    conn.execute(
                        """
                        UPDATE products SET
                            code=?, name=?, unit=?, unit_price=?, quote_description=?, price_source=?, category_id=?,
                            hout_dikte_mm=?, hout_breedte_mm=?, hout_afkort_verlies_pct=?, purchase_price=?, margin_pct=?,
                            min_oppervlakte_m2=?, max_oppervlakte_m2=?, product_dikte_mm=?,
                            supplier_id=?, supplier_article=?, work_hours=?
                        WHERE id=?
                        """,
                        (
                            code,
                            name,
                            unit,
                            price,
                            qd,
                            src,
                            cat_id,
                            hd,
                            hb,
                            hav,
                            pur,
                            mar,
                            min_o,
                            max_o,
                            d_los,
                            sid,
                            s_art,
                            wh,
                            self._editing_product_id,
                        ),
                    )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return

        errs = validate_bom_acyclic()
        if errs:
            messagebox.showwarning("Stuklijst", "\n".join(errs))

        self._refresh_products_tree()
        self._refresh_pick_combo()
        self._refresh_dim_parent_combo()
        messagebox.showinfo("Opgeslagen", "Product opgeslagen.")

    def _product_delete(self) -> None:
        if self._editing_product_id is None:
            messagebox.showinfo("Geen selectie", "Selecteer eerst een product.")
            return
        if not messagebox.askyesno("Verwijderen", "Product echt verwijderen?"):
            return
        try:
            with transaction() as conn:
                conn.execute("DELETE FROM products WHERE id = ?", (self._editing_product_id,))
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return
        self._product_new()
        self._refresh_products_tree()
        self._refresh_dim_parent_combo()

    # --- BOM ---
    def _build_bom_tab(self) -> None:
        f = ttk.Frame(self._frm_bom)
        f.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        ttk.Label(f, text="Samengesteld product (ouder)").grid(row=0, column=0, sticky=tk.W)
        self.combo_bom_parent = ttk.Combobox(f, width=50, state="readonly")
        self.combo_bom_parent.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=4)
        self.combo_bom_parent.bind("<<ComboboxSelected>>", self._on_bom_parent_change)

        ttk.Label(f, text="Onderdeel toevoegen").grid(row=2, column=0, sticky=tk.W, pady=(12, 0))
        row3 = ttk.Frame(f)
        row3.grid(row=3, column=0, columnspan=3, sticky=tk.W)
        self.combo_bom_child = ttk.Combobox(row3, width=40, state="readonly")
        self.combo_bom_child.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row3, text="Aantal").pack(side=tk.LEFT)
        self.var_bom_qty = tk.StringVar(value="1")
        ttk.Entry(row3, textvariable=self.var_bom_qty, width=8).pack(side=tk.LEFT, padx=4)
        ttk.Button(row3, text="Toevoegen aan stuklijst", command=self._bom_add).pack(
            side=tk.LEFT, padx=8
        )

        self.tree_bom = ttk.Treeview(
            f, columns=("child_id", "code", "name", "qty"), show="headings", height=12
        )
        self.tree_bom.heading("child_id", text="ID")
        self.tree_bom.heading("code", text="Code")
        self.tree_bom.heading("name", text="Naam onderdeel")
        self.tree_bom.heading("qty", text="Aantal per stuk ouder")
        self.tree_bom.column("child_id", width=50)
        self.tree_bom.column("code", width=100)
        self.tree_bom.column("name", width=320)
        self.tree_bom.column("qty", width=140)
        self.tree_bom.grid(row=4, column=0, columnspan=3, sticky=tk.NSEW, pady=8)
        f.rowconfigure(4, weight=1)
        f.columnconfigure(0, weight=1)

        bf = ttk.Frame(f)
        bf.grid(row=5, column=0, columnspan=3, sticky=tk.W)
        ttk.Button(bf, text="Verwijder geselecteerde regel", command=self._bom_remove).pack(
            side=tk.LEFT, padx=4
        )

        self._bom_parent_id: int | None = None

    def _refresh_bom_parent_combo(self) -> None:
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM products ORDER BY name"
            ).fetchall()
        self._product_labels = {f"{r['code']} — {r['name']}": r["id"] for r in rows}
        labels = list(self._product_labels.keys())
        self.combo_bom_parent["values"] = labels
        self.combo_bom_child["values"] = labels
        if labels and not self.combo_bom_parent.get():
            self.combo_bom_parent.current(0)
            self._on_bom_parent_change()

    def _on_bom_parent_change(self, _evt=None) -> None:
        label = self.combo_bom_parent.get()
        self._bom_parent_id = self._product_labels.get(label)
        self._refresh_bom_tree()

    def _refresh_bom_tree(self) -> None:
        for i in self.tree_bom.get_children():
            self.tree_bom.delete(i)
        if not self._bom_parent_id:
            return
        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT b.child_id, p.code, p.name, b.qty
                FROM bom b
                JOIN products p ON p.id = b.child_id
                WHERE b.parent_id = ?
                ORDER BY p.code
                """,
                (self._bom_parent_id,),
            ).fetchall()
        for r in rows:
            self.tree_bom.insert(
                "",
                tk.END,
                values=(r["child_id"], r["code"], r["name"], f"{r['qty']:g}"),
            )

    def _bom_add(self) -> None:
        if not self._bom_parent_id:
            return
        clabel = self.combo_bom_child.get()
        cid = self._product_labels.get(clabel)
        if not cid:
            messagebox.showwarning("Kies", "Kies een onderdeel.")
            return
        if cid == self._bom_parent_id:
            messagebox.showerror("Fout", "Een product kan niet zichzelf als onderdeel hebben.")
            return
        try:
            qty = float(self.var_bom_qty.get().replace(",", "."))
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Geef een geldig aantal > 0.")
            return
        try:
            with transaction() as conn:
                conn.execute(
                    """INSERT INTO bom (parent_id, child_id, qty) VALUES (?, ?, ?)
                       ON CONFLICT(parent_id, child_id) DO UPDATE SET qty = excluded.qty""",
                    (self._bom_parent_id, cid, qty),
                )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return
        self._refresh_bom_tree()
        self._refresh_products_tree()

    def _bom_remove(self) -> None:
        sel = self.tree_bom.selection()
        if not sel or not self._bom_parent_id:
            return
        cid = int(self.tree_bom.item(sel[0], "values")[0])
        with transaction() as conn:
            conn.execute(
                "DELETE FROM bom WHERE parent_id = ? AND child_id = ?",
                (self._bom_parent_id, cid),
            )
        self._refresh_bom_tree()
        self._refresh_products_tree()

    # --- Maatregels (B×H in mm) ---
    def _build_dimension_rules_tab(self) -> None:
        f = ttk.Frame(self._frm_dim)
        f.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        hint = (
            "Per product (bijv. draaikiep raam): definieer rechthoekige bereiken in millimeters. "
            "Leeg laten bij min/max = geen grens aan die kant. "
            "Alle passende regels tellen op bij dezelfde maat (zoals cellen in een B×H-tabel). "
            "Vaste stuklijst (BOM) blijft altijd gelden; maatregels voegen extra onderdelen toe."
        )
        ttk.Label(f, text=hint, wraplength=920, justify=tk.LEFT).grid(
            row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8)
        )

        ttk.Label(f, text="Product (ouder)").grid(row=1, column=0, sticky=tk.W)
        self.combo_dim_parent = ttk.Combobox(f, width=52, state="readonly")
        self.combo_dim_parent.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=4)
        self.combo_dim_parent.bind("<<ComboboxSelected>>", self._on_dim_parent_change)

        self.tree_dim = ttk.Treeview(
            f,
            columns=("rid", "ccode", "cname", "wmin", "wmax", "hmin", "hmax", "qty"),
            show="headings",
            height=11,
            selectmode="browse",
        )
        for c, t, w in [
            ("rid", "ID", 45),
            ("ccode", "Onderdeel", 100),
            ("cname", "Naam", 220),
            ("wmin", "B min", 70),
            ("wmax", "B max", 70),
            ("hmin", "H min", 70),
            ("hmax", "H max", 70),
            ("qty", "Aantal", 70),
        ]:
            self.tree_dim.heading(c, text=t)
            self.tree_dim.column(c, width=w)
        sb = ttk.Scrollbar(f, orient=tk.VERTICAL, command=self.tree_dim.yview)
        self.tree_dim.configure(yscrollcommand=sb.set)
        self.tree_dim.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW, pady=8)
        sb.grid(row=3, column=2, sticky=tk.NS, pady=8)
        f.rowconfigure(3, weight=1)
        f.columnconfigure(0, weight=1)

        self.tree_dim.bind("<<TreeviewSelect>>", self._on_dim_rule_select)

        form = ttk.LabelFrame(f, text="Regel")
        form.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=8)

        r = 0
        ttk.Label(form, text="Onderdeel").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.combo_dim_child = ttk.Combobox(form, width=44, state="readonly")
        self.combo_dim_child.grid(row=r, column=1, columnspan=3, sticky=tk.W, padx=4, pady=2)
        r += 1
        ttk.Label(form, text="Breedte min mm (leeg = −∞)").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_dim_wmin = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_dim_wmin, width=12).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        ttk.Label(form, text="Breedte max mm (leeg = +∞)").grid(
            row=r, column=2, sticky=tk.W, padx=8
        )
        self.var_dim_wmax = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_dim_wmax, width=12).grid(
            row=r, column=3, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(form, text="Hoogte min mm").grid(row=r, column=0, sticky=tk.W, padx=4, pady=2)
        self.var_dim_hmin = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_dim_hmin, width=12).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        ttk.Label(form, text="Hoogte max mm").grid(row=r, column=2, sticky=tk.W, padx=8)
        self.var_dim_hmax = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_dim_hmax, width=12).grid(
            row=r, column=3, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(form, text="Aantal per stuk ouder").grid(
            row=r, column=0, sticky=tk.W, padx=4, pady=2
        )
        self.var_dim_qty = tk.StringVar(value="1")
        ttk.Entry(form, textvariable=self.var_dim_qty, width=12).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )

        bf = ttk.Frame(form)
        bf.grid(row=r + 1, column=0, columnspan=4, pady=10, sticky=tk.W)
        ttk.Button(bf, text="Nieuwe regel", command=self._dim_rule_new).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(bf, text="Opslaan", command=self._dim_rule_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="Verwijderen", command=self._dim_rule_delete).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(bf, text="Vernieuwen", command=self._refresh_dim_rules_ui).pack(
            side=tk.LEFT, padx=4
        )

        self._dim_parent_id: int | None = None
        self._editing_dim_rule_id: int | None = None

    @staticmethod
    def _parse_float_opt(s: str) -> float | None:
        s = (s or "").strip()
        if not s:
            return None
        return float(s.replace(",", "."))

    def _refresh_dim_parent_combo(self) -> None:
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM products ORDER BY name"
            ).fetchall()
        self._dim_product_labels = {f"{r['code']} — {r['name']}": r["id"] for r in rows}
        labels = list(self._dim_product_labels.keys())
        self.combo_dim_parent["values"] = labels
        self.combo_dim_child["values"] = labels

    def _on_dim_parent_change(self, _evt=None) -> None:
        label = self.combo_dim_parent.get()
        self._dim_parent_id = self._dim_product_labels.get(label)
        self._dim_rule_new()
        self._refresh_dim_tree()

    def _refresh_dim_tree(self) -> None:
        for i in self.tree_dim.get_children():
            self.tree_dim.delete(i)
        if not self._dim_parent_id:
            return

        def fmt(v: float | None) -> str:
            if v is None:
                return "—"
            return f"{v:g}"

        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT d.id, d.child_id, d.qty, d.w_min, d.w_max, d.h_min, d.h_max,
                       p.code AS ccode, p.name AS cname
                FROM dimension_bom_rules d
                JOIN products p ON p.id = d.child_id
                WHERE d.parent_id = ?
                ORDER BY d.w_min, d.h_min, p.code
                """,
                (self._dim_parent_id,),
            ).fetchall()
        for row in rows:
            self.tree_dim.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["ccode"],
                    row["cname"],
                    fmt(row["w_min"]),
                    fmt(row["w_max"]),
                    fmt(row["h_min"]),
                    fmt(row["h_max"]),
                    f"{row['qty']:g}",
                ),
            )

    def _on_dim_rule_select(self, _evt=None) -> None:
        sel = self.tree_dim.selection()
        if not sel or not self._dim_parent_id:
            return
        rid = int(self.tree_dim.item(sel[0], "values")[0])
        self._editing_dim_rule_id = rid
        with get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM dimension_bom_rules WHERE id = ?", (rid,)
            ).fetchone()
            if not row:
                return
            child = conn.execute(
                "SELECT code, name FROM products WHERE id = ?", (row["child_id"],)
            ).fetchone()
        if child:
            self.combo_dim_child.set(f"{child['code']} — {child['name']}")
        self.var_dim_wmin.set("" if row["w_min"] is None else str(row["w_min"]))
        self.var_dim_wmax.set("" if row["w_max"] is None else str(row["w_max"]))
        self.var_dim_hmin.set("" if row["h_min"] is None else str(row["h_min"]))
        self.var_dim_hmax.set("" if row["h_max"] is None else str(row["h_max"]))
        self.var_dim_qty.set(str(row["qty"]))

    def _dim_rule_new(self) -> None:
        self._editing_dim_rule_id = None
        self.var_dim_wmin.set("")
        self.var_dim_wmax.set("")
        self.var_dim_hmin.set("")
        self.var_dim_hmax.set("")
        self.var_dim_qty.set("1")
        if self.combo_dim_child["values"]:
            self.combo_dim_child.set("")

    def _dim_rule_save(self) -> None:
        if not self._dim_parent_id:
            messagebox.showwarning("Kies", "Kies eerst een product (ouder).")
            return
        clabel = self.combo_dim_child.get()
        cid = self._dim_product_labels.get(clabel)
        if not cid:
            messagebox.showwarning("Kies", "Kies een onderdeel.")
            return
        if cid == self._dim_parent_id:
            messagebox.showerror("Fout", "Ouder en onderdeel mogen niet hetzelfde zijn.")
            return
        try:
            qty = float(self.var_dim_qty.get().replace(",", "."))
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Ongeldig aantal.")
            return
        try:
            wmin = self._parse_float_opt(self.var_dim_wmin.get())
            wmax = self._parse_float_opt(self.var_dim_wmax.get())
            hmin = self._parse_float_opt(self.var_dim_hmin.get())
            hmax = self._parse_float_opt(self.var_dim_hmax.get())
        except ValueError:
            messagebox.showerror("Fout", "Ongeldige getallen bij min/max.")
            return
        if wmin is not None and wmax is not None and wmin > wmax:
            messagebox.showerror("Fout", "Breedte min mag niet groter zijn dan max.")
            return
        if hmin is not None and hmax is not None and hmin > hmax:
            messagebox.showerror("Fout", "Hoogte min mag niet groter zijn dan max.")
            return

        try:
            with transaction() as conn:
                if self._editing_dim_rule_id is None:
                    conn.execute(
                        """
                        INSERT INTO dimension_bom_rules
                        (parent_id, child_id, qty, w_min, w_max, h_min, h_max)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (self._dim_parent_id, cid, qty, wmin, wmax, hmin, hmax),
                    )
                else:
                    conn.execute(
                        """
                        UPDATE dimension_bom_rules SET
                          child_id = ?, qty = ?, w_min = ?, w_max = ?, h_min = ?, h_max = ?
                        WHERE id = ? AND parent_id = ?
                        """,
                        (cid, qty, wmin, wmax, hmin, hmax, self._editing_dim_rule_id, self._dim_parent_id),
                    )
        except Exception as e:
            messagebox.showerror("Database", str(e))
            return

        self._refresh_dim_tree()
        self._dim_rule_new()
        messagebox.showinfo("Opgeslagen", "Maatregel opgeslagen.")

    def _dim_rule_delete(self) -> None:
        if self._editing_dim_rule_id is None:
            messagebox.showinfo("Geen selectie", "Selecteer een regel in de lijst.")
            return
        if not messagebox.askyesno("Verwijderen", "Deze maatregel verwijderen?"):
            return
        with transaction() as conn:
            conn.execute(
                "DELETE FROM dimension_bom_rules WHERE id = ? AND parent_id = ?",
                (self._editing_dim_rule_id, self._dim_parent_id),
            )
        self._dim_rule_new()
        self._refresh_dim_tree()

    def _refresh_dim_rules_ui(self) -> None:
        self._refresh_dim_parent_combo()
        if self.combo_dim_parent.get():
            self._on_dim_parent_change()
        else:
            self._refresh_dim_tree()

    # --- Offertes ---
    def _build_quotes_tab(self) -> None:
        left = ttk.Frame(self._frm_quotes)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=4)
        ttk.Label(left, text="Offertes").pack(anchor=tk.W)
        self.list_quotes = tk.Listbox(left, width=36, height=22)
        self.list_quotes.pack(fill=tk.Y, expand=True)
        self.list_quotes.bind("<<ListboxSelect>>", self._on_quote_select)

        bf = ttk.Frame(left)
        bf.pack(fill=tk.X, pady=6)
        ttk.Button(bf, text="Nieuwe offerte", command=self._quote_new).pack(fill=tk.X)
        ttk.Button(bf, text="Offerte exporteren (HTML)", command=self._quote_export).pack(
            fill=tk.X, pady=4
        )

        right = ttk.LabelFrame(self._frm_quotes, text="Offertegegevens")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.var_q_num = tk.StringVar()
        self.var_q_cust = tk.StringVar()
        self.var_q_email = tk.StringVar()
        self.var_q_vat = tk.StringVar(value="21")
        r = 0
        ttk.Label(right, text="Offertenummer").grid(row=r, column=0, sticky=tk.W)
        ttk.Entry(right, textvariable=self.var_q_num, width=28, state="readonly").grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(right, text="Klant / project").grid(row=r, column=0, sticky=tk.W)
        ttk.Entry(right, textvariable=self.var_q_cust, width=48).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(right, text="Adres").grid(row=r, column=0, sticky=tk.NW)
        self.txt_q_addr = tk.Text(right, width=48, height=3, wrap=tk.WORD)
        self.txt_q_addr.grid(row=r, column=1, sticky=tk.W, padx=4)
        r += 1
        ttk.Label(right, text="E-mail").grid(row=r, column=0, sticky=tk.W)
        ttk.Entry(right, textvariable=self.var_q_email, width=48).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(right, text="BTW %").grid(row=r, column=0, sticky=tk.W)
        ttk.Entry(right, textvariable=self.var_q_vat, width=8).grid(
            row=r, column=1, sticky=tk.W, padx=4
        )
        r += 1
        ttk.Label(right, text="Opmerkingen").grid(row=r, column=0, sticky=tk.NW)
        self.txt_q_notes = tk.Text(right, width=48, height=3, wrap=tk.WORD)
        self.txt_q_notes.grid(row=r, column=1, sticky=tk.W, padx=4)
        r += 1

        ttk.Button(right, text="Koptekst opslaan", command=self._quote_save_header).grid(
            row=r, column=1, sticky=tk.W, pady=6
        )
        r += 1

        ttk.Label(right, text="Regels", font=("", 10, "bold")).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=(12, 4)
        )
        r += 1

        self.tree_ql = ttk.Treeview(
            right,
            columns=("lid", "pname", "qty", "bw", "bh", "unitp", "disc", "line"),
            show="headings",
            height=10,
        )
        for c, t, w in [
            ("lid", "Regel", 40),
            ("pname", "Product", 200),
            ("qty", "Aantal", 50),
            ("bw", "B mm", 60),
            ("bh", "H mm", 60),
            ("unitp", "Prijs/st", 68),
            ("disc", "Kort %", 50),
            ("line", "Regel €", 72),
        ]:
            self.tree_ql.heading(c, text=t)
            self.tree_ql.column(c, width=w)
        self.tree_ql.grid(row=r, column=0, columnspan=2, sticky=tk.NSEW, pady=4)
        self.tree_ql.bind("<<TreeviewSelect>>", self._on_quote_line_select)
        right.rowconfigure(r, weight=1)
        right.columnconfigure(1, weight=1)

        addf = ttk.Frame(right)
        addf.grid(row=r + 1, column=0, columnspan=2, sticky=tk.W, pady=4)
        self.combo_q_product = ttk.Combobox(addf, width=38, state="readonly")
        self.combo_q_product.pack(side=tk.LEFT, padx=(0, 6))
        self.var_q_line_qty = tk.StringVar(value="1")
        ttk.Entry(addf, textvariable=self.var_q_line_qty, width=5).pack(side=tk.LEFT)
        ttk.Button(addf, text="Regel toevoegen", command=self._quote_add_line).pack(
            side=tk.LEFT, padx=8
        )
        ttk.Button(addf, text="Regel verwijderen", command=self._quote_remove_line).pack(
            side=tk.LEFT
        )

        addf2 = ttk.Frame(right)
        addf2.grid(row=r + 2, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(addf2, text="B mm").pack(side=tk.LEFT, padx=(0, 4))
        self.var_q_line_w = tk.StringVar()
        ttk.Entry(addf2, textvariable=self.var_q_line_w, width=8).pack(side=tk.LEFT)
        ttk.Label(addf2, text="H mm").pack(side=tk.LEFT, padx=(8, 4))
        self.var_q_line_h = tk.StringVar()
        ttk.Entry(addf2, textvariable=self.var_q_line_h, width=8).pack(side=tk.LEFT)
        ttk.Label(addf2, text="Korting %").pack(side=tk.LEFT, padx=(12, 4))
        self.var_q_line_disc = tk.StringVar(value="0")
        ttk.Entry(addf2, textvariable=self.var_q_line_disc, width=5).pack(side=tk.LEFT)
        ttk.Button(addf2, text="Regel bijwerken", command=self._quote_update_line).pack(
            side=tk.LEFT, padx=12
        )
        ttk.Label(
            addf2,
            text="(B×H voor maatregels; leeg = alleen vaste BOM-prijs)",
            foreground="#555",
            font=("", 9),
        ).pack(side=tk.LEFT, padx=8)

        self._current_quote_id: int | None = None
        self._editing_quote_line_id: int | None = None

    def _refresh_quote_product_combo(self) -> None:
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM products ORDER BY name"
            ).fetchall()
        self._quote_product_map = {f"{r['code']} — {r['name']}": r["id"] for r in rows}
        self.combo_q_product["values"] = list(self._quote_product_map.keys())

    def _refresh_quotes(self) -> None:
        self.list_quotes.delete(0, tk.END)
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, quote_number, customer_name, created_at FROM quotes ORDER BY id DESC"
            ).fetchall()
        self._quote_list_ids = []
        for row in rows:
            self._quote_list_ids.append(row["id"])
            self.list_quotes.insert(
                tk.END,
                f"{row['quote_number']} — {row['customer_name'] or '(geen naam)'} ({row['created_at'][:10]})",
            )

    def _on_quote_select(self, _evt=None) -> None:
        sel = self.list_quotes.curselection()
        if not sel:
            return
        qid = self._quote_list_ids[sel[0]]
        self._load_quote_by_id(qid)

    def _load_quote_by_id(self, qid: int) -> None:
        self._current_quote_id = qid
        self._editing_quote_line_id = None
        with get_connection() as conn:
            q = conn.execute("SELECT * FROM quotes WHERE id = ?", (qid,)).fetchone()
        if not q:
            return
        self.var_q_num.set(q["quote_number"])
        self.var_q_cust.set(q["customer_name"] or "")
        self.txt_q_addr.delete("1.0", tk.END)
        self.txt_q_addr.insert("1.0", q["customer_address"] or "")
        self.var_q_email.set(q["customer_email"] or "")
        self.var_q_vat.set(str(q["vat_rate"]))
        self.txt_q_notes.delete("1.0", tk.END)
        self.txt_q_notes.insert("1.0", q["notes"] or "")
        self._refresh_quote_lines()

    def _refresh_quote_lines(self) -> None:
        for i in self.tree_ql.get_children():
            self.tree_ql.delete(i)
        if not self._current_quote_id:
            return
        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT ql.id, p.name, ql.qty, ql.unit_price_override, ql.line_discount_pct, ql.product_id,
                       ql.width_mm, ql.height_mm
                FROM quote_lines ql
                JOIN products p ON p.id = ql.product_id
                WHERE ql.quote_id = ?
                ORDER BY ql.sort_order, ql.id
                """,
                (self._current_quote_id,),
            ).fetchall()

        for row in rows:
            wm = row["width_mm"] if "width_mm" in row.keys() else None
            hm = row["height_mm"] if "height_mm" in row.keys() else None
            unitp = quote_line_unit_price(
                row["product_id"], row["unit_price_override"], wm, hm
            )
            qty = float(row["qty"])
            disc = float(row["line_discount_pct"]) / 100.0
            line = unitp * qty * (1.0 - disc)
            ovr = row["unit_price_override"]
            up_show = f"{unitp:.2f}" + (" *" if ovr is not None else "")
            bw_show = "" if wm is None else f"{wm:g}"
            bh_show = "" if hm is None else f"{hm:g}"
            self.tree_ql.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["name"],
                    f"{qty:g}",
                    bw_show,
                    bh_show,
                    up_show,
                    f"{row['line_discount_pct']:g}",
                    f"{line:.2f}",
                ),
            )

    def _quote_new(self) -> None:
        num = next_quote_number()
        now = datetime.now().isoformat(timespec="seconds")
        with transaction() as conn:
            cur = conn.execute(
                """INSERT INTO quotes (created_at, quote_number, customer_name, customer_address, customer_email, notes, vat_rate)
                   VALUES (?, ?, '', '', '', '', 21)""",
                (now, num),
            )
            qid = cur.lastrowid
        self._refresh_quotes()
        if self.list_quotes.size() > 0:
            self.list_quotes.selection_clear(0, tk.END)
            self.list_quotes.selection_set(0)
            self.list_quotes.activate(0)
        self._load_quote_by_id(qid)

    def _quote_save_header(self) -> None:
        if not self._current_quote_id:
            messagebox.showinfo("Geen offerte", "Maak of selecteer een offerte.")
            return
        try:
            vat = float(self.var_q_vat.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Fout", "Ongeldig BTW-percentage.")
            return
        with transaction() as conn:
            conn.execute(
                """UPDATE quotes SET customer_name=?, customer_address=?, customer_email=?, notes=?, vat_rate=?
                   WHERE id=?""",
                (
                    self.var_q_cust.get().strip(),
                    self.txt_q_addr.get("1.0", tk.END).strip(),
                    self.var_q_email.get().strip(),
                    self.txt_q_notes.get("1.0", tk.END).strip(),
                    vat,
                    self._current_quote_id,
                ),
            )
        messagebox.showinfo("Opgeslagen", "Offerte-kop opgeslagen.")
        self._refresh_quotes()

    def _on_quote_line_select(self, _evt=None) -> None:
        sel = self.tree_ql.selection()
        if not sel or not self._current_quote_id:
            self._editing_quote_line_id = None
            return
        lid = int(self.tree_ql.item(sel[0], "values")[0])
        self._editing_quote_line_id = lid
        with get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM quote_lines WHERE id = ?", (lid,)
            ).fetchone()
            if not row:
                return
            p = conn.execute(
                "SELECT code, name FROM products WHERE id = ?", (row["product_id"],)
            ).fetchone()
        if p:
            self.combo_q_product.set(f"{p['code']} — {p['name']}")
        self.var_q_line_qty.set(str(row["qty"]))
        self.var_q_line_disc.set(str(row["line_discount_pct"]))
        wm = row["width_mm"] if "width_mm" in row.keys() else None
        hm = row["height_mm"] if "height_mm" in row.keys() else None
        self.var_q_line_w.set("" if wm is None else str(wm))
        self.var_q_line_h.set("" if hm is None else str(hm))

    def _quote_add_line(self) -> None:
        if not self._current_quote_id:
            messagebox.showinfo("Geen offerte", "Maak of selecteer een offerte.")
            return
        label = self.combo_q_product.get()
        pid = self._quote_product_map.get(label)
        if not pid:
            messagebox.showwarning("Kies", "Kies een product.")
            return
        try:
            qty = float(self.var_q_line_qty.get().replace(",", "."))
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Ongeldig aantal.")
            return
        try:
            w_mm = self._parse_float_opt(self.var_q_line_w.get())
            h_mm = self._parse_float_opt(self.var_q_line_h.get())
        except ValueError:
            messagebox.showerror("Fout", "Ongeldige maat (gebruik getallen in mm).")
            return
        if (w_mm is None) != (h_mm is None):
            messagebox.showwarning(
                "Maat",
                "Vul breedte én hoogte in (mm), of laat beide leeg voor alleen de vaste BOM-prijs.",
            )
            return
        try:
            disc = float(self.var_q_line_disc.get().replace(",", "."))
            if disc < 0 or disc > 100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Korting moet tussen 0 en 100 liggen.")
            return
        with transaction() as conn:
            mx = conn.execute(
                "SELECT COALESCE(MAX(sort_order),0) FROM quote_lines WHERE quote_id=?",
                (self._current_quote_id,),
            ).fetchone()[0]
            conn.execute(
                """INSERT INTO quote_lines (quote_id, product_id, qty, unit_price_override, line_discount_pct, sort_order, width_mm, height_mm)
                   VALUES (?, ?, ?, NULL, ?, ?, ?, ?)""",
                (self._current_quote_id, pid, qty, disc, mx + 1, w_mm, h_mm),
            )
        self._refresh_quote_lines()

    def _quote_update_line(self) -> None:
        if not self._current_quote_id or not self._editing_quote_line_id:
            messagebox.showinfo("Geen regel", "Selecteer een regel in de lijst.")
            return
        try:
            qty = float(self.var_q_line_qty.get().replace(",", "."))
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Ongeldig aantal.")
            return
        try:
            w_mm = self._parse_float_opt(self.var_q_line_w.get())
            h_mm = self._parse_float_opt(self.var_q_line_h.get())
        except ValueError:
            messagebox.showerror("Fout", "Ongeldige maat (gebruik getallen in mm).")
            return
        if (w_mm is None) != (h_mm is None):
            messagebox.showwarning(
                "Maat",
                "Vul breedte én hoogte in (mm), of laat beide leeg.",
            )
            return
        try:
            disc = float(self.var_q_line_disc.get().replace(",", "."))
            if disc < 0 or disc > 100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Korting moet tussen 0 en 100 liggen.")
            return
        with transaction() as conn:
            conn.execute(
                """UPDATE quote_lines SET qty=?, width_mm=?, height_mm=?, line_discount_pct=?
                   WHERE id=? AND quote_id=?""",
                (
                    qty,
                    w_mm,
                    h_mm,
                    disc,
                    self._editing_quote_line_id,
                    self._current_quote_id,
                ),
            )
        self._refresh_quote_lines()

    def _quote_remove_line(self) -> None:
        sel = self.tree_ql.selection()
        if not sel or not self._current_quote_id:
            return
        lid = int(self.tree_ql.item(sel[0], "values")[0])
        with transaction() as conn:
            conn.execute("DELETE FROM quote_lines WHERE id = ?", (lid,))
        self._editing_quote_line_id = None
        self._refresh_quote_lines()

    def _quote_export(self) -> None:
        if not self._current_quote_id:
            messagebox.showinfo("Geen offerte", "Selecteer een offerte.")
            return
        self._quote_save_header()
        path = export_quote_html(self._current_quote_id)
        messagebox.showinfo("Export", f"HTML opgeslagen:\n{path}")
        webbrowser.open(path.as_uri())

    # --- Materiaallijst (picking) ---
    def _build_pick_tab(self) -> None:
        f = ttk.Frame(self._frm_pick)
        f.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        ttk.Label(
            f,
            text="Kies een product om onderdelen te tonen (BOM + optioneel maatregels). Vul B×H in mm in als het product maatregels heeft.",
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)
        row = ttk.Frame(f)
        row.pack(fill=tk.X, pady=8)
        self.combo_pick = ttk.Combobox(row, width=44, state="readonly")
        self.combo_pick.pack(side=tk.LEFT, padx=(0, 8))
        self.var_pick_qty = tk.StringVar(value="1")
        ttk.Entry(row, textvariable=self.var_pick_qty, width=6).pack(side=tk.LEFT, padx=4)
        ttk.Label(row, text="B mm").pack(side=tk.LEFT, padx=(8, 2))
        self.var_pick_w = tk.StringVar()
        ttk.Entry(row, textvariable=self.var_pick_w, width=8).pack(side=tk.LEFT)
        ttk.Label(row, text="H mm").pack(side=tk.LEFT, padx=(6, 2))
        self.var_pick_h = tk.StringVar()
        ttk.Entry(row, textvariable=self.var_pick_h, width=8).pack(side=tk.LEFT)
        ttk.Button(row, text="Uitrollen", command=self._pick_expand).pack(side=tk.LEFT, padx=12)

        self.tree_pick = ttk.Treeview(
            f,
            columns=("code", "name", "unit", "qty", "unitp", "tot"),
            show="headings",
            height=18,
        )
        for c, t, w in [
            ("code", "Code", 100),
            ("name", "Onderdeel", 280),
            ("unit", "Ehd", 50),
            ("qty", "Aantal", 80),
            ("unitp", "Prijs/st", 80),
            ("tot", "Totaal", 90),
        ]:
            self.tree_pick.heading(c, text=t)
            self.tree_pick.column(c, width=w)
        self.tree_pick.pack(fill=tk.BOTH, expand=True, pady=8)

    def _refresh_pick_combo(self) -> None:
        with get_connection() as conn:
            rows = conn.execute(
                "SELECT id, code, name FROM products ORDER BY name"
            ).fetchall()
        self._pick_map = {f"{r['code']} — {r['name']}": r["id"] for r in rows}
        self.combo_pick["values"] = list(self._pick_map.keys())

    def _pick_expand(self) -> None:
        self._refresh_pick_combo()
        label = self.combo_pick.get()
        pid = self._pick_map.get(label)
        if not pid:
            messagebox.showwarning("Kies", "Kies een product.")
            return
        try:
            m = float(self.var_pick_qty.get().replace(",", "."))
            if m <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Fout", "Ongeldig aantal.")
            return
        w_mm = self._parse_float_opt(self.var_pick_w.get())
        h_mm = self._parse_float_opt(self.var_pick_h.get())
        if (w_mm is None) != (h_mm is None):
            messagebox.showwarning(
                "Maat",
                "Vul breedte én hoogte in (mm), of laat beide leeg voor alleen de vaste BOM.",
            )
            return
        for i in self.tree_pick.get_children():
            self.tree_pick.delete(i)
        try:
            lines = expand_bom_with_dimensions(pid, m, w_mm, h_mm)
        except ValueError as e:
            messagebox.showerror("Stuklijst", str(e))
            return
        # Aggregeer zelfde product-id (voor netheid)
        agg: dict[int, tuple] = {}
        for ln in lines:
            if ln.product_id in agg:
                o = agg[ln.product_id]
                nq = o[0] + ln.qty
                agg[ln.product_id] = (nq, ln)
            else:
                agg[ln.product_id] = (ln.qty, ln)
        for _pid, (qty, ln) in sorted(agg.items(), key=lambda x: x[1].code):
            tot = qty * ln.unit_price
            self.tree_pick.insert(
                "",
                tk.END,
                values=(
                    ln.code,
                    ln.name,
                    ln.unit,
                    f"{qty:g}",
                    f"{ln.unit_price:.2f}",
                    f"{tot:.2f}",
                ),
            )


def run() -> None:
    init_db()
    seed_demo_data()
    app = KozijnenApp()
    app._refresh_pick_combo()
    app.mainloop()
