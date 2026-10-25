#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistem Local de Gestiune și Catalogare Componente Electronice - v2.0
Cu suport integrat pentru Datasheet-uri PDF (Local + Auto-Discovery + Web Fallback)
Zero dependențe externe (100% Python Standard Library) - compatibil Windows & Linux.
"""

from __future__ import annotations
import csv
import os
import shutil
import sqlite3
import subprocess
import sys
import urllib.parse
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from contextlib import contextmanager
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterator, Any

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "inventar_atelier.db"
DATASHEETS_DIR = BASE_DIR / "datasheets"

CATEGORIES = [
    "Semiconductori (Tranzistoare / MOSFET / IGBT)",
    "Integrate SMPS / Power Management / LDO",
    "Auto / ECU (CAN, LIN, Drivere Injecție/Relee)",
    "Audio / Amplificatoare / DSP / Optice",
    "Microcontrolere / Memorii (EEPROM, Flash, RAM)",
    "Diode / Zener / TVS / Punți Redresoare",
    "Pasive (Condensatoare, Rezistențe, Inductanțe)",
    "Electromecanice (Relee, Mufe, Switch-uri, Senzori)",
    "Placă Donatoare / Modul Întreg",
]

STATES = ["Nou", "Recuperat (Testat OK)", "Din Placă Donatoare"]


@dataclass(slots=True)
class Componenta:
    locatie: str
    part_number: str
    smd_marking: str
    categorie: str
    capsula: str
    specificatii: str
    stare: str
    cantitate: int
    provenienta_note: str
    datasheet: str = ""
    id: int | None = None


class DatabaseManager:
    """Gestiune conexiune SQLite3 cu WAL, auto-migrare schemă și căutare rapidă."""

    def __init__(self, db_path: Path = DB_PATH) -> None:
        self.db_path = db_path
        DATASHEETS_DIR.mkdir(parents=True, exist_ok=True)
        self._init_and_migrate_db()

    @contextmanager
    def session(self) -> Iterator[sqlite3.Cursor]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA foreign_keys = ON;")
        cursor = conn.cursor()
        try:
            yield cursor
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            cursor.close()
            conn.close()

    def _init_and_migrate_db(self) -> None:
        with self.session() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS componente (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    locatie TEXT NOT NULL,
                    part_number TEXT NOT NULL,
                    smd_marking TEXT DEFAULT '',
                    categorie TEXT NOT NULL,
                    capsula TEXT DEFAULT '',
                    specificatii TEXT DEFAULT '',
                    stare TEXT NOT NULL DEFAULT 'Nou',
                    cantitate INTEGER NOT NULL DEFAULT 1,
                    provenienta_note TEXT DEFAULT '',
                    datasheet TEXT DEFAULT '',
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                """
            )
            # Auto-migrare: dacă baza de date provine din v1, adăugăm coloana 'datasheet' fără pierdere de date
            existing_cols = {row["name"] for row in cur.execute("PRAGMA table_info(componente);")}
            if "datasheet" not in existing_cols:
                cur.execute("ALTER TABLE componente ADD COLUMN datasheet TEXT DEFAULT '';")

            cur.execute("CREATE INDEX IF NOT EXISTS idx_locatie ON componente(locatie);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_pn ON componente(part_number);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_smd ON componente(smd_marking);")

            if (count := cur.execute("SELECT COUNT(*) FROM componente;").fetchone()[0]) == 0:
                seed_data = [
                    Componenta("P1-001", "MMBT3904", "AM / 1AM", CATEGORIES[0], "SOT-23", "NPN 40V 200mA 300MHz", "Nou", 40, "Echivalent uzual comandă relee/opto (Fumis RCY-1)", "datasheets/MMBT3904.pdf"),
                    Componenta("P1-002", "TPS5120DBT", "TPS5120", CATEGORIES[1], "TSSOP-30", "Dual Buck Controller 5V / 3.3V", "Recuperat (Testat OK)", 2, "Clarion QY BASE PWB IC902", "datasheets/TPS5120DBT.pdf"),
                    Componenta("P1-003", "R1514S037B", "R1514S", CATEGORIES[1], "HSOP-6", "LDO Regulator 3.6V/3.7V Standby", "Recuperat (Testat OK)", 2, "Clarion QY Standby BU-3.3V IC905", ""),
                    Componenta("S1-A1", "TB2926HQ", "TB2926HQ", CATEGORIES[3], "HZIP-25", "4x45W BTL Audio Power Amp", "Recuperat (Testat OK)", 1, "Clarion QY Audio Final IC901", ""),
                    Componenta("P1-004", "BA5830FM", "BA5830FM", CATEGORIES[3], "HSOP-28", "4-Channel CD/DVD Motor Driver 5V-8V", "Recuperat (Testat OK)", 3, "Mecanisme optice Clarion / Becker", ""),
                    Componenta("P1-005", "TJA1050", "A1050", CATEGORIES[2], "SOIC-8", "High-Speed CAN Transceiver 1Mbps 5V", "Nou", 15, "Stoc reparații ECU / CAN Bus", "datasheets/TJA1050.pdf"),
                    Componenta("S2-B3", "STP10NK60Z", "10NK60Z", CATEGORIES[0], "TO-220", "MOSFET N-Ch 600V 10A 0.65R Zener Protected", "Nou", 8, "Surse SMPS primar", ""),
                ]
                for comp in seed_data:
                    self.add(comp, _cursor=cur)

    def add(self, comp: Componenta, _cursor: sqlite3.Cursor | None = None) -> int:
        sql = """
            INSERT INTO componente (
                locatie, part_number, smd_marking, categorie, capsula,
                specificatii, stare, cantitate, provenienta_note, datasheet
            ) VALUES (
                :locatie, :part_number, :smd_marking, :categorie, :capsula,
                :specificatii, :stare, :cantitate, :provenienta_note, :datasheet
            );
        """
        if _cursor is not None:
            _cursor.execute(sql, asdict(comp))
            return _cursor.lastrowid or 0
        with self.session() as cur:
            cur.execute(sql, asdict(comp))
            return cur.lastrowid or 0

    def update(self, comp: Componenta) -> None:
        with self.session() as cur:
            cur.execute(
                """
                UPDATE componente SET
                    locatie = :locatie,
                    part_number = :part_number,
                    smd_marking = :smd_marking,
                    categorie = :categorie,
                    capsula = :capsula,
                    specificatii = :specificatii,
                    stare = :stare,
                    cantitate = :cantitate,
                    provenienta_note = :provenienta_note,
                    datasheet = :datasheet,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = :id;
                """,
                asdict(comp),
            )

    def delete(self, comp_id: int) -> None:
        with self.session() as cur:
            cur.execute("DELETE FROM componente WHERE id = ?;", (comp_id,))

    def adjust_qty(self, comp_id: int, delta: int) -> None:
        with self.session() as cur:
            cur.execute(
                "UPDATE componente SET cantitate = MAX(0, cantitate + ?), updated_at = CURRENT_TIMESTAMP WHERE id = ?;",
                (delta, comp_id),
            )

    def search(self, query: str = "", categorie: str = "Toate", stare: str = "Toate") -> list[sqlite3.Row]:
        sql = "SELECT * FROM componente WHERE 1=1"
        params: list[Any] = []

        if (q := query.strip()):
            for token in q.split():
                sql += """ AND (
                    part_number LIKE ? OR
                    smd_marking LIKE ? OR
                    specificatii LIKE ? OR
                    capsula LIKE ? OR
                    locatie LIKE ? OR
                    provenienta_note LIKE ?
                )"""
                like_tok = f"%{token}%"
                params.extend([like_tok] * 6)

        if categorie != "Toate":
            sql += " AND categorie = ?"
            params.append(categorie)

        if stare != "Toate":
            sql += " AND stare = ?"
            params.append(stare)

        sql += " ORDER BY locatie ASC, part_number ASC;"
        with self.session() as cur:
            return cur.execute(sql, params).fetchall()

    def suggest_next_location(self, prefix: str = "P1-") -> str:
        with self.session() as cur:
            rows = cur.execute(
                "SELECT locatie FROM componente WHERE locatie LIKE ?;",
                (f"{prefix}%",),
            ).fetchall()
        max_num = 0
        for r in rows:
            loc = r["locatie"]
            suffix = loc[len(prefix):]
            if suffix.isdigit():
                max_num = max(max_num, int(suffix))
        return f"{prefix}{max_num + 1:03d}"


class DatasheetResolver:
    """Rezolvă și deschide inteligent PDF-urile (Local Explicit -> Auto-Discovery în ./datasheets/ -> Web URL/Search)."""

    @staticmethod
    def resolve_local_file(part_number: str, stored_path: str) -> Path | None:
        # 1. Verificăm calea explicit salvată în baza de date (relativă la BASE_DIR sau absolută)
        if (sp := stored_path.strip()) and not sp.lower().startswith(("http://", "https://")):
            candidate = Path(sp)
            if not candidate.is_absolute():
                candidate = BASE_DIR / candidate
            if candidate.is_file():
                return candidate

        # 2. Auto-Discovery: căutăm automat <part_number>.pdf în folderul ./datasheets/ (case-insensitive)
        clean_pn = "".join(c for c in part_number.strip() if c.isalnum() or c in ("-", "_")).lower()
        if clean_pn and DATASHEETS_DIR.is_dir():
            for pdf_file in DATASHEETS_DIR.glob("*.pdf"):
                if pdf_file.stem.lower() == clean_pn or clean_pn in pdf_file.stem.lower():
                    return pdf_file
        return None

    @classmethod
    def get_status_badge(cls, part_number: str, stored_path: str) -> str:
        sp = stored_path.strip()
        if sp.lower().startswith(("http://", "https://")):
            return "🌐 Link Web"
        if cls.resolve_local_file(part_number, sp) is not None:
            return "📄 PDF Local"
        if sp:
            return "⚠️ PDF Lipsă"
        return "🔍 Căutare Web"

    @staticmethod
    def open_native(file_path: Path) -> None:
        """Lansează fișierul PDF cu aplicația nativă a sistemului de operare (Structural Pattern Matching)."""
        target = str(file_path.resolve())
        match sys.platform:
            case s if s.startswith("win"):
                os.startfile(target)  # type: ignore[attr-defined]
            case "darwin":
                subprocess.Popen(["open", target])
            case _:
                subprocess.Popen(["xdg-open", target])


class InventarApp(tk.Tk):
    def __init__(self, db: DatabaseManager) -> None:
        super().__init__()
        self.db = db
        self.selected_id: int | None = None

        self.title("Inventar Atelier Electronică & ECU - v2.0 (cu Gestiune Datasheet PDF)")
        self.geometry("1360x780")
        self.minsize(1100, 640)

        self._configure_styles()
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_table()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Datasheet.TButton", font=("Segoe UI", 10, "bold"), foreground="#004d99")

    def _bind_shortcuts(self) -> None:
        self.bind("<F3>", lambda _: self._open_datasheet())
        self.bind("<Control-f>", lambda _: self.search_entry.focus_set())

    def _build_ui(self) -> None:
        # --- Bara de Căutare și Filtrare (Sus) ---
        top_frame = ttk.LabelFrame(self, text=" Căutare Instantanee (Ctrl+F) | Dublu-Clic sau F3 pe o piesă pentru a deschide Datasheet-ul PDF ", padding=8)
        top_frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        ttk.Label(top_frame, text="Caută:").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh_table())
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=34, font=("Segoe UI", 11))
        self.search_entry.pack(side=tk.LEFT, padx=(0, 12))
        self.search_entry.focus_set()

        ttk.Label(top_frame, text="Categorie:").pack(side=tk.LEFT, padx=(0, 4))
        self.cat_filter_var = tk.StringVar(value="Toate")
        cat_cb = ttk.Combobox(
            top_frame,
            textvariable=self.cat_filter_var,
            values=["Toate"] + CATEGORIES,
            state="readonly",
            width=34,
        )
        cat_cb.pack(side=tk.LEFT, padx=(0, 12))
        cat_cb.bind("<<ComboboxSelected>>", lambda _: self.refresh_table())

        ttk.Label(top_frame, text="Stare:").pack(side=tk.LEFT, padx=(0, 4))
        self.state_filter_var = tk.StringVar(value="Toate")
        state_cb = ttk.Combobox(
            top_frame,
            textvariable=self.state_filter_var,
            values=["Toate"] + STATES,
            state="readonly",
            width=18,
        )
        state_cb.pack(side=tk.LEFT, padx=(0, 12))
        state_cb.bind("<<ComboboxSelected>>", lambda _: self.refresh_table())

        ttk.Button(top_frame, text="Resetează Filtre", command=self._reset_filters).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_frame, text="Export CSV", command=self._export_csv).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top_frame, text="📂 Folder Datasheets", command=self._open_datasheets_folder).pack(side=tk.RIGHT, padx=4)

        # --- Tabelul Central (Treeview) ---
        table_frame = ttk.Frame(self, padding=(10, 4))
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = (
            "id", "locatie", "part_number", "smd_marking", "capsula",
            "specificatii", "categorie", "stare", "cantitate", "ds_status", "provenienta_note", "datasheet_raw"
        )
        self.tree = ttk.Treeview(
            table_frame,
            columns=cols,
            displaycolumns=("id", "locatie", "part_number", "smd_marking", "capsula", "specificatii", "categorie", "stare", "cantitate", "ds_status", "provenienta_note"),
            show="headings",
            selectmode="browse",
        )

        headers = {
            "id": ("ID", 40, tk.CENTER),
            "locatie": ("Locație", 70, tk.CENTER),
            "part_number": ("Cod Piesă (Part No.)", 140, tk.W),
            "smd_marking": ("Marcaj SMD", 90, tk.CENTER),
            "capsula": ("Capsulă", 80, tk.CENTER),
            "specificatii": ("Parametri / Specificații Tehnice", 240, tk.W),
            "categorie": ("Categorie", 200, tk.W),
            "stare": ("Stare", 125, tk.CENTER),
            "cantitate": ("Stoc", 50, tk.CENTER),
            "ds_status": ("Datasheet (F3)", 110, tk.CENTER),
            "provenienta_note": ("Note / Placă Donatoare / Echivalențe", 200, tk.W),
        }
        for col, (text, width, anchor) in headers.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=anchor)

        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("zero_stock", foreground="#999999")
        self.tree.tag_configure("used_part", foreground="#004d99")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _: self._open_datasheet())

        # --- Formular Adăugare / Editare (Jos) ---
        form_frame = ttk.LabelFrame(self, text=" Adăugare / Editare Componentă & Atașare Datasheet PDF ", padding=10)
        form_frame.pack(fill=tk.X, padx=10, pady=(4, 10))

        self.vars: dict[str, tk.StringVar] = {
            k: tk.StringVar()
            for k in ("locatie", "part_number", "smd_marking", "capsula", "specificatii", "categorie", "stare", "cantitate", "provenienta_note", "datasheet")
        }
        self.vars["categorie"].set(CATEGORIES[0])
        self.vars["stare"].set(STATES[0])
        self.vars["cantitate"].set("1")

        # Rândul 1 în formular
        ttk.Label(form_frame, text="Locație (ex: P1-008):").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
        loc_sub = ttk.Frame(form_frame)
        loc_sub.grid(row=0, column=1, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(loc_sub, textvariable=self.vars["locatie"], width=11).pack(side=tk.LEFT)
        ttk.Button(loc_sub, text="Auto P1", width=8, command=lambda: self.vars["locatie"].set(self.db.suggest_next_location("P1-"))).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(form_frame, text="Cod Piesă (Part No.):").grid(row=0, column=2, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["part_number"], width=22).grid(row=0, column=3, sticky=tk.W, padx=4, pady=3)

        ttk.Label(form_frame, text="Marcaj SMD (Top Mark):").grid(row=0, column=4, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["smd_marking"], width=16).grid(row=0, column=5, sticky=tk.W, padx=4, pady=3)

        ttk.Label(form_frame, text="Capsulă (SOT-23, SOIC-8):").grid(row=0, column=6, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["capsula"], width=16).grid(row=0, column=7, sticky=tk.W, padx=4, pady=3)

        # Rândul 2 în formular
        ttk.Label(form_frame, text="Categorie:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Combobox(form_frame, textvariable=self.vars["categorie"], values=CATEGORIES, state="readonly", width=34).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=4, pady=3)

        ttk.Label(form_frame, text="Specificații (V, A, Rds, Funcție):").grid(row=1, column=3, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["specificatii"], width=36).grid(row=1, column=4, columnspan=2, sticky=tk.W, padx=4, pady=3)

        ttk.Label(form_frame, text="Stare:").grid(row=1, column=6, sticky=tk.W, padx=4, pady=3)
        ttk.Combobox(form_frame, textvariable=self.vars["stare"], values=STATES, state="readonly", width=18).grid(row=1, column=7, sticky=tk.W, padx=4, pady=3)

        # Rândul 3 în formular (Cantitate, Note, Datasheet)
        ttk.Label(form_frame, text="Cantitate:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["cantitate"], width=10).grid(row=2, column=1, sticky=tk.W, padx=4, pady=3)

        ttk.Label(form_frame, text="Note / Proveniență:").grid(row=2, column=2, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(form_frame, textvariable=self.vars["provenienta_note"], width=34).grid(row=2, column=3, columnspan=2, sticky=tk.EW, padx=4, pady=3)

        ttk.Label(form_frame, text="Datasheet (PDF / URL):").grid(row=2, column=5, sticky=tk.W, padx=4, pady=3)
        ds_sub = ttk.Frame(form_frame)
        ds_sub.grid(row=2, column=6, columnspan=2, sticky=tk.EW, padx=4, pady=3)
        ttk.Entry(ds_sub, textvariable=self.vars["datasheet"], width=24).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(ds_sub, text="📎 Atașează PDF...", command=self._browse_datasheet).pack(side=tk.LEFT, padx=(4, 0))

        # Rândul 4: Butoane de Acțiune
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=3, column=0, columnspan=8, sticky=tk.EW, pady=(8, 0))

        ttk.Button(btn_frame, text="Adaugă ca Piesă Nouă", style="Accent.TButton", command=self._add_item).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Salvează Modificările", command=self._update_item).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="Golește Formularul", command=self._clear_form).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="📄 Deschide Datasheet (F3)", style="Datasheet.TButton", command=self._open_datasheet).pack(side=tk.LEFT, padx=14)
        ttk.Button(btn_frame, text="Șterge Piesa", command=self._delete_item).pack(side=tk.LEFT, padx=8)

        self.status_lbl = ttk.Label(btn_frame, text="", font=("Segoe UI", 9, "italic"))
        self.status_lbl.pack(side=tk.LEFT, padx=14)

        ttk.Button(btn_frame, text="+1 Stoc", width=9, command=lambda: self._quick_qty(1)).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="-1 Stoc (Folosit)", width=15, command=lambda: self._quick_qty(-1)).pack(side=tk.RIGHT, padx=4)

    def _browse_datasheet(self) -> None:
        path_str = filedialog.askopenfilename(
            title="Selectează fișierul Datasheet PDF",
            filetypes=[("Documente PDF", "*.pdf"), ("Toate fișierele", "*.*")],
        )
        if not path_str:
            return
        src_path = Path(path_str).resolve()
        pn = self.vars["part_number"].get().strip()

        # Oferim copierea automată în ./datasheets/ pentru portabilitate completă între PC-uri
        if DATASHEETS_DIR not in src_path.parents:
            default_name = f"{pn}.pdf" if pn else src_path.name
            clean_name = "".join(c for c in default_name if c.isalnum() or c in (".", "-", "_"))
            dest_path = DATASHEETS_DIR / clean_name
            if messagebox.askyesno(
                "Copiere în folderul local './datasheets/'",
                f"Vrei să copiez PDF-ul în folderul local al aplicației:\n{dest_path}\npentru ca inventarul să fie 100% portabil?",
            ):
                shutil.copy2(src_path, dest_path)
                src_path = dest_path

        try:
            rel_path = src_path.relative_to(BASE_DIR).as_posix()
            self.vars["datasheet"].set(rel_path)
        except ValueError:
            self.vars["datasheet"].set(str(src_path))

    def _open_datasheet(self) -> None:
        pn = self.vars["part_number"].get().strip()
        ds_raw = self.vars["datasheet"].get().strip()

        if not pn and not ds_raw:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă din tabel (sau completează Codul Piesei) pentru a deschide Datasheet-ul!")
            return

        # 1. Dacă este URL web direct (http:// sau https://)
        if ds_raw.lower().startswith(("http://", "https://")):
            webbrowser.open(ds_raw)
            return

        # 2. Căutăm fișierul PDF local (explicit sau prin auto-discovery în ./datasheets/<part_number>.pdf)
        if (local_pdf := DatasheetResolver.resolve_local_file(pn, ds_raw)) is not None:
            try:
                DatasheetResolver.open_native(local_pdf)
            except Exception as exc:
                messagebox.showerror("Eroare la deschiderea PDF-ului", f"Nu am putut lansa cititorul PDF:\n{exc}")
            return

        # 3. Fallback inteligent: oferim fie atașarea unui PDF local, fie căutarea directă pe Web după "<part_number> datasheet pdf"
        if messagebox.askyesno(
            "Datasheet PDF local negăsit",
            f"Nu am găsit un fișier PDF local pentru '{pn}' în folderul './datasheets/'.\n\n"
            f"Vrei să caut automat pe Web datasheet-ul PDF pentru '{pn}'?\n"
            f"(Dacă îl descarci în folderul 'datasheets/{pn}.pdf', data viitoare se va deschide direct local!)",
        ):
            query = urllib.parse.quote_plus(f"{pn} datasheet filetype:pdf")
            webbrowser.open(f"https://www.google.com/search?q={query}")

    def _open_datasheets_folder(self) -> None:
        DATASHEETS_DIR.mkdir(parents=True, exist_ok=True)
        DatasheetResolver.open_native(DATASHEETS_DIR)

    def _collect_form(self) -> Componenta | None:
        loc = self.vars["locatie"].get().strip().upper()
        pn = self.vars["part_number"].get().strip()
        if not loc or not pn:
            messagebox.showwarning("Câmpuri obligatorii", "Completează cel puțin 'Locație' și 'Cod Piesă (Part No.)'!")
            return None
        try:
            qty = max(0, int(self.vars["cantitate"].get().strip()))
        except ValueError:
            messagebox.showwarning("Cantitate invalidă", "Cantitatea trebuie să fie un număr întreg!")
            return None

        return Componenta(
            id=self.selected_id,
            locatie=loc,
            part_number=pn,
            smd_marking=self.vars["smd_marking"].get().strip(),
            categorie=self.vars["categorie"].get(),
            capsula=self.vars["capsula"].get().strip().upper(),
            specificatii=self.vars["specificatii"].get().strip(),
            stare=self.vars["stare"].get(),
            cantitate=qty,
            provenienta_note=self.vars["provenienta_note"].get().strip(),
            datasheet=self.vars["datasheet"].get().strip(),
        )

    def refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = self.db.search(
            query=self.search_var.get(),
            categorie=self.cat_filter_var.get(),
            stare=self.state_filter_var.get(),
        )
        total_qty = 0
        for r in rows:
            qty = r["cantitate"]
            total_qty += qty
            tags = ()
            if qty == 0:
                tags = ("zero_stock",)
            elif r["stare"] != "Nou":
                tags = ("used_part",)

            ds_raw = r["datasheet"] or ""
            ds_badge = DatasheetResolver.get_status_badge(r["part_number"], ds_raw)

            self.tree.insert(
                "",
                tk.END,
                values=(
                    r["id"],
                    r["locatie"],
                    r["part_number"],
                    r["smd_marking"],
                    r["capsula"],
                    r["specificatii"],
                    r["categorie"],
                    r["stare"],
                    qty,
                    ds_badge,
                    r["provenienta_note"],
                    ds_raw,
                ),
                tags=tags,
            )
        self.status_lbl.config(text=f"Afișate: {len(rows)} repere | Total bucăți: {total_qty}")

    def _on_select(self, _event: Any) -> None:
        if not (sel := self.tree.selection()):
            return
        vals = self.tree.item(sel[0], "values")
        self.selected_id = int(vals[0])
        # Mapăm valorile din rândul selectat înapoi în formular (inclusiv datasheet_raw de pe indexul 11)
        self.vars["locatie"].set(str(vals[1]))
        self.vars["part_number"].set(str(vals[2]))
        self.vars["smd_marking"].set(str(vals[3]))
        self.vars["capsula"].set(str(vals[4]))
        self.vars["specificatii"].set(str(vals[5]))
        self.vars["categorie"].set(str(vals[6]))
        self.vars["stare"].set(str(vals[7]))
        self.vars["cantitate"].set(str(vals[8]))
        self.vars["provenienta_note"].set(str(vals[10]))
        self.vars["datasheet"].set(str(vals[11]))

    def _add_item(self) -> None:
        if (comp := self._collect_form()) is None:
            return
        comp.id = None
        self.db.add(comp)
        self.refresh_table()
        self._clear_form()

    def _update_item(self) -> None:
        if self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează mai întâi un rând din tabel pentru a-l modifica.")
            return
        if (comp := self._collect_form()) is None:
            return
        self.db.update(comp)
        self.refresh_table()

    def _delete_item(self) -> None:
        if self.selected_id is None:
            return
        if messagebox.askyesno("Confirmare ștergere", f"Sigur vrei să ștergi componenta ID #{self.selected_id}?"):
            self.db.delete(self.selected_id)
            self._clear_form()
            self.refresh_table()

    def _quick_qty(self, delta: int) -> None:
        if self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă din tabel pentru a ajusta rapid stocul.")
            return
        self.db.adjust_qty(self.selected_id, delta)
        current_id = self.selected_id
        self.refresh_table()
        for item in self.tree.get_children():
            if int(self.tree.item(item, "values")[0]) == current_id:
                self.tree.selection_set(item)
                self._on_select(None)
                break

    def _clear_form(self) -> None:
        self.selected_id = None
        self.tree.selection_remove(self.tree.selection())
        for k in ("locatie", "part_number", "smd_marking", "capsula", "specificatii", "provenienta_note", "datasheet"):
            self.vars[k].set("")
        self.vars["cantitate"].set("1")

    def _reset_filters(self) -> None:
        self.search_var.set("")
        self.cat_filter_var.set("Toate")
        self.state_filter_var.set("Toate")
        self.refresh_table()

    def _export_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fișier CSV", "*.csv")],
            initialfile="inventar_atelier_export.csv",
        )
        if not path:
            return
        rows = self.db.search()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["ID", "Locatie", "Part_Number", "SMD_Marking", "Capsula", "Specificatii", "Categorie", "Stare", "Cantitate", "Datasheet", "Note"])
            for r in rows:
                writer.writerow([
                    r["id"], r["locatie"], r["part_number"], r["smd_marking"],
                    r["capsula"], r["specificatii"], r["categorie"], r["stare"],
                    r["cantitate"], r["datasheet"], r["provenienta_note"]
                ])
        messagebox.showinfo("Export finalizat", f"Inventarul a fost exportat în:\n{path}")


if __name__ == "__main__":
    db_manager = DatabaseManager()
    app = InventarApp(db_manager)
    app.mainloop()
