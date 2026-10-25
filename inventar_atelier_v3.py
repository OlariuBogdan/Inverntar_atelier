#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Inventar Atelier Electronică & ECU - v3.0
Gestiune locală de componente, cu datasheet-uri PDF, stoc minim, istoric de mișcări,
import/export CSV, backup automat și etichete pentru sertare.

100% biblioteca standard Python (3.10+), fără pachete externe. Windows / Linux / macOS.
Compatibil cu baza de date din v1/v2 (inventar_atelier.db): la prima pornire în v3 se face
automat o copie de siguranță în ./backups/ și apoi se adaugă coloanele noi.

Rulare:   python inventar_atelier_v3.py            (Windows: py inventar_atelier_v3.py)
Opțional: python inventar_atelier_v3.py --db D:\\atelier\\inventar.db
"""

from __future__ import annotations

import argparse
import csv
import html
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import urllib.parse
import webbrowser
from contextlib import contextmanager
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterator

try:
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # permite folosirea/testarea stratului de date fără Tk
    tk = None  # type: ignore[assignment]

APP_VERSION = "3.0"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_DB = BASE_DIR / "inventar_atelier.db"
CURRENCY = "lei"          # moneda afișată la valoarea stocului
BACKUP_KEEP = 20          # câte backup-uri automate se păstrează
SCHEMA_VERSION = 3

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
    "Consumabile / Scule (Flux, Pastă, Spray, Ață)",
]

STATE_NEW = "Nou"
STATE_DEFECT = "Defect / Rebut"
STATES = [
    STATE_NEW,
    "Recuperat (Testat OK)",
    "Din Placă Donatoare",
    "Recuperat (De testat)",
    STATE_DEFECT,
]


# --------------------------------------------------------------------------- model

@dataclass(slots=True)
class Componenta:
    locatie: str
    part_number: str
    smd_marking: str = ""
    categorie: str = CATEGORIES[0]
    capsula: str = ""
    specificatii: str = ""
    stare: str = STATE_NEW
    cantitate: int = 1
    provenienta_note: str = ""
    datasheet: str = ""
    stoc_minim: int = 0
    pret: float = 0.0
    furnizor: str = ""
    echivalente: str = ""
    id: int | None = None


COMP_FIELDS = [f.name for f in fields(Componenta) if f.name != "id"]

NEW_COLUMNS = {
    "datasheet": "TEXT DEFAULT ''",
    "stoc_minim": "INTEGER NOT NULL DEFAULT 0",
    "pret": "REAL NOT NULL DEFAULT 0",
    "furnizor": "TEXT DEFAULT ''",
    "echivalente": "TEXT DEFAULT ''",
    "created_at": "TIMESTAMP",
}

SEARCH_COLS = [
    "part_number", "smd_marking", "specificatii", "capsula", "locatie",
    "provenienta_note", "echivalente", "furnizor", "categorie",
]

# (antet CSV, câmp)
EXPORT_COLS = [
    ("ID", "id"), ("Locatie", "locatie"), ("Part_Number", "part_number"),
    ("SMD_Marking", "smd_marking"), ("Capsula", "capsula"), ("Specificatii", "specificatii"),
    ("Categorie", "categorie"), ("Stare", "stare"), ("Cantitate", "cantitate"),
    ("Stoc_Minim", "stoc_minim"), ("Pret", "pret"), ("Furnizor", "furnizor"),
    ("Echivalente", "echivalente"), ("Datasheet", "datasheet"), ("Note", "provenienta_note"),
]


def natural_key(text: Any) -> list:
    """Sortare naturală: P1-9 înaintea lui P1-10. Lista alternează mereu str/int."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(text))]


def to_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(str(value).strip().replace(",", ".")))
    except (ValueError, TypeError):
        return default


def to_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(str(value).strip().replace(",", "."))
    except (ValueError, TypeError):
        return default


def csv_safe(value: Any) -> Any:
    """Previne injectarea de formule când CSV-ul e deschis în Excel/Calc."""
    if isinstance(value, str) and value[:1] in ("=", "+", "@", "\t", "\r"):
        return "'" + value
    if isinstance(value, str) and value[:1] == "-" and to_float(value, None) is None:  # type: ignore[arg-type]
        return "'" + value
    return value


def like_escape(token: str) -> str:
    return token.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


# --------------------------------------------------------------------------- baza de date

class DatabaseManager:
    """SQLite3 (WAL), migrare automată, istoric de mișcări, backup."""

    def __init__(self, db_path: Path = DEFAULT_DB) -> None:
        self.db_path = Path(db_path)
        self.base = self.db_path.parent
        self.datasheets_dir = self.base / "datasheets"
        self.backup_dir = self.base / "backups"
        self.datasheets_dir.mkdir(parents=True, exist_ok=True)

        existed = self.db_path.exists() and self.db_path.stat().st_size > 0
        if existed and self._user_version() < SCHEMA_VERSION:
            self.backup(tag="pre-v3")
        self._init_and_migrate_db()

    # -- infrastructură
    @contextmanager
    def session(self) -> Iterator[sqlite3.Cursor]:
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
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

    def _user_version(self) -> int:
        try:
            conn = sqlite3.connect(self.db_path)
            try:
                return conn.execute("PRAGMA user_version;").fetchone()[0]
            finally:
                conn.close()
        except sqlite3.Error:
            return 0

    def _init_and_migrate_db(self) -> None:
        conn = sqlite3.connect(self.db_path)
        try:
            conn.execute("PRAGMA journal_mode = WAL;")
        finally:
            conn.close()

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
            existing = {row["name"] for row in cur.execute("PRAGMA table_info(componente);")}
            for col, ddl in NEW_COLUMNS.items():
                if col not in existing:
                    cur.execute(f"ALTER TABLE componente ADD COLUMN {col} {ddl};")
            cur.execute("UPDATE componente SET created_at = COALESCE(updated_at, CURRENT_TIMESTAMP) WHERE created_at IS NULL;")

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS miscari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    comp_id INTEGER,
                    part_number TEXT NOT NULL,
                    delta INTEGER NOT NULL,
                    cantitate_dupa INTEGER NOT NULL,
                    motiv TEXT DEFAULT '',
                    referinta TEXT DEFAULT '',
                    ts TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                """
            )
            cur.execute("CREATE INDEX IF NOT EXISTS idx_locatie ON componente(locatie);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_pn ON componente(part_number);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_smd ON componente(smd_marking);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_misc_comp ON miscari(comp_id);")

            if cur.execute("SELECT COUNT(*) FROM componente;").fetchone()[0] == 0 and \
                    cur.execute("SELECT COUNT(*) FROM miscari;").fetchone()[0] == 0:
                for comp in self._seed_data():
                    self._add(cur, comp, motiv="Stoc inițial (date exemplu)")
            cur.execute(f"PRAGMA user_version = {SCHEMA_VERSION};")

    @staticmethod
    def _seed_data() -> list[Componenta]:
        c = CATEGORIES
        return [
            Componenta("P1-001", "MMBT3904", "AM / 1AM", c[0], "SOT-23", "NPN 40V 200mA 300MHz", "Nou", 40, "Echivalent uzual comandă relee/opto (Fumis RCY-1)", "datasheets/MMBT3904.pdf"),
            Componenta("P1-002", "TPS5120DBT", "TPS5120", c[1], "TSSOP-30", "Dual Buck Controller 5V / 3.3V", "Recuperat (Testat OK)", 2, "Clarion QY BASE PWB IC902", "datasheets/TPS5120DBT.pdf"),
            Componenta("P1-003", "R1514S037B", "R1514S", c[1], "HSOP-6", "LDO Regulator 3.6V/3.7V Standby", "Recuperat (Testat OK)", 2, "Clarion QY Standby BU-3.3V IC905"),
            Componenta("S1-A1", "TB2926HQ", "TB2926HQ", c[3], "HZIP-25", "4x45W BTL Audio Power Amp", "Recuperat (Testat OK)", 1, "Clarion QY Audio Final IC901"),
            Componenta("P1-004", "BA5830FM", "BA5830FM", c[3], "HSOP-28", "4-Channel CD/DVD Motor Driver 5V-8V", "Recuperat (Testat OK)", 3, "Mecanisme optice Clarion / Becker"),
            Componenta("P1-005", "TJA1050", "A1050", c[2], "SOIC-8", "High-Speed CAN Transceiver 1Mbps 5V", "Nou", 15, "Stoc reparații ECU / CAN Bus", "datasheets/TJA1050.pdf"),
            Componenta("S2-B3", "STP10NK60Z", "10NK60Z", c[0], "TO-220", "MOSFET N-Ch 600V 10A 0.65R Zener Protected", "Nou", 8, "Surse SMPS primar"),
        ]

    # -- backup
    def backup(self, tag: str = "manual") -> Path:
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        dest = self.backup_dir / f"inventar_{datetime.now():%Y%m%d_%H%M%S}_{tag}.db"
        src = sqlite3.connect(self.db_path)
        dst = sqlite3.connect(dest)
        try:
            src.backup(dst)  # corect și cu WAL, spre deosebire de copierea simplă a fișierului
        finally:
            dst.close()
            src.close()
        self._prune_backups()
        return dest

    def backup_if_due(self, hours: int = 20) -> Path | None:
        newest = max((p.stat().st_mtime for p in self.backup_dir.glob("*.db")), default=0.0) if self.backup_dir.is_dir() else 0.0
        if datetime.now() - datetime.fromtimestamp(newest) > timedelta(hours=hours):
            return self.backup(tag="auto")
        return None

    def _prune_backups(self) -> None:
        autos = sorted(self.backup_dir.glob("*_auto.db"))
        for old in autos[:-BACKUP_KEEP]:
            try:
                old.unlink()
            except OSError:
                pass

    # -- scriere
    def _log(self, cur: sqlite3.Cursor, comp_id: int | None, pn: str, delta: int, after: int, motiv: str, ref: str) -> None:
        cur.execute(
            "INSERT INTO miscari (comp_id, part_number, delta, cantitate_dupa, motiv, referinta) VALUES (?,?,?,?,?,?);",
            (comp_id, pn, delta, after, motiv, ref),
        )

    def _add(self, cur: sqlite3.Cursor, comp: Componenta, motiv: str = "Piesă nouă", ref: str = "") -> int:
        data = asdict(comp)
        cols = ", ".join(COMP_FIELDS)
        marks = ", ".join(f":{c}" for c in COMP_FIELDS)
        cur.execute(f"INSERT INTO componente ({cols}, created_at) VALUES ({marks}, CURRENT_TIMESTAMP);", data)
        new_id = cur.lastrowid or 0
        if comp.cantitate:
            self._log(cur, new_id, comp.part_number, comp.cantitate, comp.cantitate, motiv, ref)
        return new_id

    def add(self, comp: Componenta, ref: str = "") -> int:
        with self.session() as cur:
            return self._add(cur, comp, ref=ref)

    def update(self, comp: Componenta, ref: str = "") -> None:
        if comp.id is None:
            raise ValueError("Piesa nu are ID")
        with self.session() as cur:
            old = cur.execute("SELECT cantitate FROM componente WHERE id = ?;", (comp.id,)).fetchone()
            sets = ", ".join(f"{c} = :{c}" for c in COMP_FIELDS)
            cur.execute(f"UPDATE componente SET {sets}, updated_at = CURRENT_TIMESTAMP WHERE id = :id;", asdict(comp))
            if old is not None and old["cantitate"] != comp.cantitate:
                self._log(cur, comp.id, comp.part_number, comp.cantitate - old["cantitate"], comp.cantitate, "Corecție manuală", ref)

    def delete(self, comp_id: int) -> None:
        with self.session() as cur:
            row = cur.execute("SELECT part_number, cantitate FROM componente WHERE id = ?;", (comp_id,)).fetchone()
            if row is None:
                return
            if row["cantitate"]:
                self._log(cur, comp_id, row["part_number"], -row["cantitate"], 0, "Piesă ștearsă", "")
            cur.execute("DELETE FROM componente WHERE id = ?;", (comp_id,))

    def adjust_qty(self, comp_id: int, delta: int, ref: str = "", motiv: str | None = None) -> int | None:
        """Ajustează stocul (nu coboară sub 0). Returnează noua cantitate sau None dacă nu există."""
        with self.session() as cur:
            row = cur.execute("SELECT part_number, cantitate FROM componente WHERE id = ?;", (comp_id,)).fetchone()
            if row is None:
                return None
            new = max(0, row["cantitate"] + delta)
            real = new - row["cantitate"]
            if real != 0:
                cur.execute("UPDATE componente SET cantitate = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?;", (new, comp_id))
                self._log(cur, comp_id, row["part_number"], real, new, motiv or ("Intrare stoc" if real > 0 else "Consum"), ref)
            return new

    def import_rows(self, comps: list[Componenta]) -> tuple[int, int]:
        """Adaugă piese; dacă există aceeași piesă (cod + locație + stare), adună cantitatea. Atomic."""
        added = merged = 0
        with self.session() as cur:
            for comp in comps:
                dup = self._find_dup(cur, comp.part_number, comp.locatie, comp.stare, None)
                if dup is not None:
                    new = dup["cantitate"] + comp.cantitate
                    cur.execute("UPDATE componente SET cantitate = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?;", (new, dup["id"]))
                    if comp.cantitate:
                        self._log(cur, dup["id"], dup["part_number"], comp.cantitate, new, "Import CSV (adunat)", "")
                    merged += 1
                else:
                    self._add(cur, comp, motiv="Import CSV")
                    added += 1
        return added, merged

    # -- citire
    @staticmethod
    def _find_dup(cur: sqlite3.Cursor, pn: str, loc: str, stare: str, exclude_id: int | None) -> sqlite3.Row | None:
        return cur.execute(
            """SELECT * FROM componente
               WHERE lower(part_number) = lower(?) AND upper(locatie) = upper(?) AND stare = ?
                 AND (? IS NULL OR id != ?) LIMIT 1;""",
            (pn.strip(), loc.strip(), stare, exclude_id, exclude_id),
        ).fetchone()

    def find_duplicate(self, pn: str, loc: str, stare: str, exclude_id: int | None = None) -> sqlite3.Row | None:
        with self.session() as cur:
            return self._find_dup(cur, pn, loc, stare, exclude_id)

    def get(self, comp_id: int) -> sqlite3.Row | None:
        with self.session() as cur:
            return cur.execute("SELECT * FROM componente WHERE id = ?;", (comp_id,)).fetchone()

    def search(self, query: str = "", categorie: str = "Toate", stare: str = "Toate",
               only_low: bool = False) -> list[sqlite3.Row]:
        sql = "SELECT * FROM componente WHERE 1=1"
        params: list[Any] = []
        for token in query.split():
            sql += " AND (" + " OR ".join(f"{c} LIKE ? ESCAPE '\\'" for c in SEARCH_COLS) + ")"
            params.extend([f"%{like_escape(token)}%"] * len(SEARCH_COLS))
        if categorie != "Toate":
            sql += " AND categorie = ?"
            params.append(categorie)
        if stare != "Toate":
            sql += " AND stare = ?"
            params.append(stare)
        if only_low:
            sql += " AND stoc_minim > 0 AND cantitate <= stoc_minim AND stare != ?"
            params.append(STATE_DEFECT)
        with self.session() as cur:
            return cur.execute(sql + ";", params).fetchall()

    def all_categories(self) -> list[str]:
        with self.session() as cur:
            extra = [r[0] for r in cur.execute("SELECT DISTINCT categorie FROM componente;") if r[0] not in CATEGORIES]
        return CATEGORIES + sorted(extra)

    def all_states(self) -> list[str]:
        with self.session() as cur:
            extra = [r[0] for r in cur.execute("SELECT DISTINCT stare FROM componente;") if r[0] not in STATES]
        return STATES + sorted(extra)

    def history(self, comp_id: int | None = None, limit: int = 1000) -> list[sqlite3.Row]:
        with self.session() as cur:
            if comp_id is None:
                return cur.execute("SELECT * FROM miscari ORDER BY id DESC LIMIT ?;", (limit,)).fetchall()
            return cur.execute("SELECT * FROM miscari WHERE comp_id = ? ORDER BY id DESC LIMIT ?;", (comp_id, limit)).fetchall()

    def known_prefixes(self) -> list[str]:
        """Prefixele de locație folosite (ex. 'P1-', 'S2-B'), cele mai frecvente primele."""
        counts: dict[str, int] = {}
        with self.session() as cur:
            for r in cur.execute("SELECT locatie FROM componente;"):
                m = re.match(r"^(.*?)(\d+)$", r["locatie"])
                if m:
                    counts[m.group(1)] = counts.get(m.group(1), 0) + 1
        return [p for p, _ in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))]

    def suggest_next_location(self, prefix: str = "P1-") -> str:
        max_num, width = 0, 3
        with self.session() as cur:
            rows = cur.execute("SELECT locatie FROM componente WHERE locatie LIKE ? ESCAPE '\\';",
                               (like_escape(prefix) + "%",)).fetchall()
        for r in rows:
            m = re.match(r"^(.*?)(\d+)$", r["locatie"])
            if m and m.group(1) == prefix:
                max_num = max(max_num, int(m.group(2)))
                width = max(width, len(m.group(2)))
        return f"{prefix}{max_num + 1:0{width}d}"


# --------------------------------------------------------------------------- datasheet

class DatasheetResolver:
    """PDF explicit (BD) -> auto-descoperire în ./datasheets/ -> link web -> căutare web."""

    def __init__(self, base_dir: Path) -> None:
        self.base = base_dir
        self.dir = base_dir / "datasheets"

    def index(self) -> list[tuple[str, Path]]:
        """Scanează folderul o singură dată (nu per rând)."""
        if not self.dir.is_dir():
            return []
        return sorted(((p.stem.lower(), p) for p in self.dir.glob("*.pdf")), key=lambda t: t[0])

    @staticmethod
    def _clean(pn: str) -> str:
        return "".join(c for c in pn.strip() if c.isalnum() or c in "-_").lower()

    def resolve(self, pn: str, stored: str, index: list[tuple[str, Path]] | None = None) -> Path | None:
        sp = stored.strip()
        if sp and not sp.lower().startswith(("http://", "https://")):
            cand = Path(sp)
            if not cand.is_absolute():
                cand = self.base / cand
            if cand.is_file():
                return cand
        clean = self._clean(pn)
        if not clean:
            return None
        idx = index if index is not None else self.index()
        for stem, path in idx:                       # 1. potrivire exactă
            if stem == clean:
                return path
        if len(clean) >= 4:                          # 2. fișierul începe cu codul (ex. tja1050_datasheet)
            for stem, path in idx:
                if stem.startswith(clean):
                    return path
        if len(clean) >= 6:                          # 3. codul apare în nume (ultimul resort)
            for stem, path in idx:
                if clean in stem:
                    return path
        return None

    def badge(self, pn: str, stored: str, index: list[tuple[str, Path]] | None = None) -> str:
        sp = stored.strip()
        if sp.lower().startswith(("http://", "https://")):
            return "🌐 Link Web"
        if self.resolve(pn, sp, index) is not None:
            return "📄 PDF Local"
        return "⚠️ PDF Lipsă" if sp else "🔍 Căutare Web"

    @staticmethod
    def open_native(path: Path) -> None:
        """Deschide un PDF sau un folder. Alte tipuri de fișiere sunt refuzate (siguranță)."""
        target = path.resolve()
        if target.is_file() and target.suffix.lower() != ".pdf":
            raise ValueError(f"Se deschid doar fișiere PDF, nu '{target.suffix}'.")
        if sys.platform.startswith("win"):
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(target)])
        else:
            subprocess.Popen(["xdg-open", str(target)])


# --------------------------------------------------------------------------- export / import / etichete

def write_csv(path: str | Path, rows: list[sqlite3.Row], extra_need: bool = False) -> None:
    """CSV cu ';' (se deschide corect în Excel cu setări românești), UTF-8 cu BOM."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        header = [h for h, _ in EXPORT_COLS] + (["Lipsa_pana_la_minim"] if extra_need else [])
        w.writerow(header)
        for r in rows:
            line = []
            for _, key in EXPORT_COLS:
                v = r[key]
                if key == "pret":
                    v = f"{float(v or 0):.2f}".replace(".", ",")
                line.append(csv_safe(v if v is not None else ""))
            if extra_need:
                line.append(max(0, r["stoc_minim"] - r["cantitate"]))
            w.writerow(line)


def read_csv(path: str | Path) -> tuple[list[Componenta], int]:
    """Citește un CSV (v2 sau v3, separator detectat automat). Returnează (piese, rânduri invalide)."""
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        text = f.read()
    first = text.splitlines()[0] if text else ""
    delim = max((",", ";", "\t"), key=first.count)
    reader = csv.DictReader(text.splitlines(), delimiter=delim)
    alias = {h.lower(): key for h, key in EXPORT_COLS}
    alias.update({"locație": "locatie", "cod piesa": "part_number", "part no": "part_number",
                  "cantitate": "cantitate", "note": "provenienta_note"})
    comps: list[Componenta] = []
    invalid = 0
    for raw in reader:
        d = {alias.get((k or "").strip().lower(), ""): (v or "").strip() for k, v in raw.items()}
        loc, pn = d.get("locatie", "").upper(), d.get("part_number", "")
        if not loc or not pn:
            invalid += 1
            continue
        comps.append(Componenta(
            locatie=loc, part_number=pn,
            smd_marking=d.get("smd_marking", ""),
            categorie=d.get("categorie") or CATEGORIES[0],
            capsula=d.get("capsula", "").upper(),
            specificatii=d.get("specificatii", ""),
            stare=d.get("stare") or STATE_NEW,
            cantitate=max(0, to_int(d.get("cantitate", "1"), 1)),
            provenienta_note=d.get("provenienta_note", ""),
            datasheet=d.get("datasheet", ""),
            stoc_minim=max(0, to_int(d.get("stoc_minim", "0"))),
            pret=max(0.0, to_float(d.get("pret", "0"))),
            furnizor=d.get("furnizor", ""),
            echivalente=d.get("echivalente", ""),
        ))
    return comps, invalid


def build_labels_html(rows: list[sqlite3.Row]) -> str:
    esc = html.escape
    cards = []
    for r in rows:
        cards.append(
            '<div class="l"><div class="loc">{loc}</div><div class="pn">{pn}</div>'
            '<div class="sm">{smd}{cap}</div><div class="sp">{spec}</div></div>'.format(
                loc=esc(r["locatie"]), pn=esc(r["part_number"]),
                smd=esc(("SMD: " + r["smd_marking"]) if r["smd_marking"] else ""),
                cap=esc((" · " + r["capsula"]) if r["capsula"] else ""),
                spec=esc(r["specificatii"] or "")))
    return (
        "<!doctype html><html lang='ro'><head><meta charset='utf-8'><title>Etichete inventar</title><style>"
        "body{margin:8mm;font-family:Arial,Helvetica,sans-serif}"
        ".l{display:inline-block;vertical-align:top;width:60mm;height:27mm;margin:1mm;padding:2mm 2.5mm;"
        "border:0.3mm solid #000;box-sizing:border-box;overflow:hidden;page-break-inside:avoid}"
        ".loc{font-size:15pt;font-weight:bold}.pn{font-size:11pt;font-weight:bold}"
        ".sm{font-size:8pt;color:#333}.sp{font-size:7.5pt;color:#333;line-height:1.15}"
        "@media print{body{margin:0}}</style></head><body>" + "".join(cards) + "</body></html>"
    )


# --------------------------------------------------------------------------- interfața

class InventarApp(tk.Tk if tk else object):  # type: ignore[misc]
    COLS = ("id", "locatie", "part_number", "smd_marking", "capsula", "specificatii", "categorie",
            "stare", "cantitate", "stoc_minim", "ds_status", "echivalente", "provenienta_note")
    HEADERS = {
        "id": ("ID", 45, "center"), "locatie": ("Locație", 75, "center"),
        "part_number": ("Cod Piesă (Part No.)", 140, "w"), "smd_marking": ("Marcaj SMD", 90, "center"),
        "capsula": ("Capsulă", 80, "center"), "specificatii": ("Parametri / Specificații", 230, "w"),
        "categorie": ("Categorie", 190, "w"), "stare": ("Stare", 155, "center"),
        "cantitate": ("Stoc", 50, "center"), "stoc_minim": ("Min", 45, "center"),
        "ds_status": ("Datasheet (F3)", 130, "center"), "echivalente": ("Echivalențe", 150, "w"),
        "provenienta_note": ("Note / Proveniență", 190, "w"),
    }
    FORM_KEYS = ("locatie", "part_number", "smd_marking", "capsula", "specificatii", "categorie", "stare",
                 "cantitate", "stoc_minim", "pret", "furnizor", "echivalente", "provenienta_note", "datasheet")

    def __init__(self, db: DatabaseManager) -> None:
        super().__init__()
        self.db = db
        self.resolver = DatasheetResolver(db.base)
        self.selected_id: int | None = None
        self.sort_col, self.sort_rev = "locatie", False
        self._after_id: str | None = None
        self._shown: list[sqlite3.Row] = []

        self.title(f"Inventar Atelier Electronică & ECU - v{APP_VERSION}")
        self.geometry("1440x860")
        self.minsize(1150, 680)

        self._configure_styles()
        self._build_menu()
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_table()

    # -- stil
    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        fam = "Segoe UI" if "Segoe UI" in set(tkfont.families()) else tkfont.nametofont("TkDefaultFont").actual("family")
        self.ui_font = fam
        rowh = int(tkfont.Font(family=fam, size=10).metrics("linespace") * 1.55)
        style.configure("Treeview", rowheight=rowh, font=(fam, 10))
        style.configure("Treeview.Heading", font=(fam, 10, "bold"))
        style.configure("Accent.TButton", font=(fam, 10, "bold"))
        style.configure("Datasheet.TButton", font=(fam, 10, "bold"), foreground="#004d99")

        def fixed_map(option: str) -> list:  # remediază bug-ul Tk 8.6.9 cu culorile de rând (tag-uri)
            return [e for e in style.map("Treeview", query_opt=option) if e[:2] != ("!disabled", "!selected")]
        style.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))

    # -- meniu
    def _build_menu(self) -> None:
        bar = tk.Menu(self)
        f = tk.Menu(bar, tearoff=0)
        f.add_command(label="Export CSV (tot inventarul)", command=lambda: self._export_csv(all_rows=True))
        f.add_command(label="Export CSV (rândurile afișate)", command=lambda: self._export_csv(all_rows=False))
        f.add_command(label="Export listă de comandat (sub stoc minim)", command=self._export_order_list)
        f.add_separator()
        f.add_command(label="Import CSV...", command=self._import_csv)
        f.add_separator()
        f.add_command(label="Backup acum", command=self._backup_now)
        f.add_command(label="Deschide folderul de backup", command=lambda: self._open_folder(self.db.backup_dir))
        f.add_separator()
        f.add_command(label="Ieșire", command=self.destroy)
        bar.add_cascade(label="Fișier", menu=f)

        u = tk.Menu(bar, tearoff=0)
        u.add_command(label="Etichete pentru rândurile afișate (HTML)", command=lambda: self._labels(selected_only=False))
        u.add_command(label="Etichetă pentru piesa selectată", command=lambda: self._labels(selected_only=True))
        u.add_separator()
        u.add_command(label="Istoric piesa selectată", command=lambda: self._show_history(all_parts=False))
        u.add_command(label="Istoric general mișcări", command=lambda: self._show_history(all_parts=True))
        u.add_separator()
        u.add_command(label="Folder datasheets", command=lambda: self._open_folder(self.db.datasheets_dir))
        bar.add_cascade(label="Unelte", menu=u)

        h = tk.Menu(bar, tearoff=0)
        h.add_command(label="Scurtături de tastatură", command=self._show_shortcuts)
        bar.add_cascade(label="Ajutor", menu=h)
        self.config(menu=bar)

    # -- layout
    def _build_ui(self) -> None:
        top = ttk.LabelFrame(self, text=" Căutare instantanee (Ctrl+F) · caută în cod, marcaj SMD, capsulă, specificații, locație, note, echivalențe, furnizor ", padding=8)
        top.pack(fill=tk.X, padx=10, pady=(8, 4))

        ttk.Label(top, text="Caută:").pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._schedule_refresh())
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=32, font=(self.ui_font, 11))
        self.search_entry.pack(side=tk.LEFT, padx=(0, 12))
        self.search_entry.focus_set()

        ttk.Label(top, text="Categorie:").pack(side=tk.LEFT, padx=(0, 4))
        self.cat_filter_var = tk.StringVar(value="Toate")
        self.cat_cb = ttk.Combobox(top, textvariable=self.cat_filter_var, state="readonly", width=32)
        self.cat_cb.pack(side=tk.LEFT, padx=(0, 12))
        self.cat_cb.bind("<<ComboboxSelected>>", lambda _: self.refresh_table())

        ttk.Label(top, text="Stare:").pack(side=tk.LEFT, padx=(0, 4))
        self.state_filter_var = tk.StringVar(value="Toate")
        self.state_cb = ttk.Combobox(top, textvariable=self.state_filter_var, state="readonly", width=20)
        self.state_cb.pack(side=tk.LEFT, padx=(0, 12))
        self.state_cb.bind("<<ComboboxSelected>>", lambda _: self.refresh_table())

        self.low_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Doar sub stoc minim", variable=self.low_var, command=self.refresh_table).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="Resetează filtre", command=self._reset_filters).pack(side=tk.LEFT, padx=4)
        self._refresh_choices()

        frame = ttk.Frame(self, padding=(10, 4))
        frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(frame, columns=self.COLS, show="headings", selectmode="browse")
        for col, (text, width, anchor) in self.HEADERS.items():
            self.tree.heading(col, text=text, command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=width, anchor=anchor, stretch=col in ("specificatii", "provenienta_note"))
        ys = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        xs = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.tag_configure("zero_stock", foreground="#999999")
        self.tree.tag_configure("used_part", foreground="#004d99")
        self.tree.tag_configure("low_stock", foreground="#a40000", background="#fde8e8")
        self.tree.tag_configure("defect", foreground="#8a5a5a")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _: self._open_datasheet())
        self.tree.bind("<Delete>", lambda _: self._delete_item())
        self.tree.bind("<Button-3>", self._context_menu)
        self.tree.bind("<Button-2>", self._context_menu)

        self.ctx = tk.Menu(self, tearoff=0)
        self.ctx.add_command(label="Deschide datasheet (F3)", command=self._open_datasheet)
        self.ctx.add_command(label="+1 stoc", command=lambda: self._quick_qty(1))
        self.ctx.add_command(label="-1 stoc (folosit)", command=lambda: self._quick_qty(-1))
        self.ctx.add_separator()
        self.ctx.add_command(label="Duplică în formular", command=self._duplicate_item)
        self.ctx.add_command(label="Istoric piesă", command=lambda: self._show_history(False))
        self.ctx.add_command(label="Etichetă", command=lambda: self._labels(True))
        self.ctx.add_separator()
        self.ctx.add_command(label="Șterge piesa", command=self._delete_item)

        self.status_lbl = ttk.Label(self, anchor="w", font=(self.ui_font, 9, "italic"))
        self.status_lbl.pack(fill=tk.X, padx=12)

        form = ttk.LabelFrame(self, text=" Adăugare / Editare componentă ", padding=10)
        form.pack(fill=tk.X, padx=10, pady=(4, 10))
        for i in (3, 5, 7):
            form.columnconfigure(i, weight=1)

        self.vars = {k: tk.StringVar() for k in self.FORM_KEYS}
        self._reset_form_defaults()
        self.prefix_var = tk.StringVar(value="P1-")

        def lbl(r: int, c: int, text: str) -> None:
            ttk.Label(form, text=text).grid(row=r, column=c, sticky="w", padx=4, pady=3)

        def ent(r: int, c: int, key: str, width: int, span: int = 1) -> ttk.Entry:
            e = ttk.Entry(form, textvariable=self.vars[key], width=width)
            e.grid(row=r, column=c, columnspan=span, sticky="ew" if span > 1 else "w", padx=4, pady=3)
            return e

        lbl(0, 0, "Locație:")
        loc = ttk.Frame(form)
        loc.grid(row=0, column=1, sticky="w", padx=4, pady=3)
        self.loc_entry = ttk.Entry(loc, textvariable=self.vars["locatie"], width=10)
        self.loc_entry.pack(side=tk.LEFT)
        self.prefix_cb = ttk.Combobox(loc, textvariable=self.prefix_var, width=7, values=self.db.known_prefixes() or ["P1-"])
        self.prefix_cb.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(loc, text="Auto", width=5, command=self._auto_location).pack(side=tk.LEFT, padx=(4, 0))
        lbl(0, 2, "Cod piesă (Part No.):")
        ent(0, 3, "part_number", 24)
        lbl(0, 4, "Marcaj SMD (top mark):")
        ent(0, 5, "smd_marking", 16)
        lbl(0, 6, "Capsulă (SOT-23, SOIC-8):")
        ent(0, 7, "capsula", 16)

        lbl(1, 0, "Categorie:")
        self.form_cat = ttk.Combobox(form, textvariable=self.vars["categorie"], state="readonly", width=40)
        self.form_cat.grid(row=1, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        lbl(1, 3, "Specificații:")
        ent(1, 4, "specificatii", 36, span=2)
        lbl(1, 6, "Stare:")
        self.form_state = ttk.Combobox(form, textvariable=self.vars["stare"], state="readonly", width=22)
        self.form_state.grid(row=1, column=7, sticky="w", padx=4, pady=3)

        lbl(2, 0, "Cantitate / minim:")
        q = ttk.Frame(form)
        q.grid(row=2, column=1, sticky="w", padx=4, pady=3)
        ttk.Entry(q, textvariable=self.vars["cantitate"], width=6).pack(side=tk.LEFT)
        ttk.Label(q, text=" / ").pack(side=tk.LEFT)
        ttk.Entry(q, textvariable=self.vars["stoc_minim"], width=6).pack(side=tk.LEFT)
        lbl(2, 2, f"Preț / buc ({CURRENCY}):")
        ent(2, 3, "pret", 10)
        lbl(2, 4, "Furnizor:")
        ent(2, 5, "furnizor", 18)
        lbl(2, 6, "Echivalențe:")
        ent(2, 7, "echivalente", 22)

        lbl(3, 0, "Note / Proveniență:")
        ent(3, 1, "provenienta_note", 40, span=3)
        lbl(3, 4, "Datasheet (PDF / URL):")
        ds = ttk.Frame(form)
        ds.grid(row=3, column=5, columnspan=3, sticky="ew", padx=4, pady=3)
        ttk.Entry(ds, textvariable=self.vars["datasheet"], width=26).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(ds, text="📎 Atașează PDF...", command=self._browse_datasheet).pack(side=tk.LEFT, padx=(4, 0))
        self._refresh_choices()

        btns = ttk.Frame(form)
        btns.grid(row=4, column=0, columnspan=8, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="Adaugă ca piesă nouă (Ctrl+N)", style="Accent.TButton", command=self._add_item).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Salvează modificările (Ctrl+S)", command=self._update_item).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Duplică", command=self._duplicate_item).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Golește", command=self._clear_form).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="📄 Datasheet (F3)", style="Datasheet.TButton", command=self._open_datasheet).pack(side=tk.LEFT, padx=12)
        ttk.Button(btns, text="Istoric", command=lambda: self._show_history(False)).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Șterge piesa", command=self._delete_item).pack(side=tk.LEFT, padx=8)

        quick = ttk.Frame(form)
        quick.grid(row=5, column=0, columnspan=8, sticky="ew", pady=(8, 0))
        ttk.Label(quick, text="Mișcare rapidă de stoc pentru piesa selectată · Fișă / lucrare (apare în istoric):").pack(side=tk.LEFT, padx=(4, 6))
        self.ref_var = tk.StringVar()
        ttk.Entry(quick, textvariable=self.ref_var, width=24).pack(side=tk.LEFT, padx=4)
        ttk.Button(quick, text="-1 Stoc (folosit)", width=16, command=lambda: self._quick_qty(-1)).pack(side=tk.LEFT, padx=4)
        ttk.Button(quick, text="+1 Stoc", width=9, command=lambda: self._quick_qty(1)).pack(side=tk.LEFT, padx=4)

    def _bind_shortcuts(self) -> None:
        self.bind("<F3>", lambda _: self._open_datasheet())
        self.bind("<Control-f>", lambda _: (self.search_entry.focus_set(), self.search_entry.select_range(0, tk.END)))
        self.bind("<Control-n>", lambda _: self._new_entry())
        self.bind("<Control-s>", lambda _: self._save())
        self.search_entry.bind("<Escape>", lambda _: self.search_var.set(""))

    # -- tabel
    def _schedule_refresh(self) -> None:
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(150, self.refresh_table)

    def _sort_by(self, col: str) -> None:
        self.sort_rev = (not self.sort_rev) if col == self.sort_col else False
        self.sort_col = col
        self.refresh_table()

    def _refresh_choices(self) -> None:
        cats, states = self.db.all_categories(), self.db.all_states()
        self.cat_cb["values"] = ["Toate"] + cats
        self.state_cb["values"] = ["Toate"] + states
        if hasattr(self, "form_cat"):
            self.form_cat["values"] = cats
            self.form_state["values"] = states
        if hasattr(self, "prefix_cb"):
            self.prefix_cb["values"] = self.db.known_prefixes() or ["P1-"]

    @staticmethod
    def _row_tags(r: sqlite3.Row) -> tuple:
        if r["stare"] == STATE_DEFECT:
            return ("defect",)
        if r["stoc_minim"] > 0 and r["cantitate"] <= r["stoc_minim"]:
            return ("low_stock",)
        if r["cantitate"] == 0:
            return ("zero_stock",)
        if r["stare"] != STATE_NEW:
            return ("used_part",)
        return ()

    def refresh_table(self) -> None:
        self._after_id = None
        rows = self.db.search(self.search_var.get(), self.cat_filter_var.get(), self.state_filter_var.get(), self.low_var.get())
        index = self.resolver.index()
        recs = [(r, self.resolver.badge(r["part_number"], r["datasheet"] or "", index)) for r in rows]

        def key(rec: tuple) -> Any:
            r, badge = rec
            c = self.sort_col
            if c in ("id", "cantitate", "stoc_minim"):
                return r[c]
            return natural_key(badge if c == "ds_status" else (r[c] or ""))
        recs.sort(key=key, reverse=self.sort_rev)

        self.tree.delete(*self.tree.get_children())
        total_qty, value, low, zero = 0, 0.0, 0, 0
        for r, badge in recs:
            tags = self._row_tags(r)
            total_qty += r["cantitate"]
            value += r["cantitate"] * (r["pret"] or 0)
            low += "low_stock" in tags
            zero += r["cantitate"] == 0
            vals = [badge if c == "ds_status" else r[c] for c in self.COLS]
            self.tree.insert("", tk.END, iid=str(r["id"]), values=vals, tags=tags)
        self._shown = rows

        for col, (text, _, _) in self.HEADERS.items():
            arrow = (" ▼" if self.sort_rev else " ▲") if col == self.sort_col else ""
            self.tree.heading(col, text=text + arrow)
        if self.selected_id is not None and self.tree.exists(str(self.selected_id)):
            self.tree.selection_set(str(self.selected_id))
            self.tree.see(str(self.selected_id))
        val_txt = f" | Valoare ≈ {value:,.2f} {CURRENCY}".replace(",", " ") if value else ""
        self._status(f"Afișate: {len(rows)} repere | Total bucăți: {total_qty}{val_txt} | Sub/la minim: {low} | Stoc 0: {zero}")

    def _status(self, text: str) -> None:
        self.status_lbl.config(text=text)

    def _on_select(self, _event: Any) -> None:
        sel = self.tree.selection()
        if not sel or int(sel[0]) == self.selected_id:
            return
        self._load_form(int(sel[0]))

    def _context_menu(self, event: Any) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.ctx.tk_popup(event.x_root, event.y_root)

    # -- formular
    def _reset_form_defaults(self) -> None:
        for k in self.FORM_KEYS:
            self.vars[k].set("")
        self.vars["categorie"].set(CATEGORIES[0])
        self.vars["stare"].set(STATE_NEW)
        self.vars["cantitate"].set("1")
        self.vars["stoc_minim"].set("0")
        self.vars["pret"].set("0")

    def _load_form(self, comp_id: int) -> None:
        row = self.db.get(comp_id)
        if row is None:
            return
        self.selected_id = comp_id
        for k in self.FORM_KEYS:
            v = row[k]
            if k == "pret":
                v = f"{float(v or 0):.2f}".rstrip("0").rstrip(".") or "0"
            self.vars[k].set("" if v is None else str(v))

    def _collect_form(self) -> Componenta | None:
        loc = self.vars["locatie"].get().strip().upper()
        pn = self.vars["part_number"].get().strip()
        if not loc or not pn:
            messagebox.showwarning("Câmpuri obligatorii", "Completează cel puțin „Locație” și „Cod piesă”.")
            return None
        for key, label in (("cantitate", "Cantitatea"), ("stoc_minim", "Stocul minim")):
            raw = self.vars[key].get().strip()
            if not re.fullmatch(r"\d+", raw or "0"):
                messagebox.showwarning("Valoare invalidă", f"{label} trebuie să fie un număr întreg, zero sau mai mare.")
                return None
        pret_raw = self.vars["pret"].get().strip() or "0"
        if to_float(pret_raw, -1) < 0:
            messagebox.showwarning("Preț invalid", "Prețul trebuie să fie un număr (ex. 2,50).")
            return None
        v = {k: self.vars[k].get().strip() for k in self.FORM_KEYS}
        return Componenta(
            id=self.selected_id, locatie=loc, part_number=pn, smd_marking=v["smd_marking"],
            categorie=self.vars["categorie"].get(), capsula=v["capsula"].upper(),
            specificatii=v["specificatii"], stare=self.vars["stare"].get(),
            cantitate=to_int(v["cantitate"] or "0"), provenienta_note=v["provenienta_note"],
            datasheet=v["datasheet"], stoc_minim=to_int(v["stoc_minim"] or "0"),
            pret=to_float(pret_raw), furnizor=v["furnizor"], echivalente=v["echivalente"],
        )

    def _clear_form(self) -> None:
        self.selected_id = None
        self.tree.selection_remove(self.tree.selection())
        keep_cat, keep_state = self.vars["categorie"].get(), self.vars["stare"].get()
        self._reset_form_defaults()
        self.vars["categorie"].set(keep_cat)
        self.vars["stare"].set(keep_state)

    def _new_entry(self) -> None:
        self._clear_form()
        self.loc_entry.focus_set()

    def _auto_location(self) -> None:
        self.vars["locatie"].set(self.db.suggest_next_location(self.prefix_var.get().strip() or "P1-"))

    # -- acțiuni
    def _save(self) -> None:
        (self._update_item if self.selected_id is not None else self._add_item)()

    def _add_item(self) -> None:
        comp = self._collect_form()
        if comp is None:
            return
        comp.id = None
        ref = self.ref_var.get().strip()
        dup = self.db.find_duplicate(comp.part_number, comp.locatie, comp.stare)
        if dup is not None:
            ans = messagebox.askyesnocancel(
                "Piesă existentă",
                f"{dup['part_number']} există deja în {dup['locatie']} (stare: {dup['stare']}, stoc {dup['cantitate']}).\n\n"
                f"DA = adaugă {comp.cantitate} buc. la piesa existentă\n"
                f"NU = creează un rând separat\nAnulare = nu face nimic")
            if ans is None:
                return
            if ans:
                new = self.db.adjust_qty(dup["id"], comp.cantitate, ref, "Intrare stoc")
                self.selected_id = dup["id"]
                self._load_form(dup["id"])
                self.refresh_table()
                self._status(f"{dup['part_number']}: stoc {dup['cantitate']} → {new}")
                return
        new_id = self.db.add(comp, ref)
        self._refresh_choices()
        self._clear_form()
        self.refresh_table()
        if self.tree.exists(str(new_id)):
            self.tree.see(str(new_id))
        self._status(f"Adăugat: {comp.part_number} în {comp.locatie}")

    def _update_item(self) -> None:
        if self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează mai întâi un rând din tabel pentru a-l modifica.")
            return
        comp = self._collect_form()
        if comp is None:
            return
        self.db.update(comp, self.ref_var.get().strip())
        self._refresh_choices()
        self.refresh_table()
        self._status(f"Salvat: {comp.part_number}")

    def _duplicate_item(self) -> None:
        if self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă pe care s-o duplici.")
            return
        self.selected_id = None
        self.tree.selection_remove(self.tree.selection())
        self.vars["locatie"].set("")
        self.vars["cantitate"].set("1")
        self.loc_entry.focus_set()
        self._status("Piesă duplicată în formular: completează locația și apasă „Adaugă ca piesă nouă”.")

    def _delete_item(self) -> None:
        if self.selected_id is None:
            return
        row = self.db.get(self.selected_id)
        if row is None:
            return
        if messagebox.askyesno("Confirmare ștergere",
                               f"Ștergi {row['part_number']} din {row['locatie']} (stoc {row['cantitate']})?\n"
                               "Ștergerea rămâne în istoric, iar backup-ul zilnic permite recuperarea."):
            self.db.delete(self.selected_id)
            self._clear_form()
            self.refresh_table()

    def _quick_qty(self, delta: int) -> None:
        if self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă din tabel pentru a ajusta rapid stocul.")
            return
        row = self.db.get(self.selected_id)
        if row is None:
            return
        new = self.db.adjust_qty(self.selected_id, delta, self.ref_var.get().strip())
        self._load_form(self.selected_id)
        self.refresh_table()
        if new == row["cantitate"]:
            self._status(f"{row['part_number']}: stocul e deja 0.")
        else:
            warn = "  ⚠ la/sub stoc minim" if row["stoc_minim"] and new is not None and new <= row["stoc_minim"] else ""
            self._status(f"{row['part_number']}: stoc {row['cantitate']} → {new}{warn}")

    def _reset_filters(self) -> None:
        self.search_var.set("")
        self.cat_filter_var.set("Toate")
        self.state_filter_var.set("Toate")
        self.low_var.set(False)
        self.refresh_table()

    # -- datasheet
    def _browse_datasheet(self) -> None:
        path_str = filedialog.askopenfilename(title="Selectează fișierul datasheet PDF",
                                              filetypes=[("Documente PDF", "*.pdf"), ("Toate fișierele", "*.*")])
        if not path_str:
            return
        src = Path(path_str).resolve()
        pn = self.vars["part_number"].get().strip()
        ds_dir = self.db.datasheets_dir.resolve()
        if ds_dir not in src.parents:
            default = f"{pn}.pdf" if pn else src.name
            clean = "".join(c for c in default if c.isalnum() or c in ".-_") or src.name
            dest = ds_dir / clean
            if messagebox.askyesno("Copiere în ./datasheets/",
                                   f"Copiez PDF-ul în folderul aplicației?\n{dest}\n(pentru portabilitate între calculatoare)"):
                if dest.exists() and dest.resolve() != src and not messagebox.askyesno("Fișier existent", f"{dest.name} există deja. Îl suprascriu?"):
                    return
                shutil.copy2(src, dest)
                src = dest.resolve()
        try:
            self.vars["datasheet"].set(src.relative_to(self.db.base.resolve()).as_posix())
        except ValueError:
            self.vars["datasheet"].set(str(src))

    def _open_datasheet(self) -> None:
        pn = self.vars["part_number"].get().strip()
        raw = self.vars["datasheet"].get().strip()
        if not pn and not raw:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă din tabel (sau completează codul piesei).")
            return
        if raw.lower().startswith(("http://", "https://")):
            webbrowser.open(raw)
            return
        local = self.resolver.resolve(pn, raw)
        if local is not None:
            try:
                self.resolver.open_native(local)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Eroare la deschiderea PDF-ului", str(exc))
            return
        if messagebox.askyesno("Datasheet local negăsit",
                               f"Nu am găsit un PDF local pentru „{pn}”.\n\nCaut pe web „{pn} datasheet”?\n"
                               f"(Dacă îl salvezi ca datasheets/{pn}.pdf, data viitoare se deschide direct.)"):
            webbrowser.open("https://www.google.com/search?q=" + urllib.parse.quote_plus(f"{pn} datasheet filetype:pdf"))

    def _open_folder(self, folder: Path) -> None:
        folder.mkdir(parents=True, exist_ok=True)
        try:
            self.resolver.open_native(folder)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Eroare", str(exc))

    # -- export / import / backup / etichete
    def _export_csv(self, all_rows: bool) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fișier CSV", "*.csv")],
                                            initialfile=f"inventar_{datetime.now():%Y%m%d}.csv")
        if not path:
            return
        rows = self.db.search() if all_rows else self._shown
        write_csv(path, sorted(rows, key=lambda r: natural_key(r["locatie"])))
        messagebox.showinfo("Export finalizat", f"{len(rows)} rânduri exportate în:\n{path}\n\n(separator „;”, se deschide direct în Excel)")

    def _export_order_list(self) -> None:
        rows = self.db.search(only_low=True)
        if not rows:
            messagebox.showinfo("Listă de comandat", "Nicio piesă nu este la sau sub stocul minim.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fișier CSV", "*.csv")],
                                            initialfile=f"de_comandat_{datetime.now():%Y%m%d}.csv")
        if path:
            write_csv(path, sorted(rows, key=lambda r: natural_key(r["part_number"])), extra_need=True)
            messagebox.showinfo("Export finalizat", f"{len(rows)} repere de comandat exportate în:\n{path}")

    def _import_csv(self) -> None:
        path = filedialog.askopenfilename(title="Import CSV", filetypes=[("Fișier CSV", "*.csv"), ("Toate fișierele", "*.*")])
        if not path:
            return
        try:
            comps, invalid = read_csv(path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Import eșuat", f"Nu am putut citi fișierul:\n{exc}")
            return
        if not comps:
            messagebox.showwarning("Import", f"Nu am găsit rânduri valide (lipsesc Locatie / Part_Number). Rânduri ignorate: {invalid}.")
            return
        if not messagebox.askyesno("Confirmă importul",
                                   f"{len(comps)} rânduri valide ({invalid} ignorate).\n"
                                   "Piesele identice (cod + locație + stare) își adună cantitățile.\n"
                                   "Se face un backup înainte. Continui?"):
            return
        self.db.backup(tag="pre-import")
        added, merged = self.db.import_rows(comps)
        self._refresh_choices()
        self.refresh_table()
        messagebox.showinfo("Import finalizat", f"Adăugate: {added}\nCantități adunate la piese existente: {merged}")

    def _backup_now(self) -> None:
        dest = self.db.backup(tag="manual")
        messagebox.showinfo("Backup", f"Copie de siguranță salvată:\n{dest}")

    def _labels(self, selected_only: bool) -> None:
        if selected_only:
            if self.selected_id is None:
                messagebox.showinfo("Selectează o piesă", "Selectează o piesă pentru etichetă.")
                return
            row = self.db.get(self.selected_id)
            rows = [row] if row else []
        else:
            rows = self._shown
        if not rows:
            return
        out = self.db.base / "etichete.html"
        out.write_text(build_labels_html(sorted(rows, key=lambda r: natural_key(r["locatie"]))), encoding="utf-8")
        webbrowser.open(out.resolve().as_uri())
        self._status(f"Etichete generate: {len(rows)} (se deschid în browser, apoi Ctrl+P).")

    def _show_history(self, all_parts: bool) -> None:
        if not all_parts and self.selected_id is None:
            messagebox.showinfo("Selectează o piesă", "Selectează o piesă pentru a-i vedea istoricul.")
            return
        rows = self.db.history(None if all_parts else self.selected_id)
        win = tk.Toplevel(self)
        win.title("Istoric general mișcări" if all_parts else f"Istoric: {self.vars['part_number'].get()}")
        win.geometry("860x440")
        cols = ("ts", "part_number", "delta", "cantitate_dupa", "motiv", "referinta")
        t = ttk.Treeview(win, columns=cols, show="headings")
        for c, (h, w) in zip(cols, (("Data / ora", 150), ("Piesă", 140), ("Δ", 55), ("Stoc după", 80), ("Motiv", 170), ("Fișă / lucrare", 200))):
            t.heading(c, text=h)
            t.column(c, width=w, anchor="center" if c in ("delta", "cantitate_dupa") else "w")
        for r in rows:
            t.insert("", tk.END, values=(r["ts"], r["part_number"], f"{r['delta']:+d}", r["cantitate_dupa"], r["motiv"], r["referinta"]))
        t.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    def _show_shortcuts(self) -> None:
        messagebox.showinfo("Scurtături", "Ctrl+F  caută\nCtrl+N  piesă nouă (formular gol)\nCtrl+S  salvează (nouă sau modificată)\n"
                            "F3 / dublu-clic  deschide datasheet\nDelete  șterge piesa selectată (în tabel)\n"
                            "Esc (în căutare)  golește căutarea\nClic pe antet  sortează coloana\nClic dreapta  meniu rapid\n\n"
                            "Câmpul „Fișă / lucrare” se atașează la fiecare +1 / -1 și apare în istoric.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inventar Atelier Electronică & ECU")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="calea către fișierul bazei de date")
    args = parser.parse_args()
    if tk is None:
        sys.exit("Tkinter nu este instalat. Pe Windows reinstalează Python bifând „tcl/tk”; pe Linux: sudo apt install python3-tk")
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # text clar pe ecrane HiDPI
        except Exception:  # noqa: BLE001
            pass
    db = DatabaseManager(args.db)
    try:
        db.backup_if_due()
    except Exception:  # noqa: BLE001  (backup-ul nu trebuie să blocheze pornirea)
        pass
    InventarApp(db).mainloop()


if __name__ == "__main__":
    main()
