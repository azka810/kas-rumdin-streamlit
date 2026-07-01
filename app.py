# FINAL FILE - V5.1 BUDGET PADEBUOLO FIX - DD/MM/YYYY + PERGERAKAN BELANJA
from __future__ import annotations

import io
import json
import os
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd
import streamlit as st

APP_TITLE = "V.4 Padebuolo Next"
APP_VERSION = "V.5.5 Padebuolo Next - AI Assistant State Fix"
DEFAULT_PASSWORD = "rumdin123"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
INSTANCE_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "instance"))
INSTANCE_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = Path(os.getenv("DB_PATH", INSTANCE_DIR / "kas_rumdin.db"))

CATEGORIES = [
    "Saldo Awal",
    "Inject Dana / Top Up",
    "Iuran Bulanan",
    "Sewa",
    "Internet",
    "Listrik",
    "Air / PDAM",
    "Laundry",
    "Perlengkapan Rumah",
    "Pemeliharaan",
    "Renovasi",
    "Aset Rumah",
    "Kebersihan",
    "Lainnya",
]
METHODS = ["Kas", "Transfer", "QRIS", "Lainnya"]
DEFAULT_FUNDS = ["Kas Rayhan", "Kas Azka"]
AI_MODEL_DEFAULT = "gpt-4.1-mini"
AI_LOGO_CANDIDATES = [BASE_DIR / "ai_assistant_logo.png", BASE_DIR / "assets" / "ai_assistant_logo.png"]

DEFAULT_BUDGETS = [
    ("Kas Rayhan", "Sewa", 200_000),
    ("Kas Rayhan", "Indihome", 165_000),
    ("Kas Rayhan", "Listrik", 300_000),
    ("Kas Rayhan", "Air", 50_000),
    ("Kas Azka", "Sewa", 245_000),
    ("Kas Azka", "Indihome", 165_000),
    ("Kas Azka", "Listrik", 300_000),
    ("Kas Azka", "Air", 50_000),
]

PERSON_LABELS = {
    "Kas Rayhan": "Rayhan",
    "Kas Azka": "Azka",
}

BUDGET_COMPONENT_ORDER = {
    "Sewa": 1,
    "Indihome": 2,
    "Listrik": 3,
    "Air": 4,
}

ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")
DMY_SLASH_RE = re.compile(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$")
MONTH_NAMES_ID = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}
MONTH_NAMES_ID_LOOKUP = {name.lower(): num for num, name in MONTH_NAMES_ID.items()}
MONTH_NAMES_ID_LOOKUP.update({
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "mei": 5,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "agu": 8,
    "ags": 8,
    "aug": 8,
    "sep": 9,
    "okt": 10,
    "oct": 10,
    "nov": 11,
    "des": 12,
    "dec": 12,
})


def parse_any_date(value: Any) -> pd.Timestamp:
    """Parse dates consistently using Indonesian/DD-MM-YYYY assumptions.

    Database dates are stored as ISO YYYY-MM-DD for sorting. User-facing dates may
    come from Excel/CSV as DD/MM/YYYY, DD-MM-YYYY, Excel serial numbers, or
    Indonesian month names such as "10 April 2026". This function keeps those
    cases from being flipped into US month/day order.
    """
    if value is None:
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(value.date()) if not pd.isna(value) else pd.NaT
    if isinstance(value, datetime):
        return pd.Timestamp(value.date())
    if isinstance(value, date):
        return pd.Timestamp(value)
    try:
        if pd.isna(value):
            return pd.NaT
    except Exception:
        pass

    # Excel often stores dates as serial numbers. Keep this before string parsing.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            num = float(value)
            if 20_000 <= num <= 80_000:
                return pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            pass

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat", "-"}:
        return pd.NaT

    # Drop time suffix if the value is like "2026-04-10 00:00:00".
    text_date_only = text.split()[0] if ISO_DATE_RE.match(text) else text

    # Internal DB / ISO format: YYYY-MM-DD must never be parsed day-first.
    if ISO_DATE_RE.match(text_date_only):
        return pd.to_datetime(text_date_only[:10], format="%Y-%m-%d", errors="coerce")

    # Indonesian numeric format: DD/MM/YYYY or DD-MM-YYYY.
    m = DMY_SLASH_RE.match(text_date_only)
    if m:
        d, mth, y = map(int, m.groups())
        try:
            return pd.Timestamp(year=y, month=mth, day=d)
        except Exception:
            return pd.NaT

    # Indonesian month-name format: 10 April 2026 / 10 Apr 2026.
    m = re.match(r"^(\d{1,2})\s+([A-Za-zÀ-ÿ]+)\s+(\d{4})$", text, flags=re.IGNORECASE)
    if m:
        d = int(m.group(1))
        month_text = m.group(2).lower()
        y = int(m.group(3))
        month_num = MONTH_NAMES_ID_LOOKUP.get(month_text)
        if month_num:
            try:
                return pd.Timestamp(year=y, month=month_num, day=d)
            except Exception:
                return pd.NaT

    # Last fallback still assumes Indonesian/DD-first order.
    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def format_date_short_id(value: Any) -> str:
    parsed = parse_any_date(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def format_date_id(value: Any) -> str:
    """Display date in DD/MM/YYYY format, e.g. 10/04/2026."""
    parsed = parse_any_date(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%d/%m/%Y")


def month_key_from_date(value: Any) -> str:
    parsed = parse_any_date(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m")


def format_month_key_id(month_key: Any) -> str:
    text = str(month_key).strip()
    parsed = pd.to_datetime(f"{text}-01", format="%Y-%m-%d", errors="coerce")
    if pd.isna(parsed):
        return text
    return f"{MONTH_NAMES_ID[int(parsed.month)]} {int(parsed.year)}"


def date_input_id(container: Any, label: str, value: date, **kwargs: Any) -> date:
    """Streamlit date input with Indonesian numeric display when supported."""
    label = f"{label} (DD/MM/YYYY)"
    try:
        return container.date_input(label, value=value, format="DD/MM/YYYY", **kwargs)
    except TypeError:
        return container.date_input(label, value=value, **kwargs)

CATEGORY_RULES = {
    "Saldo Awal": ["saldo awal", "opening balance", "awal kas"],
    "Inject Dana / Top Up": ["inject", "top up", "topup", "tambah dana", "tambahan dana", "isi kas", "isi saldo", "setor", "setoran"],
    "Iuran Bulanan": ["iuran", "kas bulanan", "bulanan", "patungan kas"],
    "Sewa": ["sewa", "kontrakan", "rent", "kos", "kost"],
    "Internet": ["wifi", "wi-fi", "internet", "indihome", "biznet", "first media", "router", "modem"],
    "Listrik": ["listrik", "pln", "token", "pulsa listrik", "kwh"],
    "Air / PDAM": ["pdam", "air", "galon", "aqua", "le minerale", "isi ulang"],
    "Laundry": ["laundry", "cuci", "setrika"],
    "Perlengkapan Rumah": ["lampu", "sapu", "pel", "ember", "keset", "sabun", "tisu", "tisue", "detergen", "deterjen", "piring", "gelas", "sendok", "garpu", "wajan", "panci", "kabel", "terminal", "stop kontak", "sprei", "bantal", "guling", "hanger", "gantungan"],
    "Pemeliharaan": ["service", "servis", "perbaikan", "benerin", "betulin", "maintenance", "tukang", "tambal", "ganti", "instalasi", "pasang"],
    "Renovasi": ["renov", "renovasi", "cat", "semen", "paku", "bor", "triplek", "keramik"],
    "Aset Rumah": ["kipas", "dispenser", "kompor", "kasur", "lemari", "rak", "meja", "kursi", "magic com", "rice cooker", "ac", "kulkas"],
    "Kebersihan": ["kebersihan", "sampah", "cleaning", "karbol", "wipol", "sikat", "lap", "kanebo", "baygon", "obat nyamuk"],
}

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] {
        background: rgba(250, 250, 250, 0.85);
        border: 1px solid rgba(49, 51, 63, 0.12);
        border-radius: 16px;
        padding: 14px 16px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .small-note {
        color: rgba(49, 51, 63, 0.68);
        font-size: 0.92rem;
        line-height: 1.45;
    }
    .pill {
        display: inline-block;
        padding: 0.25rem 0.55rem;
        border-radius: 999px;
        background: rgba(49, 51, 63, 0.08);
        margin-right: 0.25rem;
        margin-bottom: 0.25rem;
        font-size: 0.86rem;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def get_config_value(key: str, default: str = "") -> str:
    """Read config from Streamlit secrets first, then environment, then default."""
    try:
        value = st.secrets.get(key, None)  # type: ignore[attr-defined]
        if value is not None:
            return str(value)
    except Exception:
        pass
    return os.getenv(key, default)


def compact_rp(value: Any) -> str:
    try:
        n = float(value or 0)
    except Exception:
        n = 0.0
    sign = "-" if n < 0 else ""
    n = abs(n)
    if n >= 1_000_000_000:
        return f"{sign}Rp{n/1_000_000_000:.2f} M".replace(".00", "")
    if n >= 1_000_000:
        return f"{sign}Rp{n/1_000_000:.2f} jt".replace(".00", "")
    if n >= 1_000:
        return f"{sign}Rp{n/1_000:.1f} rb".replace(".0", "")
    return f"{sign}Rp{n:,.0f}".replace(",", ".")


def full_rp(value: Any) -> str:
    try:
        n = int(round(float(value or 0)))
    except Exception:
        n = 0
    return f"Rp{n:,}".replace(",", ".")


def clean_amount(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)) and not pd.isna(value):
        return max(0, int(round(value)))
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-"}:
        return 0
    digits = "".join(ch for ch in text if ch.isdigit() or ch == "-")
    try:
        return max(0, int(digits or 0))
    except Exception:
        return 0



def normalize_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).lower().strip()


def infer_category(description: Any, note: Any = "", tx_type: str = "Keluar") -> str:
    """Infer a transaction category from description/note keywords."""
    combined = f"{normalize_text(description)} {normalize_text(note)}"
    if tx_type == "Masuk":
        for category in ["Saldo Awal", "Inject Dana / Top Up", "Iuran Bulanan"]:
            if any(keyword in combined for keyword in CATEGORY_RULES.get(category, [])):
                return category
        return "Iuran Bulanan"
    for category, keywords in CATEGORY_RULES.items():
        if category in {"Saldo Awal", "Inject Dana / Top Up", "Iuran Bulanan"}:
            continue
        if any(keyword in combined for keyword in keywords):
            return category
    return "Lainnya"


def recategorize_by_rules(only_lainnya: bool = True) -> int:
    df = fetch_transactions()
    if df.empty:
        return 0
    now = datetime.utcnow().isoformat(timespec="seconds")
    conn = get_conn()
    changed = 0
    for _, row in df.iterrows():
        current = str(row.get("category") or "Lainnya")
        if only_lainnya and current != "Lainnya":
            continue
        new_category = infer_category(row.get("description", ""), row.get("note", ""), str(row.get("type", "Keluar")))
        if new_category and new_category != current:
            conn.execute("UPDATE transactions SET category=?, updated_at=? WHERE id=?", (new_category, now, int(row["id"])))
            changed += 1
    conn.commit()
    return changed


def recategorize_by_keyword(keywords: List[str], new_category: str, only_lainnya: bool = True) -> int:
    cleaned = [k.strip().lower() for k in keywords if k.strip()]
    if not cleaned:
        return 0
    df = fetch_transactions()
    if df.empty:
        return 0
    now = datetime.utcnow().isoformat(timespec="seconds")
    conn = get_conn()
    changed = 0
    for _, row in df.iterrows():
        current = str(row.get("category") or "Lainnya")
        if only_lainnya and current != "Lainnya":
            continue
        haystack = f"{normalize_text(row.get('description', ''))} {normalize_text(row.get('note', ''))}"
        if any(keyword in haystack for keyword in cleaned):
            conn.execute("UPDATE transactions SET category=?, updated_at=? WHERE id=?", (new_category, now, int(row["id"])))
            changed += 1
    conn.commit()
    return changed


def parse_date(value: Any) -> str:
    parsed = parse_any_date(value)
    if pd.isna(parsed):
        return date.today().isoformat()
    return parsed.date().isoformat()


def infer_swapped_source_month(df: pd.DataFrame) -> int:
    """Infer the original month when old imports were parsed as MM/DD/YYYY.

    Example wrong stored value: 2026-11-04. The day value (4) is likely the
    original month, so the repaired value should be 2026-04-11.
    """
    if df.empty or "date" not in df.columns:
        return int(date.today().month)
    counts: Dict[int, int] = {}
    for raw in df["date"].dropna().astype(str).tolist():
        text = raw.strip()
        if not ISO_DATE_RE.match(text):
            continue
        ts = pd.to_datetime(text[:10], format="%Y-%m-%d", errors="coerce")
        if pd.isna(ts):
            continue
        # Only ambiguous dates can be flipped by an old month-first parser.
        if 1 <= int(ts.day) <= 12 and 1 <= int(ts.month) <= 12 and int(ts.day) != int(ts.month):
            counts[int(ts.day)] = counts.get(int(ts.day), 0) + 1
    if counts:
        return max(counts.items(), key=lambda item: item[1])[0]
    return int(date.today().month)


def swapped_date_candidates(df: pd.DataFrame, correct_month: int) -> pd.DataFrame:
    """Return rows that look like old MM/DD parsing and can be safely previewed.

    If correct_month is 4, a stored ISO date like 2026-11-04 is proposed as
    2026-04-11. Correct dates like 2026-04-11 are not touched because the day is
    11, not the selected correct month 4.
    """
    rows: List[Dict[str, Any]] = []
    if df.empty or "date" not in df.columns:
        return pd.DataFrame(rows)

    for _, row in df.iterrows():
        old_text = str(row.get("date") or "").strip()
        if not ISO_DATE_RE.match(old_text):
            continue
        old_ts = pd.to_datetime(old_text[:10], format="%Y-%m-%d", errors="coerce")
        if pd.isna(old_ts):
            continue
        old_day = int(old_ts.day)
        old_month = int(old_ts.month)
        if old_day != int(correct_month):
            continue
        if old_month == int(correct_month):
            continue
        if not (1 <= old_month <= 12):
            continue
        try:
            new_ts = pd.Timestamp(year=int(old_ts.year), month=int(correct_month), day=old_month)
        except Exception:
            continue
        rows.append(
            {
                "id": int(row.get("id")),
                "Tanggal Lama": old_ts.strftime("%d/%m/%Y"),
                "Tanggal Baru": new_ts.strftime("%d/%m/%Y"),
                "date_lama_iso": old_ts.strftime("%Y-%m-%d"),
                "date_baru_iso": new_ts.strftime("%Y-%m-%d"),
                "Sumber Dana": row.get("fund", ""),
                "Jenis": row.get("type", ""),
                "Kategori": row.get("category", ""),
                "Keterangan": row.get("description", ""),
                "Nominal": int(row.get("amount") or 0),
            }
        )
    return pd.DataFrame(rows)


def apply_swapped_date_fix(candidates: pd.DataFrame) -> int:
    if candidates.empty:
        return 0
    conn = get_conn()
    now = datetime.utcnow().isoformat(timespec="seconds")
    changed = 0
    for _, row in candidates.iterrows():
        conn.execute(
            "UPDATE transactions SET date=?, updated_at=? WHERE id=?",
            (str(row["date_baru_iso"]), now, int(row["id"])),
        )
        changed += 1
    conn.commit()
    return changed


def _key_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _tx_signature(row: Dict[str, Any]) -> Tuple[str, str, int, str]:
    return (
        _key_text(row.get("fund", "")),
        _key_text(row.get("type", "")),
        int(clean_amount(row.get("amount", 0))),
        _key_text(row.get("description", "")),
    )


def _add_occurrence_keys(rows: List[Dict[str, Any]]) -> Dict[Tuple[str, str, int, str, int], Dict[str, Any]]:
    counters: Dict[Tuple[str, str, int, str], int] = {}
    keyed: Dict[Tuple[str, str, int, str, int], Dict[str, Any]] = {}
    for row in rows:
        sig = _tx_signature(row)
        counters[sig] = counters.get(sig, 0) + 1
        keyed[sig + (counters[sig],)] = row
    return keyed


def load_seed_transactions_for_date_repair() -> pd.DataFrame:
    """Load the clean starting transactions used as the source of truth for dates.

    This fixes the real issue where older app versions already saved 11/04/2026
    as 2026-11-04 in SQLite. Merely changing display format cannot repair that.
    The seed file is treated as the source of truth and dates are parsed as DD/MM/YYYY
    or Excel serials.
    """
    tx_path = DATA_DIR / "seed_transactions.csv"
    xlsx_path = DATA_DIR / "Rekap Kas Rumdin.xlsx"
    try:
        if tx_path.exists():
            raw = pd.read_csv(tx_path)
            rows = normalize_transactions_from_df(raw)
            return pd.DataFrame(rows)
        if xlsx_path.exists():
            xl = pd.ExcelFile(xlsx_path)
            sheet_name = "Master Kas" if "Master Kas" in xl.sheet_names else xl.sheet_names[0]
            rows = normalize_transactions_from_df(pd.read_excel(xl, sheet_name=sheet_name))
            return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame()
    return pd.DataFrame()


def seed_date_repair_candidates(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    seed_df = load_seed_transactions_for_date_repair()
    if seed_df.empty:
        return pd.DataFrame()

    seed_rows = seed_df.to_dict("records")
    current_rows = df.to_dict("records")
    seed_keyed = _add_occurrence_keys(seed_rows)
    current_keyed = _add_occurrence_keys(current_rows)

    rows: List[Dict[str, Any]] = []
    for key, cur in current_keyed.items():
        seed = seed_keyed.get(key)
        if seed is None:
            continue
        cur_iso = parse_date(cur.get("date"))
        seed_iso = parse_date(seed.get("date"))
        if cur_iso == seed_iso:
            continue
        rows.append(
            {
                "id": int(cur.get("id")),
                "Tanggal Lama": format_date_id(cur_iso),
                "Tanggal Seharusnya": format_date_id(seed_iso),
                "date_lama_iso": cur_iso,
                "date_baru_iso": seed_iso,
                "Sumber Dana": cur.get("fund", ""),
                "Jenis": cur.get("type", ""),
                "Kategori": cur.get("category", ""),
                "Keterangan": cur.get("description", ""),
                "Nominal": int(clean_amount(cur.get("amount", 0))),
            }
        )
    return pd.DataFrame(rows)


def apply_seed_date_repair(candidates: pd.DataFrame) -> int:
    if candidates.empty:
        return 0
    conn = get_conn()
    now = datetime.utcnow().isoformat(timespec="seconds")
    changed = 0
    for _, row in candidates.iterrows():
        conn.execute(
            "UPDATE transactions SET date=?, updated_at=? WHERE id=?",
            (str(row["date_baru_iso"]), now, int(row["id"])),
        )
        changed += 1
    conn.execute(
        "INSERT OR REPLACE INTO app_meta(key, value) VALUES('date_seed_repair_v48_applied_at', ?)",
        (now,),
    )
    conn.commit()
    return changed


def auto_repair_dates_from_seed_once() -> None:
    """Auto-fix old wrongly parsed seed dates once, without touching new custom rows."""
    try:
        conn = get_conn()
        already = conn.execute(
            "SELECT value FROM app_meta WHERE key='date_seed_repair_v48_applied_at'"
        ).fetchone()
        if already:
            return
        df = pd.DataFrame(run_query("SELECT * FROM transactions ORDER BY date ASC, id ASC"))
        candidates = seed_date_repair_candidates(df)
        if not candidates.empty:
            apply_seed_date_repair(candidates)
        else:
            now = datetime.utcnow().isoformat(timespec="seconds")
            conn.execute(
                "INSERT OR REPLACE INTO app_meta(key, value) VALUES('date_seed_repair_v48_applied_at', ?)",
                (now,),
            )
            conn.commit()
    except Exception:
        # Do not block app startup just because repair preview failed.
        pass


@st.cache_resource(show_spinner=False)
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    conn = get_conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            fund TEXT NOT NULL,
            type TEXT NOT NULL CHECK(type IN ('Masuk', 'Keluar')),
            amount INTEGER NOT NULL CHECK(amount >= 0),
            category TEXT NOT NULL,
            description TEXT NOT NULL,
            method TEXT DEFAULT 'Kas',
            note TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS budgets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            person TEXT NOT NULL,
            component TEXT NOT NULL,
            amount INTEGER NOT NULL CHECK(amount >= 0),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.commit()
    seed_if_empty()
    auto_repair_dates_from_seed_once()
    ensure_standard_budget_once()


def seed_if_empty() -> None:
    conn = get_conn()
    count = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
    if count > 0:
        return
    now = datetime.utcnow().isoformat(timespec="seconds")
    tx_path = DATA_DIR / "seed_transactions.csv"
    bd_path = DATA_DIR / "seed_budgets.csv"
    if tx_path.exists():
        tx_df = pd.read_csv(tx_path)
        for _, row in tx_df.iterrows():
            conn.execute(
                """
                INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    parse_date(row.get("date")),
                    str(row.get("fund", "Kas")),
                    str(row.get("type", "Keluar")),
                    clean_amount(row.get("amount")),
                    str(row.get("category", "Lainnya")),
                    str(row.get("description", "")),
                    str(row.get("method", "Kas")),
                    "" if pd.isna(row.get("note", "")) else str(row.get("note", "")),
                    now,
                    now,
                ),
            )
    if bd_path.exists():
        bd_df = pd.read_csv(bd_path)
        for _, row in bd_df.iterrows():
            conn.execute(
                "INSERT INTO budgets(person, component, amount, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
                (str(row.get("person", "Kas")), str(row.get("component", "")), clean_amount(row.get("amount")), now, now),
            )
    conn.execute("INSERT OR REPLACE INTO app_meta(key, value) VALUES('seed_loaded_at', ?)", (now,))
    conn.commit()


def run_query(query: str, params: Iterable[Any] = ()) -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute(query, tuple(params)).fetchall()
    return [dict(r) for r in rows]


def execute(query: str, params: Iterable[Any] = ()) -> None:
    conn = get_conn()
    conn.execute(query, tuple(params))
    conn.commit()


def add_transaction(tx_date: str, fund: str, tx_type: str, amount: int, category: str, description: str, method: str = "Kas", note: str = "") -> None:
    now = datetime.utcnow().isoformat(timespec="seconds")
    execute(
        """
        INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (tx_date, fund, tx_type, amount, category, description, method, note, now, now),
    )


def update_transaction(tx_id: int, payload: Dict[str, Any]) -> None:
    now = datetime.utcnow().isoformat(timespec="seconds")
    execute(
        """
        UPDATE transactions
        SET date=?, fund=?, type=?, amount=?, category=?, description=?, method=?, note=?, updated_at=?
        WHERE id=?
        """,
        (
            payload["date"],
            payload["fund"],
            payload["type"],
            payload["amount"],
            payload["category"],
            payload["description"],
            payload["method"],
            payload["note"],
            now,
            tx_id,
        ),
    )


def delete_transaction(tx_id: int) -> None:
    execute("DELETE FROM transactions WHERE id=?", (tx_id,))


def add_budget(person: str, component: str, amount: int) -> None:
    now = datetime.utcnow().isoformat(timespec="seconds")
    execute(
        "INSERT INTO budgets(person, component, amount, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
        (person, component, amount, now, now),
    )


def delete_budget(budget_id: int) -> None:
    execute("DELETE FROM budgets WHERE id=?", (budget_id,))


def reset_standard_budgets(mark_meta: bool = True) -> None:
    """Replace budget table with the agreed monthly budget for Rayhan and Azka."""
    conn = get_conn()
    now = datetime.utcnow().isoformat(timespec="seconds")
    with conn:
        conn.execute("DELETE FROM budgets")
        for person, component, amount in DEFAULT_BUDGETS:
            conn.execute(
                "INSERT INTO budgets(person, component, amount, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
                (person, component, int(amount), now, now),
            )
        if mark_meta:
            conn.execute(
                "INSERT OR REPLACE INTO app_meta(key, value) VALUES('standard_budget_v51_applied_at', ?)",
                (now,),
            )


def ensure_standard_budget_once() -> None:
    """One-time migration for deployed databases that still contain wrong budget rows."""
    conn = get_conn()
    try:
        already = conn.execute(
            "SELECT value FROM app_meta WHERE key='standard_budget_v51_applied_at'"
        ).fetchone()
        if already:
            return
        reset_standard_budgets(mark_meta=True)
    except Exception:
        pass


def format_budget_number(value: Any) -> str:
    try:
        n = int(round(float(value or 0)))
    except Exception:
        n = 0
    return f"{n:,}"


def budget_table_for_person(budgets: pd.DataFrame, person: str) -> pd.DataFrame:
    if budgets.empty:
        base = pd.DataFrame(columns=["Komponen", "Total"])
    else:
        base = budgets[budgets["person"].astype(str).eq(person)].copy()
        if base.empty:
            base = pd.DataFrame(columns=["component", "amount"])
        base = base[["component", "amount"]].rename(columns={"component": "Komponen", "amount": "Total"})
    if not base.empty:
        base["Total"] = base["Total"].apply(clean_amount)
        base["_order"] = base["Komponen"].astype(str).map(BUDGET_COMPONENT_ORDER).fillna(99)
        base = base.sort_values(["_order", "Komponen"]).drop(columns=["_order"])
    total = int(base["Total"].sum()) if not base.empty else 0
    out = base.copy()
    out.loc[len(out)] = ["Total", total]
    out["Total"] = out["Total"].apply(format_budget_number)
    return out


def fetch_transactions() -> pd.DataFrame:
    rows = run_query("SELECT * FROM transactions ORDER BY date ASC, id ASC")
    if not rows:
        return pd.DataFrame(columns=["id", "date", "fund", "type", "amount", "category", "description", "method", "note", "netto", "running_balance"])
    df = pd.DataFrame(rows)
    df["amount"] = df["amount"].astype(int)
    df["netto"] = df.apply(lambda r: r["amount"] if r["type"] == "Masuk" else -r["amount"], axis=1)
    df["running_balance"] = df.groupby("fund", sort=False)["netto"].cumsum()
    return df


def fetch_budgets() -> pd.DataFrame:
    rows = run_query("SELECT * FROM budgets ORDER BY person ASC, component ASC, id ASC")
    if not rows:
        return pd.DataFrame(columns=["id", "person", "component", "amount"])
    df = pd.DataFrame(rows)
    df["amount"] = df["amount"].astype(int)
    return df


def get_fund_list(df: pd.DataFrame | None = None) -> List[str]:
    funds = set(DEFAULT_FUNDS)
    if df is not None and not df.empty and "fund" in df.columns:
        funds.update([str(x) for x in df["fund"].dropna().unique().tolist()])
    budget_df = fetch_budgets()
    if not budget_df.empty:
        funds.update([str(x) for x in budget_df["person"].dropna().unique().tolist()])
    return sorted(funds)


def summarize(df: pd.DataFrame) -> Dict[str, Any]:
    if df.empty:
        return {
            "total_in": 0,
            "total_out": 0,
            "saldo": 0,
            "fund_balance": pd.DataFrame(columns=["fund", "saldo"]),
            "monthly_out": pd.DataFrame(columns=["month", "keluar"]),
            "category_out": pd.DataFrame(columns=["category", "keluar"]),
        }
    total_in = int(df.loc[df["type"] == "Masuk", "amount"].sum())
    total_out = int(df.loc[df["type"] == "Keluar", "amount"].sum())
    fund_balance = df.groupby("fund", as_index=False)["netto"].sum().rename(columns={"netto": "saldo"}).sort_values("saldo", ascending=False)
    tmp = df.copy()
    tmp["month"] = tmp["date"].apply(month_key_from_date)
    monthly_out = tmp[tmp["type"] == "Keluar"].groupby("month", as_index=False)["amount"].sum().rename(columns={"amount": "keluar"})
    category_out = tmp[tmp["type"] == "Keluar"].groupby("category", as_index=False)["amount"].sum().rename(columns={"amount": "keluar"}).sort_values("keluar", ascending=False)
    return {
        "total_in": total_in,
        "total_out": total_out,
        "saldo": total_in - total_out,
        "fund_balance": fund_balance,
        "monthly_out": monthly_out,
        "category_out": category_out,
    }


def display_df(df: pd.DataFrame, use_container_width: bool = True, hide_index: bool = True) -> None:
    view = df.copy()
    rename_map = {}
    for col in list(view.columns):
        if str(col).lower() == "date":
            view[col] = view[col].apply(format_date_id)
            rename_map[col] = "Tanggal"
        elif str(col).lower() == "tanggal":
            view[col] = view[col].apply(format_date_id)
    if rename_map:
        view = view.rename(columns=rename_map)
    st.dataframe(view, use_container_width=use_container_width, hide_index=hide_index)


def get_balance_by_fund_label(df: pd.DataFrame, label: str) -> int:
    """Return current balance for a fund label using exact or loose matching."""
    if df.empty:
        return 0
    fund_balance = summarize(df)["fund_balance"]
    if fund_balance.empty:
        return 0

    target = label.strip().lower()
    fb = fund_balance.copy()
    fb["fund_norm"] = fb["fund"].astype(str).str.strip().str.lower()

    exact = fb[fb["fund_norm"] == target]
    if not exact.empty:
        return int(exact["saldo"].sum())

    loose = fb[fb["fund_norm"].str.contains(target, na=False)]
    if not loose.empty:
        return int(loose["saldo"].sum())
    return 0


def ledger_view(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    out = df.copy().sort_values(["date", "id"]).copy()
    out["Tanggal"] = out["date"].apply(format_date_id)
    out["Sumber Dana"] = out["fund"]
    out["Kategori"] = out["category"]
    out["Keterangan"] = out["description"]
    out["Masuk"] = out.apply(lambda r: compact_rp(r["amount"]) if r["type"] == "Masuk" else "-", axis=1)
    out["Keluar"] = out.apply(lambda r: compact_rp(r["amount"]) if r["type"] == "Keluar" else "-", axis=1)
    out["Netto"] = out["netto"].apply(compact_rp)
    out["Metode"] = out["method"].fillna("Kas")
    out["Catatan"] = out["note"].fillna("")
    return out[["id", "Tanggal", "Sumber Dana", "Kategori", "Keterangan", "Masuk", "Keluar", "Netto", "Metode", "Catatan"]]


def detail_ledger_view(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    out = df.copy().sort_values(["date", "id"]).copy()
    out["Tanggal"] = out["date"].apply(format_date_id)
    out["Masuk Detail"] = out.apply(lambda r: full_rp(r["amount"]) if r["type"] == "Masuk" else "-", axis=1)
    out["Keluar Detail"] = out.apply(lambda r: full_rp(r["amount"]) if r["type"] == "Keluar" else "-", axis=1)
    out["Netto Detail"] = out["netto"].apply(full_rp)
    return out[["id", "Tanggal", "fund", "category", "description", "Masuk Detail", "Keluar Detail", "Netto Detail", "method", "note"]].rename(
        columns={"fund": "Sumber Dana", "category": "Kategori", "description": "Keterangan", "method": "Metode", "note": "Catatan"}
    )



def monthly_category_summary(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Return monthly spending by category in long and pivot format."""
    if df.empty:
        empty_long = pd.DataFrame(columns=["Bulan", "Kategori", "Nominal"])
        empty_pivot = pd.DataFrame()
        return empty_long, empty_pivot

    out = df[df["type"] == "Keluar"].copy()
    if out.empty:
        empty_long = pd.DataFrame(columns=["Bulan", "Kategori", "Nominal"])
        empty_pivot = pd.DataFrame()
        return empty_long, empty_pivot

    out["tanggal_dt"] = out["date"].apply(parse_any_date)
    out = out.dropna(subset=["tanggal_dt"])
    if out.empty:
        empty_long = pd.DataFrame(columns=["Bulan", "Kategori", "Nominal"])
        empty_pivot = pd.DataFrame()
        return empty_long, empty_pivot

    out["BulanKey"] = out["tanggal_dt"].dt.to_period("M").astype(str)
    out["Bulan"] = out["BulanKey"].apply(format_month_key_id)
    out["Kategori"] = out["category"].fillna("Lainnya").astype(str)
    grouped = (
        out.groupby(["BulanKey", "Bulan", "Kategori"], as_index=False)["amount"]
        .sum()
        .rename(columns={"amount": "Nominal"})
        .sort_values(["BulanKey", "Kategori"])
    )
    pivot_key = grouped.pivot_table(index="BulanKey", columns="Kategori", values="Nominal", aggfunc="sum", fill_value=0).sort_index()
    pivot = pivot_key.copy()
    pivot.index = [format_month_key_id(idx) for idx in pivot.index]
    pivot.index.name = "Bulan"
    pivot = pivot.astype(int)
    return grouped[["Bulan", "Kategori", "Nominal"]], pivot

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def backup_json_bytes() -> bytes:
    tx = fetch_transactions().drop(columns=["netto", "running_balance"], errors="ignore").to_dict(orient="records")
    budgets = fetch_budgets().to_dict(orient="records")
    payload = {
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "app": APP_TITLE,
        "transactions": tx,
        "budgets": budgets,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def export_excel_bytes() -> bytes:
    tx = fetch_transactions()
    budgets = fetch_budgets()
    summary = summarize(tx)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        detail_ledger_view(tx).to_excel(writer, sheet_name="Buku Besar", index=False)
        if not summary["fund_balance"].empty:
            fb = summary["fund_balance"].copy()
            fb["saldo_detail"] = fb["saldo"].apply(full_rp)
            fb.to_excel(writer, sheet_name="Saldo Kas", index=False)
        budgets.to_excel(writer, sheet_name="Budget Bulanan", index=False)
        monthly_long, monthly_pivot = monthly_category_summary(tx)
        if not monthly_long.empty:
            monthly_long_export = monthly_long.copy()
            monthly_long_export["Nominal Detail"] = monthly_long_export["Nominal"].apply(full_rp)
            monthly_long_export.to_excel(writer, sheet_name="Belanja Bulanan", index=False)
            monthly_pivot.to_excel(writer, sheet_name="Pivot Belanja Bulanan")
        raw = tx.drop(columns=["netto", "running_balance"], errors="ignore").copy()
        if "date" in raw.columns:
            raw.insert(1, "Tanggal", raw["date"].apply(format_date_id))
            raw.insert(2, "Tanggal Singkat", raw["date"].apply(format_date_short_id))
        raw.to_excel(writer, sheet_name="Raw Transactions", index=False)
    return buffer.getvalue()


def normalize_transactions_from_df(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Map common Excel/CSV formats into app transaction rows."""
    if df is None or df.empty:
        return []
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lower_map = {c.lower(): c for c in df.columns}

    def col(*names: str) -> str | None:
        for name in names:
            if name.lower() in lower_map:
                return lower_map[name.lower()]
        return None

    c_date = col("date", "tanggal")
    c_fund = col("fund", "sumber dana", "person")
    c_type = col("type", "jenis")
    c_amount = col("amount", "nominal")
    c_in = col("masuk", "pemasukan")
    c_out = col("keluar", "pengeluaran")
    c_cat = col("category", "kategori")
    c_desc = col("description", "keterangan", "uraian")
    c_method = col("method", "metode")
    c_note = col("note", "catatan")

    rows: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        tx_date = parse_date(r.get(c_date)) if c_date else date.today().isoformat()
        fund = str(r.get(c_fund, "Kas" if not DEFAULT_FUNDS else DEFAULT_FUNDS[0])).strip() if c_fund else DEFAULT_FUNDS[0]
        category = str(r.get(c_cat, "Lainnya")).strip() if c_cat else "Lainnya"
        description = str(r.get(c_desc, category)).strip() if c_desc else category
        method = str(r.get(c_method, "Kas")).strip() if c_method else "Kas"
        note = "" if not c_note or pd.isna(r.get(c_note)) else str(r.get(c_note)).strip()

        if c_in or c_out:
            masuk = clean_amount(r.get(c_in)) if c_in else 0
            keluar = clean_amount(r.get(c_out)) if c_out else 0
            if masuk > 0:
                rows.append({"date": tx_date, "fund": fund, "type": "Masuk", "amount": masuk, "category": category or "Saldo Awal", "description": description or "Pemasukan", "method": method, "note": note})
            if keluar > 0:
                rows.append({"date": tx_date, "fund": fund, "type": "Keluar", "amount": keluar, "category": category or "Lainnya", "description": description or "Pengeluaran", "method": method, "note": note})
        else:
            tx_type = str(r.get(c_type, "Keluar")).strip().title() if c_type else "Keluar"
            tx_type = "Masuk" if tx_type.lower() in {"masuk", "in", "income", "pemasukan"} else "Keluar"
            amount = clean_amount(r.get(c_amount)) if c_amount else 0
            if amount > 0:
                rows.append({"date": tx_date, "fund": fund, "type": tx_type, "amount": amount, "category": category, "description": description, "method": method, "note": note})
    return rows


def normalize_budgets_from_df(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df is None or df.empty:
        return []
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lower_map = {c.lower(): c for c in df.columns}

    def col(*names: str) -> str | None:
        for name in names:
            if name.lower() in lower_map:
                return lower_map[name.lower()]
        return None

    c_person = col("person", "sumber dana", "nama", "orang")
    c_component = col("component", "komponen", "kebutuhan", "kategori")
    c_amount = col("amount", "nominal", "biaya", "budget")
    rows = []
    for _, r in df.iterrows():
        person = str(r.get(c_person, DEFAULT_FUNDS[0])).strip() if c_person else DEFAULT_FUNDS[0]
        component = str(r.get(c_component, "Kebutuhan")).strip() if c_component else "Kebutuhan"
        amount = clean_amount(r.get(c_amount)) if c_amount else 0
        if amount > 0:
            rows.append({"person": person, "component": component, "amount": amount})
    return rows


def import_rows(transactions: List[Dict[str, Any]], budgets: List[Dict[str, Any]], replace: bool = False) -> Tuple[int, int]:
    conn = get_conn()
    now = datetime.utcnow().isoformat(timespec="seconds")
    if replace:
        conn.execute("DELETE FROM transactions")
        conn.execute("DELETE FROM budgets")
    tx_count = 0
    bd_count = 0
    for tx in transactions:
        amount = clean_amount(tx.get("amount"))
        if amount <= 0:
            continue
        conn.execute(
            """
            INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                parse_date(tx.get("date")),
                str(tx.get("fund", DEFAULT_FUNDS[0])).strip() or DEFAULT_FUNDS[0],
                "Masuk" if str(tx.get("type", "Keluar")).lower() == "masuk" else "Keluar",
                amount,
                str(tx.get("category", "Lainnya")).strip() or "Lainnya",
                str(tx.get("description", "Transaksi")).strip() or "Transaksi",
                str(tx.get("method", "Kas")).strip() or "Kas",
                str(tx.get("note", "")).strip(),
                now,
                now,
            ),
        )
        tx_count += 1
    for bd in budgets:
        amount = clean_amount(bd.get("amount"))
        if amount <= 0:
            continue
        conn.execute(
            "INSERT INTO budgets(person, component, amount, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
            (str(bd.get("person", DEFAULT_FUNDS[0])).strip() or DEFAULT_FUNDS[0], str(bd.get("component", "Kebutuhan")).strip() or "Kebutuhan", amount, now, now),
        )
        bd_count += 1
    conn.commit()
    return tx_count, bd_count


def render_login() -> bool:
    password = get_config_value("APP_PASSWORD", DEFAULT_PASSWORD)
    if password == "":
        return True
    if st.session_state.get("authenticated"):
        return True
    st.title("🏠 V.4 Padebuolo Next")
    st.caption("Masuk dulu pak.")
    with st.form("login_form"):
        user_input = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Masuk", type="primary")
    if submitted:
        if user_input == password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Password salah.")
    return False


def page_dashboard(df: pd.DataFrame, budgets: pd.DataFrame) -> None:
    st.title("🏠 Dashboard V.4 Padebuolo Next")
    st.markdown("<div class='small-note'>Angka utama ditampilkan ringkas. Arahkan/keterangan help pada kartu atau buka detail angka untuk nominal penuh.</div>", unsafe_allow_html=True)

    s = summarize(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Saldo Kas", compact_rp(s["saldo"]), help=full_rp(s["saldo"]))
    c2.metric("Total Masuk", compact_rp(s["total_in"]), help=full_rp(s["total_in"]))
    c3.metric("Total Keluar", compact_rp(s["total_out"]), help=full_rp(s["total_out"]))
    monthly_budget = int(budgets["amount"].sum()) if not budgets.empty else 0
    est_month = (s["saldo"] / monthly_budget) if monthly_budget > 0 else 0
    c4.metric("Estimasi Bertahan", f"{est_month:.1f} bulan" if monthly_budget > 0 else "-", help=f"Budget bulanan: {full_rp(monthly_budget)}")

    st.divider()
    left, right = st.columns([1.1, 1])
    with left:
        st.subheader("Saldo per Sumber Dana")
        fb = s["fund_balance"].copy()
        if fb.empty:
            st.info("Belum ada transaksi.")
        else:
            fb_display = fb.copy()
            fb_display["Saldo"] = fb_display["saldo"].apply(compact_rp)
            display_df(fb_display[["fund", "Saldo"]].rename(columns={"fund": "Sumber Dana"}))
            chart_fb = fb.set_index("fund")[["saldo"]]
            st.bar_chart(chart_fb)
    with right:
        st.subheader("Pengeluaran per Kategori")
        cat = s["category_out"].head(8).copy()
        if cat.empty:
            st.info("Belum ada pengeluaran.")
        else:
            cat_display = cat.copy()
            cat_display["Keluar"] = cat_display["keluar"].apply(compact_rp)
            display_df(cat_display[["category", "Keluar"]].rename(columns={"category": "Kategori"}))

    st.subheader("Tren Pengeluaran Bulanan")
    monthly = s["monthly_out"].copy()
    if monthly.empty:
        st.info("Belum ada data bulanan.")
    else:
        st.line_chart(monthly.set_index("month")[["keluar"]])

    with st.expander("Detail angka penuh"):
        detail = {
            "Saldo Kas": full_rp(s["saldo"]),
            "Total Masuk": full_rp(s["total_in"]),
            "Total Keluar": full_rp(s["total_out"]),
            "Budget Bulanan": full_rp(monthly_budget),
            "Estimasi Kas Bertahan": f"{est_month:.2f} bulan" if monthly_budget > 0 else "-",
        }
        display_df(pd.DataFrame(detail.items(), columns=["Indikator", "Detail"]))


def page_input(df: pd.DataFrame) -> None:
    st.title("➕ Input Transaksi")
    funds = get_fund_list(df)

    st.subheader("Mode Cepat")
    c1, c2, c3 = st.columns(3)
    quick_mode = c1.selectbox("Jenis input cepat", ["Inject Dana / Top Up", "Pengeluaran Split", "Transaksi Biasa"])
    tx_date = date_input_id(c2, "Tanggal", value=date.today())
    method = c3.selectbox("Metode", METHODS, index=0)

    if quick_mode == "Inject Dana / Top Up":
        with st.form("inject_form", clear_on_submit=True):
            col1, col2, col3 = st.columns([1, 1, 2])
            fund = col1.selectbox("Masuk ke sumber dana", funds)
            amount = col2.number_input("Nominal inject", min_value=0, step=10_000, format="%d")
            description = col3.text_input("Keterangan", value=f"Inject dana {fund}")
            note = st.text_area("Catatan", placeholder="Opsional")
            submitted = st.form_submit_button("Simpan Inject Dana", type="primary")
        if submitted:
            if amount <= 0 or not description.strip():
                st.error("Nominal dan keterangan wajib diisi.")
            else:
                add_transaction(tx_date.isoformat(), fund, "Masuk", int(amount), "Inject Dana / Top Up", description.strip(), method, note.strip())
                st.success(f"Inject dana {compact_rp(amount)} ke {fund} tersimpan.")
                st.rerun()

    elif quick_mode == "Pengeluaran Split":
        with st.form("split_form", clear_on_submit=True):
            st.caption("Cocok buat belanja bareng. Nominal total akan dibagi rata/proporsi ke sumber dana terpilih.")
            col1, col2 = st.columns([1.2, 1])
            selected_funds = col1.multiselect("Dibagi ke", funds, default=funds[:2])
            total_amount = col2.number_input("Nominal total", min_value=0, step=10_000, format="%d")
            col3, col4 = st.columns([1, 2])
            category = col3.selectbox("Kategori", CATEGORIES, index=CATEGORIES.index("Lainnya"))
            description = col4.text_input("Keterangan", placeholder="Misal: Split beli galon / bayar internet")
            note = st.text_area("Catatan", placeholder="Opsional")
            submitted = st.form_submit_button("Simpan Split", type="primary")
        if submitted:
            if not selected_funds:
                st.error("Pilih minimal satu sumber dana.")
            elif total_amount <= 0 or not description.strip():
                st.error("Nominal dan keterangan wajib diisi.")
            else:
                base = int(total_amount) // len(selected_funds)
                remainder = int(total_amount) % len(selected_funds)
                for idx, fund in enumerate(selected_funds):
                    share = base + (1 if idx < remainder else 0)
                    add_transaction(tx_date.isoformat(), fund, "Keluar", share, category, description.strip(), method, note.strip())
                st.success(f"Split {compact_rp(total_amount)} untuk {len(selected_funds)} sumber dana tersimpan.")
                st.rerun()

    else:
        with st.form("regular_form", clear_on_submit=True):
            col1, col2, col3 = st.columns([1, 1, 1])
            fund = col1.selectbox("Sumber dana", funds)
            tx_type = col2.selectbox("Jenis", ["Keluar", "Masuk"])
            amount = col3.number_input("Nominal", min_value=0, step=10_000, format="%d")
            col4, col5 = st.columns([1, 2])
            default_cat = "Lainnya" if tx_type == "Keluar" else "Iuran Bulanan"
            category = col4.selectbox("Kategori", CATEGORIES, index=CATEGORIES.index(default_cat))
            description = col5.text_input("Keterangan")
            note = st.text_area("Catatan", placeholder="Opsional")
            submitted = st.form_submit_button("Simpan Transaksi", type="primary")
        if submitted:
            if amount <= 0 or not description.strip():
                st.error("Nominal dan keterangan wajib diisi.")
            else:
                add_transaction(tx_date.isoformat(), fund, tx_type, int(amount), category, description.strip(), method, note.strip())
                st.success("Transaksi tersimpan.")
                st.rerun()


def page_ledger(df: pd.DataFrame) -> None:
    st.title("📒 Buku Besar")
    if df.empty:
        st.info("Belum ada transaksi.")
        return

    st.subheader("Rekap Sisa Kas")
    azka_balance = get_balance_by_fund_label(df, "Azka")
    rayhan_balance = get_balance_by_fund_label(df, "Rayhan")
    total_balance = summarize(df)["saldo"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Sisa Kas Azka", compact_rp(azka_balance), help=full_rp(azka_balance))
    c2.metric("Sisa Kas Rayhan", compact_rp(rayhan_balance), help=full_rp(rayhan_balance))
    c3.metric("Total Sisa Kas", compact_rp(total_balance), help=full_rp(total_balance))
    st.caption("Rekap sisa kas dihitung dari seluruh transaksi masuk dikurangi keluar per sumber dana. Tabel buku besar di bawah hanya menampilkan transaksi, bukan saldo berjalan per baris.")

    filters = st.container(border=True)
    with filters:
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1.4])
        funds = ["Semua"] + get_fund_list(df)
        selected_fund = c1.selectbox("Sumber dana", funds)
        selected_type = c2.selectbox("Jenis", ["Semua", "Masuk", "Keluar"])
        selected_category = c3.selectbox("Kategori", ["Semua"] + sorted(df["category"].dropna().astype(str).unique().tolist()))
        keyword = c4.text_input("Cari keterangan/catatan")

    fdf = df.copy()
    if selected_fund != "Semua":
        fdf = fdf[fdf["fund"] == selected_fund]
    if selected_type != "Semua":
        fdf = fdf[fdf["type"] == selected_type]
    if selected_category != "Semua":
        fdf = fdf[fdf["category"] == selected_category]
    if keyword.strip():
        kw = keyword.strip().lower()
        fdf = fdf[fdf["description"].str.lower().str.contains(kw, na=False) | fdf["note"].fillna("").str.lower().str.contains(kw, na=False)]

    st.caption("Buku besar transaksi: Masuk, Keluar, dan Netto. Sisa kas ditaruh sebagai rekap di atas supaya tidak membingungkan per baris.")
    display_df(ledger_view(fdf))

    with st.expander("Lihat nominal penuh"):
        display_df(detail_ledger_view(fdf))

    st.divider()
    st.subheader("Edit / Hapus Transaksi")
    tx_ids = fdf["id"].astype(int).tolist()
    if not tx_ids:
        st.info("Tidak ada transaksi sesuai filter.")
        return
    selected_id = st.selectbox("Pilih ID transaksi", tx_ids, format_func=lambda x: f"ID {x}")
    row = df[df["id"] == selected_id].iloc[0].to_dict()
    with st.form("edit_tx_form"):
        c1, c2, c3 = st.columns(3)
        tx_date = date_input_id(c1, "Tanggal", value=parse_any_date(row["date"]).date(), key="edit_date")
        fund = c2.text_input("Sumber dana", value=str(row["fund"]))
        tx_type = c3.selectbox("Jenis", ["Masuk", "Keluar"], index=0 if row["type"] == "Masuk" else 1)
        c4, c5, c6 = st.columns([1, 1, 2])
        amount = c4.number_input("Nominal", min_value=0, value=int(row["amount"]), step=10_000, format="%d")
        category = c5.selectbox("Kategori", CATEGORIES, index=CATEGORIES.index(row["category"]) if row["category"] in CATEGORIES else CATEGORIES.index("Lainnya"))
        description = c6.text_input("Keterangan", value=str(row["description"]))
        c7, c8 = st.columns([1, 2])
        method = c7.selectbox("Metode", METHODS, index=METHODS.index(row.get("method") or "Kas") if (row.get("method") or "Kas") in METHODS else 0)
        note = c8.text_input("Catatan", value=str(row.get("note") or ""))
        save, delete = st.columns([1, 1])
        submitted = save.form_submit_button("Simpan Perubahan", type="primary")
        deleted = delete.form_submit_button("Hapus Transaksi")
    if submitted:
        if amount <= 0 or not description.strip() or not fund.strip():
            st.error("Sumber dana, nominal, dan keterangan wajib diisi.")
        else:
            update_transaction(
                int(selected_id),
                {
                    "date": tx_date.isoformat(),
                    "fund": fund.strip(),
                    "type": tx_type,
                    "amount": int(amount),
                    "category": category,
                    "description": description.strip(),
                    "method": method,
                    "note": note.strip(),
                },
            )
            st.success("Transaksi berhasil diubah.")
            st.rerun()
    if deleted:
        delete_transaction(int(selected_id))
        st.warning("Transaksi dihapus.")
        st.rerun()



def page_monthly_category(df: pd.DataFrame) -> None:
    st.title("📊 Pergerakan Belanja Bulanan per Kategori")
    st.caption("Menu ini khusus membaca transaksi jenis Keluar, lalu mengelompokkan pengeluaran per bulan dan kategori belanja.")

    if df.empty:
        st.info("Belum ada transaksi.")
        return

    out = df[df["type"] == "Keluar"].copy()
    if out.empty:
        st.info("Belum ada transaksi pengeluaran.")
        return

    out["tanggal_dt"] = out["date"].apply(parse_any_date)
    out = out.dropna(subset=["tanggal_dt"])
    if out.empty:
        st.warning("Tanggal transaksi belum bisa dibaca. Cek format tanggal di data transaksi.")
        return

    out["BulanKey"] = out["tanggal_dt"].dt.to_period("M").astype(str)
    available_months = sorted(out["BulanKey"].dropna().astype(str).unique().tolist())
    available_categories = sorted(out["category"].fillna("Lainnya").astype(str).unique().tolist())

    with st.container(border=True):
        c1, c2, c3, c4 = st.columns([1, 1, 1.3, 1])
        start_month = c1.selectbox("Dari bulan", available_months, index=0, format_func=format_month_key_id)
        end_month = c2.selectbox("Sampai bulan", available_months, index=len(available_months) - 1, format_func=format_month_key_id)
        selected_categories = c3.multiselect("Kategori", available_categories, default=available_categories)
        selected_fund = c4.selectbox("Sumber dana", ["Semua"] + get_fund_list(df), index=0)

    if start_month > end_month:
        st.error("Bulan awal tidak boleh lebih besar dari bulan akhir.")
        return

    fdf = out[(out["BulanKey"] >= start_month) & (out["BulanKey"] <= end_month)].copy()
    if selected_fund != "Semua":
        fdf = fdf[fdf["fund"] == selected_fund]
    if selected_categories:
        fdf = fdf[fdf["category"].isin(selected_categories)]
    else:
        st.warning("Pilih minimal satu kategori.")
        return

    if fdf.empty:
        st.info("Tidak ada pengeluaran sesuai filter.")
        return

    grouped, pivot = monthly_category_summary(fdf)
    total_spending = int(fdf["amount"].sum())
    monthly_total = pivot.sum(axis=1).sort_values(ascending=False) if not pivot.empty else pd.Series(dtype="int64")
    category_total = pivot.sum(axis=0).sort_values(ascending=False) if not pivot.empty else pd.Series(dtype="int64")

    biggest_month = monthly_total.index[0] if not monthly_total.empty else "-"
    biggest_month_value = int(monthly_total.iloc[0]) if not monthly_total.empty else 0
    biggest_category = category_total.index[0] if not category_total.empty else "-"
    biggest_category_value = int(category_total.iloc[0]) if not category_total.empty else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Belanja", compact_rp(total_spending), help=full_rp(total_spending))
    c2.metric("Bulan Belanja Tertinggi", biggest_month, help=full_rp(biggest_month_value))
    c3.metric("Kategori Terbesar", biggest_category, help=full_rp(biggest_category_value))

    st.divider()
    st.subheader("Grafik Pergerakan Bulanan")
    chart_type = st.radio("Jenis grafik", ["Line chart", "Bar chart"], horizontal=True)
    if chart_type == "Line chart":
        st.line_chart(pivot)
    else:
        st.bar_chart(pivot)

    st.subheader("Tabel Belanja Bulanan per Kategori")
    table = pivot.copy()
    table["Total"] = table.sum(axis=1)
    table_display = table.reset_index().copy()
    for col in table_display.columns:
        if col != "Bulan":
            table_display[col] = table_display[col].apply(compact_rp)
    display_df(table_display)

    with st.expander("Lihat nominal penuh"):
        detail_table = table.reset_index().copy()
        for col in detail_table.columns:
            if col != "Bulan":
                detail_table[col] = detail_table[col].apply(full_rp)
        display_df(detail_table)

    st.subheader("Ranking Kategori Belanja")
    category_rank = category_total.reset_index()
    category_rank.columns = ["Kategori", "Nominal"]
    category_rank["Nominal Ringkas"] = category_rank["Nominal"].apply(compact_rp)
    category_rank["Porsi"] = category_rank["Nominal"].apply(lambda x: f"{(x / total_spending * 100):.1f}%" if total_spending else "0%")
    display_df(category_rank[["Kategori", "Nominal Ringkas", "Porsi"]])

    with st.expander("Data long format / siap pivot"):
        long_display = grouped.copy()
        long_display["Nominal Ringkas"] = long_display["Nominal"].apply(compact_rp)
        long_display["Nominal Detail"] = long_display["Nominal"].apply(full_rp)
        display_df(long_display[["Bulan", "Kategori", "Nominal Ringkas", "Nominal Detail"]])

    c1, c2 = st.columns(2)
    c1.download_button(
        "Download Pivot CSV",
        data=to_csv_bytes(table.reset_index()),
        file_name="pergerakan_belanja_bulanan_pivot.csv",
        mime="text/csv",
        use_container_width=True,
    )
    c2.download_button(
        "Download Long CSV",
        data=to_csv_bytes(grouped),
        file_name="pergerakan_belanja_bulanan_long.csv",
        mime="text/csv",
        use_container_width=True,
    )


def get_openai_model() -> str:
    return get_config_value("OPENAI_MODEL", AI_MODEL_DEFAULT) or AI_MODEL_DEFAULT


def get_openai_key() -> str:
    return get_config_value("OPENAI_API_KEY", "").strip()


def get_ai_logo_path() -> Path | None:
    for path in AI_LOGO_CANDIDATES:
        try:
            if path.exists():
                return path
        except Exception:
            pass
    return None


def df_records_for_ai(df: pd.DataFrame, max_rows: int = 80) -> List[Dict[str, Any]]:
    if df.empty:
        return []
    cols = [c for c in ["date", "fund", "type", "amount", "category", "description", "method", "note"] if c in df.columns]
    out = df.sort_values(["date", "id"], ascending=[False, False]).head(max_rows)[cols].copy()
    if "date" in out.columns:
        out["date"] = out["date"].apply(format_date_id)
    if "amount" in out.columns:
        out["amount"] = out["amount"].apply(lambda x: int(clean_amount(x)))
    return out.to_dict(orient="records")


def ai_context_payload(df: pd.DataFrame, budgets: pd.DataFrame) -> Dict[str, Any]:
    summary = summarize(df)
    payload: Dict[str, Any] = {
        "app": APP_TITLE,
        "format_tanggal": "DD/MM/YYYY",
        "catatan": "AI Assistant hanya membaca data dan memberi insight. AI tidak mengubah transaksi, budget, atau database.",
        "ringkasan_kas": {
            "total_masuk": int(summary.get("total_in", 0)),
            "total_keluar": int(summary.get("total_out", 0)),
            "saldo_total": int(summary.get("saldo", 0)),
        },
        "saldo_per_sumber_dana": [],
        "budget_bulanan": [],
        "belanja_bulanan_per_kategori": [],
        "belanja_per_kategori_total": [],
        "transaksi_terbaru": df_records_for_ai(df, 80),
    }

    if not df.empty and "date" in df.columns:
        dates = df["date"].apply(parse_any_date).dropna()
        if not dates.empty:
            payload["rentang_data"] = {
                "mulai": format_date_id(dates.min()),
                "sampai": format_date_id(dates.max()),
                "jumlah_transaksi": int(len(df)),
            }

    fund_balance = summary.get("fund_balance", pd.DataFrame())
    if isinstance(fund_balance, pd.DataFrame) and not fund_balance.empty:
        payload["saldo_per_sumber_dana"] = [
            {"sumber_dana": str(r["fund"]), "saldo": int(r["saldo"])}
            for _, r in fund_balance.iterrows()
        ]

    if not budgets.empty:
        budget_rows = budgets.copy()
        budget_rows["amount"] = budget_rows["amount"].apply(clean_amount)
        payload["budget_bulanan"] = [
            {"orang": str(r["person"]), "komponen": str(r["component"]), "budget": int(r["amount"])}
            for _, r in budget_rows.iterrows()
        ]
        payload["total_budget_bulanan"] = int(budget_rows["amount"].sum())

    monthly_long, _pivot = monthly_category_summary(df)
    if not monthly_long.empty:
        payload["belanja_bulanan_per_kategori"] = [
            {"bulan": str(r["Bulan"]), "kategori": str(r["Kategori"]), "nominal": int(r["Nominal"])}
            for _, r in monthly_long.iterrows()
        ]
    category_out = summary.get("category_out", pd.DataFrame())
    if isinstance(category_out, pd.DataFrame) and not category_out.empty:
        payload["belanja_per_kategori_total"] = [
            {"kategori": str(r["category"]), "nominal": int(r["keluar"])}
            for _, r in category_out.iterrows()
        ]

    return payload


def local_ai_fallback_answer(question: str, df: pd.DataFrame, budgets: pd.DataFrame) -> str:
    """Fallback insight kalau OPENAI_API_KEY belum dipasang."""
    summary = summarize(df)
    lines = [
        "Mode lokal aktif karena OPENAI_API_KEY belum dipasang. Ini ringkasan otomatis tanpa API:",
        "",
        f"- Total masuk: {full_rp(summary.get('total_in', 0))}",
        f"- Total keluar: {full_rp(summary.get('total_out', 0))}",
        f"- Sisa kas total: {full_rp(summary.get('saldo', 0))}",
    ]
    fb = summary.get("fund_balance", pd.DataFrame())
    if isinstance(fb, pd.DataFrame) and not fb.empty:
        lines.append("")
        lines.append("Saldo per sumber dana:")
        for _, row in fb.iterrows():
            lines.append(f"- {row['fund']}: {full_rp(row['saldo'])}")

    cat = summary.get("category_out", pd.DataFrame())
    if isinstance(cat, pd.DataFrame) and not cat.empty:
        lines.append("")
        lines.append("Kategori belanja terbesar:")
        for _, row in cat.head(5).iterrows():
            lines.append(f"- {row['category']}: {full_rp(row['keluar'])}")

    if not budgets.empty:
        total_budget = int(budgets["amount"].sum())
        saldo = int(summary.get("saldo", 0))
        lines.append("")
        lines.append(f"Total budget bulanan: {full_rp(total_budget)}")
        if total_budget > 0:
            lines.append(f"Estimasi kas bertahan sekitar {saldo / total_budget:.1f} bulan.")

    lines.append("")
    lines.append("Supaya bisa jawab pertanyaan bebas seperti chat AI, pasang OPENAI_API_KEY di Streamlit Secrets.")
    return "\n".join(lines)


def call_openai_ai_assistant(question: str, df: pd.DataFrame, budgets: pd.DataFrame) -> str:
    api_key = get_openai_key()
    if not api_key:
        return local_ai_fallback_answer(question, df, budgets)

    context_payload = ai_context_payload(df, budgets)
    system_prompt = """
Kamu adalah AI Assistant untuk aplikasi kas rumah dinas bernama Padebuolo Next.
Tugasmu membaca ringkasan transaksi, budget, saldo, dan belanja bulanan yang diberikan dalam konteks JSON.
Jawab dalam bahasa Indonesia yang santai, jelas, dan praktis seperti ngobrol dengan pemilik kas.
Jangan mengarang data di luar konteks. Kalau data tidak cukup, bilang data belum cukup.
Jangan memberi instruksi untuk mengubah database kecuali berupa saran. Kamu tidak punya izin melakukan update/delete transaksi.
Prioritaskan jawaban dengan angka rupiah, ringkasan insight, dan langkah praktis.
Format tanggal harus DD/MM/YYYY.
""".strip()

    user_payload = {
        "pertanyaan_user": question,
        "data_kas": context_payload,
    }

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        response = client.responses.create(
            model=get_openai_model(),
            input=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False, default=str)},
            ],
            max_output_tokens=1200,
        )
        text = getattr(response, "output_text", None)
        if text:
            return str(text).strip()

        # Fallback extraction for SDK variants.
        chunks: List[str] = []
        for item in getattr(response, "output", []) or []:
            for content in getattr(item, "content", []) or []:
                maybe_text = getattr(content, "text", None)
                if maybe_text:
                    chunks.append(str(maybe_text))
        return "\n".join(chunks).strip() or "AI belum mengembalikan jawaban. Coba ulangi pertanyaannya."
    except Exception as exc:
        return (
            "Gagal memanggil OpenAI API. Cek OPENAI_API_KEY, OPENAI_MODEL, billing, atau requirements.txt.\n\n"
            f"Detail teknis: {type(exc).__name__}: {exc}"
        )


def page_ai_assistant(df: pd.DataFrame, budgets: pd.DataFrame) -> None:
    st.title("🤖 AI Assistant")
    logo_path = get_ai_logo_path()
    c_logo, c_text = st.columns([0.22, 0.78])
    with c_logo:
        if logo_path:
            st.image(str(logo_path), use_container_width=True)
        else:
            st.markdown("# 🤖")
    with c_text:
        st.markdown("### Asisten baca data kas Padebuolo")
        st.caption("Mode aman: AI hanya membaca data dan memberi insight. AI tidak bisa edit, hapus, atau mengubah database.")

    st.info("Versi ini sengaja tidak memakai komponen chat bawaan Streamlit (`st.chat_input`) supaya tidak kena error dynamic import `ChatInput...js` di Streamlit Cloud.")

    api_key = get_openai_key()
    if not api_key:
        st.warning("OPENAI_API_KEY belum dipasang. AI Assistant akan jalan dalam mode ringkasan lokal. Untuk chat AI penuh, isi OPENAI_API_KEY di Streamlit Secrets.")
    else:
        st.success(f"OpenAI API aktif. Model: {get_openai_model()}")

    with st.expander("Contoh pertanyaan", expanded=False):
        examples = [
            "Ringkas kondisi kas rumah dinas sekarang.",
            "Bulan apa pengeluaran paling besar dan kategori apa penyebabnya?",
            "Sisa kas Azka dan Rayhan masing-masing aman nggak?",
            "Kategori belanja mana yang paling boros dan saran hematnya apa?",
            "Bandingkan realisasi belanja dengan budget bulanan.",
            "Transaksi yang masih Lainnya sebaiknya dikategorikan apa?",
        ]
        st.write("\n".join([f"- {x}" for x in examples]))

    if "ai_messages" not in st.session_state:
        st.session_state.ai_messages = []
    if "ai_question_box" not in st.session_state:
        st.session_state.ai_question_box = ""

    quick_cols = st.columns(4)
    quick_prompts = [
        ("Ringkasan kas", "Ringkas kondisi kas rumah dinas sekarang, termasuk total masuk, total keluar, sisa kas per orang, dan hal yang perlu diperhatikan."),
        ("Kategori boros", "Kategori belanja mana yang paling besar? Jelaskan penyebab yang terlihat dari data dan beri saran hemat."),
        ("Budget check", "Bandingkan saldo dan belanja dengan budget bulanan Rayhan dan Azka. Kas masih aman berapa bulan?"),
        ("Cek Lainnya", "Cek transaksi kategori Lainnya. Beri saran keyword/kategori yang bisa dipakai untuk kategorisasi massal."),
    ]
    for col, (label, prompt) in zip(quick_cols, quick_prompts):
        if col.button(label, use_container_width=True):
            st.session_state.ai_question_box = prompt
            st.rerun()

    st.divider()

    with st.form("ai_assistant_form", clear_on_submit=False):
        question = st.text_area(
            "Tanya AI soal kas rumah dinas",
            key="ai_question_box",
            height=120,
            placeholder="Contoh: Bulan April pengeluaran paling besar kategori apa?",
        )
        submitted = st.form_submit_button("Tanya AI", use_container_width=True)

    if submitted:
        question_clean = str(question or "").strip()
        if not question_clean:
            st.warning("Isi pertanyaannya dulu, pak.")
        else:
            st.session_state.ai_messages.append({"role": "user", "content": question_clean})
            with st.spinner("AI lagi baca data kas..."):
                answer = call_openai_ai_assistant(question_clean, df, budgets)
            st.session_state.ai_messages.append({"role": "assistant", "content": answer})
            st.rerun()

    if st.session_state.ai_messages:
        st.markdown("### Riwayat tanya jawab")
        for i, msg in enumerate(st.session_state.ai_messages, start=1):
            if msg.get("role") == "user":
                st.markdown(f"**🧑 Pertanyaan {i}:**")
                st.markdown(f"> {msg.get('content', '')}")
            else:
                st.markdown("**🤖 Jawaban AI:**")
                st.markdown(msg.get("content", ""))
                st.divider()
    else:
        st.caption("Belum ada riwayat. Pilih tombol cepat atau tulis pertanyaan manual.")

    c1, c2 = st.columns(2)
    if c1.button("Hapus riwayat AI", use_container_width=True):
        st.session_state.ai_messages = []
        st.rerun()
    c2.download_button(
        "Download konteks data untuk AI",
        data=json.dumps(ai_context_payload(df, budgets), ensure_ascii=False, indent=2, default=str).encode("utf-8"),
        file_name="padebuolo_ai_context.json",
        mime="application/json",
        use_container_width=True,
    )

def page_bulk_category(df: pd.DataFrame) -> None:
    st.title("🧠 Kategorisasi Massal")
    st.caption("Pakai menu ini kalau banyak transaksi masih kebaca sebagai Lainnya. Sistem akan membaca kata kunci dari keterangan/catatan transaksi.")
    if df.empty:
        st.info("Belum ada transaksi.")
        return

    total_lainnya = int((df["category"] == "Lainnya").sum()) if "category" in df.columns else 0
    total_rows = len(df)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Transaksi", f"{total_rows}")
    c2.metric("Masih Lainnya", f"{total_lainnya}")
    c3.metric("Proporsi Lainnya", f"{(total_lainnya / total_rows * 100):.1f}%" if total_rows else "0%")

    st.subheader("1) Auto-kategorisasi dari keyword")
    only_lainnya_auto = st.toggle("Hanya ubah transaksi yang kategorinya masih Lainnya", value=True, key="auto_only_lainnya")
    preview = df.copy()
    if only_lainnya_auto:
        preview = preview[preview["category"] == "Lainnya"]
    preview["Kategori Usulan"] = preview.apply(lambda r: infer_category(r.get("description", ""), r.get("note", ""), str(r.get("type", "Keluar"))), axis=1)
    preview = preview[preview["Kategori Usulan"] != preview["category"]]

    st.write(f"Transaksi yang akan berubah kalau tombol diterapkan: **{len(preview)}**")
    if not preview.empty:
        show = preview.copy().head(50)
        show["Tanggal"] = show["date"].apply(format_date_id)
        show["Nominal"] = show["amount"].apply(compact_rp)
        display_df(show[["id", "Tanggal", "fund", "type", "Nominal", "category", "Kategori Usulan", "description"]].rename(columns={"fund": "Sumber Dana", "type": "Jenis", "category": "Kategori Lama", "description": "Keterangan"}))
    else:
        st.info("Belum ada transaksi yang cocok dengan keyword default.")

    if st.button("Terapkan Auto-kategorisasi", type="primary"):
        changed = recategorize_by_rules(only_lainnya=only_lainnya_auto)
        st.success(f"Berhasil mengubah {changed} transaksi.")
        st.rerun()

    st.divider()
    st.subheader("2) Ubah massal berdasarkan kata kunci sendiri")
    st.caption("Contoh: isi keyword `wifi, internet, indihome`, pilih kategori Internet, lalu terapkan.")
    with st.form("manual_keyword_category"):
        c1, c2 = st.columns([2, 1])
        keywords_text = c1.text_input("Keyword keterangan/catatan", placeholder="wifi, internet, laundry, galon")
        new_category = c2.selectbox("Ubah jadi kategori", CATEGORIES, index=CATEGORIES.index("Lainnya"))
        only_lainnya_manual = st.checkbox("Hanya ubah yang masih Lainnya", value=True)
        submitted = st.form_submit_button("Terapkan Keyword Ini", type="primary")
    if submitted:
        keywords = [x.strip() for x in keywords_text.split(",") if x.strip()]
        if not keywords:
            st.error("Isi minimal satu keyword.")
        elif new_category == "Lainnya":
            st.error("Pilih kategori selain Lainnya supaya ada perubahan.")
        else:
            changed = recategorize_by_keyword(keywords, new_category, only_lainnya=only_lainnya_manual)
            st.success(f"Berhasil mengubah {changed} transaksi menjadi {new_category}.")
            st.rerun()

    st.divider()
    st.subheader("Daftar keyword default")
    rules_rows = []
    for category, keywords in CATEGORY_RULES.items():
        rules_rows.append({"Kategori": category, "Keyword": ", ".join(keywords)})
    display_df(pd.DataFrame(rules_rows))


def page_settings(df: pd.DataFrame) -> None:
    st.title("⚙️ Pengaturan & Info Deploy")
    st.write("Lokasi database aktif:")
    st.code(str(DB_PATH))
    st.write("Sumber dana aktif:")
    st.markdown(" ".join(f"<span class='pill'>{fund}</span>" for fund in get_fund_list(df)), unsafe_allow_html=True)
    st.info("Untuk tambah sumber dana baru, cukup input transaksi dengan nama sumber dana baru di halaman Buku Besar > edit, atau import file dengan sumber dana tersebut.")

    st.subheader("Sinkron Tanggal dari Excel Awal")
    st.caption("Ini sumber truth-nya dari file data/seed_transactions.csv atau data/Rekap Kas Rumdin.xlsx. Cocok buat kasus data lama terlanjur kebaca 04/11/2026 padahal harusnya 11/04/2026. Transaksi tambahan yang tidak ada di Excel awal tidak disentuh.")
    seed_candidates = seed_date_repair_candidates(df)
    if seed_candidates.empty:
        st.success("Tanggal transaksi yang cocok dengan Excel awal sudah sinkron.")
    else:
        st.warning(f"Ada {len(seed_candidates)} transaksi yang tanggalnya beda dari Excel awal. Preview dulu sebelum diperbaiki.")
        preview_seed = seed_candidates.copy()
        preview_seed["Nominal"] = preview_seed["Nominal"].apply(compact_rp)
        display_df(preview_seed[["id", "Tanggal Lama", "Tanggal Seharusnya", "Sumber Dana", "Jenis", "Kategori", "Keterangan", "Nominal"]])
        confirm_seed = st.checkbox("Saya mau paksa tanggal ikut Excel awal", key="confirm_seed_date_repair")
        if st.button("Sinkronkan Tanggal dari Excel Awal", type="primary", disabled=not confirm_seed):
            changed = apply_seed_date_repair(seed_candidates)
            st.success(f"Berhasil menyinkronkan {changed} tanggal dari Excel awal.")
            st.rerun()

    st.divider()
    st.subheader("Perbaikan Tanggal Ketuker")
    st.caption("Pakai ini sekali saja kalau data lama terlanjur kebaca MM/DD/YYYY. Contoh salah: 11/04/2026 tampil sebagai 04/11/2026. Perbaikan akan mengubah 2026-11-04 menjadi 2026-04-11.")
    if df.empty:
        st.info("Belum ada transaksi untuk dicek.")
    else:
        inferred_month = infer_swapped_source_month(df)
        months = list(range(1, 13))
        default_index = months.index(inferred_month) if inferred_month in months else int(date.today().month) - 1
        correct_month = st.selectbox(
            "Bulan yang benar untuk data yang ketuker",
            months,
            index=default_index,
            format_func=lambda x: f"{x:02d} - {MONTH_NAMES_ID[x]}",
            help="Kalau contoh lo 11/4/2026 harusnya 11 April 2026, pilih 04 - April.",
        )
        candidates = swapped_date_candidates(df, int(correct_month))
        if candidates.empty:
            st.success("Tidak ada kandidat tanggal ketuker untuk bulan ini.")
        else:
            st.warning(f"Ditemukan {len(candidates)} kandidat tanggal yang kemungkinan ketuker. Cek preview dulu sebelum klik perbaiki.")
            preview = candidates.copy()
            preview["Nominal"] = preview["Nominal"].apply(compact_rp)
            display_df(preview[["id", "Tanggal Lama", "Tanggal Baru", "Sumber Dana", "Jenis", "Kategori", "Keterangan", "Nominal"]])
            with st.expander("Detail ISO yang akan diubah"):
                display_df(candidates[["id", "date_lama_iso", "date_baru_iso", "Keterangan"]])
            confirm = st.checkbox("Saya sudah cek preview dan ingin memperbaiki tanggal di atas")
            if st.button("Perbaiki tanggal ketuker", type="primary", disabled=not confirm):
                changed = apply_swapped_date_fix(candidates)
                st.success(f"Berhasil memperbaiki {changed} tanggal.")
                st.rerun()

    st.subheader("Catatan Penyimpanan")
    st.markdown(
        """
        - App ini pakai **SQLite** sebagai database backend.
        - Kalau jalan lokal/VPS/PythonAnywhere, data tersimpan di file `.db`.
        - Kalau deploy ke Streamlit Community Cloud, database lokal bisa cocok untuk demo, tapi untuk pemakaian jangka panjang tetap biasakan **Backup JSON/Excel** atau naik kelas ke Supabase/PostgreSQL.
        """
    )


def main() -> None:
    init_db()
    if not render_login():
        return

    df = fetch_transactions()
    budgets = fetch_budgets()

    st.sidebar.title("🏠 V.4 Padebuolo Next")
    page = st.sidebar.radio(
        "Menu",
        ["Dashboard", "Input Transaksi", "Buku Besar", "Pergerakan Belanja", "🤖 AI Assistant", "Kategorisasi Massal", "Budget Bulanan", "Import / Export", "Pengaturan"],
    )
    st.sidebar.divider()
    st.sidebar.caption(f"{APP_VERSION}")
    if st.sidebar.button("Logout"):
        st.session_state.pop("authenticated", None)
        st.rerun()

    if page == "Dashboard":
        page_dashboard(df, budgets)
    elif page == "Input Transaksi":
        page_input(df)
    elif page == "Buku Besar":
        page_ledger(df)
    elif page == "Pergerakan Belanja":
        page_monthly_category(df)
    elif page == "🤖 AI Assistant":
        page_ai_assistant(df, budgets)
    elif page == "Kategorisasi Massal":
        page_bulk_category(df)
    elif page == "Budget Bulanan":
        page_budget(df, budgets)
    elif page == "Import / Export":
        page_import_export(df, budgets)
    else:
        page_settings(df)


if __name__ == "__main__":
    main()
