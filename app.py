from __future__ import annotations

import io
import json
import os
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd
import streamlit as st

APP_TITLE = "Kas Rumah Dinas"
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
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if value is None or str(value).strip() == "":
        return date.today().isoformat()
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return date.today().isoformat()
    return parsed.date().isoformat()


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
    tmp["month"] = pd.to_datetime(tmp["date"], errors="coerce").dt.to_period("M").astype(str)
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


def display_df(df: pd.DataFrame, use_container_width: bool = True) -> None:
    st.dataframe(df, use_container_width=use_container_width, hide_index=True)


def ledger_view(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    out = df.copy().sort_values(["date", "id"]).copy()
    out["Tanggal"] = pd.to_datetime(out["date"], errors="coerce").dt.strftime("%d/%m/%Y")
    out["Sumber Dana"] = out["fund"]
    out["Kategori"] = out["category"]
    out["Keterangan"] = out["description"]
    out["Masuk"] = out.apply(lambda r: compact_rp(r["amount"]) if r["type"] == "Masuk" else "-", axis=1)
    out["Keluar"] = out.apply(lambda r: compact_rp(r["amount"]) if r["type"] == "Keluar" else "-", axis=1)
    out["Netto"] = out["netto"].apply(compact_rp)
    out["Sisa Kas"] = out["running_balance"].apply(compact_rp)
    out["Metode"] = out["method"].fillna("Kas")
    out["Catatan"] = out["note"].fillna("")
    return out[["id", "Tanggal", "Sumber Dana", "Kategori", "Keterangan", "Masuk", "Keluar", "Netto", "Sisa Kas", "Metode", "Catatan"]]


def detail_ledger_view(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    out = df.copy().sort_values(["date", "id"]).copy()
    out["Tanggal"] = pd.to_datetime(out["date"], errors="coerce").dt.strftime("%d/%m/%Y")
    out["Masuk Detail"] = out.apply(lambda r: full_rp(r["amount"]) if r["type"] == "Masuk" else "-", axis=1)
    out["Keluar Detail"] = out.apply(lambda r: full_rp(r["amount"]) if r["type"] == "Keluar" else "-", axis=1)
    out["Netto Detail"] = out["netto"].apply(full_rp)
    out["Sisa Kas Detail"] = out["running_balance"].apply(full_rp)
    return out[["id", "Tanggal", "fund", "category", "description", "Masuk Detail", "Keluar Detail", "Netto Detail", "Sisa Kas Detail", "method", "note"]].rename(
        columns={"fund": "Sumber Dana", "category": "Kategori", "description": "Keterangan", "method": "Metode", "note": "Catatan"}
    )


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
        raw = tx.drop(columns=["netto", "running_balance"], errors="ignore")
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
    st.title("🏠 Kas Rumah Dinas")
    st.caption("Masuk dulu pak. Default password: rumdin123, bisa diganti di secrets/env APP_PASSWORD.")
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
    st.title("🏠 Dashboard Kas Rumah Dinas")
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
    tx_date = c2.date_input("Tanggal", value=date.today())
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

    st.caption("Kolom Sisa Kas = saldo berjalan per sumber dana setelah transaksi tersebut.")
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
        tx_date = c1.date_input("Tanggal", value=pd.to_datetime(row["date"]).date(), key="edit_date")
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


def page_budget(df: pd.DataFrame, budgets: pd.DataFrame) -> None:
    st.title("🧾 Budget Bulanan")
    c1, c2, c3 = st.columns(3)
    total_budget = int(budgets["amount"].sum()) if not budgets.empty else 0
    c1.metric("Total Budget Bulanan", compact_rp(total_budget), help=full_rp(total_budget))
    saldo = summarize(df)["saldo"]
    c2.metric("Saldo Kas", compact_rp(saldo), help=full_rp(saldo))
    c3.metric("Estimasi Bertahan", f"{saldo / total_budget:.1f} bulan" if total_budget > 0 else "-")

    st.subheader("Daftar Budget")
    if budgets.empty:
        st.info("Belum ada budget.")
    else:
        show = budgets.copy()
        show["Nominal"] = show["amount"].apply(compact_rp)
        display_df(show[["id", "person", "component", "Nominal"]].rename(columns={"person": "Sumber Dana", "component": "Komponen"}))
        with st.expander("Detail nominal budget"):
            d = budgets.copy()
            d["Nominal Detail"] = d["amount"].apply(full_rp)
            display_df(d[["id", "person", "component", "Nominal Detail"]].rename(columns={"person": "Sumber Dana", "component": "Komponen"}))

    st.divider()
    st.subheader("Tambah Budget")
    funds = get_fund_list(df)
    with st.form("budget_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([1, 1.5, 1])
        person = c1.selectbox("Sumber dana/orang", funds)
        component = c2.text_input("Komponen", placeholder="Sewa, listrik, internet, air, dll")
        amount = c3.number_input("Nominal", min_value=0, step=10_000, format="%d")
        submitted = st.form_submit_button("Tambah Budget", type="primary")
    if submitted:
        if not component.strip() or amount <= 0:
            st.error("Komponen dan nominal wajib diisi.")
        else:
            add_budget(person, component.strip(), int(amount))
            st.success("Budget ditambahkan.")
            st.rerun()

    if not budgets.empty:
        st.subheader("Hapus Budget")
        selected = st.selectbox("Pilih budget", budgets["id"].astype(int).tolist(), format_func=lambda x: f"ID {x}")
        if st.button("Hapus Budget Terpilih"):
            delete_budget(int(selected))
            st.warning("Budget dihapus.")
            st.rerun()


def page_import_export(df: pd.DataFrame, budgets: pd.DataFrame) -> None:
    st.title("📦 Import / Export")
    st.subheader("Export")
    c1, c2, c3, c4 = st.columns(4)
    c1.download_button("Backup JSON", data=backup_json_bytes(), file_name="backup_kas_rumdin.json", mime="application/json", use_container_width=True)
    c2.download_button("Buku Besar CSV", data=to_csv_bytes(detail_ledger_view(df)), file_name="buku_besar_kas_rumdin.csv", mime="text/csv", use_container_width=True)
    c3.download_button("Export Excel", data=export_excel_bytes(), file_name="rekap_kas_rumdin.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if DB_PATH.exists():
        c4.download_button("Download DB", data=DB_PATH.read_bytes(), file_name="kas_rumdin.db", mime="application/octet-stream", use_container_width=True)

    st.divider()
    st.subheader("Import")
    st.caption("Bisa import backup JSON dari app ini, CSV transaksi, atau Excel lama dengan sheet Master Kas/Biaya Bulanan.")
    uploaded = st.file_uploader("Upload file", type=["json", "csv", "xlsx", "xls"])
    replace = st.toggle("Replace semua data saat import", value=False, help="Kalau aktif, transaksi dan budget lama akan dihapus dulu.")
    if uploaded is not None:
        tx_rows: List[Dict[str, Any]] = []
        budget_rows: List[Dict[str, Any]] = []
        try:
            if uploaded.name.lower().endswith(".json"):
                payload = json.loads(uploaded.getvalue().decode("utf-8"))
                tx_rows = payload.get("transactions", [])
                budget_rows = payload.get("budgets", [])
            elif uploaded.name.lower().endswith(".csv"):
                imported_df = pd.read_csv(uploaded)
                tx_rows = normalize_transactions_from_df(imported_df)
            else:
                xl = pd.ExcelFile(uploaded)
                tx_sheet = None
                budget_sheet = None
                for sheet in xl.sheet_names:
                    low = sheet.lower()
                    if tx_sheet is None and ("trans" in low or "master" in low or "buku" in low):
                        tx_sheet = sheet
                    if budget_sheet is None and ("budget" in low or "biaya" in low):
                        budget_sheet = sheet
                if tx_sheet:
                    tx_rows = normalize_transactions_from_df(pd.read_excel(xl, sheet_name=tx_sheet))
                if budget_sheet:
                    budget_rows = normalize_budgets_from_df(pd.read_excel(xl, sheet_name=budget_sheet))
            st.info(f"Terdeteksi {len(tx_rows)} transaksi dan {len(budget_rows)} budget.")
            with st.expander("Preview transaksi import"):
                display_df(pd.DataFrame(tx_rows).head(20))
            if st.button("Proses Import", type="primary"):
                tx_count, bd_count = import_rows(tx_rows, budget_rows, replace=replace)
                st.success(f"Import selesai: {tx_count} transaksi dan {bd_count} budget masuk.")
                st.rerun()
        except Exception as exc:
            st.error(f"Gagal membaca file: {exc}")

    st.divider()
    st.subheader("Reset Data")
    st.warning("Reset akan menghapus semua transaksi dan budget, lalu mengisi ulang data contoh dari Excel awal.")
    confirm = st.text_input("Ketik RESET untuk konfirmasi")
    if st.button("Reset ke Data Awal"):
        if confirm == "RESET":
            conn = get_conn()
            conn.execute("DELETE FROM transactions")
            conn.execute("DELETE FROM budgets")
            conn.commit()
            seed_if_empty()
            st.success("Data sudah direset ke data awal.")
            st.rerun()
        else:
            st.error("Konfirmasi belum sesuai.")



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
        show["Tanggal"] = pd.to_datetime(show["date"], errors="coerce").dt.strftime("%d/%m/%Y")
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

    st.sidebar.title("🏠 Kas Rumdin")
    page = st.sidebar.radio(
        "Menu",
        ["Dashboard", "Input Transaksi", "Buku Besar", "Kategorisasi Massal", "Budget Bulanan", "Import / Export", "Pengaturan"],
    )
    st.sidebar.divider()
    st.sidebar.caption("Default password: rumdin123")
    if st.sidebar.button("Logout"):
        st.session_state.pop("authenticated", None)
        st.rerun()

    if page == "Dashboard":
        page_dashboard(df, budgets)
    elif page == "Input Transaksi":
        page_input(df)
    elif page == "Buku Besar":
        page_ledger(df)
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
