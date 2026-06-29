# Kas Rumah Dinas — Streamlit App

Aplikasi rekap kas rumah dinas berbasis Streamlit + SQLite.

## Fitur

- Dashboard saldo kas, total masuk, total keluar, dan estimasi kas bertahan.
- Input transaksi biasa.
- Mode cepat inject dana/top up.
- Mode cepat split pengeluaran Rayhan/Azka atau sumber dana lain.
- Buku besar dengan kolom Masuk, Keluar, Netto, dan Sisa Kas berjalan.
- Angka tampil compact, detail nominal tersedia di expander/download.
- Budget bulanan.
- Import dari JSON/CSV/Excel lama.
- Export JSON, CSV, Excel, dan database SQLite.
- Login sederhana dengan password.

## Cara jalan lokal

### Windows

1. Extract ZIP.
2. Buka folder `kas_rumdin_streamlit`.
3. Double click `run_app.bat`.
4. Buka URL yang muncul, biasanya `http://localhost:8501`.

### Manual

```bash
cd kas_rumdin_streamlit
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # Mac/Linux
pip install -r requirements.txt
streamlit run app.py
```

Default password:

```text
rumdin123
```

## Ganti password

### Lokal

Bisa pakai environment variable:

```bash
set APP_PASSWORD=password-baru
streamlit run app.py
```

Atau copy `.streamlit/secrets.toml.example` menjadi `.streamlit/secrets.toml`, lalu isi:

```toml
APP_PASSWORD = "password-baru"
```

### Streamlit Community Cloud

Isi secrets saat deploy:

```toml
APP_PASSWORD = "password-baru"
```

## Struktur folder

```text
kas_rumdin_streamlit/
  app.py
  requirements.txt
  run_app.bat
  run_app.sh
  data/
    seed_transactions.csv
    seed_budgets.csv
    Rekap Kas Rumdin.xlsx
  .streamlit/
    config.toml
    secrets.toml.example
```

## Catatan database

Database otomatis dibuat di:

```text
instance/kas_rumdin.db
```

Kalau app pertama kali dibuka, database akan diisi dari `data/seed_transactions.csv` dan `data/seed_budgets.csv`.

Untuk pemakaian serius, rajin klik **Backup JSON** atau **Export Excel** di menu Import / Export.
