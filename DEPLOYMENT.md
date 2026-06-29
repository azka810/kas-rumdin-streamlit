# Deployment Guide — Kas Rumah Dinas Streamlit

## Opsi 1 — Streamlit Community Cloud

Cocok untuk demo atau penggunaan ringan.

1. Upload semua isi folder ini ke repository GitHub.
2. Pastikan `app.py` dan `requirements.txt` ada di root repository.
3. Buka Streamlit Community Cloud.
4. Klik **Create app**.
5. Pilih repository, branch, dan entrypoint file `app.py`.
6. Di bagian secrets, isi:

```toml
APP_PASSWORD = "password-yang-kuat"
```

7. Deploy.

Catatan: SQLite di Community Cloud cocok untuk demo. Untuk data yang sangat penting dan harus permanen, tetap backup rutin atau gunakan database eksternal seperti Supabase/PostgreSQL.

## Opsi 2 — VPS / Server sendiri

```bash
pip install -r requirements.txt
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

Agar jalan terus, gunakan `tmux`, `screen`, `systemd`, atau reverse proxy Nginx.

## Opsi 3 — Render/Railway

Bisa juga, tapi perlu command start:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port $PORT
```

Kalau pakai SQLite, pastikan pakai persistent disk/storage agar database tidak hilang saat redeploy.
