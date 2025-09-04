# SSO Automation

Automasi login ke SSO SIASN, unduh data Monitoring Usulan via API, konversi ke Excel, dan upload ke Google Drive. Proses dijalankan otomatis setiap 15 menit.

## Fitur

- Login SSO otomatis (Selenium) dengan simpan/muat cookies dan localStorage (`src/sso_login.py:1`, `src/utils.py:1`).
- OTP/TOTP opsional via service lokal (`config/settings.json:1` kunci `totp_url`).
- Unduh API Monitoring Usulan memakai bearer token dari localStorage (`src/download_monitoring_usulan.py:1`).
- Konversi JSON → XLSX (openpyxl) (`src/download_monitoring_usulan.py:1`).
- Upload/replace ke Google Drive (PyDrive2) (`src/drive_upload.py:1`).
- Scheduler built-in tiap 15 menit di entrypoint (`src/main.py:1`).

## Prasyarat (Lokal)

- Python 3.11 + pip.
- Chrome/Chromium dan driver kompatibel di host, atau biarkan Selenium 4.12 mengelola otomatis. Alternatif: taruh `chromedriver/` di PATH (lihat `run.sh:1`).
- SSO credentials di `config/credentials.env:1`.
- Google Drive API client di `config/drive/drive_config.json:1` dan simpan token setelah autentikasi ke `config/drive/credentials.json:1`.

## Konfigurasi

- `config/settings.json:1`
  - `sso_url`, `redirect_url`, selector form (`username_field`, `password_field`, `login_button`), `timeout`, `totp_url`.
- `config/credentials.env:1`
  - `SSO_USERNAME`, `SSO_PASSWORD`.
- `src/main.py:1`
  - `target_folder_id` untuk folder tujuan di Google Drive.

## Menjalankan Secara Lokal

```bash
pip install -r requirements.txt
python src/main.py
```

Output:

- JSON: `data/downloads/monitoring_usulan.json:1`
- Excel: `data/downloads/monitoring_usulan.xlsx:1`
- Log: `data/logs/app.log:1`

## Scheduler 15 Menit

- Sudah tertanam di `src/main.py:1` (loop `time.sleep`).
- Menghentikan: Ctrl+C.
- Mengubah interval: edit `interval_minutes` di `src/main.py:1` (default 15). Jika ingin via env/setting, beri tahu saya untuk menambahkan `SCHEDULE_MINUTES` atau kunci di `settings.json`.

## Alternatif: Cron (opsional)

- Template kosong tersedia di `config/scheduler.cron:1`.
- Contoh host cron:

```cron
*/15 * * * * /usr/bin/python /path/to/repo/src/main.py >> /path/to/repo/data/logs/cron.log 2>&1
```

## Menjalankan Dengan Docker

- `Dockerfile:1` sudah memasang Chromium + Chromedriver (headless) dan dependency yang diperlukan.
- Build image:

```bash
docker build -t sso-automation .
```

- Jalankan container (mount config & data):

Windows PowerShell (detached/daemon, setiap 15 menit)

```powershell
docker run -d --name sso-automation --restart unless-stopped `
  -e SCHEDULE_MINUTES=15 `
  -v "${PWD}/config:/app/config" `
  -v "${PWD}/data:/app/data" `
  sso-automation
```

Linux/Mac (detached/daemon, setiap 15 menit)

```bash
docker run --memory="1.5g" --memory-swap="2g" -d --name sso-automation --restart unless-stopped \
  --add-host=host.docker.internal:host-gateway \
  -e SCHEDULE_MINUTES=15 \
  -v "$PWD/config:/app/config" \
  -v "$PWD/data:/app/data" \
  sso-automation
```

Lihat log container:

```bash
docker logs -f sso-automation
```

Hentikan & hapus container:

```bash
docker stop sso-automation && docker rm sso-automation
```

- Catatan OAuth Drive: lakukan autentikasi PyDrive2 sekali di host agar `config/drive/credentials.json:1` terisi, lalu mount ke container. Autentikasi berbasis browser di dalam container tidak praktis.

### Akses TOTP dari Dalam Container

- Jangan gunakan `0.0.0.0` sebagai URL target; itu hanya alamat bind, bukan alamat yang bisa diakses.
- Jika service TOTP berjalan di host pada port `8001`, set `totp_url` ke `http://host.docker.internal:8001/totp` (sudah default di `config/settings.json:1`).
- Alternatif override cepat tanpa mengubah file: tambahkan env saat `docker run`:

  - PowerShell:

    ```powershell
    docker run -d --name sso-automation --restart unless-stopped `
      -e TOTP_URL=http://host.docker.internal:8001/totp `
      -v "${PWD}/config:/app/config" -v "${PWD}/data:/app/data" `
      sso-automation
    ```

  - Linux: jika `host.docker.internal` tidak resolve, tambahkan flag host-gateway:

    ```bash
    docker run -d --name sso-automation --restart unless-stopped \
      --add-host=host.docker.internal:host-gateway \
      -e TOTP_URL=http://host.docker.internal:8001/totp \
      -v "$PWD/config:/app/config" -v "$PWD/data:/app/data" \
      sso-automation
    ```

  - Jika TOTP dijalankan sebagai container lain di jaringan yang sama, gunakan nama servicenya, mis. `http://totp-service:8001/totp` dan pastikan berada pada network Docker yang sama.

## Struktur Proyek

- `src/main.py:1` — entry point + scheduler 15 menit.
- `src/browser.py:1` — setup Selenium/Chrome.
- `src/sso_login.py:1` — alur login SSO + TOTP opsional.
- `src/utils.py:1` — simpan/muat cookies & localStorage.
- `src/download_monitoring_usulan.py:1` — unduh API + konversi JSON → Excel.
- `src/drive_upload.py:1` — upload/replace file ke Google Drive.
- `scripts/convert_now.py:1` — konversi JSON yang ada menjadi XLSX dan upload (tanpa login/unduh).
- `config/*.json|*.env` — konfigurasi aplikasi.
- `Dockerfile:1`, `run.sh:1` — container dan skrip jalan lokal.

## Manual/Once-off Tasks

- Jalankan satu siklus lalu hentikan: jalankan aplikasi; ketika log "Job selesai." muncul, hentikan proses.
- Hanya konversi & upload dari JSON yang sudah ada: `python scripts/convert_now.py:1`.

## Troubleshooting

- Login/OTP gagal: pastikan `totp_url` aktif dan mengembalikan JSON dengan key `totp`; cek `data/logs/error.png:1`.
- Driver/Chrome mismatch: gunakan Selenium Manager atau sesuaikan versi Chromedriver.
- Upload Drive gagal: pastikan `config/drive/drive_config.json:1` valid dan `config/drive/credentials.json:1` sudah terisi.
