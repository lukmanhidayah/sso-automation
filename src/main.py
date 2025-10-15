import json
import os
import time
from datetime import datetime
from browser import setup_driver
from sso_login import login_sso
from download_monitoring_usulan import (
    download_monitoring_usulan_paginated,
    convert_monitoring_json_to_excel,
    download_pertek_documents_from_json,
    download_sk_documents_from_json,
)
from drive_upload import upload_file_to_drive


def run_once():
    # Hapus file sesi lama saat mulai
    for path in (
        "data/sso_cookies.json",
        "data/sso_localstorage.json",
    ):
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            # Abaikan error penghapusan agar tidak mengganggu startup
            pass

    with open("config/settings.json") as f:
        config = json.load(f)

    driver = setup_driver(headless=True)
    try:
        if login_sso(driver, config):
            print("Login SSO successful. Current URL:", driver.current_url)
            json_out = "data/downloads/monitoring_usulan.json"
            xlsx_out = "data/downloads/monitoring_usulan.xlsx"
            # Folder Drive untuk Excel (Sheets)
            excel_folder_id = "15_2IHRXVeajrzO0oYaJ_-ARnkDJsW7YY"
            # Folder Drive untuk dokumen PDF Pertek
            pdf_folder_id = "15e0vW-4SJjCjBP8ksIFc1Pw1oUZgJX1F"
            download_monitoring_usulan_paginated(out_path=json_out)
            # Convert JSON to Excel with selected fields
            convert_monitoring_json_to_excel(
                json_path=json_out,
                excel_path=xlsx_out,
                pertek_drive_folder_id=pdf_folder_id,
            )

            # Download SK documents to separate folder
            try:
                sk_folder_id = "1YCHZI7-x2aDZI-K4W_bFhns4IbrEn0WC"
                download_sk_documents_from_json(
                    json_path=json_out,
                    out_dir="data/downloads/monitoring_usulan_ttd_sk",
                    excel_path=xlsx_out,
                    sk_drive_folder_id=sk_folder_id,
                )
            except Exception as e:
                print(f"Gagal download SK: {e}")

            # Download Pertek documents after conversion
            try:
                download_pertek_documents_from_json(
                    json_path=json_out,
                    out_dir="data/downloads/monitoring_usulan_ttd_pertek",
                    excel_path=xlsx_out,
                    pertek_drive_folder_id=pdf_folder_id,
                )
            except Exception as e:
                print(f"Gagal download Pertek: {e}")

            # Upload to Google Drive after conversion
            try:
                upload_file_to_drive(
                    xlsx_out,
                    excel_folder_id,
                    convert_spreadsheet=True,
                    replace_by_title=True,
                    custom_title="monitoring_usulan",
                )
            except Exception as e:
                print(f"Gagal upload ke Google Drive: {e}")
            # Lakukan aksi lain, misalnya navigasi ke dashboard
        else:
            print("Login SSO failed.")
    finally:
        driver.quit()


if __name__ == "__main__":
    # Izinkan override interval via environment variable SCHEDULE_MINUTES
    try:
        interval_minutes = int(os.getenv("SCHEDULE_MINUTES", "15"))
    except ValueError:
        interval_minutes = 15
    print(f"Scheduler aktif: menjalankan job setiap {interval_minutes} menit.")
    while True:
        start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{start_time}] Menjalankan job...")
        try:
            run_once()
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Job selesai.")
        except KeyboardInterrupt:
            print("Dihentikan oleh pengguna.")
            break
        except Exception as e:
            print(f"Terjadi error saat menjalankan job: {e}")
        # Tunggu hingga siklus berikutnya
        try:
            for _ in range(interval_minutes * 60):
                time.sleep(1)
        except KeyboardInterrupt:
            print("Dihentikan oleh pengguna saat menunggu jadwal berikutnya.")
            break
