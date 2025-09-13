import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.download_monitoring_usulan import convert_monitoring_json_to_excel
from src.drive_upload import upload_file_to_drive


if __name__ == "__main__":
    convert_monitoring_json_to_excel(
        "data/downloads/monitoring_usulan.json",
        "data/downloads/monitoring_usulan.xlsx",
        pertek_drive_folder_id="15e0vW-4SJjCjBP8ksIFc1Pw1oUZgJX1F",
    )
    print("Converted to data/downloads/monitoring_usulan.xlsx")
    try:
        upload_file_to_drive(
            "data/downloads/monitoring_usulan.xlsx",
            "15_2IHRXVeajrzO0oYaJ_-ARnkDJsW7YY",
            convert_spreadsheet=True,
            replace_by_title=True,
            custom_title="monitoring_usulan",
        )
    except Exception as e:
        print(f"Gagal upload ke Google Drive: {e}")
