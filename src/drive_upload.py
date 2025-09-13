import os
from typing import Optional

from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive


def _build_gauth(project_root: str) -> GoogleAuth:
    gauth = GoogleAuth()
    gauth.settings = {
        "client_config_backend": "file",
        "client_config_file": os.path.join(project_root, "config", "drive", "drive_config.json"),
        "save_credentials": True,
        "save_credentials_backend": "file",
        "save_credentials_file": os.path.join(project_root, "config", "drive", "credentials.json"),
        "oauth_scope": ["https://www.googleapis.com/auth/drive"],
    }

    credentials_path = os.path.join(project_root, "config", "drive", "credentials.json")
    gauth.LoadCredentialsFile(credentials_path)
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    gauth.SaveCredentialsFile(credentials_path)
    return gauth


def upload_file_to_drive(
    file_path: str,
    target_folder_id: str,
    *,
    convert_spreadsheet: bool = True,
    replace_by_title: bool = True,
    custom_title: Optional[str] = None,
) -> str:
    """Upload a single file to Google Drive.

    - If ``replace_by_title`` is True, deletes any existing file in the folder with the same title.
    - If ``convert_spreadsheet`` and the file is Excel/CSV, upload as Google Sheets.
    - Uses config at ``config/drive/`` for OAuth client and stored credentials.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")

    # Resolve project root (one level up from this file's folder)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)

    gauth = _build_gauth(project_root)
    drive = GoogleDrive(gauth)

    fname = os.path.basename(file_path)
    title = custom_title or (
        os.path.splitext(fname)[0]
        if convert_spreadsheet and fname.lower().endswith((".xlsx", ".xls", ".csv"))
        else fname
    )

    # Look for an existing file with the same title in the target folder
    file_obj = None
    if replace_by_title:
        try:
            # Build Drive query safely without complex f-string escaping
            safe_title = title.replace("'", "\\'")
            q = "title = '{}' and '{}' in parents and trashed=false".format(safe_title, target_folder_id)
            existing = drive.ListFile({"q": q}).GetList()
            if existing:
                file_obj = existing[0]
                print(f"Replace konten file lama: {file_obj.get('title')} ({file_obj.get('id')})")
            else:
                print(f"File lama tidak ditemukan, upload baru: {title}")
        except Exception as list_err:
            print(f"Peringatan: gagal memeriksa file yang ada: {list_err}")

    if not file_obj:
        metadata = {"title": title, "parents": [{"id": target_folder_id}]}
        file_obj = drive.CreateFile(metadata)

    file_obj.SetContentFile(file_path)
    if convert_spreadsheet and fname.lower().endswith((".xlsx", ".xls", ".csv")):
        file_obj.Upload({"convert": True})
    else:
        file_obj.Upload()

    # Fetch metadata to build a share/view link
    try:
        file_obj.FetchMetadata(fields='id,alternateLink,webViewLink')
    except Exception:
        pass
    file_id = file_obj.get('id')
    web_link = file_obj.get('webViewLink') or file_obj.get('alternateLink')
    if not web_link and file_id:
        web_link = f"https://drive.google.com/file/d/{file_id}/view"

    print(f"Berhasil upload/replace: {title} (ID: {file_id})")
    return web_link or ""
