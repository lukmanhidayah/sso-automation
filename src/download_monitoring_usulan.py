import json
import os
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from typing import Any, Dict, List

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None  # type: ignore


API_URL = (
    "https://api-siasn.bkn.go.id/siasn-instansi/pengadaan/usulan/monitoring"
    "?no_peserta=&nama=&tgl_usulan=2025-08-22&jenis_pengadaan_id=02&jenis_formasi_id=&status_usulan=&periode=&limit=10000&offset=0"
)

# Mapping of status_usulan ID to descriptive name
STATUS_USULAN_MAP = {
    "1": "Input Berkas",
    "2": "Berkas Disimpan (Terverifikasi)",
    "3": "Surat Usulan",
    "4": "Approval Surat Usulan",
    "5": "Perbaikan Dokumen",
    "6": "Tidak Memenuhi Syarat",
    "7": "Menunggu Cetak SK – Menyetujui",
    "8": "Menunggu Cetak SK – Perbaikan Pertek",
    "9": "Menunggu Cetak SK – Pembatalan Pertek",
    "10": "Cetak SK",
    "11": "Profil PNS telah diperbaharui",
    "12": "Terima Usulan",
    "13": "Validasi Usulan – Tidak Memenuhi Syarat",
    "14": "Validasi Usulan – Perbaikan Dokumen",
    "15": "Validasi Usulan – Disetujui",
    "16": "Berkas Disetujui",
    "17": "Menunggu Paraf – Paraf Pertek",
    "18": "Menunggu Paraf – Gagal Paraf Pertek",
    "19": "Sdh di paraf - Pertek",
    "20": "Menunggu Tanda tangan- TTD Pertek",
    "21": "Berkas Ditolak - TTD Pertek",
    "22": "Sdh di TTD - Pertek",
    "23": "Surat Keluar",
    "24": "Perbaikan Pertek (Menunggu Approval Instansi)",
    "25": "Terima Usulan Penetapan – Pembatalan",
    "26": "Pembatalan Pertek (Menunggu Approval Instansi)",
    "27": "Menunggu SK – Paraf / TTE",
    "28": "Setuju Paraf SK",
    "29": "Tolak TTD SK",
    "30": "Setuju TTD SK",
    "31": "Telah Update di Profile PNS",
    "32": "Pembuatan SK Berhasil",
    "33": "Menunggu Layanan",
    "34": "Perbaikan Dokumen - Menunggu Approval",
    "35": "Tolak Paraf SK",
    "36": "Menunggu TTD - SK",
    "37": "Approval Perbaikan Pertek",
    "38": "Approval Pembatalan Pertek",
    "39": "Perbaikan SK",
    "40": "Berkas Disimpan (Terverifikasi) - Perbaikan SK",
    "41": "Validasi Usulan - Perbaikan SK",
    "42": "Validasi Usulan - Perbaikan SK (Disetujui)",
    "43": "Menunggu Paraf - Perbaikan SK",
    "44": "Menunggu TTD - Perbaikan SK",
    "45": "Sudah TTD - Perbaikan SK",
    "46": "Menunggu TTD SK - Instansi",
    "47": "Tolak TTD SK - Instansi",
    "48": "Setuju TTD SK - Instansi",
    "49": "Sudah TTD - SK",
    "50": "Perbaikan Dokumen - MYSAPK",
    "51": "Input Berkas - Perbaikan MySAPK ",
    "52": "Perbaikan Dokumen - Approval",
    "53": "Setuju TTD Pertek",
    "55": "Approval Tingkat Provinsi",
    "56": "Perbaikan Approval",
    "57": "Perbaikan Pertek",
    "58": "Validasi Usulan - Perbaikan Pertek",
    "59": "Menunggu Buat Sk",
    "60": "Proses Persidangan",
    "61": "Input Berkas - SK PNS",
    "62": "Menunggu TTD SK PNS - Instansi",
    "63": "Setuju TTD Digital SK PNS",
    "64": "Pembuatan SK Basah PNS Berhasil",
    "65": "Pembatalan NIP/Pertek",
    "66": "Perbaikan SK Provinsi",
    "67": "Validasi Perbaikan Dokumen - BTS",
    "68": "Validasi Perbaikan SK - BTS",
    "69": "Perbaikan Dokumen SK - BTS",
    "70": "Tolak Perbaikan SK - TMS",
    "71": "Validasi Perbaikan Dokumen SK - BTS",
    "72": "Perbaikan Dokumen Pertek - BTS",
    "73": "Tolak Perbaikan Pertek - TMS",
    "74": "Validasi Perbaikan Pertek - BTS",
    "75": "Pembatalan PERTEK",
    "76": "Menunggu Rekom OTDA",
    "99": "Usulan Dihapus",
}


def load_sso_token(path: str) -> str:
    if not os.path.exists(path):
        raise FileNotFoundError(f"LocalStorage JSON not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    token = data.get("sso_token")
    if not token:
        raise ValueError("Key 'sso_token' not found in localStorage JSON")
    return token


def download_monitoring_usulan(
    out_path: str, localstorage_path: str = "data/sso_localstorage.json"
) -> None:
    token = load_sso_token(localstorage_path)

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
        "Authorization": f"Bearer {token}",
        "Connection": "keep-alive",
        "Origin": "https://siasn-instansi.bkn.go.id",
        "Referer": "https://siasn-instansi.bkn.go.id/layananPengadaan/monitoringUsulan",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0"
        ),
        "sec-ch-ua": '"Not;A=Brand";v="99", "Microsoft Edge";v="139", "Chromium";v="139"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        # Cookie header is optional; include if needed
        # "Cookie": "d732bf4cc14bde954e7b20b82e794606=730362334080c57616252771951a529f",
    }

    req = Request(API_URL, headers=headers, method="GET")
    try:
        with urlopen(req) as resp:
            status = resp.getcode()
            body = resp.read()
            if status != 200:
                raise HTTPError(API_URL, status, f"HTTP {status}", resp.headers, None)
    except HTTPError as e:
        # Try to read error body for debugging
        detail = None
        try:
            detail = e.read().decode("utf-8", errors="ignore")  # type: ignore[attr-defined]
        except Exception:
            pass
        msg = f"Request failed: {e.code} {e.reason}"
        if detail:
            msg += f"\n{detail}"
        raise RuntimeError(msg)
    except URLError as e:
        raise RuntimeError(f"Network error: {e.reason}")

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(body)


def _extract_rows(payload: Dict[str, Any]) -> List[List[Any]]:
    items = payload.get("data") or []
    rows: List[List[Any]] = []
    for it in items:
        # Nested fields
        nested = (it or {}).get("usulan_data") or {}
        nested_data = nested.get("data") or {}

        no_peserta = nested_data.get("no_peserta") or ""
        nama = (it or {}).get("nama") or nested_data.get("nama") or ""
        jenis_pengadaan = "PPPK"  # Hardcoded value
        jenis_formasi_nama = (it or {}).get("jenis_formasi_nama") or ""
        tgl_usulan = (it or {}).get("tgl_usulan") or ""
        tgl_pengiriman_kelayanan = (it or {}).get("tgl_pengiriman_kelayanan") or ""
        status_usulan = (it or {}).get("status_usulan") or ""

        rows.append(
            [
                no_peserta,
                nama,
                jenis_pengadaan,
                jenis_formasi_nama,
                tgl_usulan,
                tgl_pengiriman_kelayanan,
                status_usulan,
            ]
        )
    return rows


def convert_monitoring_json_to_excel(json_path: str, excel_path: str) -> None:
    """Convert downloaded monitoring_usulan JSON into an Excel file.

    Columns: no_peserta, nama, jenis_pengadaan, jenis_formasi_nama, tgl_usulan, tgl_pengiriman_kelayanan, status_usulan
    """
    if Workbook is None:
        raise RuntimeError(
            "openpyxl not available. Please install it (e.g., pip install openpyxl)."
        )

    if not os.path.exists(json_path):
        raise FileNotFoundError(f"JSON file not found: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    rows = _extract_rows(payload if isinstance(payload, dict) else {})

    # Convert status_usulan ID to descriptive name
    for r in rows:
        status_id = str(r[6])
        r[6] = STATUS_USULAN_MAP.get(status_id, status_id)

    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "monitoring_usulan"
    # Header
    ws.append(
        [
            "No. Peserta",
            "Nama",
            "Jenis Pengadaan",
            "Jenis Formasi",
            "Tanggal Buat Usulan",
            "Tanggal Pengiriman",
            "Status Usulan",
        ]
    )
    # Set minimum column width to 50
    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(
            ws.column_dimensions[col_letter].width or 0, 50
        )
    for r in rows:
        ws.append(r)
        # Set minimum column width to 50 for all columns
        for col_idx in range(1, ws.max_column + 1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            ws.column_dimensions[col_letter].width = 50
    wb.save(excel_path)
