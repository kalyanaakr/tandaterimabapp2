"""
Form Perpindahan Dokumen BAPP — VERSI SATU FILE, konsep seperti Google Form.

Alur: Form (pengirim/penerima/dari/ke + pilih bundle) -> Summary (+ tanda
tangan digital pengirim & penerima) -> Submit -> Berhasil (bisa download PDF
ringkasan yang sudah ada tanda tangannya). Tidak ada sidebar/menu — cuma
satu alur form dari atas ke bawah. Sinkronisasi data ada di ikon pengaturan
(⚙️) di pojok kanan atas.

Cara pakai singkat:
- Lokal: service_account.json + .env
- Streamlit Cloud: Google service account disimpan di Streamlit Secrets
  pada bagian [gcp_service_account], bukan di GitHub.
- Isi DB_SHEET_ID, MASTER_SHEET_ID, dan bila perlu DRIVE_ROOT_FOLDER_ID.
- Jalankan: streamlit run app.py
"""
import io
import os
import tempfile
from datetime import datetime, date

import numpy as np
import streamlit as st
import pandas as pd
import gspread
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from streamlit_drawable_canvas import st_canvas
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv

# Catatan deployment:
# requirements.txt wajib memuat streamlit-drawable-canvas, reportlab, Pillow,
# gspread, google-auth, google-api-python-client, pandas, numpy, dan python-dotenv.

# =============================================================================
# 1. KONFIGURASI (lokal dari .env, Cloud dari Streamlit Secrets)
# =============================================================================
load_dotenv()

# Streamlit Cloud memakai st.secrets, sedangkan lokal memakai .env.
# Nilai dari Secrets diprioritaskan jika tersedia.
def _setting(name: str, default: str = "") -> str:
    value = os.getenv(name, default)
    try:
        secret_value = st.secrets.get(name, None)
        if secret_value is not None:
            # Jangan menganggap string kosong sebagai konfigurasi aktif.
            secret_value = str(secret_value).strip()
            if secret_value:
                value = secret_value
    except Exception:
        pass
    return str(value or "").strip()


DB_SHEET_ID = _setting("DB_SHEET_ID")
MASTER_SHEET_ID = _setting("MASTER_SHEET_ID")
MASTER_SHEET_WORKSHEET_NAME = _setting("MASTER_SHEET_WORKSHEET_NAME", "data")
GOOGLE_SERVICE_ACCOUNT_FILE = _setting("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
DRIVE_ROOT_FOLDER_ID = _setting("DRIVE_ROOT_FOLDER_ID")
ADMIN_PASSWORD = _setting("ADMIN_PASSWORD")
ADMIN_EMAIL = _setting("ADMIN_EMAIL")

APP_TITLE = "Form Perpindahan Dokumen BAPP"

# Nama tab yang dibuat OTOMATIS di spreadsheet DB_SHEET_ID (boleh sama dengan
# MASTER_SHEET_ID kalau digabung jadi satu spreadsheet).
TAB_MASTER_BUNDLE = "master_bundle"
TAB_TRANSAKSI = "transaksi_perpindahan"
TAB_DETAIL = "detail_perpindahan"
TAB_RIWAYAT = "riwayat_lokasi_bundle"

TAB_HEADERS = {
    TAB_MASTER_BUNDLE: ["nomor_bundle", "direktorat", "jumlah_bapp", "lokasi_saat_ini", "waktu", "status"],
    TAB_TRANSAKSI: [
        "id_transaksi", "tanggal", "waktu", "nama_pengirim", "nama_penerima",
        "dari_lokasi", "ke_lokasi", "jumlah_bundle", "total_bapp", "status",
        "drive_folder_id", "drive_folder_url", "created_at",
    ],
    TAB_DETAIL: ["id_transaksi", "nomor_bundle", "direktorat", "jumlah_bapp"],
    TAB_RIWAYAT: ["nomor_bundle", "dari_lokasi", "ke_lokasi", "waktu", "id_transaksi"],
}

# PENTING: kalau tab transaksi_perpindahan / detail_perpindahan versi LAMA
# (sebelum ada nama_pengirim/nama_penerima/direktorat) masih ada di
# spreadsheet Anda, HAPUS dulu tab itu secara manual di Google Sheets
# (klik kanan tab -> Delete) sebelum pakai versi ini, supaya tab baru dibuat
# otomatis dengan header yang benar. Tab master_bundle & riwayat_lokasi_bundle
# TIDAK perlu dihapus, strukturnya tidak berubah.

# Pemetaan kolom Master BAPP: nama internal -> judul kolom PERSIS di spreadsheet.
COLUMN_MAP = {
    "status_bapp_fisik": "Status BAPP Fisik",
    "waktu_bapp_diterima": "Waktu BAPP diterima",
    "nomor_penerimaan": "Nomor Penerimaan",
    "nomor_urut_penerimaan": "Nomor Urut Penerimaan",
    "barcode_penerimaan": "Barcode Penerimaan",
    "status_bapp_fisik_kedua": "Status BAPP Fisik Kedua",
    "waktu_bapp_diterima_baru": "Waktu BAPP diterima Baru",
    "nomor_penerimaan_baru": "Nomor Penerimaan Baru",
    "nomor_transaksi": "Nomor Transaksi",
    "serial_number": "Serial Number",
    "npsn": "NPSN",
    "nama_sekolah": "Nama Sekolah",
    "termin": "Termin",
    "provinsi": "Provinsi",
    "kabupaten_kota": "Kabupaten/Kota",
    "direktorat": "Direktorat",
    "status": "Status",
    "tanggal_bapp": "Tanggal BAPP",
    "nomor_map": "Nomor MAP",
    "nama_koordinator": "Nama Koordinator",
    "nama_pengirim": "Nama Pengirim",
    "pic_pengirim": "PIC Pengirim",
    "nomor_urut_penerimaan_baru": "Nomor Urut Penerimaan Baru",
    "pic_penerimaan": "PIC Penerimaan",
    "korwil_penerimaan_bapp_baru": "Korwil Penerimaan BAPP Baru",
    "timestamp_penerimaan_baru": "Timestamp Penerimaan Baru",
    # Nomor bundle = "Nomor Penerimaan Baru" (semua BAPP dengan nilai sama di
    # kolom ini dianggap satu bundle fisik). Ganti kalau kolom Anda berbeda.
    "nomor_bundle": "Nomor Penerimaan Baru",
}

# Nilai di kolom "Status BAPP Fisik Kedua" yang berarti bundle sudah siap
# dianggap masuk Tim Gate.
STATUS_AWAL_TIM_GATE = {"Open", "Diterima"}

# Alur perpindahan wajib: key = lokasi asal, value = satu-satunya tujuan valid.
ALUR_LOKASI = {"Tim Gate": "Tim SN", "Tim SN": "Tim Scan"}
LOKASI_AWAL = "Tim Gate"

DT_FORMAT = "%Y-%m-%d %H:%M:%S"

SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# =============================================================================
# 2. KONEKSI GOOGLE SHEETS (dipakai sebagai "database")
# =============================================================================

def _get_google_credentials(scopes):
    """
    Ambil credential Google dengan dua mode:
    1. Streamlit Cloud: dari st.secrets["gcp_service_account"].
    2. Lokal: dari file service_account.json (atau file yang ditentukan
       GOOGLE_SERVICE_ACCOUNT_FILE).

    Dengan pola ini service_account.json TIDAK perlu di-upload ke GitHub.
    """
    try:
        if "gcp_service_account" in st.secrets:
            service_account_info = dict(st.secrets["gcp_service_account"])

            # TOML dapat menyimpan private_key dengan \n literal. google-auth
            # membutuhkan newline yang sebenarnya.
            if "private_key" in service_account_info:
                service_account_info["private_key"] = str(
                    service_account_info["private_key"]
                ).replace("\\n", "\n")

            return Credentials.from_service_account_info(
                service_account_info,
                scopes=scopes,
            )
    except Exception as e:
        raise RuntimeError(
            "Google credential dari Streamlit Secrets gagal dibaca. "
            "Pastikan Secrets memiliki blok [gcp_service_account] "
            "dan field service-account yang lengkap."
        ) from e

    if not os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            f"File credential tidak ditemukan: {GOOGLE_SERVICE_ACCOUNT_FILE}. "
            "Untuk Streamlit Cloud, gunakan Settings → Secrets "
            "dengan blok [gcp_service_account]."
        )

    return Credentials.from_service_account_file(
        GOOGLE_SERVICE_ACCOUNT_FILE,
        scopes=scopes,
    )


@st.cache_resource(show_spinner=False)
def get_client():
    if not DB_SHEET_ID:
        raise RuntimeError(
            "DB_SHEET_ID belum diisi. Isi di .env (lokal) atau Streamlit Secrets (Cloud)."
        )
    creds = _get_google_credentials(SHEETS_SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    return get_client().open_by_key(DB_SHEET_ID)


_WORKSHEET_CACHE = {}


def get_or_create_worksheet(tab_name: str):
    """Ambil worksheet; kalau belum ada, buat otomatis + isi baris header.
    Objek worksheet disimpan di cache proses supaya tidak perlu tanya ulang
    metadata sheet ke Google tiap kali dipanggil (ini salah satu penyebab
    utama aplikasi terasa lambat, karena tiap interaksi di Streamlit
    menjalankan ulang seluruh script dari atas).
    """
    if tab_name in _WORKSHEET_CACHE:
        return _WORKSHEET_CACHE[tab_name]
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=len(TAB_HEADERS[tab_name]) + 2)
        ws.append_row(TAB_HEADERS[tab_name], value_input_option="USER_ENTERED")
    _WORKSHEET_CACHE[tab_name] = ws
    return ws


def read_rows_with_index(tab_name: str):
    """List (nomor_baris, dict) untuk tiap baris data (baris header dilewati).
    Dibaca manual (bukan get_all_records()) supaya tidak error kalau ada
    kolom kosong/duplikat di baris header.
    """
    ws = get_or_create_worksheet(tab_name)
    values = ws.get_all_values()
    if not values:
        return TAB_HEADERS[tab_name], []
    header = values[0]
    rows = []
    for i, raw in enumerate(values[1:], start=2):
        d = {header[j]: (raw[j] if j < len(raw) else "") for j in range(len(header))}
        rows.append((i, d))
    return header, rows


def read_records(tab_name: str) -> list:
    _, rows = read_rows_with_index(tab_name)
    return [d for _, d in rows]


def append_row(tab_name: str, row: list):
    append_rows(tab_name, [row])


def append_rows(tab_name: str, rows: list):
    if not rows:
        return
    ws = get_or_create_worksheet(tab_name)
    ws.append_rows(rows, value_input_option="USER_ENTERED")


def batch_update_cells(tab_name: str, cell_updates: list):
    """cell_updates: list berisi tuple (nomor_baris, nama_kolom, nilai_baru)."""
    if not cell_updates:
        return
    ws = get_or_create_worksheet(tab_name)
    header = ws.row_values(1)
    data = []
    for row_number, col_name, value in cell_updates:
        col_idx = header.index(col_name) + 1
        a1 = rowcol_to_a1(row_number, col_idx)
        data.append({"range": a1, "values": [[value]]})
    ws.batch_update(data, value_input_option="USER_ENTERED")


# =============================================================================
# 3. BACA MASTER BAPP LANGSUNG DARI SHEET ASLINYA + SINKRONISASI
# =============================================================================

@st.cache_data(ttl=120, show_spinner="Membaca Master BAPP dari Google Sheet...")
def get_master_dataframe() -> pd.DataFrame:
    ws = get_client().open_by_key(MASTER_SHEET_ID).worksheet(MASTER_SHEET_WORKSHEET_NAME)

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    header_row = values[0]
    col_index = {}
    for i, name in enumerate(header_row):
        name = name.strip()
        if name and name not in col_index:  # abaikan kolom kosong & duplikat
            col_index[name] = i

    mapped = []
    for raw_row in values[1:]:
        row = {}
        for key, sheet_header in COLUMN_MAP.items():
            idx = col_index.get(sheet_header)
            row[key] = raw_row[idx].strip() if idx is not None and idx < len(raw_row) else ""
        if row.get("nomor_bundle"):
            mapped.append(row)

    return pd.DataFrame(mapped)


def sync_master_bapp(progress_callback=None) -> dict:
    """Sinkronkan Master BAPP -> tab master_bundle. Bundle lama tidak pernah
    tertimpa lokasi/waktunya; hanya bundle baru yang ditempatkan di Tim Gate.
    """
    get_master_dataframe.clear()
    df = get_master_dataframe()

    if df.empty:
        return {"total_baris_sheet": 0, "total_bundle_di_sheet": 0, "bundle_baru": 0}

    if progress_callback:
        progress_callback(1, 3, f"Membaca {len(df)} baris dari Master BAPP...")

    _, existing_rows = read_rows_with_index(TAB_MASTER_BUNDLE)
    existing_by_nomor = {d.get("nomor_bundle"): row_no for row_no, d in existing_rows}

    grouped = df.groupby("nomor_bundle")
    new_bundle_rows, new_riwayat_rows, update_cells = [], [], []
    bundle_baru = 0

    if progress_callback:
        progress_callback(2, 3, f"Memproses {grouped.ngroups} bundle...")

    for nomor_bundle, g in grouped:
        jumlah = len(g)
        direktorat = g["direktorat"].iloc[0]

        if nomor_bundle in existing_by_nomor:
            row_no = existing_by_nomor[nomor_bundle]
            update_cells.append((row_no, "jumlah_bapp", jumlah))
            update_cells.append((row_no, "direktorat", direktorat))
            continue

        status_kedua_set = set(g["status_bapp_fisik_kedua"])
        if not (status_kedua_set & STATUS_AWAL_TIM_GATE):
            continue  # belum lolos penerimaan kedua, ditunda ke sync berikutnya

        waktu_awal = datetime.now().strftime(DT_FORMAT)
        new_bundle_rows.append([nomor_bundle, direktorat, jumlah, LOKASI_AWAL, waktu_awal, "Active"])
        new_riwayat_rows.append([nomor_bundle, "", LOKASI_AWAL, waktu_awal, ""])
        bundle_baru += 1

    if update_cells:
        batch_update_cells(TAB_MASTER_BUNDLE, update_cells)
    if new_bundle_rows:
        append_rows(TAB_MASTER_BUNDLE, new_bundle_rows)
    if new_riwayat_rows:
        append_rows(TAB_RIWAYAT, new_riwayat_rows)

    if progress_callback:
        progress_callback(3, 3, "Selesai")

    _read_master_bundle_cached.clear()

    return {
        "total_baris_sheet": len(df),
        "total_bundle_di_sheet": grouped.ngroups,
        "bundle_baru": bundle_baru,
    }


# =============================================================================
# 4. BUNDLE: LOOKUP BERDASARKAN LOKASI + DIREKTORAT + DIGIT TERAKHIR
# =============================================================================

def _row_to_bundle(d: dict) -> dict:
    return {
        "nomor_bundle": d.get("nomor_bundle", ""),
        "direktorat": d.get("direktorat", ""),
        "jumlah_bapp": int(d.get("jumlah_bapp") or 0),
        "lokasi_saat_ini": d.get("lokasi_saat_ini", ""),
        "status": d.get("status", "Active"),
    }


@st.cache_data(ttl=20, show_spinner=False)
def _read_master_bundle_cached():
    """Dicache 20 detik supaya dropdown Direktorat & pencarian digit tidak
    menarik ulang seluruh tab master_bundle dari Google Sheets di SETIAP
    interaksi (tiap klik/ketik di Streamlit menjalankan ulang seluruh
    script). Cache ini dibersihkan otomatis setiap habis Sinkronisasi Data
    atau setelah transaksi perpindahan berhasil, jadi datanya tetap segar
    persis saat dibutuhkan.
    """
    return read_records(TAB_MASTER_BUNDLE)


def get_bundles_by_lokasi(lokasi: str):
    records = _read_master_bundle_cached()
    return [
        _row_to_bundle(r) for r in records
        if r.get("lokasi_saat_ini") == lokasi and (r.get("status") or "Active") == "Active"
    ]


def get_direktorat_options(dari_lokasi: str):
    bundles = get_bundles_by_lokasi(dari_lokasi)
    return sorted({b["direktorat"] for b in bundles if b["direktorat"]})


def cari_bundle_by_digit(dari_lokasi: str, direktorat_filter: str, digits: str, exclude_nomor: set):
    """Cari bundle di lokasi tertentu (+ opsional direktorat) yang nomornya
    berakhiran `digits`, tidak termasuk yang sudah dipilih (`exclude_nomor`).
    """
    digits = digits.strip()
    if not digits:
        return []
    bundles = get_bundles_by_lokasi(dari_lokasi)
    hasil = []
    for b in bundles:
        if b["nomor_bundle"] in exclude_nomor:
            continue
        if direktorat_filter != "Semua Direktorat" and b["direktorat"] != direktorat_filter:
            continue
        if b["nomor_bundle"].endswith(digits):
            hasil.append(b)
    return hasil


# =============================================================================
# 5. TRANSAKSI PERPINDAHAN
# =============================================================================

def generate_id_transaksi(tanggal: date) -> str:
    prefix = f"TRF-{tanggal.strftime('%Y%m%d')}-"
    ids = [r.get("id_transaksi", "") for r in read_records(TAB_TRANSAKSI)]
    nums = [int(i.split("-")[-1]) for i in ids if i.startswith(prefix)]
    urut = max(nums) + 1 if nums else 1
    return f"{prefix}{urut:04d}"


def create_transaksi_perpindahan(dari_lokasi, ke_lokasi, nama_pengirim, nama_penerima,
                                  nomor_bundle_list, drive_folder_id=None, drive_folder_url=None):
    """Catat satu transaksi perpindahan untuk banyak bundle (bisa lintas
    direktorat) sekaligus. Selalu membaca ulang lokasi PALING BARU tepat
    sebelum menulis, supaya tidak menimpa perpindahan yang baru saja
    dilakukan orang lain (brief: pencegahan double-submit).
    """
    if ke_lokasi != ALUR_LOKASI.get(dari_lokasi):
        raise ValueError(f"Perpindahan {dari_lokasi} -> {ke_lokasi} tidak diperbolehkan.")

    now = datetime.now()
    _, indexed_rows = read_rows_with_index(TAB_MASTER_BUNDLE)
    by_nomor = {d.get("nomor_bundle"): (row_no, d) for row_no, d in indexed_rows}

    invalid, valid = [], []
    for nb in nomor_bundle_list:
        entry = by_nomor.get(nb)
        if entry is None or entry[1].get("lokasi_saat_ini") != dari_lokasi:
            invalid.append(nb)
        else:
            direktorat = entry[1].get("direktorat", "")
            jumlah = int(entry[1].get("jumlah_bapp") or 0)
            valid.append((nb, entry[0], direktorat, jumlah))

    if invalid:
        raise ValueError(
            "Perpindahan dibatalkan, bundle berikut sudah tidak berada di "
            f"{dari_lokasi} (mungkin baru saja dipindahkan orang lain): {', '.join(invalid)}. "
            "Kembali ke Form untuk memeriksa ulang daftar bundle."
        )
    if not valid:
        raise ValueError("Tidak ada bundle valid untuk dipindahkan.")

    id_transaksi = generate_id_transaksi(now.date())
    total_bapp = sum(v[3] for v in valid)

    append_row(TAB_TRANSAKSI, [
        id_transaksi, now.strftime("%d/%m/%Y"), now.strftime("%H:%M"),
        nama_pengirim, nama_penerima, dari_lokasi, ke_lokasi,
        len(valid), total_bapp, "Completed",
        drive_folder_id or "", drive_folder_url or "", now.strftime(DT_FORMAT),
    ])
    append_rows(TAB_DETAIL, [[id_transaksi, nb, direktorat, jumlah] for nb, _, direktorat, jumlah in valid])
    append_rows(TAB_RIWAYAT, [[nb, dari_lokasi, ke_lokasi, now.strftime(DT_FORMAT), id_transaksi] for nb, _, _, _ in valid])

    cell_updates = []
    for nb, row_no, _, _ in valid:
        cell_updates.append((row_no, "lokasi_saat_ini", ke_lokasi))
        cell_updates.append((row_no, "waktu", now.strftime(DT_FORMAT)))
    batch_update_cells(TAB_MASTER_BUNDLE, cell_updates)
    _read_master_bundle_cached.clear()

    return id_transaksi, len(valid), total_bapp


def update_transaksi_drive_info(id_transaksi, drive_folder_id, drive_folder_url):
    _, rows = read_rows_with_index(TAB_TRANSAKSI)
    for row_no, d in rows:
        if d.get("id_transaksi") == id_transaksi:
            batch_update_cells(TAB_TRANSAKSI, [
                (row_no, "drive_folder_id", drive_folder_id),
                (row_no, "drive_folder_url", drive_folder_url),
            ])
            return True
    return False


# =============================================================================
# 6. UPLOAD BUKTI KE GOOGLE DRIVE (opsional)
# =============================================================================

def drive_is_enabled() -> bool:
    return bool(DRIVE_ROOT_FOLDER_ID)


def _get_drive_service():
    creds = _get_google_credentials(
        ["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)


def _find_or_create_drive_folder(service, name: str, parent_id: str) -> str:
    query = (
        f"name = '{name}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed = false"
    )
    res = service.files().list(q=query, fields="files(id, name)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def _grant_admin_drive_access(service, folder_id: str):
    """Berikan akses baca ke folder transaksi untuk email admin.

    Ini membuat admin tetap bisa membuka PDF walaupun admin bukan orang yang
    mengisi form. ADMIN_EMAIL sebaiknya diisi dengan akun Google admin.
    Kalau ADMIN_EMAIL kosong, tidak ada permission tambahan yang dibuat.
    """
    if not ADMIN_EMAIL:
        return

    # Cek permission yang sudah ada agar tidak membuat duplikat.
    try:
        existing = service.permissions().list(
            fileId=folder_id,
            fields="permissions(id,emailAddress,type,role)"
        ).execute().get("permissions", [])
        if any(
            p.get("type") == "user"
            and p.get("emailAddress", "").lower() == ADMIN_EMAIL.lower()
            for p in existing
        ):
            return
    except Exception:
        # Kalau pengecekan permission gagal, tetap coba membuat permission.
        pass

    service.permissions().create(
        fileId=folder_id,
        body={"type": "user", "role": "reader", "emailAddress": ADMIN_EMAIL},
        sendNotificationEmail=False,
    ).execute()


def upload_bukti_transaksi(id_transaksi: str, tanggal, local_file_paths: list):
    """Upload semua file bukti ke PERPINDAHAN DOKUMEN/<tahun>/<bulan>/<id_transaksi>/."""
    if not drive_is_enabled():
        raise RuntimeError("DRIVE_ROOT_FOLDER_ID belum diisi. Isi di .env (lokal) atau Streamlit Secrets (Cloud).")

    service = _get_drive_service()
    tahun_id = _find_or_create_drive_folder(service, str(tanggal.year), DRIVE_ROOT_FOLDER_ID)
    bulan_id = _find_or_create_drive_folder(service, f"{tanggal.month:02d}", tahun_id)
    transaksi_folder_id = _find_or_create_drive_folder(service, id_transaksi, bulan_id)
    _grant_admin_drive_access(service, transaksi_folder_id)

    for path in local_file_paths:
        media = MediaFileUpload(path, resumable=True)
        service.files().create(
            body={"name": os.path.basename(path), "parents": [transaksi_folder_id]},
            media_body=media, fields="id",
        ).execute()

    folder_url = f"https://drive.google.com/drive/folders/{transaksi_folder_id}"
    return transaksi_folder_id, folder_url


# =============================================================================
# 7. TANDA TANGAN DIGITAL & PDF RINGKASAN
# =============================================================================

def _ambil_image_data_canvas(canvas_obj):
    """
    Ambil image_data dari st_canvas dengan aman.

    streamlit-drawable-canvas dapat melempar RuntimeError ketika canvas belum
    mempunyai image_data_url (misalnya pada render awal di Streamlit Cloud).
    Kondisi tersebut berarti belum ada tanda tangan, bukan berarti aplikasi
    harus crash.
    """
    if canvas_obj is None:
        return None

    try:
        image_data = canvas_obj.image_data
    except (RuntimeError, AttributeError, KeyError, TypeError):
        return None
    except Exception:
        # Jangan biarkan error internal widget merusak seluruh halaman.
        return None

    if image_data is None:
        return None

    try:
        arr = np.asarray(image_data)
    except Exception:
        return None

    if arr.size == 0 or arr.ndim < 2:
        return None

    return arr


def _ada_goresan_ttd(image_data) -> bool:
    """Cek apakah kanvas tanda tangan sudah digambar (bukan cuma kanvas putih
    kosong). image_data adalah array RGBA dari st_canvas."""
    if image_data is None:
        return False

    try:
        arr = np.asarray(image_data)
        if arr.size == 0 or arr.ndim < 2:
            return False

        rgb = arr[:, :, :3]
        return bool(np.any(rgb != 255))
    except Exception:
        return False


def _simpan_ttd_png(image_data, path: str):
    if image_data is None:
        raise ValueError("Data tanda tangan kosong.")

    arr = np.asarray(image_data)
    if arr.size == 0 or arr.ndim < 2:
        raise ValueError("Data tanda tangan tidak valid.")

    # Pastikan RGBA. Beberapa versi widget bisa menghasilkan RGB.
    if arr.shape[2] == 3:
        alpha = np.full(arr.shape[:2] + (1,), 255, dtype=arr.dtype)
        arr = np.concatenate([arr, alpha], axis=2)

    img = Image.fromarray(arr.astype("uint8"), "RGBA")
    # Tempel di atas latar putih supaya tidak transparan saat dimasukkan ke PDF.
    background = Image.new("RGB", img.size, (255, 255, 255))
    background.paste(img, mask=img.split()[3])
    background.save(path)


def _teks_pdf(t) -> str:
    """Font bawaan PDF cuma dukung Latin-1; ganti karakter yang tidak
    didukung supaya tidak error saat ada nama dengan simbol aneh."""
    return str(t).encode("latin-1", "replace").decode("latin-1")


def generate_pdf_ringkasan(data: dict, sig_pengirim_path: str, sig_penerima_path: str) -> bytes:
    """Bikin PDF ringkasan perpindahan + tanda tangan pengirim & penerima."""
    buffer = io.BytesIO()
    c = pdfcanvas.Canvas(buffer, pagesize=A4)
    _, tinggi_hal = A4
    y = tinggi_hal - 25 * mm

    c.setFont("Helvetica-Bold", 16)
    c.drawString(20 * mm, y, _teks_pdf("Bukti Perpindahan Dokumen BAPP"))
    y -= 12 * mm

    c.setFont("Helvetica", 11)
    baris = [
        f"ID Transaksi   : {data['id_transaksi']}",
        f"Tanggal/Waktu  : {data['waktu'].strftime('%d %B %Y, %H:%M')}",
        f"Pengirim       : {data['nama_pengirim']}",
        f"Penerima       : {data['nama_penerima']}",
        f"Dari -> Ke     : {data['dari_lokasi']} -> {data['ke_lokasi']}",
    ]
    for b in baris:
        c.drawString(20 * mm, y, _teks_pdf(b))
        y -= 7 * mm

    y -= 3 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20 * mm, y, _teks_pdf("Rincian per Direktorat"))
    y -= 8 * mm

    c.setFont("Helvetica", 11)
    for direktorat, agg in sorted(data["by_direktorat"].items()):
        c.drawString(20 * mm, y, _teks_pdf(f"- {direktorat}: {agg['jumlah_bundle']} Bundle / {agg['total_bapp']} BAPP"))
        y -= 7 * mm

    y -= 3 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20 * mm, y, _teks_pdf(f"TOTAL: {data['total_bundle']} Bundle / {data['total_bapp']} BAPP"))
    y -= 25 * mm

    lebar_ttd, tinggi_ttd = 70 * mm, 30 * mm
    x_pengirim, x_penerima = 20 * mm, 120 * mm
    y_gambar = y - tinggi_ttd

    c.drawImage(sig_pengirim_path, x_pengirim, y_gambar, width=lebar_ttd, height=tinggi_ttd,
                preserveAspectRatio=True, anchor="sw", mask="auto")
    c.drawImage(sig_penerima_path, x_penerima, y_gambar, width=lebar_ttd, height=tinggi_ttd,
                preserveAspectRatio=True, anchor="sw", mask="auto")

    y_label = y_gambar - 6 * mm
    c.setFont("Helvetica", 10)
    c.drawCentredString(x_pengirim + lebar_ttd / 2, y_label, _teks_pdf(f"( {data['nama_pengirim']} )"))
    c.drawCentredString(x_penerima + lebar_ttd / 2, y_label, _teks_pdf(f"( {data['nama_penerima']} )"))
    y_label -= 6 * mm
    c.drawCentredString(x_pengirim + lebar_ttd / 2, y_label, "Pengirim")
    c.drawCentredString(x_penerima + lebar_ttd / 2, y_label, "Penerima")

    c.showPage()
    c.save()
    return buffer.getvalue()


# =============================================================================
# 8. ADMIN: RIWAYAT TRANSAKSI + AKSES PDF
# =============================================================================

def _get_admin_password():
    return ADMIN_PASSWORD


def _get_admin_email():
    return ADMIN_EMAIL


def _admin_authenticated() -> bool:
    return bool(st.session_state.get("admin_authenticated", False))


def render_admin_page():
    st.subheader("🔐 Admin — Riwayat Perpindahan")
    st.caption(
        "Admin dapat melihat seluruh transaksi yang tersimpan di spreadsheet "
        "dan membuka PDF bukti, meskipun admin bukan pengisi form."
    )

    expected_password = _get_admin_password()
    if not expected_password:
        st.error(
            "Akses admin belum dikonfigurasi. Tambahkan ADMIN_PASSWORD di "
            "Streamlit Secrets (atau .env untuk lokal)."
        )
        return

    if not _admin_authenticated():
        with st.form("form_login_admin"):
            password = st.text_input("Password Admin", type="password")
            masuk = st.form_submit_button("🔓 Masuk Admin", type="primary", use_container_width=True)
        if masuk:
            if password == expected_password:
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error("Password admin salah.")
        return

    col1, col2 = st.columns([5, 1])
    with col1:
        st.success(f"Login admin aktif{(' — ' + _get_admin_email()) if _get_admin_email() else ''}")
    with col2:
        if st.button("Keluar", use_container_width=True):
            st.session_state.admin_authenticated = False
            st.rerun()

    _, rows = read_rows_with_index(TAB_TRANSAKSI)
    if not rows:
        st.info("Belum ada transaksi perpindahan.")
        return

    # Tampilkan transaksi terbaru di atas.
    rows = list(reversed(rows))
    pilihan = []
    for _, d in rows:
        tid = d.get("id_transaksi", "")
        if tid:
            pilihan.append(tid)

    selected_id = st.selectbox("Pilih ID Transaksi", pilihan)
    selected = next((d for _, d in rows if d.get("id_transaksi") == selected_id), None)
    if not selected:
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Bundle", selected.get("jumlah_bundle", "0"))
    c2.metric("BAPP", selected.get("total_bapp", "0"))
    c3.metric("Status", selected.get("status", ""))

    st.write(f"**Tanggal:** {selected.get('tanggal', '')} {selected.get('waktu', '')}")
    st.write(f"**Pengirim:** {selected.get('nama_pengirim', '')}")
    st.write(f"**Penerima:** {selected.get('nama_penerima', '')}")
    st.write(f"**Perpindahan:** {selected.get('dari_lokasi', '')} → {selected.get('ke_lokasi', '')}")

    drive_url = selected.get("drive_folder_url", "").strip()
    if drive_url:
        st.link_button("📄 Buka PDF / Folder Bukti di Google Drive", drive_url, use_container_width=True)
        st.caption("Jika akun admin belum memiliki izin Drive, pastikan ADMIN_EMAIL berisi email Google admin dan transaksi dibuat ulang setelah konfigurasi tersebut aktif.")
    else:
        st.warning(
            "PDF untuk transaksi ini belum memiliki link Google Drive. "
            "Pastikan DRIVE_ROOT_FOLDER_ID sudah diisi agar bukti PDF tersimpan permanen dan dapat diakses admin."
        )

    st.divider()
    st.write("**Daftar transaksi**")
    table_rows = []
    for _, d in rows:
        table_rows.append({
            "ID Transaksi": d.get("id_transaksi", ""),
            "Tanggal": d.get("tanggal", ""),
            "Pengirim": d.get("nama_pengirim", ""),
            "Penerima": d.get("nama_penerima", ""),
            "Dari": d.get("dari_lokasi", ""),
            "Ke": d.get("ke_lokasi", ""),
            "Bundle": d.get("jumlah_bundle", ""),
            "BAPP": d.get("total_bapp", ""),
            "Status": d.get("status", ""),
        })
    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)


# =============================================================================
# 8. STATE WIZARD
# =============================================================================

def init_wizard_state():
    defaults = {
        "wizard_step": "form",
        "selected_bundles": [],
        "data_dari_lokasi": "Tim Gate",
        "_dari_terakhir": "Tim Gate",
        "data_nama_pengirim": "",
        "data_nama_penerima": "",
        "last_result": None,
        "canvas_version": 0,
        "ttd_pengirim_image": None,
        "ttd_penerima_image": None,
        "admin_authenticated": False,
        "show_admin": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_wizard():
    st.session_state.wizard_step = "form"
    st.session_state.selected_bundles = []
    st.session_state.last_result = None
    st.session_state.ttd_pengirim_image = None
    st.session_state.ttd_penerima_image = None
    # Ganti "versi" kanvas tanda tangan supaya widget-nya dibuat ulang dari
    # kosong untuk transaksi berikutnya (bukan melanjutkan goresan lama).
    st.session_state.canvas_version = st.session_state.get("canvas_version", 0) + 1
    st.session_state.data_nama_pengirim = ""
    st.session_state.data_nama_penerima = ""
    for k in ("widget_nama_pengirim", "widget_nama_penerima", "digit_search"):
        st.session_state.pop(k, None)


# =============================================================================
# 9. HALAMAN / STEP WIZARD
# =============================================================================

def render_step_form():
    # PENTING: widget di Streamlit cuma "hidup" selama fungsi yang membuatnya
    # ikut dijalankan. Karena render_step_form() cuma jalan waktu di step
    # "form", nilai bawaan widget (key="widget_...") akan HILANG begitu
    # pindah ke step lain. Makanya setiap nilai yang masih dibutuhkan di step
    # Summary/Berhasil (nama pengirim, nama penerima, lokasi asal) langsung
    # disalin ke variabel biasa (awalan "data_") yang tidak ikut terhapus.
    col1, col2 = st.columns(2)
    with col1:
        nama_pengirim = st.text_input(
            "Nama Pengirim", key="widget_nama_pengirim",
            value=st.session_state.get("data_nama_pengirim", ""),
        )
    with col2:
        nama_penerima = st.text_input(
            "Nama Penerima", key="widget_nama_penerima",
            value=st.session_state.get("data_nama_penerima", ""),
        )
    st.session_state.data_nama_pengirim = nama_pengirim
    st.session_state.data_nama_penerima = nama_penerima

    col3, col4 = st.columns(2)
    with col3:
        dari_lokasi = st.selectbox("Dari", options=list(ALUR_LOKASI.keys()), key="widget_form_dari")
    with col4:
        ke_lokasi = ALUR_LOKASI[dari_lokasi]
        st.text_input("Ke", value=ke_lokasi, disabled=True)
    st.session_state.data_dari_lokasi = dari_lokasi

    # Kalau "Dari" diganti, bundle yang sudah dipilih sebelumnya jadi tidak
    # relevan lagi (beda lokasi asal) -> kosongkan biar konsisten.
    if st.session_state._dari_terakhir != dari_lokasi:
        st.session_state.selected_bundles = []
        st.session_state._dari_terakhir = dari_lokasi

    st.divider()
    st.subheader("Pilih Bundle")

    direktorat_options = ["Semua Direktorat"] + get_direktorat_options(dari_lokasi)
    direktorat_filter = st.selectbox("Direktorat", options=direktorat_options, key="direktorat_filter")

    digit_input = st.text_input(
        "Cari Nomor Bundle (ketik beberapa digit terakhir)",
        key="digit_search", placeholder="contoh: 125",
    )

    sudah_dipilih = {b["nomor_bundle"] for b in st.session_state.selected_bundles}

    if digit_input.strip():
        hasil = cari_bundle_by_digit(dari_lokasi, direktorat_filter, digit_input, sudah_dipilih)
        if not hasil:
            st.info("Bundle tidak ditemukan. Cek lagi lokasi/direktorat/digit yang diketik.")
        elif len(hasil) == 1:
            b = hasil[0]
            st.write(f"**{b['nomor_bundle']}** — {b['direktorat']} — {b['jumlah_bapp']} BAPP")
            if st.button("+ Tambahkan", key=f"tambah_{b['nomor_bundle']}", type="primary"):
                st.session_state.selected_bundles.append({
                    "nomor_bundle": b["nomor_bundle"], "direktorat": b["direktorat"], "jumlah_bapp": b["jumlah_bapp"],
                })
                st.session_state.pop("digit_search", None)
                st.rerun()
        else:
            hasil = hasil[:20]
            st.write(f"Ditemukan {len(hasil)} bundle, pilih salah satu:")
            opsi = {f"{b['nomor_bundle']} — {b['direktorat']} — {b['jumlah_bapp']} BAPP": b for b in hasil}
            pilihan_label = st.radio(
                "Bundle", options=list(opsi.keys()), key="radio_pilihan_bundle", label_visibility="collapsed"
            )
            if st.button("+ Tambahkan", key="tambah_dari_radio", type="primary"):
                b = opsi[pilihan_label]
                st.session_state.selected_bundles.append({
                    "nomor_bundle": b["nomor_bundle"], "direktorat": b["direktorat"], "jumlah_bapp": b["jumlah_bapp"],
                })
                st.session_state.pop("digit_search", None)
                st.session_state.pop("radio_pilihan_bundle", None)
                st.rerun()
    else:
        st.caption("Ketik minimal 1-2 digit terakhir nomor bundle untuk mencari.")

    st.divider()
    st.subheader("Bundle Terpilih")

    if not st.session_state.selected_bundles:
        st.caption("Belum ada bundle dipilih.")
    else:
        by_direktorat = {}
        for b in st.session_state.selected_bundles:
            by_direktorat.setdefault(b["direktorat"], []).append(b)

        for direktorat, items in sorted(by_direktorat.items()):
            st.markdown(f"**{direktorat}**")
            for b in items:
                c1, c2 = st.columns([5, 1])
                c1.write(f"{b['nomor_bundle']} — {b['jumlah_bapp']} BAPP")
                if c2.button("🗑️", key=f"hapus_{b['nomor_bundle']}"):
                    st.session_state.selected_bundles = [
                        x for x in st.session_state.selected_bundles if x["nomor_bundle"] != b["nomor_bundle"]
                    ]
                    st.rerun()

        total_bundle = len(st.session_state.selected_bundles)
        total_bapp = sum(b["jumlah_bapp"] for b in st.session_state.selected_bundles)
        st.metric("Total Dipilih", f"{total_bundle} Bundle / {total_bapp} BAPP")

    st.divider()
    bisa_lanjut = bool(nama_pengirim.strip()) and bool(nama_penerima.strip()) and len(st.session_state.selected_bundles) > 0
    if st.button("Lanjut ke Summary ➡️", type="primary", disabled=not bisa_lanjut, use_container_width=True):
        st.session_state.wizard_step = "summary"
        st.rerun()


def _ringkasan_per_direktorat():
    by_direktorat = {}
    for b in st.session_state.selected_bundles:
        agg = by_direktorat.setdefault(b["direktorat"], {"jumlah_bundle": 0, "total_bapp": 0})
        agg["jumlah_bundle"] += 1
        agg["total_bapp"] += b["jumlah_bapp"]
    return by_direktorat


def render_step_summary():
    st.subheader("Ringkasan Perpindahan")
    st.write(f"**Pengirim** : {st.session_state.data_nama_pengirim}")
    st.write(f"**Penerima** : {st.session_state.data_nama_penerima}")
    dari_lokasi = st.session_state.data_dari_lokasi
    st.write(f"**Dari** : {dari_lokasi}  →  **Ke** : {ALUR_LOKASI[dari_lokasi]}")

    st.divider()
    for direktorat, agg in sorted(_ringkasan_per_direktorat().items()):
        st.write(f"**{direktorat}** — {agg['jumlah_bundle']} Bundle | {agg['total_bapp']} BAPP")

    total_bundle = len(st.session_state.selected_bundles)
    total_bapp = sum(b["jumlah_bapp"] for b in st.session_state.selected_bundles)
    st.divider()
    st.metric("TOTAL", f"{total_bundle} Bundle / {total_bapp} BAPP")

    with st.expander("Lihat daftar bundle"):
        for b in st.session_state.selected_bundles:
            st.write(f"- {b['nomor_bundle']} ({b['direktorat']}, {b['jumlah_bapp']} BAPP)")

    st.divider()
    st.subheader("Tanda Tangan Digital")
    st.caption("Gambar tanda tangan dengan mouse/jari. Ikon tempat sampah di pojok kanvas untuk menghapus & mengulang.")

    versi = st.session_state.canvas_version
    colX, colY = st.columns(2)
    with colX:
        st.write("**Tanda Tangan Pengirim**")
        canvas_pengirim = st_canvas(
            stroke_width=2, stroke_color="#000000", background_color="#FFFFFF",
            height=150, width=280, drawing_mode="freedraw",
            key=f"canvas_pengirim_{versi}",
            return_image_data=True,
        )
    with colY:
        st.write("**Tanda Tangan Penerima**")
        canvas_penerima = st_canvas(
            stroke_width=2, stroke_color="#000000", background_color="#FFFFFF",
            height=150, width=280, drawing_mode="freedraw",
            key=f"canvas_penerima_{versi}",
            return_image_data=True,
        )

    # Drawable canvas dapat mengembalikan image_data=None pada rerun tertentu
    # (terutama di Streamlit Cloud). Karena itu tanda tangan yang sudah berhasil
    # diterima dari browser disimpan di session_state. Ini membuat tombol Submit
    # tetap aktif walaupun rerun berikutnya tidak lagi mengirim image_data.
    current_pengirim = _ambil_image_data_canvas(canvas_pengirim)
    current_penerima = _ambil_image_data_canvas(canvas_penerima)

    if current_pengirim is not None:
        if _ada_goresan_ttd(current_pengirim):
            st.session_state.ttd_pengirim_image = current_pengirim
        else:
            # Kanvas benar-benar kosong (misalnya tombol hapus ditekan).
            st.session_state.ttd_pengirim_image = None

    if current_penerima is not None:
        if _ada_goresan_ttd(current_penerima):
            st.session_state.ttd_penerima_image = current_penerima
        else:
            st.session_state.ttd_penerima_image = None

    ttd_pengirim_image = st.session_state.get("ttd_pengirim_image")
    ttd_penerima_image = st.session_state.get("ttd_penerima_image")

    ttd_pengirim_ok = _ada_goresan_ttd(ttd_pengirim_image)
    ttd_penerima_ok = _ada_goresan_ttd(ttd_penerima_image)

    status_col1, status_col2 = st.columns(2)
    with status_col1:
        if ttd_pengirim_ok:
            st.success("✓ TTD Pengirim terbaca")
        else:
            st.warning("TTD Pengirim belum terbaca")
    with status_col2:
        if ttd_penerima_ok:
            st.success("✓ TTD Penerima terbaca")
        else:
            st.warning("TTD Penerima belum terbaca")

    if not (ttd_pengirim_ok and ttd_penerima_ok):
        st.caption("Tanda tangan Pengirim dan Penerima wajib diisi sebelum submit.")

    st.divider()
    colA, colB = st.columns(2)
    with colA:
        if st.button("⬅️ Kembali", use_container_width=True):
            st.session_state.wizard_step = "form"
            st.rerun()
    with colB:
        if st.button(
            "✅ Submit Perpindahan", type="primary", use_container_width=True,
            disabled=not (ttd_pengirim_ok and ttd_penerima_ok),
        ):
            dari_lokasi = st.session_state.data_dari_lokasi
            ke_lokasi = ALUR_LOKASI[dari_lokasi]
            nomor_bundle_list = [b["nomor_bundle"] for b in st.session_state.selected_bundles]

            try:
                id_transaksi, n_bundle, n_bapp = create_transaksi_perpindahan(
                    dari_lokasi=dari_lokasi, ke_lokasi=ke_lokasi,
                    nama_pengirim=st.session_state.data_nama_pengirim.strip(),
                    nama_penerima=st.session_state.data_nama_penerima.strip(),
                    nomor_bundle_list=nomor_bundle_list,
                )

                waktu = datetime.now()
                hasil = {
                    "id_transaksi": id_transaksi,
                    "nama_pengirim": st.session_state.data_nama_pengirim,
                    "nama_penerima": st.session_state.data_nama_penerima,
                    "dari_lokasi": dari_lokasi,
                    "ke_lokasi": ke_lokasi,
                    "by_direktorat": _ringkasan_per_direktorat(),
                    "total_bundle": n_bundle,
                    "total_bapp": n_bapp,
                    "waktu": waktu,
                }

                tmp_dir = tempfile.mkdtemp(prefix="ttd_")
                sig_pengirim_path = os.path.join(tmp_dir, "ttd_pengirim.png")
                sig_penerima_path = os.path.join(tmp_dir, "ttd_penerima.png")
                # Gunakan image_data yang sudah dibaca aman di atas, bukan
                # mengakses canvas.image_data lagi saat proses submit.
                _simpan_ttd_png(ttd_pengirim_image, sig_pengirim_path)
                _simpan_ttd_png(ttd_penerima_image, sig_penerima_path)

                pdf_bytes = generate_pdf_ringkasan(hasil, sig_pengirim_path, sig_penerima_path)
                hasil["pdf_bytes"] = pdf_bytes

                drive_folder_url = None
                if drive_is_enabled():
                    try:
                        pdf_path = os.path.join(tmp_dir, f"{id_transaksi}.pdf")
                        with open(pdf_path, "wb") as f:
                            f.write(pdf_bytes)
                        fid, furl = upload_bukti_transaksi(id_transaksi, waktu.date(), [pdf_path])
                        update_transaksi_drive_info(id_transaksi, fid, furl)
                        drive_folder_url = furl
                    except Exception as drive_err:
                        st.warning(f"Transaksi tersimpan, tapi upload PDF ke Drive gagal: {drive_err}")
                hasil["drive_folder_url"] = drive_folder_url

                st.session_state.last_result = hasil
                st.session_state.selected_bundles = []
                st.session_state.wizard_step = "berhasil"
                st.rerun()
            except ValueError as e:
                st.error(str(e))
            except FileNotFoundError as e:
                st.error(str(e))
            except Exception as e:
                st.error(
                    "Perpindahan gagal diproses. "
                    f"Detail: {e}"
                )


def render_step_berhasil():
    r = st.session_state.last_result
    st.success("✓ PERPINDAHAN BERHASIL")
    st.write(f"**ID Transaksi** : {r['id_transaksi']}")
    st.write(f"**Pengirim** : {r['nama_pengirim']}")
    st.write(f"**Penerima** : {r['nama_penerima']}")
    st.write(f"**Dari** : {r['dari_lokasi']}  →  **Ke** : {r['ke_lokasi']}")

    st.divider()
    for direktorat, agg in sorted(r["by_direktorat"].items()):
        st.write(f"**{direktorat}** — {agg['jumlah_bundle']} Bundle | {agg['total_bapp']} BAPP")

    st.divider()
    st.metric("TOTAL", f"{r['total_bundle']} Bundle / {r['total_bapp']} BAPP")
    st.caption(r["waktu"].strftime("%d %B %Y, %H:%M"))

    st.download_button(
        "⬇️ Download PDF Ringkasan", data=r["pdf_bytes"],
        file_name=f"{r['id_transaksi']}.pdf", mime="application/pdf",
        type="primary", use_container_width=True,
    )

    if r.get("drive_folder_url"):
        st.link_button("📂 Buka PDF di Google Drive", r["drive_folder_url"], use_container_width=True)

    if st.button("➕ Buat Perpindahan Baru", use_container_width=True):
        reset_wizard()
        st.rerun()


# =============================================================================
# 10. HALAMAN UTAMA
# =============================================================================

st.set_page_config(page_title=APP_TITLE, page_icon="📦")
init_wizard_state()

top_col1, top_col2 = st.columns([6, 1])
with top_col1:
    st.title(f"📦 {APP_TITLE}")
with top_col2:
    with st.popover("⚙️"):
        st.caption("Pengaturan")
        if st.button("🔐 Admin / Riwayat PDF", use_container_width=True):
            st.session_state.show_admin = True
            st.rerun()
        st.write(f"Master BAPP: `{MASTER_SHEET_ID[:12]}...`" if MASTER_SHEET_ID else "Master BAPP: belum diatur")
        if st.button("🔄 Sinkronisasi Data"):
            try:
                with st.spinner("Menyinkronkan..."):
                    hasil = sync_master_bapp()
                st.success(
                    f"Selesai. {hasil['bundle_baru']} bundle baru ditambahkan "
                    f"(dari {hasil['total_bundle_di_sheet']} total bundle di sheet)."
                )
            except FileNotFoundError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Sinkronisasi gagal: {e}")

if st.session_state.get("show_admin", False):
    if st.button("⬅️ Kembali ke Form", use_container_width=True):
        st.session_state.show_admin = False
        st.rerun()
    render_admin_page()
else:
    STEP_RENDERERS = {
        "form": render_step_form,
        "summary": render_step_summary,
        "berhasil": render_step_berhasil,
    }
    STEP_RENDERERS[st.session_state.wizard_step]()
