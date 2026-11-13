"""
generate.py — Script utama laporan harian puskesmas (SimpusDrugsMate)
Jalankan manual: python generate.py [--uji]
Atau otomatis via simpusdrugsmate_harian.py
"""
VERSI = "2026-11-13.1"   # dasar: v2026-11-05.1 — push_github: selaraskan dengan GitHub (Actions) sebelum push, konflik dimenangkan lokal, hasil push diperiksa

import os
import sys
import json
import re
import subprocess
import ctypes
import threading
import time
from datetime import datetime, timedelta, date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent   # folder SimpusDrugsMate
sys.path.insert(0, str(ROOT))

# ── Config (bersama.pengaturan; berhenti keras kalau config tidak valid) ──
try:
    import config
    from bersama import hari as _hari
    from bersama import telegram as _telegram
    from bersama.log import pasang_log
    from bersama.pengaturan import tambah_argumen_mode
except Exception as _e:   # ConfigError / ImportError
    print(f"[ERROR] Config SimpusDrugsMate tidak bisa dimuat: {_e}")
    sys.exit(1)

P = config.PENGATURAN
BASE_DIR = ROOT
log = pasang_log(P, "generate")

NAMA_PUSKESMAS   = P.ambil("identitas", "NAMA_PUSKESMAS")
NAMA             = P.ambil("identitas", "NAMA")
NIP              = P.ambil("identitas", "NIP")
JABATAN          = P.ambil("identitas", "JABATAN")
GITHUB_USERNAME  = P.ambil("github", "GITHUB_USERNAME")
GITHUB_REPO      = P.ambil("github", "GITHUB_REPO")
GITHUB_TOKEN     = P.ambil("github", "GITHUB_TOKEN", wajib=False)   # tidak dipakai untuk push (lihat push_github)
PATH_MUTASI      = P.path_mutasi
PATH_OUTPUT      = P.path_output      # produksi: akar repo; mode uji: <repo>/uji (dari [testing])
PATH_REPO        = P.path_repo        # akar repo git — SELALU sama baik produksi maupun uji
PATH_EKINERJA    = P.path_ekinerja
WKHTMLTOPDF_PATH = P.ambil("path", "WKHTMLTOPDF_PATH", wajib=False,
                           default=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
GDRIVE_FOLDER_ID = P.gdrive_folder_id
GDRIVE_CREDS     = P.dir_rahasia / "credentials.json"
GDRIVE_TOKEN     = P.dir_rahasia / "token.json"
PASSWORD         = P.ambil("keamanan", "PASSWORD")
LOOKBACK_REG     = P.ambil_int("pengaturan", "LOOKBACK_REGISTER")
LOOKBACK_MUT     = P.ambil_int("pengaturan", "LOOKBACK_MUTASI")
TIMEOUT_POPUP    = P.ambil_int("pengaturan", "TIMEOUT_POPUP")
STARTUP_DELAY    = P.ambil_int("pengaturan", "STARTUP_DELAY_MENIT", default=7)
RETRY_MAKS       = P.ambil_int("pengaturan", "RETRY_EXCEL", default=3)
RETRY_TUNGGU     = P.ambil_int("pengaturan", "RETRY_TUNGGU_DETIK", default=30)
LAST_RUN_FILE    = P.penanda("last_run.txt")

# Mode uji: halaman di subfolder repo (uji/...) dan nama file PDF diberi awalan, supaya
# tidak pernah bentrok dengan produksi (pencarian link GDrive berdasarkan nama persis).
AWALAN_URL  = (P.awalan_url + "/") if P.awalan_url else ""
AWALAN_FILE = P.awalan_nama_file

BULAN_ID = {
    1:"Januari",2:"Februari",3:"Maret",4:"April",
    5:"Mei",6:"Juni",7:"Juli",8:"Agustus",
    9:"September",10:"Oktober",11:"November",12:"Desember"
}

# ── Nama sheet Register per bulan ───────────────────────────────
SHEET_REGISTER = {
    1:"RegisterJan", 2:"RegisterFeb", 3:"RegisterMar",
    4:"RegisterApr", 5:"RegisterMei", 6:"Register Jun",
    7:"Register Jul", 8:"Register Agu", 9:"Register Sep",
    10:"Register Okt", 11:"Register Nov", 12:"Register Des"
}

# ══════════════════════════════════════════════════════════════════
# TELEGRAM
# ══════════════════════════════════════════════════════════════════
def kirim_telegram(pesan: str, penting: bool = False):
    """Kirim lewat bersama.telegram. Pesan di sini sudah membawa emoji sendiri (🚨/ℹ️/✅/⚠️),
    jadi tanpa awalan otomatis. Di mode uji otomatis diawali [UJI]."""
    _telegram.kirim_telegram(pesan, penting=penting, p=P, polos=True)


def file_terbuka(path: Path) -> bool:
    """Cek apakah file Excel sedang dibuka (ada file lock ~$)."""
    lock = path.parent / f"~${path.name}"
    return lock.exists()


def popup_excel_terbuka(nama_file: str) -> str:
    """
    Tampilkan popup Windows via MessageBox (bisa muncul dari Task Scheduler).
    Return: 'lanjut' | 'tunda15' | 'tunda30' | 'batal'
    """
    # MB_YESNOCANCEL = 3, icon Warning = 48, topmost = 0x40000
    # Yes=6 (Lanjut), No=7 (Tunda), Cancel=2 (Batal)
    pesan = (
        f"File Excel masih terbuka:\n{nama_file}\n\n"
        f"Klik YES untuk lanjutkan sekarang\n"
        f"Klik NO untuk tunda 15 menit\n"
        f"Klik CANCEL untuk batalkan\n\n"
        f"Auto-batal dalam {TIMEOUT_POPUP} menit jika tidak ada respon."
    )
    try:
        # Jalankan di thread terpisah dengan timeout
        hasil = {"val": 2}  # default: batal

        def tampilkan():
            ret = ctypes.windll.user32.MessageBoxW(
                0, pesan,
                "⚠️ Laporan Puskesmas — File Excel Terbuka",
                0x00000003 | 0x00000030 | 0x00040000  # YesNoCancel + Warning + Topmost
            )
            hasil["val"] = ret

        t = threading.Thread(target=tampilkan)
        t.start()
        t.join(timeout=TIMEOUT_POPUP * 60)

        if t.is_alive():
            log.warning("Popup timeout — proses dibatalkan.")
            return "batal"

        ret = hasil["val"]
        if ret == 6:   return "lanjut"
        elif ret == 7: return "tunda15"
        else:          return "batal"

    except Exception as e:
        log.warning(f"Popup gagal: {e} — lanjut otomatis")
        return "lanjut"


def tunggu_excel(path: Path) -> bool:
    """
    Tangani file Excel yang terbuka.
    Return True jika boleh lanjut, False jika dibatalkan.
    """
    if not file_terbuka(path):
        return True

    log.warning(f"File terbuka: {path.name}")
    pilihan = popup_excel_terbuka(path.name)

    if pilihan == "lanjut":
        log.info(f"Lanjut meski {path.name} mungkin masih terbuka.")
        return True
    elif pilihan == "tunda15":
        log.info("Tunda 15 menit...")
        time.sleep(900)
        return tunggu_excel(path)
    else:
        log.warning(f"Dibatalkan — {path.name} masih terbuka.")
        kirim_telegram(f"⚠️ Generate dibatalkan — {path.name} masih terbuka.", penting=True)
        return False


def baca_excel_dengan_retry(path: Path, fungsi_baca, *args):
    """
    Coba baca file Excel, retry jika gagal karena sync OneDrive.
    """
    for percobaan in range(1, RETRY_MAKS + 1):
        try:
            return fungsi_baca(path, *args)
        except PermissionError as e:
            if percobaan < RETRY_MAKS:
                log.warning(f"Percobaan {percobaan}/{RETRY_MAKS} gagal ({e}) — tunggu {RETRY_TUNGGU}s...")
                time.sleep(RETRY_TUNGGU)
            else:
                raise
        except Exception:
            raise


# ══════════════════════════════════════════════════════════════════
# HARI LIBUR NASIONAL
# ══════════════════════════════════════════════════════════════════
HK = _hari.dari_pengaturan(P, log=log)
hari_kerja = HK.hari_kerja   # kompatibel dengan pemanggilan lama: hari_kerja(tgl) -> bool


def path_register(tahun: int) -> Path:
    return P.path_register(tahun)


def baca_register(tanggal_list: list) -> dict:
    """
    Baca data Register untuk list tanggal tertentu.
    Return: {date: [baris_data]}
    """
    hasil = {}
    if not tanggal_list:
        return hasil

    from openpyxl import load_workbook

    for tgl in tanggal_list:
        tahun = tgl.year
        path = path_register(tahun)

        if not path.exists():
            log.warning(f"File Register tidak ditemukan: {path}")
            kirim_telegram(f"⚠️ File Register {tahun} tidak ditemukan!\nPath: {path}", penting=True)
            continue

        bulan = tgl.month
        nama_sheet = SHEET_REGISTER.get(bulan)
        if not nama_sheet:
            log.warning(f"Nama sheet bulan {bulan} belum dikonfigurasi.")
            continue

        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if nama_sheet not in wb.sheetnames:
                log.warning(f"Sheet '{nama_sheet}' tidak ditemukan di {path.name}")
                continue

            ws = wb[nama_sheet]
            rows = []
            for row in ws.iter_rows(min_row=6, values_only=True):
                tgl_cell = row[0]  # kolom A
                if tgl_cell is None:
                    continue
                if isinstance(tgl_cell, datetime):
                    tgl_row = tgl_cell.date()
                elif isinstance(tgl_cell, date):
                    tgl_row = tgl_cell
                else:
                    continue

                if tgl_row == tgl:
                    nama_obat = row[11]  # L
                    if not row[4] and not nama_obat:
                        continue
                    # Format RM 8 digit dengan leading zero
                    rm_raw = row[3]
                    if rm_raw:
                        try:
                            rm_fmt = str(int(float(str(rm_raw)))).zfill(8)
                        except:
                            rm_fmt = str(rm_raw).zfill(8)
                    else:
                        rm_fmt = ""
                    rows.append({
                        "tanggal"    : tgl_row,
                        "no"         : row[1],   # B
                        "id_"        : row[2],   # C
                        "rm"         : rm_fmt,
                        "nama"       : row[4],   # E
                        "umur"       : row[6],   # G (skip F=Waktu)
                        "alamat"     : row[7],   # H
                        "jenis"      : row[8],   # I
                        "poli"       : row[9],   # J
                        "ispa_diare" : row[10],  # K
                        "nama_obat"  : row[11],  # L
                        "jumlah"     : row[12],  # M
                        "kode_obat"  : row[13],  # N
                        "do_"        : row[15],  # P (bukan Q — ada kolom kosong di O)
                        "sis"        : row[16],  # Q
                        "generik"    : row[17],  # R
                        "sat"        : (row[25] if len(row) > 25 else None),  # Z = kolom 1an (satuan dosis)
                    })

            if rows:
                hasil[tgl] = rows
            wb.close()

        except Exception as e:
            log.error(f"Error baca Register {tgl}: {e}")
            kirim_telegram(f"🚨 Error baca Register {tgl}:\n{e}", penting=True)

    return hasil


def baca_mutasi(tanggal_list: list) -> dict:
    """
    Baca data Mutasi untuk list tanggal.
    Return: {date: {"penerimaan": [...], "distribusi": [...]}}
    """
    hasil = {}
    if not tanggal_list or not PATH_MUTASI.exists():
        if not PATH_MUTASI.exists():
            log.error(f"File Mutasi tidak ditemukan: {PATH_MUTASI}")
            kirim_telegram(f"🚨 File Mutasi tidak ditemukan!\nPath: {PATH_MUTASI}", penting=True)
        return hasil

    from openpyxl import load_workbook
    try:
        wb = load_workbook(PATH_MUTASI, read_only=True, data_only=True)
        # Cari sheet mutasi (bisa mutasi2024, mutasi2025, dll)
        sheet_mutasi = None
        for nama in wb.sheetnames:
            if nama.lower().startswith("mutasi"):
                sheet_mutasi = nama
                break

        if not sheet_mutasi:
            log.error("Sheet mutasi tidak ditemukan di Mutasi.xlsx")
            return hasil

        ws = wb[sheet_mutasi]
        tgl_set = set(tanggal_list)

        for row in ws.iter_rows(min_row=10, values_only=True):
            tgl_cell = row[1]  # kolom B (Tgl)
            if tgl_cell is None:
                continue
            if isinstance(tgl_cell, datetime):
                tgl_row = tgl_cell.date()
            elif isinstance(tgl_cell, date):
                tgl_row = tgl_cell
            else:
                continue

            if tgl_row not in tgl_set:
                continue

            jenis = str(row[7]).strip() if row[7] else ""  # kolom H Keluar/Masuk
            item = {
                "tanggal"    : tgl_row,
                "nama_barang": row[3],   # D
                "kode"       : row[4],   # E
                "batch"      : row[5],   # F
                "dari"       : row[8],   # I (Dari)
                "ke"         : row[9],   # J (Ke)
                "jumlah"     : row[10],  # K
            }

            if tgl_row not in hasil:
                hasil[tgl_row] = {"penerimaan": [], "distribusi": []}

            if jenis.lower() == "masuk":
                hasil[tgl_row]["penerimaan"].append(item)
            elif jenis.lower() == "keluar":
                hasil[tgl_row]["distribusi"].append(item)

        wb.close()
    except Exception as e:
        log.error(f"Error baca Mutasi: {e}")
        kirim_telegram(f"🚨 Error baca Mutasi:\n{e}", penting=True)

    return filter_distribusi_minimal(hasil)


def filter_distribusi_minimal(hasil: dict) -> dict:
    """
    Distribusi: hanya catat unit (kolom 'Ke') yang menerima minimal
    config.BATAS_MIN_ITEM_DISTRIBUSI item obat dalam SATU HARI. Unit yang menerima kurang dari
    itu dibuang dari laporan harian dan rekap bulanan (keduanya memakai baca_mutasi yang sama).
    Penerimaan tidak difilter -- hanya distribusi yang dimaksud "unit menerima" di sini.
    Angka batas ada di config.py supaya gampang diubah tanpa mengubah logika ini.
    """
    batas = getattr(config, "BATAS_MIN_ITEM_DISTRIBUSI", 1)
    if batas <= 1:
        return hasil
    for data in hasil.values():
        dist = data.get("distribusi", [])
        if not dist:
            continue
        per_unit = {}
        for item in dist:
            per_unit.setdefault(item.get("ke") or "Lain-lain", []).append(item)
        data["distribusi"] = [item for grup in per_unit.values() if len(grup) >= batas for item in grup]
    return hasil


# ══════════════════════════════════════════════════════════════════
# ENKRIPSI HTML
# ══════════════════════════════════════════════════════════════════
def enkripsi_konten(html_konten: str, password: str) -> str:
    """
    Enkripsi konten HTML dengan password via XOR + base64.
    Browser akan dekripsi menggunakan hash di URL.
    """
    import base64
    key = password.encode("utf-8")
    data = html_konten.encode("utf-8")
    encrypted = bytes([data[i] ^ key[i % len(key)] for i in range(len(data))])
    return base64.b64encode(encrypted).decode("utf-8")


def buat_halaman_terenkripsi(judul: str, html_konten: str, password: str, extra_js: str = "") -> str:
    """Bungkus konten terenkripsi dalam halaman HTML dengan dekriptor JS."""
    encrypted = enkripsi_konten(html_konten, password)
    return f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{judul}</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }}
  #pesan-error {{ display:none; text-align:center; margin-top: 100px; color: #c0392b; font-size: 18px; }}
  #konten {{ display:none; }}
  #toast {{ position:fixed; bottom:24px; left:50%; transform:translateX(-50%);
             background:#2c3e50; color:white; padding:10px 20px;
             border-radius:20px; font-size:13px; opacity:0;
             transition:opacity 0.3s; pointer-events:none; z-index:9999; }}
  #toast.show {{ opacity:1; }}
</style>
</head>
<body>
<div id="toast"></div>
<div id="pesan-error">
  <h2>⛔ Akses Tidak Valid</h2>
  <p>Link tidak lengkap atau password salah.</p>
</div>
<div id="konten"></div>
<script>
function tampilToast(pesan) {{
  var toast = document.getElementById('toast');
  toast.textContent = pesan;
  toast.classList.add('show');
  setTimeout(function() {{ toast.classList.remove('show'); }}, 2500);
}}
function salinTeks(teks, label) {{
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(teks).then(function() {{
      tampilToast('✓ ' + label + ' disalin');
    }}).catch(function() {{ fallbackSalin(teks, label); }});
  }} else {{
    fallbackSalin(teks, label);
  }}
}}
function fallbackSalin(teks, label) {{
  var el = document.createElement('textarea');
  el.value = teks;
  el.style.position = 'fixed';
  el.style.opacity = '0';
  document.body.appendChild(el);
  el.focus(); el.select();
  try {{ document.execCommand('copy'); tampilToast('✓ ' + label + ' disalin'); }}
  catch(e) {{ tampilToast('Gagal salin'); }}
  document.body.removeChild(el);
}}
{extra_js}
(function(){{
  var hash = window.location.hash.slice(1);
  if (!hash) {{ document.getElementById('pesan-error').style.display='block'; return; }}
  var enc = "{encrypted}";
  try {{
    var key = unescape(encodeURIComponent(hash));
    var data = atob(enc);
    var result = '';
    for (var i = 0; i < data.length; i++) {{
      result += String.fromCharCode(data.charCodeAt(i) ^ key.charCodeAt(i % key.length));
    }}
    var decoded = decodeURIComponent(escape(result));
    var el = document.getElementById('konten');
    el.innerHTML = decoded;
    el.style.display = 'block';
  }} catch(e) {{
    document.getElementById('pesan-error').style.display = 'block';
  }}
}})();
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════
# TEMPLATE HTML KONTEN
# ══════════════════════════════════════════════════════════════════
CSS_UMUM = """
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px;
       color: #1a1a1a; background: white; padding: 12px; }
.halaman { width: 100%; }
.header { text-align: center; border-bottom: 2px solid var(--warna); padding-bottom: 12px; margin-bottom: 20px; }
.header h1 { font-size: 15px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }
.header h2 { font-size: 13px; font-weight: normal; margin-top: 4px; color: var(--warna); }
.header .meta { font-size: 11px; color: #555; margin-top: 6px; }
.judul-seksi { font-size: 12px; font-weight: bold; background: var(--warna);
               color: white; padding: 6px 12px; margin: 16px 0 0; border-radius: 4px 4px 0 0;
               width: 100%; display: block; }
table { width: 100%; border-collapse: collapse; font-size: 11px; table-layout: fixed; }
thead th { background: var(--warna-gelap); color: white; padding: 6px 8px;
           text-align: left; font-weight: 600; word-wrap: break-word; }
tbody tr:hover { background: var(--warna-hover); }
td { padding: 5px 8px; border-bottom: 1px solid #e0e0e0; vertical-align: middle;
     word-wrap: break-word; overflow-wrap: break-word; }
.grup-header td { background: var(--warna-muda); font-weight: bold; padding: 6px 8px; }
.total-row td { background: #eee; font-weight: bold; }
.footer { text-align: right; font-size: 10px; color: #888;
          margin-top: 20px; border-top: 1px solid #ccc; padding-top: 8px; width: 100%; }
@media print {
  body { padding: 0; }
  .no-print { display: none; }
}
</style>
"""

# Tema warna per jenis
TEMA = {
    "register"             : {"warna": "#3498db", "warna_gelap": "#2980b9", "warna_hover": "#d6eaf8", "warna_muda": "#d6eaf8"},
    "distribusi"           : {"warna": "#2ecc71", "warna_gelap": "#27ae60", "warna_hover": "#d5f5e3", "warna_muda": "#d5f5e3"},
    "penerimaan"           : {"warna": "#a855f7", "warna_gelap": "#9333ea", "warna_hover": "#f3e8ff", "warna_muda": "#f3e8ff"},
    "Laporan"              : {"warna": "#e67e22", "warna_gelap": "#ca6f1e", "warna_hover": "#fdebd0", "warna_muda": "#fdebd0"},
}

def css_tema(jenis: str) -> str:
    t = TEMA.get(jenis, TEMA["register"])
    return f"""<style>
:root {{
  --warna: {t['warna']};
  --warna-gelap: {t['warna_gelap']};
  --warna-hover: {t['warna_hover']};
  --warna-muda: {t['warna_muda']};
}}
</style>"""

def css_tema_pdf(jenis: str) -> str:
    """CSS dengan nilai warna inline — untuk kompatibilitas wkhtmltopdf."""
    t = TEMA.get(jenis, TEMA["register"])
    w  = t['warna']
    wg = t['warna_gelap']
    wh = t['warna_hover']
    wm = t['warna_muda']
    # CSS kolom berbeda untuk register vs mutasi
    if jenis == "register":
        css_kolom = """
thead th:nth-child(1)  { width: 4%; }
thead th:nth-child(2)  { width: 10%; }
thead th:nth-child(3)  { width: 10%; }
thead th:nth-child(4)  { width: 12%; }
thead th:nth-child(5)  { width: 7%; }
thead th:nth-child(6)  { width: 9%; }
thead th:nth-child(7)  { width: 20%; }
thead th:nth-child(8)  { width: 10%; }
thead th:nth-child(9)  { width: 12%; }
thead th:nth-child(10) { width: 6%; }"""
    else:
        css_kolom = """
thead th:nth-child(1) { width: 60%; }
thead th:nth-child(2) { width: 20%; }
thead th:nth-child(3) { width: 20%; }"""

    return f"""<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: Arial, sans-serif; font-size: 12px;
        color: #1a1a1a; background: white; padding: 12px; }}
.halaman {{ width: 100%; }}
.header {{ text-align: center; border-bottom: 2px solid {w};
           padding-bottom: 12px; margin-bottom: 20px; }}
.header h1 {{ font-size: 15px; font-weight: bold; text-transform: uppercase;
              letter-spacing: 1px; }}
.header h2 {{ font-size: 13px; font-weight: normal; margin-top: 4px; color: {w}; }}
.header .meta {{ font-size: 11px; color: #555; margin-top: 6px; }}
.judul-seksi {{ font-size: 12px; font-weight: bold; background: {w};
                color: white; padding: 6px 12px; margin: 16px 0 0;
                border-radius: 4px 4px 0 0; width: 100%; display: block; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; table-layout: fixed; }}
thead th {{ background: {wg}; color: white; padding: 6px 8px;
            text-align: left; font-weight: bold; word-wrap: break-word; }}
{css_kolom}
td {{ padding: 5px 8px; border-bottom: 1px solid #e0e0e0; vertical-align: middle;
      word-wrap: break-word; overflow-wrap: break-word; }}
.grup-header td {{ background: {wm}; font-weight: bold; padding: 6px 8px; }}
.total-row td {{ background: #eee; font-weight: bold; }}
.summary-tbl {{ border-collapse: collapse; font-size: 12px; }}
.summary-tbl td {{ padding: 2px 4px; border-bottom: 1px solid #eee; }}
.summary-wrap {{ display: flex; gap: 48px; margin: 8px 0 18px; }}
.summary-wrap h4 {{ font-size: 11px; color: #666; margin-bottom: 4px; font-weight: normal; }}
.footer {{ text-align: right; font-size: 10px; color: #888;
           margin-top: 20px; border-top: 1px solid #ccc; padding-top: 8px; width: 100%; }}
tr {{ page-break-inside: avoid; }}
</style>"""

def bersih(val) -> str:
    if val is None: return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", ""): return ""
    return s

def fmt_tgl(tgl: date) -> str:
    return f"{tgl.day} {BULAN_ID[tgl.month]} {tgl.year}"

def fmt_umur(umur) -> str:
    if umur is None: return "-"
    try:
        u = float(umur)
        if u < 1:
            return f"{int(u*12)} bln"
        return f"{int(u)} thn"
    except:
        return str(umur)

def inisial(nama: str) -> str:
    """Ubah nama lengkap ke inisial. 'Anggi Meilani' → 'A.M.'"""
    if not nama:
        return ""
    kata = str(nama).strip().split()
    return ".".join(k[0].upper() for k in kata if k) + "."

def fmt_pecahan(nilai) -> str:
    """
    Nilai dosis dari Register. Sel berformat pecahan menyimpan angka (1/3 = 0.3333333333333333),
    jadi diubah kembali ke pecahan sederhana: 0.3333 -> '1/3', 1.5 -> '1 1/2'.
    Teks (mis. '1/2 cth') dikembalikan apa adanya.
    """
    if isinstance(nilai, bool) or not isinstance(nilai, (int, float)):
        return str(nilai)
    if float(nilai).is_integer():
        return str(int(nilai))
    from fractions import Fraction
    f = Fraction(abs(nilai)).limit_denominator(12)
    if abs(float(f) - abs(nilai)) < 1e-6:
        utuh, sisa = divmod(f.numerator, f.denominator)
        teks = f"{utuh} {sisa}/{f.denominator}" if utuh else f"{sisa}/{f.denominator}"
        return ("-" if nilai < 0 else "") + teks
    # bukan pecahan sederhana: desimal pendek, jangan biarkan panjang
    return f"{nilai:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def fmt_desimal(nilai) -> str:
    """Angka dosis berstatuan ml: desimal dipertahankan (2.5 -> '2,5', 0.6 -> '0,6')."""
    if isinstance(nilai, bool) or not isinstance(nilai, (int, float)):
        return str(nilai)
    if float(nilai).is_integer():
        return str(int(nilai))
    return f"{nilai:.3f}".rstrip("0").rstrip(".").replace(".", ",")


def fmt_dosis(do_, sis, sat=None) -> str:
    """
    Teks dosis, mis. '3 x 1/3 cth' atau '3 x 2,5 ml'.
    sat = isi kolom 1an di Register (satuan sesuai tulisan di signa; kosong untuk baris lama).
    Satuan ml -> desimal dipertahankan; selain itu angka dikembalikan ke pecahan sederhana.
    """
    sat_s = str(sat).strip().lower() if sat else ""
    is_ml = bool(sat_s) and sat_s.split()[0] in ("ml", "cc")
    do_s = fmt_pecahan(do_) if do_ else ""
    sis_s = (fmt_desimal(sis) if is_ml else fmt_pecahan(sis)) if sis else ""
    if sis_s and sat_s:
        sis_s = f"{sis_s} {sat_s}"
    if do_s and sis_s:
        return f"{do_s} x {sis_s}"
    if do_s:
        return do_s
    return "-"


def html_register(tgl: date, rows: list, nama_pkm: str, nama: str, nip: str) -> str:
    """Generate konten HTML untuk halaman Register."""
    if not rows:
        return "<p>Tidak ada data kunjungan pada tanggal ini.</p>"

    # ── Kelompokkan baris per pasien berdasarkan No ──────────────
    # Pasien baru = baris yang punya Nama Pasien (tidak kosong)
    # Baris obat lanjutan = baris yang Nama Pasiennya kosong tapi ada Nama Obat
    pasien_list = []  # list of {"info": {...}, "obat": [...]}
    pasien_unik_no = set()

    for r in rows:
        nama_pasien = r.get("nama")
        nama_obat   = r.get("nama_obat")

        if nama_pasien:
            # Baris pasien baru
            entry = {"info": r, "obat": []}
            if nama_obat:
                entry["obat"].append({
                    "kode"     : bersih(str(r.get("kode_obat") or "")),
                    "nama_obat": nama_obat,
                    "dosis"    : fmt_dosis(r.get("do_"), r.get("sis"), r.get("sat")),
                    "jumlah"   : r.get("jumlah") or "",
                })
            pasien_list.append(entry)
            pasien_unik_no.add(r.get("no"))
        elif nama_obat and pasien_list:
            # Baris obat lanjutan — tempel ke pasien terakhir
            pasien_list[-1]["obat"].append({
                "kode"     : bersih(str(r.get("kode_obat") or "")),
                "nama_obat": nama_obat,
                "dosis"    : fmt_dosis(r.get("do_"), r.get("sis"), r.get("sat")),
                "jumlah"   : r.get("jumlah") or "",
            })

    total = len(pasien_list)

    # ── Hitung statistik resep ───────────────────────────────────
    total_item    = sum(len(p["obat"]) for p in pasien_list)
    total_generik = 0
    for r in rows:
        gen = str(r.get("generik") or "").strip().lower()
        if gen == "generik":
            total_generik += 1

    # ── Summary ─────────────────────────────────────────────────
    per_jenis = {}
    per_poli  = {}
    for p in pasien_list:
        j  = p["info"].get("jenis") or "Lain-lain"
        pl = p["info"].get("poli")  or "Lain-lain"
        per_jenis[j]  = per_jenis.get(j, 0) + 1
        per_poli[pl]  = per_poli.get(pl, 0) + 1

    baris_jenis = "".join(
        f'<tr><td>{j}</td><td style="text-align:right;padding-left:20px">{n}</td></tr>'
        for j, n in sorted(per_jenis.items())
    )
    baris_poli = "".join(
        f'<tr><td>{p}</td><td style="text-align:right;padding-left:20px">{n}</td></tr>'
        for p, n in sorted(per_poli.items())
    )
    baris_resep = f"""
<tr><td>Total R/</td><td style="text-align:right;padding-left:20px"><strong>{total}</strong></td></tr>
<tr><td>Total item obat</td><td style="text-align:right;padding-left:20px"><strong>{total_item}</strong></td></tr>
<tr><td>Item generik</td><td style="text-align:right;padding-left:20px"><strong>{total_generik}</strong></td></tr>
"""

    # ── Tabel detail ────────────────────────────────────────────
    baris_html = ""
    for i, p in enumerate(pasien_list, 1):
        info      = p["info"]
        obat_list = p["obat"]
        warna     = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        nama_inisial = inisial(info.get('nama') or '')

        # Baris pertama: data pasien + obat pertama sejajar
        obat0 = obat_list[0] if obat_list else None
        baris_html += f"""
<tr style="background:{warna}">
  <td style="vertical-align:middle;text-align:center;border-bottom:1px solid #ddd">{i}</td>
  <td style="vertical-align:middle;border-bottom:1px solid #ddd">{info.get('id_') or ''}</td>
  <td style="vertical-align:middle;border-bottom:1px solid #ddd">{info.get('rm') or ''}</td>
  <td style="vertical-align:middle;border-bottom:1px solid #ddd"><strong>{nama_inisial}</strong></td>
  <td style="vertical-align:middle;text-align:center;border-bottom:1px solid #ddd">{fmt_umur(info.get('umur'))}</td>
  <td style="vertical-align:middle;border-bottom:1px solid #ddd">{info.get('jenis') or ''}</td>
  <td style="vertical-align:middle;border-bottom:1px solid #ddd">{info.get('poli') or ''}</td>
  <td style="border-bottom:1px solid #ddd">{obat0['kode'] if obat0 else ''}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{obat0['dosis'] if obat0 else ''}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{obat0['jumlah'] if obat0 else ''}</td>
</tr>"""
        # Baris obat berikutnya
        for ob in obat_list[1:]:
            baris_html += f"""
<tr style="background:{warna}">
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd">{ob['kode']}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{ob['dosis']}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{ob['jumlah']}</td>
</tr>"""

    waktu_buat = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # HTML untuk web (pakai CSS variables)
    html_web = f"""<!-- dibuat: {waktu_buat} -->
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Register {fmt_tgl(tgl)}</title>
{css_tema('register')}
{CSS_UMUM}
<style>
.summary-tbl {{ border-collapse: collapse; font-size: 12px; }}
.summary-tbl td {{ padding: 2px 4px; border-bottom: 1px solid #eee; }}
.summary-wrap {{ display: flex; gap: 48px; margin: 8px 0 18px; }}
.summary-wrap h4 {{ font-size: 11px; color: #666; margin-bottom: 4px; font-weight: normal; }}
</style>
</head>
<body>
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>Register Pelayanan Apotek</h2>
    <div class="meta">{fmt_tgl(tgl)}</div>
  </div>
  <p style="margin:12px 0 6px"><strong>Ringkasan Kunjungan</strong> — Total: <strong>{total} pasien</strong></p>
  <div class="summary-wrap">
    <div><h4>Jenis Pasien</h4><table class="summary-tbl"><tbody>{baris_jenis}</tbody></table></div>
    <div><h4>Poli</h4><table class="summary-tbl"><tbody>{baris_poli}</tbody></table></div>
    <div><h4>Resep</h4><table class="summary-tbl"><tbody>{baris_resep}</tbody></table></div>
  </div>
  <div class="judul-seksi">Detail Kunjungan</div>
  <table>
    <thead><tr>
      <th style="width:30px">No</th><th>ID</th><th>RM</th><th>Nama Pasien</th>
      <th style="width:50px">Umur</th><th>Jenis</th><th>Poli</th>
      <th>Kode Obat</th><th style="width:60px">Dosis</th><th style="width:40px">Jml</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>
  <div class="footer">{nama} &nbsp;|&nbsp; NIP {nip} &nbsp;|&nbsp; {nama_pkm}</div>
</div>
</body>
</html>"""

    # HTML untuk PDF (CSS inline, tanpa CSS variables)
    html_pdf = f"""<!-- dibuat: {waktu_buat} -->
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<title>Register {fmt_tgl(tgl)}</title>
{css_tema_pdf('register')}
</head>
<body>
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>Register Pelayanan Apotek</h2>
    <div class="meta">{fmt_tgl(tgl)}</div>
  </div>
  <p style="margin:12px 0 6px"><strong>Ringkasan Kunjungan</strong> — Total: <strong>{total} pasien</strong></p>
  <div class="summary-wrap">
    <div><h4>Jenis Pasien</h4><table class="summary-tbl"><tbody>{baris_jenis}</tbody></table></div>
    <div><h4>Poli</h4><table class="summary-tbl"><tbody>{baris_poli}</tbody></table></div>
  </div>
  <div class="judul-seksi">Detail Kunjungan</div>
  <table>
    <thead><tr>
      <th style="width:30px">No</th><th>ID</th><th>RM</th><th>Nama Pasien</th>
      <th style="width:50px">Umur</th><th>Jenis</th><th>Poli</th>
      <th>Kode Obat</th><th style="width:60px">Dosis</th><th style="width:40px">Jml</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>
</div>
</body>
</html>"""

    meta = {"total": total, "jenis": "register"}
    return html_web, html_pdf, meta


def html_por(tgl: date, rows: list, nama_pkm: str, nama: str, nip: str):
    """
    Generate HTML untuk Rekap Harian POR (Non-Trajectory).
    Filter: hanya pasien yang ada ispa_diare.
    Kolom: No, Nama (inisial), Umur, Diagnosis, Nama Obat, Dosis, Jumlah.
    """
    if not rows:
        return "<p>Tidak ada data POR.</p>", "<p>Tidak ada data.</p>", {}

    # Kelompokkan per pasien
    pasien_list = []
    for r in rows:
        nama_pasien = r.get("nama")
        nama_obat   = r.get("nama_obat")
        diagnosis   = str(r.get("ispa_diare") or "").strip()
        if nama_pasien and diagnosis:
            entry = {"info": r, "obat": [], "diagnosis": diagnosis}
            if nama_obat:
                entry["obat"].append({
                    "nama_obat": nama_obat,
                    "dosis"    : fmt_dosis(r.get("do_"), r.get("sis"), r.get("sat")),
                    "jumlah"   : r.get("jumlah") or "",
                })
            pasien_list.append(entry)
        elif nama_obat and pasien_list:
            pasien_list[-1]["obat"].append({
                "nama_obat": nama_obat,
                "dosis"    : fmt_dosis(r.get("do_"), r.get("sis"), r.get("sat")),
                "jumlah"   : r.get("jumlah") or "",
            })

    if not pasien_list:
        return "<p>Tidak ada data POR.</p>", "<p>Tidak ada data.</p>", {}

    total = len(pasien_list)

    # Hitung total dari SEMUA rows register (bukan hanya ISPA/Diare)
    # Lembar resep = jumlah pasien register (baris yang punya nama)
    total_lembar  = sum(1 for r in rows if r.get("nama"))
    # Total R/ = jumlah item obat semua pasien register
    total_r       = sum(1 for r in rows if r.get("nama_obat"))
    # Total generik = jumlah baris dengan kolom generik = "Generik"
    total_generik = sum(
        1 for r in rows
        if str(r.get("generik") or "").strip().lower() == "generik"
    )

    # Summary per diagnosis (hanya ISPA/Diare)
    per_diag = {}
    for p in pasien_list:
        d = p["diagnosis"]
        per_diag[d] = per_diag.get(d, 0) + 1

    baris_diag = "".join(
        f'<tr><td>{d}</td>'
        f'<td style="text-align:right;padding-left:20px">{n} pasien</td></tr>'
        for d, n in sorted(per_diag.items())
    )

    # Summary bawah tabel — dari register keseluruhan
    summary_bawah = f"""
<div style="margin-top:12px;padding:10px 12px;background:#fdf6ee;
            border:1px solid #f0d9b5;border-radius:6px;font-size:11px;">
  <table style="border-collapse:collapse;width:auto">
    <tbody>
      <tr><td style="padding:2px 4px;color:#666">Lembar resep</td>
          <td style="padding:2px 4px 2px 20px;text-align:right"><strong>{total_lembar}</strong></td></tr>
      <tr><td style="padding:2px 4px;color:#666">Total R/</td>
          <td style="padding:2px 4px 2px 20px;text-align:right"><strong>{total_r}</strong></td></tr>
      <tr><td style="padding:2px 4px;color:#666">Total generik</td>
          <td style="padding:2px 4px 2px 20px;text-align:right"><strong>{total_generik}</strong></td></tr>
    </tbody>
  </table>
</div>"""

    # Tabel detail
    baris_html = ""
    for i, p in enumerate(pasien_list, 1):
        info      = p["info"]
        obat_list = p["obat"]
        diag      = p["diagnosis"]
        warna     = "#f9f9f9" if i % 2 == 0 else "#ffffff"
        nm_inisial = inisial(info.get("nama") or "")
        obat0 = obat_list[0] if obat_list else None
        baris_html += f"""
<tr style="background:{warna}">
  <td style="text-align:center;border-bottom:1px solid #ddd">{i}</td>
  <td style="border-bottom:1px solid #ddd"><strong>{nm_inisial}</strong></td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{fmt_umur(info.get('umur'))}</td>
  <td style="border-bottom:1px solid #ddd">{diag}</td>
  <td style="border-bottom:1px solid #ddd">{obat0['nama_obat'] if obat0 else ''}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{obat0['dosis'] if obat0 else ''}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{obat0['jumlah'] if obat0 else ''}</td>
</tr>"""
        for ob in obat_list[1:]:
            baris_html += f"""
<tr style="background:{warna}">
  <td style="border-bottom:1px solid #ddd"></td><td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd"></td><td style="border-bottom:1px solid #ddd"></td>
  <td style="border-bottom:1px solid #ddd">{ob['nama_obat']}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{ob['dosis']}</td>
  <td style="text-align:center;border-bottom:1px solid #ddd">{ob['jumlah']}</td>
</tr>"""

    waktu_buat = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    konten = f"""<!-- dibuat: {waktu_buat} -->
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>Rekap Harian POR</h2>
    <div class="meta">{fmt_tgl(tgl)}</div>
  </div>
  <p style="margin:12px 0 6px"><strong>Ringkasan ISPA/Diare</strong> — Total: <strong>{total} pasien</strong></p>
  <div style="margin:6px 0 16px">
    <table style="border-collapse:collapse;font-size:12px"><tbody>{baris_diag}</tbody></table>
  </div>
  <div class="judul-seksi">Detail Pasien</div>
  <table>
    <thead><tr>
      <th style="width:5%">No</th>
      <th style="width:15%">Nama</th>
      <th style="width:8%">Umur</th>
      <th style="width:15%">Diagnosis</th>
      <th style="width:37%">Nama Obat</th>
      <th style="width:10%">Dosis</th>
      <th style="width:10%">Jml</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>
  {summary_bawah}
</div>"""

    tema_por = TEMA.get("Laporan", TEMA["register"])
    w  = tema_por["warna"]
    wg = tema_por["warna_gelap"]

    html_web = f"""<!DOCTYPE html>
<html lang="id"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Rekap Harian POR {fmt_tgl(tgl)}</title>
<style>
body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px;
        color: #1a1a1a; background: #f5f5f5; padding: 12px; }}
.halaman {{ max-width: 900px; margin: 0 auto; background: white;
             padding: 16px; border-radius: 8px; }}
.header {{ text-align: center; border-bottom: 2px solid {w};
           padding-bottom: 12px; margin-bottom: 16px; }}
.header h1 {{ font-size: 15px; font-weight: bold; text-transform: uppercase; }}
.header h2 {{ font-size: 13px; color: {w}; margin-top: 4px; }}
.header .meta {{ font-size: 11px; color: #555; margin-top: 4px; }}
.judul-seksi {{ font-size: 12px; font-weight: bold; background: {w};
                color: white; padding: 6px 12px; margin-bottom: 0;
                border-radius: 4px 4px 0 0; display: block; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; table-layout: fixed; }}
thead th {{ background: {wg}; color: white; padding: 6px 8px;
            text-align: left; font-weight: bold; word-wrap: break-word; }}
td {{ padding: 5px 8px; border-bottom: 1px solid #e0e0e0; vertical-align: middle;
      word-wrap: break-word; }}
tbody tr:hover {{ background: #f0f8ff; }}
</style>
</head><body>{konten}</body></html>"""

    html_pdf = css_tema_pdf("Laporan") + konten
    meta = {"total": total, "jenis": "Laporan", "per_diagnosis": per_diag}
    return html_web, html_pdf, meta


def html_mutasi(tgl: date, items: list, jenis: str, nama_pkm: str, nama: str, nip: str) -> str:
    """
    Generate HTML untuk Distribusi atau Penerimaan.
    jenis: 'distribusi' atau 'penerimaan'
    """
    if not items:
        return f"<p>Tidak ada data {jenis} pada tanggal ini.</p>"

    judul_lap = "Distribusi Obat & BMHP" if jenis == "distribusi" else "Penerimaan Obat & BMHP"
    kolom_grup = "Ke" if jenis == "distribusi" else "Dari"
    key_grup   = "ke" if jenis == "distribusi" else "dari"

    # Kelompokkan per tujuan/sumber
    grup = {}
    for item in items:
        g = item.get(key_grup) or "Lain-lain"
        if g not in grup:
            grup[g] = []
        grup[g].append(item)

    jumlah_unit = len(grup)
    total_item  = sum(len(v) for v in grup.values())
    baris_html  = ""

    for nama_grup, baris_list in sorted(grup.items()):
        subcount = len(baris_list)
        baris_html += f"""
<tr class="grup-header">
  <td colspan="2">{kolom_grup}: <strong>{nama_grup}</strong></td>
  <td style="text-align:right">{subcount} item</td>
</tr>"""
        for b in baris_list:
            baris_html += f"""
<tr>
  <td style="padding-left:20px">{b.get('nama_barang') or ''}</td>
  <td>{b.get('kode') or ''}</td>
  <td style="text-align:right">{b.get('jumlah') or ''}</td>
</tr>"""

    baris_html += f"""
<tr class="total-row">
  <td colspan="2">TOTAL</td>
  <td style="text-align:right">{total_item} item</td>
</tr>"""

    label_unit = "Penerima" if jenis == "distribusi" else "Sumber"
    isi_summary = f"""
  <p style="margin:12px 0 6px"><strong>Ringkasan</strong></p>
  <div class="summary-wrap">
    <div>
      <h4>Jumlah Kegiatan</h4>
      <table class="summary-tbl"><tbody>
        <tr><td>{label_unit}</td><td style="text-align:right;padding-left:20px"><strong>{jumlah_unit}</strong></td></tr>
        <tr><td>Total Item</td><td style="text-align:right;padding-left:20px"><strong>{total_item}</strong></td></tr>
      </tbody></table>
    </div>
    <div>
      <h4>Daftar {label_unit}</h4>
      <table class="summary-tbl"><tbody>
        {"".join(f'<tr><td>{g}</td><td style="text-align:right;padding-left:20px">{len(v)} item</td></tr>' for g, v in sorted(grup.items()))}
      </tbody></table>
    </div>
  </div>"""

    isi_tabel = f"""
  <div class="judul-seksi">{judul_lap} — {fmt_tgl(tgl)}</div>
  <table>
    <thead><tr>
      <th style="width:60%">Nama Barang</th>
      <th style="width:20%">Kode</th>
      <th style="width:20%;text-align:right">Jumlah</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>"""

    waktu_buat = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    html_web = f"""<!-- dibuat: {waktu_buat} -->
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{judul_lap} {fmt_tgl(tgl)}</title>
{css_tema(jenis)}
{CSS_UMUM}
<style>
.summary-tbl {{ border-collapse: collapse; font-size: 12px; }}
.summary-tbl td {{ padding: 2px 4px; border-bottom: 1px solid #eee; }}
.summary-wrap {{ display: flex; gap: 48px; margin: 8px 0 18px; }}
.summary-wrap h4 {{ font-size: 11px; color: #666; margin-bottom: 4px; font-weight: normal; }}
</style>
</head>
<body>
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>{judul_lap}</h2>
    <div class="meta">{fmt_tgl(tgl)}</div>
  </div>
  {isi_summary}
  {isi_tabel}
</div>
</body>
</html>"""

    html_pdf = f"""<!-- dibuat: {waktu_buat} -->
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<title>{judul_lap} {fmt_tgl(tgl)}</title>
{css_tema_pdf(jenis)}
</head>
<body>
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>{judul_lap}</h2>
    <div class="meta">{fmt_tgl(tgl)}</div>
  </div>
  {isi_summary}
  {isi_tabel}
</div>
</body>
</html>"""

    meta = {"total": total_item, "unit": jumlah_unit, "jenis": jenis}
    return html_web, html_pdf, meta


def html_rekap(tgl_mulai: date, tgl_akhir: date, semua_data: list,
               jenis: str, nama_pkm: str, nama: str, nip: str) -> str:
    """Generate HTML rekap periode / bulanan."""
    if not semua_data:
        return "<p>Tidak ada data pada periode ini.</p>"

    judul_lap = {
        "register"   : "Rekap Kunjungan Pasien",
        "distribusi" : "Rekap Distribusi Obat & BMHP",
        "penerimaan" : "Rekap Penerimaan Obat & BMHP",
    }.get(jenis, "Rekap")

    periode = f"{fmt_tgl(tgl_mulai)} s.d. {fmt_tgl(tgl_akhir)}"
    baris_html = ""

    if jenis == "register":
        total_pasien = 0
        tgl_terakhir = None
        for r in semua_data:
            tgl = r.get("tanggal")
            if tgl != tgl_terakhir:
                tgl_terakhir = tgl
                baris_html += f"""
<tr class="grup-header">
  <td colspan="6">{fmt_tgl(tgl)}</td>
</tr>"""
            if r.get("nama"):
                total_pasien += 1
            nama_obat = r.get("nama_obat")
            if r.get("nama"):
                baris_html += f"""
<tr>
  <td>{r.get('no') or ''}</td>
  <td>{r.get('nama') or ''}</td>
  <td>{fmt_umur(r.get('umur'))}</td>
  <td>{r.get('jenis') or ''}</td>
  <td>{r.get('poli') or ''}</td>
  <td>{nama_obat or ''}</td>
</tr>"""
            elif nama_obat:
                baris_html += f"""
<tr>
  <td colspan="5" class="obat">{nama_obat}</td>
  <td>{r.get('jumlah') or ''}</td>
</tr>"""

        return f"""
{CSS_UMUM}
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>{judul_lap}</h2>
    <div class="meta">{periode}</div>
  </div>
  <table>
    <thead><tr>
      <th>No</th><th>Nama Pasien</th><th>Umur</th>
      <th>Jenis</th><th>Poli</th><th>Obat</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>
</div>"""

    else:
        kolom_grup = "Ke" if jenis == "distribusi" else "Dari"
        key_grup   = "ke" if jenis == "distribusi" else "dari"
        tgl_terakhir = None
        total_keseluruhan = 0

        for item in semua_data:
            tgl = item.get("tanggal")
            if tgl != tgl_terakhir:
                tgl_terakhir = tgl
                baris_html += f"""
<tr class="grup-header">
  <td colspan="4">{fmt_tgl(tgl)}</td>
</tr>"""
            total_keseluruhan += int(item.get("jumlah") or 0)
            baris_html += f"""
<tr>
  <td>{item.get(key_grup) or ''}</td>
  <td>{item.get('nama_barang') or ''}</td>
  <td>{item.get('kode') or ''}</td>
  <td style="text-align:right">{item.get('jumlah') or ''}</td>
</tr>"""

        baris_html += f"""
<tr class="total-row">
  <td colspan="3">TOTAL KESELURUHAN</td>
  <td style="text-align:right">{total_keseluruhan}</td>
</tr>"""

        return f"""
{CSS_UMUM}
<div class="halaman">
  <div class="header">
    <h1>{nama_pkm}</h1>
    <h2>{judul_lap}</h2>
    <div class="meta">{periode}</div>
  </div>
  <table>
    <thead><tr>
      <th>{kolom_grup}</th><th>Nama Barang</th>
      <th>Kode</th><th style="text-align:right">Jumlah</th>
    </tr></thead>
    <tbody>{baris_html}</tbody>
  </table>
</div>"""


# ══════════════════════════════════════════════════════════════════
# SIMPAN FILE HTML
# ══════════════════════════════════════════════════════════════════
def baca_tanggal_buat(path_output: Path) -> str:
    """Baca tanggal pembuatan dari HTML comment file yang sudah ada."""
    if not path_output.exists():
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with open(path_output, encoding="utf-8") as f:
            baris_pertama = f.readline()
        import re
        m = re.search(r'<!-- dibuat: (.+?) -->', baris_pertama)
        if m:
            return m.group(1)
    except:
        pass
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def simpan_html(path_output: Path, judul: str, konten_html: str, metadata: dict = None):
    path_output.parent.mkdir(parents=True, exist_ok=True)
    tgl_buat = baca_tanggal_buat(path_output)
    konten_html = re.sub(
        r'<!-- dibuat: .+? -->',
        f'<!-- dibuat: {tgl_buat} -->',
        konten_html
    )
    path_output.write_text(konten_html, encoding="utf-8")
    log.info(f"Disimpan: {path_output}")

    if metadata:
        meta_path = path_output.with_suffix(".json")
        import json as _json
        meta_path.write_text(_json.dumps(metadata, ensure_ascii=False), encoding="utf-8")


def simpan_pdf_ekinerja(html_konten: str, jenis: str, tgl: date) -> str:
    """
    Generate PDF ke folder Ekinerja.
    Return path PDF kalau berhasil, kosong kalau gagal.
    """
    if not PATH_EKINERJA:
        return ""
    path_pdf = path_pdf_ekinerja(jenis, tgl)
    if not path_pdf:
        return ""
    footer = f"{NAMA} \u00a0|\u00a0 NIP {NIP} \u00a0|\u00a0 {NAMA_PUSKESMAS}"
    ok = html_ke_pdf(html_konten, path_pdf, footer_kiri=footer)
    return str(path_pdf) if ok else ""


def path_harian(jenis: str, tgl: date) -> Path:
    bulan_dir = tgl.strftime("%Y-%m")
    nama_file = tgl.strftime("%Y-%m-%d") + ".html"
    return PATH_OUTPUT / bulan_dir / jenis / nama_file


def path_rekap(jenis: str, tgl_mulai: date, tgl_akhir: date) -> Path:
    bulan_dir = tgl_mulai.strftime("%Y-%m")
    if tgl_mulai.month == tgl_akhir.month and tgl_mulai.year == tgl_akhir.year:
        nama_file = f"rekap-{tgl_mulai.strftime('%Y-%m')}.html"
    else:
        nama_file = f"rekap-{tgl_mulai.strftime('%Y-%m-%d')}_{tgl_akhir.strftime('%Y-%m-%d')}.html"
    return PATH_OUTPUT / bulan_dir / "rekap" / jenis / nama_file


# ══════════════════════════════════════════════════════════════════
# PDF GENERATION
# ══════════════════════════════════════════════════════════════════

def path_pdf_ekinerja(jenis: str, tgl: date) -> Path:
    """Path PDF di folder Ekinerja."""
    if not PATH_EKINERJA:
        return None
    bulan = tgl.month
    tahun = tgl.year
    nama_bulan = f"{bulan:02d} {BULAN_ID[bulan]}"
    folder = PATH_EKINERJA / str(tahun) / nama_bulan / jenis
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{AWALAN_FILE}{jenis}-{tgl.strftime('%Y-%m-%d')}.pdf"


def html_ke_pdf(html_konten: str, path_pdf: Path, footer_kiri: str = "") -> bool:
    """Konversi HTML ke PDF pakai pdfkit. Return True kalau berhasil."""
    import tempfile
    try:
        import pdfkit
        config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

        # Footer PDF — hanya nama + NIP, tanpa nomor halaman
        footer_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  body {{ margin: 0; padding: 0 10mm;
          font-family: Arial, sans-serif; font-size: 8pt; color: #888; }}
  .footer {{ border-top: 1px solid #ccc; padding-top: 3px; }}
</style></head>
<body>
<div class="footer">
  <span>{NAMA} &nbsp;|&nbsp; NIP {NIP}</span>
</div>
</body></html>"""

        with tempfile.NamedTemporaryFile(mode='w', suffix='.html',
                                          delete=False, encoding='utf-8') as f:
            f.write(footer_html)
            footer_path = f.name

        options = {
            "page-size"           : "A4",
            "orientation"         : "Portrait",
            "margin-top"          : "10mm",
            "margin-bottom"       : "15mm",
            "margin-left"         : "10mm",
            "margin-right"        : "10mm",
            "encoding"            : "UTF-8",
            "enable-local-file-access": None,
            "quiet"               : "",
            "footer-html"         : footer_path,
            "footer-spacing"      : "3",
        }
        pdfkit.from_string(html_konten, str(path_pdf), options=options, configuration=config)
        log.info(f"PDF disimpan: {path_pdf}")
        Path(footer_path).unlink(missing_ok=True)
        return True
    except ImportError:
        log.error("pdfkit tidak terinstall. Jalankan: py -m pip install pdfkit")
        return False
    except Exception as e:
        log.error(f"Gagal buat PDF {path_pdf.name}: {e}")
        kirim_telegram(f"🚨 Gagal buat PDF {path_pdf.name}:\n{e}", penting=True)
        return False


def gdrive_link_dari_path(path_file: Path) -> str:
    """
    Dapat shareable link GDrive dari file lokal yang sudah sync.
    Pakai Drive API untuk cari file berdasarkan nama dan parent folder.
    """
    service = get_gdrive_service()
    if not service:
        return ""
    try:
        # Strip hash (#password) dari nama file kalau ada
        nama = path_file.name.split("#")[0]
        q = f"name='{nama}' and trashed=false"
        hasil = service.files().list(
            q=q, fields="files(id, name)", pageSize=5
        ).execute()
        files = hasil.get("files", [])
        if not files:
            log.warning(f"File {nama} tidak ditemukan di GDrive (belum sync?)")
            return ""
        file_id = files[0]["id"]
        return f"https://drive.google.com/file/d/{file_id}/view"
    except Exception as e:
        log.warning(f"Gagal dapat link GDrive untuk {path_file.name}: {e}")
        return ""


# ══════════════════════════════════════════════════════════════════
# SIMPAN HTML + GENERATE PDF
# ══════════════════════════════════════════════════════════════════

_gdrive_service = None

def get_gdrive_service():
    """Return authenticated Google Drive service, dengan auto-refresh token."""
    global _gdrive_service
    if _gdrive_service:
        return _gdrive_service
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build

        SCOPES = ["https://www.googleapis.com/auth/drive"]
        creds = None
        if GDRIVE_TOKEN.exists():
            creds = Credentials.from_authorized_user_file(str(GDRIVE_TOKEN), SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            GDRIVE_TOKEN.write_text(creds.to_json())
        if not creds or not creds.valid:
            raise RuntimeError("Token Google Drive tidak valid. Jalankan test_gdrive.py untuk login ulang.")
        _gdrive_service = build("drive", "v3", credentials=creds)
        return _gdrive_service
    except Exception as e:
        log.error(f"Google Drive auth gagal: {e}")
        return None


def gdrive_cari_atau_buat_folder(service, nama: str, parent_id: str = None) -> str:
    """Cari folder di GDrive, buat jika belum ada. Return folder ID."""
    query = f"name='{nama}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    hasil = service.files().list(q=query, fields="files(id, name)").execute()
    files = hasil.get("files", [])
    if files:
        return files[0]["id"]
    # Buat folder baru
    meta = {"name": nama, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        meta["parents"] = [parent_id]
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def gdrive_upload(service, path_file: Path, folder_id: str) -> str:
    """
    Upload atau update file ke Google Drive folder.
    Return shareable URL.
    """
    from googleapiclient.http import MediaFileUpload

    nama = path_file.name
    # Cek apakah file sudah ada
    query = f"name='{nama}' and '{folder_id}' in parents and trashed=false"
    hasil = service.files().list(q=query, fields="files(id)").execute()
    files = hasil.get("files", [])

    media = MediaFileUpload(str(path_file), mimetype="text/html", resumable=False)

    if files:
        # Update file yang sudah ada
        file_id = files[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
    else:
        # Upload baru
        meta = {"name": nama, "parents": [folder_id]}
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
        file_id = f["id"]
        # Set permission: anyone with link can view
        service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()

    return f"https://drive.google.com/file/d/{file_id}/view"


# Cache folder IDs supaya tidak query berulang
_folder_cache = {}

def gdrive_folder_untuk(jenis: str, bulan_str: str, service) -> str:
    """
    Cari/buat folder hierarki di GDrive:
    [Root/GDRIVE_FOLDER_ID] / [2026-09] / [jenis]
    Return folder ID.
    """
    cache_key = f"{bulan_str}/{jenis}"
    if cache_key in _folder_cache:
        return _folder_cache[cache_key]

    root_id = GDRIVE_FOLDER_ID if GDRIVE_FOLDER_ID else None
    bulan_id = gdrive_cari_atau_buat_folder(service, bulan_str, root_id)
    jenis_id  = gdrive_cari_atau_buat_folder(service, jenis, bulan_id)
    _folder_cache[cache_key] = jenis_id
    return jenis_id


def simpan_dan_upload_gdrive(path_output: Path, judul: str, konten_html: str,
                              metadata: dict = None, jenis: str = "register") -> str:
    """
    Simpan HTML lokal + upload ke Google Drive.
    Return URL Google Drive (atau kosong kalau gagal).
    """
    simpan_html(path_output, judul, konten_html, metadata)

    service = get_gdrive_service()
    if not service:
        log.warning("GDrive tidak tersedia, file hanya tersimpan lokal.")
        return ""

    try:
        bulan_str = path_output.parent.parent.name  # misal "2026-09"
        folder_id = gdrive_folder_untuk(jenis, bulan_str, service)
        url = gdrive_upload(service, path_output, folder_id)
        log.info(f"GDrive: {url}")
        return url
    except Exception as e:
        log.error(f"Upload GDrive gagal ({path_output.name}): {e}")
        kirim_telegram(f"⚠️ Upload GDrive gagal: {path_output.name}\n{e}", penting=False)
        return ""


# ══════════════════════════════════════════════════════════════════
# PUSH KE GITHUB (hanya dashboard)
# ══════════════════════════════════════════════════════════════════
def _selaraskan_dengan_remote(repo, cabang: str):
    """Ambil commit baru dari GitHub (mis. hasil GitHub Actions) lalu pasang commit lokal di atasnya.
    Kalau konflik (dashboard.html diubah di dua tempat), versi LOKAL yang menang: file itu baru saja
    dibuat dari data terbaru di PC ini."""
    import git
    repo.git.fetch("origin")
    try:
        ketinggalan = int(repo.git.rev_list("--count", f"{cabang}..origin/{cabang}") or 0)
    except git.GitCommandError:
        return   # remote belum punya cabang ini
    if ketinggalan == 0:
        return
    log.info(f"GitHub punya {ketinggalan} commit baru (mis. dari Actions) — selaraskan dulu...")
    try:
        # Saat rebase, "theirs" = commit lokal yang sedang dipasang ulang.
        repo.git.pull("--rebase", "--autostash", "-X", "theirs", "origin", cabang)
    except git.GitCommandError:
        try:
            repo.git.rebase("--abort")
        except git.GitCommandError:
            pass
        raise


def push_github():
    """
    Push HANYA dashboard.html ke GitHub (JSON/HTML harian sengaja diabaikan oleh .gitignore
    di akar repo; lihat README/PANDUAN untuk isinya).

    Akar repo git SELALU PATH_REPO (sama untuk produksi dan uji), supaya satu clone saja yang
    perlu dirawat. Yang dibatasi per mode adalah pathspec `git add`:
      - produksi: "**/dashboard.html" TAPI mengecualikan folder uji/ (supaya dashboard uji
        yang sedang dikerjakan tidak ikut terdorong oleh run produksi).
      - uji: hanya dashboard di dalam folder uji sendiri (PATH_OUTPUT, relatif ke PATH_REPO).
    Kalau PATH_OUTPUT bukan PATH_REPO maupun subfolder-nya, berhenti dengan error jelas
    (konfigurasi [testing] PATH_OUTPUT salah) daripada diam-diam push dari folder yang salah.
    """
    try:
        import git
        repo = git.Repo(PATH_REPO)

        # Hapus index.lock kalau ada (sisa proses git yang tidak selesai)
        lock_file = Path(PATH_REPO) / ".git" / "index.lock"
        if lock_file.exists():
            try:
                lock_file.unlink()
                log.info("Git index.lock dihapus otomatis.")
            except Exception as e:
                log.warning(f"Gagal hapus index.lock: {e}")

        if P.uji:
            try:
                rel_output = Path(PATH_OUTPUT).resolve().relative_to(Path(PATH_REPO).resolve())
            except ValueError:
                log.error(f"Mode uji: PATH_OUTPUT ({PATH_OUTPUT}) bukan subfolder dari PATH_REPO "
                         f"({PATH_REPO}). Push dibatalkan supaya tidak salah folder.")
                return False
            pola = f"{rel_output.as_posix()}/**/dashboard.html"
            repo.git.add(pola)
            log.info(f"Mode uji: git add dibatasi ke {pola}")
        else:
            repo.git.add("**/dashboard.html", "--", ":(exclude)uji/**")

        waktu = datetime.now().strftime("%Y-%m-%d %H:%M")
        pesan_commit = f"update dashboard uji {waktu}" if P.uji else f"update dashboard {waktu}"
        if repo.index.diff("HEAD"):          # hanya kalau ada yang di-stage (tidak ada commit kosong)
            repo.index.commit(pesan_commit)

        origin = repo.remote(name="origin")
        cabang = repo.active_branch.name

        # Ada yang perlu dikirim? (commit lokal yang belum ada di GitHub, termasuk sisa push yang dulu ditolak)
        try:
            origin.fetch()
            maju = int(repo.git.rev_list("--count", f"origin/{cabang}..{cabang}") or 0)
        except Exception:
            maju = 1   # remote/cabang belum dikenal → coba push
        if maju == 0:
            log.info("Tidak ada perubahan dashboard untuk di-push.")
            return True

        alasan = ""
        for percobaan in (1, 2):
            _selaraskan_dengan_remote(repo, cabang)
            hasil = origin.push(cabang)
            ditolak = [h for h in hasil
                       if h.flags & (h.ERROR | h.REJECTED | h.REMOTE_REJECTED | h.REMOTE_FAILURE)]
            if not ditolak:
                log.info("Push dashboard ke GitHub berhasil.")
                return True
            alasan = "; ".join((h.summary or "").strip() for h in ditolak)
            log.warning(f"Push ditolak GitHub (percobaan {percobaan}/2): {alasan}")
        petunjuk = ""
        if "email privacy" in alasan.lower() or "GH007" in alasan:
            petunjuk = (" — email commit pribadi diblokir GitHub: pakai email noreply "
                        "(git config user.email) atau matikan 'Block command line pushes that expose my email'")
        raise RuntimeError(f"Push ditolak GitHub: {alasan}{petunjuk}")
    except Exception as e:
        log.error(f"Push GitHub gagal: {e}")
        kirim_telegram(
            f"🚨 *Push GitHub GAGAL*\n{e}\n\nDashboard tidak terupdate.",
            penting=True
        )
        return False


# ══════════════════════════════════════════════════════════════════
# REKAP BULANAN
# ══════════════════════════════════════════════════════════════════
# Rekap bulanan (otomatis di akhir bulan) → PDF di Ekinerja, link lewat Google Drive.
# Folder & nama file: Ekinerja\<tahun>\<NN Bulan>\RekapRegister\RekapRegister-YYYY-MM-<hari terakhir>.pdf
JENIS_REKAP = {
    "RekapRegister"   : {"icon": "💊", "label": "Register",   "sumber": "register"},
    "RekapDistribusi" : {"icon": "📦", "label": "Distribusi", "sumber": "distribusi"},
    "RekapPenerimaan" : {"icon": "🚛", "label": "Penerimaan", "sumber": "penerimaan"},
}
REKAP_PAKAI_INISIAL = True    # nama pasien di PDF rekap register ditulis inisial (sama dengan PDF harian)
SIMPAN_REKAP_HTML   = False   # HTML rekap lama (hanya lokal, tidak pernah dipush) tidak dibuat lagi


def html_rekap_pdf(tgl_mulai: date, tgl_akhir: date, semua_data: list, jenis: str, nama_pkm: str) -> str:
    """HTML rekap bulanan khusus PDF (CSS inline, tanpa variabel CSS — kompatibel wkhtmltopdf)."""
    t = TEMA.get(jenis, TEMA["register"])
    w, wg, wm = t["warna"], t["warna_gelap"], t["warna_muda"]
    judul_lap = {
        "register"   : "Rekap Kunjungan Pasien",
        "distribusi" : "Rekap Distribusi Obat & BMHP",
        "penerimaan" : "Rekap Penerimaan Obat & BMHP",
    }.get(jenis, "Rekap")
    periode = f"{fmt_tgl(tgl_mulai)} s.d. {fmt_tgl(tgl_akhir)}"
    baris, tgl_terakhir = "", None

    if jenis == "register":
        total_pasien, jml_kolom = 0, 7
        lebar = [5, 14, 9, 11, 11, 42, 8]
        kepala = ["No", "Pasien", "Umur", "Jenis", "Poli", "Obat", "Jml"]
        for r in semua_data:
            tgl = r.get("tanggal")
            if tgl != tgl_terakhir:
                tgl_terakhir = tgl
                baris += f'<tr class="grup"><td colspan="{jml_kolom}">{fmt_tgl(tgl)}</td></tr>'
            nama_obat = r.get("nama_obat")
            if r.get("nama"):
                total_pasien += 1
                nm = inisial(r.get("nama")) if REKAP_PAKAI_INISIAL else str(r.get("nama"))
                baris += (
                    f'<tr><td>{r.get("no") or ""}</td><td>{nm}</td>'
                    f'<td>{fmt_umur(r.get("umur"))}</td><td>{r.get("jenis") or ""}</td>'
                    f'<td>{r.get("poli") or ""}</td><td>{nama_obat or ""}</td>'
                    f'<td style="text-align:right">{r.get("jumlah") or ""}</td></tr>'
                )
            elif nama_obat:
                baris += (
                    f'<tr><td colspan="5"></td><td>{nama_obat}</td>'
                    f'<td style="text-align:right">{r.get("jumlah") or ""}</td></tr>'
                )
        meta = f"{periode} &nbsp;|&nbsp; Total {total_pasien} pasien"
    else:
        kolom_grup = "Ke" if jenis == "distribusi" else "Dari"
        key_grup   = "ke" if jenis == "distribusi" else "dari"
        jml_kolom, total = 4, 0
        lebar = [26, 44, 16, 14]
        kepala = [kolom_grup, "Nama Barang", "Kode", "Jumlah"]
        for item in semua_data:
            tgl = item.get("tanggal")
            if tgl != tgl_terakhir:
                tgl_terakhir = tgl
                baris += f'<tr class="grup"><td colspan="{jml_kolom}">{fmt_tgl(tgl)}</td></tr>'
            try:
                total += int(item.get("jumlah") or 0)
            except (TypeError, ValueError):
                pass
            baris += (
                f'<tr><td>{item.get(key_grup) or ""}</td><td>{item.get("nama_barang") or ""}</td>'
                f'<td>{item.get("kode") or ""}</td>'
                f'<td style="text-align:right">{item.get("jumlah") or ""}</td></tr>'
            )
        baris += (f'<tr class="total"><td colspan="3">TOTAL KESELURUHAN</td>'
                  f'<td style="text-align:right">{total}</td></tr>')
        meta = periode

    cols = "".join(f'<col style="width:{x}%">' for x in lebar)
    kanan = ' style="text-align:right"'
    th = "".join(f'<th{kanan if k in ("Jml", "Jumlah") else ""}>{k}</th>' for k in kepala)
    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><title>{judul_lap} {periode}</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: Arial, sans-serif; font-size: 11px; color: #1a1a1a; background: white; padding: 8px; }}
.header {{ text-align: center; border-bottom: 2px solid {w}; padding-bottom: 10px; margin-bottom: 14px; }}
.header h1 {{ font-size: 15px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }}
.header h2 {{ font-size: 13px; font-weight: normal; margin-top: 4px; color: {w}; }}
.header .meta {{ font-size: 11px; color: #555; margin-top: 5px; }}
table {{ width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10.5px; }}
thead {{ display: table-header-group; }}
tr {{ page-break-inside: avoid; }}
th {{ background: {wg}; color: white; padding: 5px 6px; text-align: left; font-weight: bold; }}
td {{ padding: 4px 6px; border-bottom: 1px solid #e0e0e0; vertical-align: top; word-wrap: break-word; }}
.grup td {{ background: {wm}; font-weight: bold; }}
.total td {{ background: #eee; font-weight: bold; }}
</style></head><body>
<div class="header"><h1>{nama_pkm}</h1><h2>{judul_lap}</h2><div class="meta">{meta}</div></div>
<table><colgroup>{cols}</colgroup><thead><tr>{th}</tr></thead><tbody>{baris}</tbody></table>
</body></html>"""


def generate_rekap_bulanan(tahun: int, bulan: int):
    """Generate rekap seluruh bulan → PDF di Ekinerja (RekapRegister / RekapDistribusi / RekapPenerimaan)."""
    import calendar
    tgl_mulai = date(tahun, bulan, 1)
    hari_terakhir = calendar.monthrange(tahun, bulan)[1]
    tgl_akhir = date(tahun, bulan, hari_terakhir)

    log.info(f"Generate rekap bulanan {BULAN_ID[bulan]} {tahun}...")

    tanggal_list = [
        tgl_mulai + timedelta(days=i)
        for i in range((tgl_akhir - tgl_mulai).days + 1)
        if hari_kerja(tgl_mulai + timedelta(days=i))
    ]

    data_reg = baca_register(tanggal_list)
    data_mut = baca_mutasi(tanggal_list)

    kumpulan = {
        "register"  : [r for tgl in sorted(data_reg) for r in data_reg[tgl]],
        "distribusi": [i for tgl in sorted(data_mut) for i in data_mut[tgl].get("distribusi", [])],
        "penerimaan": [i for tgl in sorted(data_mut) for i in data_mut[tgl].get("penerimaan", [])],
    }

    hasil = {}
    for jenis_rekap, cfg_r in JENIS_REKAP.items():
        sumber = cfg_r["sumber"]
        data = kumpulan.get(sumber) or []
        if not data:
            log.info(f"Rekap {sumber}: tidak ada data bulan ini, dilewati.")
            continue
        try:
            if SIMPAN_REKAP_HTML:
                simpan_html(path_rekap(sumber, tgl_mulai, tgl_akhir),
                            f"Rekap {sumber} {BULAN_ID[bulan]} {tahun}",
                            html_rekap(tgl_mulai, tgl_akhir, data, sumber, NAMA_PUSKESMAS, NAMA, NIP))
            html_pdf = html_rekap_pdf(tgl_mulai, tgl_akhir, data, sumber, NAMA_PUSKESMAS)
            pdf = simpan_pdf_ekinerja(html_pdf, jenis_rekap, tgl_akhir)
            if pdf:
                hasil[jenis_rekap] = Path(pdf)
        except Exception as e:
            log.error(f"Error rekap {sumber}: {e}")
            kirim_telegram(f"🚨 Error rekap {sumber} {bulan}/{tahun}:\n{e}", penting=True)

    log.info(f"Rekap bulanan {BULAN_ID[bulan]} {tahun} selesai ({len(hasil)} PDF).")
    return hasil


# ══════════════════════════════════════════════════════════════════
# DASHBOARD KALENDER
# ══════════════════════════════════════════════════════════════════

# Konfigurasi jenis file — tambah di sini kalau ada jenis baru
JENIS_FILE = {
    # Jenis generate otomatis dari Excel
    "register"              : {"icon": "💊", "label": "Register",             "scan_gdrive": True,  "ekstensi": [".pdf"]},
    "distribusi"            : {"icon": "📦", "label": "Distribusi",           "scan_gdrive": True,  "ekstensi": [".pdf"]},
    "penerimaan"            : {"icon": "🚛", "label": "Penerimaan",           "scan_gdrive": True,  "ekstensi": [".pdf"]},
    # Jenis manual / via bot
    "ApelLokmin"            : {"icon": "🍎", "label": "Apel Lokmin",          "scan_gdrive": True,  "ekstensi": [".pdf", ".jpg", ".jpeg", ".png"]},
    "Laporan"               : {"icon": "📑", "label": "Laporan",               "scan_gdrive": True, "ekstensi": [".pdf", ".jpg", ".jpeg", ".png"]},
    "Penyimpanan"           : {"icon": "🗃️", "label": "Penyimpanan",          "scan_gdrive": True,  "ekstensi": [".pdf", ".jpg", ".jpeg", ".png"]},
    "Rekap"                 : {"icon": "📈", "label": "Rekap",                "scan_gdrive": True,  "ekstensi": [".pdf", ".jpg", ".jpeg", ".png"]},
    "TugasTambahan"         : {"icon": "💉", "label": "Tugas Tambahan",       "scan_gdrive": True,  "ekstensi": [".pdf", ".jpg", ".jpeg", ".png"]},
}

def gabung_penerimaan_lengkap(tgl: date):
    """
    Penerimaan + berita acara: gabung PDF hasil generate dengan pindaian bersuffix satu huruf
    (penerimaan-YYYY-MM-DD-a.pdf, -b, ...) menjadi penerimaan-YYYY-MM-DD-lengkap.pdf
    di folder Ekinerja yang sama. Urutan: pindaian/berita acara (a, b, ...) dulu, hasil generate selalu terakhir.
    File asli dan pindaian tidak diubah. Digabung ulang hanya bila ada yang lebih baru.
    Return Path file -lengkap bila ada, selain itu None.
    """
    if not PATH_EKINERJA:
        return None
    import re as _re
    folder = PATH_EKINERJA / str(tgl.year) / f"{tgl.month:02d} {BULAN_ID[tgl.month]}" / "penerimaan"
    dasar  = f"{AWALAN_FILE}penerimaan-{tgl.strftime('%Y-%m-%d')}"
    asli   = folder / f"{dasar}.pdf"
    lengkap = folder / f"{dasar}-lengkap.pdf"
    try:
        if not folder.exists():
            return None
        pola = _re.compile(rf"^{_re.escape(dasar)}-([A-Za-z])\.pdf$", _re.IGNORECASE)
        scan = sorted((f for f in folder.iterdir() if pola.match(f.name)),
                      key=lambda f: pola.match(f.name).group(1).lower())
        if not scan or not asli.exists():
            return lengkap if lengkap.exists() else None
        bagian = scan + [asli]
        if lengkap.exists() and lengkap.stat().st_mtime >= max(f.stat().st_mtime for f in bagian):
            return lengkap
        from pypdf import PdfWriter, PdfReader
        w = PdfWriter()
        for f in bagian:
            w.append(PdfReader(str(f)))
        tmp = lengkap.with_name(lengkap.stem + ".tmp_gabung")
        with open(tmp, "wb") as out:
            w.write(out)
        w.close()
        os.replace(tmp, lengkap)
        log.info(f"Penerimaan digabung: {lengkap.name} ← {', '.join(f.name for f in bagian)}")
        return lengkap
    except Exception as e:
        log.warning(f"Gagal gabung penerimaan {tgl}: {e}")
        return lengkap if lengkap.exists() else None


def generate_dashboard(tahun: int, bulan: int):
    """Generate halaman kalender dashboard untuk satu bulan."""
    import calendar
    base_url  = f"https://{GITHUB_USERNAME}.github.io/{GITHUB_REPO}"
    bulan_str = f"{tahun}-{bulan:02d}"
    nama_bulan = f"{BULAN_ID[bulan]} {tahun}"

    # Scan file yang ada di output folder (register/distribusi/penerimaan dari GitFiles)
    tersedia = {}
    bulan_dir = PATH_OUTPUT / bulan_str
    JENIS_GITFILES = {"register", "distribusi", "penerimaan"}
    for jenis, cfg_j in JENIS_FILE.items():
        if not cfg_j.get("scan_gdrive"):
            continue
        ekstensi = cfg_j.get("ekstensi", [".pdf"])

        # Jenis generate otomatis → scan dari GitFiles (*.html)
        if jenis in JENIS_GITFILES:
            jenis_dir = bulan_dir / jenis
            if jenis_dir.exists():
                for f in jenis_dir.glob("*.html"):
                    try:
                        tgl = date.fromisoformat(f.stem)
                        if tgl not in tersedia:
                            tersedia[tgl] = {}
                        tersedia[tgl][jenis] = f"{base_url}/{AWALAN_URL}{bulan_str}/{jenis}/{f.stem}.html#{PASSWORD}"
                    except:
                        pass
        else:
            # Jenis manual → scan dari folder Ekinerja GDrive
            # tersedia[tgl][jenis] berisi dict:
            #   {"tunggal": Path, "huruf": [Path,...], "angka": [Path,...]}
            if not PATH_EKINERJA:
                continue
            nama_bulan_folder = f"{bulan:02d} {BULAN_ID[bulan]}"
            folder_ekinerja = PATH_EKINERJA / str(tahun) / nama_bulan_folder / jenis
            if not folder_ekinerja.exists():
                continue
            import re as _re
            for ext in ekstensi:
                for f in sorted(folder_ekinerja.glob(f"*{ext}")):
                    stem = f.stem
                    # Format: NamaJenis-YYYY-MM-DD[-suffix]
                    # suffix bisa: huruf (a-z) → gabung, angka (0-9) → pisah, kosong → tunggal
                    m = _re.match(
                        rf"^{_re.escape(jenis)}-(\d{{4}}-\d{{2}}-\d{{2}})(?:-([a-zA-Z0-9]+))?$",
                        stem
                    )
                    if not m:
                        continue
                    try:
                        tgl = date.fromisoformat(m.group(1))
                    except:
                        continue
                    suffix = m.group(2) or ""

                    if tgl not in tersedia:
                        tersedia[tgl] = {}
                    if jenis not in tersedia[tgl]:
                        tersedia[tgl][jenis] = {"tunggal": None, "huruf": [], "angka": [], "folder": folder_ekinerja}

                    if not suffix:
                        tersedia[tgl][jenis]["tunggal"] = f
                    elif suffix.isalpha():
                        tersedia[tgl][jenis]["huruf"].append(f)
                    elif suffix.isdigit():
                        tersedia[tgl][jenis]["angka"].append(f)

    # Rekap bulanan: PDF di Ekinerja (dibuat tiap akhir bulan); link lewat Google Drive
    rekap_links = {}
    if PATH_EKINERJA:
        folder_bulan_rekap = PATH_EKINERJA / str(tahun) / f"{bulan:02d} {BULAN_ID[bulan]}"
        for jenis_r in JENIS_REKAP:
            folder_r = folder_bulan_rekap / jenis_r
            if not folder_r.exists():
                continue
            pdf_r = sorted(folder_r.glob(f"{AWALAN_FILE}{jenis_r}-*.pdf"))
            if not pdf_r:
                continue
            url_r = gdrive_link_dari_path(pdf_r[-1])
            if url_r:
                rekap_links[jenis_r] = (url_r, pdf_r[-1].name)
            else:
                log.warning(f"Link GDrive rekap {pdf_r[-1].name} belum ditemukan (belum sync?)")

    # Build kalender
    cal = calendar.monthcalendar(tahun, bulan)
    hari_nama = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]

    header_hari = "".join(f'<th>{h}</th>' for h in hari_nama)

    baris_kalender = ""
    maks_item = 0
    for minggu in cal:
        baris_kalender += "<tr>"
        for hari in minggu:
            if hari == 0:
                baris_kalender += '<td class="kosong"></td>'
            else:
                tgl = date(tahun, bulan, hari)
                is_minggu = tgl.weekday() == 6
                is_libur  = tgl in HK.libur(tahun)
                is_today  = tgl == date.today()

                kelas = "hari"
                if is_minggu or is_libur:
                    kelas += " libur"
                if is_today:
                    kelas += " today"

                items = []
                if tgl in tersedia:
                    import json as _json
                    for jenis, cfg_j in JENIS_FILE.items():
                        if jenis not in tersedia[tgl]:
                            continue
                        gdrive_url = ""
                        jml_label  = ""
                        JENIS_GITFILES = {"register", "distribusi", "penerimaan"}

                        if jenis in JENIS_GITFILES:
                            path_json = bulan_dir / jenis / (tgl.strftime("%Y-%m-%d") + ".json")
                            if path_json.exists():
                                try:
                                    meta = _json.loads(path_json.read_text(encoding="utf-8"))
                                    gdrive_url = meta.get("gdrive_url", "")
                                    if jenis == "register":
                                        jml_label = str(meta.get("total", ""))
                                    elif jenis == "distribusi":
                                        jml_label = f"{meta.get('unit', '')}u"
                                    else:
                                        jml_label = str(meta.get("total", ""))
                                except:
                                    pass
                            nama_file = f"{AWALAN_FILE}{jenis}-{tgl.strftime('%Y-%m-%d')}.pdf"
                            # Penerimaan + berita acara → pakai versi -lengkap bila ada
                            if jenis == "penerimaan":
                                p_lengkap = gabung_penerimaan_lengkap(tgl)
                                if p_lengkap:
                                    url_lengkap = gdrive_link_dari_path(p_lengkap)
                                    if url_lengkap:
                                        gdrive_url = url_lengkap
                                        nama_file  = p_lengkap.name
                            # Untuk GitFiles: tidak tampilkan kalau belum ada GDrive link
                            if not gdrive_url:
                                continue
                            url_copy  = gdrive_url
                        else:
                            # Jenis manual — data berupa dict {tunggal, huruf, angka}
                            data_manual = tersedia[tgl][jenis]
                            jenis_icon  = cfg_j["icon"]
                            jenis_nama  = cfg_j["label"]
                            tgl_pendek  = tgl.strftime("%d/%m")
                            tgl_label   = f"{tgl.day} {BULAN_ID[tgl.month]} {tgl.year}"

                            def _buat_icon_row(path_lokal, label_suffix=""):
                                url = gdrive_link_dari_path(path_lokal)
                                if not url:
                                    return ""
                                nama  = path_lokal.name
                                label = f"{jenis_nama} {label_suffix or tgl_label}"
                                teks  = f"{url} | {nama}"
                                return (
                                    f'<div class="ic" title="{jenis_nama} {tgl_pendek} — {nama}">'
                                    f'<span class="ico" onclick="salinTeks(\'{teks}\', \'{label}\')">{jenis_icon}</span>'
                                    f'<span class="jml" onclick="window.open(\'{url}\',\'_blank\')">buka</span>'
                                    f'</div>'
                                )

                            # File tunggal (tanpa suffix)
                            if data_manual.get("tunggal"):
                                items.append(_buat_icon_row(data_manual["tunggal"]))
                            # File suffix huruf → digabung jadi satu PDF dulu (akan dihandle nanti)
                            # Untuk sekarang: tampilkan icon bila PDF gabungan sudah ada
                            if data_manual.get("huruf"):
                                tgl_str_dash = tgl.strftime("%Y-%m-%d")
                                pdf_gabung   = Path(data_manual["folder"]) / f"{jenis}-{tgl_str_dash}.pdf"
                                if pdf_gabung.exists():
                                    items.append(_buat_icon_row(pdf_gabung))
                            # File suffix angka → masing-masing icon sendiri
                            for f_angka in data_manual.get("angka", []):
                                items.append(_buat_icon_row(f_angka, f"{jenis_nama} {tgl_pendek}"))
                            continue  # skip blok teks_copy di bawah

                        teks_copy   = f"{url_copy} | {nama_file}"
                        label_toast = f"{cfg_j['label']} {tgl.day} {BULAN_ID[tgl.month]} {tgl.year}"
                        jenis_icon  = cfg_j["icon"]
                        jenis_nama  = cfg_j["label"]
                        tgl_pendek  = tgl.strftime("%d/%m")
                        # Klik icon → copy link, klik angka → buka GDrive
                        items.append(
                            f'<div class="ic" title="{jenis_nama} {tgl_pendek}">'
                            f'<span class="ico" onclick="salinTeks(\'{teks_copy}\', \'{label_toast}\')">{jenis_icon}</span>'
                            f'<span class="jml" onclick="window.open(\'{gdrive_url}\',\'_blank\')">{jml_label}</span>'
                            f'</div>'
                        )

                items = [i for i in items if i]
                maks_item = max(maks_item, len(items))
                isi_td = (
                    f'<div class="tgl-num">{hari}</div>'
                    f'<div class="ikon-grid">{"".join(items)}</div>'
                ) if items else f'<div class="tgl-num">{hari}</div>'

                baris_kalender += f'<td class="{kelas}">{isi_td}</td>'
        baris_kalender += "</tr>"

    # Tinggi baris seragam: ikut hari dengan ikon terbanyak (2 ikon per baris)
    n_baris_ikon  = max(1, -(-maks_item // 2))
    tinggi_sempit = 30 + n_baris_ikon * 34    # layar sempit: teks di bawah ikon
    tinggi_lebar  = 34 + n_baris_ikon * 26    # layar lebar: teks di samping ikon

    # Rekap section
    rekap_html = ""
    if rekap_links:
        rekap_html = '<div class="rekap-section"><strong>Rekap Bulan Ini:</strong> '
        for jenis_r, (url_r, nama_file_r) in rekap_links.items():
            cfg_r = JENIS_REKAP[jenis_r]
            teks_r = f"{url_r} | {nama_file_r}"
            rekap_html += (
                f'<span class="rekap-link" title="{nama_file_r}">'
                f'<span onclick="salinTeks(\'{teks_r}\', \'Rekap {cfg_r["label"]} {nama_bulan}\')">{cfg_r["icon"]}</span> '
                f'<span class="rekap-buka" onclick="window.open(\'{url_r}\',\'_blank\')">{cfg_r["label"]}</span>'
                f'</span> '
            )
        rekap_html += '</div>'

    # Legenda bawah — hitung total per jenis
    import json as _json2
    total_per_jenis = {}
    for jenis, cfg_j in JENIS_FILE.items():
        total  = 0
        satuan = "dokumen"
        if jenis == "register":
            satuan = "pasien"
            # Baca semua JSON di folder register bulan ini
            json_dir = bulan_dir / "register"
            if json_dir.exists():
                for jf in json_dir.glob("*.json"):
                    try:
                        meta = _json2.loads(jf.read_text(encoding="utf-8"))
                        total += int(meta.get("total", 0))
                    except: pass
        elif jenis == "distribusi":
            satuan = "unit"
            json_dir = bulan_dir / "distribusi"
            if json_dir.exists():
                for jf in json_dir.glob("*.json"):
                    try:
                        meta = _json2.loads(jf.read_text(encoding="utf-8"))
                        total += int(meta.get("unit", 0))
                    except: pass
        else:
            # Hitung jumlah hari yang ada dokumen
            total = sum(1 for tgl, data in tersedia.items() if jenis in data)

        if total > 0:
            total_per_jenis[jenis] = (cfg_j["icon"], cfg_j["label"], total, satuan)

    legenda_bawah = ""
    if total_per_jenis:
        baris_legenda = "".join(
            f'<div class="leg-row">'
            f'<span class="leg-ico">{icon}</span>'
            f'<span class="leg-nama">{label}</span>'
            f'<span class="leg-total">{total} {satuan}</span>'
            f'</div>'
            for icon, label, total, satuan in total_per_jenis.values()
        )
        legenda_bawah = f'<div class="legenda-bawah">{baris_legenda}</div>'

    # ── Navigasi bulan ───────────────────────────────────────────
    import calendar as _cal
    # Bulan sebelumnya
    if bulan == 1:
        bln_prev, thn_prev = 12, tahun - 1
    else:
        bln_prev, thn_prev = bulan - 1, tahun
    # Bulan berikutnya
    if bulan == 12:
        bln_next, thn_next = 1, tahun + 1
    else:
        bln_next, thn_next = bulan + 1, tahun

    str_prev = f"{thn_prev}-{bln_prev:02d}"
    str_next = f"{thn_next}-{bln_next:02d}"
    path_prev = PATH_OUTPUT / str_prev / "dashboard.html"
    path_next = PATH_OUTPUT / str_next / "dashboard.html"

    url_prev = f"{base_url}/{AWALAN_URL}{str_prev}/dashboard.html#{PASSWORD}" if path_prev.exists() else ""
    url_next = f"{base_url}/{AWALAN_URL}{str_next}/dashboard.html#{PASSWORD}" if path_next.exists() else ""

    nav_html = '<div class="nav-bulan">'
    if url_prev:
        nav_html += f'<a class="nav-btn" href="{url_prev}">&#8592; {BULAN_ID[bln_prev]}</a>'
    else:
        nav_html += '<span class="nav-btn nav-disabled"></span>'
    nav_html += f'<span class="nav-label">{nama_bulan}</span>'
    if url_next:
        nav_html += f'<a class="nav-btn" href="{url_next}">{BULAN_ID[bln_next]} &#8594;</a>'
    else:
        nav_html += '<span class="nav-btn nav-disabled"></span>'
    nav_html += '</div>'

    html = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard {nama_bulan} — {NAMA_PUSKESMAS}</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px;
        background: #f0f4f8; color: #1a1a1a; padding: 8px; }}
.container {{ max-width: 700px; margin: 0 auto; background: white;
              border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.08);
              overflow: hidden; }}
.header {{ background: #3498db; color: white; padding: 14px 20px; text-align: center; }}
.header h1 {{ font-size: 13px; font-weight: normal; opacity: 0.9; }}
.nav-bulan {{ display: flex; align-items: center; justify-content: space-between;
               margin-top: 8px; }}
.nav-btn {{ color: white; text-decoration: none; font-size: 13px; font-weight: 500;
             padding: 4px 10px; border-radius: 16px; background: rgba(255,255,255,0.2);
             min-width: 80px; text-align: center; }}
.nav-btn:hover {{ background: rgba(255,255,255,0.35); }}
.nav-disabled {{ min-width: 80px; }}
.nav-label {{ font-size: 19px; font-weight: bold; }}
.body {{ padding: 10px; }}
.legenda-bawah {{ margin-top: 12px; padding: 10px 12px;
                   background: #f8fafc; border-radius: 8px;
                   border: 1px solid #e4eaf0; }}
.leg-row {{ display: flex; align-items: center; gap: 6px;
             padding: 3px 0; font-size: 11px;
             border-bottom: 1px solid #eef2f7; }}
.leg-row:last-child {{ border-bottom: none; }}
.leg-ico {{ font-size: 13px; width: 20px; }}
.leg-nama {{ flex: 1; color: #555; }}
.leg-total {{ font-weight: bold; color: #2c3e50; font-size: 11px; }}
table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
th {{ background: #2980b9; color: white; padding: 7px 4px;
      text-align: center; font-size: 11px; font-weight: 600; }}
td {{ border: 1px solid #e4eaf0; vertical-align: top;
      height: var(--tinggi-sel, 72px); width: 14.28%; padding: 3px 2px; }}
td.kosong {{ background: #f8fafc; border-color: #eef2f7; }}
td.libur {{ background: #fff0f0; }}
td.today {{ background: #ebf5ff; outline: 2px solid #3498db; }}
.tgl-num {{ font-size: 18px; font-weight: bold; color: #2c3e50;
             line-height: 1; padding: 1px 0 0 2px; }}
td.libur .tgl-num {{ color: #e74c3c; }}
td.today .tgl-num {{ color: #2980b9; }}
.ikon-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr));
               gap: 3px 2px; margin-top: 4px; }}
.ic {{ display: flex; flex-direction: column; align-items: center; gap: 1px;
        cursor: pointer; border-radius: 4px; padding: 2px 0;
        transition: background 0.15s; user-select: none; }}
.ic:hover {{ background: #eef6ff; }}
.ic:active {{ background: #d6eaf8; }}
.ic .ico {{ font-size: 15px; line-height: 1.1; }}
.ic .jml {{ font-size: 8px; color: #2471a3; line-height: 1.1; min-height: 9px;
             text-decoration: underline dotted; }}
@media (min-width: 520px) {{
  td {{ height: var(--tinggi-lebar, 72px); padding: 4px 5px; }}
  .tgl-num {{ font-size: 22px; }}
  .ic {{ flex-direction: row; justify-content: flex-start; gap: 4px; padding: 2px 3px; }}
  .ic .jml {{ font-size: 10px; }}
}}
.rekap-section {{ margin-top: 12px; padding: 9px 12px;
                   background: #f0f7ff; border-radius: 8px;
                   font-size: 11px; display: flex; gap: 10px;
                   flex-wrap: wrap; align-items: center;
                   border: 1px solid #d6eaf8; }}
.rekap-link {{ color: #2471a3; font-weight: 500; cursor: pointer;
                padding: 3px 8px; background: white; border-radius: 6px;
                border: 1px solid #aed6f1; font-size: 11px; }}
.rekap-link:hover {{ background: #d6eaf8; }}
.rekap-buka {{ text-decoration: underline dotted; }}
.footer {{ text-align: center; font-size: 10px; color: #aaa;
            padding: 10px; border-top: 1px solid #eee; margin-top: 12px; }}
.toast {{ position: fixed; bottom: 20px; left: 50%;
           transform: translateX(-50%);
           background: #2ecc71; color: white; padding: 8px 18px;
           border-radius: 20px; font-size: 12px; display: none;
           box-shadow: 0 2px 8px rgba(0,0,0,0.2); z-index: 99;
           white-space: nowrap; }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>{NAMA_PUSKESMAS}</h1>
    {nav_html}
  </div>
  <div class="body">
    <table style="--tinggi-sel:{tinggi_sempit}px;--tinggi-lebar:{tinggi_lebar}px">
      <thead><tr>{header_hari}</tr></thead>
      <tbody>{baris_kalender}</tbody>
    </table>
    {rekap_html}
    {legenda_bawah}
    <div class="footer">Diperbarui: {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
  </div>
</div>
<div class="toast" id="toast"></div>
<script>
function tampilToast(pesan) {{
  var t = document.getElementById('toast');
  t.textContent = pesan;
  t.style.display = 'block';
  setTimeout(function() {{ t.style.display = 'none'; }}, 2200);
}}
function salinTeks(teks, label) {{
  var baris = teks.replace(/\\\\n/g, '\\n');
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(baris).then(function() {{
      tampilToast('✓ ' + label + ' disalin');
    }}).catch(function() {{ fallbackSalin(baris, label); }});
  }} else {{
    fallbackSalin(baris, label);
  }}
}}
function fallbackSalin(teks, label) {{
  var el = document.createElement('textarea');
  el.value = teks;
  el.style.position = 'fixed';
  el.style.opacity = '0';
  document.body.appendChild(el);
  el.select();
  try {{
    document.execCommand('copy');
    tampilToast('✓ ' + label + ' disalin');
  }} catch(e) {{
    tampilToast('Gagal salin — buka manual');
  }}
  document.body.removeChild(el);
}}
</script>
</body>
</html>"""

    # Simpan — terenkripsi, salinLink ada di wrapper
    out_path = PATH_OUTPUT / bulan_str / "dashboard.html"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    halaman = buat_halaman_terenkripsi(f"Dashboard {nama_bulan}", html, PASSWORD)
    out_path.write_text(halaman, encoding="utf-8")
    log.info(f"Dashboard disimpan: {out_path}")
    return out_path


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════
def cek_last_run() -> bool:
    """Return True kalau script sudah jalan hari ini."""
    if not LAST_RUN_FILE.exists():
        return False
    try:
        isi = LAST_RUN_FILE.read_text().strip()
        return isi == date.today().isoformat()
    except:
        return False


def catat_last_run():
    """Tulis tanggal hari ini ke last_run.txt."""
    LAST_RUN_FILE.write_text(date.today().isoformat())


def main(startup_mode: bool = False, tgl_dari: date = None, tgl_sampai: date = None, paksa: bool = False, jenis_filter: str = None):
    # ── Startup mode: cek sudah jalan hari ini ───────────────────
    if startup_mode:
        if cek_last_run():
            log.info("Startup mode: sudah jalan hari ini, skip.")
            return
        log.info(f"Startup mode: delay {STARTUP_DELAY} menit...")
        time.sleep(STARTUP_DELAY * 60)

    log.info("=" * 60)
    log.info(f"generate.py v{VERSI}")
    log.info(f"Mulai generate — {datetime.now().strftime('%A, %d %B %Y %H:%M')}")
    if startup_mode:
        log.info("Mode: STARTUP (data kemarin)")
    if tgl_dari:
        log.info(f"Mode: RENTANG {tgl_dari} s.d. {tgl_sampai or tgl_dari}")
    if jenis_filter:
        log.info(f"Filter jenis: {jenis_filter}")
    log.info("=" * 60)

    hari_ini = date.today()
    tgl_proses = hari_ini - timedelta(days=1) if startup_mode else hari_ini

    berhasil_reg = 0
    berhasil_dist = 0
    berhasil_pen = 0

    # ── Cek file Excel terbuka ───────────────────────────────────
    path_reg = path_register(tgl_proses.year)
    if path_reg.exists():
        if not tunggu_excel(path_reg):
            return
    if PATH_MUTASI.exists():
        if not tunggu_excel(PATH_MUTASI):
            return

    # ── Tentukan tanggal yang perlu diproses ────────────────────
    if tgl_dari:
        # Mode rentang — proses semua tanggal dalam range
        tgl_akhir_range = tgl_sampai or tgl_dari
        tanggal_reg = [
            tgl_dari + timedelta(days=i)
            for i in range((tgl_akhir_range - tgl_dari).days + 1)
            if hari_kerja(tgl_dari + timedelta(days=i))
        ]
        tanggal_mut = [
            tgl_dari + timedelta(days=i)
            for i in range((tgl_akhir_range - tgl_dari).days + 1)
        ]
    else:
        # Mode normal — lookback dari tgl_proses
        tanggal_reg = []
        i = 0
        while len(tanggal_reg) < LOOKBACK_REG and i < 14:
            tgl = tgl_proses - timedelta(days=i)
            if hari_kerja(tgl):
                tanggal_reg.append(tgl)
            i += 1
        tanggal_mut = [tgl_proses - timedelta(days=i) for i in range(LOOKBACK_MUT)]

    # ── Baca data (dengan retry untuk OneDrive) ──────────────────
    perlu_register = jenis_filter in (None, "register")
    perlu_mutasi   = jenis_filter in (None, "distribusi", "penerimaan")

    data_reg = {}
    data_mut = {}

    if perlu_register:
        log.info(f"Baca Register untuk: {[str(t) for t in tanggal_reg]}")
        try:
            data_reg = baca_register(tanggal_reg)
        except Exception as e:
            log.error(f"Gagal baca Register: {e}")
            kirim_telegram(f"🚨 Gagal baca Register:\n{e}", penting=True)

    if perlu_mutasi:
        log.info(f"Baca Mutasi untuk: {[str(t) for t in tanggal_mut]}")
        try:
            data_mut = baca_mutasi(tanggal_mut)
        except Exception as e:
            log.error(f"Gagal baca Mutasi: {e}")
            kirim_telegram(f"🚨 Gagal baca Mutasi:\n{e}", penting=True)

    # ── Generate Register ────────────────────────────────────────
    url_reg  = {}
    url_dist = {}
    url_pen  = {}
    pdf_paths = []

    if jenis_filter in (None, "register"):
        for tgl, rows in data_reg.items():
            try:
                path = path_harian("register", tgl)
                html_web, html_pdf, meta = html_register(tgl, rows, NAMA_PUSKESMAS, NAMA, NIP)
                simpan_html(path, f"Register {fmt_tgl(tgl)}", html_web, metadata=meta)
                pdf_path = simpan_pdf_ekinerja(html_pdf, "register", tgl)
                if pdf_path:
                    pdf_paths.append(("register", tgl, Path(pdf_path)))
                berhasil_reg += 1
            except Exception as e:
                log.error(f"Error generate Register {tgl}: {e}")
                kirim_telegram(f"🚨 Error generate Register {tgl}:\n{e}", penting=True)

    # ── Generate Distribusi & Penerimaan ─────────────────────────
    if jenis_filter in (None, "distribusi", "penerimaan"):
        for tgl, data in data_mut.items():
            try:
                dist = data.get("distribusi", [])
                pen  = data.get("penerimaan", [])

                if dist:
                    path = path_harian("distribusi", tgl)
                    if not paksa and path.with_suffix(".json").exists():
                        try:
                            import json as _json
                            meta_cek = _json.loads(path.with_suffix(".json").read_text(encoding="utf-8"))
                            if meta_cek.get("gdrive_url"):
                                url_dist[tgl] = meta_cek["gdrive_url"]
                                log.info(f"Skip distribusi {tgl} — link sudah ada")
                                berhasil_dist += 1
                                dist = []
                        except:
                            pass
                    if dist:
                        html_web, html_pdf, meta = html_mutasi(tgl, dist, "distribusi", NAMA_PUSKESMAS, NAMA, NIP)
                        simpan_html(path, f"Distribusi {fmt_tgl(tgl)}", html_web, metadata=meta)
                        pdf_path = simpan_pdf_ekinerja(html_pdf, "distribusi", tgl)
                        if pdf_path:
                            pdf_paths.append(("distribusi", tgl, Path(pdf_path)))
                        berhasil_dist += 1

                if pen:
                    path = path_harian("penerimaan", tgl)
                    if not paksa and path.with_suffix(".json").exists():
                        try:
                            import json as _json
                            meta_cek = _json.loads(path.with_suffix(".json").read_text(encoding="utf-8"))
                            if meta_cek.get("gdrive_url"):
                                url_pen[tgl] = meta_cek["gdrive_url"]
                                log.info(f"Skip penerimaan {tgl} — link sudah ada")
                                berhasil_pen += 1
                                pen = []
                        except:
                            pass
                    if pen:
                        html_web, html_pdf, meta = html_mutasi(tgl, pen, "penerimaan", NAMA_PUSKESMAS, NAMA, NIP)
                        simpan_html(path, f"Penerimaan {fmt_tgl(tgl)}", html_web, metadata=meta)
                        pdf_path = simpan_pdf_ekinerja(html_pdf, "penerimaan", tgl)
                        if pdf_path:
                            pdf_paths.append(("penerimaan", tgl, Path(pdf_path)))
                        berhasil_pen += 1

            except Exception as e:
                log.error(f"Error generate Mutasi {tgl}: {e}")
                kirim_telegram(f"🚨 Error generate Mutasi {tgl}:\n{e}", penting=True)

    # ── Ambil link GDrive setelah semua PDF selesai ──────────────
    if pdf_paths:
        log.info("Tunggu sync Google Drive...")
        time.sleep(10)
        import json as _json
        for jenis_pdf, tgl, path_pdf in pdf_paths:
            url = gdrive_link_dari_path(path_pdf)
            if url:
                if jenis_pdf == "register":   url_reg[tgl]  = url
                elif jenis_pdf == "distribusi": url_dist[tgl] = url
                elif jenis_pdf == "penerimaan": url_pen[tgl]  = url
                # Update metadata JSON supaya dashboard bisa baca link GDrive
                path_html = path_harian(jenis_pdf, tgl)
                meta_path = path_html.with_suffix(".json")
                if meta_path.exists():
                    try:
                        meta = _json.loads(meta_path.read_text(encoding="utf-8"))
                        meta["gdrive_url"] = url
                        meta_path.write_text(_json.dumps(meta, ensure_ascii=False), encoding="utf-8")
                    except:
                        pass
            else:
                log.warning(f"Link GDrive tidak ditemukan untuk {jenis_pdf}/{path_pdf.name} — mungkin belum sync")

    # ── Cek akhir bulan → rekap bulanan ─────────────────────────
    import calendar
    hari_terakhir = calendar.monthrange(hari_ini.year, hari_ini.month)[1]
    if hari_ini.day == hari_terakhir:
        log.info("Akhir bulan terdeteksi — generate rekap bulanan...")
        try:
            generate_rekap_bulanan(hari_ini.year, hari_ini.month)
            time.sleep(10)   # beri waktu Google Drive menyinkronkan PDF rekap
        except Exception as e:
            log.error(f"Error rekap bulanan: {e}")
            kirim_telegram(f"🚨 Error rekap bulanan {hari_ini.month}/{hari_ini.year}:\n{e}", penting=True)

    # ── Generate Dashboard ───────────────────────────────────────
    try:
        generate_dashboard(hari_ini.year, hari_ini.month)
    except Exception as e:
        log.error(f"Error generate dashboard: {e}")
        kirim_telegram(f"🚨 Error generate dashboard:\n{e}", penting=True)

    # ── Push ke GitHub ───────────────────────────────────────────
    push_ok = push_github()

    # ── Ringkasan ─────────────────────────────────────────────────
    base_url  = f"https://{GITHUB_USERNAME}.github.io/{GITHUB_REPO}"
    bulan_str  = hari_ini.strftime("%Y-%m")
    dash_url   = f"{base_url}/{AWALAN_URL}{bulan_str}/dashboard.html#{PASSWORD}"

    # Bangun link — hanya dari GDrive, skip kalau belum ada
    link_reg_list  = []
    link_dist_list = []
    link_pen_list  = []

    for tgl in sorted(url_reg.keys(), reverse=True):
        t = tgl.strftime("%Y-%m-%d")
        link_reg_list.append(f"{t}: {url_reg[tgl]}")

    for tgl in sorted(url_dist.keys(), reverse=True):
        t = tgl.strftime("%Y-%m-%d")
        link_dist_list.append(f"{t}: {url_dist[tgl]}")

    for tgl in sorted(url_pen.keys(), reverse=True):
        t = tgl.strftime("%Y-%m-%d")
        link_pen_list.append(f"{t}: {url_pen[tgl]}")

    link_bagian = f"\n\n📅 Dashboard: {dash_url}"
    if link_reg_list:
        link_bagian += "\n\n🔗 Register:\n" + "\n".join(link_reg_list)
    if link_dist_list:
        link_bagian += "\n\n🔗 Distribusi:\n" + "\n".join(link_dist_list)
    if link_pen_list:
        link_bagian += "\n\n🔗 Penerimaan:\n" + "\n".join(link_pen_list)

    # Tandai ada yang gagal
    ada_error = (
        (len(tanggal_reg) > 0 and berhasil_reg == 0) or
        (berhasil_reg < len(data_reg))
    )
    status_reg  = f"{'⚠️' if len(data_reg)==0 and len(tanggal_reg)>0 else '✅'} {berhasil_reg} file"
    status_dist = f"✅ {berhasil_dist} file"
    status_pen  = f"✅ {berhasil_pen} file"

    ringkasan = (
        f"{'⚠️' if ada_error else '✅'} Generate selesai — {hari_ini.strftime('%d/%m/%Y')}\n"
        f"Register: {status_reg}\n"
        f"Distribusi: {status_dist}\n"
        f"Penerimaan: {status_pen}\n"
        f"GitHub: {'✅ berhasil' if push_ok else '❌ gagal'}"
        f"{link_bagian}"
    )
    log.info(ringkasan.replace("*", "").replace("✅", "OK").replace("⚠️", "WARN").replace("❌", "GAGAL"))
    kirim_telegram(ringkasan)

    # Catat last run
    catat_last_run()
    log.info("Selesai.")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate laporan harian puskesmas")
    parser.add_argument("--startup", action="store_true",
                        help="Mode startup: delay + proses data kemarin")
    parser.add_argument("--dari", type=str, default=None,
                        help="Tanggal mulai (YYYY-MM-DD), misal: 2026-09-01")
    parser.add_argument("--sampai", type=str, default=None,
                        help="Tanggal akhir (YYYY-MM-DD), default = --dari")
    parser.add_argument("--paksa", action="store_true",
                        help="Paksa generate ulang semua meskipun link GDrive sudah ada")
    parser.add_argument("--jenis", type=str, default=None,
                        help=f"Filter jenis: {', '.join(JENIS_FILE.keys())}")
    parser.add_argument("--rekap", type=str, default=None,
                        help="Buat rekap bulanan (PDF) untuk bulan tertentu YYYY-MM, lalu perbarui dashboard bulan itu")
    parser.add_argument("--link-saja", action="store_true",
                        help="Hanya ambil link GDrive dari file yang sudah ada + update dashboard, tanpa generate ulang")
    tambah_argumen_mode(parser)
    args = parser.parse_args()
    if P.uji:
        log.info(f"Mode: UJI | PATH_OUTPUT={PATH_OUTPUT} | PATH_REPO={PATH_REPO} | PATH_EKINERJA={PATH_EKINERJA}")

    # Mode --rekap YYYY-MM: buat rekap bulanan (PDF) bulan tsb + perbarui dashboard + push
    if args.rekap:
        log.info(f"generate.py v{VERSI} — mode REKAP {args.rekap}")
        th, bl = (int(x) for x in args.rekap.split("-"))
        generate_rekap_bulanan(th, bl)
        log.info("Tunggu sync Google Drive...")
        time.sleep(10)
        generate_dashboard(th, bl)
        push_github()
        sys.exit(0)

    # Mode --link-saja: scan GDrive, update JSON, regenerate dashboard
    if args.link_saja:
        log.info("=" * 60)
        log.info(f"generate.py v{VERSI} — mode LINK SAJA")
        log.info("=" * 60)
        hari_ini  = date.today()
        tahun     = hari_ini.year
        bulan     = hari_ini.month
        bulan_str = hari_ini.strftime("%Y-%m")
        service   = get_gdrive_service()
        diperbarui = 0

        if service:
            bulan_dir = PATH_OUTPUT / bulan_str
            for jenis in ("register", "distribusi", "penerimaan"):
                jenis_dir = bulan_dir / jenis
                if not jenis_dir.exists():
                    continue
                for jf in jenis_dir.glob("*.json"):
                    try:
                        import json as _json_ls
                        meta = _json_ls.loads(jf.read_text(encoding="utf-8"))
                        if meta.get("gdrive_url"):
                            continue  # sudah ada, skip
                        # Cari PDF di Ekinerja
                        tgl_str  = jf.stem
                        nama_pdf = f"{AWALAN_FILE}{jenis}-{tgl_str}.pdf"
                        if PATH_EKINERJA:
                            nama_bulan_folder = f"{bulan:02d} {BULAN_ID[bulan]}"
                            path_pdf = PATH_EKINERJA / str(tahun) / nama_bulan_folder / jenis / nama_pdf
                            url = gdrive_link_dari_path(path_pdf)
                            if url:
                                meta["gdrive_url"] = url
                                jf.write_text(_json_ls.dumps(meta, ensure_ascii=False), encoding="utf-8")
                                log.info(f"Link diperbarui: {jenis}/{tgl_str}")
                                diperbarui += 1
                    except Exception as e:
                        log.warning(f"Gagal update {jf.name}: {e}")

        log.info(f"Selesai. {diperbarui} file diperbarui.")
        try:
            out = generate_dashboard(tahun, bulan)
            log.info(f"Dashboard disimpan: {out}")
            push_github()
        except Exception as e:
            log.error(f"Error update dashboard: {e}")
        sys.exit(0)

    # Normalisasi --jenis: cocokkan ke kunci JENIS_FILE secara case-insensitive
    jenis_filter = None
    if args.jenis:
        args_jenis_lower = args.jenis.lower()
        for k in JENIS_FILE:
            if k.lower() == args_jenis_lower:
                jenis_filter = k
                break
        if jenis_filter is None:
            print(f"Jenis tidak dikenal: {args.jenis}")
            print(f"Pilihan: {', '.join(JENIS_FILE.keys())}")
            sys.exit(1)

    if args.dari:
        tgl_dari   = date.fromisoformat(args.dari)
        tgl_sampai = date.fromisoformat(args.sampai) if args.sampai else tgl_dari
        main(startup_mode=False, tgl_dari=tgl_dari, tgl_sampai=tgl_sampai, paksa=args.paksa, jenis_filter=jenis_filter)
    else:
        main(startup_mode=args.startup, paksa=args.paksa, jenis_filter=jenis_filter)