"""
dashboard_actions.py — Perbarui link + dashboard TANPA PC (dijalankan GitHub Actions).

Yang dikerjakan (setara `generate.py --link-saja` + push, tapi lewat Google Drive API):
  1. Baca data harian (JSON/HTML) dari folder GitFiles di Drive dan daftar file Ekinerja di Drive.
  2. Isi `gdrive_url` yang masih kosong di JSON harian (lalu tulis balik ke Drive).
  3. Bangun ulang dashboard bulan itu dengan generate_dashboard() ASLI dari generate.py
     (supaya tampilannya selalu sama dengan yang dibuat di PC).
  4. Salin hasilnya ke <repo>/<YYYY-MM>/dashboard.html (commit/push dilakukan workflow).

Yang TIDAK dikerjakan di server: membuat PDF dari Excel (tetap di PC) dan menggabung
penerimaan-...-a/-b.pdf menjadi -lengkap.pdf (server hanya memakai -lengkap yang sudah ada).

Kebutuhan (lihat workflow): secret GDRIVE_SA_JSON, GDRIVE_EKINERJA_ID, GDRIVE_GITFILES_ID,
DASHBOARD_PASSWORD. File generate.py harus ada di folder yang sama dengan skrip ini.

Jalankan lokal (uji): python dashboard_actions.py --bulan 2026-10 --repo .
"""
VERSI = "2026-11-07.2"

import argparse
import importlib.util
import json
import logging
import os
import shutil
import sys
import tempfile
import types
from datetime import date, datetime
from pathlib import Path

log = logging.getLogger("dashboard_actions")

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}
JENIS_GITFILES = ("register", "distribusi", "penerimaan")
MIME_FOLDER = "application/vnd.google-apps.folder"


# ══════════════════════════════════════════════════════════════════
# GOOGLE DRIVE (service account)
# ══════════════════════════════════════════════════════════════════
class Drive:
    """Pembungkus tipis Google Drive API v3. Antarmuka ini yang ditiru FakeDrive saat uji."""

    def __init__(self, info_sa: dict):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        cred = service_account.Credentials.from_service_account_info(
            info_sa, scopes=["https://www.googleapis.com/auth/drive"])
        self.svc = build("drive", "v3", credentials=cred, cache_discovery=False)

    def anak(self, folder_id: str) -> list:
        """Isi folder: [{id, name, mimeType}, ...]"""
        hasil, token = [], None
        while True:
            r = self.svc.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                fields="nextPageToken, files(id, name, mimeType)",
                pageSize=1000, pageToken=token,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            hasil += r.get("files", [])
            token = r.get("nextPageToken")
            if not token:
                return hasil

    def baca(self, file_id: str) -> bytes:
        return self.svc.files().get_media(fileId=file_id, supportsAllDrives=True).execute()

    def tulis(self, file_id: str, data: bytes, mime: str = "application/json") -> None:
        from googleapiclient.http import MediaInMemoryUpload
        self.svc.files().update(
            fileId=file_id, media_body=MediaInMemoryUpload(data, mimetype=mime),
            supportsAllDrives=True,
        ).execute()


def peta_folder(drive, folder_id: str) -> dict:
    """{nama: id} untuk subfolder langsung di folder_id."""
    return {f["name"]: f["id"] for f in drive.anak(folder_id) if f["mimeType"] == MIME_FOLDER}


def link_drive(file_id: str) -> str:
    return f"https://drive.google.com/file/d/{file_id}/view"


# ══════════════════════════════════════════════════════════════════
# HARI LIBUR (server tidak punya salinan cache PC)
# ══════════════════════════════════════════════════════════════════
class Libur:
    """Pengganti bersama.hari: libur nasional dibaca dari libur.json (lihat petunjuk)."""

    def __init__(self, tabel: dict):
        self.tabel = tabel

    def libur(self, tahun: int) -> set:
        return self.tabel.get(tahun, set())

    def hari_kerja(self, tgl: date) -> bool:
        return tgl.weekday() != 6 and tgl not in self.libur(tgl.year)


def muat_libur(path: Path) -> dict:
    """libur.json: {"2026": ["2026-01-01", ...], "2027": [...]} atau daftar tanggal datar."""
    if not path.exists():
        log.warning(f"{path.name} tidak ada — hanya hari Minggu yang ditandai libur.")
        return {}
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    tanggal = []
    if isinstance(data, dict):
        for v in data.values():
            tanggal += v
    else:
        tanggal = list(data)
    tabel = {}
    for s in tanggal:
        d = date.fromisoformat(s)
        tabel.setdefault(d.year, set()).add(d)
    return tabel


# ══════════════════════════════════════════════════════════════════
# MEMUAT generate.py TANPA config PC (modul pengganti)
# ══════════════════════════════════════════════════════════════════
def _pasang_modul_pengganti(nilai: dict, path_output: Path, path_ekinerja: Path, libur: Libur):
    """Daftarkan modul `config` dan `bersama.*` tiruan supaya generate.py bisa diimpor di server."""

    class _Pengaturan:
        uji = False
        awalan_url = ""
        awalan_nama_file = ""
        gdrive_folder_id = ""
        path_mutasi = None
        path_repo = path_output
        dir_rahasia = path_output.parent

        def __init__(self):
            self.path_output = path_output
            self.path_ekinerja = path_ekinerja

        def ambil(self, bagian, kunci, wajib=True, default=None):
            v = nilai.get((bagian, kunci))
            if v is not None:
                return v
            return default if default is not None else ""

        def ambil_int(self, bagian, kunci, default=0, **kw):
            return default

        def penanda(self, nama):
            return self.dir_rahasia / nama

        def __getattr__(self, nama):
            return None

    cfg = types.ModuleType("config")
    cfg.PENGATURAN = _Pengaturan()

    bersama = types.ModuleType("bersama")
    bersama.__path__ = []
    m_hari = types.ModuleType("bersama.hari")
    m_hari.dari_pengaturan = lambda P, log=None: libur
    m_tele = types.ModuleType("bersama.telegram")
    m_tele.kirim_telegram = lambda *a, **k: None
    m_log = types.ModuleType("bersama.log")
    m_log.pasang_log = lambda P, nama: logging.getLogger(nama)
    m_peng = types.ModuleType("bersama.pengaturan")
    m_peng.tambah_argumen_mode = lambda parser: None
    bersama.hari, bersama.telegram, bersama.log, bersama.pengaturan = m_hari, m_tele, m_log, m_peng

    ditambah = {"config": cfg, "bersama": bersama, "bersama.hari": m_hari,
                "bersama.telegram": m_tele, "bersama.log": m_log, "bersama.pengaturan": m_peng}
    sebelumnya = {k: sys.modules.get(k) for k in ditambah}
    sys.modules.update(ditambah)
    return sebelumnya


def _lepas_modul_pengganti(sebelumnya: dict):
    for k, v in sebelumnya.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v


def muat_generate(path_generate: Path, nilai: dict, path_output: Path, path_ekinerja: Path, libur: Libur):
    sebelumnya = _pasang_modul_pengganti(nilai, path_output, path_ekinerja, libur)
    try:
        spec = importlib.util.spec_from_file_location("generate_pc", str(path_generate))
        g = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(g)
        return g
    finally:
        _lepas_modul_pengganti(sebelumnya)


# ══════════════════════════════════════════════════════════════════
# INTI
# ══════════════════════════════════════════════════════════════════
def jalankan(drive, tahun: int, bulan: int, *, id_ekinerja: str, id_gitfiles: str, password: str,
             repo_root: Path, path_generate: Path, path_libur: Path,
             nama_puskesmas: str = "Puskesmas Tambang Ulang",
             github_repository: str = "") -> dict:
    """Perbarui dashboard satu bulan. Return ringkasan {link_diperbarui, dashboard}."""
    bulan_str = f"{tahun}-{bulan:02d}"
    nama_bulan_folder = f"{bulan:02d} {BULAN_ID[bulan]}"
    pemilik, _, nama_repo = (github_repository or "pemilik/repo").partition("/")

    tmp = Path(tempfile.mkdtemp(prefix="dash_"))
    out_root, ek_root = tmp / "GitFiles", tmp / "Ekinerja"
    out_root.mkdir()
    ek_root.mkdir()
    try:
        nilai = {
            ("identitas", "NAMA_PUSKESMAS"): nama_puskesmas,
            ("keamanan", "PASSWORD"): password,
            ("github", "GITHUB_USERNAME"): pemilik,
            ("github", "GITHUB_REPO"): nama_repo,
        }
        libur = Libur(muat_libur(path_libur))
        g = muat_generate(path_generate, nilai, out_root, ek_root, libur)
        awalan = getattr(g, "AWALAN_FILE", "") or ""

        # ── Ekinerja: placeholder file + peta nama→link ─────────────
        link_peta, nama_folder_ek = {}, {}   # rel path → url ; (jenis) → {nama: id}
        f_tahun = peta_folder(drive, id_ekinerja).get(str(tahun))
        f_bulan = peta_folder(drive, f_tahun).get(nama_bulan_folder) if f_tahun else None
        if not f_bulan:
            log.warning(f"Folder Ekinerja/{tahun}/{nama_bulan_folder} tidak ditemukan di Drive.")
        else:
            sub = peta_folder(drive, f_bulan)
            for jenis in sorted(set(g.JENIS_FILE) | set(getattr(g, "JENIS_REKAP", {}))):
                jid = sub.get(jenis)
                if not jid:
                    continue
                folder_lokal = ek_root / str(tahun) / nama_bulan_folder / jenis
                folder_lokal.mkdir(parents=True, exist_ok=True)
                nama_folder_ek[jenis] = {}
                for f in drive.anak(jid):
                    if f["mimeType"] == MIME_FOLDER:
                        continue
                    (folder_lokal / f["name"]).touch()
                    rel = f"{tahun}/{nama_bulan_folder}/{jenis}/{f['name']}"
                    link_peta[rel] = link_drive(f["id"])
                    nama_folder_ek[jenis][f["name"]] = f["id"]

        # ── GitFiles: JSON harian (asli) + placeholder HTML ─────────
        json_id = {}   # (jenis, nama) → file id Drive
        f_gf = peta_folder(drive, id_gitfiles).get(bulan_str)
        if not f_gf:
            log.warning(f"Folder GitFiles/{bulan_str} tidak ditemukan di Drive — kalender akan kosong.")
        else:
            sub = peta_folder(drive, f_gf)
            for jenis in JENIS_GITFILES:
                jid = sub.get(jenis)
                if not jid:
                    continue
                folder_lokal = out_root / bulan_str / jenis
                folder_lokal.mkdir(parents=True, exist_ok=True)
                for f in drive.anak(jid):
                    if f["name"].endswith(".json"):
                        (folder_lokal / f["name"]).write_bytes(drive.baca(f["id"]))
                        json_id[(jenis, f["name"])] = f["id"]
                    elif f["name"].endswith(".html"):
                        (folder_lokal / f["name"]).touch()

        # ── Isi gdrive_url yang kosong ──────────────────────────────
        diperbarui = 0
        for (jenis, nama_json), fid in sorted(json_id.items()):
            jf = out_root / bulan_str / jenis / nama_json
            try:
                meta = json.loads(jf.read_text(encoding="utf-8"))
            except Exception as e:
                log.warning(f"JSON {jenis}/{nama_json} tidak terbaca: {e}")
                continue
            if meta.get("gdrive_url"):
                continue
            nama_pdf = f"{awalan}{jenis}-{Path(nama_json).stem}.pdf"
            pdf_id = nama_folder_ek.get(jenis, {}).get(nama_pdf)
            if not pdf_id:
                continue
            meta["gdrive_url"] = link_drive(pdf_id)
            data = json.dumps(meta, ensure_ascii=False).encode("utf-8")
            jf.write_bytes(data)
            drive.tulis(fid, data)
            diperbarui += 1
            log.info(f"Link diperbarui: {jenis}/{Path(nama_json).stem}")

        # ── Tombol bulan sebelumnya/berikutnya: dashboard yang sudah ada di repo ──
        for d in Path(repo_root).glob("*/dashboard.html"):
            if len(d.parent.name) == 7 and d.parent.name[4] == "-" and d.parent.name != bulan_str:
                (out_root / d.parent.name).mkdir(exist_ok=True)
                (out_root / d.parent.name / "dashboard.html").touch()

        # ── Ganti fungsi yang bergantung pada disk PC ──────────────
        def _link_dari_path(p):
            try:
                rel = Path(p).relative_to(ek_root).as_posix()
            except ValueError:
                return ""
            return link_peta.get(rel, "")

        def _lengkap(tgl):
            folder = ek_root / str(tgl.year) / f"{tgl.month:02d} {BULAN_ID[tgl.month]}" / "penerimaan"
            p = folder / f"{awalan}penerimaan-{tgl.strftime('%Y-%m-%d')}-lengkap.pdf"
            return p if p.exists() else None

        g.gdrive_link_dari_path = _link_dari_path
        g.gabung_penerimaan_lengkap = _lengkap

        # ── Bangun dashboard memakai kode asli ──────────────────────
        hasil = Path(g.generate_dashboard(tahun, bulan))
        tujuan = Path(repo_root) / bulan_str / "dashboard.html"
        tujuan.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(hasil, tujuan)
        log.info(f"Dashboard {bulan_str} ditulis: {tujuan} ({diperbarui} link diperbarui)")
        return {"link_diperbarui": diperbarui, "dashboard": tujuan}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════
def _wajib_env(nama: str) -> str:
    v = os.environ.get(nama, "").strip()
    if not v:
        sys.exit(f"Variabel/secret {nama} belum diisi.")
    return v


def main(argv=None):
    ap = argparse.ArgumentParser(description="Perbarui dashboard dari Google Drive (tanpa PC)")
    ap.add_argument("--bulan", default="", help="YYYY-MM (kosong = bulan ini)")
    ap.add_argument("--repo", default=".", help="folder checkout repo (berisi <YYYY-MM>/dashboard.html)")
    ap.add_argument("--generate", default=str(Path(__file__).with_name("generate.py")))
    ap.add_argument("--libur", default=str(Path(__file__).with_name("libur.json")))
    a = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    log.info(f"dashboard_actions.py v{VERSI}")

    if a.bulan.strip():
        try:
            tahun, bulan = (int(x) for x in a.bulan.strip().split("-"))
            date(tahun, bulan, 1)
        except Exception:
            sys.exit(f"Format --bulan salah: {a.bulan!r} (harus YYYY-MM)")
    else:
        h = date.today()
        tahun, bulan = h.year, h.month

    if not Path(a.generate).exists():
        sys.exit(f"generate.py tidak ditemukan di {a.generate}. Unggah salinannya ke folder server/.")
    try:
        info_sa = json.loads(_wajib_env("GDRIVE_SA_JSON"))
    except json.JSONDecodeError:
        sys.exit("GDRIVE_SA_JSON bukan JSON yang valid (salin ulang seluruh isi file kunci).")

    drive = Drive(info_sa)
    try:
        r = jalankan(
            drive, tahun, bulan,
            id_ekinerja=_wajib_env("GDRIVE_EKINERJA_ID"), id_gitfiles=_wajib_env("GDRIVE_GITFILES_ID"),
            password=_wajib_env("DASHBOARD_PASSWORD"), repo_root=Path(a.repo),
            path_generate=Path(a.generate), path_libur=Path(a.libur),
            nama_puskesmas=os.environ.get("NAMA_PUSKESMAS", "Puskesmas Tambang Ulang"),
            github_repository=os.environ.get("GITHUB_REPOSITORY", ""),
        )
    except Exception as e:
        if type(e).__name__ == "HttpError":
            sys.exit(f"Google Drive menolak permintaan: {e}\n"
                     f"Pastikan folder Ekinerja & GitFiles sudah dibagikan ke email service account (Editor).")
        raise
    print(f"Selesai: {r['link_diperbarui']} link diperbarui → {r['dashboard']}")


if __name__ == "__main__":
    main()