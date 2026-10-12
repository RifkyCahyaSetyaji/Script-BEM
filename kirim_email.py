"""
=======================================================
  Script Kirim Email SMTP ke Banyak Penerima Sekaligus
  Author  : Prototype AI - SIBK
  Created : Oktober 2026
=======================================================

Fitur:
  - Kirim ke banyak penerima (daftar dari file atau langsung)
  - Baca daftar penerima dari CSV (.csv) atau Excel (.xlsx / .xls)
  - Mendukung format pesan HTML dan teks biasa
  - Lampiran file dokumen (PDF, DOCX, XLSX, PNG, dll.)
  - Log hasil pengiriman per penerima ke file TXT & JSON
  - Dukungan Gmail, Outlook/Hotmail, Yahoo, atau SMTP custom

Kolom wajib di file CSV/XLSX:
  - nama  : Nama lengkap penerima
  - email : Alamat email penerima
"""

import csv
import smtplib
import json
import os
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

try:
    import pandas as pd
    PANDAS_TERSEDIA = True
except ImportError:
    PANDAS_TERSEDIA = False
    # CSV didukung langsung menggunakan modul bawaan 'csv' tanpa perlu pandas.
    # Pandas hanya dibutuhkan jika membaca file Excel (.xlsx / .xls).

# ─────────────────────────────────────────────────────
# 1. KONFIGURASI SMTP (Isi sesuai akun pengirim Anda)
# ─────────────────────────────────────────────────────

SMTP_CONFIG = {
    # ─── Gmail (default) ────────────────────────────
    "host"         : "smtp.gmail.com",
    "port"         : 587,            # TLS
    # "port"       : 465,            # SSL — ganti smtplib.SMTP ke SMTP_SSL

    # ─── Outlook / Hotmail ─── (uncomment jika pakai)
    # "host"       : "smtp.office365.com",
    # "port"       : 587,

    # ─── Yahoo ─────────────── (uncomment jika pakai)
    # "host"       : "smtp.mail.yahoo.com",
    # "port"       : 587,

    # ─── SMTP Server Kampus / Kustom ─────────────────
    # "host"       : "mail.kampus.ac.id",
    # "port"       : 587,

    "email"        : "rifky_setyaji@student.ub.ac.id",   # ← Ganti email pengirim
    "password"     : "groe qclf mcgy gclm",       # ← App Password (bukan password utama!)
    "display_name" : "SGE FILKOM UB 2026"         # ← Nama tampil pengirim
}

# ─────────────────────────────────────────────────────
# 2. DAFTAR PENERIMA & LAMPIRAN
#    Semua data (nama, email, path dokumen) dibaca dari penerima.csv
#
#    Format kolom CSV:
#      nama     : Nama lengkap penerima   (WAJIB)
#      email    : Alamat email penerima   (WAJIB)
#      lampiran : Path absolut file dokumen yang dilampirkan per penerima (OPSIONAL)
# ─────────────────────────────────────────────────────

FILE_CSV_PENERIMA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "penerima.csv")

# DAFTAR_PENERIMA akan diisi saat program berjalan dari CSV
# (dipanggil di bagian bawah setelah fungsi load_penerima_dari_file didefinisikan)

# ─────────────────────────────────────────────────────
# 3. ISI EMAIL
# ─────────────────────────────────────────────────────

SUBJEK = "📢 TRANSPARANSI GRAND RECRUITMENT JUNIOR STAFF SGE FILKOM UB 2026"

# Template HTML — gunakan {nama} dan {link_dokumen} sebagai placeholder
BODY_HTML = """
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #2d3748; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px;">
  <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 24px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
    <h2 style="color: #2e7d32; margin-top: 0; font-size: 18px; border-bottom: 2px solid #e8f5e9; padding-bottom: 10px;">
      📢 TRANSPARANSI GRAND RECRUITMENT JUNIOR STAFF SGE FILKOM UB 2026
    </h2>

    <p style="font-size: 15px;">Halo, Sobat SGE! 👋🏻</p>

    <p>
      Berikut hasil transparansi Grand Recruitment Junior Staff SGE FILKOM UB 2026 atas nama <strong>{nama}</strong> yang dapat diakses melalui tautan berikut:
    </p>

    {link_dokumen}

    <p>
      Terima kasih atas partisipasi dan antusiasme seluruh peserta dalam rangkaian rekrutmen Junior Staff SGE FILKOM UB 2026. Transparansi ini merupakan bentuk keterbukaan dan pertanggungjawaban dalam proses rekrutmen yang telah dilaksanakan.
    </p>

    <br>
    <p style="margin-bottom: 4px;">Salam hangat,</p>
    <p style="margin: 0; font-weight: bold; color: #2e7d32; font-size: 15px;">SGE FILKOM UB 2026 💚</p>
    <hr style="border: none; border-top: 1px solid #edf2f7; margin-top: 24px;">
    <p style="font-size: 12px; color: #a0aec0; margin: 0;">Email ini dikirim otomatis. Harap tidak membalas langsung ke alamat ini.</p>
  </div>
</body>
</html>
"""

# Versi teks biasa (fallback jika client email tidak support HTML)
BODY_TEXT = """\
📢 TRANSPARANSI GRAND RECRUITMENT JUNIOR STAFF SGE FILKOM UB 2026

Halo, Sobat SGE! 👋🏻

Berikut hasil transparansi Grand Recruitment Junior Staff SGE FILKOM UB 2026 atas nama {nama} yang dapat diakses melalui tautan berikut:

{link_dokumen_text}

Terima kasih atas partisipasi dan antusiasme seluruh peserta dalam rangkaian rekrutmen Junior Staff SGE FILKOM UB 2026. Transparansi ini merupakan bentuk keterbukaan dan pertanggungjawaban dalam proses rekrutmen yang telah dilaksanakan.

Salam hangat,
SGE FILKOM UB 2026 💚
"""

# ─────────────────────────────────────────────────────
# 4. LAMPIRAN GLOBAL (OPSIONAL)
#    Jika ada file yang ingin dilampirkan ke SEMUA penerima,
#    daftarkan di sini. Lampiran per-penerima diatur lewat
#    kolom 'lampiran' di penerima.csv.
# ─────────────────────────────────────────────────────

# Direktori tempat script ini berada (untuk referensi path relatif)
_DIR = os.path.dirname(os.path.abspath(__file__))

FILE_LAMPIRAN = [
    # Kosong = lampiran diambil sepenuhnya dari kolom 'lampiran' di CSV
    # Tambahkan di sini jika ingin ada lampiran yang dikirim ke SEMUA penerima:
    # os.path.join(_DIR, "dokumen", "brosur.pdf"),
]

# ═══════════════════════════════════════════════════════
#              KODE UTAMA — TIDAK PERLU DIUBAH
# ═══════════════════════════════════════════════════════

def setup_logging():
    """Inisialisasi logger ke konsol dan file log."""
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(
        log_dir,
        f"email_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    )
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8")
        ]
    )
    logging.info(f"Log disimpan ke: {log_file}")
    return log_file


def load_penerima_dari_file(filepath: str) -> list:
    """
    Muat daftar penerima dari file eksternal.
    Format yang didukung:
      - CSV  (.csv)        : Harus memiliki kolom 'nama' dan 'email'
      - Excel (.xlsx/.xls) : Harus memiliki kolom 'nama' dan 'email' di sheet pertama
      - JSON (.json)       : [{"nama": "...", "email": "..."}, ...]
      - TXT  (.txt)        : satu email per baris (nama dari prefix sebelum @)

    Dependensi tambahan untuk CSV/XLSX:
      pip install pandas openpyxl
    """
    penerima = []
    ext = os.path.splitext(filepath)[1].lower()

    # ─── CSV ──────────────────────────────────────────────────────────────
    if ext == ".csv":
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers:
                raise ValueError(f"File CSV '{filepath}' kosong.")
            headers = [h.strip().lower() for h in headers]
            if "email" not in headers:
                raise ValueError(
                    f"Kolom 'email' tidak ditemukan di '{filepath}'. "
                    f"Kolom tersedia: {headers}"
                )
            email_idx = headers.index("email")
            nama_idx = headers.index("nama") if "nama" in headers else -1
            lampiran_idx = headers.index("lampiran") if "lampiran" in headers else -1

            for row in reader:
                if not row:
                    continue
                email = row[email_idx].strip() if email_idx < len(row) else ""
                if not email:
                    continue
                if nama_idx != -1 and nama_idx < len(row) and row[nama_idx].strip():
                    nama = row[nama_idx].strip()
                else:
                    nama = email.split("@")[0].replace(".", " ").title()
                lampiran = row[lampiran_idx].strip() if (lampiran_idx != -1 and lampiran_idx < len(row)) else ""
                penerima.append({
                    "nama": nama,
                    "email": email,
                    "lampiran": lampiran
                })

    # ─── XLSX / XLS ───────────────────────────────────────────────────────
    elif ext in (".xlsx", ".xls"):
        if not PANDAS_TERSEDIA:
            raise ImportError(
                "pandas + openpyxl diperlukan untuk baca Excel. "
                "Jalankan: pip install pandas openpyxl"
            )
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        df = pd.read_excel(filepath, engine=engine, dtype=str).fillna("")
        df.columns = [c.strip().lower() for c in df.columns]
        if "email" not in df.columns:
            raise ValueError(
                f"Kolom 'email' tidak ditemukan di '{filepath}'. "
                f"Kolom tersedia: {list(df.columns)}"
            )
        if "nama" not in df.columns:
            logging.warning("Kolom 'nama' tidak ditemukan, menggunakan prefix email sebagai nama.")
            df["nama"] = df["email"].apply(
                lambda e: e.split("@")[0].replace(".", " ").title()
            )
        cols = ["nama", "email"]
        if "lampiran" in df.columns:
            cols.append("lampiran")
        penerima = df[cols].to_dict(orient="records")

    # ─── JSON ─────────────────────────────────────────────────────────────
    elif ext == ".json":
        with open(filepath, "r", encoding="utf-8") as f:
            penerima = json.load(f)

    # ─── TXT (satu email per baris) ───────────────────────────────────────
    elif ext == ".txt":
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                email = line.strip()
                if email and "@" in email:
                    penerima.append({
                        "nama" : email.split("@")[0].replace(".", " ").title(),
                        "email": email
                    })
    else:
        raise ValueError(
            f"Format file tidak didukung: '{ext}'. "
            "Gunakan: .csv, .xlsx, .xls, .json, atau .txt"
        )

    # Filter baris dengan email kosong
    penerima = [p for p in penerima if str(p.get("email", "")).strip()]

    logging.info(f"Memuat {len(penerima)} penerima dari '{filepath}'")
    return penerima


def buat_pesan(penerima: dict) -> MIMEMultipart:
    """
    Susun objek email: headers, HTML body, teks fallback, dan lampiran/link dokumen.
    Lampiran/link diambil dari:
      1. Kolom 'lampiran' di dict penerima (URL web/Google Docs atau path file lokal)
      2. Daftar global FILE_LAMPIRAN (file lokal untuk semua penerima)
    """
    msg = MIMEMultipart("alternative")
    msg["Subject"] = SUBJEK
    msg["From"]    = f"{SMTP_CONFIG['display_name']} <{SMTP_CONFIG['email']}>"
    msg["To"]      = penerima["email"]

    nama = penerima.get("nama", penerima["email"])

    # ── Kumpulkan & pisahkan lampiran (URL Google Docs vs file fisik lokal) ──
    raw_lampiran = str(penerima.get("lampiran", "")).strip()
    urls = []
    file_lokal_penerima = []

    if raw_lampiran:
        # Mendukung multiple nilai dipisah tanda titik koma ';'
        for item in raw_lampiran.split(";"):
            item = item.strip()
            if not item:
                continue
            if item.startswith("http://") or item.startswith("https://"):
                urls.append(item)
            else:
                file_lokal_penerima.append(item)

    # ── Format tombol / link dokumen untuk HTML & Text ──
    if urls:
        tombol_html_list = []
        text_links = []
        for u in urls:
            tombol_html_list.append(f"""
  <div style="margin: 18px 0; padding: 16px 20px; background: #f1f8e9; border: 1px solid #c8e6c9; border-radius: 8px;">
    <p style="margin: 0 0 10px 0; font-weight: bold; color: #2e7d32; font-size: 14px;">🔗 Link Transparansi:</p>
    <a href="{u}" target="_blank" style="display: inline-block; padding: 10px 22px; background: #2e7d32; color: #ffffff; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 13px;">Buka Dokumen Transparansi</a>
    <p style="margin: 10px 0 0 0; font-size: 12px; color: #555;">
      Atau akses tautan langsung: <br>
      <a href="{u}" target="_blank" style="color: #2e7d32; word-break: break-all;">{u}</a>
    </p>
  </div>""")
            text_links.append(f"🔗 Link Transparansi: {u}")

        link_dokumen_html = "\n".join(tombol_html_list)
        link_dokumen_text = "\n" + "\n".join(text_links) + "\n"
    else:
        link_dokumen_html = ""
        link_dokumen_text = ""

    msg.attach(MIMEText(BODY_TEXT.format(nama=nama, link_dokumen_text=link_dokumen_text), "plain", "utf-8"))
    msg.attach(MIMEText(BODY_HTML.format(nama=nama, link_dokumen=link_dokumen_html), "html",  "utf-8"))

    # ── Lampirkan file lokal jika ada ──
    semua_file = FILE_LAMPIRAN + file_lokal_penerima
    for filepath in semua_file:
        if not os.path.isfile(filepath):
            logging.warning(f"    ⚠ Lampiran file tidak ditemukan, dilewati: {os.path.basename(filepath)}")
            continue
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename=\"{os.path.basename(filepath)}\""
        )
        msg.attach(part)
        logging.info(f"    📎 File Lampiran: {os.path.basename(filepath)}")

    if urls:
        for u in urls:
            logging.info(f"    🔗 Link Dokumen: {u}")

    return msg


def kirim_semua(daftar_penerima: list):
    """
    Hubungkan ke server SMTP sekali, lalu kirim email ke seluruh daftar penerima.
    Menggunakan STARTTLS (port 587). Untuk SSL (port 465), ganti ke smtplib.SMTP_SSL.
    """
    berhasil = []
    gagal    = []
    total    = len(daftar_penerima)

    logging.info(f"{'='*55}")
    logging.info(f"Mulai pengiriman ke {total} penerima...")
    logging.info(f"Server : {SMTP_CONFIG['host']}:{SMTP_CONFIG['port']}")
    logging.info(f"Dari   : {SMTP_CONFIG['email']}")
    logging.info(f"{'='*55}")

    try:
        # ── Buka koneksi & login SMTP sekali untuk semua penerima ──
        with smtplib.SMTP(SMTP_CONFIG["host"], SMTP_CONFIG["port"], timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SMTP_CONFIG["email"], SMTP_CONFIG["password"])
            logging.info("Login SMTP berhasil ✓")
            logging.info(f"{'-'*55}")

            for idx, penerima in enumerate(daftar_penerima, start=1):
                email_tujuan = penerima.get("email", "").strip()
                nama_tujuan  = penerima.get("nama", email_tujuan)

                # Validasi format email dasar
                if not email_tujuan or "@" not in email_tujuan:
                    logging.warning(
                        f"[{idx:>3}/{total}] ⚠ LEWAT   → Email tidak valid: '{email_tujuan}'"
                    )
                    gagal.append({"penerima": penerima, "error": "Format email tidak valid"})
                    continue

                try:
                    msg = buat_pesan(penerima)
                    server.sendmail(
                        SMTP_CONFIG["email"],
                        email_tujuan,
                        msg.as_string()
                    )
                    logging.info(f"[{idx:>3}/{total}] ✅ BERHASIL → {nama_tujuan} <{email_tujuan}>")
                    berhasil.append(email_tujuan)

                except smtplib.SMTPRecipientsRefused as e:
                    logging.error(f"[{idx:>3}/{total}] ❌ DITOLAK  → {email_tujuan} | {e}")
                    gagal.append({"penerima": penerima, "error": str(e)})
                except Exception as e:
                    logging.error(f"[{idx:>3}/{total}] ❌ GAGAL    → {email_tujuan} | {e}")
                    gagal.append({"penerima": penerima, "error": str(e)})

    except smtplib.SMTPAuthenticationError:
        logging.critical("❌ AUTENTIKASI GAGAL: Email atau App Password salah!")
        logging.critical("   Gmail: Aktifkan 2FA → buat App Password di:")
        logging.critical("   https://myaccount.google.com/apppasswords")
        return
    except (smtplib.SMTPConnectError, ConnectionRefusedError, TimeoutError) as e:
        logging.critical(f"❌ KONEKSI GAGAL: Tidak dapat terhubung ke server SMTP | {e}")
        return

    # ─── Ringkasan Akhir ───────────────────────────────
    logging.info(f"{'='*55}")
    logging.info(
        f"SELESAI | Total: {total} | "
        f"✅ Berhasil: {len(berhasil)} | ❌ Gagal: {len(gagal)}"
    )

    if gagal:
        logging.warning("Email yang GAGAL dikirim:")
        for item in gagal:
            logging.warning(f"  • {item['penerima'].get('email')} — {item['error']}")

    # Simpan ringkasan JSON
    ringkasan = {
        "timestamp"     : datetime.now().isoformat(),
        "total"         : total,
        "berhasil"      : berhasil,
        "gagal"         : gagal
    }
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    ringkasan_path = os.path.join(log_dir, "hasil_pengiriman.json")
    with open(ringkasan_path, "w", encoding="utf-8") as f:
        json.dump(ringkasan, f, indent=2, ensure_ascii=False)
    logging.info(f"Ringkasan tersimpan → {ringkasan_path}")


# ─────────────────────────────────────────────────────
# TITIK MASUK PROGRAM
# ─────────────────────────────────────────────────────
if __name__ == "__main__":
    setup_logging()

    # Muat semua data (nama, email, lampiran) dari penerima.csv
    if not os.path.isfile(FILE_CSV_PENERIMA):
        logging.critical(f"File CSV tidak ditemukan: {FILE_CSV_PENERIMA}")
        logging.critical("Pastikan file 'penerima.csv' ada di folder yang sama dengan script ini.")
        exit(1)

    DAFTAR_PENERIMA = load_penerima_dari_file(FILE_CSV_PENERIMA)

    if not DAFTAR_PENERIMA:
        logging.critical("Tidak ada penerima yang valid di CSV. Proses dihentikan.")
        exit(1)

    # Preview daftar penerima sebelum dikirim
    logging.info(f"Preview penerima ({len(DAFTAR_PENERIMA)} orang):")
    for idx, p in enumerate(DAFTAR_PENERIMA, 1):
        raw_dok = str(p.get("lampiran", "")).strip()
        if raw_dok.startswith("http://") or raw_dok.startswith("https://"):
            lampiran_info = f"[Link Docs] {raw_dok[:55]}..."
        elif raw_dok:
            lampiran_info = f"[File] {os.path.basename(raw_dok)}"
        else:
            lampiran_info = "(tidak ada)"
        logging.info(f"  {idx}. {p.get('nama','?')} <{p.get('email','?')}> | dok: {lampiran_info}")

    kirim_semua(DAFTAR_PENERIMA)
