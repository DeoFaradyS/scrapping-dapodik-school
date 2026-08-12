"""
Scraper Dapodik: Provinsi -> Kab/Kota -> Kecamatan -> Sekolah -> Detail
Jenjang: SD, SMP, SMA, SMK | Status: Swasta
Prioritas: JANGAN ADA YANG MISS. Lambat gapapa, asal data lengkap.
Tiap level ada retry + kalo tetep gagal, dicatet & lanjut (gak crash total).

CARA PAKAI:
    pip install playwright openpyxl pandas
    playwright install chromium

    python scrape_full.py --provinsi "Gorontalo"
    python scrape_full.py --provinsi "Gorontalo,Bali"
    python scrape_full.py                              # semua provinsi, auto-skip yang selesai
    python scrape_full.py --ulang --provinsi "Gorontalo"
    python scrape_full.py --test                       # 1 provinsi, 1 kab, 1 kec doang
    python scrape_full.py --debug                       # browser kelihatan

OUTPUT:
    {NamaProvinsi}.xlsx    -> 1 file per provinsi, isi semua jenjang (SD/SMP/SMA/SMK)
    progress_selesai.txt   -> "Jenjang|Provinsi" yang udah kelar, run berikutnya auto-skip
"""

import asyncio
import os
import re
import sys
import pandas as pd
from playwright.async_api import async_playwright

BASE = "https://dapo.kemendikdasmen.go.id"
PROGRESS_FILE = "progress_selesai.txt"
ROW_SELECTOR = "tr.cursor-pointer"
LEVEL_NAMES = ["Kab/Kota", "Kecamatan", "Sekolah"]  # depth 1, 2, 3
JENJANG_LIST = ["SD", "SMP", "SMA", "SMK"]  # fix, semua jenjang selalu diambil

DEBUG = "--debug" in sys.argv
TEST = "--test" in sys.argv
ULANG = "--ulang" in sys.argv


def ambil_arg_list(flag, default):
    if flag not in sys.argv:
        return default
    idx = sys.argv.index(flag)
    if idx + 1 >= len(sys.argv):
        return default
    return [s.strip() for s in sys.argv[idx + 1].split(",")]


PROVINSI_FILTER = ambil_arg_list("--provinsi", None)
if PROVINSI_FILTER:
    PROVINSI_FILTER = [s.lower() for s in PROVINSI_FILTER]

STATUS_BLACKLIST = {
    "sangat baik", "baik", "cukup", "sedang", "kurang", "sangat kurang",
    "swasta", "negeri", "sudah", "belum", "ya", "tidak", "aktif", "nonaktif",
    "sudah sinkron", "belum sinkron", "sinkron", "belum sync", "sudah sync",
}


def start_url(jenjang):
    return f"{BASE}/progres?jenjang={jenjang}&status_sekolah=Swasta"


# ---------- progress & penyimpanan ----------

def baca_progress_selesai():
    if not os.path.exists(PROGRESS_FILE):
        return set()
    with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def tandai_selesai(key):
    with open(PROGRESS_FILE, "a", encoding="utf-8") as f:
        f.write(key + "\n")


def simpan_provinsi_incremental(nama_prov, rows_baru):
    """Simpan/gabung ke file excel per provinsi, misal 'Gorontalo.xlsx'."""
    if not rows_baru:
        return
    file_out = f"{nama_prov}.xlsx"
    df_baru = pd.DataFrame(rows_baru)
    if os.path.exists(file_out):
        try:
            df_lama = pd.read_excel(file_out)
            df_baru = pd.concat([df_lama, df_baru], ignore_index=True)
        except Exception:
            pass
    df_baru.to_excel(file_out, index=False)
    return file_out


# ---------- parsing nama dari row tabel ----------

def bersihkan_nama_file(nama):
    nama = re.sub(r'[\\/:*?"<>|\t\n\r]', "-", nama).strip().rstrip(". ")
    return nama[:150] if nama else "tanpa_nama"


def nama_dari_cols(cols, fallback):
    kandidat = [
        c for c in cols
        if c and not c.endswith("%")
        and not c.replace(".", "").replace(",", "").isdigit()
        and c.lower() not in STATUS_BLACKLIST
        and re.search(r"[A-Za-z]{2,}", c)
    ]
    if not kandidat:
        return fallback
    for c in kandidat:
        if re.match(r"^(Prov\.|Kab\.|Kota|Kec\.)", c, re.I):
            return c
    return max(kandidat, key=len)


async def get_rows_text(page):
    rows = await page.query_selector_all(ROW_SELECTOR)
    out = []
    for r in rows:
        txt = await r.inner_text()
        parts = [c.strip() for c in re.split(r"[\n\t]+", txt) if c.strip()]
        out.append((r, parts))
    return out


def extract_fields_from_text(text):
    def cari(label, pat=r"[^\n]{1,150}"):
        m = re.search(re.escape(label) + r"\s*[:\-]?\s*\n\s*(" + pat + r")", text, re.I)
        return m.group(1).strip() if m else ""

    npsn = re.search(r"NPSN\s*\n\s*(\d{6,10})", text, re.I)
    status = re.search(r"\bStatus\s*\n\s*(Swasta|Negeri)\b", text, re.I)
    akreditasi = re.search(r"Akreditasi\s*([A-Z])\b", text, re.I)
    alamat = re.search(r"Alamat\s*\n\s*([^\n]{5,200})", text, re.I)

    return {
        "Kepala Sekolah": cari("Kepala Sekolah"),
        "NPSN": npsn.group(1) if npsn else "",
        "Status": status.group(1) if status else "",
        "Bentuk Pendidikan": cari("Bentuk Pendidikan"),
        "Status Kepemilikan": cari("Status Kepemilikan"),
        "Peserta Didik": cari("Peserta Didik", r"\d+"),
        "PTK": cari("PTK", r"\d+"),
        "Rombel": cari("Rombel", r"\d+"),
        "Akreditasi": akreditasi.group(1) if akreditasi else "",
        "Alamat": alamat.group(1).strip() if alamat else "",
    }


# ---------- navigasi tahan banting (retry) ----------

async def goto_aman(page, url, tries=4, wait_until="networkidle", timeout=90000):
    for _ in range(tries):
        try:
            await page.goto(url, wait_until=wait_until, timeout=timeout)
            await page.wait_for_timeout(1200)
            return True
        except Exception:
            await page.wait_for_timeout(2000)
    print(f"        [!] goto gagal total: {url}")
    return False


async def klik_baris_aman(page, index, tries=4, timeout=45000):
    """Klik baris ke-`index` di tabel, re-query tiap percobaan biar gak stale."""
    for _ in range(tries):
        try:
            rows = await get_rows_text(page)
            if index >= len(rows):
                return False
            row_el, _ = rows[index]
            async with page.expect_navigation(wait_until="networkidle", timeout=timeout):
                await row_el.click()
            await page.wait_for_timeout(2000)
            return True
        except Exception:
            await page.wait_for_timeout(2000)
    print(f"        [!] klik baris {index} gagal total")
    return False


# ---------- crawl rekursif: 1 fungsi buat semua level ----------

async def crawl(page, jenjang, path, parent_url, depth, hasil):
    """path = [nama_provinsi, nama_kab, nama_kec, ...] sejauh ini.
    depth 1..3 = level Kab/Kota, Kecamatan, Sekolah. Abis klik sekolah (depth 3),
    kita udah di halaman detail -> extract & catat, gak rekursi lagi."""
    rows = await get_rows_text(page)
    jumlah = 1 if TEST else len(rows)
    label = LEVEL_NAMES[depth - 1]
    print(f"{'  ' * depth}{label} di {path[-1]}: {jumlah}")

    for i in range(jumlah):
        rows_now = await get_rows_text(page)
        if i >= len(rows_now):
            continue
        _, cols = rows_now[i]
        nama = bersihkan_nama_file(nama_dari_cols(cols, f"{label.lower()}_{i+1}"))
        path_baru = path + [nama]

        ok = await klik_baris_aman(page, i)
        if not ok:
            hasil.append(baris_gagal(jenjang, path_baru, f"GAGAL buka {label}"))
            await goto_aman(page, parent_url)
            continue
        url_sekarang = page.url

        if depth == 3:  # baru klik sekolah -> ini halaman detail
            await catat_sekolah(page, jenjang, path_baru, hasil)
        else:
            await crawl(page, jenjang, path_baru, url_sekarang, depth + 1, hasil)

        await goto_aman(page, parent_url)


def baris_gagal(jenjang, path, pesan):
    kolom = ["Provinsi", "Kab/Kota", "Kecamatan", "Nama Sekolah"]
    row = {"Jenjang": jenjang, "Status": pesan}
    for k, v in zip(kolom, path):
        row[k] = v
    return row


async def catat_sekolah(page, jenjang, path, hasil):
    try:
        h1 = await page.query_selector("h1")
        nama_asli = (await h1.inner_text()).strip() if h1 else ""
    except Exception:
        nama_asli = ""
    nama_sek = bersihkan_nama_file(nama_asli) if nama_asli else path[-1]

    try:
        text = await page.inner_text("body")
        field_data = extract_fields_from_text(text)
    except Exception as e:
        field_data = {}
        print(f"        [!] extract gagal buat {nama_sek}: {e}")

    print(f"        + {nama_sek}")
    hasil.append({
        "Jenjang": jenjang,
        "Provinsi": path[0],
        "Kab/Kota": path[1],
        "Kecamatan": path[2],
        "Nama Sekolah": nama_sek,
        "Status": "sukses",
        **field_data,
    })


# ---------- main ----------

async def main():
    selesai = baca_progress_selesai()
    if selesai and not ULANG:
        print(f"Udah selesai sebelumnya ({len(selesai)}): {', '.join(sorted(selesai))}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=not DEBUG)
        page = await browser.new_page()

        for jenjang in JENJANG_LIST:
            url_awal = start_url(jenjang)
            await goto_aman(page, url_awal)

            prov_rows = await get_rows_text(page)
            print(f"\n########## Jenjang {jenjang}: {len(prov_rows)} provinsi ##########")
            jumlah_prov = 1 if TEST else len(prov_rows)

            for pi in range(jumlah_prov):
                prov_rows = await get_rows_text(page)
                if pi >= len(prov_rows):
                    continue
                _, cols = prov_rows[pi]
                nama_prov = bersihkan_nama_file(nama_dari_cols(cols, f"provinsi_{pi+1}"))
                key = f"{jenjang}|{nama_prov}"

                if key in selesai and not ULANG:
                    print(f"\n=== {jenjang} - {nama_prov} -> SKIP (udah selesai) ===")
                    continue
                if PROVINSI_FILTER and not any(f in nama_prov.lower() for f in PROVINSI_FILTER):
                    continue

                print(f"\n=== {jenjang} - Provinsi {pi+1}/{len(prov_rows)}: {nama_prov} ===")
                hasil = []

                ok = await klik_baris_aman(page, pi)
                if not ok:
                    hasil.append(baris_gagal(jenjang, [nama_prov], "GAGAL buka provinsi"))
                    simpan_provinsi_incremental(nama_prov, hasil)
                    await goto_aman(page, url_awal)
                    continue

                await crawl(page, jenjang, [nama_prov], page.url, 1, hasil)

                await goto_aman(page, url_awal)
                file_out = simpan_provinsi_incremental(nama_prov, hasil)
                tandai_selesai(key)
                print(f"=== {key} SELESAI, {len(hasil)} baris -> {file_out} ===")

        print(f"\nSemua target udah diproses.")
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())