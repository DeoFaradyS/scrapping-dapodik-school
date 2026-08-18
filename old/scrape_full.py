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
    python scrape_full.py --provinsi "Surabaya" --kabkota "Surabaya"
    python scrape_full.py --jenjang "SD,SMP"            # cuma jenjang tertentu
    python scrape_full.py --provinsi "Gorontalo" --jenjang "SD"
    python scrape_full.py                              # semua provinsi & jenjang

OUTPUT:
    {Provinsi}_{KabKota}_{Jenjang}.xlsx   -> 1 file per kombo provinsi+kab/kota+jenjang
    Tiap run scrape fresh sesuai filter, gak ada resume/skip antar-run.
"""

import asyncio
import re
import sys
import pandas as pd
from playwright.async_api import async_playwright

BASE = "https://dapo.kemendikdasmen.go.id"
ROW_SELECTOR = "tr.cursor-pointer"
LEVEL_NAMES = ["Kab/Kota", "Kecamatan", "Sekolah"]  # depth 1, 2, 3
JENJANG_LIST = ["SD", "SMP", "SMA", "SMK"]  # fix, semua jenjang selalu diambil
RETRY_MAKS = 4
FIELD_WAJIB = ["NPSN", "Nama Sekolah", "Status"]  # kosong -> PERLU_CEK_MANUAL

STATS = {"sekolah": 0, "gagal": 0, "cek_manual": 0}

DEBUG = "--debug" in sys.argv
TEST = "--test" in sys.argv


def ambil_arg_list(flag, default):
    if flag not in sys.argv:
        return default
    idx = sys.argv.index(flag)
    if idx + 1 >= len(sys.argv):
        return default
    return [s.strip() for s in sys.argv[idx + 1].split(",")]


def normalize(nama):
    """lowercase + trim + rapetin spasi + strip prefix Kab./Kota/Prov. biar 'Surabaya' match 'Kota Surabaya'."""
    n = re.sub(r"\s+", " ", nama.strip().lower())
    n = re.sub(r"^(kab\.?|kabupaten|kota|prov\.?|provinsi)\s+", "", n)
    return n.strip()


PROVINSI_FILTER = ambil_arg_list("--provinsi", None)
if PROVINSI_FILTER:
    PROVINSI_FILTER = [normalize(s) for s in PROVINSI_FILTER]

KABKOTA_FILTER = ambil_arg_list("--kabkota", None)
if KABKOTA_FILTER:
    KABKOTA_FILTER = [normalize(s) for s in KABKOTA_FILTER]

JENJANG_FILTER = ambil_arg_list("--jenjang", JENJANG_LIST)
JENJANG_FILTER = [s.strip().upper() for s in JENJANG_FILTER]

STATUS_BLACKLIST = {
    "sangat baik", "baik", "cukup", "sedang", "kurang", "sangat kurang",
    "swasta", "negeri", "sudah", "belum", "ya", "tidak", "aktif", "nonaktif",
    "sudah sinkron", "belum sinkron", "sinkron", "belum sync", "sudah sync",
}


def start_url(jenjang):
    return f"{BASE}/progres?jenjang={jenjang}&status_sekolah=Swasta"


# ---------- penyimpanan ----------

def simpan_hasil(jenjang, nama_prov, hasil):
    """Pecah hasil satu provinsi jadi 1 file per Kab/Kota: '{Prov}_{KabKota}_{Jenjang}.xlsx'."""
    if not hasil:
        return []
    df = pd.DataFrame(hasil)
    files = []
    for kabkota, grup in df.groupby("Kab/Kota", dropna=False):
        file_out = bersihkan_nama_file(f"{nama_prov}_{kabkota}_{jenjang}") + ".xlsx"
        grup.to_excel(file_out, index=False)
        files.append(file_out)
    return files


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


def cari_index_by_nama(rows, nama, fallback_label):
    """Cari baris yg nama-nya cocok. None kalo gak ketemu/ambigu (biar caller fallback ke index)."""
    cocok = [i for i, (_, cols) in enumerate(rows)
             if nama_dari_cols(cols, fallback_label) == nama]
    return cocok[0] if len(cocok) == 1 else None


async def klik_baris_aman(page, index, nama=None, fallback_label="row", tries=RETRY_MAKS, timeout=45000):
    """Klik baris. Percobaan pertama pake index asli. Percobaan berikutnya (abis reload
    parent) cari ulang baris via cocok nama persis dulu, kalo ambigu/gak ketemu baru fallback index."""
    for percobaan in range(tries):
        try:
            rows = await get_rows_text(page)
            target = index
            if percobaan > 0 and nama:
                found = cari_index_by_nama(rows, nama, fallback_label)
                target = found if found is not None else index
            if target >= len(rows):
                return False
            row_el, _ = rows[target]
            async with page.expect_navigation(wait_until="networkidle", timeout=timeout):
                await row_el.click()
            await page.wait_for_timeout(2000)
            return True
        except Exception:
            await page.wait_for_timeout(2000)
    print(f"        [!] klik baris {index} ({nama}) gagal total")
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
        fallback_label = f"{label.lower()}_{i+1}"
        nama = bersihkan_nama_file(nama_dari_cols(cols, fallback_label))

        if depth == 1 and KABKOTA_FILTER and not any(f in normalize(nama) for f in KABKOTA_FILTER):
            continue

        path_baru = path + [nama]

        ok = await klik_baris_aman(page, i, nama=nama_dari_cols(cols, fallback_label), fallback_label=fallback_label)
        if not ok:
            hasil.append(baris_gagal(jenjang, path_baru, f"PERLU_CEK_MANUAL: gagal buka {label}"))
            await goto_aman(page, parent_url)
            continue
        url_sekarang = page.url

        if depth == 3:  # baru klik sekolah -> ini halaman detail
            await catat_sekolah(page, jenjang, path_baru, hasil)
        else:
            await crawl(page, jenjang, path_baru, url_sekarang, depth + 1, hasil)

        await goto_aman(page, parent_url)


def baris_gagal(jenjang, path, pesan):
    STATS["gagal"] += 1
    STATS["cek_manual"] += 1
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

    STATS["sekolah"] += 1
    row = {
        "Jenjang": jenjang,
        "Provinsi": path[0],
        "Kab/Kota": path[1],
        "Kecamatan": path[2],
        "Nama Sekolah": nama_sek,
        "Status": "sukses",
        **field_data,
    }
    kosong = [f for f in FIELD_WAJIB if not str(row.get(f, "")).strip()]
    if kosong:
        row["Status"] = f"PERLU_CEK_MANUAL: kosong ({', '.join(kosong)})"
        STATS["cek_manual"] += 1

    print(f"        + {nama_sek}  [{STATS['sekolah']} sekolah total | cek manual {STATS['cek_manual']}]")
    hasil.append(row)


# ---------- main ----------

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=not DEBUG)
        page = await browser.new_page()

        for ji, jenjang in enumerate(JENJANG_FILTER):
            url_awal = start_url(jenjang)
            await goto_aman(page, url_awal)

            prov_rows = await get_rows_text(page)
            print(f"\n########## Jenjang {jenjang} ({ji+1}/{len(JENJANG_FILTER)}): {len(prov_rows)} provinsi ##########")
            jumlah_prov = 1 if TEST else len(prov_rows)

            for pi in range(jumlah_prov):
                prov_rows = await get_rows_text(page)
                if pi >= len(prov_rows):
                    continue
                _, cols = prov_rows[pi]
                fallback_label = f"provinsi_{pi+1}"
                nama_prov = bersihkan_nama_file(nama_dari_cols(cols, fallback_label))

                if PROVINSI_FILTER and not any(f in normalize(nama_prov) for f in PROVINSI_FILTER):
                    continue

                print(f"\n=== {jenjang} - Provinsi {pi+1}/{len(prov_rows)}: {nama_prov} "
                      f"| {STATS['sekolah']} sekolah terkumpul ===")
                hasil = []

                ok = await klik_baris_aman(page, pi, nama=nama_dari_cols(cols, fallback_label), fallback_label=fallback_label)
                if not ok:
                    hasil.append(baris_gagal(jenjang, [nama_prov], "PERLU_CEK_MANUAL: gagal buka provinsi"))
                    simpan_hasil(jenjang, nama_prov, hasil)
                    await goto_aman(page, url_awal)
                    continue

                await crawl(page, jenjang, [nama_prov], page.url, 1, hasil)

                await goto_aman(page, url_awal)
                files_out = simpan_hasil(jenjang, nama_prov, hasil)
                print(f"=== {jenjang}|{nama_prov} SELESAI, {len(hasil)} baris -> {len(files_out)} file ===")

        print(f"\nSemua target udah diproses.")
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())