"""
Scraper Sekolah Swasta (SD/SMP/SMA/SMK) - data.kemendikdasmen.go.id

Kenapa pake Playwright padahal ambil data via API?
API-nya (api.data.belajar.id) diproteksi WAF yg cek TLS fingerprint koneksi
HTTPS - request polos dari Python (`requests`) ketolak 403 walau header disamain
persis kayak browser. Playwright buka Chromium ASLI, jadi TLS-nya genuine,
WAF lolos. Tapi kita GAK KLIK UI SAMA SEKALI - cuma manggil fetch() dari dalem
browser context (page.evaluate), jadi tetep ngebut kayak API murni.

API dipake:
  LIST   : https://api.data.belajar.id/data-portal-backend/v2/master-data/satuan-pendidikan/daftar-data-induk/{kodeWilayah}
  DETAIL : https://api.data.belajar.id/data-portal-backend/v1/master-data/satuan-pendidikan/details/{npsn}

CARA PAKAI:
    pip install playwright
    playwright install chromium
    python scrape_api.py

    Muncul menu pilih provinsi, ketik nomor (bisa banyak dipisah koma, atau 0 = semua).

OUTPUT:
    {NamaProvinsi}_Swasta.csv  -> 1 file per provinsi
    Kolom: NPSN | Nama Sekolah | Bentuk | Telpon | Email | Naungan | Kabupaten | Kecamatan | Alamat
    Ditulis streaming per baris, gak numpuk di memory.
"""

import asyncio
import csv
import json
import os
import sys
from playwright.async_api import async_playwright

HALAMAN_AWAL = ("https://data.kemendikdasmen.go.id/data-induk/satpen/daftar"
                 "?pembina=kemendikdasmen&jalur=formal&bentuk=sd,smp,sma,smk"
                 "&status=2&jenis=pendidikan-umum,pendidikan-kejuruan")

LIST_API = ("https://api.data.belajar.id/data-portal-backend/v2/master-data/satuan-pendidikan"
            "/daftar-data-induk/{kode}?pembina=kemendikdasmen&jalurPendidikan=formal"
            "&bentukPendidikan=sd,smp,sma,smk&statusSatuanPendidikan=2"
            "&jenisPendidikan=pendidikan-umum,pendidikan-kejuruan&limit={limit}&offset={offset}")

DETAIL_API = "https://api.data.belajar.id/data-portal-backend/v1/master-data/satuan-pendidikan/details/{npsn}"

RETRY_MAKS = 4
BATCH_DETAIL = 15  # berapa banyak detail sekolah ditarik bareng2 (concurrent fetch dalem browser)

DEBUG = "--debug" in sys.argv
TEST = "--test" in sys.argv  # cuma ambil 1 halaman list pertama per provinsi

STATS = {"sekolah": 0, "gagal_detail": 0}

# JS snippet: fetch() dari dalem browser context, return JSON atau {__error: status}
JS_FETCH = """
async (url) => {
    try {
        const res = await fetch(url, {
            headers: {
                'accept': '*/*',
                'x-bot': '0',
                'x-bot-type': 'user',
                'x-client-app': 'data-portal-fe',
                'x-request-id': crypto.randomUUID()
            }
        });
        if (!res.ok) return {__error: res.status};
        return await res.json();
    } catch (e) {
        return {__error: String(e)};
    }
}
"""


async def fetch_json(page, url, tries=RETRY_MAKS):
    for percobaan in range(tries):
        try:
            hasil = await page.evaluate(JS_FETCH, url)
            if isinstance(hasil, dict) and "__error" in hasil:
                if DEBUG:
                    print(f"        [!] fetch gagal ({hasil['__error']}) @ {url}")
                await page.wait_for_timeout(1500 * (percobaan + 1))
                continue
            return hasil
        except Exception as e:
            if DEBUG:
                print(f"        [!] evaluate error ({e}) @ {url}")
            await page.wait_for_timeout(1500 * (percobaan + 1))
    return None


def load_provinsi():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "provinsi.json")
    if not os.path.exists(path):
        print("[!] provinsi.json gak ketemu di folder yg sama. Taruh file itu di sini dulu.")
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return dict(sorted(data.items()))


def pilih_provinsi(daftar):
    nama_list = list(daftar.keys())
    print("\n=== Pilih Provinsi ===")
    for i, nama in enumerate(nama_list, 1):
        print(f"{i:2d}. {nama}")
    print(" 0. SEMUA provinsi")
    pilihan = input("\nMasukin nomor (pisah koma kalo lebih dari 1, misal 1,3,5): ").strip()

    if pilihan == "0":
        return daftar

    idx_list = [s.strip() for s in pilihan.split(",") if s.strip()]
    hasil = {}
    for idx in idx_list:
        if idx.isdigit() and 1 <= int(idx) <= len(nama_list):
            nama = nama_list[int(idx) - 1]
            hasil[nama] = daftar[nama]
        else:
            print(f"    [!] '{idx}' diabaikan, bukan nomor valid.")
    return hasil


def bersihkan_nama_file(nama):
    import re
    nama = re.sub(r'[\\/:*?"<>|\t\n\r]', "-", nama).strip()
    return nama[:150] if nama else "tanpa_nama"


async def deteksi_limit(page, kode):
    for coba in (100, 50, 20):
        url = LIST_API.format(kode=kode, limit=coba, offset=0)
        data = await fetch_json(page, url, tries=2)
        if data and data.get("meta", {}).get("limit") == coba:
            return coba
    return 20


async def ambil_daftar_sekolah(page, kode, limit):
    semua = []
    offset = 0
    total = None
    while True:
        url = LIST_API.format(kode=kode, limit=limit, offset=offset)
        data = await fetch_json(page, url)
        if not data:
            print(f"        [!] gagal ambil list @ offset {offset}, skip sisanya.")
            break
        rows = data.get("data", [])
        meta = data.get("meta", {})
        if total is None:
            total = meta.get("total", 0)
            print(f"        total sekolah: {total}")
        semua.extend(rows)
        offset += limit
        if TEST:
            break
        if offset >= total or not rows:
            break
    return semua


async def ambil_detail_batch(page, npsn_list):
    """Tarik detail beberapa NPSN sekaligus (concurrent) dalem 1 browser context."""
    tugas = [fetch_json(page, DETAIL_API.format(npsn=npsn)) for npsn in npsn_list]
    hasil_list = await asyncio.gather(*tugas)
    out = {}
    for npsn, data in zip(npsn_list, hasil_list):
        if not data:
            STATS["gagal_detail"] += 1
            out[npsn] = {"Telpon": "", "Email": "", "Naungan": "PERLU_CEK_MANUAL"}
            continue
        sp = data.get("satuanPendidikan", {})
        out[npsn] = {
            "Telpon": sp.get("nomorTelepon", "") or "",
            "Email": sp.get("email", "") or "",
            "Naungan": sp.get("namaYayasan", "") or "",
        }
    return out


async def scrape_provinsi(page, nama_prov, kode):
    print(f"\n=== {nama_prov} (kode {kode}) ===")
    limit = await deteksi_limit(page, kode)
    print(f"    pake limit={limit} per request")

    rows = await ambil_daftar_sekolah(page, kode, limit)
    print(f"    {len(rows)} sekolah ketemu, mulai ambil detail (telpon/email/naungan)...")

    file_out = bersihkan_nama_file(f"{nama_prov}_Swasta") + ".csv"
    kolom = ["NPSN", "Nama Sekolah", "Bentuk", "Telpon", "Email", "Naungan", "Kabupaten", "Kecamatan", "Alamat"]

    with open(file_out, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=kolom)
        writer.writeheader()

        for i in range(0, len(rows), BATCH_DETAIL):
            batch = rows[i:i + BATCH_DETAIL]
            npsn_list = [r["npsn"] for r in batch]
            detail_map = await ambil_detail_batch(page, npsn_list)

            for r in batch:
                detail = detail_map.get(r["npsn"], {"Telpon": "", "Email": "", "Naungan": "PERLU_CEK_MANUAL"})
                row = {
                    "NPSN": r.get("npsn", ""),
                    "Nama Sekolah": r.get("nama", ""),
                    "Bentuk": r.get("bentukPendidikan", ""),
                    "Telpon": detail["Telpon"],
                    "Email": detail["Email"],
                    "Naungan": detail["Naungan"],
                    "Kabupaten": r.get("namaKabupaten", ""),
                    "Kecamatan": r.get("namaKecamatan", ""),
                    "Alamat": r.get("alamatJalan", ""),
                }
                writer.writerow(row)
            f.flush()
            STATS["sekolah"] += len(batch)
            print(f"        {min(i + BATCH_DETAIL, len(rows))}/{len(rows)} sekolah tercatat...")

    print(f"=== {nama_prov} SELESAI -> {file_out} ({len(rows)} baris) ===")


async def main():
    daftar = load_provinsi()
    daftar = {k: v for k, v in daftar.items() if v}
    terpilih = pilih_provinsi(daftar)

    if not terpilih:
        print("Gak ada provinsi valid dipilih, keluar.")
        return

    print(f"\nMulai scrape {len(terpilih)} provinsi: {', '.join(terpilih.keys())}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=not DEBUG)
        page = await browser.new_page()
        # buka halaman aslinya dulu, biar origin/cookie/context kebentuk bener
        await page.goto(HALAMAN_AWAL, wait_until="networkidle", timeout=90000)
        await page.wait_for_timeout(1000)

        for nama_prov, kode in terpilih.items():
            await scrape_provinsi(page, nama_prov, kode)

        await browser.close()

    print(f"\nSemua selesai. Total {STATS['sekolah']} sekolah, {STATS['gagal_detail']} gagal ambil detail (cek kolom Naungan='PERLU_CEK_MANUAL').")


if __name__ == "__main__":
    asyncio.run(main())