# Scraper Data Induk Satuan Pendidikan (API-based)

Ambil data sekolah swasta (SD/SMP/SMA/SMK) dari [data.kemendikdasmen.go.id](https://data.kemendikdasmen.go.id/data-induk/satpen/daftar), langsung dari API resminya (`api.data.belajar.id`) — bukan klik-klik UI. Output 1 file CSV per provinsi.

> **Ganti dari versi lama** (`scrape_full.py`, drill Provinsi→Kab/Kota→Kecamatan→Sekolah di dapo.kemendikdasmen.go.id). Situs sumber udah pindah/redesign, struktur baru gak butuh drill-down — data ditarik langsung per provinsi dari API.

## Instalasi

```bash
pip install playwright
playwright install chromium
```

Playwright tetep dipake (bukan buat klik UI, tapi buat numpang TLS/koneksi browser asli — API-nya diproteksi WAF yang nolak request polos dari Python `requests`, harus lewat browser beneran).

## File yang Dibutuhin

| File            | Fungsi                                                             |
| --------------- | ------------------------------------------------------------------ |
| `scrape_api.py` | Script utama                                                       |
| `provinsi.json` | Daftar nama+kode wilayah 39 provinsi, dipake buat isi menu pilihan |

Taruh 2 file itu di folder yang sama.

## Cara Pakai

```bash
python scrape_api.py
```

Muncul menu pilih provinsi, tinggal ketik nomor:

```
=== Pilih Provinsi ===
 1. Aceh
 2. Bali
 3. Banten
 ...
39. Sumatera Utara
 0. SEMUA provinsi

Masukin nomor (pisah koma kalo lebih dari 1, misal 1,3,5):
```

- `1` → cuma Aceh
- `1,5,10` → beberapa provinsi
- `0` → semua Indonesia (59rb+ sekolah, lama)

### Flag

| Flag      | Fungsi                                                                       |
| --------- | ---------------------------------------------------------------------------- |
| `--test`  | Cuma ambil 1 halaman list pertama per provinsi (±20 sekolah), buat cek cepet |
| `--debug` | Browser Chromium kelihatan + print detail tiap request yang gagal            |

Bisa digabung: `python scrape_api.py --test --debug`

## Cara Kerja

1. Buka `data.kemendikdasmen.go.id` sekali (biar context browser kebentuk — origin/cookie beres).
2. Per provinsi terpilih, panggil **List API** (`daftar-data-induk/{kodeWilayah}`) loop tiap halaman → dapet NPSN, Nama, Bentuk, Kabupaten, Kecamatan, Alamat semua sekolah.
3. Per sekolah, panggil **Detail API** (`details/{npsn}`) buat ambil Telpon, Email, Naungan — ditarik **15 sekaligus** (concurrent) per batch biar ngebut, bukan 1-1.
4. Gagal ambil detail (abis retry 4x) → tetep dicatet, kolom Naungan diisi `PERLU_CEK_MANUAL`, gak diskip diem-diem.
5. Ditulis ke CSV **streaming per batch** (flush langsung ke disk), gak numpuk di memory.

Semua request pake header niru browser asli (`Origin`, `x-client-app`, dll) + `x-request-id` random tiap call, biar gak keblokir anti-bot.

## Output

| File                    | Isi                                                                                                          |
| ----------------------- | ------------------------------------------------------------------------------------------------------------ |
| `{Provinsi}_Swasta.csv` | 1 file per provinsi. Kolom: NPSN, Nama Sekolah, Bentuk, Telpon, Email, Naungan, Kabupaten, Kecamatan, Alamat |

Semua jenjang (SD/SMP/SMA/SMK) kegabung dalem 1 file per provinsi, dibedain lewat kolom **Bentuk**.

## Kalo Ada Baris `PERLU_CEK_MANUAL` (di kolom Naungan)

Brarti Detail API gagal ditarik abis retry 4x buat NPSN itu (biasa network/rate-limit sesaat). NPSN & data dasar (Nama/Kabupaten/dll) tetep lengkap, cuma Telpon/Email/Naungan yang kosong. Run ulang provinsi yang sama kalo mau coba lagi — atau cek manual NPSN itu ke situsnya langsung.

## Tuning Kecepatan

`BATCH_DETAIL = 15` di awal script — jumlah request detail yang ditembak bareng. Kalo mulai banyak `[!] fetch gagal` pas run `--debug`, turunin ke `10`. Kalo lancar dan mau lebih ngebut, bisa naikin ke `25-30`.

## Ambil Ulang `provinsi.json`

Kalo situsnya ganti kode wilayah suatu saat, `provinsi.json` bisa dibikin ulang manual dari API publik:

```
https://data.kemendikdasmen.go.id/data-induk/satpen/jumlah?...
```

Buka Network tab (F12) pas load halaman itu, cari request yang responnya list `{kodeWilayah, namaWilayah, ...}` per provinsi, copy `data[]`-nya.

## Changelog

- **v3 (API-based)**: Situs sumber ganti total ke `data.kemendikdasmen.go.id` (drill-down lama gak berlaku lagi). Scraper dirombak pake API langsung (`api.data.belajar.id`) — Playwright cuma dipake numpang TLS/browser context, gak klik UI. Output disederhanain jadi CSV per provinsi (bukan per kombo Provinsi+KabKota+Jenjang), kolom dikurangin sesuai kebutuhan (NPSN, Nama, Bentuk, Telpon, Email, Naungan, Kabupaten, Kecamatan, Alamat — Akreditasi & Rekapitulasi/Peserta-Didik-PTK dihapus).
- **v2**: Hapus progress tracking (`progress_selesai.txt`, `--ulang`) — tiap run scrape fresh. Tambah filter `--kabkota`. Output dipecah per kombo Provinsi+KabKota+Jenjang. Retry level pake cocok nama dulu baru fallback index. Baris gagal/field kosong ditandai `PERLU_CEK_MANUAL`.
- **v1**: Versi awal — scrape provinsi/kab/kota/kecamatan/sekolah drill-down manual di dapo.kemendikdasmen.go.id, progress resume via `progress_selesai.txt`.
