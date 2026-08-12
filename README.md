# Scraper Dapodik

Ambil data sekolah swasta (SD/SMP/SMA/SMK) dari [dapo.kemendikdasmen.go.id](https://dapo.kemendikdasmen.go.id/progres), jenjang Provinsi → Kab/Kota → Kecamatan → Sekolah → Detail. Output 1 file Excel per kombo Provinsi + Kab/Kota + Jenjang.

## Instalasi

```bash
pip install playwright openpyxl pandas
playwright install chromium
```

## Command

| Command                                                            | Fungsi                                     |
| ------------------------------------------------------------------ | ------------------------------------------ |
| `python scrape_full.py`                                            | Semua provinsi & jenjang (SD/SMP/SMA/SMK)  |
| `python scrape_full.py --provinsi "Gorontalo"`                     | Cuma 1 provinsi                            |
| `python scrape_full.py --provinsi "Gorontalo,Bali"`                | Beberapa provinsi                          |
| `python scrape_full.py --provinsi "Surabaya" --kabkota "Surabaya"` | 1 provinsi + 1 kab/kota                    |
| `python scrape_full.py --jenjang "SD,SMP"`                         | Cuma jenjang tertentu                      |
| `python scrape_full.py --provinsi "Gorontalo" --jenjang "SD"`      | Kombinasi provinsi + jenjang               |
| `python scrape_full.py --test`                                     | Test cepet: 1 provinsi, 1 kab, 1 kec doang |
| `python scrape_full.py --debug`                                    | Browser kelihatan (buat cek jalannya)      |

Flag bisa digabung, misal:

```bash
python scrape_full.py --test --debug --provinsi "Gorontalo"
```

Filter provinsi/kab-kota gak case-sensitive, gak butuh spasi persis, dan otomatis nerima ketikan tanpa prefix ("Surabaya" tetep ketemu "Kota Surabaya", "Bandung" ketemu "Kab. Bandung"). Nama file & isi Excel tetep pake nama ASLI hasil scrape, bukan ketikan filter.

**Tiap run scrape fresh** sesuai filter — gak nyimpen state antar-run, gak ada resume/skip otomatis.

## Cara Kerja

1. Jalan provinsi per provinsi, jenjang per jenjang, difilter dulu di level Kab/Kota kalo `--kabkota` diisi (jadi kota lain gak ikut nyasar).
2. Tiap level (kab/kec/sekolah) retry maks 4x kalo gagal buka. Percobaan pertama pake index baris asli; percobaan berikutnya cari ulang baris pake cocok nama persis dulu (biar gak salah klik gara-gara tabel geser abis reload), baru fallback ke index asli kalo nama ambigu/gak ketemu.
3. Gagal abis retry maks, atau field wajib (NPSN / Nama Sekolah / Status) kosong abis diambil → dicatet dengan Status `PERLU_CEK_MANUAL: <alasan>`, bukan diskip diem-diem. Field lain (Akreditasi, Kepala Sekolah, Alamat, PTK, Rombel, Peserta Didik) boleh kosong, wajar, gak diflag.
4. Progress ditampilin real-time: `[187 sekolah total | cek manual 2]`.

## Output

| File                                  | Isi                                                                                                                                                                                        |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `{Provinsi}_{KabKota}_{Jenjang}.xlsx` | 1 file per kombo provinsi+kab/kota+jenjang. Kolom: Jenjang, Provinsi, Kab/Kota, Kecamatan, Nama Sekolah, NPSN, Status, Kepala Sekolah, Akreditasi, Peserta Didik, PTK, Rombel, Alamat, dll |

## Data yang Diambil

Identitas sekolah, NPSN, alamat, kepala sekolah, akreditasi, jumlah Peserta Didik/PTK/Rombel. **Tidak** termasuk data fasilitas (toilet, dll).

## Kalo Ada Baris `PERLU_CEK_MANUAL`

Cek kolom Status di Excel — isinya alasan (gagal buka halaman, atau field wajib kosong). Run ulang provinsi/kab-kota itu pake `--provinsi "..." --kabkota "..."` buat coba lagi.

## GitHub Actions (`scrape.yml`)

Trigger manual dari tab Actions, jalan per jenjang paralel (matrix SD/SMP/SMA/SMK) khusus Jawa Timur. Ada input opsional `kabkota` buat filter kab/kota tertentu; kosongin buat semua kab/kota di Jawa Timur.

## Changelog

- **v2**: Hapus progress tracking (`progress_selesai.txt`, `--ulang`) — tiap run scrape fresh. Tambah filter `--kabkota` (normalize match, bukan cuma provinsi). Output dipecah per kombo Provinsi+KabKota+Jenjang (sebelumnya 1 file gabung per provinsi). Retry level pake cocok nama dulu baru fallback index. Baris gagal/field kosong ditandai `PERLU_CEK_MANUAL` (sebelumnya `GAGAL ...` / status "sukses" tanpa validasi).
- **v1**: Versi awal — scrape provinsi/kab/kota/kecamatan/sekolah, 1 file per provinsi gabung semua jenjang, progress resume via `progress_selesai.txt`.
