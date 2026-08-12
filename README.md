# Scraper Dapodik

Ambil data sekolah swasta (SD/SMP/SMA/SMK) dari [dapo.kemendikdasmen.go.id](https://dapo.kemendikdasmen.go.id/progres), jenjang Provinsi → Kab/Kota → Kecamatan → Sekolah → Detail. Output 1 file Excel per provinsi.

## Instalasi

```bash
pip install playwright openpyxl pandas
playwright install chromium
```

## Command

| Command | Fungsi |
|---|---|
| `python scrape_full.py` | Semua provinsi & jenjang (SD/SMP/SMA/SMK), auto-skip yang udah selesai |
| `python scrape_full.py --provinsi "Gorontalo"` | Cuma 1 provinsi |
| `python scrape_full.py --provinsi "Gorontalo,Bali"` | Beberapa provinsi |
| `python scrape_full.py --jenjang "SD,SMP"` | Cuma jenjang tertentu |
| `python scrape_full.py --provinsi "Gorontalo" --jenjang "SD"` | Kombinasi provinsi + jenjang |
| `python scrape_full.py --ulang --provinsi "Gorontalo"` | Paksa scrape ulang provinsi yang udah "selesai" |
| `python scrape_full.py --ulang` | Hapus semua progress, mulai dari nol |
| `python scrape_full.py --test` | Test cepet: 1 provinsi, 1 kab, 1 kec doang |
| `python scrape_full.py --debug` | Browser kelihatan (buat cek jalannya) |

Flag bisa digabung, misal:
```bash
python scrape_full.py --test --debug --provinsi "Gorontalo"
```

## Cara Kerja

1. Jalan provinsi per provinsi, jenjang per jenjang. Tiap level (kab/kec/sekolah) ada retry 4x kalo gagal buka — kalo tetep gagal, dicatet sebagai baris "GAGAL buka ..." di excel (bukan skip diem-diem).
2. Progress ditampilin real-time pake counter jalan: `[187 sekolah total | gagal 2]`. Gak ada persen (total sekolah nasional gak diketahui di depan), tapi angka ini nunjukin script masih hidup & terus maju.

## Output

| File | Isi |
|---|---|
| `{NamaProvinsi}.xlsx` | 1 file per provinsi, gabungan semua jenjang. Kolom: Jenjang, Provinsi, Kab/Kota, Kecamatan, Nama Sekolah, NPSN, Status, Kepala Sekolah, Akreditasi, Peserta Didik, PTK, Rombel, Alamat, dll |
| `progress_selesai.txt` | Daftar `Jenjang\|Provinsi` yang udah kelar — run berikutnya auto-skip yang udah selesai |

## Data yang Diambil

Identitas sekolah, NPSN, alamat, kepala sekolah, akreditasi, jumlah Peserta Didik/PTK/Rombel. **Tidak** termasuk data fasilitas (toilet, dll).

## Kalo Macet / Error

- Script udah retry otomatis 4x tiap level, jadi error koneksi sesaat gak bikin berhenti total.
- Kalo mau lanjutin abis macet: run command yang sama lagi — provinsi yang udah selesai otomatis di-skip (baca `progress_selesai.txt`).
- Kalo ada baris "GAGAL buka ..." di hasil excel: run ulang provinsi itu pake `--ulang --provinsi "NamaProvinsi"`.