# SIMPADU RT — Koneksi Google Apps Script (GAS)

Aplikasi `index.html` sekarang bisa **tersinkron ke server** lewat Google Apps Script + Google Sheets. Semua warga yang membuka aplikasi akan melihat **data yang sama** (warga, kas, pengumuman, pengaduan, dll), bukan data lokal per perangkat.

## Cara Setup (sekali saja)

1. **Buat project GAS**
   - Buka https://script.google.com → klik **New project**
   - Hapus isi `Code.gs` bawaan, tempel seluruh isi file `Code.gs` dari folder ini

2. **(Opsional) Kunci akses**
   - Di `Code.gs`, isi `TOKEN` dengan kode rahasia, mis. `"rt14-modinan-2026"`
   - Jika diisi, semua request dari aplikasi harus menyertakan `?token=...`

3. **Deploy sebagai Web App**
   - Klik **Deploy → New deployment → Web app**
   - `Execute as` : **Me**
   - `Who has access` : **Anyone** ← wajib, supaya semua warga bisa akses tanpa login Google
   - Klik **Deploy**, izinkan akses jika diminta

4. **Salin URL Web App**
   - Hasil deploy berupa URL, mis. `https://script.google.com/macros/s/XXXX/exec`
   - Buka `index.html`, cari baris:
     ```js
     const GAS_URL="";
     ```
   - Tempel URL di antara tanda kutip:
     ```js
     const GAS_URL="https://script.google.com/macros/s/XXXX/exec";
     ```

5. **Uji**
   - Buka aplikasi → di pojok kanan atas muncul chip status sinkronisasi (☁️)
   - Buka **Pengaturan → Data & Cadangan** → tombol **🔄 Sinkronkan Sekarang**
   - Buka aplikasi dari HP/komputer lain → data yang sama muncul

## Cara Kerja

| Arah | Kapan | Mekanisme |
|---|---|---|
| Kirim ke server | Setiap data berubah (simpan/edit/hapus) | `save()` → debounce 1,5 detik → `POST action=save` |
| Muat dari server | Saat aplikasi dibuka | `GET action=get` → data server menimpa data lokal |
| Manual | Tombol di Pengaturan / chip status | `gas-sync-now`, `gas-pull`, `gas-push` |

- Data disimpan di **Google Sheets** (sheet `SIMPADU_RT`, kolom `key | value | updated`). Satu baris = satu modul (warga, transaksi, pengumuman, dll).
- `localStorage` tetap dipakai sebagai **cache offline** — aplikasi tetap jalan walau tanpa internet, lalu tersinkron saat online.
- Konflik edit bersamaan: **yang terakhir menulis menang** (cukup untuk skala RT).

## Catatan Penting

- **Jangan** ubah nama sheet `SIMPADU_RT` atau kolomnya — backend bergantung padanya.
- Setiap **perubahan di `Code.gs`** harus di-deploy ulang (Deploy → Manage deployments → Edit → New version).
- URL Web App lama tetap valid; buat versi baru hanya saat kode berubah.
- Data sensitif (data pribadi warga) tersimpan di Google Sheets milik Anda — atur izin berbagi sheet dengan bijak.

## Deployment Web (GitHub + Netlify)

Aplikasi di-hosting di **Netlify** dan terhubung ke **GitHub** untuk auto-deploy.

| Item | Nilai |
|---|---|
| Repo GitHub | `https://github.com/Antoksuryanto/rt14-modinan` |
| URL Warga | `https://rt14modinan.netlify.app` |
| URL Backend GAS | `https://script.google.com/macros/s/AKfycbxXCs1OnTC_dIvJQwxX19_iU0BMgG9PZ_SmzRtXUOD_FYIaCMB9WJ-dxF9WCcba7WkaZg/exec` |

### Alur Sinkronisasi

1. **Admin ubah data** → `save()` → debounce 1,5 detik → `POST action=save` → Google Sheets
2. **Warga buka halaman** → `GET action=get` → data server dimuat
3. **Admin ubah kode** → `git push` → Netlify auto-deploy (~30 detik)

### Cara Update Kode

```bash
git add -A
git commit -m "deskripsi perubahan"
git push
```

Netlify otomatis build & deploy versi baru. Tidak perlu deploy manual.