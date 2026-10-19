# Alethra Office (Android)

Aplikasi **Office Web** (Word, Excel, PowerPoint, PDF) yang dibungkus menjadi **APK Android** dengan
[Capacitor](https://capacitorjs.com). Seluruh kode dan library dimasukkan ke dalam APK, jadi aplikasi
**berjalan 100% offline** (tanpa CDN, tanpa server).

## Cara mendapatkan APK (tanpa Android Studio)

1. Buat repo baru di GitHub, lalu push seluruh isi folder ini ke branch `main`.
2. Buka tab **Actions** -> workflow **Build Android APK** (jalan otomatis tiap push; atau klik *Run workflow*).
3. Setelah selesai (± 5-8 menit), buka tab **Releases** dan unduh `Alethra-Office.apk` (atau dari **Artifacts** pada run tersebut).
4. Kirim APK ke HP, lalu pasang (izinkan "Pasang dari sumber tidak dikenal").

APK juga otomatis dipublikasikan di tab **Releases** pada setiap push.

## Struktur repo

| Path | Isi |
|---|---|
| `www/` | Aplikasi web: `index.html`, `shell.js/css` (Beranda & daftar berkas), editor `word/ excel/ ppt/ pdf/` |
| `www/lib/` | Semua library pihak ketiga (lokal, untuk offline) |
| `www/native-bridge.js` | Jembatan WebView: simpan berkas, cetak, akses plugin native |
| `android/` | Proyek Android (Capacitor) + plugin native `FilesPlugin.java` (daftar/kelola berkas) dan `FileSaverPlugin.java` (simpan/cetak) |
| `scripts/vendor-libs.js` | Menyalin ulang library dari `node_modules` ke `www/lib` |
| `.github/workflows/build-apk.yml` | Workflow build APK |

## Tampilan & fitur

- **Beranda**: tab *Semua dokumen*, *Terbaru*, *Favorit*, dan filter *Lokasi*; navigasi bawah Beranda / Word / Excel / PPoint / PDF.
- **Tab Word / Excel / PPoint / PDF**: header kiri berisi tombol **Baru** (dokumen kosong) dan **Buka** (pemilih berkas),
  menggantikan logo; di bawahnya daftar berkas jenis itu.
- **Menu titik tiga** pada tiap berkas: Buka dengan..., Ubah Nama, Favorit, Buat jalan pintas, Bagikan,
  Bagikan melalui Bluetooth, Simpan ke Google Drive, Hapus, Detail.
- Ikon cari, urutkan (nama / tanggal / ukuran), dan info aplikasi di pojok kanan header.

## Izin akses berkas

Agar daftar dokumen di HP tampil, aplikasi meminta **"Akses semua berkas"** (Android 11+) saat pertama dibuka:
tekan *Izinkan akses*, aktifkan toggle, lalu kembali ke aplikasi. (Android 10 ke bawah: izin penyimpanan biasa.)
Aplikasi memindai `.docx`, `.xlsx`, `.pptx`, `.pdf` di penyimpanan internal (kecuali `Android/` dan folder tersembunyi).
Izin ini cocok untuk sideload; Google Play membatasi izin tersebut.

## Perilaku khusus di Android

- **Simpan**: berkas yang dibuka dari daftar **ditimpa di lokasi aslinya** (tidak lagi muncul `(1)`, `(2)`);
  dokumen baru disimpan ke folder **Download** dan simpan berikutnya menimpa berkas yang sama.
  Tanpa izin akses berkas, penyimpanan memakai folder Download biasa.
- **Buka dari aplikasi lain**: file manager / WhatsApp / dsb. bisa membuka `.docx/.xlsx/.pptx/.pdf` langsung di aplikasi ini.
- **Cetak** membuka dialog cetak Android (pilih *Simpan sebagai PDF* untuk membuat PDF).
- Tombol Back: dari tab jenis kembali ke Beranda; di dalam editor gunakan menu Berkas > Kembali ke menu utama.
- Fitur yang memang butuh internet (gambar dari URL, video YouTube, *Dari Web* di Excel, pencarian Wikipedia/terjemahan
  di PowerPoint) hanya jalan saat HP online; sisanya offline penuh.

## Update APK tanpa uninstall

Key penandatangan tetap: file `android/app/office.jks` ikut di repo (password ada di `android/app/build.gradle`),
dan `versionCode` otomatis naik mengikuti nomor run GitHub. Jadi setiap APK baru dari repo ini
**bisa langsung dipasang menimpa versi lama** tanpa uninstall dan tanpa kehilangan data.

- Setiap push ke `main` membuat **GitHub Release** baru (`v1.<nomor run>`) berisi `Alethra-Office.apk`.
- Jangan hapus/ganti `office.jks`; kalau key berubah, update tidak bisa lagi menimpa dan harus uninstall.
- Karena keystore ada di repo, jaga agar repo tetap **private** (cukup untuk dipakai sendiri/sideload).
  Untuk Google Play, gunakan keystore rahasia lewat env `KEYSTORE_FILE`, `KEYSTORE_PASSWORD`, `KEY_ALIAS`, `KEY_PASSWORD`.

## Build lokal (opsional)

Butuh Node 22+, JDK 21, dan Android SDK.

```bash
npm ci
npx cap sync android
cd android && ./gradlew assembleDebug
# hasil: android/app/build/outputs/apk/debug/app-debug.apk
```

Ubah `appId` / nama aplikasi di `capacitor.config.json` dan `android/app/build.gradle` bila perlu.

## Menu "Format" kontekstual di Excel

Seperti *Format Tabel/Gambar/Garis* di Word, Excel kini punya tab oranye yang muncul otomatis saat objek di lembar dipilih
(dan kembali ke tab sebelumnya saat objek tidak lagi dipilih):

| Objek terpilih | Tab | Fitur utama |
|---|---|---|
| Gambar | FORMAT GAMBAR | kecerahan, kontras, efek warna, gaya cepat, garis tepi, sudut membulat, bayangan, ukuran (cm), kunci rasio, skala %, ukuran asli, atur ulang |
| Bentuk / Kotak Teks | FORMAT BENTUK | ubah jenis bentuk, gaya cepat, isi, garis tepi (warna/tebal/jenis), opasitas, bayangan, font teks, rata teks |
| Bagan | FORMAT BAGAN | pilih data, tukar baris/kolom, jenis, skema warna, judul, judul sumbu, legenda, label data, garis kisi, latar & tepi |
| Sel ber-sparkline | FORMAT SPARKLINE | jenis, titik tertinggi/terendah/negatif, penanda, sumbu, warna, tebal, edit data |

Semua tab objek juga punya: ukuran & rotasi, **Pas ke Sel**, **Pindah ke Sel…**, susunan (majukan/mundurkan/ratakan/putar),
Panel Seleksi, nama objek, teks alternatif, gandakan, dan hapus. Perubahan berlaku ke semua objek yang dipilih (Shift+klik) dan bisa di-*Undo*.

Catatan: gambar dengan efek/bingkai, bagan, bentuk, kotak teks, dan sparkline hanya tersimpan penuh di sesi aplikasi; saat disimpan ke `.xlsx`
hanya gambar (tanpa efek/bingkai) yang ikut tersimpan.

## Menu "Format" kontekstual di PowerPoint

Sama seperti Word dan Excel, PowerPoint kini punya tab oranye yang muncul otomatis saat objek di slide dipilih
(kembali ke tab sebelumnya saat pilihan dilepas). Fiturnya menyesuaikan jenis objek:

| Objek terpilih | Tab | Fitur utama |
|---|---|---|
| Gambar | FORMAT GAMBAR | kecerahan, kontras, saturasi, efek warna, gaya cepat, garis tepi, sudut membulat/oval, bayangan/cahaya/pantulan, pangkas (4 sisi), opasitas, ganti gambar, ukuran asli, skala %, atur ulang |
| Video / Audio | FORMAT MEDIA | mulai otomatis, ulangi, senyap, tanpa kontrol, sembunyikan ikon audio saat tayang, gaya (tepi, sudut, efek, opasitas), rasio 16:9 / 4:3 / 1:1 |
| Bentuk / Garis / Kotak Teks | FORMAT BENTUK | ubah jenis bentuk, gaya cepat, isian, garis tepi (warna/tebal/jenis), opasitas isian, efek, font, ukuran, tebal/miring/garis bawah, warna, rata teks, arah teks |
| Tabel | FORMAT TABEL | gaya tabel, baris judul/belang/kolom pertama, warna & tebal tepi, isian sel, rata sel (sel/baris/kolom/tabel), sisipkan/hapus baris & kolom |

Semua tab juga punya: tinggi, lebar, rotasi (cm/derajat), **kunci rasio**, **Pas ke Slide**, posisi horizontal/vertikal (cm), **Tengahkan**,
susun (depan/belakang/ratakan/putar/balik), **Panel Seleksi**, nama objek, teks alternatif, gandakan, dan hapus.
Perubahan berlaku ke semua objek yang dipilih dan bisa di-*Undo*. Untuk kotak teks/placeholder tab Format tersedia
tetapi tidak memindahkan tab Beranda saat Anda mengetik.

### Geser, ubah ukuran, putar (mouse **dan sentuhan HP**)

- Geser objek dengan menyeret; muncul **garis panduan** magenta saat sejajar dengan tengah/tepi slide atau objek lain (tahan Alt untuk menonaktifkan).
- Handle ubah ukuran bekerja benar walau objek sedang **diputar**; sudut gambar/video menjaga rasio (Shift untuk bentuk).
- **Handle putar** di atas objek (otomatis menempel ke kelipatan 15°).
- Di HP: seret objek/handle dengan jari, **cubit dua jari** untuk skala, **ketuk dua kali** bentuk untuk edit teks.

Catatan penyimpanan `.pptx`: efek gambar (kecerahan, kontras, efek warna, pangkas, sudut, garis tepi, balik) "dipanggang" ke dalam gambar;
opasitas gambar/isian tersimpan; opsi media dan bayangan/efek bentuk khusus tersimpan penuh hanya di sesi aplikasi.
