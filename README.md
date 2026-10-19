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
