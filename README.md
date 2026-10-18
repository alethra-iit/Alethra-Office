# Office Offline (Android)

Aplikasi **Office Web** (Word, Excel, PowerPoint, PDF) yang dibungkus menjadi **APK Android** dengan
[Capacitor](https://capacitorjs.com). Seluruh kode dan library dimasukkan ke dalam APK, jadi aplikasi
**berjalan 100% offline** (tanpa CDN, tanpa server).

## Cara mendapatkan APK (tanpa Android Studio)

1. Buat repo baru di GitHub, lalu push seluruh isi folder ini ke branch `main`.
2. Buka tab **Actions** -> workflow **Build Android APK** (jalan otomatis tiap push; atau klik *Run workflow*).
3. Setelah selesai (± 5-8 menit), buka tab **Releases** dan unduh `Office-Offline.apk` (atau dari **Artifacts** pada run tersebut).
4. Kirim APK ke HP, lalu pasang (izinkan "Pasang dari sumber tidak dikenal").

APK juga otomatis dipublikasikan di tab **Releases** pada setiap push.

## Struktur repo

| Path | Isi |
|---|---|
| `www/` | Aplikasi web (index.html, `word/`, `excel/`, `ppt/`, `pdf/`) |
| `www/lib/` | Semua library pihak ketiga (lokal, untuk offline) |
| `www/native-bridge.js` | Jembatan WebView: simpan berkas ke Download, cetak, tombol Back |
| `android/` | Proyek Android (Capacitor) + plugin native `FileSaverPlugin.java` |
| `scripts/vendor-libs.js` | Menyalin ulang library dari `node_modules` ke `www/lib` |
| `.github/workflows/build-apk.yml` | Workflow build APK |

## Perilaku khusus di Android

- **Simpan** (`.docx`, `.xlsx`, `.pptx`, `.pdf`, ...) -> tersimpan otomatis di folder **Download** perangkat
  (Android 9 ke bawah: `Android/data/com.office.offline/files/Download`).
- **Cetak** membuka dialog cetak Android (pilih *Simpan sebagai PDF* untuk membuat PDF).
- **Buka berkas** memakai pemilih berkas bawaan Android.
- Fitur yang memang butuh internet (gambar dari URL, video YouTube, *Dari Web* di Excel, pencarian
  Wikipedia/terjemahan di PowerPoint) hanya jalan saat HP online; sisanya offline penuh.

## Update APK tanpa uninstall

Key penandatangan tetap: file `android/app/office.jks` ikut di repo (password ada di `android/app/build.gradle`),
dan `versionCode` otomatis naik mengikuti nomor run GitHub. Jadi setiap APK baru dari repo ini
**bisa langsung dipasang menimpa versi lama** tanpa uninstall dan tanpa kehilangan data.

- Setiap push ke `main` membuat **GitHub Release** baru (`v1.<nomor run>`) berisi `Office-Offline.apk`.
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
