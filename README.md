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

## Pembaruan di dalam aplikasi (menu Tentang)

Ketuk **ikon Alethra Office** di pojok kiri atas Beranda untuk membuka halaman **Tentang** (halaman penuh, bukan popup).
Cara kerjanya sama dengan Alethra Browser, lewat **GitHub Releases**:

- **Periksa pembaruan**: membandingkan nomor rilis terbaru (`v1.N`) dengan `versionCode` yang terpasang.
- **Unduh & pasang**: APK diunduh dengan bilah progres (bisa dibatalkan), lalu diperiksa ukuran, SHA-256, dan
  **tanda tangannya harus sama** dengan aplikasi terpasang sebelum pemasang Android dibuka. Android akan meminta izin
  *Pasang aplikasi tidak dikenal* sekali saja; setelah diizinkan, kembali ke aplikasi dan pemasangan dilanjutkan.
- **Lewati versi ini** dan sakelar **Periksa pembaruan otomatis** (maksimal sekali sehari saat aplikasi dibuka;
  bila ada versi baru muncul titik di ikon aplikasi dan sebuah notifikasi singkat).
- Alamat repositori ditanam saat build dari `GITHUB_REPOSITORY` (otomatis di GitHub Actions). Untuk build lokal isi
  `updateRepo=pemilik/repo` di `android/gradle.properties`.
- **Repositori harus publik**: pemeriksa memakai API GitHub tanpa login, jadi repositori privat dianggap "rilis tidak
  ditemukan". Bila ingin kode tetap privat, buat repositori publik khusus rilis dan isi `updateRepo` ke sana
  (APK tetap harus ditandatangani `office.jks` yang sama).

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

## Nama berkas saat menyimpan

- **Simpan pertama** dokumen baru (Word, Excel, PowerPoint, PDF) kini menampilkan dialog **"Simpan sebagai"** untuk mengisi nama
  (ekstensi otomatis). Judul dokumen ikut berubah ke nama yang dipilih.
- Menu **Berkas** punya dua tombol: **Simpan** (menimpa berkas yang sama tanpa bertanya) dan **Simpan Sebagai…** (selalu minta nama baru,
  hasilnya berkas baru di folder Download; simpan berikutnya menimpa berkas baru itu).
- Berkas yang dibuka dari daftar tetap ditimpa di lokasi aslinya. Jika nama sudah dipakai, Android menambahkan `(1)` agar tidak menimpa berkas lain.

## Rumus (persamaan) di Word

Rumus bawaan Microsoft Word (OMML / `<m:oMath>`) sekarang **tampil** saat `.docx` dibuka dan **tetap utuh** saat disimpan.

- **Buka**: Mammoth membuang rumus, jadi `www/word/omml.js` mengganti tiap rumus dengan token sebelum Mammoth membaca berkas,
  lalu token itu dipulihkan sebagai **MathML** (pecahan, akar, pangkat/indeks, ∑ ∏ ∫ ∮, kurung, matriks, fungsi, limit, aksen, dll.).
  Rumus tampil sebagai satu blok yang tidak ikut terketik (`contenteditable=false`).
- **Simpan**: OMML asli disimpan di atribut `data-omml` dan dimasukkan kembali ke `document.xml`, jadi rumus tetap bisa
  diedit di Microsoft Word.
- Butuh WebView Android dengan MathML (Chromium 109+, Android 7+ dengan WebView terbarui). Di WebView lama rumus
  ditampilkan sebagai teks linear, dan tetap utuh saat disimpan.
- Batasan: hanya isi utama dokumen yang didukung (bukan header/footer).

### Editor rumus v2 (palet visual, gaya, dan galeri)

Dialog **Sisipkan/Edit Rumus** dirombak supaya mudah dipakai orang awam tanpa hafal LaTeX:

- **Tombol mewakili fungsi aslinya**: tiap tombol menampilkan bentuk rumus sebenarnya dengan kotak isi □
  (mis. integral tentu lengkap dengan batas atas/bawah, sigma, limit, matriks, determinan). **Akar kuadrat (√) dan akar
  pangkat n (ⁿ√) kini berbeda**: yang kedua menampilkan "n" di pangkat akar.
- **Tab palet**: Dasar, Kalkulus, Aljabar (matriks 2×2/3×3, determinan, sistem persamaan, fungsi bersyarat), Trigonometri,
  Aksen, Yunani, Operator, Relasi, Panah, Himpunan, Statistik, Fisika & Geometri, Kimia, Lainnya, dan **Contoh**
  (±50 rumus siap pakai, dikelompokkan; ketuk untuk menyisipkan di kursor).
- **⇥ Kotak berikutnya** (atau tombol `Tab` di kotak rumus) melompat ke kotak `{}`/`[]` kosong berikutnya setelah
  menyisipkan templat.
- **Gaya rumus** (toolbar di bawah kotak rumus): **B** tebal, *I* miring, A tegak, warna rumus, warna sorot, ukuran (8-48 pt),
  ⌫ hapus gaya. Berlaku pada teks yang diblok; tanpa blok = seluruh rumus. Tombol tebal/miring berfungsi sebagai sakelar.
  Di LaTeX: `\textcolor{red}{..}` (nama Inggris/Indonesia atau hex), `\colorbox{..}{..}`, `\size{16}{..}`, `\color{..}`,
  `\large`, `\mathbf`, `\mathit`, `\boldsymbol`, serta panah berketerangan `\xrightarrow[bawah]{atas}`.
- **Bilah ketuk-rumus** kini punya **B · I · A▾ (warna & sorot)** untuk mengubah gaya rumus langsung tanpa membuka dialog
  (gaya ditulis langsung ke OMML, jadi bagian rumus lain tidak berubah).
- Gaya disimpan sebagai gaya Word asli (`w:color`, `w:sz`, `w:shd`, `m:sty`, `m:ctrlPr`), jadi tetap terbaca di Microsoft Word.
- **Rumus Umum ▾** di pita bertambah dan dikelompokkan (Geometri, Aljabar, Trigonometri, Kalkulus, Statistika, Fisika, Kimia).

### Sisipkan & edit rumus

Tab **SISIPKAN → grup Rumus**:

| Tombol | Fungsi |
|---|---|
| **π Rumus** (atau `Alt` + `=`) | Buka editor rumus untuk menyisipkan rumus baru di kursor. Bila ada teks terpilih, teks itu jadi isi awal. |
| **Edit Rumus** | Edit rumus yang sedang terpilih (termasuk rumus dari berkas Word). |
| **Rumus Umum ▾** | Galeri siap pakai: luas lingkaran, Pythagoras, rumus ABC, binomial, Taylor, Fourier, turunan, dll. |

Mengetuk rumus di halaman memunculkan bilah kecil **Edit · Sebaris/Blok · Hapus**; mengetuk dua kali langsung membuka editor.

Editor memakai penulisan **LaTeX** (`www/word/eqedit.js`) dengan pratinjau langsung dan palet tombol (struktur, huruf Yunani,
operator, relasi, panah). Tombol struktur menyisipkan templat dan menaruh kursor di slot pertama; teks yang dipilih di
kotak masuk ke slot itu. Contoh: `\frac{a}{b}`, `\sqrt[3]{x}`, `\sum_{i=1}^{n} i^2`, `\int_0^1 x\,dx`,
`\begin{pmatrix} a & b \\ c & d \end{pmatrix}`, `\begin{cases} x & x>0 \\ -x & x\le 0 \end{cases}`.

- Hasilnya OMML standar Word (`<m:oMath>` / `<m:oMathPara>`, strukturnya sama dengan keluaran pandoc dan lolos validasi skema OOXML),
  bukan gambar, sehingga tersimpan sebagai rumus sungguhan di `.docx`.
- **Sebaris** = di dalam kalimat; **Blok** = di paragraf sendiri, rata tengah. Rumus beberapa baris (`\\`) otomatis blok.
- Saat membuka rumus yang sudah ada, OMML diubah ke LaTeX untuk diedit. Bila isinya **tidak diubah**, OMML asli dipakai apa adanya
  (tidak ada yang berubah). Bila rumus memuat elemen yang tak punya padanan di editor (mis. `m:box`), muncul catatan bahwa
  bentuknya bisa sedikit berubah setelah diedit.
- Aksen/huruf khusus dipetakan ke karakter Unicode (`\mathbb{R}` → ℝ), sehingga tampil benar di Word.
- Catatan teknis: `insertHTML` di Chromium memecah paragraf bila fragmennya berisi `<math>` dan melempar `<span contenteditable=false>`
  ke luar `<p>`. Karena itu editor menyisipkan "cangkang" `<span class="eq">` dibungkus spasi lebar-nol (dibuang saat disimpan),
  lalu mengisi MathML lewat DOM. Undo/redo tetap berfungsi.
