/* Jembatan WebView Android <-> aplikasi.
   - Menyediakan window.OfficeNative (plugin native Files / FileSaver / App) untuk shell.js.
   - <a download href="blob:..."> tidak jalan di WebView, jadi disimpan lewat plugin native:
       * berkas yang dibuka dari daftar  -> ditimpa di lokasi aslinya
       * berkas baru                     -> disimpan di folder Download (lalu disimpan ulang ke berkas yang sama)
       * tanpa izin akses berkas         -> MediaStore (folder Download)
   - window.print() diarahkan ke dialog cetak Android (bisa "Simpan sebagai PDF").
   Di browser biasa (bukan Android) berkas ini hanya menyiapkan objek kosong. */
(function(){
  const C = window.Capacitor;
  const N = window.OfficeNative = { isNative:false, access:false, paths:{}, toast:function(m){ console.log(m); } };
  if (!C || !C.isNativePlatform || !C.isNativePlatform()) return;

  const reg = n => (C.registerPlugin ? C.registerPlugin(n) : (C.Plugins && C.Plugins[n]));
  const FileSaver = reg('FileSaver'), Files = reg('Files'), App = reg('App');
  if (!FileSaver || !Files) return;
  N.isNative = true; N.FileSaver = FileSaver; N.Files = Files; N.App = App;

  function toast(msg){
    let t = document.getElementById('nb-toast');
    if (!t){
      t = document.createElement('div'); t.id = 'nb-toast';
      t.style.cssText = 'position:fixed;left:50%;bottom:calc(88px + var(--sab,0px));transform:translateX(-50%);max-width:88%;background:#323130;color:#fff;padding:10px 16px;border-radius:6px;font:13px system-ui,sans-serif;z-index:2147483647;box-shadow:0 2px 8px #0005;display:none;text-align:center';
      document.body.appendChild(t);
    }
    t.textContent = msg; t.style.display = 'block';
    clearTimeout(t._h); t._h = setTimeout(() => t.style.display = 'none', 3500);
  }
  N.toast = toast;

  N.checkAccess = async function(){
    try { const r = await Files.checkAccess(); N.access = !!(r && r.granted); }
    catch(e){ N.access = false; }
    return N.access;
  };
  N.checkAccess();
  if (App && App.addListener) App.addListener('resume', () => N.checkAccess());

  function blobToBase64(blob){
    return new Promise((ok, err) => {
      const r = new FileReader();
      r.onload = () => ok(String(r.result).split(',')[1] || '');
      r.onerror = () => err(r.error);
      r.readAsDataURL(blob);
    });
  }
  const shortPath = p => String(p || '').replace(/^\/storage\/emulated\/0\//, '');

  async function saveFromHref(href, name){
    try{
      const blob = await (await fetch(href)).blob();
      const data = await blobToBase64(blob);
      const key = String(name || '').toLowerCase();
      if (await N.checkAccess()){
        const res = await Files.writeFile({ name: name || 'berkas', data, path: N.paths[key] || '' });
        N.paths[key] = res.path;
        toast('Tersimpan: ' + shortPath(res.path));
        window.dispatchEvent(new CustomEvent('office:saved', { detail: res }));
      } else {
        const res = await FileSaver.save({ name: name || 'berkas', mime: blob.type || 'application/octet-stream', data });
        toast('Tersimpan di Download: ' + ((res && res.name) || name));
      }
    }catch(e){ toast('Gagal menyimpan: ' + (e && e.message || e)); }
  }

  /* 1) a.click() lewat skrip (pola yang dipakai semua modul) */
  const origClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function(){
    if (this.hasAttribute('download') && /^(blob:|data:)/.test(this.href)){
      saveFromHref(this.href, this.getAttribute('download')); return;
    }
    return origClick.apply(this, arguments);
  };

  /* 2) klik asli pengguna pada tautan unduhan */
  document.addEventListener('click', e => {
    const a = e.target.closest && e.target.closest('a[download]');
    if (a && /^(blob:|data:)/.test(a.href)){ e.preventDefault(); saveFromHref(a.href, a.getAttribute('download')); }
  }, true);

  /* 3) cetak */
  window.print = function(){
    FileSaver.print({ title: document.title || 'Dokumen' }).catch(e => toast('Gagal mencetak: ' + (e && e.message || e)));
  };
})();
