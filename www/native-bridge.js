/* Jembatan WebView Android <-> aplikasi.
   - <a download href="blob:..."> tidak jalan di WebView, jadi disimpan lewat plugin native FileSaver ke folder Download.
   - window.print() tidak jalan di WebView, jadi diarahkan ke plugin native (Print Service Android: bisa "Simpan sebagai PDF").
   Di browser biasa (bukan Android) berkas ini tidak melakukan apa-apa. */
(function(){
  const C = window.Capacitor;
  if (!C || !C.isNativePlatform || !C.isNativePlatform()) return;
  const reg = n => (C.registerPlugin ? C.registerPlugin(n) : (C.Plugins && C.Plugins[n]));
  const P = reg('FileSaver');
  if (!P) return;

  function toast(msg){
    let t = document.getElementById('nb-toast');
    if (!t){
      t = document.createElement('div'); t.id = 'nb-toast';
      t.style.cssText = 'position:fixed;left:50%;bottom:28px;transform:translateX(-50%);max-width:88%;background:#323130;color:#fff;padding:10px 16px;border-radius:4px;font:13px system-ui,sans-serif;z-index:2147483647;box-shadow:0 2px 8px #0005;display:none';
      document.body.appendChild(t);
    }
    t.textContent = msg; t.style.display = 'block';
    clearTimeout(t._h); t._h = setTimeout(() => t.style.display = 'none', 3500);
  }

  function blobToBase64(blob){
    return new Promise((ok, err) => {
      const r = new FileReader();
      r.onload = () => ok(String(r.result).split(',')[1] || '');
      r.onerror = () => err(r.error);
      r.readAsDataURL(blob);
    });
  }

  async function saveFromHref(href, name){
    try{
      const blob = await (await fetch(href)).blob();
      const data = await blobToBase64(blob);
      const res = await P.save({ name: name || 'berkas', mime: blob.type || 'application/octet-stream', data });
      toast('Tersimpan di Download: ' + (res && res.name || name));
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
    P.print({ title: document.title || 'Dokumen' }).catch(e => toast('Gagal mencetak: ' + (e && e.message || e)));
  };

  /* 4) tombol Back Android: kembali ke beranda jika sedang di dalam aplikasi, selain itu keluar */
  const App = reg('App');
  if (App && App.addListener){
    App.addListener('backButton', () => {
      const appEl = document.getElementById('app');
      if (appEl && !appEl.hidden){ toast('Gunakan menu Berkas > Kembali ke menu utama'); }
      else App.exitApp();
    });
  }
})();
