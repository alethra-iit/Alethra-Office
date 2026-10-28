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
  const FileSaver = reg('FileSaver'), Files = reg('Files'), App = reg('App'), Updater = reg('Updater');
  if (!FileSaver || !Files) return;
  N.isNative = true; N.FileSaver = FileSaver; N.Files = Files; N.App = App; N.Updater = Updater;

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

  N.named = {};      /* nama berkas yang sudah dikonfirmasi pengguna / sudah punya lokasi */
  N.forceAs = false; /* true = tombol "Simpan Sebagai..." ditekan: minta nama baru */

  /* dialog nama berkas (muncul saat simpan pertama dan saat "Simpan Sebagai...") */
  function askName(name){
    return new Promise(ok => {
      const m = /(\.[A-Za-z0-9]{1,5})$/.exec(name || ''), ext = m ? m[1] : '';
      const base = String(name || 'Dokumen').replace(/\.[A-Za-z0-9]{1,5}$/, '');
      const bg = document.createElement('div');
      bg.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483646;display:flex;align-items:center;justify-content:center;padding:16px;font:14px system-ui,sans-serif';
      const btn = 'font:inherit;padding:8px 16px;border-radius:6px;border:1px solid #999;background:#f3f3f3;color:#222;cursor:pointer';
      bg.innerHTML = '<div style="background:#fff;color:#222;border-radius:10px;padding:18px;width:min(92vw,380px);box-shadow:0 8px 30px rgba(0,0,0,.35)">' +
        '<div style="font-weight:600;font-size:16px;margin-bottom:10px">Simpan sebagai</div>' +
        '<div style="display:flex;align-items:center;gap:6px"><input type="text" autocomplete="off" spellcheck="false" style="flex:1;min-width:0;font:inherit;padding:8px;border:1px solid #999;border-radius:6px;background:#fff;color:#222"><span style="color:#555"></span></div>' +
        '<div style="font-size:12px;color:#666;margin-top:8px">Disimpan di folder Download</div>' +
        '<div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px"><button data-c style="' + btn + '">Batal</button><button data-o style="' + btn + ';background:#2b579a;border-color:#2b579a;color:#fff">Simpan</button></div></div>';
      const inp = bg.querySelector('input');
      inp.value = base; bg.querySelector('span').textContent = ext;
      const done = v => { bg.remove(); ok(v); };
      const go = () => {
        let v = inp.value.replace(/[\\/:*?"<>|]+/g, '').trim();
        if (!v){ inp.focus(); return; }
        if (ext && v.toLowerCase().endsWith(ext.toLowerCase())) v = v.slice(0, -ext.length);
        done(v + ext);
      };
      bg.querySelector('[data-c]').onclick = () => done(null);
      bg.querySelector('[data-o]').onclick = go;
      bg.addEventListener('keydown', e => { if (e.key === 'Enter'){ e.preventDefault(); go(); } else if (e.key === 'Escape'){ e.preventDefault(); done(null); } });
      bg.addEventListener('mousedown', e => { if (e.target === bg) done(null); });
      document.body.appendChild(bg);
      inp.focus(); inp.select();
    });
  }
  N.askName = askName;

  async function saveFromHref(href, name){
    try{
      const key = String(name || '').toLowerCase();
      const ask = N.forceAs || !(N.paths[key] || N.named[key]);
      const asNew = N.forceAs;
      N.forceAs = false;
      let outName = name || 'berkas';
      if (ask){
        const n = await askName(outName);
        if (!n){ toast('Penyimpanan dibatalkan'); return; }
        outName = n;
      }
      const blob = await (await fetch(href)).blob();
      const data = await blobToBase64(blob);
      let finalName = outName;
      if (await N.checkAccess()){
        const res = await Files.writeFile({ name: outName, data, path: asNew || (ask && outName !== name) ? '' : (N.paths[key] || '') });
        finalName = res.name || outName;
        N.paths[key] = res.path; N.paths[finalName.toLowerCase()] = res.path;
        toast('Tersimpan: ' + shortPath(res.path));
        window.dispatchEvent(new CustomEvent('office:saved', { detail: res }));
      } else {
        const res = await FileSaver.save({ name: outName, mime: blob.type || 'application/octet-stream', data });
        finalName = (res && res.name) || outName;
        toast('Tersimpan di Download: ' + finalName);
      }
      N.named[key] = true; N.named[finalName.toLowerCase()] = true;
      if (finalName !== name) window.dispatchEvent(new CustomEvent('office:renamed', { detail: { from: name, to: finalName } }));
    }catch(e){ toast('Gagal menyimpan: ' + (e && e.message || e)); }
  }
  /* tombol bertanda data-saveas ("Simpan Sebagai...") memaksa dialog nama */
  document.addEventListener('click', e => {
    if (e.target.closest && e.target.closest('[data-saveas]')) N.forceAs = true;
  }, true);

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
