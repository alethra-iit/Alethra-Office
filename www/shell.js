/* Shell aplikasi: Beranda + tab Word / Excel / PPoint / PDF dengan daftar berkas perangkat.
   Editor tiap format (word/excel/ppt/pdf) dimuat saat berkas dibuka atau dibuat baru. */
(function(){
'use strict';

const N = window.OfficeNative || { isNative:false, paths:{}, toast:function(){} };
const Files = N.Files;
const $ = (s, r) => (r || document).querySelector(s);
const toast = m => N.toast(m);
window.Apps = window.Apps || {};

/* ------------------------------------------------------------------ data dasar */
const FM = {
  word:  { id:'word',  nav:'Word',   name:'Word',       c:'#2b579a', l:'W',   accept:'.docx,.html,.htm,.txt' },
  excel: { id:'excel', nav:'Excel',  name:'Excel',      c:'#217346', l:'X',   accept:'.xlsx,.csv' },
  ppt:   { id:'ppt',   nav:'PPoint', name:'PowerPoint', c:'#d24726', l:'P',   accept:'.pptx' },
  pdf:   { id:'pdf',   nav:'PDF',    name:'PDF',        c:'#c4262e', l:'PDF', accept:'.pdf' }
};
const TYPE_BY_EXT = { docx:'word', xlsx:'excel', pptx:'ppt', pdf:'pdf' };
const MIME = {
  docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf:'application/pdf'
};
const TYPE_LABEL = { docx:'Dokumen Word (.docx)', xlsx:'Lembar kerja Excel (.xlsx)', pptx:'Presentasi PowerPoint (.pptx)', pdf:'Dokumen PDF (.pdf)' };
const KEYS = { files:'ofc.files.v1', favs:'ofc.favs.v1', recents:'ofc.recents.v1', sort:'ofc.sort.v1' };
const PAGE = 300;

const LS = {
  get(k, d){ try{ const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); }catch(e){ return d; } },
  set(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){} }
};

/* ------------------------------------------------------------------ ikon */
const svg = (body, vb) => `<svg viewBox="${vb || '0 0 24 24'}" xmlns="http://www.w3.org/2000/svg">${body}</svg>`;
const S1 = 'fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"';
const IC = {
  logo: svg('<path d="M12 1.5 22 6 12 10.5 2 6z" fill="#4277c9"/><path d="M2 7.4 12 11.9 22 7.4V9.6L12 14.1 2 9.6z" fill="#3aa66b"/><path d="M2 11 12 15.5 22 11V13.2L12 17.7 2 13.2z" fill="#e5694a"/><path d="M2 14.6 12 19.1 22 14.6V16.8L12 21.3 2 16.8z" fill="#c4262e"/>'),
  home: svg('<path d="M12 3.2 2.8 11.4h2.6V20h5.1v-5.6h3V20h5.1v-8.6h2.6z" fill="currentColor"/>'),
  word: svg('<rect x="9.5" y="3.5" width="11" height="17" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M13 8.5h5M13 12h5M13 15.5h5" stroke="currentColor" stroke-width="1.5"/><rect x="2.5" y="7" width="10" height="10" rx="1.4" fill="currentColor"/><text x="7.5" y="14.6" font-size="7.6" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">W</text>'),
  excel: svg('<rect x="9.5" y="3.5" width="11" height="17" rx="1.6" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M13 8.5h5M13 12h5M13 15.5h5" stroke="currentColor" stroke-width="1.5"/><rect x="2.5" y="7" width="10" height="10" rx="1.4" fill="currentColor"/><text x="7.5" y="14.6" font-size="7.6" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">X</text>'),
  ppt: svg('<circle cx="14.5" cy="12" r="7.3" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M14.5 4.7v7.3h7.3" fill="none" stroke="currentColor" stroke-width="1.6"/><rect x="2.5" y="7" width="10" height="10" rx="1.4" fill="currentColor"/><text x="7.5" y="14.6" font-size="7.6" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">P</text>'),
  pdf: svg('<path d="M6.2 20.4c3.4-1.2 6.6-6 7.6-12.4.5-3.2-1.9-3.4-2.1-1-.4 4.7 2.9 8.9 8 8.9 2.2 0 2.1-2-.2-2-3.4 0-9 .9-12.7 7.1-.8 1.2.1 1.9 1 1.4z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>'),
  search: svg(`<circle cx="10.5" cy="10.5" r="6.3" ${S1}/><path d="m15.4 15.4 5 5" ${S1}/>`),
  sort: svg(`<path d="M4 7h16M4 12h10.5M4 17h5.5" ${S1} stroke-width="2.2"/>`),
  back: svg(`<path d="M15 5 8 12l7 7" ${S1} stroke-width="2.2"/>`),
  plus: svg(`<path d="M12 5v14M5 12h14" ${S1} stroke-width="2.4"/>`),
  folder: svg(`<path d="M3.5 7.5a1.5 1.5 0 0 1 1.5-1.5h4l2 2.2h8A1.5 1.5 0 0 1 20.5 9.7V17A1.5 1.5 0 0 1 19 18.5H5A1.5 1.5 0 0 1 3.5 17z" ${S1}/>`),
  appicon: svg('<rect x="2" y="2" width="9.6" height="9.6" rx="2" fill="#2b579a"/><text x="6.8" y="9.3" font-size="7.2" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">W</text><rect x="12.4" y="2" width="9.6" height="9.6" rx="2" fill="#217346"/><text x="17.2" y="9.3" font-size="7.2" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">X</text><rect x="2" y="12.4" width="9.6" height="9.6" rx="2" fill="#d24726"/><text x="6.8" y="19.9" font-size="7.2" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">P</text><rect x="12.4" y="12.4" width="9.6" height="9.6" rx="2" fill="#c4262e"/><text x="17.2" y="19.9" font-size="7.2" font-weight="800" fill="#fff" text-anchor="middle" font-family="Arial,sans-serif">A</text>'),
  docs: svg(`<circle cx="12" cy="12" r="9.2" ${S1} stroke-width="2"/><text x="12" y="16.6" font-size="12.5" font-weight="500" fill="currentColor" text-anchor="middle" font-family="Arial,sans-serif">D</text>`),
  history: svg(`<path d="M3.5 12a8.5 8.5 0 1 0 2.6-6.1" ${S1} stroke-width="2.1"/><path d="M3.2 3.8v4.9h4.9" ${S1} stroke-width="2.1"/><path d="M12 7.6V12l3.2 2" ${S1} stroke-width="2.1"/>`),
  star: svg('<path d="m12 2.8 2.8 5.9 6.4.8-4.7 4.4 1.2 6.4L12 17.2 6.3 20.3l1.2-6.4L2.8 9.5l6.4-.8z" fill="currentColor"/>'),
  storage: svg('<rect x="3.5" y="4.2" width="17" height="4.4" rx="2.2" fill="currentColor"/><rect x="3.5" y="9.8" width="17" height="4.4" rx="2.2" fill="currentColor"/><rect x="3.5" y="15.4" width="17" height="4.4" rx="2.2" fill="currentColor"/>'),
  dots: svg('<circle cx="12" cy="5.2" r="2" fill="currentColor"/><circle cx="12" cy="12" r="2" fill="currentColor"/><circle cx="12" cy="18.8" r="2" fill="currentColor"/>'),
  inbox: svg('<path d="M12 2.5v3M7.2 4.2l1.6 2.2M16.8 4.2l-1.6 2.2" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" fill="none"/><path d="M2.8 13.2 5.6 8.4A1.6 1.6 0 0 1 7 7.6h10a1.6 1.6 0 0 1 1.4.8l2.8 4.8v6.2a1.6 1.6 0 0 1-1.6 1.6H4.4a1.6 1.6 0 0 1-1.6-1.6z" fill="currentColor"/><path d="M2.8 13.2h5a1.5 1.5 0 0 1 1.5 1.5 2.7 2.7 0 0 0 5.4 0 1.5 1.5 0 0 1 1.5-1.5h5" fill="none" stroke="var(--bg)" stroke-width="1.5"/>'),
  openwith: svg(`<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" ${S1}/>`),
  rename: svg(`<path d="M3.5 8.5h11v7h-11zM18 5.5v13M16 5.5h4M16 18.5h4" ${S1}/>`),
  starplus: svg(`<path d="m12 3.5 2.4 5 5.4.7-4 3.8 1 5.4L12 15.7 7.2 18.4l1-5.4-4-3.8 5.4-.7z" ${S1}/><path d="M19 15v6M16 18h6" ${S1}/>`),
  linkplus: svg(`<path d="M10 14a4 4 0 0 0 5.6 0l2.2-2.2a4 4 0 0 0-5.6-5.6l-1 1M14 10a4 4 0 0 0-5.6 0l-2.2 2.2a4 4 0 0 0 5.6 5.6l1-1" ${S1}/><path d="M19 16v6M16 19h6" ${S1}/>`),
  share: svg(`<circle cx="6" cy="12" r="2.6" ${S1}/><circle cx="17.5" cy="5.8" r="2.6" ${S1}/><circle cx="17.5" cy="18.2" r="2.6" ${S1}/><path d="m8.3 10.8 7-3.8M8.3 13.2l7 3.8" ${S1}/>`),
  bt: svg(`<path d="m6.5 7.5 11 9-5.5 4.5V3l5.5 4.5-11 9" ${S1} stroke-width="2"/>`),
  drive: svg(`<path d="M8.5 4h7l5 8.8-3.5 6.2H7L3.5 12.8z" ${S1}/><path d="M19 15v6M16 18h6" ${S1}/>`),
  trash: svg(`<path d="M4.5 7h15M9.5 7V4.2h5V7M6.5 7l1 13h9l1-13M10.2 11v6M13.8 11v6" ${S1}/>`),
  info: svg(`<rect x="4" y="3.5" width="16" height="17" rx="2" ${S1}/><path d="M12 11v5.5" ${S1} stroke-width="2.2"/><circle cx="12" cy="7.8" r="1.2" fill="currentColor"/>`),
  check: svg('<path d="m5 12.5 4.5 4.5L19 7.5" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>')
};

/* ------------------------------------------------------------------ util */
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
const extOf = n => { const i = String(n).lastIndexOf('.'); return i >= 0 ? String(n).slice(i + 1).toLowerCase() : ''; };
const dirOf = p => p.slice(0, p.lastIndexOf('/'));
const shortPath = p => String(p || '').replace(/^\/storage\/emulated\/0\/?/, '') || '/';
const nf = new Intl.NumberFormat('id-ID', { maximumFractionDigits: 1 });
function fmtSize(b){
  if (b < 1024) return b + ' B';
  if (b < 1048576) return nf.format(b / 1024) + ' KB';
  if (b < 1073741824) return nf.format(b / 1048576) + ' MB';
  return nf.format(b / 1073741824) + ' GB';
}
const fmtDate = ms => new Date(ms).toLocaleDateString('id-ID', { day:'2-digit', month:'2-digit', year:'numeric' });
const fmtDateTime = ms => new Date(ms).toLocaleString('id-ID', { dateStyle:'long', timeStyle:'short' });
const norm = o => ({ path:o.path, name:o.name, size:+o.size || 0, mtime:+o.mtime || 0, type:TYPE_BY_EXT[o.ext || extOf(o.name)] || null });

/* ------------------------------------------------------------------ state */
const S = {
  tab:'home', sub:'all', q:'', search:false, folder:null, limit:PAGE,
  sort: LS.get(KEYS.sort, 'date-desc'),
  files: [], favs: new Set(LS.get(KEYS.favs, [])), recents: LS.get(KEYS.recents, {}),
  access: null, scanning:false, inEditor:false
};
let byPath = new Map();
const indexFiles = () => { byPath = new Map(S.files.map(f => [f.path, f])); };
const saveFavs = () => LS.set(KEYS.favs, [...S.favs]);
const saveRecents = () => LS.set(KEYS.recents, S.recents);
function noteRecent(p){
  S.recents[p] = Date.now();
  const keys = Object.keys(S.recents).sort((a, b) => S.recents[b] - S.recents[a]);
  keys.slice(60).forEach(k => delete S.recents[k]);
  saveRecents();
}

const SORTS = {
  'name-asc':  (a, b) => a.name.localeCompare(b.name, 'id', { sensitivity:'base', numeric:true }),
  'name-desc': (a, b) => b.name.localeCompare(a.name, 'id', { sensitivity:'base', numeric:true }),
  'date-desc': (a, b) => b.mtime - a.mtime,
  'date-asc':  (a, b) => a.mtime - b.mtime,
  'size-desc': (a, b) => b.size - a.size,
  'size-asc':  (a, b) => a.size - b.size
};
const SORT_LABEL = {
  'date-desc':'Terbaru dulu', 'date-asc':'Terlama dulu', 'name-asc':'Nama (A-Z)',
  'name-desc':'Nama (Z-A)', 'size-desc':'Ukuran terbesar', 'size-asc':'Ukuran terkecil'
};

function visible(){
  let a;
  if (S.tab === 'home'){
    if (S.sub === 'recent') a = Object.keys(S.recents).sort((x, y) => S.recents[y] - S.recents[x]).map(p => byPath.get(p)).filter(Boolean);
    else if (S.sub === 'fav') a = [...S.favs].map(p => byPath.get(p)).filter(Boolean);
    else a = S.files.slice();
  } else a = S.files.filter(f => f.type === S.tab);
  if (S.folder) a = a.filter(f => dirOf(f.path) === S.folder);
  const q = S.q.trim().toLowerCase();
  if (q) a = a.filter(f => f.name.toLowerCase().includes(q));
  if (!(S.tab === 'home' && S.sub === 'recent')) a.sort(SORTS[S.sort] || SORTS['date-desc']);
  return a;
}

/* ------------------------------------------------------------------ render */
const hdr = $('#hdr'), subtabs = $('#subtabs'), list = $('#list'), bnav = $('#bnav'), scanbar = $('#scanbar');
const shell = $('#shell'), app = $('#app'), layer = $('#layer'), picker = $('#picker');

function renderHeader(){
  if (S.search){
    hdr.innerHTML = `<button class="ib" data-a="search-close" aria-label="Kembali">${IC.back}</button>` +
      `<input id="q" type="search" placeholder="Cari berkas" value="${esc(S.q)}" autocomplete="off" enterkeyhint="search">`;
    return;
  }
  const left = S.tab === 'home'
    ? `<div class="logo">${IC.logo}</div>`
    : `<div class="hl" style="--c:${FM[S.tab].c}">` +
        `<button class="pill fill" data-a="new">${IC.plus}<span>Baru</span></button>` +
        `<button class="pill" data-a="open">${IC.folder}<span>Buka</span></button></div>`;
  hdr.innerHTML = left + '<div class="sp"></div>' +
    `<button class="ib" data-a="search" aria-label="Cari">${IC.search}</button>` +
    `<button class="ib" data-a="sort" aria-label="Urutkan">${IC.sort}</button>` +
    `<button class="ib" data-a="about" aria-label="Tentang">${IC.appicon}</button>`;
}

function renderSubtabs(){
  subtabs.hidden = S.tab !== 'home';
  if (subtabs.hidden) return;
  const t = (k, ic, label) => `<button class="st${S.sub === k ? ' on' : ''}" data-sub="${k}" aria-label="${label}">${IC[ic]}</button>`;
  subtabs.innerHTML = t('all', 'docs', 'Semua dokumen') + t('recent', 'history', 'Terbaru') + t('fav', 'star', 'Favorit') +
    '<div class="sp"></div>' +
    `<button class="st${S.folder ? ' flt' : ''}" data-a="folder" aria-label="Lokasi">${IC.storage}</button>`;
}

function renderNav(){
  const items = [['home', 'Beranda', IC.home, '#e0390b']].concat(
    Object.values(FM).map(f => [f.id, f.nav, IC[f.id], f.c]));
  bnav.innerHTML = items.map(([id, label, icon, c]) =>
    `<button class="ni${S.tab === id ? ' on' : ''}" data-tab="${id}" style="--c:${c}"><span class="pb">${icon}</span><span>${label}</span></button>`).join('');
}

function row(f){
  const m = FM[f.type];
  return `<div class="row" data-p="${esc(f.path)}"><div class="fi${f.type === 'pdf' ? ' pdf' : ''}" style="background:${m.c}">${m.l}</div>` +
    `<div class="fb"><div class="fn">${esc(f.name)}</div><div class="fm">${fmtSize(f.size)} | ${fmtDate(f.mtime)}</div></div>` +
    `<button class="dots" data-m="${esc(f.path)}" aria-label="Menu">${IC.dots}</button></div>`;
}

function empty(msg, btn){
  return `<div class="empty">${IC.inbox}${msg ? `<p>${msg}</p>` : ''}${btn || ''}</div>`;
}

function renderList(){
  let html;
  if (!N.isNative){
    html = empty('Daftar berkas hanya tersedia di aplikasi Android. Gunakan tombol Baru / Buka.');
  } else if (S.access === false){
    html = empty('Izinkan akses ke semua berkas agar dokumen di perangkat Anda tampil di sini.',
      '<button class="btn" data-a="grant">Izinkan akses</button>');
  } else {
    const a = visible();
    if (!a.length){
      html = empty(S.q ? 'Tidak ada berkas yang cocok.' : (S.scanning ? 'Memindai berkas…' : (S.tab === 'home' && S.sub !== 'all' ? '' : 'Belum ada berkas.')));
    } else {
      html = a.slice(0, S.limit).map(row).join('') +
        (a.length > S.limit ? `<button class="more" data-a="more">Tampilkan lebih banyak (${a.length - S.limit})</button>` : '');
    }
  }
  list.innerHTML = html;
}

function renderScan(){ scanbar.hidden = !S.scanning; }
function render(){ renderHeader(); renderSubtabs(); renderList(); renderNav(); renderScan(); }

/* ------------------------------------------------------------------ navigasi */
function setTab(t){
  S.tab = t; S.q = ''; S.search = false; S.folder = null; S.limit = PAGE;
  render(); list.scrollTop = 0;
}
function closeSearch(){ S.search = false; S.q = ''; S.limit = PAGE; renderHeader(); renderList(); }

function act(a){
  switch (a){
    case 'new': start(S.tab, null); break;
    case 'open': pickFile(S.tab); break;
    case 'search': S.search = true; renderHeader(); { const i = $('#q'); if (i) i.focus(); } break;
    case 'search-close': closeSearch(); break;
    case 'sort': sortSheet(); break;
    case 'about': aboutDialog(); break;
    case 'folder': folderSheet(); break;
    case 'more': S.limit += PAGE; renderList(); break;
    case 'grant':
      N.Files.requestAccess().then(() => toast("Aktifkan 'Izinkan akses semua berkas', lalu kembali ke aplikasi")).catch(e => toast(e.message || 'Gagal'));
      break;
  }
}

shell.addEventListener('click', e => {
  const t = e.target.closest('[data-a],[data-tab],[data-sub],[data-m],.row');
  if (!t) return;
  if (t.dataset.m) return menu(t.dataset.m);
  if (t.dataset.tab) return setTab(t.dataset.tab);
  if (t.dataset.sub){ S.sub = t.dataset.sub; S.limit = PAGE; renderSubtabs(); renderList(); return; }
  if (t.dataset.a) return act(t.dataset.a);
  if (t.classList.contains('row')){ const f = byPath.get(t.dataset.p); if (f) openEntry(f); }
});
shell.addEventListener('input', e => {
  if (e.target.id === 'q'){ S.q = e.target.value; S.limit = PAGE; renderList(); }
});

/* ------------------------------------------------------------------ sheet & dialog */
function closeLayer(){ layer.innerHTML = ''; layer._pick = null; layer._ok = null; }

function sheet(title, items, onPick){
  layer.innerHTML = '<div class="dim" data-x="1"></div><div class="sheet">' +
    (title ? `<div class="s-t">${esc(title)}</div>` : '') +
    items.map((it, i) => `<button class="si" data-i="${i}">${it.icon || ''}<span class="sl"><span>${esc(it.label)}</span>${it.sub ? `<small>${esc(it.sub)}</small>` : ''}</span>${it.check ? `<span class="ck">${IC.check}</span>` : ''}</button>`).join('') +
    '</div>';
  layer._pick = i => { closeLayer(); onPick(items[i]); };
}

function dialog(o){
  layer.innerHTML = '<div class="dim" data-x="1"></div><div class="dlg">' +
    `<h3>${esc(o.title)}</h3><div class="dlg-b">${o.html || ''}</div><div class="dlg-a">` +
    (o.cancel === null ? '' : `<button class="tb" data-d="c">${o.cancel || 'Batal'}</button>`) +
    `<button class="tb pri" data-d="ok">${o.ok || 'OK'}</button></div></div>`;
  layer._ok = o.onOk || (() => {});
  const inp = layer.querySelector('input');
  if (inp){ inp.focus(); inp.select(); }
}

async function runOk(){
  const el = layer.querySelector('.dlg'), fn = layer._ok;
  let r;
  try { r = await fn(el); } catch(e){ r = false; }
  if (r !== false) closeLayer();
}

layer.addEventListener('click', e => {
  if (e.target.dataset.x) return closeLayer();
  const si = e.target.closest('.si');
  if (si && layer._pick) return layer._pick(+si.dataset.i);
  const d = e.target.closest('[data-d]');
  if (d) return d.dataset.d === 'ok' ? runOk() : closeLayer();
});
layer.addEventListener('keydown', e => { if (e.key === 'Enter' && layer._ok && e.target.tagName === 'INPUT') runOk(); });

function sortSheet(){
  const items = Object.keys(SORT_LABEL).map(k => ({ k, label:SORT_LABEL[k], check:S.sort === k }));
  sheet('Urutkan berdasarkan', items, it => { S.sort = it.k; LS.set(KEYS.sort, S.sort); S.limit = PAGE; renderList(); });
}

function folderSheet(){
  const cnt = new Map();
  S.files.forEach(f => { const d = dirOf(f.path); cnt.set(d, (cnt.get(d) || 0) + 1); });
  const items = [{ dir:null, label:`Semua lokasi (${S.files.length})`, check:!S.folder }];
  [...cnt.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40).forEach(([d, c]) =>
    items.push({ dir:d, label:`${d.split('/').pop() || '/'} (${c})`, sub:shortPath(d), check:S.folder === d }));
  sheet('Lokasi penyimpanan', items, it => { S.folder = it.dir; S.limit = PAGE; renderSubtabs(); renderList(); });
}

async function aboutDialog(){
  let ver = '';
  try { const i = await N.App.getInfo(); ver = `${i.version} (${i.build})`; } catch(e){}
  dialog({
    title:'Alethra Office',
    html:'<p>Word, Excel, PowerPoint, dan PDF yang berjalan sepenuhnya <b>offline</b> di perangkat Anda.</p>' +
         (ver ? `<p class="mut">Versi ${esc(ver)}</p>` : ''),
    ok:'Tutup', cancel:null
  });
}

/* ------------------------------------------------------------------ menu berkas */
function menu(p){
  const f = byPath.get(p); if (!f) return;
  const fav = S.favs.has(p);
  const items = [
    { k:'with',  icon:IC.openwith, label:'Buka dengan...' },
    { k:'rename',icon:IC.rename,   label:'Ubah Nama' },
    { k:'fav',   icon:IC.starplus, label:fav ? 'Hapus dari Favorit' : 'Tambahkan ke Favorit' },
    { k:'short', icon:IC.linkplus, label:'Buat jalan pintas' },
    { k:'share', icon:IC.share,    label:'Bagikan' },
    { k:'bt',    icon:IC.bt,       label:'Bagikan melalui Bluetooth' },
    { k:'drive', icon:IC.drive,    label:'Simpan ke Google Drive' },
    { k:'del',   icon:IC.trash,    label:'Hapus' },
    { k:'info',  icon:IC.info,     label:'Detail' }
  ];
  sheet(f.name, items, it => doMenu(it.k, f));
}

function dropEntry(p){
  S.files = S.files.filter(f => f.path !== p);
  S.favs.delete(p); delete S.recents[p];
  delete N.paths[p.split('/').pop().toLowerCase()];
  indexFiles(); saveFavs(); saveRecents(); LS.set(KEYS.files, S.files); renderList();
}

function movePath(oldP, newP, newName){
  const f = byPath.get(oldP); if (!f) return;
  delete N.paths[f.name.toLowerCase()];
  f.path = newP; f.name = newName;
  if (S.favs.delete(oldP)) S.favs.add(newP);
  if (S.recents[oldP]){ S.recents[newP] = S.recents[oldP]; delete S.recents[oldP]; }
  indexFiles(); saveFavs(); saveRecents(); LS.set(KEYS.files, S.files); renderList();
}

async function doMenu(k, f){
  const mime = MIME[extOf(f.name)] || '*/*';
  try {
    switch (k){
      case 'with': await Files.openWith({ path:f.path, mime }); break;
      case 'share': await Files.share({ path:f.path, mime }); break;
      case 'bt': {
        const r = await Files.share({ path:f.path, mime, pkg:'com.android.bluetooth' });
        if (r && r.fallback) toast('Bluetooth tidak ditemukan, pilih aplikasi dari daftar');
        break;
      }
      case 'drive': {
        const r = await Files.share({ path:f.path, mime, pkg:'com.google.android.apps.docs' });
        if (r && r.fallback) toast('Google Drive belum terpasang, pilih aplikasi lain');
        break;
      }
      case 'fav':
        if (S.favs.has(f.path)){ S.favs.delete(f.path); toast('Dihapus dari Favorit'); }
        else { S.favs.add(f.path); toast('Ditambahkan ke Favorit'); }
        saveFavs(); renderList();
        break;
      case 'short': {
        const r = await Files.pinShortcut({ path:f.path, name:f.name });
        toast(r && r.supported ? 'Konfirmasi penambahan jalan pintas di layar utama' : 'Peluncur layar tidak mendukung jalan pintas');
        break;
      }
      case 'rename': {
        const ext = extOf(f.name), base = f.name.slice(0, f.name.length - ext.length - 1);
        dialog({
          title:'Ubah Nama', ok:'Ubah',
          html:`<input class="inp" value="${esc(base)}" maxlength="150" autocomplete="off">`,
          onOk: async el => {
            const v = el.querySelector('input').value.replace(/[\\/:*?"<>|]/g, '_').trim();
            if (!v) return false;
            const nn = v + '.' + ext;
            if (nn === f.name) return;
            try {
              const r = await Files.rename({ path:f.path, newName:nn });
              movePath(f.path, r.path, r.name);
              toast('Nama diubah');
            } catch(e){ toast(e.message || 'Gagal mengubah nama'); return false; }
          }
        });
        break;
      }
      case 'del':
        dialog({
          title:'Hapus berkas?', ok:'Hapus',
          html:`<p>${esc(f.name)}</p><p class="mut">Berkas akan dihapus permanen dari perangkat.</p>`,
          onOk: async () => {
            try { await Files.delete({ path:f.path }); dropEntry(f.path); toast('Berkas dihapus'); }
            catch(e){ toast(e.message || 'Gagal menghapus'); return false; }
          }
        });
        break;
      case 'info':
        dialog({
          title:'Detail', ok:'Tutup', cancel:null,
          html:'<table class="dt">' +
            `<tr><td>Nama</td><td>${esc(f.name)}</td></tr>` +
            `<tr><td>Jenis</td><td>${esc(TYPE_LABEL[extOf(f.name)] || extOf(f.name))}</td></tr>` +
            `<tr><td>Lokasi</td><td>${esc(shortPath(dirOf(f.path)))}</td></tr>` +
            `<tr><td>Ukuran</td><td>${fmtSize(f.size)} (${f.size.toLocaleString('id-ID')} byte)</td></tr>` +
            `<tr><td>Diubah</td><td>${esc(fmtDateTime(f.mtime))}</td></tr></table>`
        });
        break;
    }
  } catch(e){ toast((e && e.message) || 'Terjadi kesalahan'); }
}

/* ------------------------------------------------------------------ editor */
const loaded = {};
function loadModule(id){
  if (window.Apps[id]) return Promise.resolve();
  return loaded[id] || (loaded[id] = new Promise((ok, err) => {
    const s = document.createElement('script');
    s.src = id + '/' + id + '.js'; s.onload = ok;
    s.onerror = () => { delete loaded[id]; err(new Error('Gagal memuat ' + s.src)); };
    document.head.appendChild(s);
  }));
}

async function start(id, file){
  try { await loadModule(id); }
  catch(e){ return toast(e.message); }
  S.inEditor = true; shell.hidden = true; app.hidden = false; app.innerHTML = '';
  /* Tiap editor dipasang di wadah baru. Editor memasang listener di elemen root-nya;
     bila root dipakai ulang (#app), listener menumpuk tiap kali dokumen dibuka dan satu ketukan
     diproses berkali-kali (ribbon terbuka lalu langsung menutup, tombol catatan terbalik dua kali). */
  const host = document.createElement('div'); host.style.height = '100%'; app.appendChild(host);
  window.Apps[id].mount(host, { file, onHome:goHome });
  history.replaceState(null, '', '#' + id);
}

function goHome(){
  app.hidden = true; app.innerHTML = ''; shell.hidden = false; S.inEditor = false;
  document.title = 'Alethra Office';
  history.replaceState(null, '', location.pathname);
  render(); rescan();
}

function pickFile(id){
  picker.accept = FM[id].accept;
  picker.onchange = () => { const f = picker.files[0]; picker.value = ''; if (f) start(id, f); };
  picker.click();
}

function b64File(b64, name, mime){
  const bin = atob(b64), u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return new File([u], name, { type:mime || '' });
}

async function fileFromEntry(f){
  const mime = MIME[extOf(f.name)] || '';
  try {
    const r = await fetch(window.Capacitor.convertFileSrc(f.path));
    if (!r.ok) throw new Error('http ' + r.status);
    const b = await r.blob();
    if (!b.size && f.size) throw new Error('kosong');
    return new File([b], f.name, { type:mime, lastModified:f.mtime || Date.now() });
  } catch(e){
    const d = await Files.readFile({ uri:f.path });
    return b64File(d.data, f.name, mime);
  }
}

async function openEntry(f){
  try {
    toast('Membuka ' + f.name + '…');
    const file = await fileFromEntry(f);
    N.paths[f.name.toLowerCase()] = f.path;
    noteRecent(f.path);
    await start(f.type, file);
  } catch(e){ toast('Gagal membuka: ' + ((e && e.message) || e)); }
}

/* dibuka dari aplikasi lain / jalan pintas */
const typeFromMime = m => { const k = Object.keys(MIME).find(x => MIME[x] === m); return k ? TYPE_BY_EXT[k] : null; };
async function handleLaunch(r){
  let name = r.name || 'berkas';
  let ext = extOf(name);
  if (!TYPE_BY_EXT[ext] && r.mime){
    const k = Object.keys(MIME).find(x => MIME[x] === r.mime);
    if (k){ ext = k; name += '.' + k; }
  }
  const type = TYPE_BY_EXT[ext] || typeFromMime(r.mime);
  if (!type) return toast('Jenis berkas tidak didukung');
  try {
    toast('Membuka ' + name + '…');
    let file;
    if (r.path){
      const f = byPath.get(r.path) || { path:r.path, name, size:0, mtime:0, type };
      file = await fileFromEntry(f);
      N.paths[name.toLowerCase()] = r.path;
      noteRecent(r.path);
    } else {
      const d = await Files.readFile({ uri:r.uri });
      file = b64File(d.data, name, MIME[ext]);
    }
    await start(type, file);
  } catch(e){ toast('Gagal membuka: ' + ((e && e.message) || e)); }
}
async function checkLaunch(){
  if (!Files) return;
  try { const r = await Files.getLaunch(); if (r && (r.path || r.uri)) handleLaunch(r); } catch(e){}
}

/* ------------------------------------------------------------------ pindai berkas */
let scanning = false;
async function rescan(){
  if (!N.isNative || scanning) return;
  scanning = true;
  try {
    S.access = await N.checkAccess();
    if (!S.access){ S.files = []; indexFiles(); render(); return; }
    S.scanning = true; renderScan(); if (!S.files.length) renderList();
    const r = await Files.list();
    S.files = (r.files || []).map(norm).filter(f => f.type);
    indexFiles(); LS.set(KEYS.files, S.files);
  } catch(e){
    if (e && e.code !== 'no_access') toast('Gagal memindai berkas');
  } finally {
    scanning = false; S.scanning = false;
    if (!S.inEditor) render();
  }
}

window.addEventListener('office:saved', e => {
  const d = e.detail || {};
  if (!d.path) return;
  const ex = byPath.get(d.path);
  if (ex){ ex.size = +d.size || ex.size; ex.mtime = +d.mtime || Date.now(); }
  else { const o = norm(d); if (o.type) S.files.push(o); }
  indexFiles(); noteRecent(d.path); LS.set(KEYS.files, S.files);
});

/* ------------------------------------------------------------------ tombol Back & siklus hidup */
if (N.App && N.App.addListener){
  N.App.addListener('backButton', () => {
    if (layer.innerHTML){ closeLayer(); return; }
    if (S.inEditor){ toast('Gunakan menu Berkas > Kembali ke menu utama'); return; }
    if (S.search){ closeSearch(); return; }
    if (S.tab !== 'home'){ setTab('home'); return; }
    N.App.exitApp();
  });
  N.App.addListener('resume', () => { if (!S.inEditor) rescan(); checkLaunch(); });
}
if (Files && Files.addListener){ try { Files.addListener('launch', () => checkLaunch()); } catch(e){} }

/* ------------------------------------------------------------------ mulai */
if (N.isNative){
  S.files = LS.get(KEYS.files, []).filter(f => f && f.path && f.type);
  indexFiles();
}
render();
if (N.isNative){ rescan(); checkLaunch(); }
const h0 = location.hash.slice(1);
if (FM[h0] && !N.isNative) start(h0, null);

})();
