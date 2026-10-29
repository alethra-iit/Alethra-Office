/* Word editor (gaya Word 2013). Mengekspos Apps.word.mount(root,{file,onHome}) */
(function(){
const CDN={
  mammoth:'lib/mammoth.browser.min.js',
  docx:'lib/docx.iife.js',
  jszip:'lib/jszip.min.js',
  docxio:'word/docx-io.js',
  omml:'word/omml.js',
  eqedit:'word/eqedit.js'
};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>err(new Error('Gagal memuat '+n));document.head.appendChild(s);}));
}

const CSS=`
.wd{--blue:#2b579a;--ln:#d4d4d4;--hf:"Calibri Light",Carlito,Calibri,sans-serif;--bf:Calibri,Carlito,"Segoe UI",sans-serif;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative}
.wd *{box-sizing:border-box}
.wd button,.wd select,.wd input{font:inherit;color:inherit}
.wd [hidden]{display:none!important}
.wd em{font-style:normal}
.wd .qat{height:26px;display:flex;align-items:center;justify-content:flex-start;position:relative;padding:0 4px 0 6px;font-size:13px}
.wd .qat .l{position:static;flex:none;display:flex;gap:0;align-items:center}
.wd .qat .fn{flex:1;min-width:0;max-width:none;margin:0 8px 0 4px;overflow-x:auto;overflow-y:hidden;white-space:nowrap;text-overflow:clip;text-align:center;scrollbar-width:none;touch-action:pan-x;-webkit-overflow-scrolling:touch;cursor:grab}
.wd .qat .fn::-webkit-scrollbar{display:none}
.wd .qat .w{background:var(--blue);color:#fff;font-weight:700;width:18px;height:16px;display:grid;place-items:center;font-size:11px;margin-right:2px}
.wd .qat button{background:none;border:0;cursor:pointer;width:22px;height:22px;color:#555}
.wd .qat button:hover{background:#e1e1e1}
.wd .tabs{display:flex;height:26px;border-bottom:1px solid var(--ln);background:#fff}
.wd .tab{background:none;border:1px solid transparent;border-bottom:0;padding:0 14px;cursor:pointer;height:26px;font-size:12px}
.wd .tab:hover{color:var(--blue)}
.wd .tab.on{border-color:var(--ln);background:#fff;color:var(--blue);height:26px}
.wd .tab.file{background:var(--blue);color:#fff;padding:0 22px;margin-right:6px;border:0}
.wd .ribbon{height:88px;background:#fff;border-bottom:1px solid var(--ln);display:flex;padding:3px 4px 2px;overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;white-space:nowrap;flex:none}
.wd .grp{display:flex;flex-direction:column;border-right:1px solid #e3e3e3;padding:0 6px}
.wd .gb{flex:none;height:64px;margin-bottom:3px;display:flex;gap:4px;align-items:center}
.wd .gn{text-align:center;font-size:11px;color:#666;height:15px;line-height:15px}
.wd .rows{display:flex;flex-direction:column;gap:2px}
.wd .row{display:flex;gap:2px;align-items:center;height:26px}
.wd .b{min-width:24px;height:24px;border:1px solid transparent;background:none;display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:0 4px;cursor:pointer}
.wd .b:hover{background:#e6e6e6;border-color:#cfcfcf}
.wd .b.on{background:#cde6f7;border-color:#92c0e0}
.wd .b.big{flex-direction:column;height:62px;min-width:48px;gap:3px;font-size:11px}
.wd .b.big em{font-size:24px;line-height:1.1}
.wd .b.rw{height:20px;width:auto;min-width:26px;justify-content:center}
.wd select{height:23px;border:1px solid #ababab;background:#fff}
.wd .clr{position:relative;min-width:32px;height:24px;display:inline-flex;gap:1px;align-items:center;justify-content:center;padding:0 2px;cursor:pointer}
.wd .clr:hover{background:#e6e6e6}
.wd .clr span{padding:0 3px;line-height:1.1}
.wd .clr input{position:absolute;left:0;bottom:0;width:0;height:0;opacity:0;pointer-events:none;border:0;padding:0}
.wd .ddw{display:inline-flex}
.wd .dd{gap:1px;padding:0 3px}
.wd .dd em{display:grid;place-items:center}.wd .ic{display:block}
.wd .dd.vl{justify-content:space-between;padding:0 6px;border-color:#ababab;background:#fff;height:23px}
.wd .dd.vl .v{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.wd .dd .ar,.wd .clr .ar{font-style:normal;font-size:8px;color:#666}
.wd .nl small{color:#777;font-size:10px}.wd .nl input{width:46px}.wd .ri{display:grid;place-items:center;width:24px;color:#444}
.wd .wpop{position:absolute;z-index:40;background:#fff;border:1px solid #bbb;box-shadow:0 2px 8px rgba(0,0,0,.28);padding:2px 0;min-width:140px;overflow-y:auto;font-size:12px}
.wd .wpop .mi{display:flex;align-items:center;gap:10px;padding:6px 14px 6px 12px;cursor:pointer;white-space:nowrap}
.wd .wpop .mi:hover{background:#e6eefb}.wd .wpop .mi.cur{background:#dce8f7;font-weight:600}.wd .wpop .mi b{margin-left:auto;color:var(--blue)}
.wd .wpop .pal{padding:6px}.wd .wpop .pr{display:flex;gap:3px;margin-bottom:3px}
.wd .wpop .sw{width:20px;height:20px;border:1px solid #bbb;padding:0;cursor:pointer}.wd .wpop .sw:hover{outline:2px solid var(--blue);outline-offset:1px}
.wd .gal{display:flex;gap:2px;width:600px;height:64px;border:1px solid var(--ln);padding:2px;overflow:hidden}
.wd .sty{flex:none;width:86px;height:58px;background:none;border:1px solid transparent;cursor:pointer;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:4px;overflow:hidden}
.wd .sty:hover{border-color:#f5c35d;background:#fdf1c9}
.wd .sty span{font:11pt Calibri,Carlito,sans-serif;white-space:nowrap}
.wd .sty small{font-size:11px;color:#555;white-space:nowrap}
.wd .rbar{height:20px;background:#fff;border-bottom:1px solid var(--ln);display:flex;justify-content:center;overflow:hidden;flex:none}
.wd .ruler{position:relative;height:20px;flex:none;font-size:9px;color:#555;user-select:none}
.wd .ruler .rs{position:absolute;top:4px;height:12px;background:#dfe3e8}
.wd .ruler .mid{position:absolute;top:4px;height:12px;background:#fff;border:1px solid #cfd4da}
.wd .ruler .tk{position:absolute;top:8px;width:1px;height:4px;background:#888}
.wd .ruler b{position:absolute;top:3px;transform:translateX(-50%);font-weight:400}
.wd .ruler .mk{position:absolute;top:12px;width:0;height:0;border:5px solid transparent;border-bottom:7px solid #777;cursor:ew-resize}
.wd .desk{flex:1;overflow:auto;padding:12px 0 24px;background:linear-gradient(#f8f8f8,#ececec)}
.wd .wsheet{width:max-content;min-width:100%;display:flex;justify-content:center}
.wd .paper{position:relative;background:#fff;border:1px solid #ababab;flex:none}
.wd .paper.bd{outline-offset:-24px}
.wd .paper.wm::after{content:var(--wmt,"DRAF");position:absolute;inset:0;display:grid;place-items:center;font:700 var(--wms,110px) Calibri,sans-serif;color:rgba(0,0,0,.08);transform:rotate(-45deg);pointer-events:none}
.wd .cm{position:absolute;width:30px;height:30px;border:0 solid #bbb;pointer-events:none}
.wd .cm.a{left:calc(var(--ml) - 30px);top:calc(var(--mt) - 30px);border-width:0 1px 1px 0}
.wd .cm.b{left:calc(100% - var(--mr));top:calc(var(--mt) - 30px);border-width:0 0 1px 1px}
.wd .cm.c{left:calc(var(--ml) - 30px);top:calc(100% - var(--mb));border-width:1px 1px 0 0}
.wd .cm.d{left:calc(100% - var(--mr));top:calc(100% - var(--mb));border-width:1px 0 0 1px}
.wd .page{--hc:#2f5496;outline:none;font:11pt/var(--lh,1.08) var(--bf);color:#000;word-wrap:break-word}
.wd .page.col{background-image:repeating-linear-gradient(to bottom,transparent 0,transparent calc(var(--ph) - 2px),#c9c9c9 calc(var(--ph) - 2px),#c9c9c9 var(--ph))}
.wd .pgap{user-select:none;pointer-events:none;display:flow-root}
.wd .hfl{position:absolute;inset:0;pointer-events:none}
.wd .hf{position:absolute;pointer-events:auto;font:10pt/1.2 var(--bf);color:#555;outline:none;overflow:hidden;min-height:18px;cursor:text;word-wrap:break-word}
.wd .hf:hover{outline:1px dashed #9ab}
.wd .hf:focus{outline:1px dashed #2b579a;background:rgba(43,87,154,.06);color:#000}
.wd .hf:empty:not(:focus):hover::before{content:attr(data-ph);color:#999}
/* ===== Responsif: tablet & ponsel ===== */
.wd .tabs{overflow-x:auto;scrollbar-width:none}.wd .tabs::-webkit-scrollbar{display:none}
.wd .tab{flex:none;white-space:nowrap}
.wd .ribbon{-webkit-overflow-scrolling:touch}

.wd.rc .ribbon{display:none}
.wd .dg{min-width:min(340px,92vw)}.wd .db input,.wd .db select{max-width:100%}
.wd .toast{max-width:calc(100% - 32px)}
@media (pointer:coarse){
 .wd .qat{height:36px}.wd .qat button{width:28px;height:32px;font-size:16px}
 .wd .tabs{height:36px}.wd .tab{height:36px;padding:0 16px;font-size:13px}.wd .tab.on{height:36px}
 .wd .ribbon{height:106px}.wd .gb{height:80px}.wd .gn{height:16px;line-height:16px}
 .wd .row{height:36px}.wd .b{min-width:34px;height:34px}.wd .b.rw{height:24px;min-width:34px}.wd .rows{gap:3px}
 .wd .b.big{height:78px;min-width:56px}
 .wd select{height:34px;font-size:16px}.wd .clr{min-width:40px;height:34px}.wd .dd.vl{height:34px}.wd .wpop .mi{padding:11px 18px 11px 14px;font-size:15px}.wd .wpop .sw{width:28px;height:28px}
 .wd .gal{height:84px}.wd .sty{height:78px}
 .wd .status{height:30px}.wd .status button{width:28px;height:28px;font-size:16px}
 .wd .menu{top:72px;bottom:30px}.wd .menu button{padding:16px 24px}
 .wd .find input[type=text],.wd .db input,.wd .db select,.wd input[type=number]{font-size:16px;height:34px}
 .wd .find button,.wd .dbt button,.wd .db button{height:34px;padding:0 14px}
}
@media (max-width:900px){
 .wd .rbar{display:none}
}
@media (max-width:700px){

 .wd .menu{width:min(260px,86vw)}
 .wd .status .s{gap:10px}.wd .status{font-size:10px;white-space:nowrap}.wd .status .s{white-space:nowrap;flex:none}.wd .status .s:first-child>span:nth-child(n+2){display:none}.wd .zr{width:56px!important;vertical-align:middle}.wd .status button{width:26px}.wd .status .zv{display:inline-block;min-width:28px;text-align:right}
 .wd .find{left:8px;right:8px;top:auto;bottom:40px;grid-template-columns:auto 1fr}
 .wd .find .fb{flex-wrap:wrap}
 .wd .sp{right:8px;top:80px;width:min(220px,80vw)}
 .wd .desk{padding:8px 0 16px}
}
@media (max-height:500px){.wd .desk{padding:6px 0 12px}}
@media print{.wd .hf{outline:none!important;background:none!important;color:#555!important}.wd .hf:empty::before{content:none!important}.wd .pgap{background:none!important}.wd .wsheet{zoom:1!important}}
.wd .page p{margin:0 0 var(--ps,8pt)}
.wd .page h1,.wd .g1{font:400 16pt var(--hf);color:var(--hc);margin:12pt 0 0}
.wd .page h2,.wd .g2{font:400 13pt var(--hf);color:var(--hc);margin:2pt 0 0}
.wd .page h3{font:400 12pt var(--hf);color:#1f3763;margin:2pt 0 0}
.wd .s-title{font:28pt/1 var(--hf);margin:0 0 8pt}
.wd .s-sub{color:#5a5a5a;letter-spacing:.75pt}
.wd .s-em1{font-style:italic;color:#404040}.wd .s-em2{font-style:italic}
.wd .s-em3{font-style:italic;font-weight:700;color:var(--hc,#2f5496)}.wd .s-strong{font-weight:700}
.wd .page p.s-nosp{margin:0}
.wd .sty .s-title{font-size:20pt}.wd .sty .g1,.wd .sty .g2{margin:0}
.wd .page table{border-collapse:collapse;margin:0 0 8pt}
.wd .page td,.wd .page th{border:1px solid #000;padding:2pt 5.4pt;min-width:60px;vertical-align:top}
.wd .page img{max-width:100%}
.wd .page .eq{display:inline-block;vertical-align:middle;user-select:all;-webkit-user-select:all}
.wd .page .eq.eqd{display:block;text-align:center;margin:6pt 0}
.wd .page .eq math{font-family:"Cambria Math","STIX Two Math","Noto Sans Math","Latin Modern Math",math;font-size:1.1em;color:inherit}
.wd .page blockquote{margin:8pt 36pt;font-style:italic;color:#404040}
.wd .page .fnote{font-size:9pt;margin:0}
.wd .page.marks p::after,.wd .page.marks h1::after,.wd .page.marks h2::after,.wd .page.marks h3::after,.wd .page.marks li::after{content:"¶";color:#2b579a;font-size:10pt}
.wd .find{position:absolute;right:24px;top:150px;z-index:5;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:10px;display:grid;grid-template-columns:auto 170px;gap:6px;align-items:center}
.wd .find input[type=text]{height:24px;border:1px solid #aaa;padding:0 4px}
.wd .find .fb{grid-column:1/3;display:flex;gap:4px;justify-content:flex-end}
.wd .find button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.wd .status{background:var(--blue);color:#fff;height:22px;display:flex;align-items:center;justify-content:space-between;padding:0 8px;font-size:11px;text-transform:uppercase;flex:none}
.wd .status .s{display:flex;gap:18px;align-items:center}
.wd .status button{background:none;border:0;color:#fff;cursor:pointer;width:20px;height:20px;font-size:13px;text-transform:none}
.wd .status button.on{background:rgba(255,255,255,.25)}
.wd .status input{width:130px}
.wd .menu{position:absolute;z-index:10;top:52px;left:0;bottom:22px;width:260px;background:#2b579a;color:#fff;padding:8px 0}
.wd .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.wd .menu button:hover{background:rgba(255,255,255,.18)}
.wd .menu hr{border:0;border-top:1px solid rgba(255,255,255,.3);margin:6px 0}
.wd .toast{position:absolute;right:16px;bottom:36px;background:#323130;color:#fff;padding:8px 14px;z-index:20;max-width:360px}
.wd .page a{color:#0563c1}
.wd .page .toc a,.wd .page .tof a,.wd .page .bib a{color:inherit;text-decoration:none}
.wd .page.fx1 :is(h1,h2,h3){text-shadow:1px 1px 2px rgba(0,0,0,.3)}
.wd .page.fx2 :is(h1,h2){border-bottom:1px solid var(--hc)}
.wd .page.ln{counter-reset:ln}
.wd .page.ln :is(p,h1,h2,h3,li){position:relative;counter-increment:ln}
.wd .page.ln :is(p,h1,h2,h3,li)::before{content:counter(ln);position:absolute;left:-34px;width:26px;text-align:right;font:9pt Calibri,sans-serif;color:#777}
.wd .page img.sel,.wd .page table.sel,.wd .page hr.sel{outline:2px solid #2b579a}
/* ===== Tab kontekstual (Format Tabel / Gambar / Garis) + gagang objek ===== */
.wd .tab.ctx{color:#9c4a00;background:#fff1e0;border-color:#f0c79a;font-weight:600}
.wd .tab.ctx.on{background:#fff;color:#9c4a00;border-bottom:2px solid #fff}
.wd .ribbon.cx{background:#fffaf3}
.wd .hnd{position:absolute;pointer-events:none;z-index:6}
.wd .hnd i{position:absolute;pointer-events:auto;touch-action:none;background:#fff;border:2px solid #2b579a;width:12px;height:12px;border-radius:50%;margin:-6px 0 0 -6px;font-style:normal}
.wd .hnd i[data-h=e]{left:100%;top:50%;cursor:ew-resize}
.wd .hnd i[data-h=s]{left:50%;top:100%;cursor:ns-resize}
.wd .hnd i[data-h=se]{left:100%;top:100%;cursor:nwse-resize}
.wd .hnd i[data-h=mv]{left:0;top:-26px;margin:0;width:24px;height:24px;border-radius:4px;display:grid;place-items:center;font-size:15px;line-height:1;cursor:move;color:#2b579a}
.wd .hnd.nomv i[data-h=mv],.wd .hnd.hr i[data-h=s],.wd .hnd.hr i[data-h=se]{display:none}
.wd .dropl{position:absolute;z-index:7;background:#2b579a;pointer-events:none}
@media (pointer:coarse){.wd .hnd i{width:20px;height:20px;margin:-10px 0 0 -10px}.wd .hnd i[data-h=mv]{width:32px;height:32px;top:-38px;margin:0;font-size:20px}}
@media print{.wd .hnd,.wd .dropl{display:none!important}}
.wd .page.marks .xe::after{content:" XE";color:#2b579a;font-size:8pt}
.wd .page.marks .ta::after{content:" TA";color:#7030a0;font-size:8pt}
.wd .page .xe,.wd .page .ta{border-bottom:1px dotted #999}
.wd input[type=number]{height:23px;border:1px solid #ababab;width:58px}
.wd .nl{display:flex;align-items:center;gap:3px}
.wd .ov{position:absolute;inset:0;background:rgba(0,0,0,.25);z-index:30;display:grid;place-items:center}
.wd .dg{background:#fff;border:1px solid #888;box-shadow:0 4px 18px rgba(0,0,0,.3);padding:14px 16px;min-width:340px;max-width:92%}
.wd .dg h3{margin:0 0 10px;font-size:14px;font-weight:600;color:var(--blue)}
.wd .db{display:flex;flex-direction:column;gap:8px;max-height:60vh;overflow:auto}
.wd .db input,.wd .db select{height:26px;border:1px solid #aaa;padding:0 4px}
.wd .db select[size]{height:auto}
.wd .dbt{display:flex;gap:6px;justify-content:flex-end;margin-top:12px}
.wd .dbt button,.wd .db button{height:26px;border:1px solid #aaa;background:#f3f3f3;padding:0 12px;cursor:pointer}
.wd .dbt .ok{background:var(--blue);color:#fff;border-color:var(--blue)}
.wd .sli{display:flex;justify-content:space-between;gap:8px;align-items:center;border-bottom:1px solid #eee;padding:3px 0}
.wd .sp{position:absolute;right:24px;top:150px;z-index:5;width:220px;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:8px}
.wd .spi{display:flex;justify-content:space-between;align-items:center;padding:3px 6px;cursor:pointer}
.wd .spi:hover{background:#e6f1fb}.wd .spi.on{background:#cde6f7}
.wd .spi button{background:none;border:0;cursor:pointer}
@media print{body *{visibility:hidden}.wd .paper,.wd .paper *{visibility:visible}.wd .cm{display:none}.wd .paper{position:absolute;left:0;top:0;border:0;zoom:1!important}}
/* ===== Label teks di ribbon, ikon SVG, popup tabel ===== */
.wd .b .lbl{display:block;font-size:11px;line-height:1.15;white-space:nowrap;text-align:center}
.wd .b .lbl .ar{margin-left:2px}
.wd .b.big{justify-content:flex-start;padding:3px 5px 0}
.wd .b.big em{display:grid;place-items:center;height:30px;font-size:inherit}
.wd .b.big .ic{width:28px;height:28px}
.wd .b.rw,.wd .b.sm{justify-content:flex-start;gap:5px;padding:0 6px;height:20px}
.wd .b.rw .ic,.wd .b.sm .ic{width:16px;height:16px}
.wd .b.rw em,.wd .b.sm em{display:grid;place-items:center}
.wd .rows>.ddw{display:flex}.wd .rows>.ddw>.b{flex:1}
.wd .clr .lbl2{font-size:11px;padding:0 3px;white-space:nowrap}
.wd .clr.big{flex-direction:column;justify-content:flex-start;height:62px;min-width:48px;gap:3px;padding:3px 5px 0}
.wd .clr.big>span:first-child{display:grid;place-items:center;padding:0 2px 1px}
.wd .clr.big .ic{width:26px;height:26px}
.wd .rows>.clr{justify-content:flex-start;height:20px;gap:4px;padding:0 6px}
.wd .rows>.clr .ic{width:16px;height:16px}
.wd .nlt{font-size:11px;color:#444;white-space:nowrap}
.wd .qat .ic{width:16px;height:16px;display:block;margin:auto}
.wd .menu button{display:flex;align-items:center;gap:12px}
.wd .menu .ic{width:20px;height:20px;flex:none}
.wd .status .ic{width:15px;height:15px;display:block;margin:auto}
.wd .tbg{padding:6px 8px 4px}
.wd .tbg .lb{margin:0 0 5px;text-align:center;color:#555;min-height:16px}
.wd .tbg .gr{display:grid;grid-template-columns:repeat(10,18px);gap:2px}
.wd .tbg .gr i{width:18px;height:18px;border:1px solid #aaa;background:#fff;display:block}
.wd .tbg .gr i.on{background:#cde6f7;border-color:#2b579a}
.wd .wpop hr{border:0;border-top:1px solid #ddd;margin:3px 0}
.wd .db label{display:flex;justify-content:space-between;align-items:center;gap:12px}
.wd .db label input[type=text]{width:72px}
@media (pointer:coarse){
 .wd .b.big .ic,.wd .clr.big .ic{width:30px;height:30px}
 .wd .b.big em{height:32px}
 .wd .b.rw,.wd .b.sm,.wd .rows>.clr{height:24px}
 .wd .clr.big{height:78px}
 .wd .qat .ic{width:18px;height:18px}
 .wd .tbg .gr{grid-template-columns:repeat(10,26px)}.wd .tbg .gr i{width:26px;height:26px}
}
`;

const FONTS=['Calibri','Arial','Times New Roman','Cambria','Georgia','Verdana','Courier New','Segoe UI','Tahoma','Trebuchet MS'];
const SIZES=[8,9,10,11,12,14,16,18,20,24,28,36,48,72];
const SZ={a4:[794,1123],letter:[816,1056],legal:[816,1344]};
const MG={normal:[96,96,96,96],narrow:[48,48,48,48],moderate:[96,72,96,72],wide:[96,192,96,192]};
const STATE=['bold','italic','underline','strikeThrough','subscript','superscript','insertUnorderedList','insertOrderedList','justifyLeft','justifyCenter','justifyRight','justifyFull'];
const CMDS=new Set([...STATE,'outdent','indent','undo','redo','removeFormat','insertHorizontalRule','selectAll','cut','copy']);
const STY=[['Normal','p',''],['Tanpa Spasi','p','s-nosp'],['Judul 1','h1',''],['Judul 2','h2',''],['Judul','p','s-title'],['Subjudul','p','s-sub'],
  ['Penekanan Halus','i','s-em1'],['Penekanan','i','s-em2'],['Penekanan Kuat','i','s-em3'],['Tebal','i','s-strong']];
const TABS=[['home','BERANDA'],['insert','SISIPKAN'],['design','DESAIN'],['layout','TATA LETAK'],['ref','REFERENSI'],['review','TINJAU'],['view','TAMPILAN']];

const esc=t=>t.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const ICS=b=>`<svg class="ic" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${b}</svg>`;
const ICT=(t,fs)=>`<text x="12" y="16.5" font-size="${fs||13}" font-weight="700" text-anchor="middle" fill="currentColor" stroke="none" font-family="Arial,sans-serif">${t}</text>`;
const ICPG='<rect x="4.5" y="3" width="15" height="18" rx="1.5"/>';
const DDI={
 case:ICS(ICT('Aa')),
 lh:ICS('<path d="M7 5v14M4 8l3-3 3 3M4 16l3 3 3-3M14 7h7M14 12h7M14 17h7"/>'),
 sym:ICS(ICT('Ω',15)),
 eqg:ICS(ICT('Σ',15)),
 hdrs:ICS(ICPG+'<path d="M4.5 8h15M8 5.6h8"/>'),
 ftrs:ICS(ICPG+'<path d="M4.5 16h15M8 18.4h8"/>'),
 pnum:ICS(ICPG+'<path d="M8.5 8h7M8.5 11h5"/><text x="12" y="19" font-size="5.5" font-weight="700" text-anchor="middle" fill="currentColor" stroke="none" font-family="Arial,sans-serif">#</text>'),
 thm:ICS('<path d="M12 3a9 9 0 1 0 0 18c1.6 0 2.1-1.1 1.5-2.1-.6-1 0-2 1.4-2H17a4 4 0 0 0 4-4c0-5-4-9.9-9-9.9z"/><circle cx="7.6" cy="11.2" r="1" fill="currentColor"/><circle cx="10.5" cy="7.2" r="1" fill="currentColor"/><circle cx="15" cy="7.6" r="1" fill="currentColor"/>'),
 hcol:ICS('<rect x="3" y="3" width="8" height="8" rx="1.5" fill="#4472c4" stroke="none"/><rect x="13" y="3" width="8" height="8" rx="1.5" fill="#ed7d31" stroke="none"/><rect x="3" y="13" width="8" height="8" rx="1.5" fill="#70ad47" stroke="none"/><rect x="13" y="13" width="8" height="8" rx="1.5" fill="#ffc000" stroke="none"/>'),
 fpair:ICS('<text x="2" y="18" font-size="16" font-weight="700" fill="currentColor" stroke="none" font-family="Georgia,serif">A</text><text x="13" y="18" font-size="12" fill="currentColor" stroke="none" font-family="Arial,sans-serif">a</text>'),
 pspace:ICS('<path d="M5 4h14M5 20h14M12 7v10M9 10l3-3 3 3M9 14l3 3 3-3"/>'),
 fx:ICS('<path d="M10 3l1.8 5.2L17 10l-5.2 1.8L10 17l-1.8-5.2L3 10l5.2-1.8zM18 15l.8 2.2L21 18l-2.2.8L18 21l-.8-2.2L15 18l2.2-.8z"/>'),
 wmk:ICS(ICPG+'<text x="12" y="14" font-size="6.5" font-weight="700" text-anchor="middle" fill="currentColor" stroke="none" font-family="Arial,sans-serif" transform="rotate(-35 12 12)">ABC</text>'),
 bsty:ICS('<rect x="3.5" y="3.5" width="17" height="17"/><rect x="7.5" y="7.5" width="9" height="9" stroke-dasharray="2 1.6"/>'),
 bw:ICS('<path d="M4 6h16" stroke-width="1"/><path d="M4 11.5h16" stroke-width="2.2"/><path d="M4 18h16" stroke-width="3.8"/>'),
 margin:ICS(ICPG+'<rect x="7.5" y="6.5" width="9" height="11" stroke-dasharray="2 1.6"/>'),
 orient:ICS('<rect x="3" y="4" width="8" height="12" rx="1"/><rect x="12" y="10" width="9" height="6" rx="1"/>'),
 pgsize:ICS('<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/>'),
 cols:ICS(ICPG+'<path d="M12 6v12M7.5 8.5h2.5M7.5 11.5h2.5M7.5 14.5h2.5M14 8.5h2.5M14 11.5h2.5M14 14.5h2.5"/>'),
 dir:ICS('<path d="M10 4a4 4 0 0 0 0 8V4M14 4v8M10 4h8"/><path d="M5 17h13M15 14l3 3-3 3"/>'),
 brk:ICS('<path d="M5 3h14v7H5zM5 14h14v7H5z"/><path d="M2 12h2M8 12h2M14 12h2M20 12h2"/>'),
 ln:ICS('<path d="M4 6h2M4 12h2M4 18h2M9 6h11M9 12h11M9 18h11"/>'),
 hyph:ICS(ICT('a-b',11)),
 pos:ICS('<rect x="3" y="3" width="18" height="18" rx="1.5"/><rect x="5.5" y="5.5" width="7" height="6" fill="currentColor" stroke="none"/><path d="M14.5 7h4M14.5 10h4M5.5 14.5h13M5.5 17.5h13"/>'),
 wrap:ICS('<rect x="4" y="8" width="7" height="7" fill="currentColor" stroke="none"/><path d="M4 4.5h16M4 19.5h16M13.5 8.5h6.5M13.5 11.5h6.5M13.5 14.5h6.5"/>'),
 algn:ICS('<path d="M4 3v18"/><rect x="7" y="6" width="12" height="4"/><rect x="7" y="14" width="7" height="4"/>'),
 rot:ICS('<path d="M20 12a8 8 0 1 1-3-6.2"/><path d="M20 4v5h-5"/>'),
 addtxt:ICS('<path d="M4 6h10M4 12h10M4 18h6M18 9v8M14 13h8"/>')
};
/* ===== Ikon SVG gaya Word 2013 (garis abu-abu + aksen biru Word) ===== */
const WB='#2b579a';
const ICX=(x,y,fs,t,o)=>{o=o||{};return `<text x="${x}" y="${y}" font-size="${fs}"${o.w?` font-weight="${o.w}"`:''}${o.i?' font-style="italic"':''} text-anchor="middle" fill="${o.f||'currentColor'}" stroke="none" font-family="${o.ff||'Arial,sans-serif'}">${t}</text>`;};
const WM={
 mnew:ICS(ICPG+'<path d="M12 9v6M9 12h6"/>'),
 mopen:ICS('<path d="M3.5 7h6l2 2h9v10h-17z"/>'),
 msave:ICS('<path d="M4.5 4.5h12l3 3v12h-15z"/><rect x="7.5" y="4.5" width="7" height="5"/><rect x="7.5" y="13.5" width="9" height="6"/>'),
 mprint:ICS('<rect x="6.5" y="3.5" width="11" height="5"/><rect x="3.5" y="8.5" width="17" height="8" rx="1"/><rect x="6.5" y="14" width="11" height="6.5"/>'),
 mback:ICS('<path d="M20 12H5M10 6l-6 6 6 6"/>'),
 edit:ICS('<rect x="5" y="3.5" width="14" height="17"/><path d="M8 8h8M8 12h8M8 16h5"/>'),
 read:ICS('<path d="M3 6c3-1 6-1 9 1.5 3-2.5 6-2.5 9-1.5v12c-3-1-6-1-9 1.5C9 17 6 17 3 18z"/><path d="M12 7.5v12"/>')
};
const WI={
 paste:ICS('<rect x="4.5" y="5" width="11" height="14.5" rx="1.2"/><rect x="7.5" y="3" width="5" height="3.6" rx=".8" fill="#fff"/><rect x="10.5" y="9.5" width="9.5" height="11.5" rx="1" fill="#fff" stroke="'+WB+'"/><path d="M12.8 13.5h5M12.8 16.5h5" stroke="'+WB+'"/>'),
 cut:ICS('<circle cx="7" cy="17.5" r="2.6"/><circle cx="17" cy="17.5" r="2.6"/><path d="M8.7 15.4L16 3.5M15.3 15.4L8 3.5"/>'),
 copy:ICS('<rect x="8.5" y="8.5" width="11.5" height="12" rx="1" fill="#fff"/><path d="M15.5 8.5v-4a1 1 0 0 0-1-1h-9.5a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h3.5"/>'),
 painter:ICS('<rect x="4" y="3.5" width="12" height="5" rx="1" fill="#ffd966" stroke="#8a6d00"/><path d="M16 6h3.2v5.5H11V14"/><rect x="9" y="14" width="4" height="7" rx="1.2" fill="#fff"/>'),
 grow:ICS('<path d="M3 19L8.5 5l5.5 14M5 14.5h7"/><path d="M18 11V4M15.2 6.8L18 4l2.8 2.8" stroke="'+WB+'"/>'),
 shrink:ICS('<path d="M3 19L8 7l5 12M4.8 15h6.4"/><path d="M18 4v7M15.2 8.2L18 11l2.8-2.8" stroke="'+WB+'"/>'),
 clear:ICS('<path d="M3 19L8 6l5 13M5 14.5h6"/><path d="M12.5 20.5h8M13 20.5l-2.8-2.8a1.4 1.4 0 0 1 0-2l5.5-5.5a1.4 1.4 0 0 1 2 0l3 3a1.4 1.4 0 0 1 0 2l-5.2 5.3" stroke="#c0455e" fill="#f9d3da"/>'),
 bold:ICS(ICX(12,18,18,'B',{w:700})),
 italic:ICS(ICX(12,18,18,'I',{i:1,ff:'Georgia,serif'})),
 underline:ICS(ICX(12,16.5,16,'U',{w:700})+'<path d="M6 21h12"/>'),
 strike:ICS(ICX(12,16,11,'abc')+'<path d="M3.5 12.5h17" stroke="'+WB+'"/>'),
 sub:ICS(ICX(7,15,12,'x',{w:700})+ICX(17,21,9,'2',{w:700,f:WB})),
 sup:ICS(ICX(7,19,12,'x',{w:700})+ICX(17,10,9,'2',{w:700,f:WB})),
 hilite:ICS('<path d="M14.5 3.5l6 6-9 9-4.2.4-.4-4.2z" fill="#fff"/><path d="M5 20.5h6"/>'),
 fcolor:ICS(ICX(12,17,18,'A',{w:700})),
 ul:ICS('<circle cx="5" cy="6.5" r="1.5" fill="currentColor" stroke="none"/><circle cx="5" cy="12" r="1.5" fill="currentColor" stroke="none"/><circle cx="5" cy="17.5" r="1.5" fill="currentColor" stroke="none"/><path d="M10 6.5h10M10 12h10M10 17.5h10"/>'),
 ol:ICS(ICX(5,8.8,6.5,'1',{w:700,f:WB})+ICX(5,14.3,6.5,'2',{w:700,f:WB})+ICX(5,19.8,6.5,'3',{w:700,f:WB})+'<path d="M10 6.5h10M10 12h10M10 17.5h10"/>'),
 outd:ICS('<path d="M4 4.5h16M11 9.5h9M11 14.5h9M4 19.5h16"/><path d="M8 9.2L4.2 12 8 14.8z" fill="'+WB+'" stroke="'+WB+'"/>'),
 ind:ICS('<path d="M4 4.5h16M11 9.5h9M11 14.5h9M4 19.5h16"/><path d="M4 9.2L7.8 12 4 14.8z" fill="'+WB+'" stroke="'+WB+'"/>'),
 marks:ICS(ICX(12,19,19,'¶',{f:WB})),
 shade:ICS('<path d="M4.5 11.5l6.5-6.5 7 7-6.5 6.5z" fill="#fff"/><path d="M19.8 14.6c1.2 1.6 1.7 2.5 1.7 3.2a1.7 1.7 0 0 1-3.4 0c0-.7.5-1.6 1.7-3.2z" fill="'+WB+'" stroke="none"/>'),
 find:ICS('<circle cx="10" cy="10" r="5.8"/><path d="M14.5 14.5l6 6" stroke-width="2.2"/>'),
 repl:ICS('<path d="M4 7h11M12 4l3 3-3 3" stroke="'+WB+'"/><path d="M20 17H9M12 14l-3 3 3 3"/>'),
 sall:ICS('<rect x="3.5" y="3.5" width="12" height="12" stroke-dasharray="2.4 1.8"/><path d="M11 11l9 3-4 1.5L14 20z" fill="#fff" stroke="'+WB+'"/>'),
 eq:ICS(ICT('π',17)),
 eqedit:ICS('<path d="M4.5 19.5l1-4 10.5-10.5 3 3L8.5 18.5z"/><path d="M14 7l3 3" stroke="'+WB+'"/>'),
 pgbrk:ICS('<path d="M6 3h12v7H6zM6 14h12v7H6z"/><path d="M2.5 12h3M9 12h2M13 12h2M18.5 12h3" stroke="'+WB+'"/>'),
 table:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><rect x="3.5" y="5" width="17" height="4.5" fill="#bcd0ec"/><path d="M3.5 9.5h17M3.5 14.2h17M9.2 5v14M14.8 5v14"/><rect x="3.5" y="5" width="17" height="14" stroke="'+WB+'"/>'),
 image:ICS('<rect x="3.5" y="4.5" width="17" height="15" rx="1" fill="#fff"/><circle cx="9" cy="9.5" r="1.8" fill="#f2b01e" stroke="none"/><path d="M3.5 17l5-5 4 4 3-3 5 5" stroke="#2b8a3e"/>'),
 link:ICS('<path d="M10 14a3.5 3.5 0 0 0 5 0l3.5-3.5a3.5 3.5 0 0 0-5-5L12 7" stroke="'+WB+'"/><path d="M14 10a3.5 3.5 0 0 0-5 0L5.5 13.5a3.5 3.5 0 0 0 5 5L12 17" stroke="'+WB+'"/>'),
 hr:ICS('<path d="M4 12h16" stroke-width="2.4"/><path d="M4 6.5h16M4 17.5h16" stroke-dasharray="2 2" stroke-width="1.2"/>'),
 date:ICS('<rect x="3.5" y="5" width="14" height="14" rx="1.2" fill="#fff"/><path d="M3.5 9h14M7.5 3.5v3M13.5 3.5v3"/><circle cx="17" cy="17" r="4.5" fill="#fff" stroke="'+WB+'"/><path d="M17 14.5V17l1.8 1.2" stroke="'+WB+'"/>'),
 toc:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M8 7.5h8" stroke="'+WB+'" stroke-width="2.2"/><path d="M8 11.5h8M8 15h8M8 18h5"/>'),
 fn:ICS(ICX(8,17,12,'ab',{w:700})+ICX(17.5,10.5,9,'1',{w:700,f:WB})),
 en:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M8 7.5h8M8 11h8"/><path d="M8 17h8" stroke="'+WB+'"/>'),
 nfn:ICS(ICX(8,15,12,'ab',{w:700})+ICX(17,9,9,'1',{w:700,f:WB})+'<path d="M17 14v6M14.5 17.5L17 20l2.5-2.5" stroke="'+WB+'"/>'),
 snotes:ICS('<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="2.8" stroke="'+WB+'"/>'),
 cite:ICS('<path d="M6 7h5v5c0 3-1.8 5-4.5 6M14 7h5v5c0 3-1.8 5-4.5 6" stroke="'+WB+'"/>'),
 src:ICS('<rect x="4" y="4" width="4.5" height="16" rx=".8"/><rect x="9.5" y="6" width="4.5" height="14" rx=".8" stroke="'+WB+'"/><path d="M15.5 7l4.2 1.1-3.3 12.4-4.2-1.1z"/>'),
 bib:ICS('<path d="M12 6.5C10 5 7 4.5 4 5v13c3-.5 6 0 8 1.5 2-1.5 5-2 8-1.5V5c-3-.5-6 0-8 1.5zM12 6.5v13"/>'),
 cap:ICS('<rect x="3.5" y="4" width="17" height="11" rx="1" fill="#fff"/><path d="M3.5 13l4.5-4.5 3.5 3.5 3-3 6 6"/><circle cx="16.5" cy="7.8" r="1.4" fill="#f2b01e" stroke="none"/><path d="M6 19h12" stroke="'+WB+'" stroke-width="2"/>'),
 tof:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><rect x="7" y="6" width="5" height="4" stroke="'+WB+'"/><path d="M14 7h3M14 9.5h3M7 13.5h10M7 17h10"/>'),
 xref:ICS('<path d="M4 6h9M4 11h6M4 16h9"/><path d="M14.5 13.5l5.5 5.5M20 14v5h-5" stroke="'+WB+'"/>'),
 mxe:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M8 8h8M8 12h4"/><path d="M15.5 14.5v6M12.5 17.5h6" stroke="'+WB+'" stroke-width="2"/>'),
 idx:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M8 7.5h3M13 7.5h3M8 11.5h3M13 11.5h3M8 15.5h3M13 15.5h3" /><path d="M12 5v14" stroke="'+WB+'"/>'),
 mta:ICS('<path d="M12 4v15M7 20h10M5 7h14"/><path d="M5 7l-2.5 6a3 3 0 0 0 5 0zM19 7l-2.5 6a3 3 0 0 0 5 0z"/>'),
 toa:ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M8 7.5h8M8 11.5h8M8 15.5h5" /><path d="M16 14v5M13.5 16.5h5" stroke="'+WB+'"/>'),
 upd:ICS('<path d="M19.5 12a7.5 7.5 0 1 1-2.5-5.6" stroke="'+WB+'"/><path d="M19.5 4.5v4.5h-4.5" stroke="'+WB+'"/>'),
 spell:ICS(ICX(9,11,9,'ABC',{w:700})+'<path d="M5.5 16.5l4 4 10-10" stroke="'+WB+'" stroke-width="2.2"/>'),
 wc:ICS(ICX(11,10,8.5,'ABC',{w:700})+ICX(11,19,8.5,'123',{w:700,f:WB})),
 read:WM.read,
 ruler:ICS('<rect x="2.5" y="8" width="19" height="8" fill="#fff"/><path d="M6 8v3.5M9.5 8v5M13 8v3.5M16.5 8v5"/>'),
 zin:ICS('<circle cx="10" cy="10" r="6"/><path d="M14.5 14.5l6 6" stroke-width="2.2"/><path d="M7.3 10h5.4M10 7.3v5.4" stroke="'+WB+'"/>'),
 zout:ICS('<circle cx="10" cy="10" r="6"/><path d="M14.5 14.5l6 6" stroke-width="2.2"/><path d="M7.3 10h5.4" stroke="'+WB+'"/>'),
 z100:ICS('<circle cx="10" cy="10" r="6"/><path d="M14.5 14.5l6 6" stroke-width="2.2"/>'+ICX(10,12.2,6,'1:1',{w:700,f:WB})),
 zfit:ICS('<path d="M4 8V4h4M20 8V4h-4M4 16v4h4M20 16v4h-4"/><path d="M8 12h8M10 10l-2 2 2 2M14 10l2 2-2 2" stroke="'+WB+'"/>'),
 save:ICS('<path d="M4.5 4.5h12l3 3v12h-15z"/><rect x="7.5" y="4.5" width="7" height="5"/><rect x="7.5" y="13.5" width="9" height="6" stroke="'+WB+'"/>'),
 undo:ICS('<path d="M9 6L4.5 10.5 9 15"/><path d="M4.5 10.5H15a5 5 0 0 1 0 10h-3"/>'),
 redo:ICS('<path d="M15 6l4.5 4.5L15 15"/><path d="M19.5 10.5H9a5 5 0 0 0 0 10h3"/>'),
 selp:ICS('<rect x="3.5" y="4" width="11" height="8" fill="#fff"/><rect x="9.5" y="11" width="11" height="8" fill="#cfe0f5" stroke="'+WB+'"/>'),
 fwd:ICS('<rect x="3.5" y="9.5" width="11" height="11" fill="#fff"/><rect x="9.5" y="3.5" width="11" height="11" fill="#cfe0f5" stroke="'+WB+'"/>'),
 bwd:ICS('<rect x="9.5" y="3.5" width="11" height="11" fill="#fff"/><rect x="3.5" y="9.5" width="11" height="11" fill="#cfe0f5" stroke="'+WB+'"/>'),
 bsq:ICS('<rect x="4.5" y="4.5" width="15" height="15"/>'),
 chk:ICS('<path d="M5 12.5l4.5 4.5L19 7" stroke="'+WB+'" stroke-width="2"/>')
};
const VALDD=new Set(['font','size','cstyle']);
const WPAL=[['#000000','#434343','#666666','#999999','#b7b7b7','#cccccc','#d9d9d9','#efefef','#f3f3f3','#ffffff'],['#980000','#ff0000','#ff9900','#ffff00','#00ff00','#00ffff','#4a86e8','#0000ff','#9900ff','#ff00ff'],['#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3','#c9daf8','#cfe2f3','#d9d2e9','#ead1dc'],['#dd7e6b','#ea9999','#f9cb9c','#ffe599','#b6d7a8','#a2c4c9','#a4c2f4','#9fc5e8','#b4a7d6','#d5a6bd'],['#a61c00','#cc0000','#e69138','#f1c232','#6aa84f','#45818e','#3c78d8','#3d85c6','#674ea7','#a64d79']];
const B=(ic,tip,a,m,lab)=>`<button class="b ${m||''}" title="${tip}" data-a="${a}"><em>${ic}</em>${m==='big'||m==='rw'?`<span class="lbl">${lab||tip}</span>`:''}</button>`;
/* S(nama,opsi,lebar,tooltip,reset,label,mode): label tampil di samping ikon (kecil) atau di bawah ikon (mode 'big') */
const S=(n,o,w,tip,r,lab,m)=>{const op=o.map(x=>Array.isArray(x)?`<option value="${x[0]}">${x[1]}</option>`:`<option>${x}</option>`).join(''),v=VALDD.has(n);
  const cls=v?' vl':m==='big'?' big':lab?' sm':'';
  const inner=v?'<span class="v"></span><i class="ar">▾</i>':m==='big'?`<em>${DDI[n]||''}</em><span class="lbl">${lab||''}<i class="ar">▾</i></span>`:`<em>${DDI[n]||''}</em>${lab?`<span class="lbl">${lab}</span>`:''}<i class="ar">▾</i>`;
  return `<span class="ddw"><select data-s="${n}" title="${tip||''}"${r?' data-r':''} hidden>${op}</select><button class="b dd${cls}" data-dd="${n}" title="${tip||''}"${v?` style="width:${w}px"`:''}>${inner}</button></span>`;};
const C=(n,v,t,ic,lab)=>`<span class="clr" data-cp="${n}" title="${t}"><span style="border-bottom:3px solid ${v}">${ic}</span><input type="color" data-s="${n}" value="${v}">${lab?`<span class="lbl2">${lab}</span>`:''}<i class="ar">▾</i></span>`;
const CB=(n,v,t,ic,lab)=>`<span class="clr big" data-cp="${n}" title="${t}"><span style="border-bottom:3px solid ${v}">${ic}</span><input type="color" data-s="${n}" value="${v}"><span class="lbl">${lab}<i class="ar">▾</i></span></span>`;
const G=(n,h)=>`<div class="grp"><div class="gb">${h}</div><div class="gn">${n}</div></div>`;
const R2=(a,b)=>`<div class="rows"><div class="row">${a}</div><div class="row">${b}</div></div>`;
const AL=t=>{const w=[14,9,14,9],x=i=>t==='L'||t==='J'?0:t==='C'?(14-w[i])/2:14-w[i];if(t==='J')w.fill(14);return `<svg width="16" height="14">${w.map((v,i)=>`<rect x="${1+x(i)}" y="${1+i*3.5}" width="${v}" height="1.6" fill="#444"/>`).join('')}</svg>`;};

const N=(n,v,t,st,l,lab)=>`<label class="nl" title="${t}">${l}${lab?`<span class="nlt">${lab}</span>`:''}<input type="number" data-s="${n}" value="${v}" step="${st}"><small>${(t.match(/\((\w+)\)/)||[])[1]||''}</small></label>`;
const RIB={
home:
 G('Papan klip',B(WI.paste,'Tempel','paste','big')+`<div class="rows">${B(WI.cut,'Potong','cut','rw')}${B(WI.copy,'Salin','copy','rw')}${B(WI.painter,'Kuas Format','painter','rw')}</div>`)+
 G('Font',R2(
  S('font',FONTS,128,'Font')+S('size',SIZES,48,'Ukuran font')+B(WI.grow,'Perbesar font','grow')+B(WI.shrink,'Perkecil font','shrink')+
  S('case',[['','Aa'],['upper','HURUF BESAR'],['lower','huruf kecil'],['title','Kapital Tiap Kata']],50,'Ubah huruf besar/kecil',1)+B(WI.clear,'Hapus semua format','removeFormat'),
  B(WI.bold,'Tebal (Ctrl+B)','bold')+B(WI.italic,'Miring (Ctrl+I)','italic')+B(WI.underline,'Garis bawah (Ctrl+U)','underline')+B(WI.strike,'Coret','strikeThrough')+
  B(WI.sub,'Subskrip','subscript')+B(WI.sup,'Superskrip','superscript')+C('hilite','#ffff00','Warna sorot teks',WI.hilite)+C('color','#ff0000','Warna font',WI.fcolor)))+
 G('Paragraf',R2(
  B(WI.ul,'Poin','insertUnorderedList')+B(WI.ol,'Penomoran','insertOrderedList')+B(WI.outd,'Kurangi indentasi','outdent')+B(WI.ind,'Tambah indentasi','indent')+B(WI.marks,'Tampilkan/sembunyikan ¶','marks'),
  B(AL('L'),'Rata kiri (Ctrl+L)','justifyLeft')+B(AL('C'),'Tengah (Ctrl+E)','justifyCenter')+B(AL('R'),'Rata kanan (Ctrl+R)','justifyRight')+B(AL('J'),'Rata kiri-kanan (Ctrl+J)','justifyFull')+
  S('lh',[['','↕ Spasi'],['1','1,0'],['1.15','1,15'],['1.5','1,5'],['2','2,0'],['3','3,0']],64,'Spasi baris',1)+C('shade','#d9d9d9','Pengarsiran paragraf',WI.shade)))+
 G('Gaya',`<div class="gal">${STY.map((s,i)=>`<button class="sty" data-a="style:${i}"><span class="${s[2]||(s[1]==='h1'?'g1':s[1]==='h2'?'g2':'')}">AaBbCcDd</span><small>${s[0]}</small></button>`).join('')}</div>`)+
 G('Penyuntingan',`<div class="rows">${B(WI.find,'Cari','find','rw')}${B(WI.repl,'Ganti','replace','rw')}${B(WI.sall,'Pilih semua','selectAll','rw')}</div>`),
insert:
 G('Halaman',B(WI.pgbrk,'Pemisah halaman','pagebreak','big','Pemisah<br>Halaman'))+G('Tabel',B(WI.table,'Tabel','table','big'))+
 G('Ilustrasi',B(WI.image,'Gambar','image','big'))+G('Tautan',B(WI.link,'Tautan','link','big'))+
 G('Rumus',B(WI.eq,'Sisipkan rumus (Alt+=)','eq','big','Rumus')+`<div class="rows">${B(WI.eqedit,'Edit rumus yang dipilih','eqedit','rw','Edit Rumus')}${S('eqg',[['','Rumus Umum ▾'],...[['circ','Luas lingkaran'],['pyth','Teorema Pythagoras'],['abc','Rumus kuadrat (ABC)'],['binom','Teorema binomial'],['taylor','Deret Taylor'],['trig','Identitas trigonometri'],['fourier','Deret Fourier'],['deriv','Definisi turunan'],['integ','Teorema dasar kalkulus'],['normal','Distribusi normal'],['mean','Rata-rata']]],170,'Sisipkan rumus bawaan',1,'Rumus Umum')}</div>`)+
 G('Simbol',B(WI.hr,'Garis horizontal','insertHorizontalRule','big','Garis<br>Horizontal')+S('sym',[['','Ω Simbol'],'©','®','™','°','±','×','÷','€','£','→','←','✓'],80,'Sisipkan simbol',1,'Simbol','big'))+
 G('Teks',B(WI.date,'Tanggal & waktu','date','big','Tanggal &<br>Waktu'))+
 G('Header & Footer',`<div class="rows">${
  S('hdrs',[['','Header ▾'],['blank','Header kosong'],['title','Judul dokumen'],['three','Tiga kolom'],['rm','Hapus header']],118,'Header',1,'Header')+
  S('ftrs',[['','Footer ▾'],['blank','Footer kosong'],['title','Judul dokumen'],['three','Tiga kolom'],['rm','Hapus footer']],118,'Footer',1,'Footer')+
  S('pnum',[['','Nomor Halaman ▾'],['bc','Bawah – tengah'],['br','Bawah – kanan'],['bl','Bawah – kiri'],['tr','Atas – kanan'],['pc','Halaman X dari Y'],['rm','Hapus nomor halaman']],140,'Nomor halaman',1,'Nomor Halaman')}</div>`),
design:
 G('Pemformatan dokumen',
  S('thm',[['','Tema: Office'],['ion','Tema: Ion'],['facet','Tema: Facet'],['retro','Tema: Retrospect'],['serif','Tema: Serif'],['mono','Tema: Mono']],110,'Tema',0,'Tema','big')+
  S('hcol',[['#2f5496','Warna: Biru'],['#385723','Warna: Hijau'],['#c00000','Warna: Merah'],['#c55a11','Warna: Oranye'],['#7030a0','Warna: Ungu'],['#7f7f7f','Warna: Abu-abu'],['#000000','Warna: Hitam']],110,'Warna tema',0,'Warna','big')+
  S('fpair',[['Calibri Light|Calibri','Font: Office'],['Arial|Arial','Font: Arial'],['Cambria|Calibri','Font: Cambria-Calibri'],['Georgia|Verdana','Font: Georgia-Verdana'],['Times New Roman|Times New Roman','Font: Times New Roman'],['Trebuchet MS|Segoe UI','Font: Trebuchet-Segoe']],110,'Font tema',0,'Font','big')+
  S('pspace',[['','Spasi: Default'],['none','Tanpa spasi'],['compact','Kompak'],['tight','Ketat'],['open','Terbuka'],['relaxed','Santai'],['double','Dobel']],110,'Spasi paragraf',0,'Spasi<br>Paragraf','big'))+
 G('Efek',S('fx',[['','Efek: Tanpa'],['fx1','Efek: Bayangan'],['fx2','Efek: Garis bawah']],110,'Efek judul',0,'Efek','big')+`<div class="rows">${B(WI.chk,'Atur sebagai default','setdef','rw','Atur default')}</div>`)+
 G('Latar halaman',
  S('wmk',[['','Watermark: Tanpa'],['DRAF','DRAF'],['RAHASIA','RAHASIA'],['CONTOH','CONTOH'],['JANGAN DISALIN','JANGAN DISALIN'],['*','Kustom…']],118,'Watermark',0,'Watermark','big')+
  CB('pagecolor','#ffffff','Warna halaman',ICS('<rect x="4.5" y="3" width="15" height="18" rx="1.5" fill="#fff"/><path d="M4.5 12.5l7-7M4.5 17l11.5-11.5M9 21l10.5-10.5" stroke="#9db7dd"/>'),'Warna<br>Halaman')+
  S('bsty',[['','Batas: Tanpa'],['solid','Garis'],['double','Ganda'],['dashed','Putus-putus'],['groove','Alur']],100,'Batas halaman',0,'Batas<br>Halaman','big')+
  `<div class="rows">${S('bw',[['1','1 px'],['2','2 px'],['4','4 px'],['6','6 px']],54,'Tebal batas',0,'Tebal Batas')}${C('bcol','#2b579a','Warna batas',WI.bsq,'Warna Batas')}</div>`),
layout:
 G('Penataan halaman',
  S('margin',[['normal','Margin: Normal'],['narrow','Margin: Sempit'],['moderate','Margin: Sedang'],['wide','Margin: Lebar']],120,'Margin',0,'Margin','big')+
  S('orient',[['p','Tegak'],['l','Mendatar']],80,'Orientasi',0,'Orientasi','big')+
  S('pgsize',[['a4','A4'],['letter','Letter'],['legal','Legal']],70,'Ukuran kertas',0,'Ukuran','big')+
  S('cols',[['1','1 Kolom'],['2','2 Kolom'],['3','3 Kolom']],90,'Kolom',0,'Kolom','big')+
  S('dir',[['ltr','Arah: Kiri-Kanan'],['rtl','Arah: Kanan-Kiri']],120,'Arah teks',0,'Arah<br>Teks','big')+
  `<div class="rows">${
   S('brk',[['','Pemisah…'],['p','Halaman'],['c','Kolom']],90,'Pemisah',1,'Pemisah')+
   S('ln',[['','Nomor baris: Tidak'],['1','Nomor baris: Kontinu']],130,'Nomor baris',0,'Nomor Baris')+
   S('hyph',[['','Tanpa tanda hubung'],['1','Hubung otomatis']],120,'Pemenggalan kata',0,'Pemenggalan')}</div>`)+
 G('Paragraf',R2(N('indl',0,'Indentasi kiri (cm)','0.5',ICS('<path d="M4 4h16M10 9h10M10 14h10M4 19h16M4 9l4 2.5L4 14z"/>'),'Kiri')+N('indr',0,'Indentasi kanan (cm)','0.5',ICS('<path d="M4 4h16M4 9h10M4 14h10M4 19h16M20 9l-4 2.5 4 2.5z"/>'),'Kanan'),N('spb',0,'Spasi sebelum (pt)','6',ICS('<path d="M5 4h14M12 20V8M8.5 11.5L12 8l3.5 3.5"/>'),'Sebelum')+N('spa',8,'Spasi sesudah (pt)','6',ICS('<path d="M5 20h14M12 4v12M8.5 12.5L12 16l3.5-3.5"/>'),'Sesudah')))+
 G('Susun (pilih gambar/tabel)',
  S('pos',[['','Posisi'],['i','Sebaris'],['l','Kiri (membungkus)'],['c','Tengah'],['r','Kanan (membungkus)']],118,'Posisi objek',1,'Posisi','big')+
  S('wrap',[['','Bungkus Teks'],['in','Sebaris dengan teks'],['sq','Persegi'],['tb','Atas dan bawah']],118,'Bungkus teks',1,'Bungkus<br>Teks','big')+
  B(WI.fwd,'Majukan','fwd','big')+B(WI.bwd,'Mundurkan','bwd','big')+B(WI.selp,'Panel Seleksi','selpane','big','Panel<br>Seleksi')+
  S('algn',[['','Rata'],['l','Kiri'],['c','Tengah'],['r','Kanan']],70,'Ratakan objek / paragraf',1,'Ratakan','big')+
  S('rot',[['','Putar'],['cw','Kanan 90°'],['ccw','Kiri 90°'],['fh','Balik horizontal'],['fv','Balik vertikal']],90,'Putar objek',1,'Putar','big')),
ref:
 G('Daftar isi',B(WI.toc,'Daftar isi','toc','big')+`<div class="rows">${S('addtxt',[['','Tambah Teks'],['h1','Level 1'],['h2','Level 2'],['h3','Level 3'],['p','Tidak ditampilkan']],96,'Tambah Teks',1,'Tambah Teks')}${B(WI.upd,'Perbarui Tabel','tocup','rw')}</div>`)+
 G('Catatan kaki',B(WI.fn,'Sisipkan catatan kaki','footnote','big','Sisipkan<br>Catatan Kaki')+`<div class="rows">${B(WI.en,'Sisipkan catatan akhir','endnote','rw','Catatan Akhir')}${B(WI.nfn,'Catatan kaki berikutnya','nextfn','rw','Catatan Berikutnya')}${B(WI.snotes,'Tampilkan catatan','shownotes','rw','Tampilkan Catatan')}</div>`)+
 G('Kutipan & Daftar Pustaka',B(WI.cite,'Sisipkan kutipan','cite','big','Sisipkan<br>Kutipan')+`<div class="rows">${B(WI.src,'Kelola sumber','sources','rw','Kelola Sumber')}<div class="row" style="height:22px"><em class="ri" title="Gaya sitasi">${ICS('<path d="M7 7h4v4c0 3-1.5 5-4 6M15 7h4v4c0 3-1.5 5-4 6"/>')}</em><span class="nlt">Gaya:</span>${S('cstyle',['APA','MLA','Chicago','IEEE'],70,'Gaya sitasi')}</div>${B(WI.bib,'Daftar pustaka','bib','rw','Daftar Pustaka')}</div>`)+
 G('Keterangan',B(WI.cap,'Sisipkan keterangan','caption','big','Sisipkan<br>Keterangan')+`<div class="rows">${B(WI.tof,'Sisipkan tabel gambar','tof','rw','Tabel Gambar')}${B(WI.upd,'Perbarui tabel','capup','rw','Perbarui Tabel')}${B(WI.xref,'Referensi silang','xref','rw','Referensi Silang')}</div>`)+
 G('Indeks',B(WI.mxe,'Tandai entri','markxe','big','Tandai<br>Entri')+`<div class="rows">${B(WI.idx,'Sisipkan indeks','idx','rw','Sisipkan Indeks')}${B(WI.upd,'Perbarui indeks','idxup','rw','Perbarui Indeks')}</div>`)+
 G('Daftar kutipan hukum',B(WI.mta,'Tandai kutipan','markta','big','Tandai<br>Kutipan')+`<div class="rows">${B(WI.toa,'Sisipkan daftar kutipan','toa','rw','Sisipkan Daftar')}${B(WI.upd,'Perbarui tabel','toaup','rw','Perbarui Tabel')}</div>`),

review:
 G('Pemeriksaan',B(WI.spell,'Ejaan','spell','big')+B(WI.wc,'Hitung kata','wordcount','big','Hitung<br>Kata'))+G('Mode',B(WI.read,'Mode baca','read','big')),
view:
 G('Tampilan',B(WI.read,'Mode baca','read','big'))+G('Tampilkan',B(WI.ruler,'Penggaris','ruler','big')+B(WI.marks,'Tanda paragraf','marks','big','Tanda<br>Paragraf'))+
 G('Zoom',B(WI.z100,'100%','zoom:100','big')+B(WI.zout,'75%','zoom:75','big')+B(WI.zin,'150%','zoom:150','big')+B(WI.zfit,'Lebar halaman','zoom:fit','big','Lebar<br>Halaman'))
};

/* ===== Ribbon kontekstual: muncul saat tabel / gambar / garis dipilih ===== */
const CTX=[['tbl','FORMAT TABEL'],['pic','FORMAT GAMBAR'],['hr','FORMAT GARIS']];
Object.assign(DDI,{
 tsty:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><rect x="3.5" y="5" width="17" height="4.5" fill="#4472c4"/><path d="M3.5 9.5h17M3.5 14.2h17M9.2 5v14M14.8 5v14"/>'),
 tcal:ICS('<rect x="3.5" y="4.5" width="17" height="15" fill="#fff"/><path d="M7 9h10M9 12h6M7 15h10" stroke="'+WB+'"/>'),
 tbd:ICS('<rect x="4.5" y="4.5" width="15" height="15" stroke-dasharray="2 2"/><path d="M4.5 4.5h15" stroke="'+WB+'" stroke-width="2.4"/>'),
 tbs:ICS('<path d="M3.5 7h17"/><path d="M3.5 12h17" stroke-dasharray="3 2"/><path d="M3.5 17h17" stroke-dasharray="1 2.4"/>'),
 tfit:ICS('<rect x="4.5" y="6" width="15" height="12" fill="#fff"/><path d="M2.5 12h4M17.5 12h4M4.5 12l2-1.6M4.5 12l2 1.6M19.5 12l-2-1.6M19.5 12l-2 1.6" stroke="'+WB+'"/>')
});
const CXI={
 above:ICS('<rect x="4" y="10" width="16" height="9" fill="#fff"/><path d="M4 14.5h16M10 10v9M15 10v9"/><path d="M12 3.5v4.5M9.8 5.7L12 3.5l2.2 2.2" stroke="'+WB+'"/>'),
 below:ICS('<rect x="4" y="5" width="16" height="9" fill="#fff"/><path d="M4 9.5h16M10 5v9M15 5v9"/><path d="M12 20.5V16M9.8 18.3L12 20.5l2.2-2.2" stroke="'+WB+'"/>'),
 left:ICS('<rect x="10" y="4" width="10" height="16" fill="#fff"/><path d="M15 4v16M10 9.3h10M10 14.7h10"/><path d="M3.5 12H8M5.7 9.8L3.5 12l2.2 2.2" stroke="'+WB+'"/>'),
 right:ICS('<rect x="4" y="4" width="10" height="16" fill="#fff"/><path d="M9 4v16M4 9.3h10M4 14.7h10"/><path d="M20.5 12H16M18.3 9.8l2.2 2.2-2.2 2.2" stroke="'+WB+'"/>'),
 drow:ICS('<rect x="4" y="5" width="16" height="14" fill="#fff"/><path d="M4 9.7h16M4 14.3h16M10 5v14M15 5v14"/><path d="M4 9.7h16v4.6H4z" fill="#f4b6b6" stroke="#c0392b"/>'),
 dcol:ICS('<rect x="4" y="5" width="16" height="14" fill="#fff"/><path d="M4 9.7h16M4 14.3h16M9.3 5v14M14.7 5v14"/><path d="M9.3 5h5.4v14H9.3z" fill="#f4b6b6" stroke="#c0392b"/>'),
 dtbl:ICS('<rect x="4" y="5" width="16" height="14" fill="#f4b6b6" stroke="#c0392b"/><path d="M4 9.7h16M4 14.3h16M9.3 5v14M14.7 5v14" stroke="#c0392b"/><path d="M8 8l8 8M16 8l-8 8" stroke="#7a1f14" stroke-width="2"/>'),
 merge:ICS('<rect x="3.5" y="6" width="17" height="12" fill="#fff"/><path d="M12 6v3M12 15v3" stroke-dasharray="1.6 1.6"/><path d="M7 12h10M7 12l2.2-2M7 12l2.2 2M17 12l-2.2-2M17 12l-2.2 2" stroke="'+WB+'"/>'),
 split:ICS('<rect x="3.5" y="6" width="17" height="12" fill="#fff"/><path d="M12 6v12"/><path d="M4.5 12h5M4.5 12l2-1.8M4.5 12l2 1.8M19.5 12h-5M19.5 12l-2-1.8M19.5 12l-2 1.8" stroke="'+WB+'"/>'),
 drows:ICS('<rect x="4" y="3.5" width="16" height="17" fill="#fff"/><path d="M4 9.2h16M4 14.8h16" stroke="'+WB+'"/>'),
 dcols:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><path d="M9.2 5v14M14.8 5v14" stroke="'+WB+'"/>'),
 up:ICS('<path d="M12 19V6M6.5 11.5L12 6l5.5 5.5" stroke="'+WB+'" stroke-width="2"/><path d="M5 20h14"/>'),
 down:ICS('<path d="M12 5v13M6.5 12.5L12 18l5.5-5.5" stroke="'+WB+'" stroke-width="2"/><path d="M5 4h14"/>'),
 lock:ICS('<rect x="5.5" y="11" width="13" height="9" rx="1.2" fill="#fff"/><path d="M8.5 11V8a3.5 3.5 0 0 1 7 0v3"/>'),
 reset:ICS('<path d="M5 12a7 7 0 1 0 2.2-5.1M5 4.5v4h4" stroke="'+WB+'"/>'),
 fitw:ICS('<rect x="7" y="6" width="10" height="12" fill="#fff"/><path d="M2.5 12H6M21.5 12H18M2.5 12l2-1.6M2.5 12l2 1.6M21.5 12l-2-1.6M21.5 12l-2 1.6" stroke="'+WB+'"/>'),
 shadow:ICS('<rect x="4.5" y="4.5" width="11" height="11" fill="#fff"/><path d="M17.5 7.5v10h-10M19.5 9.5v10h-10" stroke="#999"/>'),
 alt:ICS('<rect x="3.5" y="4.5" width="17" height="15" rx="1" fill="#fff"/><path d="M7 15l2.2-6 2.2 6M7.8 13h2.8M14 9v6M14 9.5h2.4a1.6 1.6 0 0 1 0 3.2H14" stroke="'+WB+'"/>'),
 del:ICS('<path d="M5.5 7h13M10 7V4.5h4V7M7 7l.8 12.5h8.4L17 7M10.5 10.5v6M13.5 10.5v6"/>'),
 hdr:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><rect x="3.5" y="5" width="17" height="4.7" fill="#bcd0ec"/><path d="M3.5 9.7h17M3.5 14.3h17M9.2 5v14M14.8 5v14"/>'),
 band:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><path d="M3.5 9.7h17v4.6h-17z" fill="#d9e2f3"/><path d="M3.5 9.7h17M3.5 14.3h17"/>'),
 fcol:ICS('<rect x="3.5" y="5" width="17" height="14" fill="#fff"/><rect x="3.5" y="5" width="5.7" height="14" fill="#bcd0ec"/><path d="M3.5 9.7h17M3.5 14.3h17M9.2 5v14"/>'),
 w:ICS('<path d="M4 6v12M20 6v12M4 12h16M4 12l2.4-2M4 12l2.4 2M20 12l-2.4-2M20 12l-2.4 2" stroke="'+WB+'"/>'),
 h:ICS('<path d="M6 4h12M6 20h12M12 4v16M12 4l-2 2.4M12 4l2 2.4M12 20l-2-2.4M12 20l2-2.4" stroke="'+WB+'"/>'),
 ind:ICS('<path d="M4 4v16M8 12h12M8 12l2.4-2M8 12l2.4 2" stroke="'+WB+'"/>'),
 pen:ICS('<path d="M5 19l1-4L16.5 4.5l3 3L9 18z" fill="#fff"/><path d="M5 19l4-1" stroke="'+WB+'"/>'),
 rad:ICS('<path d="M4.5 19.5V10a5.5 5.5 0 0 1 5.5-5.5h9.5" stroke="'+WB+'"/>')
};
const CXH={
 tbl:
  G('Baris & Kolom',`<div class="rows">${B(CXI.above,'Sisipkan baris di atas','tins:above','rw','Sisip Atas')}${B(CXI.below,'Sisipkan baris di bawah','tins:below','rw','Sisip Bawah')}${B(CXI.drow,'Hapus baris','tdel:row','rw','Hapus Baris')}</div>`+
   `<div class="rows">${B(CXI.left,'Sisipkan kolom di kiri','tins:left','rw','Sisip Kiri')}${B(CXI.right,'Sisipkan kolom di kanan','tins:right','rw','Sisip Kanan')}${B(CXI.dcol,'Hapus kolom','tdel:col','rw','Hapus Kolom')}</div>`+
   B(CXI.dtbl,'Hapus tabel','tdel:tbl','big','Hapus<br>Tabel'))+
  G('Gabung',B(CXI.merge,'Gabungkan sel (pilih beberapa sel, atau gabung dengan sel di kanan)','tmerge','big','Gabung<br>Sel')+B(CXI.split,'Pisahkan sel','tsplit','big','Pisahkan<br>Sel'))+
  G('Ukuran Sel',`<div class="rows"><div class="row">${N('trh',0,'Tinggi baris (cm)','0.1',CXI.h,'Tinggi')}</div><div class="row">${N('tcw',0,'Lebar kolom (cm)','0.1',CXI.w,'Lebar')}</div></div>`+
   `<div class="rows"><div class="row">${N('tw',0,'Lebar tabel (cm)','0.5',CXI.w,'Tabel')}</div><div class="row">${S('tfit',[['','Sesuaikan Otomatis'],['win','Sesuaikan lebar halaman'],['cnt','Sesuaikan isi'],['fix','Kunci lebar kolom']],130,'Sesuaikan otomatis',1,'Sesuaikan')}</div></div>`+
   `<div class="rows">${B(CXI.drows,'Distribusikan tinggi baris','tdistr:r','rw','Samakan Baris')}${B(CXI.dcols,'Distribusikan lebar kolom','tdistr:c','rw','Samakan Kolom')}</div>`)+
  G('Perataan',S('tcal',[['','Rata Sel'],['tl','Atas Kiri'],['tc','Atas Tengah'],['tr','Atas Kanan'],['ml','Tengah Kiri'],['mc','Tengah'],['mr','Tengah Kanan'],['bl','Bawah Kiri'],['bc','Bawah Tengah'],['br','Bawah Kanan']],110,'Perataan teks di dalam sel',1,'Rata<br>Sel','big')+
   S('algn',[['','Rata'],['l','Kiri'],['c','Tengah'],['r','Kanan']],70,'Ratakan tabel di halaman (kiri / tengah / kanan)',1,'Rata<br>Tabel','big'))+
  G('Posisi Tabel',S('pos',[['','Posisi'],['i','Sebaris'],['l','Kiri (membungkus)'],['c','Tengah'],['r','Kanan (membungkus)']],118,'Posisi tabel',1,'Posisi','big')+
   S('wrap',[['','Bungkus Teks'],['in','Sebaris dengan teks'],['sq','Persegi'],['tb','Atas dan bawah']],118,'Bungkus teks',1,'Bungkus<br>Teks','big')+
   `<div class="rows">${B(CXI.up,'Pindahkan tabel ke atas satu blok','tmove:up','rw','Naikkan')}${B(CXI.down,'Pindahkan tabel ke bawah satu blok','tmove:down','rw','Turunkan')}<div class="row" style="height:20px">${N('tind',0,'Indentasi tabel (cm)','0.1',CXI.ind,'Geser')}</div></div>`)+
  G('Gaya Tabel',S('tsty',[['','Gaya Tabel'],['grid','Kisi Sederhana'],['plain','Tanpa Garis'],['out','Garis Luar'],['hz','Garis Horizontal'],['blue','Biru'],['gray','Abu-abu'],['green','Hijau'],['orange','Oranye']],120,'Gaya tabel',1,'Gaya<br>Tabel','big')+
   `<div class="rows">${B(CXI.hdr,'Baris judul (tebal)','thead','rw','Baris Judul')}${B(CXI.band,'Baris belang','tband','rw','Baris Belang')}${B(CXI.fcol,'Kolom pertama (tebal)','tfc','rw','Kolom Pertama')}</div>`)+
  G('Garis & Warna',S('tbd',[['','Tepi'],['all','Semua Tepi'],['out','Tepi Luar'],['in','Tepi Dalam'],['none','Tanpa Tepi'],['top','Tepi Atas'],['bottom','Tepi Bawah'],['left','Tepi Kiri'],['right','Tepi Kanan']],100,'Terapkan garis tepi',1,'Tepi','big')+
   `<div class="rows"><div class="row">${S('tbs',[['solid','Garis Polos'],['dashed','Putus-putus'],['dotted','Titik-titik'],['double','Ganda']],96,'Jenis garis',0,'Jenis')}</div><div class="row">${N('tbw',1,'Tebal garis (px)','1',CXI.pen,'Tebal')}</div></div>`+
   `<div class="rows">${C('tbc','#000000','Warna garis',CXI.pen,'Warna Garis')}${C('tsh','#ffff99','Warna latar sel',WI.shade,'Pengarsiran')}${B(WI.shade,'Hapus warna latar sel','tshn','rw','Tanpa Latar')}</div>`),
 pic:
  G('Ukuran',R2(N('iw',0,'Lebar gambar (px)','5',CXI.w,'Lebar'),N('ih',0,'Tinggi gambar (px)','5',CXI.h,'Tinggi'))+
   `<div class="rows">${B(CXI.lock,'Kunci rasio aspek','ilock','rw','Kunci Rasio')}${B(CXI.reset,'Kembalikan ke ukuran asli','ireset','rw','Ukuran Asli')}${B(CXI.fitw,'Lebarkan sesuai halaman','ifitw','rw','Lebar Halaman')}</div>`+
   `<div class="rows"><div class="row">${N('isc',100,'Skala (%)','10',CXI.w,'Skala')}</div></div>`)+
  G('Susun',S('pos',[['','Posisi'],['i','Sebaris'],['l','Kiri (membungkus)'],['c','Tengah'],['r','Kanan (membungkus)']],118,'Posisi objek',1,'Posisi','big')+
   S('wrap',[['','Bungkus Teks'],['in','Sebaris dengan teks'],['sq','Persegi'],['tb','Atas dan bawah']],118,'Bungkus teks',1,'Bungkus<br>Teks','big')+
   S('algn',[['','Rata'],['l','Kiri'],['c','Tengah'],['r','Kanan']],70,'Ratakan objek',1,'Ratakan','big')+
   S('rot',[['','Putar'],['cw','Kanan 90°'],['ccw','Kiri 90°'],['fh','Balik horizontal'],['fv','Balik vertikal']],90,'Putar objek',1,'Putar','big')+
   B(WI.fwd,'Majukan','fwd','big')+B(WI.bwd,'Mundurkan','bwd','big'))+
  G('Gaya Gambar',`<div class="rows"><div class="row">${N('ibw',0,'Tebal garis tepi (px)','1',CXI.pen,'Tepi')}</div><div class="row">${N('irad',0,'Sudut membulat (px)','2',CXI.rad,'Sudut')}</div></div>`+
   `<div class="rows">${C('ibc','#2b579a','Warna garis tepi',CXI.pen,'Warna Tepi')}${B(CXI.shadow,'Bayangan','ishadow','rw','Bayangan')}${B(CXI.alt,'Teks alternatif','ialt','rw','Teks Alt')}</div>`)+
  G('Hapus',B(CXI.del,'Hapus gambar','delobj','big','Hapus<br>Gambar')),
 hr:
  G('Garis',`<div class="rows"><div class="row">${N('hrw',1,'Tebal garis (px)','1',CXI.pen,'Tebal')}</div><div class="row">${N('hrp',100,'Lebar garis (%)','5',CXI.w,'Lebar')}</div></div>`+
   `<div class="rows">${C('hrc','#808080','Warna garis',CXI.pen,'Warna')}</div>`)+
  G('Susun',S('algn',[['','Rata'],['l','Kiri'],['c','Tengah'],['r','Kanan']],70,'Ratakan garis',1,'Ratakan','big')+B(CXI.del,'Hapus garis','delobj','big','Hapus<br>Garis'))
};

function mount(root,opts={}){
  if(!document.getElementById('wd-css')){const s=document.createElement('style');s.id='wd-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Dokumen1.docx',zoom=100,lastRange=null,painting=null,autoFit=true;
  const cfg={size:'a4',land:false,m:[...MG.normal],cols:1};
  root.innerHTML=`<div class="wd">
    <div class="qat"><div class="l"><span class="w">W</span><button data-a="save" title="Simpan (Ctrl+S)">${WI.save}</button><button data-a="undo" title="Urungkan (Ctrl+Z)">${WI.undo}</button><button data-a="redo" title="Ulangi (Ctrl+Y)">${WI.redo}</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}${CTX.map(t=>`<button class="tab ctx" data-tab="${t[0]}" hidden>${t[1]}</button>`).join('')}</div>
    ${TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${RIB[t[0]]}</div>`).join('')}${CTX.map(t=>`<div class="ribbon cx" data-pane="${t[0]}" hidden>${CXH[t[0]]}</div>`).join('')}
    <div class="rbar"><div class="ruler"></div></div>
    <div class="desk"><div class="wsheet"><div class="paper"><div class="hfl"></div>
      <div class="page" contenteditable="true" spellcheck="true"></div></div></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fb"><button data-a="fnext">Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="sp" hidden></div>
    <div class="status"><div class="s"><span class="pg">Halaman 1 dari 1</span><span class="wc">0 kata</span><span>Bahasa Indonesia</span></div>
      <div class="s"><button class="on" data-a="edit" data-m="edit" title="Mode edit (tata letak cetak)">${WM.edit}</button><button data-a="read" data-m="read" title="Mode baca">${WM.read}</button>
      <span><button data-a="zoomd:-10">−</button><input type="range" min="20" max="200" value="100" class="zr"><button data-a="zoomd:10">+</button> <span class="zv">100%</span></span></div></div>
    <div class="menu" hidden>
      <button data-a="new">${WM.mnew}<span>Dokumen kosong baru</span></button><button data-a="open">${WM.mopen}<span>Buka…</span></button>
      <button data-a="save">${WM.msave}<span>Simpan</span></button><button data-a="save" data-saveas="1">${WM.msave}<span>Simpan Sebagai…</span></button><button data-a="print">${WM.mprint}<span>Cetak / simpan PDF</span></button>
      <hr><button data-a="home">${WM.mback}<span>Kembali ke menu utama</span></button></div>
    <input type="file" class="fin" accept=".docx,.html,.htm,.txt" hidden><input type="file" class="img" accept="image/*" hidden>
  </div>`;

  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const page=$('.page'),paper=$('.paper'),ruler=$('.ruler'),sheet=$('.wsheet'),menu=$('.menu'),fin=$('.fin'),fd=$('.find');
  const toast=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;$('.wd').appendChild(t);setTimeout(()=>t.remove(),3000);};
  const setName=n=>{name=n;$('.fn').textContent=n.replace(/\.docx$/i,'')+' - Word';document.title=$('.fn').textContent;};
  window.addEventListener('office:renamed',function f(e){if(!root.isConnected){removeEventListener('office:renamed',f);return;}const d=e.detail||{};if(d.from===name)setName(d.to);});
  const dims=()=>{const[w,h]=SZ[cfg.size];return cfg.land?[h,w]:[w,h];};

  /* seleksi & perintah */
  document.addEventListener('selectionchange',function sc(){
    if(!root.isConnected){document.removeEventListener('selectionchange',sc);return;}
    const s=getSelection(); if(s.rangeCount&&page.contains(s.anchorNode)) lastRange=s.getRangeAt(0).cloneRange();
  });
  function focus(){
    const s=getSelection();if(s.rangeCount&&root.querySelector('.hfl').contains(s.anchorNode))return;
    const inside=s.rangeCount&&page.contains(s.anchorNode);
    page.focus({preventScroll:true});
    if(!inside&&lastRange){s.removeAllRanges();s.addRange(lastRange);}
  }
  const exec=(c,v=null)=>{focus();document.execCommand(c,false,v);refresh();stats();};
  const curSize=()=>{const s=getSelection();let n=s.anchorNode;if(!n||!page.contains(n))return 11;n=n.nodeType===3?n.parentElement:n;return Math.round(parseFloat(getComputedStyle(n).fontSize)*.75);};
  function setSize(pt){
    focus();document.execCommand('styleWithCSS',false,false);document.execCommand('fontSize',false,7);document.execCommand('styleWithCSS',false,true);
    page.querySelectorAll('font[size="7"]').forEach(el=>{const s=document.createElement('span');s.style.fontSize=pt+'pt';s.innerHTML=el.innerHTML;el.replaceWith(s);});
    refresh();
  }
  function blocks(){
    focus();const s=getSelection();if(!s.rangeCount)return[];
    let bl=[...page.querySelectorAll('p,h1,h2,h3,li,blockquote')].filter(el=>s.containsNode(el,true)&&!el.classList.contains('pgap'));
    if(!bl.length){let n=s.anchorNode;n=n.nodeType===3?n.parentElement:n;const e=n.closest('p,h1,h2,h3,li,blockquote');if(e&&page.contains(e))bl=[e];}
    return bl;
  }
  function refresh(){
    STATE.forEach(c=>{const b=root.querySelector(`[data-a="${c}"]`);if(b)b.classList.toggle('on',document.queryCommandState(c));});
    const f=(document.queryCommandValue('fontName')||'').replace(/["']/g,'');
    if(FONTS.includes(f)){$('[data-s=font]').value=f;ddSync('font');}
    const sz=curSize();if(SIZES.includes(sz)){$('[data-s=size]').value=sz;ddSync('size');}
    syncNums();ctxCaret();
  }
  function stats(){
    const t=page.innerText.trim(),h=dims()[1];
    $('.wc').textContent=(t?t.split(/\s+/).length:0)+' kata';
    $('.pg').textContent='Halaman 1 dari '+(cfg.cols>1?Math.max(1,Math.ceil(page.offsetHeight/h)):pgCount);
    if(!paging)sched();
  }

  /* tata letak halaman & penggaris */
  function drawRuler(){
    const[w]=dims(),[,mr,,ml]=cfg.m,u=37.795/2;
    let h=`<div class="rs" style="left:0;width:${ml}px"></div><div class="rs" style="right:0;width:${mr}px"></div><div class="mid" style="left:${ml}px;width:${w-ml-mr}px"></div>`;
    for(let i=-Math.ceil(ml/u);i<=Math.ceil((w-ml)/u);i++){const x=ml+i*u;if(x<0||x>w)continue;
      h+=i%2?`<i class="tk" style="left:${x}px"></i>`:(i?`<b style="left:${x}px">${Math.abs(i/2)}</b>`:`<b style="left:${x}px">0</b>`);}
    ruler.innerHTML=h+`<div class="mk" data-side="l" style="left:${ml-5}px"></div><div class="mk" data-side="r" style="left:${w-mr-5}px"></div>`;
    ruler.style.width=w+'px';
  }
  function apply(){
    const[w,h]=dims(),[mt,mr,mb,ml]=cfg.m;
    Object.assign(page.style,{width:w+'px',minHeight:h+'px',padding:`${mt}px ${mr}px ${mb}px ${ml}px`,columnCount:cfg.cols>1?cfg.cols:'auto',columnGap:'36px'});
    page.style.setProperty('--ph',h+'px');page.classList.toggle('col',cfg.cols>1);
    [['--mt',mt],['--mr',mr],['--mb',mb],['--ml',ml]].forEach(([k,v])=>paper.style.setProperty(k,v+'px'));
    drawRuler();if(autoFit)fit();stats();
  }
  ruler.addEventListener('mousedown',e=>{
    const mk=e.target.closest('.mk');if(!mk)return;e.preventDefault();
    const side=mk.dataset.side,left=ruler.getBoundingClientRect().left;
    const mv=ev=>{const x=(ev.clientX-left)/(zoom/100),[w]=dims();
      if(side==='l')cfg.m[3]=Math.round(Math.max(0,Math.min(x,w-cfg.m[1]-120)));else cfg.m[1]=Math.round(Math.max(0,Math.min(w-x,w-cfg.m[3]-120)));apply();};
    addEventListener('mousemove',mv);addEventListener('mouseup',()=>removeEventListener('mousemove',mv),{once:true});
  });
  const desk=$('.desk');
  /* sesuaikan lebar halaman dengan layar (ponsel/tablet); tidak pernah memperbesar lewat 100% */
  function fit(){
    const avail=desk.clientWidth-12;if(avail<=0)return;
    const v=Math.max(20,Math.min(100,Math.floor(avail/(dims()[0]+2)*100)));
    if(v!==zoom)setZoom(v);
  }
  function setMode(edit){page.contentEditable=edit?'true':'false';$$('.status [data-m]').forEach(b=>b.classList.toggle('on',(b.dataset.m==='edit')===edit));$$('.b[data-a="read"]').forEach(b=>b.classList.toggle('on',!edit));}
  function setZoom(v){zoom=v;sheet.style.zoom=ruler.style.zoom=v/100;$('.zr').value=v;$('.zv').textContent=v+'%';place();}

  /* berkas */
  const setContent=h=>{page.innerHTML=h||'<p><br></p>';pick(null);stats();};
  async function openFile(file){
    const ext=file.name.split('.').pop().toLowerCase();
    try{
      hdr=ftr='';
      if(ext==='docx'){
        const buf=await file.arrayBuffer();
        await lib('mammoth',()=>window.mammoth);
        let mbuf=buf,maths=[];
        try{  /* Mammoth membuang rumus Word (OMML): ganti dulu dengan token, lalu pulihkan sebagai MathML */
          await lib('jszip',()=>window.JSZip);await lib('omml',()=>window.WordMath);
          const pr=await WordMath.prepare(buf,window.JSZip);mbuf=pr.buf;maths=pr.maths;
        }catch(e){mbuf=buf;maths=[];toast('Rumus tidak terbaca: '+e.message);}
        let html=(await mammoth.convertToHtml({arrayBuffer:mbuf})).value;
        if(maths.length)html=window.WordMath.inject(html,maths);
        setContent(html);
        try{  /* header, footer, dan tata letak halaman asli dibaca terpisah dari isi */
          await lib('jszip',()=>window.JSZip);await lib('docxio',()=>window.WordDocx);
          const r=await WordDocx.readDocx(buf,window.JSZip);
          hdr=r.hdr;ftr=r.ftr;
          if(r.size){cfg.size=r.size;cfg.land=r.land;}
          if(r.m)cfg.m=r.m;
          cfg.cols=r.cols||1;
        }catch(e){toast('Header/footer asli tidak terbaca: '+e.message);}
        apply();
      }
      else if(ext==='txt'){setContent((await file.text()).split(/\n/).map(l=>'<p>'+(esc(l)||'<br>')+'</p>').join(''));}
      else setContent(new DOMParser().parseFromString(await file.text(),'text/html').body.innerHTML);
      setName(file.name.replace(/\.(html?|txt)$/i,'.docx'));
    }catch(e){toast('Tidak bisa membuka berkas: '+e.message);}
  }
  async function saveDocx(){
    try{
      pick(null);if(eqc)eqc.hide();await lib('docx',()=>window.docx);await lib('docxio',()=>window.WordDocx);
      if(page.querySelector('.eq[data-omml]'))await lib('jszip',()=>window.JSZip);   /* perlu untuk menyimpan rumus */
      const bs=$('[data-s=bsty]').value;
      const blob=await WordDocx.toBlob({JSZip:window.JSZip,page,name,cfg,D,hdr,ftr,rtl:page.dir==='rtl',lineNumbers:page.classList.contains('ln'),
        border:bs?{style:bs,w:+$('[data-s=bw]').value,color:$('[data-s=bcol]').value}:null,pageColor:paper.style.background});
      const a=document.createElement('a');a.href=URL.createObjectURL(blob);
      a.download=/\.docx$/i.test(name)?name:name+'.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+a.download);
    }catch(e){toast('Gagal menyimpan: '+e.message);}
  }

  /* cari & ganti */
  const f1=$('.f1'),f2=$('.f2');
  function findNext(){const t=f1.value;if(!t)return;
    for(let i=0;i<40;i++){if(!window.find(t,false,false,true)){toast('Tidak ditemukan');return;}const n=getSelection().anchorNode;if(n&&page.contains(n))return;}}
  function replaceAll(){const t=f1.value;if(!t)return;let c=0;const re=new RegExp(t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'),w=document.createTreeWalker(page,NodeFilter.SHOW_TEXT);
    for(let n;n=w.nextNode();)n.nodeValue=n.nodeValue.replace(re,()=>{c++;return f2.value;});toast(c+' penggantian dibuat');stats();}


/* ===== Paginasi: tiap halaman terpisah, ukuran sama ===== */
let PG_GAP=16;let pgCount=1,paging=false,pgT=0;
const cleanHTML=()=>{const c=page.cloneNode(true);c.querySelectorAll('.pgap').forEach(x=>x.remove());return c.innerHTML;};
function sched(){clearTimeout(pgT);pgT=setTimeout(paginate,120);}
function paginate(){
  if(!root.isConnected)return;
  paging=true;
  page.querySelectorAll('.pgap').forEach(x=>x.remove());
  const h=dims()[1],mt=cfg.m[0],mr=cfg.m[1],mb=cfg.m[2],ml=cfg.m[3],G=PG_GAP,T=h+G,ch=h-mt-mb;
  if(cfg.cols>1){page.style.minHeight=h+'px';pgCount=Math.max(1,Math.ceil(page.offsetHeight/h));paging=false;renderHF(pgCount,0);return;}
  const sc=()=>page.offsetHeight?page.getBoundingClientRect().height/page.offsetHeight:1;
  const Y=el=>(el.getBoundingClientRect().top-page.getBoundingClientRect().top)/sc(),H=el=>el.getBoundingClientRect().height/sc();
  const units=[];[...page.children].forEach(c=>{if(/^(UL|OL)$/.test(c.tagName))units.push(...c.children);else units.push(c);});
  let lastBot=mt;
  for(const el of units){
    const ht=H(el);if(!ht)continue;
    const top=Y(el),k=Math.floor(top/T);
    if(top+ht>k*T+h-mb&&ht<=ch&&top>k*T+mt+2){
      const sp=document.createElement(el.parentNode===page?'div':'li');
      sp.className='pgap';sp.contentEditable='false';sp.style.display='flow-root';el.before(sp);
      const y0=Y(sp),hh=(k+1)*T+mt-Y(el),a=Math.max(0,k*T+h-y0);
      if(hh>0)sp.style.cssText=`display:flow-root;height:${hh}px;`+(sp.tagName==='DIV'?`margin:0 -${mr}px 0 -${ml}px;`:'list-style:none;')+
        `background:linear-gradient(to bottom,transparent ${a}px,#aaa ${a}px,#aaa ${a+1}px,#ececec ${a+1}px,#ececec ${a+G-1}px,#aaa ${a+G-1}px,#aaa ${a+G}px,transparent ${a+G}px)`;
      else sp.remove();
    }
    lastBot=Math.max(lastBot,Y(el)+ht);
  }
  pgCount=Math.max(1,Math.ceil((lastBot+mb+G)/T));
  page.style.minHeight=(pgCount*T-G)+'px';
  paging=false;renderHF(pgCount,G);
  $('.pg').textContent='Halaman 1 dari '+pgCount;place();
}
page.addEventListener('load',sched,true);
if(document.fonts&&document.fonts.ready)document.fonts.ready.then(sched);

/* Cetak: susun ulang tanpa celah antarhalaman supaya tiap halaman pas dengan kertas, lengkap dengan header/footer */
function printDoc(){
  const[w,h]=dims();let st=document.getElementById('wd-print');
  if(!st){st=document.createElement('style');st.id='wd-print';document.head.appendChild(st);}
  st.textContent=`@page{size:${w}px ${h}px;margin:0}`;
  PG_GAP=0;paginate();
  const done=()=>{removeEventListener('afterprint',done);PG_GAP=16;paginate();};
  addEventListener('afterprint',done);
  window.print();
}

/* ===== Header, Footer, Nomor Halaman, tanda sudut tiap halaman ===== */
let hdr='',ftr='';
const pnHTML=(html,k,n)=>{const t=document.createElement('div');t.innerHTML=html;t.querySelectorAll('.pn').forEach(x=>x.textContent=k+1);t.querySelectorAll('.pt').forEach(x=>x.textContent=n);return t.innerHTML;};
function renderHF(n,G){
  const L=$('.hfl');if(!L||L.contains(document.activeElement))return;
  const[w,h]=dims(),[mt,mr,mb,ml]=cfg.m;let o='';
  const cm=(x,y,bw)=>`<i class="cm" style="left:${x}px;top:${y}px;border-width:${bw}"></i>`;
  for(let k=0;k<n;k++){
    const b=k*(h+G),bx=`left:${ml}px;width:${Math.max(40,w-ml-mr)}px;`;
    o+=cm(ml-30,b+mt-30,'0 1px 1px 0')+cm(w-mr,b+mt-30,'0 0 1px 1px')+cm(ml-30,b+h-mb,'1px 1px 0 0')+cm(w-mr,b+h-mb,'1px 0 0 1px');
    o+=`<div class="hf" data-t="h" contenteditable="true" data-ph="Header" style="${bx}top:${b+36}px;max-height:${Math.max(20,mt-44)}px">${pnHTML(hdr,k,n)}</div>`;
    o+=`<div class="hf" data-t="f" contenteditable="true" data-ph="Footer" style="${bx}top:${b+h-mb+14}px;max-height:${Math.max(20,mb-20)}px">${pnHTML(ftr,k,n)}</div>`;
  }
  L.innerHTML=o;
}
$('.hfl').addEventListener('input',e=>{
  const el=e.target.closest('.hf');if(!el)return;
  const c=el.cloneNode(true);c.querySelectorAll('.pn,.pt').forEach(x=>x.textContent='#');
  const html=c.innerHTML==='<br>'?'':c.innerHTML;if(el.dataset.t==='h')hdr=html;else ftr=html;
  [...$('.hfl').querySelectorAll(`.hf[data-t="${el.dataset.t}"]`)].forEach((o,k)=>{if(o!==el)o.innerHTML=pnHTML(html,k,pgCount);});
});
const HFT={blank:'<div><br></div>',title:()=>`<div style="text-align:center">${esc(name.replace(/\.docx$/i,''))}</div>`,three:'<div style="display:flex;justify-content:space-between"><span>[Kiri]</span><span>[Tengah]</span><span>[Kanan]</span></div>'};
const redoHF=()=>renderHF(pgCount,cfg.cols>1?0:PG_GAP);
function setHF(t,v){
  if(!v)return;const html=v==='rm'?'':typeof HFT[v]==='function'?HFT[v]():HFT[v];
  if(t==='h')hdr=html;else ftr=html;redoHF();
  const el=$(`.hf[data-t="${t}"]`);if(el&&html){el.scrollIntoView({block:'center'});el.focus();}
}
const PNM={bc:['f','center'],br:['f','right'],bl:['f','left'],tr:['h','right'],pc:['f','center',1]};
function pageNum(v){
  if(!v)return;
  const strip=html=>{const d=document.createElement('div');d.innerHTML=html;[...d.children].forEach(c=>{if(c.querySelector('.pn')||c.classList.contains('pn'))c.remove();});return d.innerHTML;};
  if(v==='rm'){hdr=strip(hdr);ftr=strip(ftr);}
  else{const[t,al,tot]=PNM[v],pn='<span class="pn" contenteditable="false">#</span>',
    row=`<div style="text-align:${al}">${tot?`Halaman ${pn} dari <span class="pt" contenteditable="false">#</span>`:pn}</div>`;
    if(t==='h')hdr=strip(hdr)+row;else ftr=strip(ftr)+row;}
  redoHF();
}

/* ===== Desain, Tata Letak, Referensi ===== */
const store={get(k,d){try{const v=localStorage.getItem('wd.'+k);return v===null?d:JSON.parse(v);}catch{return d;}},set(k,v){try{localStorage.setItem('wd.'+k,JSON.stringify(v));}catch{}}};
const D={h:'Calibri Light',b:'Calibri',hc:'#2f5496',lh:1.08,ps:8,fx:''};
const THM={'':['Calibri Light','Calibri','#2f5496',1.08],ion:['Trebuchet MS','Verdana','#c55a11',1.15],facet:['Trebuchet MS','Trebuchet MS','#385723',1.15],retro:['Georgia','Georgia','#7030a0',1.2],serif:['Times New Roman','Times New Roman','#000000',1.15],mono:['Courier New','Courier New','#7f7f7f',1.1]};
const PSP={'':[8,1.08],none:[0,1],compact:[4,1],tight:[6,1.15],open:[10,1.3],relaxed:[8,1.5],double:[8,2]};
const TAC=['Kasus','Undang-Undang','Peraturan','Lainnya'],CAPL=['Gambar','Tabel','Persamaan'];
let selObj=null,cstyle=store.get('cs','APA');const src=store.get('src',[]);
const setSel=(n,v)=>{const s=$(`[data-s=${n}]`);if(s){s.value=v;if(s.selectedIndex<0)s.selectedIndex=0;ddSync(n);}};
const roman=n=>{let r='';for(const[v,s]of[[10,'x'],[9,'ix'],[5,'v'],[4,'iv'],[1,'i']])while(n>=v){r+=s;n-=v;}return r;};
const tid=e=>{if(!e.id)e.id=(e.classList.contains('cap')?'cap':'h')+Math.random().toString(36).slice(2,8);return e.id;};
const pageOf=el=>{const r=el.getBoundingClientRect(),p=page.getBoundingClientRect(),k=page.offsetHeight?p.height/page.offsetHeight:1;return Math.max(1,Math.floor((r.top-p.top)/k/(dims()[1]+(cfg.cols>1?0:PG_GAP)))+1);};
const attr=t=>esc(t).replace(/"/g,'&quot;');

function dlg(title,body,ok,okText){
  const d=document.createElement('div');d.className='ov';
  d.innerHTML=`<div class="dg"><h3>${title}</h3><div class="db">${body}</div><div class="dbt"><button class="x">Batal</button><button class="ok">${okText||'OK'}</button></div></div>`;
  $('.wd').appendChild(d);d.querySelector('.x').onclick=()=>d.remove();
  d.querySelector('.ok').onclick=()=>{if(ok(d)!==false)d.remove();};return d;
}

/* Desain */
function applyDesign(d){
  Object.assign(D,d);const w=$('.wd');
  w.style.setProperty('--hf',`"${D.h}",Carlito,sans-serif`);w.style.setProperty('--bf',`"${D.b}",Carlito,"Segoe UI",sans-serif`);
  page.style.setProperty('--hc',D.hc);page.style.setProperty('--lh',D.lh);page.style.setProperty('--ps',D.ps+'pt');
  page.classList.toggle('fx1',D.fx==='fx1');page.classList.toggle('fx2',D.fx==='fx2');
  setSel('hcol',D.hc);setSel('fpair',D.h+'|'+D.b);setSel('fx',D.fx);
  const pk=Object.keys(PSP).find(k=>PSP[k][0]===+D.ps&&PSP[k][1]===+D.lh);setSel('pspace',pk===undefined?'':pk);
}
function border(){
  const s=$('[data-s=bsty]').value,w=+$('[data-s=bw]').value,c=$('[data-s=bcol]').value;
  paper.classList.toggle('bd',!!s);paper.style.outline=s?`${s==='double'&&w<3?3:w}px ${s} ${c}`:'';
}

/* Tata letak */
function curBlock(){
  const s=getSelection();let n=s.anchorNode;if(!n||!page.contains(n))return null;
  n=n.nodeType===3?n.parentElement:n;const e=n.closest('p,h1,h2,h3,li,blockquote');return e&&page.contains(e)?e:null;
}
function syncNums(){
  const n=curBlock();if(!n)return;const c=getComputedStyle(n);
  const m={indl:c.marginLeft,indr:c.marginRight,spb:c.marginTop,spa:c.marginBottom};
  for(const k in m){const i=$(`[data-s=${k}]`);if(i&&document.activeElement!==i)i.value=k[0]==='i'?+(parseFloat(m[k])/37.795).toFixed(1):Math.round(parseFloat(m[k])*.75);}
}
function pick(t,force){if(selObj)selObj.classList.remove('sel');selObj=t;if(t)t.classList.add('sel');if(!$('.sp').hidden)drawSP();ctxSync(force);}
const need=()=>selObj||(toast('Pilih gambar atau tabel terlebih dahulu'),null);
const objs=()=>[...page.querySelectorAll('img,table,hr')];
function drawSP(){
  const c={IMG:0,TABLE:0,HR:0},nm={IMG:'Gambar',TABLE:'Tabel',HR:'Garis'};
  $('.sp').innerHTML='<b>Panel Seleksi</b>'+(objs().map((o,i)=>`<div class="spi${o===selObj?' on':''}" data-i="${i}"><span>${nm[o.tagName]} ${++c[o.tagName]}</span><button data-v="${i}" title="Tampilkan/sembunyikan">${o.style.visibility==='hidden'?'🚫':'👁'}</button></div>`).join('')||'<div><i>Tidak ada objek</i></div>');
}
function selpane(){const p=$('.sp');p.hidden=!p.hidden;if(!p.hidden)drawSP();}
function z(d){const o=need();if(!o)return;o.style.position='relative';const v=Math.max(0,(+o.style.zIndex||0)+d);o.style.zIndex=v;toast('Urutan tumpukan: '+v);}
$('.sp').addEventListener('click',e=>{
  const os=objs(),b=e.target.closest('button[data-v]');
  if(b){const o=os[+b.dataset.v];o.style.visibility=o.style.visibility==='hidden'?'':'hidden';drawSP();return;}
  const r=e.target.closest('.spi');if(r){const o=os[+r.dataset.i];pick(o);o.scrollIntoView({block:'center'});}
});
page.addEventListener('click',e=>{
  const a=e.target.closest('a[href^="#"]');
  if(a){e.preventDefault();const t=page.querySelector('#'+CSS.escape(a.getAttribute('href').slice(1)));if(t)t.scrollIntoView({block:'center'});return;}
  const o=e.target.closest('img,table,hr');pick(o&&page.contains(o)?o:null);
});

/* Daftar isi, catatan, keterangan, indeks */
const tocHTML=()=>{const hs=[...page.querySelectorAll('h1,h2,h3')];
  return '<p><b>Daftar Isi</b></p>'+(hs.length?hs.map(h=>`<p style="margin:0 0 4pt ${(h.tagName[1]-1)*20}px"><a href="#${tid(h)}">${esc(h.textContent)}</a><span style="float:right">${pageOf(h)}</span></p>`).join(''):'<p><i>Belum ada judul</i></p>');};
function renum(){
  const c={};
  page.querySelectorAll('p.cap').forEach(p=>{const l=p.dataset.l;c[l]=(c[l]||0)+1;tid(p);p.querySelector('b').textContent=l+' '+c[l];});
  page.querySelectorAll('a.xr').forEach(a=>{const t=page.querySelector('#'+CSS.escape(a.dataset.t));if(t)a.textContent=t.classList.contains('cap')?t.querySelector('b').textContent:t.textContent;});
}
const tofHTML=l=>`<p><b>Daftar ${l}</b></p>`+([...page.querySelectorAll(`p.cap[data-l="${l}"]`)].map(p=>`<p><a href="#${tid(p)}">${esc(p.textContent)}</a><span style="float:right">${pageOf(p)}</span></p>`).join('')||`<p><i>Belum ada keterangan ${l}</i></p>`);
const idxHTML=()=>{
  const m={};page.querySelectorAll('.xe').forEach(x=>{(m[x.dataset.xe]=m[x.dataset.xe]||new Set()).add(pageOf(x));});
  const k=Object.keys(m).sort((a,b)=>a.localeCompare(b,'id'));let L='';
  return '<p><b>Indeks</b></p>'+(k.length?k.map(e=>{const l=e[0].toUpperCase(),hd=l!==L?(L=l,`<p style="margin:6pt 0 2pt"><b>${esc(l)}</b></p>`):'';return hd+`<p style="margin:0 0 2pt 12px">${esc(e)}, ${[...m[e]].sort((a,b)=>a-b).join(', ')}</p>`;}).join(''):'<p><i>Belum ada entri yang ditandai</i></p>');
};
const toaHTML=()=>{
  const m={};page.querySelectorAll('.ta').forEach(x=>{const g=m[x.dataset.cat]=m[x.dataset.cat]||{},e=x.dataset.ta;(g[e]=g[e]||new Set()).add(pageOf(x));});
  const cs=TAC.filter(c=>m[c]);
  return '<p><b>Daftar Kutipan Hukum</b></p>'+(cs.length?cs.map(c=>`<p style="margin:6pt 0 2pt"><b>${c}</b></p>`+Object.keys(m[c]).sort((a,b)=>a.localeCompare(b,'id')).map(e=>`<p style="margin:0 0 2pt 12px">${esc(e)}, ${[...m[c][e]].sort((a,b)=>a-b).join(', ')}</p>`).join('')).join(''):'<p><i>Belum ada kutipan yang ditandai</i></p>');
};
function updAll(){
  renum();
  page.querySelectorAll('.toc').forEach(t=>t.innerHTML=tocHTML());
  page.querySelectorAll('.tof').forEach(t=>t.innerHTML=tofHTML(t.dataset.l));
  page.querySelectorAll('.idx').forEach(t=>t.innerHTML=idxHTML());
  page.querySelectorAll('.toa').forEach(t=>t.innerHTML=toaHTML());
  refreshCites();toast('Tabel diperbarui');
}
function note(k){
  const t=prompt(k==='fn'?'Teks catatan kaki':'Teks catatan akhir');if(!t)return;
  const n=page.querySelectorAll('sup.'+k).length+1,lab=k==='fn'?n:roman(n);
  exec('insertHTML',`<sup class="${k}">${lab}</sup>`);
  let box=page.querySelector('.'+k+'s');
  if(!box){box=document.createElement('div');box.className=k+'s';box.innerHTML='<hr>'+(k==='en'?'<p class="fnote"><b>Catatan Akhir</b></p>':'');
    const en=page.querySelector('.ens');if(k==='fn'&&en)en.before(box);else page.appendChild(box);}
  box.insertAdjacentHTML('beforeend',`<p class="fnote"><sup>${lab}</sup> ${esc(t)}</p>`);
}
function nextfn(){
  const all=[...page.querySelectorAll('sup.fn,sup.en')];if(!all.length)return toast('Belum ada catatan');
  const s=getSelection(),nx=(lastRange&&all.find(x=>lastRange.comparePoint(x,0)>0))||all[0];
  nx.scrollIntoView({block:'center'});const r=document.createRange();r.selectNode(nx);s.removeAllRanges();s.addRange(r);
}
function shownotes(){const b=page.querySelector('.fns')||page.querySelector('.ens');b?b.scrollIntoView({block:'center'}):toast('Belum ada catatan');}
function caption(){
  pickMenu(CAPL.map(l=>({l,v:l})),o=>{const l=o.v;
    const h=`<p class="cap" data-l="${l}" style="font-size:9pt;color:#44546a;font-style:italic"><b>${l}</b></p>`;
    if(selObj)(selObj.closest('p')||selObj).insertAdjacentHTML('afterend',h);else exec('insertHTML',h);
    renum();});
}
function tof(){
  pickMenu(CAPL.map(l=>({l,v:l})),o=>{const l=o.v;renum();exec('insertHTML',`<div class="tof" data-l="${l}">${tofHTML(l)}</div><p><br></p>`);});
}
function xref(){
  const items=[...page.querySelectorAll('h1,h2,h3,p.cap')];if(!items.length)return toast('Belum ada judul atau keterangan');renum();
  pickMenu(items.map((e,i)=>({l:(e.classList.contains('cap')?e.textContent:'Judul: '+e.textContent).slice(0,60),v:i})),o=>{
    const e=items[o.v];if(!e)return;
    exec('insertHTML',`<a class="xr" href="#${tid(e)}" data-t="${tid(e)}">${esc(e.classList.contains('cap')?e.querySelector('b').textContent:e.textContent)}</a>&nbsp;`);
  });
}
function markxe(){
  focus();const t=getSelection().toString().trim();if(!t)return toast('Pilih teks yang akan ditandai');
  const e=prompt('Entri indeks',t);if(!e)return;
  exec('insertHTML',`<span class="xe" data-xe="${attr(e)}">${esc(t)}</span>`);
}
function idx(){exec('insertHTML',`<div class="idx">${idxHTML()}</div><p><br></p>`);}
function markta(){
  focus();const t=getSelection().toString().trim();if(!t)return toast('Pilih teks kutipan yang akan ditandai');
  pickMenu(TAC.map(k=>({l:k,v:k})),o=>exec('insertHTML',`<span class="ta" data-ta="${attr(t)}" data-cat="${o.v}">${esc(t)}</span>`));
}
function toa(){exec('insertHTML',`<div class="toa">${toaHTML()}</div><p><br></p>`);}

/* Kutipan & daftar pustaka */
const A=s=>s.author||'Anonim',last=a=>a.includes(',')?a.split(',')[0].trim():a.trim().split(/\s+/).pop();
const fmtCite=(s,i)=>cstyle==='IEEE'?`[${i+1}]`:cstyle==='MLA'?`(${last(A(s))})`:cstyle==='Chicago'?`(${last(A(s))} ${s.year||'t.t.'})`:`(${last(A(s))}, ${s.year||'t.t.'})`;
const bibEntry=(s,i)=>{const a=esc(A(s)),y=esc(s.year||'t.t.'),ti=esc(s.title||''),p=esc(s.pub||'');
  return cstyle==='MLA'?`${a}. <i>${ti}</i>. ${p?p+', ':''}${y}.`:cstyle==='Chicago'?`${a}. ${y}. <i>${ti}</i>. ${p?p+'.':''}`:cstyle==='IEEE'?`[${i+1}] ${a}, “${ti},” ${p?p+', ':''}${y}.`:`${a}. (${y}). <i>${ti}</i>. ${p?p+'.':''}`;};
const bibHTML=()=>'<p><b>Daftar Pustaka</b></p>'+(src.length?src.map((s,i)=>[s,i]).sort((x,y)=>cstyle==='IEEE'?x[1]-y[1]:A(x[0]).localeCompare(A(y[0]),'id')).map(([s,i])=>`<p style="padding-left:36px;text-indent:-36px">${bibEntry(s,i)}</p>`).join(''):'<p><i>Belum ada sumber</i></p>');
function refreshCites(){
  page.querySelectorAll('.cite').forEach(c=>{const i=src.findIndex(s=>s.id===c.dataset.src);if(i>=0)c.textContent=fmtCite(src[i],i);});
  const b=page.querySelector('.bib');if(b)b.innerHTML=bibHTML();
}
function sources(){
  const d=dlg('Kelola Sumber',`<div class="sl"></div><input class="a" placeholder="Pengarang (Belakang, Depan)"><input class="y" placeholder="Tahun"><input class="t" placeholder="Judul"><input class="p" placeholder="Penerbit / situs"><button class="add">Tambah sumber</button>`,()=>{refreshCites();},'Selesai');
  const q=s=>d.querySelector(s),draw=()=>{q('.sl').innerHTML=src.length?src.map((s,i)=>`<div class="sli"><span>${esc(A(s))} (${esc(s.year||'t.t.')}) — ${esc(s.title||'')}</span><button data-d="${i}">✕</button></div>`).join(''):'<i>Belum ada sumber</i>';};
  q('.add').onclick=()=>{const s={id:'s'+Date.now().toString(36),author:q('.a').value.trim(),year:q('.y').value.trim(),title:q('.t').value.trim(),pub:q('.p').value.trim()};
    if(!s.title&&!s.author)return;src.push(s);store.set('src',src);['.a','.y','.t','.p'].forEach(k=>q(k).value='');draw();};
  q('.sl').onclick=e=>{const i=e.target.dataset.d;if(i!==undefined){src.splice(+i,1);store.set('src',src);draw();}};
  draw();
}
function cite(){
  if(!src.length){toast('Tambahkan sumber terlebih dahulu');return sources();}
  pickMenu([...src.map((x,i)=>({l:(A(x)+' ('+(x.year||'t.t.')+') — '+(x.title||'')).slice(0,70),v:i})),{l:'＋ Kelola sumber…',v:-1}],o=>{
    if(o.v<0)return sources();const x=src[o.v];if(!x)return;
    exec('insertHTML',`<span class="cite" data-src="${x.id}">${fmtCite(x,o.v)}</span>&nbsp;`);
  });
}
function bib(){
  if(page.querySelector('.bib')){refreshCites();return toast('Daftar pustaka diperbarui');}
  exec('insertHTML',`<div class="bib">${bibHTML()}</div><p><br></p>`);
}


  /* aksi */
  const showFind=()=>{fd.hidden=false;f1.focus();f1.select();};
  /* tabel: grid pemilih + dialog (sama seperti PowerPoint) */
  function insertTable(r,c){
    r=Math.max(1,Math.min(100,r|0));c=Math.max(1,Math.min(20,c|0));
    const before=new Set(page.querySelectorAll('table'));
    exec('insertHTML','<table style="width:100%;table-layout:fixed">'+('<tr>'+'<td><br></td>'.repeat(c)+'</tr>').repeat(r)+'</table><p><br></p>');
    const t=[...page.querySelectorAll('table')].find(x=>!before.has(x)),td=t&&t.querySelector('td');
    if(td){const rg=document.createRange();rg.selectNodeContents(td);rg.collapse(true);const g=getSelection();g.removeAllRanges();g.addRange(rg);}
    if(t)pick(t,true);
  }
  function tablePicker(){
    const a=curAnchor;if(!a)return;let g='';
    for(let i=0;i<80;i++)g+=`<i data-r="${Math.floor(i/10)+1}" data-c="${i%10+1}"></i>`;
    openWP(a,`<div class="tbg"><div class="lb">Sisipkan Tabel</div><div class="gr">${g}</div></div><hr><div class="mi" data-tb="dlg"><span>${WI.table.replace('class="ic"','class="ic" style="width:16px;height:16px"')} Sisipkan Tabel…</span></div>`,e=>{
      const i=e.target.closest('i[data-r]');if(i){closeWP();insertTable(+i.dataset.r,+i.dataset.c);return;}
      if(e.target.closest('[data-tb]')){closeWP();tableDlg();}});
    if(!wp||wp._o!==a)return;
    const cells=[...wp.querySelectorAll('i[data-r]')],lb=wp.querySelector('.lb');
    const hov=e=>{const i=e.target.closest('i[data-r]');if(!i)return;const r=+i.dataset.r,c=+i.dataset.c;
      cells.forEach(x=>x.classList.toggle('on',+x.dataset.r<=r&&+x.dataset.c<=c));lb.textContent=c+' x '+r+' Tabel';};
    wp.addEventListener('mouseover',hov);wp.addEventListener('pointerdown',hov);
  }
  function tableDlg(){
    const d=dlg('Sisipkan Tabel','<label>Jumlah kolom (1–20): <input type="text" inputmode="numeric" class="tc" value="5"></label><label>Jumlah baris (1–100): <input type="text" inputmode="numeric" class="tr" value="4"></label>',d=>{
      const c=parseInt(d.querySelector('.tc').value,10),r=parseInt(d.querySelector('.tr').value,10);
      if(!(c>=1&&r>=1)){toast('Masukkan angka yang valid');return false;}insertTable(r,c);});
    d.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();d.querySelector('.ok').click();}else if(e.key==='Escape')d.remove();});
    const f=d.querySelector('.tc');f.focus();f.select();
  }
  /* ===== rumus: sisip / edit (word/eqedit.js dimuat saat pertama dipakai) ===== */
  let eqc=null;
  const eqLoad=async()=>{
    if(!eqc){
      await lib('omml',()=>window.WordMath);await lib('eqedit',()=>window.WordEq);
      eqc=window.WordEq.attach({host:$('.wd'),page,focus,refresh,stats,toast});
    }
    return eqc;
  };
  const eqFail=e=>toast('Editor rumus gagal dimuat: '+e.message);
  const acts={
    menu(){menu.hidden=!menu.hidden;},
    eq(){eqLoad().then(c=>c.openNew()).catch(eqFail);},
    eqedit(){eqLoad().then(c=>c.editCurrent()).catch(eqFail);},
    new(){setName('Dokumen1.docx');hdr=ftr='';setContent('');focus();},
    open(){fin.click();},save:saveDocx,print(){pick(null);if(eqc)eqc.hide();printDoc();},home(){opts.onHome&&opts.onHome();},
    async paste(){focus();try{document.execCommand('insertText',false,await navigator.clipboard.readText());}catch{toast('Gunakan Ctrl+V untuk menempel');}},
    painter(){painting={bold:document.queryCommandState('bold'),italic:document.queryCommandState('italic'),underline:document.queryCommandState('underline'),
      font:(document.queryCommandValue('fontName')||'Calibri').replace(/["']/g,''),size:curSize(),color:document.queryCommandValue('foreColor')};toast('Pilih teks untuk menerapkan format');},
    grow(){const n=SIZES.find(s=>s>curSize());if(n)setSize(n);},
    shrink(){const n=[...SIZES].reverse().find(s=>s<curSize());if(n)setSize(n);},
    marks(){page.classList.toggle('marks');},
    style(i){const[,k,c]=STY[i];
      if(k==='i'){const t=getSelection().toString();if(t)exec('insertHTML',`<span class="${c}">${esc(t)}</span>`);return;}
      exec('formatBlock',k==='p'?'P':k.toUpperCase());blocks().forEach(b=>b.className=c);},
    find:showFind,replace:showFind,fnext:findNext,fclose(){fd.hidden=true;},
    frep(){const s=getSelection();if(s.toString().toLowerCase()===f1.value.toLowerCase()&&page.contains(s.anchorNode)){focus();document.execCommand('insertText',false,f2.value);}findNext();},
    frall:replaceAll,
    pagebreak(){exec('insertHTML','<div class="pb" style="page-break-after:always;border-top:1px dashed #999;text-align:center;color:#999;font-size:9pt" contenteditable="false">Pemisah halaman</div><p><br></p>');},
    table:tablePicker,
    image(){$('.img').click();},
    link(){const u=prompt('Alamat tautan (URL)','https://');if(u)exec('createLink',u);},
    date(){exec('insertText',new Date().toLocaleString('id-ID',{day:'numeric',month:'long',year:'numeric',hour:'2-digit',minute:'2-digit'}));},
    toc(){if(!page.querySelector('h1,h2,h3'))return toast('Belum ada Judul 1/2/3 di dokumen');exec('insertHTML',`<div class="toc">${tocHTML()}</div><p><br></p>`);},
    tocup:updAll,capup:updAll,idxup:updAll,toaup:updAll,
    footnote(){note('fn');},endnote(){note('en');},nextfn,shownotes,
    cite,sources,bib,caption,tof,xref,markxe,idx,markta,toa,selpane,
    fwd(){z(1);},bwd(){z(-1);},
    setdef(){store.set('design',D);toast('Format disimpan sebagai default');},
    spell(){page.spellcheck=!page.spellcheck;toast('Pemeriksa ejaan '+(page.spellcheck?'aktif':'nonaktif'));},
    wordcount(){const t=page.innerText;toast(`${t.trim()?t.trim().split(/\s+/).length:0} kata · ${t.replace(/\n/g,'').length} karakter · ${page.querySelectorAll('p,h1,h2,h3,li').length} paragraf`);},
    read(){const e=page.isContentEditable;setMode(!e);toast(e?'Mode baca':'Mode edit');},edit(){if(!page.isContentEditable){setMode(true);toast('Mode edit');}},
    ruler(){root.querySelector('.rbar').hidden=!root.querySelector('.rbar').hidden;},
    zoom(v){if(v==='fit'){autoFit=true;fit();}else{autoFit=false;setZoom(+v);}},zoomd(d){autoFit=false;setZoom(Math.min(200,Math.max(20,zoom+ +d)));}
  };
  const caseOf=(t,m)=>m==='upper'?t.toUpperCase():m==='lower'?t.toLowerCase():t.toLowerCase().replace(/(^|\s)\S/g,c=>c.toUpperCase());
  const sels={
    font:v=>exec('fontName',v),size:v=>setSize(v),color:v=>exec('foreColor',v),hilite:v=>exec('hiliteColor',v),
    case(v){focus();const t=getSelection().toString();if(t&&v)exec('insertText',caseOf(t,v));},
    lh(v){if(v)blocks().forEach(b=>b.style.lineHeight=v);},
    shade(v){blocks().forEach(b=>b.style.background=v);},
    sym:v=>v&&exec('insertText',v),
    eqg:v=>{if(!v)return;eqLoad().then(c=>{const g=window.WordEq.GALLERY.find(x=>x[0]===v);if(g)c.insertTex(g[2],true);}).catch(eqFail);},
    thm(v){const t=THM[v];applyDesign({h:t[0],b:t[1],hc:t[2],lh:t[3],ps:8});},
    hcol:v=>applyDesign({hc:v}),fpair(v){const[h,b]=v.split('|');applyDesign({h,b});},
    pspace(v){const p=PSP[v];applyDesign({ps:p[0],lh:p[1]});},fx:v=>applyDesign({fx:v}),
    pagecolor:v=>paper.style.background=v,
    wmk(v){if(v==='*')v=prompt('Teks watermark','RAHASIA')||'';paper.classList.toggle('wm',!!v);
      if(v){paper.style.setProperty('--wmt',JSON.stringify(v));paper.style.setProperty('--wms',Math.round(Math.min(140,760/(v.length*.6)))+'px');}},
    bsty:border,bw:border,bcol:border,
    dir:v=>{page.dir=v;},ln:v=>page.classList.toggle('ln',v==='1'),hyph(v){page.lang='id';page.style.hyphens=v?'auto':'manual';},
    brk(v){if(v==='p')acts.pagebreak();else if(v==='c')exec('insertHTML','<div class="cb" style="break-after:column;border-top:1px dashed #999;font-size:9pt;color:#999;text-align:center" contenteditable="false">Pemisah kolom</div><p><br></p>');},
    indl:v=>blocks().forEach(b=>b.style.marginLeft=(+v||0)+'cm'),indr:v=>blocks().forEach(b=>b.style.marginRight=(+v||0)+'cm'),
    spb:v=>blocks().forEach(b=>b.style.marginTop=(+v||0)+'pt'),spa:v=>blocks().forEach(b=>b.style.marginBottom=(+v||0)+'pt'),
    pos(v){const o=need();if(!o||!v)return;const st=o.style;st.float='';st.margin='';if(o.tagName==='IMG')st.display='';
      if(v==='l'){st.float='left';st.margin='0 12px 8px 0';}else if(v==='r'){st.float='right';st.margin='0 0 8px 12px';}else if(v==='c'){if(o.tagName==='IMG')st.display='block';st.margin='0 auto';}},
    wrap(v){const o=need();if(!o||!v)return;const st=o.style;
      if(v==='in'){st.float='';st.margin='';st.clear='';if(o.tagName==='IMG')st.display='';}
      else if(v==='sq'){if(!st.float||st.float==='none'){st.float='left';st.margin='0 12px 8px 0';}}
      else{st.float='none';st.clear='both';st.margin='8px 0';if(o.tagName==='IMG')st.display='block';}},
    algn(v){if(!v)return;const o=selObj;if(!o)return exec(v==='c'?'justifyCenter':v==='r'?'justifyRight':'justifyLeft');
      const st=o.style;st.float='';if(o.tagName==='IMG')st.display='block';st.marginLeft=v==='l'?'0':'auto';st.marginRight=v==='r'?'0':'auto';},
    rot(v){if(!v)return;const o=need();if(!o)return;const d=o.dataset;let r=+d.r||0,fx=+d.fx||1,fy=+d.fy||1;
      if(v==='cw')r+=90;else if(v==='ccw')r-=90;else if(v==='fh')fx*=-1;else fy*=-1;
      Object.assign(d,{r,fx,fy});o.style.transform=`rotate(${r}deg) scale(${fx},${fy})`;},
    hdrs:v=>setHF('h',v),ftrs:v=>setHF('f',v),pnum:pageNum,
    addtxt(v){if(v){exec('formatBlock',v==='p'?'P':v.toUpperCase());blocks().forEach(b=>b.className='');}},
    cstyle(v){cstyle=v;store.set('cs',v);refreshCites();},
    margin(v){cfg.m=[...MG[v]];apply();},orient(v){cfg.land=v==='l';apply();},pgsize(v){cfg.size=v;apply();},cols(v){cfg.cols=+v;apply();}
  };

  /* event */
  /* ===== Tab kontekstual, gagang ubah ukuran/geser, dan perintah tabel / gambar / garis ===== */
  var sv0=null,PXCM=37.795,hnd=null,ctxKind=null,ctxWasRc=false,lastCell=null,imgLock=true,bcfg={w:1,c:'#000000',s:'solid'};
  const bodyW=()=>{const[w]=dims();return w-cfg.m[1]-cfg.m[3];};
  const clamp=(v,a,b)=>Math.max(a,Math.min(b,v));
  const lastFill={},sv=(n,v,f)=>{const i=$(`[data-s=${n}]`);if(i&&(f||document.activeElement!==i)){i.value=v;lastFill[n]=String(i.value);}};
  sv0=sv;const cmv=v=>+(v/PXCM).toFixed(2);
  const smallScr=()=>matchMedia('(max-width:700px),(max-height:500px)').matches;
  const cellOf=n=>{if(!n)return null;n=n.nodeType===3?n.parentElement:n;const c=n&&n.closest&&n.closest('td,th');return c&&page.contains(c)?c:null;};
  const done=(keep)=>{stats();place();ctxFill(true);if(!keep)sched();};

  function showTab(id,collapse){
    const wd=$('.wd');
    $$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',!collapse&&x.dataset.tab===id));
    $$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==id);
    wd.classList.toggle('rc',!!collapse);
    const t=$(`.tab[data-tab=${id}]`),bar=$('.tabs');if(t&&bar&&!collapse)bar.scrollTo({left:Math.max(0,t.offsetLeft-60),behavior:'smooth'});
  }
  function ctxSync(force){
    const kind=!selObj?null:selObj.tagName==='TABLE'?'tbl':selObj.tagName==='IMG'?'pic':'hr';
    $$('.tab.ctx').forEach(t=>t.hidden=t.dataset.tab!==kind);
    const wd=$('.wd'),act=$('.tab.on');
    if(kind&&kind!==ctxKind){
      if(kind!=='tbl'||force){ctxWasRc=wd.classList.contains('rc');showTab(kind,false);}
      else{const t=$(`.tab[data-tab=${kind}]`),bar=$('.tabs');if(t&&bar)bar.scrollTo({left:Math.max(0,t.offsetLeft-60),behavior:'smooth'});}
    }else if(!kind&&((act&&act.classList.contains('ctx'))||$$('[data-pane]:not([hidden])').some(p=>p.classList.contains('cx')))){
      showTab('home',ctxWasRc||(smallScr()&&false));ctxWasRc=false;
    }
    ctxKind=kind;
    if(hnd)hnd.className='hnd'+(kind==='hr'?' hr':'')+(selObj&&selObj.tagName==='IMG'?' nomv0':'');
    ctxFill();place();
  }
  function ctxFill(f){
    const o=selObj;if(!o||!page.contains(o))return;const sv=(n,v)=>sv0(n,v,f);
    if(o.tagName==='IMG'){
      sv('iw',Math.round(o.offsetWidth));sv('ih',Math.round(o.offsetHeight));
      sv('isc',o.naturalWidth?Math.round(o.offsetWidth/o.naturalWidth*100):100);
      sv('ibw',parseFloat(o.style.borderWidth)||0);sv('irad',parseFloat(o.style.borderRadius)||0);
      $$('[data-a=ilock]').forEach(b=>b.classList.toggle('on',imgLock));$$('[data-a=ishadow]').forEach(b=>b.classList.toggle('on',!!o.style.boxShadow));
    }else if(o.tagName==='TABLE'){
      sv('tw',cmv(o.offsetWidth));sv('tind',cmv(parseFloat(o.style.marginLeft)||0));
      const c=curCell();if(c){sv('trh',cmv(c.parentElement.offsetHeight));sv('tcw',cmv(c.offsetWidth));}
      $$('[data-a=thead]').forEach(b=>b.classList.toggle('on',!!o.dataset.hdr));
      $$('[data-a=tband]').forEach(b=>b.classList.toggle('on',!!o.dataset.band));
      $$('[data-a=tfc]').forEach(b=>b.classList.toggle('on',!!o.dataset.fc));
    }else if(o.tagName==='HR'){
      sv('hrw',parseFloat(o.style.borderTopWidth)||parseFloat(o.style.height)||1);
      sv('hrp',Math.round(o.offsetWidth/Math.max(1,o.parentElement.clientWidth)*100));
    }
  }

  /* --- sel / grid tabel --- */
  function curCell(){
    const t=selObj&&selObj.tagName==='TABLE'?selObj:null;if(!t)return null;
    const s=getSelection();let c=s.rangeCount&&page.contains(s.anchorNode)?cellOf(s.anchorNode):null;
    if(!c&&lastRange)c=cellOf(lastRange.startContainer);
    if(!c&&lastCell&&t.contains(lastCell))c=lastCell;
    if(c&&c.closest('table')!==t)c=c.closest('table')&&t.contains(c)?[...t.querySelectorAll('td,th')].find(x=>x.contains(c)&&x.closest('table')===t):null;
    return c&&t.contains(c)?c:(t.rows[0]&&t.rows[0].cells[0])||null;
  }
  function selCells(t){
    const out=new Set(),s=getSelection(),r=(s.rangeCount&&page.contains(s.anchorNode)&&!s.isCollapsed)?s.getRangeAt(0):(lastRange&&!lastRange.collapsed?lastRange:null);
    if(r){[...t.querySelectorAll('td,th')].forEach(c=>{if(c.closest('table')===t&&r.intersectsNode(c))out.add(c);});}
    if(out.size<2){out.clear();const c=curCell();if(c)out.add(c);}
    return[...out];
  }
  const allCells=t=>[...t.querySelectorAll('td,th')].filter(c=>c.closest('table')===t);
  const cellsOrAll=t=>{const c=selCells(t);return c.length>1?c:allCells(t);};
  function tgrid(t){
    const g=[],pos=new Map();
    [...t.rows].forEach((tr,r)=>{g[r]=g[r]||[];let c=0;
      [...tr.cells].forEach(cl=>{while(g[r][c])c++;pos.set(cl,{r,c});
        for(let a=0;a<cl.rowSpan;a++)for(let b=0;b<cl.colSpan;b++){if(r+a<t.rows.length)(g[r+a]=g[r+a]||[])[c+b]=cl;}
        c+=cl.colSpan;});});
    return{g,pos,W:Math.max(0,...g.map(x=>x.length)),H:g.length};
  }
  function newCell(ref){
    const c=document.createElement(ref&&ref.tagName==='TH'?'th':'td');
    if(ref){c.style.cssText=ref.style.cssText;c.style.background='';c.style.backgroundColor='';c.style.height='';c.style.width='';}
    c.innerHTML='<br>';return c;
  }
  const insertAt=(row,nc,col,pos)=>{const ref=[...row.cells].find(c=>pos.has(c)&&pos.get(c).c>col);row.insertBefore(nc,ref||null);};
  const delTable=()=>{const t=selObj;if(!t)return;const nx=t.nextElementSibling;t.remove();pick(null);if(nx&&nx.tagName==='P'&&!nx.textContent.trim()&&!nx.nextElementSibling&&!page.querySelector('table')&&false)nx.remove();stats();};

  function insRow(below){
    const t=selObj,cell=curCell();if(!t||!cell)return;
    const{g,pos,W}=tgrid(t),p=pos.get(cell),k=below?p.r+cell.rowSpan:p.r;
    const tr=document.createElement('tr'),seen=new Set(),refRow=g[Math.min(k,g.length-1)]||[];
    for(let c=0;c<W;c++){
      const up=k>0&&k<g.length?g[k][c]:null;
      if(up&&pos.get(up).r<k){if(!seen.has(up)){up.rowSpan++;seen.add(up);}continue;}
      const ref=refRow[c]||cell,nc=newCell(ref);
      if(ref.colSpan>1&&pos.get(ref).c===c){nc.colSpan=ref.colSpan;c+=ref.colSpan-1;}
      tr.appendChild(nc);
    }
    const before=t.rows[k];
    if(before)before.parentNode.insertBefore(tr,before);else t.rows[t.rows.length-1].parentNode.appendChild(tr);
    done();
  }
  function insCol(right){
    const t=selObj,cell=curCell();if(!t||!cell)return;
    const{g,pos,W}=tgrid(t),p=pos.get(cell),k=right?p.c+cell.colSpan:p.c,seen=new Set();
    allCells(t).forEach(c=>c.style.width='');t.style.tableLayout='';
    for(let r=0;r<t.rows.length;r++){
      const at=k<W?g[r][k]:null;
      if(at&&k>0&&pos.get(at).c<k){if(!seen.has(at)){at.colSpan++;seen.add(at);}continue;}
      const ref=(k>0?g[r][k-1]:null)||at||cell,nc=newCell(ref);nc.colSpan=1;nc.rowSpan=1;
      insertAt(t.rows[r],nc,k-1,pos);
    }
    done();
  }
  function delRow(){
    const t=selObj,cell=curCell();if(!t||!cell)return;
    if(t.rows.length<2)return delTable();
    const{g,pos,W}=tgrid(t),r=pos.get(cell).r,tr=t.rows[r],seen=new Set();
    for(let c=0;c<W;c++){
      const cl=g[r][c];if(!cl||seen.has(cl))continue;seen.add(cl);const p=pos.get(cl);
      if(p.r<r)cl.rowSpan--;
      else if(cl.rowSpan>1){const nx=t.rows[r+1];cl.rowSpan--;if(nx)insertAt(nx,cl,p.c,pos);}
    }
    tr.remove();done();
  }
  function delCol(){
    const t=selObj,cell=curCell();if(!t||!cell)return;
    const{g,pos,W}=tgrid(t);if(W<2)return delTable();
    const c=pos.get(cell).c,seen=new Set();
    allCells(t).forEach(x=>x.style.width='');
    for(let r=0;r<t.rows.length;r++){const cl=g[r][c];if(!cl||seen.has(cl))continue;seen.add(cl);if(cl.colSpan>1)cl.colSpan--;else cl.remove();}
    done();
  }
  function merge(){
    const t=selObj;if(!t)return;let cs=selCells(t);const{g,pos}=tgrid(t);
    if(cs.length<2){
      const c=cs[0];if(!c)return;const p=pos.get(c),rt=(g[p.r]||[])[p.c+c.colSpan];
      if(rt&&pos.get(rt).r===p.r&&rt.rowSpan===c.rowSpan)cs=[c,rt];
      else{const bt=(g[p.r+c.rowSpan]||[])[p.c];if(bt&&bt.colSpan===c.colSpan&&pos.get(bt).c===p.c)cs=[c,bt];else return toast('Tidak ada sel di sebelahnya untuk digabung');}
    }
    const set=new Set(cs);let r0,r1,c0,c1,again=true;
    const box=()=>{r0=c0=1e9;r1=c1=-1;set.forEach(c=>{const p=pos.get(c);r0=Math.min(r0,p.r);c0=Math.min(c0,p.c);r1=Math.max(r1,p.r+c.rowSpan-1);c1=Math.max(c1,p.c+c.colSpan-1);});};
    while(again){again=false;box();for(let r=r0;r<=r1;r++)for(let c=c0;c<=c1;c++){const x=(g[r]||[])[c];if(x&&!set.has(x)){set.add(x);again=true;}}}
    const keep=g[r0][c0],rest=[...set].filter(c=>c!==keep).sort((a,b)=>(pos.get(a).r-pos.get(b).r)||(pos.get(a).c-pos.get(b).c));
    rest.forEach(c=>{const has=c.textContent.trim()||c.querySelector('img,table');if(has){if(keep.textContent.trim()||keep.querySelector('img,table'))keep.insertAdjacentHTML('beforeend','<br>');keep.innerHTML=keep.innerHTML.replace(/^<br>$/,'')+c.innerHTML;}c.remove();});
    keep.rowSpan=r1-r0+1;keep.colSpan=c1-c0+1;keep.style.width='';
    const rg=document.createRange();rg.selectNodeContents(keep);rg.collapse(true);const s=getSelection();s.removeAllRanges();s.addRange(rg);lastCell=keep;
    done();
  }
  function split(){
    const t=selObj,cell=curCell();if(!t||!cell)return;
    const{pos}=tgrid(t),p=pos.get(cell);
    if(cell.colSpan>1||cell.rowSpan>1){
      const cs=cell.colSpan,rs=cell.rowSpan;cell.colSpan=1;cell.rowSpan=1;
      for(let a=0;a<rs;a++){const row=t.rows[p.r+a];if(!row)continue;for(let b=(a?0:1);b<cs;b++)insertAt(row,newCell(cell),p.c+b-(a?0:0),pos);}
      return done();
    }
    dlg('Pisahkan Sel','<label>Jumlah kolom: <input type="text" inputmode="numeric" class="sn" value="2"></label><label>Jumlah baris: <input type="text" inputmode="numeric" class="sm" value="1"></label>',d=>{
      const n=clamp(parseInt(d.querySelector('.sn').value)||1,1,12),m=clamp(parseInt(d.querySelector('.sm').value)||1,1,12);
      if(n<2&&m<2)return;
      let G=tgrid(t),pp=G.pos.get(cell);
      if(n>1){
        const seen=new Set([cell]);for(let r=0;r<G.H;r++){const x=G.g[r][pp.c];if(x&&!seen.has(x)){seen.add(x);x.colSpan+=n-1;}}
        for(let i=1;i<n;i++){const nc=newCell(cell);cell.after(nc);}
      }
      if(m>1){
        G=tgrid(t);pp=G.pos.get(cell);const mine=new Set([cell]);
        for(let i=1;i<n;i++){const x=G.g[pp.r][pp.c+i];if(x)mine.add(x);}
        const seen=new Set();G.g[pp.r].forEach(x=>{if(x&&!mine.has(x)&&!seen.has(x)){seen.add(x);x.rowSpan+=m-1;}});
        let after=t.rows[pp.r];
        for(let j=1;j<m;j++){const tr=document.createElement('tr');for(let i=0;i<n;i++)tr.appendChild(newCell(cell));after.after(tr);after=tr;}
      }
      done();
    },'Pisahkan');
  }
  function freezeCols(t){
    const{g,pos,W}=tgrid(t),w=Array(W).fill(0);
    allCells(t).forEach(c=>{if(c.colSpan===1){const i=pos.get(c).c;w[i]=Math.max(w[i],c.offsetWidth);}});
    const known=w.filter(x=>x>0),avg=known.length?known.reduce((a,b)=>a+b,0)/known.length:t.offsetWidth/Math.max(1,W);
    for(let i=0;i<W;i++)if(!w[i])w[i]=avg;
    allCells(t).forEach(c=>{c.style.width=c.colSpan===1?w[pos.get(c).c]+'px':'';});
    t.style.tableLayout='fixed';t.style.width=w.reduce((a,b)=>a+b,0)+'px';
    return{w,pos,W};
  }
  function setColW(t,c,px){
    const f=freezeCols(t),old=f.w[c];f.w[c]=clamp(px,24,Math.max(24,bodyW()-(f.w.reduce((a,b)=>a+b,0)-old)));
    allCells(t).forEach(x=>{if(x.colSpan===1&&f.pos.get(x).c===c)x.style.width=f.w[c]+'px';});
    t.style.width=f.w.reduce((a,b)=>a+b,0)+'px';
  }
  function setTblW(t,px){
    const f=freezeCols(t),sum=f.w.reduce((a,b)=>a+b,0);px=clamp(px,24*f.W,bodyW());const k=px/sum;
    allCells(t).forEach(x=>{if(x.colSpan===1)x.style.width=f.w[f.pos.get(x).c]*k+'px';});
    t.style.width=px+'px';
  }
  const SIDES=['Top','Right','Bottom','Left'],OPP={Top:'Bottom',Bottom:'Top',Left:'Right',Right:'Left'};
  function nbrs(g,pos,c,s){
    const p=pos.get(c),out=new Set();
    if(s==='Top')for(let b=0;b<c.colSpan;b++)out.add((g[p.r-1]||[])[p.c+b]);
    else if(s==='Bottom')for(let b=0;b<c.colSpan;b++)out.add((g[p.r+c.rowSpan]||[])[p.c+b]);
    else for(let a=0;a<c.rowSpan;a++)out.add((g[p.r+a]||[])[s==='Left'?p.c-1:p.c+c.colSpan]);
    out.delete(undefined);out.delete(c);return[...out];
  }
  const setSide=(c,s,on)=>{c.style['border'+s]=on?`${bcfg.w}px ${bcfg.s} ${bcfg.c}`:'none';};
  function applyBorders(t,mode,cells){
    const{g,pos}=tgrid(t),set=new Set(cells),line=mode!=='none';
    cells.forEach(c=>SIDES.forEach(s=>{
      const nb=nbrs(g,pos,c,s),outer=!nb.length||nb.some(n=>!set.has(n));
      const hit=mode==='all'||mode==='none'||(mode==='out'&&outer)||(mode==='in'&&!outer)||(outer&&mode===s.toLowerCase());
      if(!hit)return;setSide(c,s,line);nb.forEach(n=>setSide(n,OPP[s],line));
    }));
  }
  function restyleBorders(t){   /* ubah warna/tebal/jenis pada garis yang sudah ada */
    cellsOrAll(t).forEach(c=>{const cs=getComputedStyle(c);SIDES.forEach(s=>{if(cs['border'+s+'Style']!=='none'&&parseFloat(cs['border'+s+'Width'])>0)setSide(c,s,true);});});
  }
  const STYLES={
    grid:{b:['#000000'],hd:0,band:0},plain:{b:null,hd:0,band:0},out:{b:['#000000'],outer:1,hd:0,band:0},hz:{b:['#7f7f7f'],hz:1,hd:1,band:0},
    blue:{b:['#8eaadb'],hd:'#4472c4',band:'#d9e2f3'},gray:{b:['#a6a6a6'],hd:'#d9d9d9',hdc:'#000000',band:'#f2f2f2'},
    green:{b:['#a9d18e'],hd:'#70ad47',band:'#e2efd9'},orange:{b:['#f4b183'],hd:'#ed7d31',band:'#fbe5d6'}
  };
  function tableStyle(t,k){
    const d=STYLES[k];if(!d)return;
    allCells(t).forEach(c=>{c.style.border='';c.style.background='';c.style.color='';c.style.fontWeight='';});
    t.dataset.hdr=t.dataset.band=t.dataset.fc='';
    if(d.b){bcfg.c=d.b[0];bcfg.w=1;bcfg.s='solid';sv('tbw',1);const ci=$('[data-s=tbc]');if(ci){ci.value=bcfg.c;ci.previousElementSibling.style.borderBottomColor=bcfg.c;}
      if(d.outer)applyBorders(t,'out',allCells(t));else if(d.hz){applyBorders(t,'top',allCells(t));applyBorders(t,'bottom',allCells(t));}else applyBorders(t,'all',allCells(t));
      if(d.hz){allCells(t).forEach(c=>{c.style.borderLeft='none';c.style.borderRight='none';});}}
    else applyBorders(t,'none',allCells(t));
    if(d.hd){t.dataset.hdr='1';[...t.rows[0].cells].forEach(c=>{c.style.fontWeight='700';if(typeof d.hd==='string'){c.style.background=d.hd;c.style.color=d.hdc||'#ffffff';}});}
    if(d.band){t.dataset.band='1';[...t.rows].forEach((r,i)=>{if(i>0&&i%2===1)[...r.cells].forEach(c=>c.style.background=d.band);});}
    done(1);
  }
  function setCellAlign(t,v){
    const[vv,hh]=[{t:'top',m:'middle',b:'bottom'}[v[0]],{l:'left',c:'center',r:'right'}[v[1]]];
    selCells(t).forEach(c=>{c.style.verticalAlign=vv;c.style.textAlign=hh;c.querySelectorAll('p,h1,h2,h3,li,div').forEach(b=>b.style.textAlign=hh);});
  }
  function moveBlock(o,dir){
    let b=o;while(b.parentNode&&b.parentNode!==page)b=b.parentNode;
    let n=dir==='up'?b.previousElementSibling:b.nextElementSibling;
    while(n&&n.classList.contains('pgap'))n=dir==='up'?n.previousElementSibling:n.nextElementSibling;
    if(!n)return toast(dir==='up'?'Sudah di paling atas':'Sudah di paling bawah');
    if(dir==='up')n.before(b);else n.after(b);b.scrollIntoView({block:'nearest'});done();
  }

  /* --- gagang ubah ukuran & geser --- */
  hnd=document.createElement('div');hnd.className='hnd';hnd.hidden=true;
  hnd.innerHTML='<i data-h="mv" title="Seret untuk memindahkan">✥</i><i data-h="e" title="Ubah lebar"></i><i data-h="s" title="Ubah tinggi"></i><i data-h="se" title="Ubah ukuran"></i>';
  paper.appendChild(hnd);
  const dropl=document.createElement('div');dropl.className='dropl';dropl.hidden=true;paper.appendChild(dropl);
  function place(){
    if(!hnd)return;
    if(!selObj||!page.contains(selObj)||page.getAttribute('contenteditable')==='false'){hnd.hidden=true;return;}
    const pr=paper.getBoundingClientRect(),k=(pr.width/paper.offsetWidth)||1,r=selObj.getBoundingClientRect();
    if(!r.width&&!r.height){hnd.hidden=true;return;}
    Object.assign(hnd.style,{left:(r.left-pr.left)/k-1+'px',top:(r.top-pr.top)/k-1+'px',width:r.width/k+'px',height:r.height/k+'px'});
    hnd.hidden=false;
  }
  hnd.addEventListener('pointerdown',e=>{
    const g=e.target.closest('[data-h]'),o=selObj;if(!g||!o)return;e.preventDefault();e.stopPropagation();
    const kind=g.dataset.h,pr=paper.getBoundingClientRect(),k=(pr.width/paper.offsetWidth)||1;
    const st={x:e.clientX,y:e.clientY,w:o.offsetWidth,h:o.offsetHeight};
    let tf=null,rows=null;
    if(o.tagName==='TABLE'&&kind!=='mv'){tf=freezeCols(o);rows=[...o.rows].map(r=>r.offsetHeight);}
    const ratio=st.h/Math.max(1,st.w);
    const mv=ev=>{
      const dx=(ev.clientX-st.x)/k,dy=(ev.clientY-st.y)/k;
      if(kind==='mv'){dragTo(o,ev);return;}
      if(o.tagName==='IMG'){
        let w=st.w,h=st.h;
        if(kind==='e'||kind==='se')w=st.w+dx;if(kind==='s'||kind==='se')h=st.h+dy;
        if(imgLock){if(kind==='s')w=h/ratio;else h=w*ratio;}
        w=clamp(w,16,bodyW());if(imgLock)h=w*ratio;h=Math.max(16,h);
        o.style.width=Math.round(w)+'px';o.style.height=Math.round(h)+'px';
      }else if(o.tagName==='TABLE'){
        if(kind!=='s'){const sum=tf.w.reduce((a,b)=>a+b,0),px=clamp(sum+dx,24*tf.W,bodyW()),f=px/sum;
          allCells(o).forEach(x=>{if(x.colSpan===1)x.style.width=tf.w[tf.pos.get(x).c]*f+'px';});o.style.width=px+'px';}
        if(kind!=='e'){const f=Math.max(.2,(st.h+dy)/st.h);[...o.rows].forEach((r,i)=>r.style.height=Math.max(14,rows[i]*f)+'px');}
      }else if(o.tagName==='HR'){o.style.width=clamp(st.w+dx,20,bodyW())+'px';}
      place();ctxFill();
    };
    const up=ev=>{
      document.removeEventListener('pointermove',mv);document.removeEventListener('pointerup',up);document.removeEventListener('pointercancel',up);
      if(kind==='mv')dropTo(o,ev);dropl.hidden=true;done();
    };
    document.addEventListener('pointermove',mv);document.addEventListener('pointerup',up);document.addEventListener('pointercancel',up);
  });
  function dropSpot(o,ev){
    hnd.style.visibility='hidden';const el=document.elementFromPoint(ev.clientX,ev.clientY);hnd.style.visibility='';
    if(!el||!page.contains(el))return null;
    if(o.tagName==='IMG'){
      const r=document.caretRangeFromPoint?document.caretRangeFromPoint(ev.clientX,ev.clientY):null;
      if(!r||!page.contains(r.startContainer)||o.contains(r.startContainer))return null;return{range:r};
    }
    let b=el.closest('.page > *');if(!b||b===o||o.contains(b)||b.classList.contains('pgap'))return null;
    const rc=b.getBoundingClientRect();return{block:b,after:ev.clientY>rc.top+rc.height/2,rc};
  }
  function dragTo(o,ev){
    const s=dropSpot(o,ev);if(!s){dropl.hidden=true;return;}
    const pr=paper.getBoundingClientRect(),k=(pr.width/paper.offsetWidth)||1;dropl.hidden=false;
    if(s.range){const rc=s.range.getClientRects()[0]||s.range.getBoundingClientRect();Object.assign(dropl.style,{left:(rc.left-pr.left)/k+'px',top:(rc.top-pr.top)/k+'px',width:'2px',height:Math.max(12,rc.height/k)+'px'});}
    else{const y=(s.after?s.rc.bottom:s.rc.top)-pr.top;Object.assign(dropl.style,{left:(s.rc.left-pr.left)/k+'px',top:y/k-1+'px',width:s.rc.width/k+'px',height:'3px'});}
  }
  function dropTo(o,ev){
    const s=dropSpot(o,ev);if(!s)return;
    if(s.range)s.range.insertNode(o);else if(s.after)s.block.after(o);else s.block.before(o);
    pick(o);
  }

  /* --- perintah (dipasang ke acts / sels) --- */
  const need2=()=>selObj;
  Object.assign(acts,{
    tins(v){if(v==='above')insRow(0);else if(v==='below')insRow(1);else if(v==='left')insCol(0);else insCol(1);},
    tdel(v){if(v==='row')delRow();else if(v==='col')delCol();else delTable();},
    tmerge:merge,tsplit:split,
    tdistr(v){const t=selObj;if(!t)return;
      if(v==='r'){const hs=[...t.rows].map(r=>r.offsetHeight),a=hs.reduce((x,y)=>x+y,0)/hs.length;[...t.rows].forEach(r=>r.style.height=a+'px');}
      else{const f=freezeCols(t),a=f.w.reduce((x,y)=>x+y,0)/f.W;allCells(t).forEach(c=>{c.style.width=c.colSpan===1?a+'px':'';});}
      done();},
    tmove(v){if(selObj)moveBlock(selObj,v);},
    thead(){const t=selObj;if(!t)return;t.dataset.hdr=t.dataset.hdr?'':'1';[...t.rows[0].cells].forEach(c=>c.style.fontWeight=t.dataset.hdr?'700':'');done(1);},
    tband(){const t=selObj;if(!t)return;t.dataset.band=t.dataset.band?'':'1';
      [...t.rows].forEach((r,i)=>{if(i>0&&i%2===1)[...r.cells].forEach(c=>{if(t.dataset.band)c.style.background='#f2f2f2';else if(/^(#f2f2f2|rgb\(242, 242, 242\))$/i.test(c.style.backgroundColor||''))c.style.background='';});});done(1);},
    tfc(){const t=selObj;if(!t)return;t.dataset.fc=t.dataset.fc?'':'1';[...t.rows].forEach(r=>{if(r.cells[0])r.cells[0].style.fontWeight=t.dataset.fc?'700':'';});done(1);},
    tshn(){const t=selObj;if(t)selCells(t).forEach(c=>c.style.background='');},
    ilock(){imgLock=!imgLock;ctxFill();toast(imgLock?'Rasio aspek dikunci':'Rasio aspek dibuka');},
    ireset(){const o=selObj;if(!o)return;o.style.width=o.style.height='';o.removeAttribute('width');o.removeAttribute('height');done(1);setTimeout(()=>{place();ctxFill(true);},60);},
    ifitw(){const o=selObj;if(!o)return;const r=o.offsetHeight/Math.max(1,o.offsetWidth),w=bodyW();o.style.width=w+'px';o.style.height=Math.round(w*r)+'px';done(1);},
    ishadow(){const o=selObj;if(!o)return;o.style.boxShadow=o.style.boxShadow?'':'3px 3px 8px rgba(0,0,0,.4)';ctxFill();},
    ialt(){const o=selObj;if(!o)return;dlg('Teks Alternatif','<label>Deskripsi gambar: <input type="text" class="alt" value="'+(o.getAttribute('alt')||'').replace(/"/g,'&quot;')+'"></label>',d=>{o.setAttribute('alt',d.querySelector('.alt').value);toast('Teks alternatif disimpan');});},
    delobj(){const o=selObj;if(!o)return;pick(null);o.remove();stats();}
  });
  Object.assign(sels,{
    tw(v){const t=selObj;if(t&&+v>0){setTblW(t,+v*PXCM);done(1);}},
    trh(v){const t=selObj,c=curCell();if(!t||!c)return;c.parentElement.style.height=+v>0?(+v*PXCM)+'px':'';done(1);},
    tcw(v){const t=selObj,c=curCell();if(!t||!c||!(+v>0))return;const{pos}=tgrid(t);setColW(t,pos.get(c).c,+v*PXCM);done(1);},
    tind(v){const t=selObj;if(!t)return;t.style.float='';t.style.marginLeft=(+v||0)*PXCM+'px';t.style.marginRight='';done(1);},
    tfit(v){const t=selObj;if(!v||!t)return;
      if(v==='win'){allCells(t).forEach(c=>c.style.width='');t.style.tableLayout='fixed';t.style.width='100%';t.style.marginLeft=t.style.marginRight='';}
      else if(v==='cnt'){allCells(t).forEach(c=>c.style.width='');t.style.tableLayout='';t.style.width='auto';}
      else freezeCols(t);
      done(1);},
    tsty(v){if(v&&selObj)tableStyle(selObj,v);},
    tcal(v){if(v&&selObj)setCellAlign(selObj,v);},
    tbd(v){if(v&&selObj){applyBorders(selObj,v,cellsOrAll(selObj));done(1);}},
    tbs(v){bcfg.s=v;if(selObj)restyleBorders(selObj);},
    tbw(v){bcfg.w=clamp(+v||1,.5,8);if(selObj)restyleBorders(selObj);},
    tbc(v){bcfg.c=v;if(selObj)restyleBorders(selObj);},
    tsh(v){const t=selObj;if(t)selCells(t).forEach(c=>c.style.background=v);},
    iw(v){const o=selObj;if(!o||!(+v>0))return;const r=o.offsetHeight/Math.max(1,o.offsetWidth),w=clamp(+v,10,bodyW());o.style.width=w+'px';o.style.height=Math.round((imgLock?w*r:o.offsetHeight))+'px';done(1);},
    ih(v){const o=selObj;if(!o||!(+v>0))return;const r=o.offsetWidth/Math.max(1,o.offsetHeight),h=Math.max(10,+v);let w=imgLock?h*r:o.offsetWidth;
      if(w>bodyW()){w=bodyW();} o.style.height=(imgLock&&w!==h*r?Math.round(w/r):Math.round(h))+'px';o.style.width=Math.round(w)+'px';done(1);},
    isc(v){const o=selObj;if(!o||!(+v>0)||!o.naturalWidth)return;const w=clamp(o.naturalWidth*(+v)/100,10,bodyW());o.style.width=Math.round(w)+'px';o.style.height=Math.round(w*o.naturalHeight/o.naturalWidth)+'px';done(1);},
    ibw(v){const o=selObj;if(!o)return;o.style.border=+v>0?`${+v}px solid ${$('[data-s=ibc]').value}`:'';place();},
    ibc(v){const o=selObj;if(!o)return;if(parseFloat(o.style.borderWidth)>0)o.style.borderColor=v;else{o.style.border=`1px solid ${v}`;sv('ibw',1);}place();},
    irad(v){const o=selObj;if(o)o.style.borderRadius=(+v||0)+'px';},
    hrw(v){const o=selObj;if(!o)return;o.style.border='0';o.style.borderTop=`${clamp(+v||1,1,20)}px solid ${$('[data-s=hrc]').value}`;o.style.height='0';place();},
    hrc(v){const o=selObj;if(!o)return;o.style.border='0';o.style.borderTop=`${parseFloat(o.style.borderTopWidth)||1}px solid ${v}`;o.style.height='0';},
    hrp(v){const o=selObj;if(o){o.style.width=clamp(+v||100,5,100)+'%';place();}}
  });
  ['iw','ih','isc','tw','tind','trh','tcw','hrw','hrp'].forEach(n=>{const f=sels[n];sels[n]=v=>{if(String(v)===lastFill[n])return;f(v);};});
  function ctxCaret(){
    if(selObj&&!page.contains(selObj)){pick(null);return;}
    const s=getSelection();if(!s.rangeCount||!page.contains(s.anchorNode))return;
    const ct=cellOf(s.anchorNode);if(ct)lastCell=ct;
    const t=ct&&ct.closest('table');
    if(t){if(selObj!==t&&!(selObj&&selObj.tagName!=='TABLE'))pick(t);else ctxFill();}
    else if(selObj&&selObj.tagName==='TABLE')pick(null);
  }
  page.addEventListener('keyup',ctxCaret);
  page.addEventListener('mouseup',()=>setTimeout(()=>{const s=getSelection();const ct=s.rangeCount&&page.contains(s.anchorNode)?cellOf(s.anchorNode):null;if(ct){lastCell=ct;ctxFill();}},0));
  page.addEventListener('scroll',place);desk.addEventListener('scroll',()=>{});
  new ResizeObserver(()=>place()).observe(page);

  /* ===== dropdown langsung di ribbon (pengganti popup <select> / pemilih warna bawaan) ===== */
  let wp=null,curAnchor=null;
  const closeWP=()=>{if(wp){wp.remove();wp=null;}};
  const ddSync=n=>{const sl=$(`[data-s=${n}]`),v=root.querySelector(`.dd[data-dd=${n}] .v`);if(sl&&v){const o=sl.options[sl.selectedIndex];v.textContent=o?o.text:'';}};
  function openWP(anchor,html,onPick){
    if(wp&&wp._o===anchor){closeWP();return;}
    closeWP();
    const host=$('.wd'),el=document.createElement('div');el.className='wpop';el.innerHTML=html;el._o=anchor;host.appendChild(el);
    const hr=host.getBoundingClientRect(),ar=anchor.getBoundingClientRect(),w=el.offsetWidth;
    let L=ar.left-hr.left;const T=ar.bottom-hr.top+1;
    if(L+w>hr.width-4)L=Math.max(4,hr.width-w-4);
    el.style.left=L+'px';el.style.top=T+'px';el.style.maxHeight=Math.max(120,hr.height-T-6)+'px';
    el.addEventListener('click',onPick);wp=el;
  }
  function pickMenu(items,cb){
    const a=curAnchor;if(!a)return;
    openWP(a,items.map((o,i)=>`<div class="mi" data-i="${i}"><span>${esc(o.l)}</span></div>`).join(''),e=>{const m=e.target.closest('.mi');if(!m)return;closeWP();cb(items[+m.dataset.i]);});
  }
  function openDD(btn){
    const n=btn.dataset.dd,sl=$(`[data-s=${n}]`);if(!sl)return;
    const rs=sl.hasAttribute('data-r'),cur=sl.selectedIndex,isF=n==='font';
    const html=[...sl.options].map((o,i)=>rs&&i===0?'':`<div class="mi${!rs&&i===cur?' cur':''}" data-i="${i}"${isF?` style="font-family:'${o.text}'"`:''}><span>${esc(o.text)}</span>${!rs&&i===cur?'<b>✓</b>':''}</div>`).join('');
    openWP(btn,html,e=>{const m=e.target.closest('.mi');if(!m)return;
      sl.selectedIndex=+m.dataset.i;closeWP();sl.dispatchEvent(new Event('change',{bubbles:true}));ddSync(n);});
  }
  function openCP(el){
    const inp=el.querySelector('input');
    const html=`<div class="pal">${WPAL.map(r=>`<div class="pr">${r.map(c=>`<button class="sw" data-c="${c}" style="background:${c}" title="${c}"></button>`).join('')}</div>`).join('')}<div class="mi" data-m="1"><span>🎨 Warna lainnya…</span></div></div>`;
    openWP(el,html,e=>{const b=e.target.closest('[data-c],[data-m]');if(!b)return;closeWP();
      if(b.dataset.c){inp.value=b.dataset.c;inp.dispatchEvent(new Event('input',{bubbles:true}));}else inp.click();});
  }
  root.addEventListener('click',e=>{const d=e.target.closest('.dd'),c=e.target.closest('.clr');if(d)openDD(d);else if(c&&!e.target.closest('input'))openCP(c);});
  const wpDown=e=>{if(!root.isConnected){document.removeEventListener('pointerdown',wpDown,true);return;}
    if(!wp||wp.contains(e.target)||(wp._o&&wp._o.contains(e.target)))return;closeWP();};
  document.addEventListener('pointerdown',wpDown,true);
  $$('.ribbon').forEach(r=>r.addEventListener('scroll',closeWP,{passive:true}));
  root.addEventListener('mousedown',e=>{if(e.target.closest('.b,.sty,.tab,.status button,.clr,.wpop'))e.preventDefault();});
  root.addEventListener('click',e=>{
    const t=e.target.closest('[data-a],[data-tab]');if(!t)return;
    if(t.dataset.tab){
      const small=matchMedia('(max-width:700px),(max-height:500px)').matches,wd=$('.wd');
      if(small&&t.classList.contains('on')&&!wd.classList.contains('rc')){wd.classList.add('rc');t.classList.remove('on');return;}  /* ketuk tab aktif = ciutkan pita */
      wd.classList.remove('rc');
      $$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
    const[k,v]=t.dataset.a.split(':');curAnchor=t;
    {const ae=document.activeElement;if(ae&&ae.matches&&ae.matches('.ribbon input')&&k!=='menu')ae.blur();}
    if(CMDS.has(k))exec(k);else if(acts[k])acts[k](v);
    if(k!=='menu')menu.hidden=true;
  });
  const onSel=e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||(e.type==='input'&&(e.target.tagName==='SELECT'||e.target.type==='number')))return;
    sels[s](e.target.value);if(e.target.tagName==='SELECT'&&e.target.hasAttribute('data-r'))e.target.selectedIndex=0;
    if(e.target.type==='color')e.target.previousElementSibling.style.borderBottomColor=e.target.value;};
  root.addEventListener('change',onSel);root.addEventListener('input',onSel);
  $('.desk').addEventListener('mousedown',()=>menu.hidden=true);
  $('.zr').oninput=e=>{autoFit=false;setZoom(+e.target.value);};
  if(window.ResizeObserver){let rt=0;new ResizeObserver(()=>{clearTimeout(rt);rt=setTimeout(()=>{if(!root.isConnected)return;if(autoFit)fit();},80);}).observe(desk);}
  fin.onchange=()=>{if(fin.files[0])openFile(fin.files[0]);fin.value='';};
  $('.img').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>exec('insertImage',r.result);r.readAsDataURL(f);e.target.value='';};
  page.addEventListener('click',e=>{if(e.target.closest&&e.target.closest('.eq'))eqLoad().then(c=>c.tap(e)).catch(eqFail);});
  page.addEventListener('input',stats);
  page.addEventListener('keyup',refresh);
  page.addEventListener('mouseup',()=>{refresh();
    if(painting&&!getSelection().isCollapsed){const p=painting;painting=null;
      ['bold','italic','underline'].forEach(c=>{if(document.queryCommandState(c)!==p[c])exec(c);});
      exec('fontName',p.font);setSize(p.size);if(p.color)exec('foreColor',p.color);}});
  page.addEventListener('keydown',e=>{if(e.key==='Tab'){e.preventDefault();document.execCommand('insertHTML',false,'&emsp;');}});
  root.addEventListener('keydown',e=>{
    if(e.altKey&&!e.ctrlKey&&!e.metaKey&&(e.key==='='||e.key==='+')){e.preventDefault();acts.eq();return;}
    if(!(e.ctrlKey||e.metaKey))return;const k=e.key.toLowerCase();
    if(k==='s'){e.preventDefault();saveDocx();}else if(k==='f'||k==='h'){e.preventDefault();showFind();}});
  f1.addEventListener('keydown',e=>{if(e.key==='Enter')findNext();});

  document.execCommand('styleWithCSS',false,true);
  $('[data-s=font]').value='Calibri';$('[data-s=size]').value=11;$$('.dd.vl').forEach(b=>ddSync(b.dataset.dd));
  if(matchMedia('(max-width:700px),(max-height:500px)').matches){$('.wd').classList.add('rc');$$('.tab[data-tab]').forEach(x=>x.classList.remove('on'));}
  setName(opts.file?opts.file.name:'Dokumen1.docx');
  setContent('');apply();
  applyDesign(store.get('design',{}));setSel('cstyle',cstyle);
  if(opts.file)openFile(opts.file);else focus();
}

window.Apps=window.Apps||{};
Apps.word={mount};
})();
