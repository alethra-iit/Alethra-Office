/* Rumus Word (OMML, <m:oMath>) <-> tampilan MathML di editor.
   Alur BUKA : prepare() mengganti tiap rumus di document.xml dengan token teks sebelum dibaca Mammoth
               (Mammoth membuang rumus), lalu inject() mengganti token itu dengan MathML + menyimpan OMML asli.
   Alur SIMPAN: elemen .eq membawa OMML asli di atribut data-omml; docx-io.js memasukkannya kembali ke document.xml.
   Mengekspos window.WordMath. */
(function(root){
'use strict';
const M_NS='http://schemas.openxmlformats.org/officeDocument/2006/math';
const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

/* ---------- util DOM ---------- */
const kids=(el,n)=>el?[...el.childNodes].filter(x=>x.nodeType===1&&x.localName===n):[];
const kid=(el,n)=>kids(el,n)[0]||null;
const mattr=(el,a)=>{                       // m:val dll; null bila atribut tidak ada ("" tetap "")
  if(!el)return null;
  if(el.hasAttributeNS&&el.hasAttributeNS(M_NS,a))return el.getAttributeNS(M_NS,a);
  return el.hasAttribute('m:'+a)?el.getAttribute('m:'+a):null;
};
const on=el=>{if(!el)return false;const v=mattr(el,'val');return v===null||!/^(0|false|off)$/i.test(v);};
const chrVal=(pr,tag,dflt)=>{const v=mattr(kid(pr,tag),'val');return v===null?dflt:v;};

/* ---------- OMML -> MathML ---------- */
const ACCENT={'\u0302':'^','\u0303':'~','\u0307':'\u02D9','\u0308':'\u00A8','\u0304':'\u00AF','\u20D7':'\u2192','\u0301':'\u00B4','\u0300':'`','\u030C':'\u02C7','\u0306':'\u02D8'};

function runML(r){
  const rPr=kid(r,'rPr');
  const sty=rPr?mattr(kid(rPr,'sty'),'val'):null,nor=rPr?kid(rPr,'nor'):null;
  let txt='';
  for(const c of r.childNodes)if(c.nodeType===1&&c.localName==='t')txt+=c.textContent;
  if(!txt)return '';
  if(on(nor))return `<mtext>${esc(txt)}</mtext>`;
  const bold=sty==='b'||sty==='bi',plain=sty==='p'||sty==='b';
  const st=bold?' style="font-weight:bold"':'';
  let out='';
  const re=/(\d+(?:[.,]\d+)?)|(\p{L}+)|(\s+)|([\s\S])/gu;
  for(let m;(m=re.exec(txt));){
    if(m[1])out+=`<mn${st}>${esc(m[1])}</mn>`;
    else if(m[2]){
      if(plain)out+=`<mi mathvariant="normal"${st}>${esc(m[2])}</mi>`;          // mis. sin, lim: satu token tegak
      else for(const ch of m[2])out+=`<mi${st}>${esc(ch)}</mi>`;                 // variabel: huruf miring satu-satu
    }
    else if(m[3])continue;
    else out+=`<mo>${esc(m[4]==='-'?'\u2212':m[4])}</mo>`;
  }
  return out;
}
const seq=el=>{let s='';if(el)for(const n of el.childNodes)if(n.nodeType===1)s+=node(n);return s;};
const arg=(p,name)=>`<mrow>${seq(kid(p,name))}</mrow>`;
const hasContent=el=>!!(el&&el.textContent.trim()||el&&el.getElementsByTagNameNS(M_NS,'r').length);
const mo=(c,attrs)=>`<mo${attrs||''}>${esc(ACCENT[c]||c)}</mo>`;

function node(n){
  const ln=n.localName;
  switch(ln){
    case 'r':return runML(n);
    case 'f':{
      const f=chrVal(kid(n,'fPr'),'type','bar')||'bar';
      if(f==='skw'||f==='lin')return `<mrow>${arg(n,'num')}<mo>/</mo>${arg(n,'den')}</mrow>`;
      return `<mfrac${f==='noBar'?' linethickness="0"':''}>${arg(n,'num')}${arg(n,'den')}</mfrac>`;
    }
    case 'rad':{
      const hide=on(kid(kid(n,'radPr'),'degHide'));
      if(hide||!hasContent(kid(n,'deg')))return `<msqrt>${arg(n,'e')}</msqrt>`;
      return `<mroot>${arg(n,'e')}${arg(n,'deg')}</mroot>`;
    }
    case 'sSup':return `<msup>${arg(n,'e')}${arg(n,'sup')}</msup>`;
    case 'sSub':return `<msub>${arg(n,'e')}${arg(n,'sub')}</msub>`;
    case 'sSubSup':return `<msubsup>${arg(n,'e')}${arg(n,'sub')}${arg(n,'sup')}</msubsup>`;
    case 'sPre':return `<mmultiscripts>${arg(n,'e')}<mprescripts/>${arg(n,'sub')}${arg(n,'sup')}</mmultiscripts>`;
    case 'nary':{
      const pr=kid(n,'naryPr'),chr=chrVal(pr,'chr','\u222B');
      const loc=chrVal(pr,'limLoc',/[\u222B-\u2233]/.test(chr)?'subSup':'undOvr');
      const hs=on(kid(pr,'subHide'))||!hasContent(kid(n,'sub')),hp=on(kid(pr,'supHide'))||!hasContent(kid(n,'sup'));
      const op=`<mo largeop="true" movablelimits="false" stretchy="false">${esc(chr)}</mo>`,u=loc==='undOvr';
      let s;
      if(!hs&&!hp)s=`<${u?'munderover':'msubsup'}>${op}${arg(n,'sub')}${arg(n,'sup')}</${u?'munderover':'msubsup'}>`;
      else if(!hs)s=`<${u?'munder':'msub'}>${op}${arg(n,'sub')}</${u?'munder':'msub'}>`;
      else if(!hp)s=`<${u?'mover':'msup'}>${op}${arg(n,'sup')}</${u?'mover':'msup'}>`;
      else s=op;
      return `<mrow>${s}${arg(n,'e')}</mrow>`;
    }
    case 'd':{
      const pr=kid(n,'dPr'),b=chrVal(pr,'begChr','('),e=chrVal(pr,'endChr',')'),sp=chrVal(pr,'sepChr','|');
      const es=kids(n,'e');let s='';
      es.forEach((x,i)=>{if(i&&sp)s+=`<mo separator="true" stretchy="true">${esc(sp)}</mo>`;s+=`<mrow>${seq(x)}</mrow>`;});
      return `<mrow>${b?`<mo form="prefix" fence="true" stretchy="true">${esc(b)}</mo>`:''}${s}${e?`<mo form="postfix" fence="true" stretchy="true">${esc(e)}</mo>`:''}</mrow>`;
    }
    case 'func':return `<mrow>${arg(n,'fName')}<mo>\u2061</mo>${arg(n,'e')}</mrow>`;
    case 'limLow':return `<munder>${arg(n,'e')}${arg(n,'lim')}</munder>`;
    case 'limUpp':return `<mover>${arg(n,'e')}${arg(n,'lim')}</mover>`;
    case 'acc':{
      const c=chrVal(kid(n,'accPr'),'chr','\u0302');
      return `<mover accent="true">${arg(n,'e')}${mo(c,' stretchy="true"')}</mover>`;
    }
    case 'bar':{
      const top=chrVal(kid(n,'barPr'),'pos','top')!=='bot';
      return `<${top?'mover':'munder'}>${arg(n,'e')}<mo stretchy="true">\u00AF</mo></${top?'mover':'munder'}>`;
    }
    case 'groupChr':{
      const pr=kid(n,'groupChrPr'),c=chrVal(pr,'chr','\u23DF'),top=chrVal(pr,'pos','bot')==='top';
      return `<${top?'mover':'munder'}>${arg(n,'e')}<mo stretchy="true">${esc(c)}</mo></${top?'mover':'munder'}>`;
    }
    case 'box':return arg(n,'e');
    case 'phant':{
      const pr=kid(n,'phantPr'),show=kid(pr,'show');
      return show&&!on(show)?`<mphantom>${arg(n,'e')}</mphantom>`:arg(n,'e');
    }
    case 'borderBox':return `<mrow style="border:1px solid currentColor;padding:2px">${seq(kid(n,'e'))}</mrow>`;
    case 'eqArr':return `<mtable columnalign="left">${kids(n,'e').map(x=>`<mtr><mtd>${seq(x)}</mtd></mtr>`).join('')}</mtable>`;
    case 'm':return `<mtable>${kids(n,'mr').map(r=>`<mtr>${kids(r,'e').map(x=>`<mtd>${seq(x)}</mtd>`).join('')}</mtr>`).join('')}</mtable>`;
    case 'oMath':return seq(n);
    default:
      if(/Pr$/.test(ln))return '';          // m:fPr, m:ctrlPr, dll. = properti, bukan isi
      return seq(n);                          // m:e, m:num, wadah lain -> isi saja
  }
}

function toMathML(omml,display,DP){           // omml: elemen m:oMath / m:oMathPara
  DP=DP||root.DOMParser;
  let body='';
  if(omml.localName==='oMathPara'){display=true;body=kids(omml,'oMath').map(m=>`<mrow>${seq(m)}</mrow>`).join('<mspace linebreak="newline"/>');}
  else body=`<mrow>${seq(omml)}</mrow>`;
  return `<math xmlns="http://www.w3.org/1998/Math/MathML" display="${display?'block':'inline'}">${body}</math>`;
}
const plainText=el=>[...el.getElementsByTagNameNS(M_NS,'t')].map(t=>t.textContent).join('');

/* WebView lama (Chromium < 109) belum punya MathML -> tampilkan teks linear */
let _sup=null;
function supported(){
  if(_sup!==null)return _sup;
  try{
    const d=document.createElement('div');d.style.cssText='position:absolute;visibility:hidden';
    d.innerHTML='<math><mspace height="23px" width="77px"></mspace></math>';
    document.body.appendChild(d);_sup=Math.abs(d.firstChild.firstChild.getBoundingClientRect().height-23)<1;d.remove();
  }catch(e){_sup=false;}
  return _sup;
}

/* ---------- BUKA: token pengganti rumus ---------- */
async function prepare(buf,JSZip,DP,XS){
  DP=DP||root.DOMParser;XS=XS||root.XMLSerializer;
  const none={buf,maths:[]};
  const zip=await JSZip.loadAsync(buf);
  const f=zip.file('word/document.xml');if(!f)return none;
  const xml=await f.async('string');
  if(!xml.includes('oMath'))return none;                  // dokumen tanpa rumus: jangan ubah apa pun
  const doc=new DP().parseFromString(xml,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)return none;
  const isM=e=>e.namespaceURI===M_NS&&(e.localName==='oMath'||e.localName==='oMathPara');
  const top=[...doc.getElementsByTagName('*')].filter(e=>{
    if(!isM(e))return false;
    for(let a=e.parentNode;a&&a.nodeType===1;a=a.parentNode)if(isM(a))return false;
    return true;
  });
  if(!top.length)return none;
  const ser=new XS(),maths=[];
  for(const el of top){
    const i=maths.length;
    maths.push({omml:ser.serializeToString(el),display:el.localName==='oMathPara'});
    const r=doc.createElementNS(W_NS,'w:r'),t=doc.createElementNS(W_NS,'w:t');
    t.textContent='@@ALMATH'+i+'@@';r.appendChild(t);el.parentNode.replaceChild(r,el);
  }
  zip.file('word/document.xml',ser.serializeToString(doc));
  return {buf:await zip.generateAsync({type:'arraybuffer'}),maths};
}

/* ---------- BUKA: token -> elemen rumus ---------- */
function eqHTML(m,DP){
  DP=DP||root.DOMParser;
  let inner='';
  try{
    const d=new DP().parseFromString(m.omml,'application/xml');
    const el=d.documentElement;
    inner=supported()?toMathML(el,m.display,DP):esc(plainText(el));
  }catch(e){inner='[rumus]';}
  return `<span class="eq${m.display?' eqd':''}" contenteditable="false" data-omml="${esc(m.omml)}">${inner}</span>`;
}
function inject(html,maths,DP){
  if(!maths||!maths.length)return html;
  return html.replace(/@@ALMATH(\d+)@@/g,(s,i)=>maths[+i]?eqHTML(maths[+i],DP):s);
}

root.WordMath={prepare,inject,toMathML,supported,M_NS};
})(typeof window!=='undefined'?window:globalThis);
