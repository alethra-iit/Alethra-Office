/* Editor rumus Word: LaTeX <-> OMML (<m:oMath>) + dialog sisip/edit + bilah ketuk-rumus.
   - tex2omml(tex,opt)  : teks LaTeX -> string OMML (m:oMath / m:oMathPara)
   - omml2tex(el)       : elemen OMML -> {tex,lossy,display}  (supaya rumus yang sudah ada bisa diedit ulang)
   - attach(ctx)        : pasang perilaku di editor Word (ketuk rumus, dialog, ganti/sisipkan)
   Bergantung pada window.WordMath (omml.js) hanya untuk pratinjau MathML. Mengekspos window.WordEq. */
(function(root){
'use strict';
const M_NS='http://schemas.openxmlformats.org/officeDocument/2006/math';
const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xe=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const xa=s=>xe(s).replace(/"/g,'&quot;');
class EqError extends Error{constructor(m){super(m);this.name='EqError';}}
const err=m=>{throw new EqError(m);};

/* ============================ TABEL ============================ */
const SYM={};
const defS=(chars,names)=>{names.split(' ').forEach((n,i)=>{SYM[n]=chars[i];});};
defS('αβγδϵεζηθϑικλμνξπϖρϱσςτυϕφχψω','alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi pi varpi rho varrho sigma varsigma tau upsilon phi varphi chi psi omega');
defS('ΓΔΘΛΞΠΣΥΦΨΩ','Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega');
Object.assign(SYM,{
  times:'×',div:'÷',pm:'±',mp:'∓',cdot:'⋅',ast:'∗',star:'⋆',circ:'∘',bullet:'∙',oplus:'⊕',ominus:'⊖',otimes:'⊗',oslash:'⊘',odot:'⊙',
  wedge:'∧',land:'∧',vee:'∨',lor:'∨',cap:'∩',cup:'∪',uplus:'⊎',sqcap:'⊓',sqcup:'⊔',setminus:'∖',
  leq:'≤',le:'≤',geq:'≥',ge:'≥',neq:'≠',ne:'≠',approx:'≈',equiv:'≡',sim:'∼',simeq:'≃',cong:'≅',propto:'∝',ll:'≪',gg:'≫',prec:'≺',succ:'≻',preceq:'⪯',succeq:'⪰',doteq:'≐',
  in:'∈',notin:'∉',ni:'∋',subset:'⊂',supset:'⊃',subseteq:'⊆',supseteq:'⊇',nsubseteq:'⊈',sqsubseteq:'⊑',sqsupseteq:'⊒',
  forall:'∀',exists:'∃',nexists:'∄',neg:'¬',lnot:'¬',emptyset:'∅',varnothing:'∅',infty:'∞',partial:'∂',nabla:'∇',angle:'∠',measuredangle:'∡',
  perp:'⊥',parallel:'∥',nparallel:'∦',mid:'∣',nmid:'∤',prime:'′',dprime:'″',degree:'°',
  ldots:'…',dots:'…',cdots:'⋯',vdots:'⋮',ddots:'⋱',
  to:'→',rightarrow:'→',leftarrow:'←',gets:'←',leftrightarrow:'↔',Rightarrow:'⇒',Leftarrow:'⇐',Leftrightarrow:'⇔',implies:'⟹',iff:'⟺',impliedby:'⟸',
  mapsto:'↦',uparrow:'↑',downarrow:'↓',updownarrow:'↕',nearrow:'↗',searrow:'↘',swarrow:'↙',nwarrow:'↖',longrightarrow:'⟶',longleftarrow:'⟵',
  longleftrightarrow:'⟷',hookrightarrow:'↪',rightleftharpoons:'⇌',
  hbar:'ℏ',ell:'ℓ',Re:'ℜ',Im:'ℑ',aleph:'ℵ',wp:'℘',therefore:'∴',because:'∵',square:'□',Box:'□',triangle:'△',diamond:'⋄',dagger:'†',ddagger:'‡',
  S:'§',P:'¶',copyright:'©',pounds:'£',euro:'€',checkmark:'✓',top:'⊤',bot:'⊥',vdash:'⊢',dashv:'⊣',models:'⊨',
  langle:'⟨',rangle:'⟩',lfloor:'⌊',rfloor:'⌋',lceil:'⌈',rceil:'⌉',backslash:'\\',lbrace:'{',rbrace:'}',vert:'|',Vert:'‖',
  textasciitilde:'~',
  /* spasi */
  quad:'\u2003',qquad:'\u2003\u2003',enspace:'\u2002'
});
const ESC1={'{':'{','}':'}','%':'%','&':'&','#':'#','$':'$','_':'_','|':'‖','^':'^','~':'~',' ':' ',',':'\u2009',':':'\u2005',';':'\u2004','!':''};
const SYM_REV={};   /* karakter -> perintah (yang pertama didefinisikan = pilihan utama) */
['alpha','beta','gamma','delta','epsilon','varepsilon','zeta','eta','theta','vartheta','iota','kappa','lambda','mu','nu','xi','pi','varpi','rho','varrho','sigma','varsigma','tau','upsilon','phi','varphi','chi','psi','omega',
 'Gamma','Delta','Theta','Lambda','Xi','Pi','Sigma','Upsilon','Phi','Psi','Omega','times','div','pm','mp','cdot','ast','star','circ','bullet','oplus','ominus','otimes','oslash','odot','wedge','vee','cap','cup','uplus','sqcap','sqcup','setminus',
 'leq','geq','neq','approx','equiv','sim','simeq','cong','propto','ll','gg','prec','succ','preceq','succeq','doteq','in','notin','ni','subset','supset','subseteq','supseteq','nsubseteq','sqsubseteq','sqsupseteq',
 'forall','exists','nexists','neg','emptyset','infty','partial','nabla','angle','measuredangle','perp','parallel','nparallel','mid','nmid','prime','dprime','degree','ldots','cdots','vdots','ddots',
 'to','leftarrow','leftrightarrow','Rightarrow','Leftarrow','Leftrightarrow','implies','iff','impliedby','mapsto','uparrow','downarrow','updownarrow','nearrow','searrow','swarrow','nwarrow','longrightarrow','longleftarrow','longleftrightarrow','hookrightarrow','rightleftharpoons',
 'hbar','ell','Re','Im','aleph','wp','therefore','because','square','triangle','diamond','dagger','ddagger','S','P','copyright','pounds','euro','checkmark','top','vdash','dashv','models',
 'langle','rangle','lfloor','rfloor','lceil','rceil'].forEach(n=>{if(SYM[n]&&!SYM_REV[SYM[n]])SYM_REV[SYM[n]]=n;});

const FUNCS=new Set('sin cos tan cot sec csc arcsin arccos arctan sinh cosh tanh coth log ln lg exp lim limsup liminf max min sup inf det dim ker gcd deg hom arg Pr'.split(' '));
const LIMFUNCS=new Set(['lim','limsup','liminf','max','min','sup','inf','det','gcd','Pr']);
const NARY={sum:'∑',prod:'∏',coprod:'∐',int:'∫',iint:'∬',iiint:'∭',oint:'∮',oiint:'∯',oiiint:'∰',bigcup:'⋃',bigcap:'⋂',bigvee:'⋁',bigwedge:'⋀',bigoplus:'⨁',bigotimes:'⨂',bigodot:'⨀',biguplus:'⨄',bigsqcup:'⨆'};
const NARY_REV={};Object.keys(NARY).forEach(k=>{NARY_REV[NARY[k]]=k;});
const ACC={hat:'\u0302',widehat:'\u0302',check:'\u030C',tilde:'\u0303',widetilde:'\u0303',acute:'\u0301',grave:'\u0300',dot:'\u0307',ddot:'\u0308',breve:'\u0306',bar:'\u0304',vec:'\u20D7',overrightarrow:'\u20D7'};
const ACC_REV={};['hat','check','tilde','acute','grave','dot','ddot','breve','bar','vec'].forEach(k=>{ACC_REV[ACC[k]]=k;});
const DELIM={'{':'{','}':'}','|':'‖',langle:'⟨',rangle:'⟩',lfloor:'⌊',rfloor:'⌋',lceil:'⌈',rceil:'⌉',lvert:'|',rvert:'|',lVert:'‖',rVert:'‖',vert:'|',Vert:'‖',lbrace:'{',rbrace:'}',backslash:'\\',uparrow:'↑',downarrow:'↓'};
const DELIM_REV={'{':'\\{','}':'\\}','‖':'\\|','⟨':'\\langle','⟩':'\\rangle','⌊':'\\lfloor','⌋':'\\rfloor','⌈':'\\lceil','⌉':'\\rceil','\\':'\\backslash','':'.'};
const BREAKCMD=new Set('pm mp leq le geq ge neq ne approx equiv sim simeq cong propto ll gg to rightarrow leftarrow leftrightarrow Rightarrow Leftarrow Leftrightarrow implies iff in notin ni subset supset subseteq supseteq mapsto'.split(' '));
const BREAKCH='+-=<>,;';
const MATENV={matrix:['',''],pmatrix:['(',')'],bmatrix:['[',']'],Bmatrix:['{','}'],vmatrix:['|','|'],Vmatrix:['‖','‖'],smallmatrix:['',''],array:['','']};
const ARRENV=new Set('aligned align align* alignat alignat* split gather gather* gathered eqnarray eqnarray* flalign flalign*'.split(' '));
const FONTCMD={mathrm:'p',textrm:'p',mathup:'p',textup:'p',mathbf:'b',textbf:'b',mathit:'i',textit:'i',mathnormal:'i',boldsymbol:'bi',bm:'bi',mathbfit:'bi',mathsf:'sf',textsf:'sf',mathtt:'tt',texttt:'tt',mathbb:'bb',mathcal:'cal',mathscr:'cal',mathfrak:'frak'};
/* warna (nama Inggris/Indonesia atau hex) & ukuran */
const COLORS={black:'000000',white:'FFFFFF',red:'FF0000',green:'008000',blue:'0000FF',yellow:'FFFF00',orange:'FF8000',purple:'800080',gray:'808080',grey:'808080',brown:'8B4513',cyan:'00FFFF',magenta:'FF00FF',pink:'FF69B4',teal:'008080',navy:'000080',lime:'00FF00',olive:'808000',maroon:'800000',darkgray:'404040',lightgray:'BFBFBF',
  hitam:'000000',putih:'FFFFFF',merah:'FF0000',hijau:'008000',biru:'0000FF',kuning:'FFFF00',oranye:'FF8000',jingga:'FF8000',ungu:'800080',abu:'808080',cokelat:'8B4513',merahmuda:'FF69B4'};
function colorHex(raw,model){
  const s=String(raw).trim().replace(/^#/,'');
  if(model&&/^rgb$/i.test(model)){const q=s.split(',').map(Number);if(q.length===3&&q.every(x=>x>=0&&x<=1))return q.map(x=>Math.round(x*255).toString(16).padStart(2,'0')).join('').toUpperCase();err('Warna rgb tidak valid: '+raw);}
  if(/^[0-9a-fA-F]{6}$/.test(s))return s.toUpperCase();
  if(/^[0-9a-fA-F]{3}$/.test(s))return s.split('').map(c=>c+c).join('').toUpperCase();
  const k=s.toLowerCase();if(COLORS[k])return COLORS[k];
  err('Warna tidak dikenal: '+raw+' (pakai nama seperti red/merah atau kode seperti FF0000)');
}
const SIZECMD={tiny:7,scriptsize:8,footnotesize:9,small:10,normalsize:0,large:14,Large:17,LARGE:20,huge:24,Huge:28};
const mixSty=(o,n)=>{                              /* gabung gaya huruf: tebal + miring = tebal-miring */
  if(!o)return n;const b=/b/.test(o),i=/i/.test(o);
  if(n==='b')return i?'bi':'b';if(n==='i')return b?'bi':'i';if(n==='bi')return 'bi';if(n==='p')return b?'b':'p';
  return n;
};

/* alfabet matematika Unicode (mathbb, mathcal, mathfrak, mathsf, mathtt) */
const ALPHA={
  bb:{U:0x1D538,L:0x1D552,D:0x1D7D8,ex:{C:'ℂ',H:'ℍ',N:'ℕ',P:'ℙ',Q:'ℚ',R:'ℝ',Z:'ℤ'}},
  cal:{U:0x1D49C,L:0x1D4B6,ex:{B:'ℬ',E:'ℰ',F:'ℱ',H:'ℋ',I:'ℐ',L:'ℒ',M:'ℳ',R:'ℛ',e:'ℯ',g:'ℊ',o:'ℴ'}},
  frak:{U:0x1D504,L:0x1D51E,ex:{C:'ℭ',H:'ℌ',I:'ℑ',R:'ℜ',Z:'ℨ'}},
  sf:{U:0x1D5A0,L:0x1D5BA,D:0x1D7E2},
  tt:{U:0x1D670,L:0x1D68A,D:0x1D7F6}
};
const ALPHA_CMD={bb:'mathbb',cal:'mathcal',frak:'mathfrak',sf:'mathsf',tt:'mathtt'};
const ALPHA_REV={};
function alphaMap(kind,ch){
  const a=ALPHA[kind];if(!a)return ch;
  if(a.ex&&a.ex[ch])return a.ex[ch];
  const c=ch.charCodeAt(0);
  if(c>=65&&c<=90)return String.fromCodePoint(a.U+c-65);
  if(c>=97&&c<=122)return String.fromCodePoint(a.L+c-97);
  if(a.D&&c>=48&&c<=57)return String.fromCodePoint(a.D+c-48);
  return ch;
}
for(const k in ALPHA){
  for(let c=48;c<=122;c++){const ch=String.fromCharCode(c);if(!/[0-9A-Za-z]/.test(ch))continue;const m=alphaMap(k,ch);if(m!==ch&&!ALPHA_REV[m])ALPHA_REV[m]=[k,ch];}
}

/* ============================ PENGURAI (LaTeX -> AST) ============================ */
function tokenize(src){
  const T=[];let i=0;const n=src.length;
  while(i<n){
    const c=src[i];
    if(/\s/.test(c)){i++;continue;}
    const st=i;
    if(c==='\\'){
      let j=i+1;
      if(j>=n)err('Garis miring di akhir rumus');
      if(/[A-Za-z]/.test(src[j])){while(j<n&&/[A-Za-z]/.test(src[j]))j++;if(src[j]==='*'&&/^(align|alignat|gather|flalign|eqnarray)$/.test(src.slice(i+1,j))&&false)j++;T.push({k:'cmd',v:src.slice(i+1,j),s:st,e:j});}
      else{const ch=String.fromCodePoint(src.codePointAt(j));j+=ch.length;T.push({k:'cmd',v:ch,s:st,e:j});}
      i=j;continue;
    }
    if('{}^_&\''.includes(c)){T.push({k:c,s:st,e:i+1});i++;continue;}
    const ch=String.fromCodePoint(src.codePointAt(i));i+=ch.length;
    T.push({k:'ch',v:ch,s:st,e:i});
  }
  return T;
}

function parseTex(src,opt){
  opt=opt||{};
  const T=tokenize(src);let p=0;
  const peek=()=>T[p],next=()=>T[p++];
  const R=(s,o)=>Object.assign({t:'r',s},o||{});
  const isStopGen=t=>t.k==='}'||(t.k==='cmd'&&(t.v==='right'||t.v==='middle'||t.v==='end'||t.v==='\\'));
  const isBreak=t=>t.k==='ch'?BREAKCH.includes(t.v):(t.k==='cmd'&&BREAKCMD.has(t.v));
  const expect=(k,msg)=>{const t=next();if(!t||t.k!==k)err(msg||('Diharapkan '+k));return t;};

  function rawGroup(){            /* baca isi {...} mentah (untuk \text, nama environment) */
    const t=peek();if(!t||t.k!=='{')err('Diharapkan { ');
    let depth=0,j=t.s;
    for(;j<src.length;j++){
      const c=src[j];
      if(c==='\\'){j++;continue;}
      if(c==='{')depth++;
      else if(c==='}'){depth--;if(depth===0)break;}
    }
    if(depth!==0)err('Kurung { tidak ditutup');
    const body=src.slice(t.s+1,j);
    while(p<T.length&&T[p].s<=j)p++;
    return body;
  }
  function group(){               /* {...} -> node[] */
    expect('{');
    const c=seq(t=>t.k==='}'||(t.k==='cmd'&&(t.v==='right'||t.v==='middle'||t.v==='end')));
    const t=next();if(!t||t.k!=='}')err('Kurung { tidak ditutup');
    return c;
  }
  function seq(stop){
    const out=[];
    for(;;){
      const t=peek();if(!t||stop(t))break;
      const n=term();if(!n)continue;
      if(n.t==='cset'){                             /* \color{..} / \large: berlaku sampai akhir kelompok */
        const rest=seq(stop);applyStyle(rest,n.st);out.push({t:'g',c:rest,style:true});break;
      }
      out.push(n);
    }
    return out;
  }
  function delim(){
    const t=next();if(!t)err('Pembatas hilang setelah \\left/\\right');
    if(t.k==='ch')return t.v==='.'?'':t.v;
    if(t.k==='cmd'){if(DELIM[t.v]!==undefined)return DELIM[t.v];if(SYM[t.v])return SYM[t.v];}
    err('Pembatas tidak dikenal');
  }
  function scriptArg(){
    const t=peek();if(!t||isStopGen(t))err('Argumen kurang: setelah ^, _ atau perintah masih perlu isi');
    if(t.k==='{')return group();
    const a=atom(true);return a?[a]:[];
  }
  function mapNodes(nodes,fn){                      /* kunjungi SEMUA simpul (run & struktur) */
    for(const n of nodes){
      if(!n||typeof n!=='object')continue;
      fn(n);
      for(const k of ['c','n','d','deg','e','sub','sup','lim','name'])if(Array.isArray(n[k]))mapNodes(n[k],fn);
      if(n.es)n.es.forEach(x=>mapNodes(x,fn));
      if(n.rows)n.rows.forEach(r=>{if(!r.length||Array.isArray(r[0]))r.forEach(c=>mapNodes(c,fn));else mapNodes(r,fn);});
    }
  }
  function applyStyle(nodes,st){                    /* warna/ukuran/sorot; gaya yang lebih dalam menang */
    mapNodes(nodes,n=>{for(const k in st)if(n[k]==null)n[k]=st[k];});
  }
  function operand(){             /* isi operator n-ary: sampai operator +,-,= dst. di tingkat yang sama */
    const t=peek();if(!t||isStopGen(t))return [];
    if(t.k==='{'){const g=atom();return [g];}
    const out=[];let depth=0;
    for(;;){
      const q=peek();if(!q||isStopGen(q))break;
      if(depth===0&&(isBreak(q)||(q.k==='ch'&&(q.v===')'||q.v===']'))))break;
      const n=term();if(!n)continue;out.push(n);
      if(n.t==='r'&&n.s.length===1){if('(['.includes(n.s))depth++;else if(')]'.includes(n.s))depth=Math.max(0,depth-1);}
    }
    return out;
  }
  function funcArg(){
    const t=peek();if(!t||isStopGen(t)||isBreak(t))return [];
    if(t.k==='{')return [atom()];
    if(t.k==='ch'&&t.v==='('){
      const out=[];let depth=0;
      for(;;){const q=peek();if(!q)break;const n=term();if(!n)continue;out.push(n);
        if(n.t==='r'&&n.s==='(')depth++;else if(n.t==='r'&&n.s===')'){depth--;if(depth<=0)break;}
        else if(n.t==='sup'&&n.e.length===1&&n.e[0].s===')'){break;}}
      return out;
    }
    const n=term();return n?[n]:[];
  }
  function term(){
    const a=atom();if(!a)return null;
    if(a.t==='fnop')return a;   /* diproses di bawah */
    return scripts(a);
  }
  function scripts(a){
    let nary=null;
    if(a.t==='nary'){
      const q=peek();
      if(q&&q.k==='cmd'&&(q.v==='limits'||q.v==='nolimits')){next();a.loc=q.v==='limits'?'undOvr':'subSup';}
    }
    let sub=null,sup=null,pr=0;
    for(;;){
      const q=peek();if(!q)break;
      if(q.k==='_'){next();if(sub)err('Subskrip ganda');sub=scriptArg();}
      else if(q.k==='^'){next();if(sup&&!pr)err('Superskrip ganda');sup=(sup||[]).concat(scriptArg());pr=0;}
      else if(q.k==='\''){next();sup=(sup||[]).concat(R('′'));pr=1;}
      else break;
    }
    return build(a,sub,sup);
  }
  function build(a,sub,sup){
    if(a.t==='nary'){a.sub=sub;a.sup=sup;a.e=operand();return a;}
    if(a.t==='fn'){
      let nm=[R(a.name,{sty:'p'})];
      if(sub&&sup)nm=[{t:'ssub',e:nm,sub,sup}];
      else if(sub)nm=LIMFUNCS.has(a.name)?[{t:'limlow',e:nm,lim:sub}]:[{t:'sub',e:nm,sub}];
      else if(sup)nm=[{t:'sup',e:nm,sup}];
      return {t:'func',name:nm,e:funcArg()};
    }
    if(a.t==='gchr'&&(sub||sup)){let r=a;if(sub)r={t:'limlow',e:[r],lim:sub};if(sup)r={t:'limupp',e:[r],lim:sup};return r;}
    if(!sub&&!sup)return a;
    if(a.t==='g'&&a.empty){const e=atom();return {t:'pre',sub:sub||[],sup:sup||[],e:e?[e]:[]};}
    if(sub&&sup)return {t:'ssub',e:[a],sub,sup};
    if(sub)return {t:'sub',e:[a],sub};
    return {t:'sup',e:[a],sup};
  }
  function styled(kind,nodes){
    mapNodes(nodes,r=>{
      if(r.t!=='r'||r.nor)return;
      if(ALPHA[kind]){r.s=[...r.s].map(ch=>alphaMap(kind,ch)).join('');}
      else r.sty=mixSty(r.sty,kind);
    });
    return {t:'g',c:nodes,style:true};
  }
  function optModel(){                              /* [HTML] / [rgb] sebelum nama warna */
    const q=peek();if(!q||q.k!=='ch'||q.v!=='[')return null;
    next();let m='';
    for(;;){const t=next();if(!t)err('Kurung siku [ tidak ditutup');if(t.k==='ch'&&t.v===']')break;m+=t.v||'';}
    return m;
  }
  function env(name){
    if(name==='array'&&peek()&&peek().k==='{')rawGroup();
    if(/^alignat\*?$/.test(name)&&peek()&&peek().k==='{')rawGroup();
    const isMat=MATENV[name],isCases=name==='cases'||name==='dcases';
    if(!isMat&&!isCases&&!ARRENV.has(name))err('Environment tidak dikenal: '+name);
    const rows=[];let row=[];
    for(;;){
      const cell=seq(t=>isStopGen(t)||(isMat&&t.k==='&'));
      const t=next();if(!t)err('\\begin{'+name+'} tidak ditutup');
      if(t.k==='&'){row.push(cell);continue;}
      if(t.k==='cmd'&&t.v==='\\'){row.push(cell);rows.push(row);row=[];continue;}
      if(t.k==='cmd'&&t.v==='end'){
        const en=rawGroup();if(en!==name)err('\\end{'+en+'} tidak cocok dengan \\begin{'+name+'}');
        row.push(cell);if(row.length>1||row[0].length)rows.push(row);break;
      }
      err('Tanda tidak terduga di dalam '+name);
    }
    if(isMat){
      const w=Math.max(1,...rows.map(r=>r.length));rows.forEach(r=>{while(r.length<w)r.push([]);});
      const m={t:'mat',rows};const d=MATENV[name];
      return d[0]||d[1]?{t:'d',beg:d[0],end:d[1],sep:'|',es:[[m]]}:m;
    }
    const ea={t:'eqarr',rows:rows.map(r=>r[0]||[])};
    return isCases?{t:'d',beg:'{',end:'',sep:'|',es:[[ea]]}:ea;
  }
  function atom(noEmptyScripts){
    const t=peek();if(!t)err('Rumus tidak lengkap');
    if(t.k==='^'||t.k==='_')return {t:'g',c:[],empty:true};
    next();
    switch(t.k){
      case '{':{p--;const c=group();return {t:'g',c,empty:!c.length};}
      case '}':err('Kurung } tanpa pasangan');break;
      case '&':return R('&');
      case '\'':return R('′');
      case 'ch':return R(t.v==='-'?'\u2212':t.v==='~'?'\u00A0':t.v);
    }
    const v=t.v;
    if(SYM[v]!==undefined)return R(SYM[v]);
    if(v.length===1&&!/[A-Za-z]/.test(v)){
      if(v==='\\')return {t:'g',c:[],nl:true};
      if(ESC1[v]!==undefined)return R(ESC1[v]);
      err('Perintah tidak dikenal: \\'+v);
    }
    if(SIZECMD[v]!==undefined&&!SYM[v])return {t:'cset',st:{sz:SIZECMD[v]}};
    if(FUNCS.has(v))return {t:'fn',name:v};
    if(NARY[v])return {t:'nary',chr:NARY[v],loc:/[\u222B-\u2233]/.test(NARY[v])?'subSup':'undOvr'};
    if(ACC[v]){const e=atom();return {t:'acc',chr:ACC[v],e:e?[e]:[]};}
    if(FONTCMD[v]){const a=atom();return styled(FONTCMD[v],a.t==='g'?a.c:[a]);}
    switch(v){
      case 'frac':case 'dfrac':case 'tfrac':case 'cfrac':{const n=scriptArg(),d=scriptArg();return {t:'f',type:'bar',n,d};}
      case 'sfrac':{const n=scriptArg(),d=scriptArg();return {t:'f',type:'skw',n,d};}
      case 'nicefrac':{const n=scriptArg(),d=scriptArg();return {t:'f',type:'lin',n,d};}
      case 'binom':case 'dbinom':case 'tbinom':{const n=scriptArg(),d=scriptArg();return {t:'d',beg:'(',end:')',sep:'|',es:[[{t:'f',type:'noBar',n,d}]]};}
      case 'genfrac':{
        const txt=a=>(a.length&&a[0].s)||'';
        const ld=txt(group()),rd=txt(group()),th=group().map(x=>x.s||'').join(''),sty=group(),n=group(),d=group();
        const f={t:'f',type:/^0/.test(th)?'noBar':'bar',n,d};
        return ld||rd?{t:'d',beg:ld,end:rd,sep:'|',es:[[f]]}:f;
      }
      case 'sqrt':{
        let deg=null;const q=peek();
        if(q&&q.k==='ch'&&q.v==='['){next();deg=seq(x=>x.k==='ch'&&x.v===']');const c=next();if(!c)err('Kurung siku [ tidak ditutup');}
        return {t:'rad',deg,e:scriptArg()};
      }
      case 'overline':return {t:'bar',pos:'top',e:scriptArg()};
      case 'underline':return {t:'bar',pos:'bot',e:scriptArg()};
      case 'overbrace':return {t:'gchr',chr:'⏞',pos:'top',e:scriptArg()};
      case 'underbrace':return {t:'gchr',chr:'⏟',pos:'bot',e:scriptArg()};
      case 'overset':{const l=scriptArg(),e=scriptArg();return {t:'limupp',e,lim:l};}
      case 'underset':{const l=scriptArg(),e=scriptArg();return {t:'limlow',e,lim:l};}
      case 'boxed':return {t:'box',e:scriptArg()};
      case 'phantom':case 'hphantom':case 'vphantom':return {t:'phant',e:scriptArg()};
      case 'prescript':{const sup=scriptArg(),sub=scriptArg(),e=scriptArg();return {t:'pre',sub,sup,e};}
      case 'text':case 'textnormal':case 'mbox':case 'hbox':return R(rawGroup().replace(/\\([{}%&#$_\\^~])/g,'$1'),{nor:true});
      case 'operatorname':case 'operatorname*':{const a=atom();const s=(a.t==='g'?a.c:[a]).map(x=>x.s||'').join('');return {t:'fn',name:s};}
      case 'left':{
        const beg=delim();const es=[];let sep='|',cur;
        for(;;){
          cur=seq(t=>t.k==='}'||(t.k==='cmd'&&(t.v==='right'||t.v==='middle'||t.v==='end')));
          es.push(cur);
          const q=next();if(!q)err('\\left tanpa \\right');
          if(q.k==='cmd'&&q.v==='middle'){sep=delim();continue;}
          if(q.k==='cmd'&&q.v==='right'){return {t:'d',beg,end:delim(),sep,es};}
          err('\\left tanpa \\right');
        }
      }
      case 'right':err('\\right tanpa \\left');break;
      case 'middle':err('\\middle di luar \\left...\\right');break;
      case 'begin':return env(rawGroup());
      case 'end':err('\\end tanpa \\begin');break;
      case 'big':case 'Big':case 'bigg':case 'Bigg':case 'bigl':case 'bigr':case 'Bigl':case 'Bigr':case 'biggl':case 'biggr':{const d=delim();return R(d);}
      case 'textcolor':case 'colortext':{const m=optModel(),c=colorHex(rawGroup(),m),x=scriptArg();applyStyle(x,{col:c});return {t:'g',c:x,style:true};}
      case 'color':{const m=optModel(),c=colorHex(rawGroup(),m);return {t:'cset',st:{col:c}};}
      case 'colorbox':case 'highlight':case 'bgcolor':{const m=optModel(),c=colorHex(rawGroup(),m),x=scriptArg();applyStyle(x,{hl:c});return {t:'g',c:x,style:true};}
      case 'size':case 'fontsize':{
        const n=parseFloat(rawGroup().replace(',','.'));
        if(!(n>=4&&n<=200))err('Ukuran harus antara 4 dan 200 poin');
        const x=scriptArg();applyStyle(x,{sz:n});return {t:'g',c:x,style:true};
      }
      case 'xrightarrow':case 'xleftarrow':case 'xlongequal':{            /* panah dengan keterangan di atas/bawah */
        let below=null;const q=peek();
        if(q&&q.k==='ch'&&q.v==='['){next();below=seq(x=>x.k==='ch'&&x.v===']');if(!next())err('Kurung siku [ tidak ditutup');}
        const up=scriptArg();
        let node=R(v==='xrightarrow'?'⟶':v==='xleftarrow'?'⟵':'═');
        if(up.length)node={t:'limupp',e:[node],lim:up};
        if(below&&below.length)node={t:'limlow',e:[node],lim:below};
        return node;
      }
      case 'displaystyle':case 'textstyle':case 'scriptstyle':case 'limits':case 'nolimits':case 'mathstrut':case 'relax':return null;
      case 'not':{const a=atom();return a;}
    }
    err('Perintah tidak dikenal: \\'+v);
  }

  /* akar: beberapa baris dipisah \\ */
  const lines=[];
  for(;;){
    const c=seq(t=>t.k==='cmd'&&t.v==='\\'||t.k==='}'||(t.k==='cmd'&&(t.v==='right'||t.v==='middle'||t.v==='end')));
    lines.push(c);
    const t=peek();if(!t)break;
    if(t.k==='cmd'&&t.v==='\\'){next();continue;}
    if(t.k==='}')err('Kurung } tanpa pasangan');
    err('\\'+t.v+' tanpa pasangan');
  }
  while(lines.length>1&&!lines[lines.length-1].length)lines.pop();
  return lines;
}

/* ============================ AST -> OMML ============================ */
const val=(n,v)=>`<m:${n} m:val="${xa(v)}"/>`;
function flat(nodes){const o=[];for(const n of nodes||[]){if(n.t==='g'){o.push(...flat(n.c));}else if(n.t==='fnop'||n.t==='cset'){}else o.push(n);}return o;}
/* isi <w:rPr>: font + warna + ukuran + sorot (urutan sesuai skema Word) */
function rprXML(st){
  let x='<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>';
  if(st&&st.col)x+=`<w:color w:val="${st.col}"/>`;
  if(st&&st.sz>0){const h=Math.round(st.sz*2);x+=`<w:sz w:val="${h}"/><w:szCs w:val="${h}"/>`;}
  if(st&&st.hl)x+=`<w:shd w:val="clear" w:color="auto" w:fill="${st.hl}"/>`;
  return x;
}
const STY_VAL={p:'p',b:'b',bi:'bi',i:'i'};
function runXML(s,o){
  o=o||{};let pr='';
  if(o.nor)pr='<m:nor/>';else if(o.sty&&STY_VAL[o.sty])pr=val('sty',STY_VAL[o.sty]);
  return `<m:r>${pr?`<m:rPr>${pr}</m:rPr>`:''}<w:rPr>${rprXML(o)}</w:rPr><m:t${/^\s|\s$/.test(s)?' xml:space="preserve"':''}>${xe(s)}</m:t></m:r>`;
}
/* properti struktur: m:xxxPr berisi opsi + m:ctrlPr (warna/ukuran/sorot garis pecahan, kurung, akar, dst.) */
const hasSt=n=>!!(n.col||n.sz>0||n.hl);
function prX(tag,inner,n){
  const c=hasSt(n)?`<m:ctrlPr><w:rPr>${rprXML(n)}</w:rPr></m:ctrlPr>`:'';
  return inner||c?`<m:${tag}>${inner||''}${c}</m:${tag}>`:'';
}
const stk=n=>[n.sty||'',n.nor?1:0,n.col||'',n.sz||'',n.hl||''].join('|');
function emit(nodes,ph){
  const L=flat(nodes);let out='';
  const box=()=>runXML('□');
  const E=x=>{const s=emit(x,ph);return s||(ph?box():'');};
  const arg=(tag,x)=>`<m:${tag}>${E(x)}</m:${tag}>`;
  for(let i=0;i<L.length;i++){
    const n=L[i];
    if(n.t==='r'){
      let s=n.s;
      while(L[i+1]&&L[i+1].t==='r'&&stk(L[i+1])===stk(n)){s+=L[i+1].s;i++;}
      if(s)out+=runXML(s,n);continue;
    }
    switch(n.t){
      case 'f':out+=`<m:f>${prX('fPr',n.type!=='bar'?val('type',n.type):'',n)}${arg('num',n.n)}${arg('den',n.d)}</m:f>`;break;
      case 'rad':{
        const has=n.deg&&n.deg.length,show=has||(ph&&n.deg);          /* akar pangkat n: tampilkan kotak pangkat di pratinjau */
        out+=`<m:rad>${prX('radPr',show?'':val('degHide','1'),n)}<m:deg>${has?emit(n.deg,ph):show?box():''}</m:deg>${arg('e',n.e)}</m:rad>`;break;}
      case 'sup':out+=`<m:sSup>${prX('sSupPr','',n)}${arg('e',n.e)}${arg('sup',n.sup)}</m:sSup>`;break;
      case 'sub':out+=`<m:sSub>${prX('sSubPr','',n)}${arg('e',n.e)}${arg('sub',n.sub)}</m:sSub>`;break;
      case 'ssub':out+=`<m:sSubSup>${prX('sSubSupPr','',n)}${arg('e',n.e)}${arg('sub',n.sub)}${arg('sup',n.sup)}</m:sSubSup>`;break;
      case 'pre':out+=`<m:sPre>${prX('sPrePr','',n)}${arg('sub',n.sub)}${arg('sup',n.sup)}${arg('e',n.e)}</m:sPre>`;break;
      case 'nary':{
        const hs=!(n.sub&&(n.sub.length||ph)),hp=!(n.sup&&(n.sup.length||ph));
        out+=`<m:nary>${prX('naryPr',val('chr',n.chr)+val('limLoc',n.loc)+(hs?val('subHide','1'):'')+(hp?val('supHide','1'):''),n)}<m:sub>${hs?'':E(n.sub)}</m:sub><m:sup>${hp?'':E(n.sup)}</m:sup>${arg('e',n.e)}</m:nary>`;break;}
      case 'd':out+=`<m:d>${prX('dPr',val('begChr',n.beg)+(n.es.length>1?val('sepChr',n.sep):'')+val('endChr',n.end),n)}${n.es.map(x=>`<m:e>${E(x)}</m:e>`).join('')}</m:d>`;break;
      case 'func':out+=`<m:func>${prX('funcPr','',n)}<m:fName>${emit(n.name,ph)}</m:fName>${arg('e',n.e)}</m:func>`;break;
      case 'fn':out+=runXML(n.name,Object.assign({},n,{sty:'p'}));break;
      case 'limlow':out+=`<m:limLow>${prX('limLowPr','',n)}${arg('e',n.e)}${arg('lim',n.lim)}</m:limLow>`;break;
      case 'limupp':out+=`<m:limUpp>${prX('limUppPr','',n)}${arg('e',n.e)}${arg('lim',n.lim)}</m:limUpp>`;break;
      case 'acc':out+=`<m:acc>${prX('accPr',val('chr',n.chr),n)}${arg('e',n.e)}</m:acc>`;break;
      case 'bar':out+=`<m:bar>${prX('barPr',val('pos',n.pos),n)}${arg('e',n.e)}</m:bar>`;break;
      case 'gchr':out+=`<m:groupChr>${prX('groupChrPr',val('chr',n.chr)+val('pos',n.pos),n)}${arg('e',n.e)}</m:groupChr>`;break;
      case 'box':out+=`<m:borderBox>${prX('borderBoxPr','',n)}${arg('e',n.e)}</m:borderBox>`;break;
      case 'phant':out+=`<m:phant>${prX('phantPr',val('show','0'),n)}${arg('e',n.e)}</m:phant>`;break;
      case 'eqarr':out+=`<m:eqArr>${prX('eqArrPr','',n)}${(n.rows.length?n.rows:[[]]).map(r=>`<m:e>${E(r)}</m:e>`).join('')}</m:eqArr>`;break;
      case 'mat':out+=`<m:m>${prX('mPr','',n)}${n.rows.map(r=>`<m:mr>${r.map(c=>`<m:e>${E(c)}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`;break;
    }
  }
  return out;
}
const NSDECL=` xmlns:m="${M_NS}" xmlns:w="${W_NS}"`;
/* opt: {display:bool, ph:bool(kotak kosong utk pratinjau)} -> string OMML */
function tex2omml(tex,opt){
  opt=opt||{};
  const lines=parseTex(tex,opt);
  const ph=!!opt.ph;
  if(lines.length===1&&!lines[0].length)err('Rumus masih kosong');
  const multi=lines.length>1;
  const body=lines.map(l=>`<m:oMath>${emit(l,ph)}</m:oMath>`);
  if(opt.display||multi)return `<m:oMathPara${NSDECL}>${body.join('')}</m:oMathPara>`;
  return body[0].replace('<m:oMath>',`<m:oMath${NSDECL}>`);
}

/* ============================ OMML -> LaTeX ============================ */
const kids=(el,n)=>el?[...el.childNodes].filter(x=>x.nodeType===1&&(!n||x.localName===n)):[];
const kid=(el,n)=>kids(el,n)[0]||null;
const mattr=(el,a)=>{
  if(!el)return null;
  if(el.hasAttributeNS&&el.hasAttributeNS(M_NS,a))return el.getAttributeNS(M_NS,a);
  return el.hasAttribute('m:'+a)?el.getAttribute('m:'+a):null;
};
const on=el=>{if(!el)return false;const v=mattr(el,'val');return v===null||!/^(0|false|off)$/i.test(v);};
const chrVal=(pr,tag,d)=>{const v=mattr(kid(pr,tag),'val');return v===null?d:v;};
const plain=el=>{let s='';if(el)for(const t of el.getElementsByTagNameNS(M_NS,'t'))s+=t.textContent;return s;};
const hasContent=el=>!!(el&&(plain(el).trim()||el.getElementsByTagNameNS(M_NS,'r').length));
/* ---- gaya (warna/ukuran/sorot) dari OMML ---- */
const wkid=(el,n)=>el?([...el.childNodes].find(x=>x.nodeType===1&&x.localName===n&&x.namespaceURI===W_NS)||null):null;
const wval=el=>el?(el.getAttributeNS(W_NS,'val')||el.getAttribute('w:val')||''):'';
const mrp=r=>[...r.childNodes].find(x=>x.nodeType===1&&x.localName==='rPr'&&x.namespaceURI===M_NS)||null;
const HLITE={yellow:'FFFF00',green:'00FF00',cyan:'00FFFF',magenta:'FF00FF',blue:'0000FF',red:'FF0000',darkBlue:'000080',darkCyan:'008080',darkGreen:'008000',darkMagenta:'800080',darkRed:'800000',darkYellow:'808000',darkGray:'808080',lightGray:'C0C0C0'};
function rpOf(el){
  if(el.localName==='r')return wkid(el,'rPr');
  const pr=[...el.childNodes].find(x=>x.nodeType===1&&x.namespaceURI===M_NS&&/Pr$/.test(x.localName)&&x.localName!=='rPr');
  const cp=pr?kid(pr,'ctrlPr'):null;return cp?wkid(cp,'rPr'):null;
}
/* kunci gaya "warna|ukuran|sorot"; '' bila tidak ada. Ukuran sama dengan ukuran teks dasar dianggap bawaan. */
function styKey(el,base){
  const rp=rpOf(el);if(!rp)return '';
  let col='',sz='',hl='';
  const cv=wval(wkid(rp,'color'));if(/^[0-9a-f]{6}$/i.test(cv)&&!/^000000$/.test(cv))col=cv.toUpperCase();
  const zv=+wval(wkid(rp,'sz'))/2;if(zv>0&&!(base&&Math.abs(zv-base)<.3))sz=String(zv);
  const sh=wkid(rp,'shd'),fill=sh?(sh.getAttributeNS(W_NS,'fill')||sh.getAttribute('w:fill')||''):'';
  if(/^[0-9a-f]{6}$/i.test(fill))hl=fill.toUpperCase();else{const h=HLITE[wval(wkid(rp,'highlight'))];if(h)hl=h;}
  return col||sz||hl?col+'|'+sz+'|'+hl:'';
}
function wrapSty(k,inner){
  const [col,sz,hl]=k.split('|');let o=inner;
  if(col)o='\\textcolor{'+col+'}{'+o+'}';
  if(hl)o='\\colorbox{'+hl+'}{'+o+'}';
  if(sz)o='\\size{'+sz+'}{'+o+'}';
  return o;
}
const cat=(a,b)=>(a&&b&&/\\[A-Za-z]+$/.test(a)&&/^[A-Za-z]/.test(b))?a+' '+b:a+b;
const BREAKCHARS=new Set([...BREAKCH,'\u2212',...[...BREAKCMD].map(c=>SYM[c]).filter(Boolean)]);
const SCRCMD={'double-struck':'mathbb',script:'mathcal',fraktur:'mathfrak','sans-serif':'mathsf',monospace:'mathtt'};
const MATREV={'()':'pmatrix','[]':'bmatrix','{}':'Bmatrix','||':'vmatrix','‖‖':'Vmatrix'};
const parseSafe=t=>{try{return parseTex(t);}catch(e){return null;}};
const NOBASE=new Set(['sup','sub','ssub','nary','func','limlow','limupp','pre']);
const atomic=t=>{const L=parseSafe(t);return !!L&&L.length===1&&L[0].length===1&&!NOBASE.has(L[0][0].t);};
const single=t=>{const L=parseSafe(t);return !!L&&L.length===1&&L[0].length===1;};
const bs=t=>atomic(t)?t:'{'+t+'}';
const sc=t=>/^[A-Za-z0-9]$/.test(t)?t:'{'+t+'}';
const needBrace=t=>{
  const L=parseSafe(t);if(!L||L.length!==1)return true;
  const N=L[0];if(N.length<=1)return false;
  return N[0].t==='g'||N.some(x=>x.t==='r'&&x.s.length===1&&BREAKCHARS.has(x.s));
};
const parenWrapped=t=>{
  if(t[0]!=='('||t[t.length-1]!==')')return false;
  let d=0;for(let i=0;i<t.length;i++){if(t[i]==='(')d++;else if(t[i]===')'){d--;if(d===0&&i<t.length-1)return false;}}
  return d===0;
};
function chTex(c,inEq){
  if(c==='\u2212')return '-';
  if(c==='‖')return '\\|';
  if(SYM_REV[c])return '\\'+SYM_REV[c];
  switch(c){
    case '{':return '\\{';case '}':return '\\}';case '\\':return '\\backslash ';case '$':return '\\$';case '%':return '\\%';case '#':return '\\#';
    case '&':return inEq?'&':'\\&';case '_':return '\\_';case '^':return '\\^{}';case '~':return '\\~{}';case '\'':return '\\prime';
    case '\u2009':return '\\,';case '\u2005':return '\\:';case '\u2004':return '\\;';case '\u2003':return '\\quad';case '\u2002':return '\\enspace';case '\u00A0':return '~';
  }
  return c;
}
function runTex(r,c){
  const rPr=mrp(r);let t='';
  for(const x of r.childNodes)if(x.nodeType===1&&x.localName==='t')t+=x.textContent;
  if(!t)return '';
  if(rPr&&on(kid(rPr,'nor')))return '\\text{'+t.replace(/[\\{}$&%#_^~]/g,m=>'\\'+m)+'}';
  const sty0=rPr?mattr(kid(rPr,'sty'),'val'):null,sty=(c.uni&&sty0===c.uni)?null:sty0,scr=rPr?mattr(kid(rPr,'scr'),'val'):null;
  const styCmd=scr&&SCRCMD[scr]?SCRCMD[scr]:(sty==='p'?'mathrm':sty==='b'?'mathbf':sty==='bi'?'boldsymbol':(sty==='i'&&/[^\p{L}\s]/u.test(t))?'mathit':null);
  let out='',buf='',bk=null;
  const flush=()=>{if(buf)out=cat(out,bk?'\\'+bk+'{'+buf+'}':buf);buf='';bk=null;};
  for(const ch of t){
    let kind=null,x=ch;
    if(ALPHA_REV[ch]){kind=ALPHA_CMD[ALPHA_REV[ch][0]];x=ALPHA_REV[ch][1];}
    else if(/[A-Za-z0-9]/.test(ch))kind=styCmd;
    else{flush();out=cat(out,chTex(ch,c.inEq));continue;}
    if(kind!==bk)flush();
    bk=kind;buf+=x;
  }
  flush();return out;
}
function oseq(el,c){
  let s='';const ns=kids(el);
  for(let i=0;i<ns.length;){
    const k=styKey(ns[i],c.base);
    if(k&&k!==c.cur){                              /* kelompokkan simpul berurutan yang bergaya sama -> satu pembungkus */
      const prev=c.cur;c.cur=k;let inner='',j=i;
      while(j<ns.length&&styKey(ns[j],c.base)===k){inner=cat(inner,onode(ns[j],c));j++;}
      c.cur=prev;s=cat(s,wrapSty(k,inner));i=j;
    }else{s=cat(s,onode(ns[i],c));i++;}
  }
  return s;
}
const oarg=(p,name,c)=>oseq(kid(p,name),c);
const dtex=ch=>ch===''?'.':(DELIM_REV[ch]!==undefined?DELIM_REV[ch]:ch);
function fnameTex(el,c){
  const ch=kids(el).filter(x=>!/Pr$/.test(x.localName));
  const nm=e=>{const t=plain(e).trim();return FUNCS.has(t)?t:null;};
  if(ch.length===1){
    const x=ch[0],ln=x.localName;
    if(ln==='r'){const t=plain(x).trim();if(FUNCS.has(t))return '\\'+t;}
    if((ln==='limLow'||ln==='limUpp')&&nm(kid(x,'e')))return '\\'+nm(kid(x,'e'))+(ln==='limLow'?'_':'^')+sc(oarg(x,'lim',c));
    if(ln==='sSup'&&nm(kid(x,'e')))return '\\'+nm(kid(x,'e'))+'^'+sc(oarg(x,'sup',c));
    if(ln==='sSub'&&nm(kid(x,'e')))return '\\'+nm(kid(x,'e'))+'_'+sc(oarg(x,'sub',c));
    if(ln==='sSubSup'&&nm(kid(x,'e')))return '\\'+nm(kid(x,'e'))+'_'+sc(oarg(x,'sub',c))+'^'+sc(oarg(x,'sup',c));
  }
  if(ch.every(x=>x.localName==='r')){const t=plain(el).trim();if(t)return FUNCS.has(t)?'\\'+t:'\\operatorname{'+t+'}';}
  return oseq(el,c);
}
function matTex(m,name,c){
  const rows=kids(m,'mr').map(r=>kids(r,'e').map(e=>oseq(e,c)).join(' & ')).join(' \\\\ ');
  return '\\begin{'+name+'} '+rows+' \\end{'+name+'}';
}
function arrTex(a,name,c){
  const k={inEq:true,lossy:false,cur:c.cur,base:c.base};const rows=kids(a,'e').map(e=>oseq(e,k)).join(' \\\\ ');
  if(k.lossy)c.lossy=true;
  return '\\begin{'+name+'} '+rows+' \\end{'+name+'}';
}
const CONT=new Set(['e','num','den','sub','sup','lim','deg','fName','mr','oMath']);
function onode(n,c){
  const ln=n.localName;
  switch(ln){
    case 'r':return runTex(n,c);
    case 'f':{
      const t=chrVal(kid(n,'fPr'),'type','bar'),a=oarg(n,'num',c),b=oarg(n,'den',c);
      if(t==='skw')return '\\sfrac{'+a+'}{'+b+'}';
      if(t==='lin')return '\\nicefrac{'+a+'}{'+b+'}';
      if(t==='noBar')return '\\genfrac{}{}{0pt}{}{'+a+'}{'+b+'}';
      return '\\frac{'+a+'}{'+b+'}';
    }
    case 'rad':{
      const hide=on(kid(kid(n,'radPr'),'degHide'));
      return !hide&&hasContent(kid(n,'deg'))?'\\sqrt['+oarg(n,'deg',c)+']{'+oarg(n,'e',c)+'}':'\\sqrt{'+oarg(n,'e',c)+'}';
    }
    case 'sSup':case 'sSub':case 'sSubSup':{
      const e=oarg(n,'e',c);const base=bs(e);
      const sub=ln!=='sSup'?oarg(n,'sub',c):null,sup=ln!=='sSub'?oarg(n,'sup',c):null;
      let s=base;
      if(sub!==null)s+='_'+sc(sub);
      if(sup!==null)s+=(/^(\\prime)+$/.test(sup)&&sub===null&&base===e)?"'".repeat(sup.split('\\prime').length-1):'^'+sc(sup);
      return s;
    }
    case 'sPre':return '\\prescript{'+oarg(n,'sup',c)+'}{'+oarg(n,'sub',c)+'}{'+oarg(n,'e',c)+'}';
    case 'nary':{
      const pr=kid(n,'naryPr'),chr=chrVal(pr,'chr','\u222B');let s;
      if(NARY_REV[chr])s='\\'+NARY_REV[chr];else{s='\\sum';c.lossy=true;}
      const dflt=/[\u222B-\u2233]/.test(chr)?'subSup':'undOvr',loc=chrVal(pr,'limLoc',dflt);
      if(loc!==dflt)s+=loc==='undOvr'?'\\limits':'\\nolimits';
      const hs=on(kid(pr,'subHide'))||!hasContent(kid(n,'sub')),hp=on(kid(pr,'supHide'))||!hasContent(kid(n,'sup'));
      if(!hs)s+='_'+sc(oarg(n,'sub',c));
      if(!hp)s+='^'+sc(oarg(n,'sup',c));
      let e=oarg(n,'e',c);
      if(e){if(needBrace(e))e='{'+e+'}';s+=' '+e;}
      return s;
    }
    case 'd':{
      const pr=kid(n,'dPr'),b=chrVal(pr,'begChr','('),e=chrVal(pr,'endChr',')'),sp=chrVal(pr,'sepChr','|'),es=kids(n,'e');
      if(es.length===1){
        const ch=kids(es[0]);
        if(ch.length===1){
          const x=ch[0];
          if(b==='('&&e===')'&&x.localName==='f'&&chrVal(kid(x,'fPr'),'type','bar')==='noBar')return '\\binom{'+oarg(x,'num',c)+'}{'+oarg(x,'den',c)+'}';
          if(x.localName==='m'&&MATREV[b+e])return matTex(x,MATREV[b+e],c);
          if(x.localName==='eqArr'&&b==='{'&&e==='')return arrTex(x,'cases',c);
        }
      }
      return cat('\\left'+dtex(b),es.map(x=>oseq(x,c)).join(' \\middle'+dtex(sp)+' '))+'\\right'+dtex(e);
    }
    case 'func':{
      const nm=fnameTex(kid(n,'fName'),c),e=oarg(n,'e',c);
      if(!e)return nm;
      return (single(e)||parenWrapped(e))?nm+' '+e:nm+'{'+e+'}';
    }
    case 'limLow':case 'limUpp':{
      const low=ln==='limLow',e=kid(n,'e'),eks=kids(e);
      if(eks.length===1&&eks[0].localName==='groupChr'){
        const g=eks[0],gp=kid(g,'groupChrPr'),ch=chrVal(gp,'chr','⏟'),top=chrVal(gp,'pos','bot')==='top';
        if((ch==='⏟'&&!top&&low)||(ch==='⏞'&&top&&!low))return '\\'+(low?'underbrace':'overbrace')+'{'+oarg(g,'e',c)+'}'+(low?'_':'^')+sc(oarg(n,'lim',c));
      }
      return '\\'+(low?'underset':'overset')+'{'+oarg(n,'lim',c)+'}{'+oarg(n,'e',c)+'}';
    }
    case 'acc':{
      const ch=chrVal(kid(n,'accPr'),'chr','\u0302');
      if(ACC_REV[ch])return '\\'+ACC_REV[ch]+'{'+oarg(n,'e',c)+'}';
      c.lossy=true;return '\\overset{'+ch+'}{'+oarg(n,'e',c)+'}';
    }
    case 'bar':return (chrVal(kid(n,'barPr'),'pos','top')==='bot'?'\\underline{':'\\overline{')+oarg(n,'e',c)+'}';
    case 'groupChr':{
      const pr=kid(n,'groupChrPr'),ch=chrVal(pr,'chr','⏟'),top=chrVal(pr,'pos','bot')==='top';
      if(ch==='⏟'&&!top)return '\\underbrace{'+oarg(n,'e',c)+'}';
      if(ch==='⏞'&&top)return '\\overbrace{'+oarg(n,'e',c)+'}';
      c.lossy=true;return '\\'+(top?'overset':'underset')+'{'+ch+'}{'+oarg(n,'e',c)+'}';
    }
    case 'box':c.lossy=true;return oarg(n,'e',c);
    case 'phant':{const sh=kid(kid(n,'phantPr'),'show');return sh&&!on(sh)?'\\phantom{'+oarg(n,'e',c)+'}':oarg(n,'e',c);}
    case 'borderBox':return '\\boxed{'+oarg(n,'e',c)+'}';
    case 'eqArr':return arrTex(n,'aligned',c);
    case 'm':return matTex(n,'matrix',c);
    default:
      if(/Pr$/.test(ln))return '';
      if(!CONT.has(ln))c.lossy=true;
      return oseq(n,c);
  }
}
/* el: elemen m:oMath / m:oMathPara -> {tex, lossy, display} */
function omml2tex(el,opt){
  const c={lossy:false,inEq:false,cur:'',base:(opt&&opt.basePt)||0};let tex,display=false;
  /* semua huruf/angka bergaya sama (mis. seluruh rumus tebal) -> satu pembungkus, bukan satu per huruf */
  const vs=[...el.getElementsByTagNameNS(M_NS,'r')].filter(r=>/[\p{L}\p{N}]/u.test(plain(r))&&!(mrp(r)&&on(kid(mrp(r),'nor')))).map(r=>{const p=mrp(r);return (p&&mattr(kid(p,'sty'),'val'))||'';});
  if(vs.length>1&&vs[0]&&vs.every(v=>v===vs[0])&&{p:1,b:1,bi:1}[vs[0]])c.uni=vs[0];
  const UNI={p:'mathrm',b:'mathbf',bi:'boldsymbol'};
  const wrapU=t=>c.uni&&t?'\\'+UNI[c.uni]+'{'+t+'}':t;
  if(el.localName==='oMathPara'){display=true;tex=kids(el,'oMath').map(m=>wrapU(oseq(m,c))).join(' \\\\ ');}
  else tex=wrapU(oseq(el,c));
  return {tex:tex.trim(),lossy:c.lossy,display};
}
/* ubah OMML sebaris <-> blok tanpa menyentuh isinya; null bila tidak mungkin (beberapa baris -> sebaris) */
function setDisplay(omml,display,DP,XS){
  DP=DP||root.DOMParser;XS=XS||root.XMLSerializer;
  const el=new DP().parseFromString(omml,'application/xml').documentElement;
  const ser=n=>new XS().serializeToString(n);
  const para=el.localName==='oMathPara';
  if(display)return para?omml:`<m:oMathPara${NSDECL}>${ser(el).replace(/^<\?xml[^>]*\?>\s*/,'')}</m:oMathPara>`;
  if(!para)return omml;
  const ms=kids(el,'oMath');if(ms.length!==1)return null;
  return ser(ms[0]);
}
const isPara=o=>/^\s*<m:oMathPara[\s>]/.test(o);

/* ============================ PALET & DIALOG ============================ */
const CM='‸';   /* penanda posisi kursor di templat */
/* item palet: [label, templat, keterangan, (glyph khusus untuk tombol)] */
const sy=n=>[SYM[n],'\\'+n+' ',n];
const T=(l,t,h,g)=>[l,t,h,g];
const mx=(env,rows,cols,first)=>{                          /* matriks kosong rows×cols; kursor di sel pertama */
  const r=[];for(let i=0;i<rows;i++){const c=[];for(let j=0;j<cols;j++)c.push(i||j?'':first);r.push(c.join(' & '));}
  return '\\begin{'+env+'} '+r.join(' \\\\ ')+' \\end{'+env+'}';
};
const PAL={
  d:['Dasar',[
    T('a/b','\\frac{‸}{}','Pecahan'),
    T('a⁄b','\\sfrac{‸}{}','Pecahan miring'),
    T('√','\\sqrt{‸}','Akar kuadrat'),
    T('ⁿ√','\\sqrt[‸]{}','Akar pangkat n (ganti n)','\\sqrt[n]{\\square}'),
    T('∛','\\sqrt[3]{‸}','Akar pangkat 3'),
    T('xⁿ','^{‸}','Pangkat'),
    T('x²','^{2}','Kuadrat'),
    T('xₙ','_{‸}','Indeks bawah'),
    T('xₙⁿ','_{‸}^{}','Indeks + pangkat'),
    T('ⁿₘx','\\prescript{‸}{}{}','Indeks di kiri'),
    T('eˣ','e^{‸}','Eksponen e'),
    T('( )','\\left(‸\\right)','Kurung biasa'),
    T('[ ]','\\left[‸\\right]','Kurung siku'),
    T('{ }','\\left\\{‸\\right\\}','Kurung kurawal'),
    T('| |','\\left|‸\\right|','Nilai mutlak'),
    T('‖ ‖','\\left\\|‸\\right\\|','Norma'),
    T('⟨ ⟩','\\left\\langle ‸\\right\\rangle','Kurung sudut'),
    T('⌊ ⌋','\\left\\lfloor ‸\\right\\rfloor','Pembulatan bawah'),
    T('⌈ ⌉','\\left\\lceil ‸\\right\\rceil','Pembulatan atas'),
    T('(n k)','\\binom{‸}{}','Kombinasi / binomial'),
    T('n!','‸!','Faktorial'),
    T('log','\\log_{‸}{}','Logaritma basis'),
    T('ln','\\ln\\left(‸\\right)','Logaritma natural'),
    T('±','\\pm ','Plus-minus'),
    T('×','\\times ','Kali')
  ],1],
  k:['Kalkulus',[
    T('∫','\\int ‸\\,dx','Integral tak tentu'),
    T('∫ₐᵇ','\\int_{‸}^{}{}\\,dx','Integral tentu'),
    T('∬','\\iint_{‸}{}\\,dA','Integral ganda'),
    T('∭','\\iiint_{‸}{}\\,dV','Integral tripel'),
    T('∮','\\oint_{‸}{}\\,ds','Integral kontur'),
    T('[ ]ₐᵇ','\\left[‸\\right]_{}^{}','Substitusi batas'),
    T('lim','\\lim_{x\\to ‸}{}','Limit'),
    T('lim ∞','\\lim_{x\\to\\infty}{‸}','Limit tak hingga'),
    T('d/dx','\\frac{d}{dx}\\left(‸\\right)','Turunan'),
    T('dⁿ/dxⁿ','\\frac{d^{n}}{dx^{n}}','Turunan ke-n'),
    T('∂/∂x','\\frac{\\partial ‸}{\\partial x}','Turunan parsial'),
    T('∂²/∂x²','\\frac{\\partial^{2} ‸}{\\partial x^{2}}','Parsial kedua'),
    T('f′','f\'(‸)','Turunan (aksen)'),
    T('∑','\\sum_{‸}^{}{}','Sigma (jumlah)'),
    T('∏','\\prod_{‸}^{}{}','Perkalian'),
    T('⋃','\\bigcup_{‸}^{}{}','Gabungan besar'),
    T('⋂','\\bigcap_{‸}^{}{}','Irisan besar'),
    T('∇','\\nabla ‸','Nabla / gradien'),
    T('Δ','\\Delta ‸','Delta (perubahan)'),
    T('∞','\\infty ','Tak hingga')
  ],1],
  a:['Aljabar',[
    T('2×2',mx('pmatrix',2,2,CM),'Matriks 2×2'),
    T('3×3',mx('pmatrix',3,3,CM),'Matriks 3×3'),
    T('[2×2]',mx('bmatrix',2,2,CM),'Matriks kurung siku'),
    T('|2×2|',mx('vmatrix',2,2,CM),'Determinan 2×2'),
    T('|3×3|',mx('vmatrix',3,3,CM),'Determinan 3×3'),
    T('2×1',mx('pmatrix',2,1,CM),'Vektor kolom 2'),
    T('3×1',mx('pmatrix',3,1,CM),'Vektor kolom 3'),
    T('1×3',mx('pmatrix',1,3,CM),'Vektor baris 3'),
    T('I₂','\\begin{pmatrix} 1 & 0 \\\\ 0 & 1 \\end{pmatrix}','Matriks identitas'),
    T('{ 2','\\begin{cases} ‸ \\\\ \\end{cases}','Sistem persamaan 2 baris'),
    T('{ 3','\\begin{cases} ‸ \\\\ \\\\ \\end{cases}','Sistem persamaan 3 baris'),
    T('f(x)={','f(x)=\\begin{cases} ‸ & \\text{jika } \\\\ & \\text{jika } \\end{cases}','Fungsi bersyarat'),
    T('rata =','\\begin{aligned} ‸ &= \\\\ &= \\end{aligned}','Langkah sejajar tanda ='),
    T('ax²','ax^{2}+bx+c=0','Persamaan kuadrat'),
    T('ABC','x=\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}','Rumus ABC'),
    T('(a+b)²','(a+b)^{2}=a^{2}+2ab+b^{2}','Kuadrat jumlah'),
    T('a²−b²','a^{2}-b^{2}=(a+b)(a-b)','Selisih kuadrat'),
    T('ₙCᵣ','\\prescript{}{n}{C_{r}}','Kombinasi nCr'),
    T('ₙPᵣ','\\prescript{}{n}{P_{r}}','Permutasi nPr')
  ],1],
  t:['Trigonometri',[
    T('sin','\\sin\\left(‸\\right)','Sinus'),T('cos','\\cos\\left(‸\\right)','Cosinus'),T('tan','\\tan\\left(‸\\right)','Tangen'),
    T('cot','\\cot\\left(‸\\right)','Kotangen'),T('sec','\\sec\\left(‸\\right)','Sekan'),T('csc','\\csc\\left(‸\\right)','Kosekan'),
    T('arcsin','\\arcsin\\left(‸\\right)','Arcsin'),T('arccos','\\arccos\\left(‸\\right)','Arccos'),T('arctan','\\arctan\\left(‸\\right)','Arctan'),
    T('sinh','\\sinh\\left(‸\\right)','Sinus hiperbolik'),T('cosh','\\cosh\\left(‸\\right)','Cosinus hiperbolik'),T('tanh','\\tanh\\left(‸\\right)','Tangen hiperbolik'),
    T('sin²','\\sin^{2}‸','Sinus kuadrat'),T('cos²','\\cos^{2}‸','Cosinus kuadrat'),
    T('sin²+cos²','\\sin^{2}\\theta+\\cos^{2}\\theta=1','Identitas dasar'),
    T('sin(α±β)','\\sin(\\alpha\\pm\\beta)=\\sin\\alpha\\cos\\beta\\pm\\cos\\alpha\\sin\\beta','Jumlah sudut'),
    T('tan θ','\\tan\\theta=\\frac{\\sin\\theta}{\\cos\\theta}','Tan = sin/cos'),
    T('sin 2θ','\\sin 2\\theta=2\\sin\\theta\\cos\\theta','Sudut ganda'),
    T('a/sin A','\\frac{a}{\\sin A}=\\frac{b}{\\sin B}=\\frac{c}{\\sin C}','Aturan sinus'),
    T('c²=…','c^{2}=a^{2}+b^{2}-2ab\\cos C','Aturan cosinus'),
    T('°','^{\\circ}','Derajat'),
    T('∠','\\angle ','Sudut'),
    T('π','\\pi ','Pi'),
    T('θ','\\theta ','Theta')
  ],1],
  x:['Aksen',[
    T('x̂','\\hat{‸}','Topi'),T('x̄','\\overline{‸}','Garis atas'),T('x̲','\\underline{‸}','Garis bawah'),
    T('x⃗','\\vec{‸}','Vektor'),T('ẋ','\\dot{‸}','Titik'),T('ẍ','\\ddot{‸}','Dua titik'),
    T('x̃','\\tilde{‸}','Tilde'),T('x̆','\\breve{‸}','Breve'),T('x̌','\\check{‸}','Caron'),
    T('x́','\\acute{‸}','Akut'),T('x̀','\\grave{‸}','Grave'),
    T('AB→','\\overrightarrow{‸}','Panah di atas'),
    T('⏞','\\overbrace{‸}^{}','Kurung kurawal atas'),T('⏟','\\underbrace{‸}_{}','Kurung kurawal bawah'),
    T('atas','\\overset{‸}{}','Simbol di atas'),T('bawah','\\underset{‸}{}','Simbol di bawah'),
    T('→ ket.','\\xrightarrow{‸}','Panah + keterangan atas'),
    T('→ ket.2','\\xrightarrow[]{‸}','Panah + keterangan atas/bawah'),
    T('▭','\\boxed{‸}','Kotak')
  ],1],
  y:['Yunani',('alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi pi varpi rho varrho sigma varsigma tau upsilon phi varphi chi psi omega Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega').split(' ').map(sy),0],
  o:['Operator',('pm mp times div cdot ast star circ bullet oplus ominus otimes odot wedge vee cap cup uplus sqcup sqcap setminus neg partial nabla infty prime degree dagger ddagger').split(' ').map(sy),0],
  r:['Relasi',[['=','=','sama dengan'],['<','<','kurang dari'],['>','>','lebih dari']].concat(('neq approx equiv sim simeq cong propto leq geq ll gg prec succ doteq in notin ni subset supset subseteq supseteq mid nmid perp parallel').split(' ').map(sy)),0],
  p:['Panah',('to leftarrow leftrightarrow Rightarrow Leftarrow Leftrightarrow mapsto uparrow downarrow updownarrow implies impliedby iff longrightarrow longleftarrow longleftrightarrow hookrightarrow rightleftharpoons nearrow searrow swarrow nwarrow').split(' ').map(sy),0],
  h:['Himpunan',[
    T('∈','\\in ','Anggota'),T('∉','\\notin ','Bukan anggota'),T('⊂','\\subset ','Himpunan bagian'),T('⊆','\\subseteq ','Bagian/sama'),
    T('∪','\\cup ','Gabungan'),T('∩','\\cap ','Irisan'),T('∅','\\emptyset ','Himpunan kosong'),T('∖','\\setminus ','Selisih'),
    T('∀','\\forall ','Untuk semua'),T('∃','\\exists ','Ada'),T('¬','\\neg ','Negasi'),T('∧','\\wedge ','Dan'),T('∨','\\vee ','Atau'),
    T('⇒','\\Rightarrow ','Maka'),T('⇔','\\Leftrightarrow ','Jika dan hanya jika'),T('∴','\\therefore ','Jadi'),T('∵','\\because ','Karena'),
    T('ℝ','\\mathbb{R}','Bilangan real'),T('ℕ','\\mathbb{N}','Bilangan asli'),T('ℤ','\\mathbb{Z}','Bilangan bulat'),T('ℚ','\\mathbb{Q}','Bilangan rasional'),T('ℂ','\\mathbb{C}','Bilangan kompleks'),
    T('{x|…}','\\{x\\mid ‸\\}','Himpunan {x | syarat}'),T('{ }','\\{‸\\}','Himpunan'),
    T('A∪B','A\\cup B','Gabungan A dan B'),T('A∩B','A\\cap B','Irisan A dan B'),T('Aᶜ','A^{c}','Komplemen'),T('A×B','A\\times B','Hasil kali Cartesius'),T('n(A)','n(A)','Banyak anggota')
  ],1],
  z:['Statistik',[
    T('x̄','\\bar{‸}','Rata-rata'),
    T('Σx/n','\\bar{x}=\\frac{1}{n}\\sum_{i=1}^{n}x_{i}','Rata-rata hitung'),
    T('s²','s^{2}=\\frac{1}{n-1}\\sum_{i=1}^{n}(x_{i}-\\bar{x})^{2}','Variansi sampel'),
    T('σ','\\sigma=\\sqrt{\\frac{1}{n}\\sum_{i=1}^{n}(x_{i}-\\mu)^{2}}','Simpangan baku'),
    T('P(A)','P(‸)','Peluang'),
    T('P(A|B)','P(A\\mid B)=\\frac{P(A\\cap B)}{P(B)}','Peluang bersyarat'),
    T('E[X]','\\mathbb{E}\\left[‸\\right]','Nilai harapan'),
    T('Var','\\mathrm{Var}(‸)','Variansi'),
    T('N(μ,σ²)','X\\sim N(\\mu,\\sigma^{2})','Distribusi normal'),
    T('z','z=\\frac{x-\\mu}{\\sigma}','Nilai z'),
    T('P(X=k)','P(X=k)=\\binom{n}{k}p^{k}(1-p)^{n-k}','Peluang binomial'),
    T('r','r=\\frac{\\sum(x-\\bar{x})(y-\\bar{y})}{\\sqrt{\\sum(x-\\bar{x})^{2}\\sum(y-\\bar{y})^{2}}}','Korelasi'),
    T('ₙCᵣ','\\prescript{}{n}{C_{r}}','Kombinasi'),T('(n k)','\\binom{n}{k}','Binomial'),T('n!','n!','Faktorial')
  ],1],
  f:['Fisika & Geometri',[
    T('F=ma','\\vec{F}=m\\vec{a}','Hukum II Newton'),
    T('E=mc²','E=mc^{2}','Energi massa'),
    T('Ek','E_{k}=\\frac{1}{2}mv^{2}','Energi kinetik'),
    T('F=G…','F=G\\frac{m_{1}m_{2}}{r^{2}}','Gravitasi Newton'),
    T('V=IR','V=IR','Hukum Ohm'),
    T('v=Δx/Δt','v=\\frac{\\Delta x}{\\Delta t}','Kecepatan'),
    T('m/s²','\\mathrm{m/s^{2}}','Satuan percepatan'),
    T('ħ','\\hbar ','h-bar'),
    T('△ABC','\\triangle ABC','Segitiga'),
    T('AB̅','\\overline{AB}','Ruas garis'),
    T('∠ABC','\\angle ABC','Sudut ABC'),
    T('∥','AB\\parallel CD','Sejajar'),
    T('⊥','AB\\perp CD','Tegak lurus'),
    T('≅','\\triangle ABC\\cong\\triangle DEF','Kongruen'),
    T('πr²','L=\\pi r^{2}','Luas lingkaran'),
    T('4/3πr³','V=\\frac{4}{3}\\pi r^{3}','Volume bola'),
    T('c²=a²+b²','c^{2}=a^{2}+b^{2}','Pythagoras'),
    T('jarak','d=\\sqrt{(x_{2}-x_{1})^{2}+(y_{2}-y_{1})^{2}}','Jarak dua titik')
  ],1],
  c:['Kimia',[
    T('H₂O','\\mathrm{H_{2}O}','Air'),
    T('H₂SO₄','\\mathrm{H_{2}SO_{4}}','Asam sulfat'),
    T('CO₂','\\mathrm{CO_{2}}','Karbon dioksida'),
    T('Na⁺','\\mathrm{Na^{+}}','Ion positif'),
    T('SO₄²⁻','\\mathrm{SO_{4}^{2-}}','Ion bermuatan'),
    T('→','\\rightarrow ','Panah reaksi'),
    T('⇌','\\rightleftharpoons ','Kesetimbangan'),
    T('→Δ','\\xrightarrow{\\Delta}','Reaksi dipanaskan'),
    T('↑','\\uparrow ','Gas naik'),T('↓','\\downarrow ','Endapan'),
    T('ΔH','\\Delta H','Perubahan entalpi'),
    T('¹⁴₆C','\\prescript{14}{6}{\\mathrm{C}}','Isotop'),
    T('pH','\\mathrm{pH}=-\\log\\left[\\mathrm{H^{+}}\\right]','Rumus pH'),
    T('reaksi','2\\mathrm{H_{2}}+\\mathrm{O_{2}}\\rightarrow 2\\mathrm{H_{2}O}','Contoh reaksi')
  ],1],
  l:['Lainnya',[
    T('teks','\\text{‸}','Teks biasa','\\text{abc}'),
    T('abc','\\mathrm{‸}','Huruf tegak'),
    T('𝐱','\\mathbf{‸}','Huruf tebal'),
    T('𝑥','\\mathit{‸}','Huruf miring'),
    T('𝒜','\\mathcal{‸}','Kaligrafi'),
    T('𝔄','\\mathfrak{‸}','Fraktur'),
    T('𝖠','\\mathsf{‸}','Tanpa kait'),
    T('𝚊','\\mathtt{‸}','Mesin tik'),
    T('…','\\ldots ','Titik tiga'),T('⋯','\\cdots ','Titik tiga tengah'),T('⋮','\\vdots ','Titik vertikal'),T('⋱','\\ddots ','Titik diagonal'),
    T('ℏ','\\hbar ','h-bar'),T('ℓ','\\ell ','l kecil'),T('ℵ','\\aleph ','Aleph'),T('ℜ','\\Re ','Bagian real'),T('ℑ','\\Im ','Bagian imajiner'),
    T('spasi','\\quad ','Spasi lebar'),T('spasi²','\\qquad ','Spasi sangat lebar'),T(',','\\, ','Spasi tipis')
  ],1]
};
/* rumus bawaan: [id, nama, LaTeX, kelompok] (id dipakai juga oleh dropdown "Rumus Umum" di pita) */
const GALLERY=[
  ['circ','Luas lingkaran','A=\\pi r^{2}','Geometri'],
  ['kell','Keliling lingkaran','K=2\\pi r','Geometri'],
  ['bola','Volume bola','V=\\frac{4}{3}\\pi r^{3}','Geometri'],
  ['tabung','Volume tabung','V=\\pi r^{2}t','Geometri'],
  ['kerucut','Volume kerucut','V=\\frac{1}{3}\\pi r^{2}t','Geometri'],
  ['segitiga','Luas segitiga','L=\\frac{1}{2}\\times a\\times t','Geometri'],
  ['heron','Rumus Heron','L=\\sqrt{s(s-a)(s-b)(s-c)}','Geometri'],
  ['pyth','Teorema Pythagoras','a^{2}+b^{2}=c^{2}','Geometri'],
  ['jarak','Jarak dua titik','d=\\sqrt{(x_{2}-x_{1})^{2}+(y_{2}-y_{1})^{2}}','Geometri'],
  ['garis','Persamaan garis lurus','y-y_{1}=m(x-x_{1})','Geometri'],
  ['abc','Rumus kuadrat (ABC)','x=\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}','Aljabar'],
  ['kuad','Kuadrat jumlah','(a+b)^{2}=a^{2}+2ab+b^{2}','Aljabar'],
  ['selisih','Selisih kuadrat','a^{2}-b^{2}=(a+b)(a-b)','Aljabar'],
  ['binom','Teorema binomial','(x+a)^{n}=\\sum_{k=0}^{n}\\binom{n}{k}x^{k}a^{n-k}','Aljabar'],
  ['deretar','Deret aritmetika','S_{n}=\\frac{n}{2}\\left(2a+(n-1)b\\right)','Aljabar'],
  ['deretge','Deret geometri','S_{n}=\\frac{a\\left(1-r^{n}\\right)}{1-r}','Aljabar'],
  ['logs','Sifat logaritma','\\log_{a}(bc)=\\log_{a}b+\\log_{a}c','Aljabar'],
  ['eksp','Sifat eksponen','a^{m}\\cdot a^{n}=a^{m+n}','Aljabar'],
  ['det2','Determinan 2×2','\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}=ad-bc','Aljabar'],
  ['inv','Invers matriks 2×2','A^{-1}=\\frac{1}{ad-bc}\\begin{pmatrix} d & -b \\\\ -c & a \\end{pmatrix}','Aljabar'],
  ['sistem','Sistem persamaan','\\begin{cases} 2x+y=5 \\\\ x-y=1 \\end{cases}','Aljabar'],
  ['euler','Identitas Euler','e^{i\\pi}+1=0','Aljabar'],
  ['trig','Identitas trigonometri','\\sin\\alpha\\pm\\sin\\beta=2\\sin\\frac{1}{2}(\\alpha\\pm\\beta)\\cos\\frac{1}{2}(\\alpha\\mp\\beta)','Trigonometri'],
  ['sin2','Sin² + Cos² = 1','\\sin^{2}\\theta+\\cos^{2}\\theta=1','Trigonometri'],
  ['singanda','Sinus sudut ganda','\\sin 2\\theta=2\\sin\\theta\\cos\\theta','Trigonometri'],
  ['arsin','Aturan sinus','\\frac{a}{\\sin A}=\\frac{b}{\\sin B}=\\frac{c}{\\sin C}','Trigonometri'],
  ['arcos','Aturan cosinus','c^{2}=a^{2}+b^{2}-2ab\\cos C','Trigonometri'],
  ['deriv','Definisi turunan',"f'(x)=\\lim_{h\\to 0}\\frac{f(x+h)-f(x)}{h}",'Kalkulus'],
  ['integ','Teorema dasar kalkulus','\\int_{a}^{b}f(x)\\,dx=F(b)-F(a)','Kalkulus'],
  ['rantai','Aturan rantai','\\frac{dy}{dx}=\\frac{dy}{du}\\cdot\\frac{du}{dx}','Kalkulus'],
  ['parsial','Integral parsial','\\int u\\,dv=uv-\\int v\\,du','Kalkulus'],
  ['limsin','Limit sin x / x','\\lim_{x\\to 0}\\frac{\\sin x}{x}=1','Kalkulus'],
  ['gauss','Integral Gauss','\\int_{-\\infty}^{\\infty}e^{-x^{2}}\\,dx=\\sqrt{\\pi}','Kalkulus'],
  ['jumlah','Jumlah n bilangan asli','\\sum_{k=1}^{n}k=\\frac{n(n+1)}{2}','Kalkulus'],
  ['taylor','Deret Taylor','e^{x}=1+\\frac{x}{1!}+\\frac{x^{2}}{2!}+\\frac{x^{3}}{3!}+\\cdots','Kalkulus'],
  ['fourier','Deret Fourier','f(x)=a_{0}+\\sum_{n=1}^{\\infty}\\left(a_{n}\\cos\\frac{n\\pi x}{L}+b_{n}\\sin\\frac{n\\pi x}{L}\\right)','Kalkulus'],
  ['normal','Distribusi normal','f(x)=\\frac{1}{\\sigma\\sqrt{2\\pi}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}','Statistika'],
  ['mean','Rata-rata','\\bar{x}=\\frac{1}{n}\\sum_{i=1}^{n}x_{i}','Statistika'],
  ['stdev','Simpangan baku','\\sigma=\\sqrt{\\frac{1}{n}\\sum_{i=1}^{n}(x_{i}-\\mu)^{2}}','Statistika'],
  ['bayes','Teorema Bayes','P(A\\mid B)=\\frac{P(B\\mid A)P(A)}{P(B)}','Statistika'],
  ['binp','Peluang binomial','P(X=k)=\\binom{n}{k}p^{k}(1-p)^{n-k}','Statistika'],
  ['emc','Energi massa Einstein','E=mc^{2}','Fisika'],
  ['newton','Hukum II Newton','\\vec{F}=m\\vec{a}','Fisika'],
  ['grav','Gravitasi Newton','F=G\\frac{m_{1}m_{2}}{r^{2}}','Fisika'],
  ['kinetik','Energi kinetik','E_{k}=\\frac{1}{2}mv^{2}','Fisika'],
  ['gerak','Gerak lurus berubah beraturan','s=v_{0}t+\\frac{1}{2}at^{2}','Fisika'],
  ['schrod','Persamaan Schrödinger','i\\hbar\\frac{\\partial}{\\partial t}\\Psi=\\hat{H}\\Psi','Fisika'],
  ['air','Reaksi pembentukan air','2\\mathrm{H_{2}}+\\mathrm{O_{2}}\\rightarrow 2\\mathrm{H_{2}O}','Kimia'],
  ['amonia','Sintesis amonia','\\mathrm{N_{2}}+3\\mathrm{H_{2}}\\rightleftharpoons 2\\mathrm{NH_{3}}','Kimia'],
  ['ph','Rumus pH','\\mathrm{pH}=-\\log\\left[\\mathrm{H^{+}}\\right]','Kimia'],
  ['gasideal','Gas ideal','PV=nRT','Kimia']
];
const SW_COL=['000000','7F7F7F','C00000','FF0000','FF8000','FFC000','92D050','00A000','00B0F0','0070C0','002060','7030A0','FF00FF','8B4513'];
const SW_HL=['FFFF00','92D050','00FFFF','FF99CC','FFC000','D9D9D9','BDD7EE','FFD6D6'];
const SIZES_PT=[8,9,10,11,12,14,16,18,20,24,28,36,48];

/* kotak "‸" jadi \square, skrip tanpa dasar diberi x, supaya pratinjau tombol mewakili fungsi aslinya */
const glyphOf=it=>{
  if(it[3])return it[3];
  let t=it[1];if(/^[\^_]/.test(t))t='x'+t;
  return t.replace(CM,'\\square ');
};

let cssDone=false;
function ensureCSS(){
  if(cssDone||typeof document==='undefined')return;cssDone=true;
  const st=document.createElement('style');
  st.textContent=`
.wd .eqdg{width:min(560px,96%);max-width:96%;max-height:96%;overflow:auto;padding:12px 14px;box-sizing:border-box}
.wd .eqdg .eqpv{min-height:56px;max-height:26vh;overflow:auto;border:1px solid #c8c8c8;background:#fafafa;display:grid;place-items:center;padding:6px 8px;margin-bottom:8px;color:#222}
.wd .eqdg .eqpv math{font-family:"Cambria Math","STIX Two Math","Noto Sans Math","Latin Modern Math",math;font-size:1.45em}
.wd .eqdg .eqpv .e{color:#a4262c;font-size:12px;text-align:left;white-space:pre-wrap}
.wd .eqdg .eqpv .h{color:#999;font-size:12px}
.wd .eqdg textarea{width:100%;font:15px/1.35 Consolas,"Courier New",monospace;border:1px solid #aaa;padding:5px 6px;resize:none;min-height:48px;display:block;box-sizing:border-box}
.wd .eqdg .eqnote{font-size:11px;color:#8a6d00;background:#fff8e1;border:1px solid #f0dc9a;padding:4px 6px;margin-top:6px}
.wd .eqdg .eqst{display:flex;flex-wrap:wrap;gap:4px;align-items:center;margin:8px 0 2px;position:relative}
.wd .eqdg .eqst button,.wd .eqdg .eqst select{height:32px;min-width:32px;padding:0 8px;border:1px solid #c8c8c8;background:#f7f7f7;cursor:pointer;font-size:13px;color:#222;box-sizing:border-box}
.wd .eqdg .eqst button:active,.wd .eqdg .eqst button:hover{background:#dbe7f7;border-color:var(--blue)}
.wd .eqdg .eqst .eqa{display:block;font-weight:700;line-height:15px}
.wd .eqdg .eqst .eqcu{display:block;height:3px;margin:0 0 -1px}
.wd .eqdg .eqst .sp{flex:1}
.wd .eqdg .eqhint{font-size:11px;color:#666;margin:2px 0 4px}
.wd .eqdg .eqrow{display:flex;align-items:center;justify-content:space-between;gap:6px}
.wd .eqdg .eqrow button{height:28px;padding:0 8px;border:1px solid #c8c8c8;background:#f7f7f7;cursor:pointer;font-size:12px;white-space:nowrap;flex:none}
.wd .eqdg .eqrow button:active{background:#dbe7f7}
.wd .eqdg .eqpal .gl{display:grid;place-items:center;max-width:100%;overflow:hidden}
.wd .eqdg .eqsw,.wd .eqbar .eqsw{position:absolute;z-index:6;left:0;top:36px;background:#fff;border:1px solid #888;box-shadow:0 2px 8px rgba(0,0,0,.3);padding:6px;width:max-content;max-width:94vw;text-align:left}
.wd .eqbar .eqsw{top:100%;margin-top:2px}
.wd .eqsw .gr{display:grid;grid-template-columns:repeat(7,28px);gap:4px;margin-bottom:6px}
.wd .eqsw .sl{font-size:11px;color:#555;margin:2px 0 3px}
.wd .eqsw .gr button{width:28px;height:28px;min-width:0;padding:0;border:1px solid #999;cursor:pointer}
.wd .eqsw .gr button:hover{outline:2px solid var(--blue);outline-offset:1px}
.wd .eqsw .oto{display:block;width:100%;height:28px;border:1px solid #bbb;background:#f5f5f5;cursor:pointer;font-size:12px;margin-bottom:6px;padding:0 8px}
.wd .eqsw label.cu{font-size:12px;display:flex;align-items:center;gap:6px}
.wd .eqsw input[type=color]{width:34px;height:26px;padding:0;border:1px solid #999}
.wd .eqdg .eqtabs{display:flex;gap:2px;margin:8px 0 4px;overflow-x:auto;scrollbar-width:none}
.wd .eqdg .eqtabs button{flex:none;height:30px;padding:0 10px;border:1px solid #c8c8c8;background:#f3f3f3;cursor:pointer;white-space:nowrap}
.wd .eqdg .eqtabs button.on{background:#fff;color:var(--blue);border-color:var(--blue);font-weight:600}
.wd .eqdg .eqpal{display:grid;grid-template-columns:repeat(auto-fill,minmax(66px,1fr));gap:4px;max-height:200px;overflow:auto;padding:3px;border:1px solid #e1e1e1;background:#fff;align-content:start}
.wd .eqdg .eqpal.sym{grid-template-columns:repeat(auto-fill,minmax(38px,1fr))}
.wd .eqdg .eqpal.ex{grid-template-columns:repeat(auto-fill,minmax(150px,1fr))}
.wd .eqdg .eqpal button{min-height:36px;padding:2px 4px;border:1px solid #d4d4d4;background:#fbfbfb;cursor:pointer;font-size:14px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1px;color:#222;min-width:0}
.wd .eqdg .eqpal:not(.sym) button{min-height:58px}
.wd .eqdg .eqpal button:hover,.wd .eqdg .eqpal button:active{background:#dbe7f7;border-color:var(--blue)}
.wd .eqdg .eqpal button small{font-size:9.5px;line-height:1.1;color:#555;text-align:center;overflow:hidden;max-width:100%;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.wd .eqdg .eqpal math{font-size:15px;font-family:"Cambria Math","STIX Two Math","Noto Sans Math",math;pointer-events:none}
.wd .eqdg .eqpal.ex math{font-size:14px}
.wd .eqdg .eqpal .gh{grid-column:1/-1;font-size:11px;font-weight:600;color:var(--blue);padding:4px 2px 0;border-bottom:1px solid #e1e1e1}
.wd .eqdg .eqmode{display:flex;gap:14px;margin-top:8px;font-size:12px;flex-wrap:wrap}
.wd .eqdg .eqmode label{display:flex;align-items:center;gap:5px;cursor:pointer}
.wd .eqdg .eqmode input{width:auto;height:auto}
.wd .page .eq.eqsel{outline:2px solid var(--blue);outline-offset:2px;background:rgba(43,87,154,.08)}
.wd .eqbar{position:absolute;z-index:25;display:flex;background:#fff;border:1px solid #8a8a8a;box-shadow:0 2px 8px rgba(0,0,0,.28)}
.wd .eqbar>button{border:0;background:none;height:34px;padding:0 11px;cursor:pointer;font-size:13px;white-space:nowrap;position:relative}
.wd .eqbar>button+button{border-left:1px solid #ddd}
.wd .eqbar>button:active,.wd .eqbar>button:hover{background:#dbe7f7}
.wd .eqbar button[data-q=del]{color:#a4262c}
.wd .eqbar button[data-q=col] u{position:absolute;left:9px;right:9px;bottom:4px;height:3px;background:linear-gradient(90deg,#c00,#fc0,#0a0,#07c,#70a)}`;
  document.head.appendChild(st);
}

function parseXML(s){
  const d=new root.DOMParser().parseFromString(s,'application/xml');
  if(d.getElementsByTagName('parsererror').length)throw new EqError('OMML tidak valid');
  return d.documentElement;
}
function mathHTML(omml,display){
  const WM=root.WordMath,el=parseXML(omml);
  if(WM&&WM.supported&&WM.supported())return WM.toMathML(el,display);
  return xe(WM&&WM.plainText?WM.plainText(el):'');
}

/* ---------- gaya di kotak LaTeX: bungkus / lepas perintah pada teks terpilih ---------- */
function matchBrace(s,i){let d=0;for(let j=i;j<s.length;j++){const c=s[j];if(c==='\\'){j++;continue;}if(c==='{')d++;else if(c==='}'){d--;if(d===0)return j;}}return -1;}
function unwrapCmd(sel,names,nargs){                 /* sel persis \perintah{..}(..)  -> {args:[..]} atau null */
  const m=/^\\([A-Za-z]+)\s*/.exec(sel);if(!m||!names.includes(m[1]))return null;
  let j=m[0].length;const args=[];
  for(let k=0;k<nargs;k++){
    while(/\s/.test(sel[j]||''))j++;
    if(sel[j]!=='{')return null;const e=matchBrace(sel,j);if(e<0)return null;
    args.push(sel.slice(j+1,e));j=e+1;
  }
  return j===sel.length?{args,inner:args[args.length-1]}:null;
}
const STYLE_ONE=['mathbf','textbf','boldsymbol','bm','mathit','textit','mathrm','textrm','mathup'];
const STYLE_TWO=['textcolor','colorbox','highlight','bgcolor','size'];
function stripStyles(str){                           /* buang semua perintah gaya, isinya dipertahankan */
  let out='',i=0;
  while(i<str.length){
    if(str[i]==='\\'){
      const m=/^\\([A-Za-z]+)/.exec(str.slice(i));
      if(m){
        const nm=m[1],after=i+m[0].length;
        const two=STYLE_TWO.includes(nm);
        if(two||STYLE_ONE.includes(nm)){
          let j=after,ok=true,inner='';
          for(let k=0;k<(two?2:1);k++){
            while(/\s/.test(str[j]||''))j++;
            if(str[j]!=='{'){ok=false;break;}
            const e=matchBrace(str,j);if(e<0){ok=false;break;}
            inner=str.slice(j+1,e);j=e+1;
          }
          if(ok){out+=stripStyles(inner);i=j;continue;}
        }
        out+=m[0];i=after;continue;
      }
      out+=str.slice(i,i+2);i+=2;continue;
    }
    out+=str[i++];
  }
  return out;
}
const STCMD={b:['mathbf',['mathbf','textbf','boldsymbol','bm']],i:['mathit',['mathit','textit']],p:['mathrm',['mathrm','textrm','mathup']]};
/* terapkan gaya pada string `sel`; kembalikan string baru */
function styleText(sel,kind,val){
  if(kind==='clr')return stripStyles(sel);
  if(STCMD[kind]){
    const [cmd,names]=STCMD[kind],u=unwrapCmd(sel,names,1);
    return u?u.inner:'\\'+cmd+'{'+sel+'}';
  }
  const cmd={col:'textcolor',hl:'colorbox',sz:'size'}[kind];
  const names=kind==='col'?['textcolor']:kind==='hl'?['colorbox','highlight','bgcolor']:['size'];
  const u=unwrapCmd(sel,names,2),inner=u?u.inner:sel;
  if(val===null||val===''||(kind==='sz'&&+val===0))return inner;           /* "Otomatis" / "Normal": lepas gaya */
  return '\\'+cmd+'{'+val+'}{'+inner+'}';
}
/* seluruh rumus tanpa blok: rumus beberapa baris (dipisah \\) dibungkus per baris supaya garis baru tidak hilang */
function styleWhole(all,kind,val){
  if(/\\\\/.test(all)&&!/\\begin/.test(all))return all.split(/\s*\\\\\s*/).map(l=>l.trim()?styleText(l.trim(),kind,val):l).join(' \\\\ ');
  return styleText(all,kind,val);
}

/* o: {host,title,okText,tex,display,note,onOk({tex,display,omml})->false untuk tetap terbuka} */
function openDialog(o){
  ensureCSS();
  const ov=document.createElement('div');ov.className='ov';
  const sw=(arr,k)=>`<div class="gr">${arr.map(c=>`<button type="button" data-k="${k}" data-c="${c}" style="background:#${c}" title="#${c}"></button>`).join('')}</div>`;
  ov.innerHTML=`<div class="dg eqdg"><h3>${xe(o.title||'Rumus')}</h3>
<div class="eqpv"></div>
<textarea class="eqtx" rows="2" spellcheck="false" autocapitalize="off" autocomplete="off" autocorrect="off" placeholder="Ketik rumus (LaTeX), atau ketuk tombol di bawah, mis. \\frac{a}{b}"></textarea>
${o.note?`<div class="eqnote">${xe(o.note)}</div>`:''}
<div class="eqst">
<button type="button" data-st="b" title="Tebal"><span class="eqa">B</span></button>
<button type="button" data-st="i" title="Miring"><span class="eqa" style="font-style:italic;font-family:serif">I</span></button>
<button type="button" data-st="p" title="Tegak (tidak miring)"><span class="eqa" style="font-family:serif">A</span></button>
<button type="button" data-st="col" title="Warna rumus"><span class="eqa">A</span><u class="eqcu" style="background:linear-gradient(90deg,#c00,#fc0,#0a0,#07c,#70a)"></u></button>
<button type="button" data-st="hl" title="Warna sorot (latar)"><span class="eqa">ab</span><u class="eqcu" style="background:#ff0"></u></button>
<select data-st="sz" title="Ukuran rumus" style="width:84px"><option value="">Ukuran</option><option value="0">Normal</option>${SIZES_PT.map(n=>`<option value="${n}">${n} pt</option>`).join('')}</select>
<button type="button" data-st="clr" title="Hapus semua gaya pada pilihan">⌫</button>
<div class="eqsw" hidden></div>
</div>
<div class="eqhint">Gaya berlaku pada teks yang diblok; tanpa blok = seluruh rumus.</div>
<div class="eqhint eqrow"><span>Ketuk tombol untuk menyisipkan, lalu isi kotak □.</span><button type="button" data-st="next" title="Pindah ke kotak kosong berikutnya (Tab)">⇥ Kotak berikutnya</button></div>
<div class="eqtabs">${Object.keys(PAL).map((k,i)=>`<button type="button" data-g="${k}"${i?'':' class="on"'}>${PAL[k][0]}</button>`).join('')}<button type="button" data-g="e">Contoh</button></div>
<div class="eqpal"></div>
<div class="eqmode"><label><input type="radio" name="eqm" value="i"> Sebaris (dalam teks)</label><label><input type="radio" name="eqm" value="b"> Blok (di tengah)</label></div>
<div class="dbt"><button type="button" class="x">Batal</button><button type="button" class="ok">${xe(o.okText||'Sisipkan')}</button></div></div>`;
  o.host.appendChild(ov);
  const $=s=>ov.querySelector(s),ta=$('.eqtx'),pv=$('.eqpv'),pal=$('.eqpal'),okb=$('.ok'),swp=$('.eqsw');
  const rad=v=>ov.querySelector(`input[name=eqm][value=${v}]`);
  ta.value=o.tex||'';rad(o.display?'b':'i').checked=true;
  const WM=root.WordMath,mm=WM&&WM.supported&&WM.supported();
  let good=false;
  const display=()=>rad('b').checked;
  function update(){
    const t=ta.value;
    if(!t.trim()){pv.innerHTML='<span class="h">Pratinjau rumus muncul di sini</span>';good=false;okb.disabled=true;return;}
    try{
      const om=tex2omml(t,{display:display(),ph:true});
      if(isPara(om)&&!display())rad('b').checked=true;     /* beberapa baris otomatis blok */
      pv.innerHTML=mathHTML(om,isPara(om));good=true;okb.disabled=false;
    }catch(e){pv.innerHTML=`<div class="e">${xe(e.message||String(e))}</div>`;good=false;okb.disabled=true;}
  }
  let tm=0;const later=()=>{clearTimeout(tm);tm=setTimeout(update,90);};
  const glyph=(tex)=>{try{return mathHTML(tex2omml(tex,{ph:true}),false);}catch(e){return null;}};
  function renderPal(k){
    pal._k=k;
    if(k==='e'){                                            /* contoh rumus lengkap, dikelompokkan */
      pal.className='eqpal ex';let h='',g='';
      GALLERY.forEach((it,i)=>{
        if(it[3]!==g){g=it[3];h+=`<div class="gh">${xe(g)}</div>`;}
        let m=mm?glyph(it[2]):null;
        h+=`<button type="button" data-e="${i}" title="${xa(it[1])}"><span class="gl">${m||xe(it[2])}</span><small>${xe(it[1])}</small></button>`;
      });
      pal.innerHTML=h;fit();return;
    }
    const [,items,cap]=PAL[k];
    pal.className='eqpal'+(cap?'':' sym');
    pal.innerHTML=items.map((it,i)=>{
      let g=xe(it[0]);
      if(cap&&mm){const m=glyph(glyphOf(it));if(m)g=m;}
      return `<button type="button" data-i="${i}" title="${xa(it[2]||'')}"><span class="gl">${g}</span>${cap?`<small>${xe(it[2]||'')}</small>`:''}</button>`;}).join('');
    fit();
  }
  function fit(){                                            /* perkecil glyph yang lebih besar dari tombolnya */
    pal.querySelectorAll('button').forEach(b=>{
      const m=b.querySelector('math');if(!m)return;
      const maxW=b.clientWidth-10,maxH=pal.classList.contains('ex')?46:40;
      const r=m.getBoundingClientRect();if(!r.width)return;
      const sc=Math.min(1,maxW/r.width,maxH/r.height);
      if(sc<1)m.style.fontSize=Math.max(7,Math.floor(parseFloat(getComputedStyle(m).fontSize)*sc*10)/10)+'px';
    });
  }
  function insert(tpl){
    const s=ta.selectionStart,e=ta.selectionEnd,sel=ta.value.slice(s,e);
    if(/^[\^_]/.test(tpl)){                                 /* pangkat/indeks tanpa dasar di depannya: beri kotak dasar */
      const before=ta.value.slice(0,s).replace(/\s+$/,'');
      if(!before||/[{(\[=+\-,&|]$/.test(before)||/\\\\$/.test(before))tpl='\\square '+tpl;
    }
    let at=tpl.indexOf(CM);const has=at>=0;
    let t=has?tpl.replace(CM,sel):tpl;
    if(!has&&/\\[A-Za-z]+$/.test(t)&&!/\s$/.test(t))t+=' ';
    ta.setRangeText(t,s,e,'end');
    const pos=has?s+at+sel.length:s+t.length;
    ta.setSelectionRange(pos,pos);ta.focus({preventScroll:true});later();
  }
  function nextSlot(){                                      /* loncat ke {} atau [] kosong berikutnya */
    const v=ta.value,from=ta.selectionEnd,re=/\{\}|\[\]/g;let m,hit=null,first=null;
    while((m=re.exec(v))){if(!first)first=m;if(!hit&&m.index>=from)hit=m;}
    const t=hit||first;if(!t)return false;
    ta.focus({preventScroll:true});ta.setSelectionRange(t.index+1,t.index+1);return true;
  }
  function applyStyle(kind,val){
    const v=ta.value;let s=ta.selectionStart,e=ta.selectionEnd;const whole=s===e;
    if(whole){s=0;e=v.length;}
    while(s<e&&/\s/.test(v[s]))s++;while(e>s&&/\s/.test(v[e-1]))e--;
    if(s>=e)return;
    const sel=v.slice(s,e),out=whole?styleWhole(sel,kind,val):styleText(sel,kind,val);
    ta.setRangeText(out,s,e,'select');update();
  }
  function closeSw(){swp.hidden=true;swp.innerHTML='';}
  function openSw(kind){
    if(!swp.hidden&&swp._k===kind){closeSw();return;}
    swp._k=kind;
    const col=kind==='col';
    swp.innerHTML=`<button type="button" class="oto" data-k="${kind}" data-c="">${col?'Otomatis (hapus warna)':'Tanpa sorot'}</button>`+
      `<div class="sl">${col?'Warna rumus':'Warna sorot'}</div>${sw(col?SW_COL:SW_HL,kind)}`+
      `<label class="cu">Warna lain: <input type="color" data-k="${kind}" value="${col?'#c00000':'#ffff00'}"></label>`;
    swp.hidden=false;
  }
  renderPal('d');
  ov.addEventListener('mousedown',e=>{if(e.target.closest('.eqpal button,.eqtabs button,.eqst button,.eqsw button,.eqrow button'))e.preventDefault();});
  pal.addEventListener('click',e=>{
    const be=e.target.closest('button[data-e]');
    if(be){const it=GALLERY[+be.dataset.e];
      if(!ta.value.trim()){ta.value=it[2];ta.setSelectionRange(it[2].length,it[2].length);later();}
      else{const s=ta.selectionStart;ta.setRangeText(it[2],s,ta.selectionEnd,'end');later();}
      ta.focus({preventScroll:true});return;}
    const b=e.target.closest('button[data-i]');if(b)insert(PAL[pal._k][1][+b.dataset.i][1]);
  });
  $('.eqtabs').addEventListener('click',e=>{const b=e.target.closest('button[data-g]');if(!b)return;
    ov.querySelectorAll('.eqtabs button').forEach(x=>x.classList.toggle('on',x===b));renderPal(b.dataset.g);pal.scrollTop=0;});
  ov.addEventListener('click',e=>{
    if(e.target.closest('.eqrow button[data-st=next]')){closeSw();nextSlot();return;}
    if(!e.target.closest('.eqst'))return;
    const sc=e.target.closest('.eqsw button[data-k]');
    if(sc){applyStyle(sc.dataset.k,sc.dataset.c);closeSw();return;}
    const b=e.target.closest('button[data-st]');if(!b)return;
    const k=b.dataset.st;
    if(k==='col'||k==='hl')openSw(k);
    else if(k==='next'){closeSw();nextSlot();}
    else{closeSw();applyStyle(k);}
  });
  $('.eqst').addEventListener('change',e=>{
    const c=e.target.closest('input[type=color]');
    if(c){applyStyle(c.dataset.k,c.value.slice(1).toUpperCase());closeSw();return;}
    const s=e.target.closest('select[data-st=sz]');
    if(s){if(s.value!=='')applyStyle('sz',s.value);s.value='';}
  });
  ta.addEventListener('input',later);
  ta.addEventListener('keydown',e=>{
    if(e.key==='Tab'&&!e.shiftKey&&!e.ctrlKey&&!e.altKey&&/\{\}|\[\]/.test(ta.value)){e.preventDefault();nextSlot();}
  });
  ov.querySelectorAll('input[name=eqm]').forEach(r=>r.addEventListener('change',update));
  const close=()=>{clearTimeout(tm);ov.remove();};
  $('.x').onclick=close;
  okb.onclick=()=>{
    update();if(!good)return;
    const d=display(),om=tex2omml(ta.value,{display:d});
    if(o.onOk({tex:ta.value,display:isPara(om),omml:om})!==false)close();
  };
  ov.addEventListener('keydown',e=>{
    if(e.key==='Escape'){e.stopPropagation();if(!swp.hidden)closeSw();else close();}
    else if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();okb.click();}
  });
  ov.addEventListener('mousedown',e=>{if(e.target===ov)e.stopPropagation();});
  update();ta.focus({preventScroll:true});ta.setSelectionRange(ta.value.length,ta.value.length);
  return {close,ta,update,applyStyle};
}

/* ---------- ubah gaya OMML yang sudah ada (tanpa lewat LaTeX, jadi bagian lain tetap persis) ----------
   ops: {col:'FF0000'|null(lepas), hl:..., sz:pt|null, bold:true|false|'toggle', ital:true|false|'toggle'} */
const MPR={f:'fPr',rad:'radPr',sSup:'sSupPr',sSub:'sSubPr',sSubSup:'sSubSupPr',sPre:'sPrePr',nary:'naryPr',d:'dPr',func:'funcPr',limLow:'limLowPr',limUpp:'limUppPr',acc:'accPr',bar:'barPr',groupChr:'groupChrPr',borderBox:'borderBoxPr',phant:'phantPr',eqArr:'eqArrPr',m:'mPr'};
function restyle(omml,ops,DP,XS){
  DP=DP||root.DOMParser;XS=XS||root.XMLSerializer;
  const doc=new DP().parseFromString(omml,'application/xml'),top=doc.documentElement;
  const mk=(ns,n)=>doc.createElementNS(ns,(ns===M_NS?'m:':'w:')+n);
  const runs=[...top.getElementsByTagNameNS(M_NS,'r')];
  const txt=r=>[...r.getElementsByTagNameNS(M_NS,'t')].map(t=>t.textContent).join('');
  const styV=r=>{const p=mrp(r),v=p&&mattr(kid(p,'sty'),'val');return v||null;};
  const isNor=r=>{const p=mrp(r);return !!(p&&on(kid(p,'nor')));};
  const isB=r=>{const v=styV(r);return v==='b'||v==='bi';};
  const isI=r=>{const v=styV(r);return v?(v==='i'||v==='bi'):/^\p{L}+$/u.test(txt(r));};
  const mine=runs.filter(r=>!isNor(r));
  const tgt=(want,fn)=>want==='toggle'?!(mine.length&&mine.every(fn)):want;
  const wantB=tgt(ops.bold,isB),wantI=tgt(ops.ital,isI);
  if(wantB!==undefined||wantI!==undefined){
    for(const r of mine){
      const b=wantB===undefined?isB(r):wantB,i=wantI===undefined?isI(r):wantI;
      let p=mrp(r);if(!p){p=mk(M_NS,'rPr');r.insertBefore(p,r.firstChild);}
      let s=kid(p,'sty');if(!s){s=mk(M_NS,'sty');p.insertBefore(s,p.firstChild);}
      s.setAttributeNS(M_NS,'m:val',b?(i?'bi':'b'):(i?'i':'p'));
    }
  }
  const touch=('col' in ops)||('hl' in ops)||('sz' in ops);
  const setRpr=rp=>{
    const drop=n=>{for(const x of [...rp.childNodes])if(x.nodeType===1&&x.namespaceURI===W_NS&&x.localName===n)rp.removeChild(x);};
    if('col' in ops){drop('color');if(ops.col){const c=mk(W_NS,'color');c.setAttributeNS(W_NS,'w:val',ops.col);rp.appendChild(c);}}
    if('sz' in ops){drop('sz');drop('szCs');if(ops.sz>0){const h=String(Math.round(ops.sz*2));for(const n of ['sz','szCs']){const z=mk(W_NS,n);z.setAttributeNS(W_NS,'w:val',h);rp.appendChild(z);}}}
    if('hl' in ops){drop('shd');drop('highlight');if(ops.hl){const s=mk(W_NS,'shd');s.setAttributeNS(W_NS,'w:val','clear');s.setAttributeNS(W_NS,'w:color','auto');s.setAttributeNS(W_NS,'w:fill',ops.hl);rp.appendChild(s);}}
  };
  if(touch){
    for(const r of runs){
      let rp=wkid(r,'rPr');
      if(!rp){rp=mk(W_NS,'rPr');const t=[...r.childNodes].find(x=>x.nodeType===1&&x.localName==='t');r.insertBefore(rp,t||null);}
      setRpr(rp);
    }
    for(const el of [...top.getElementsByTagName('*')]){
      const pn=el.namespaceURI===M_NS&&MPR[el.localName];if(!pn)continue;
      let pr=[...el.childNodes].find(x=>x.nodeType===1&&x.namespaceURI===M_NS&&x.localName===pn);
      if(!pr){pr=mk(M_NS,pn);el.insertBefore(pr,el.firstChild);}
      let cp=kid(pr,'ctrlPr');if(!cp){cp=mk(M_NS,'ctrlPr');pr.appendChild(cp);}
      let rp=wkid(cp,'rPr');if(!rp){rp=mk(W_NS,'rPr');cp.appendChild(rp);}
      setRpr(rp);
    }
  }
  return new XS().serializeToString(top).replace(/^<\?xml[^>]*\?>\s*/,'');
}

/* ============================ PASANG DI EDITOR WORD ============================ */
/* ctx: {host(.wd), page, focus, refresh, stats, toast} */
function attach(ctx){
  const {host,page,focus,refresh,stats,toast}=ctx;
  ensureCSS();
  let armed=null,bar=null,lastTap={el:null,t:0};
  const WM=()=>root.WordMath;
  const eqHTML=(omml,display)=>WM().eqHTML({omml,display});
  const clean=s=>(s||'').replace(/\u200b/g,'').trim();
  function blockOf(){
    const s=getSelection();let n=s.rangeCount?s.anchorNode:null;if(!n||!page.contains(n))return null;
    if(n.nodeType===3)n=n.parentElement;
    while(n&&n.parentElement!==page)n=n.parentElement;
    return n&&n!==page?n:null;
  }
  const blockEmpty=b=>!b||(!clean(b.textContent)&&!b.querySelector('.eq,img,table'));
  function hide(){
    if(armed)armed.classList.remove('eqsel');armed=null;
    if(bar){bar.remove();bar=null;}
  }
  const ZW='&#8203;';
  /* Dua perilaku Chromium yang harus dihindari (diuji langsung di Chromium 141):
     1) fragmen berisi <math> dianggap blok oleh insertHTML -> paragraf terpecah. Maka yang disisipkan hanya "cangkang" <span class="eq">
        berisi teks sementara; MathML diisi lewat DOM sesudahnya (undo/redo tetap jalan karena cangkangnya yang dicatat).
     2) <span contenteditable=false> polos jatuh di LUAR <p>; ZWSP di kedua sisi menahannya di dalam paragraf (ZWSP dibuang docx-io saat simpan).
        Rumus blok di paragraf kosong tidak butuh ZWSP. */
  const shell=(omml,display,zw)=>{
    const h=`<span class="eq${display?' eqd':''}" contenteditable="false" data-omml="${xa(omml)}" data-eqnew="1">\u22EF</span>`;
    return zw?ZW+h+ZW:h;
  };
  function fill(omml,display){
    const el=page.querySelector('[data-eqnew]');if(!el)return null;
    el.removeAttribute('data-eqnew');
    const t=document.createElement('div');t.innerHTML=eqHTML(omml,display);
    el.innerHTML=t.firstChild?t.firstChild.innerHTML:'';
    return el;
  }
  const put=(omml,display)=>{document.execCommand('insertHTML',false,shell(omml,display,!display));return fill(omml,display);};
  function setCaretStart(el){const r=document.createRange(),s=getSelection();r.selectNodeContents(el);r.collapse(true);s.removeAllRanges();s.addRange(r);}
  function insertNew(omml,display){
    focus();
    if(display){
      const b=blockOf();
      if(b&&!blockEmpty(b)){
        document.execCommand('insertParagraph');            /* pecah di kursor */
        const nb=blockOf();
        if(nb&&!blockEmpty(nb)){                            /* ada teks sesudah kursor: sisakan satu paragraf kosong di antaranya */
          document.execCommand('insertParagraph');
          const cur=blockOf(),prev=cur&&cur.previousElementSibling;if(prev)setCaretStart(prev);
        }
      }
      put(omml,true);
      const cb=blockOf(),nx=cb&&cb.nextElementSibling;
      if(nx&&/^(P|H[1-6]|LI|BLOCKQUOTE)$/.test(nx.tagName))setCaretStart(nx);else document.execCommand('insertParagraph');
    }else put(omml,false);
    refresh();stats();
  }
  /* rentang yang membungkus rumus beserta ZWSP pengapitnya */
  function fullRange(eq){
    const r=document.createRange();r.selectNode(eq);
    const pv=eq.previousSibling,nx=eq.nextSibling;
    if(pv&&pv.nodeType===3&&/\u200b$/.test(pv.nodeValue))r.setStart(pv,pv.nodeValue.length-1);
    if(nx&&nx.nodeType===3&&/^\u200b/.test(nx.nodeValue))r.setEnd(nx,1);
    return r;
  }
  function select(eq){const s=getSelection();focus();s.removeAllRanges();s.addRange(fullRange(eq));}
  function replace(eq,omml,display){
    const was=eq.classList.contains('eqd');
    select(eq);hide();
    if(display&&!was){                                      /* sebaris -> blok: lepas dulu, lalu sisipkan sebagai paragraf sendiri */
      document.execCommand('delete');insertNew(omml,true);return null;
    }
    const ne=put(omml,display);
    refresh();stats();
    return ne;
  }
  function remove(eq){select(eq);hide();document.execCommand('delete');refresh();stats();}
  function toggleMode(eq){
    const want=!eq.classList.contains('eqd'),om=setDisplay(eq.getAttribute('data-omml'),want);
    if(om===null){toast('Rumus beberapa baris hanya bisa berbentuk blok');return;}
    replace(eq,om,want);
  }
  function dialog(eq){
    let tex='',display=false,note='',lossy=false,orig=null;
    if(eq){
      orig=eq.getAttribute('data-omml');
      let basePt=0;
      try{const px=parseFloat(getComputedStyle(eq.parentElement||page).fontSize);if(px>0)basePt=Math.round(px*0.75*2)/2;}catch(e){}
      try{const r=omml2tex(parseXML(orig),{basePt});tex=r.tex;display=r.display;lossy=r.lossy;}
      catch(e){toast('Rumus tidak bisa dibuka untuk diedit: '+e.message);return;}
      if(lossy)note='Rumus ini memuat elemen Word yang tidak ada padanannya di editor; bentuknya bisa sedikit berubah jika Anda mengubah isinya.';
    }else{
      const s=getSelection(),txt=s.rangeCount&&page.contains(s.anchorNode)?s.toString().trim():'';
      tex=txt.replace(/\s*\n\s*/g,' ');
      display=blockEmpty(blockOf())&&!tex;
    }
    const t0=tex,d0=display;
    openDialog({host,title:eq?'Edit Rumus':'Sisipkan Rumus',okText:eq?'Terapkan':'Sisipkan',tex,display,note,
      onOk(res){
        if(!eq){insertNew(res.omml,res.display);return;}
        if(res.tex===t0){                       /* isi tidak diubah: pertahankan OMML asli */
          if(res.display===d0)return;
          const om=setDisplay(orig,res.display);
          if(om!==null){replace(eq,om,res.display);return;}
        }
        replace(eq,res.omml,res.display);
      }});
  }
  /* ubah gaya rumus langsung (tanpa membuka dialog); bilah tetap terbuka di rumus yang baru */
  function quick(eq,ops){
    let om;try{om=restyle(eq.getAttribute('data-omml'),ops);}catch(e){toast('Gaya gagal diterapkan: '+e.message);return;}
    const ne=replace(eq,om,eq.classList.contains('eqd'));
    if(ne&&page.contains(ne))showBar(ne);
  }
  function showBar(eq){
    hide();armed=eq;eq.classList.add('eqsel');
    bar=document.createElement('div');bar.className='eqbar';
    const sw=(arr,k)=>`<div class="gr">${arr.map(c=>`<button type="button" data-k="${k}" data-c="${c}" style="background:#${c}" title="#${c}"></button>`).join('')}</div>`;
    bar.innerHTML=`<button type="button" data-q="edit">✎ Edit</button><button type="button" data-q="b" title="Tebal"><b>B</b></button><button type="button" data-q="i" title="Miring"><i style="font-family:serif">I</i></button><button type="button" data-q="col" title="Warna & sorot">A<u></u></button><button type="button" data-q="mode">${eq.classList.contains('eqd')?'Sebaris':'Blok'}</button><button type="button" data-q="del">Hapus</button>`+
      `<div class="eqsw" hidden><button type="button" class="oto" data-k="col" data-c="">Warna otomatis</button><div class="sl">Warna rumus</div>${sw(SW_COL,'col')}<button type="button" class="oto" data-k="hl" data-c="">Tanpa sorot</button><div class="sl">Warna sorot</div>${sw(SW_HL,'hl')}</div>`;
    host.appendChild(bar);
    const hr=host.getBoundingClientRect(),er=eq.getBoundingClientRect(),bw=bar.offsetWidth,bh=bar.offsetHeight;
    let L=er.left-hr.left,T=er.top-hr.top-bh-6;
    if(T<30)T=er.bottom-hr.top+6;
    L=Math.max(4,Math.min(L,hr.width-bw-4));
    bar.style.left=L+'px';bar.style.top=T+'px';
    bar.addEventListener('mousedown',e=>e.preventDefault());
    bar.addEventListener('click',e=>{
      const t=armed;if(!t)return;
      const sc=e.target.closest('.eqsw button[data-k]');
      if(sc){quick(t,sc.dataset.k==='col'?{col:sc.dataset.c||null}:{hl:sc.dataset.c||null});return;}
      const b=e.target.closest('button[data-q]');if(!b)return;const q=b.dataset.q;
      if(q==='edit'){hide();dialog(t);}else if(q==='mode')toggleMode(t);else if(q==='del')remove(t);
      else if(q==='b')quick(t,{bold:'toggle'});
      else if(q==='i')quick(t,{ital:'toggle'});
      else if(q==='col'){const p=bar.querySelector('.eqsw');p.hidden=!p.hidden;}
    });
  }
  function onDown(e){
    if(!host.isConnected){document.removeEventListener('pointerdown',onDown,true);return;}
    if(!armed)return;
    if(bar&&bar.contains(e.target))return;
    if(e.target.closest&&e.target.closest('.eq,.ribbon,.tabs'))return;   /* tombol pita (Edit Rumus) butuh rumus yang masih terpilih */
    hide();
  }
  document.addEventListener('pointerdown',onDown,true);
  page.addEventListener('keydown',()=>hide());
  return {
    /* dipanggil dari klik di halaman; true bila klik mengenai rumus */
    tap(e){
      const q=e.target.closest&&e.target.closest('.eq');if(!q||!page.contains(q))return false;
      const now=Date.now();
      if(lastTap.el===q&&now-lastTap.t<450){lastTap={el:null,t:0};hide();dialog(q);return true;}
      lastTap={el:q,t:now};showBar(q);return true;
    },
    openNew(){if(armed){const t=armed;hide();dialog(t);}else dialog(null);},
    editCurrent(){
      let q=armed;
      if(!q){
        const s=getSelection();
        if(s.rangeCount&&page.contains(s.anchorNode)){
          const n=s.anchorNode.nodeType===3?s.anchorNode.parentElement:s.anchorNode;
          q=n.closest&&n.closest('.eq');
          if(!q&&!s.isCollapsed){const r=s.getRangeAt(0);q=[...page.querySelectorAll('.eq')].find(x=>r.intersectsNode(x))||null;}
        }
      }
      if(!q){toast('Ketuk rumus yang ingin diedit terlebih dulu');return;}
      hide();dialog(q);
    },
    insertTex(tex,display){
      try{insertNew(tex2omml(tex,{display:!!display}),!!display);}catch(e){toast('Rumus tidak valid: '+e.message);}
    },
    hide
  };
}

root.WordEq={tex2omml,omml2tex,setDisplay,restyle,openDialog,attach,GALLERY,PAL,EqError,_parse:parseTex};
})(typeof window!=='undefined'?window:globalThis);
