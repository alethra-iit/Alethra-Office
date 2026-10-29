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
const FONTCMD={mathrm:'p',textrm:'p',mathbf:'b',textbf:'b',mathit:'i',textit:'i',boldsymbol:'bi',bm:'bi',mathsf:'sf',mathtt:'tt',mathbb:'bb',mathcal:'cal',mathscr:'cal',mathfrak:'frak'};

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
      const n=term();if(n)out.push(n);
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
  function mapRuns(nodes,fn){
    for(const n of nodes){
      if(n.t==='r'){fn(n);continue;}
      for(const k of ['c','n','d','deg','e','sub','sup','lim','name'])if(Array.isArray(n[k]))mapRuns(n[k],fn);
      if(n.es)n.es.forEach(x=>mapRuns(x,fn));
      if(n.rows)n.rows.forEach(r=>{if(Array.isArray(r[0])||r.length===0)r.forEach(c=>mapRuns(c,fn));else mapRuns(r,fn);});
    }
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
    mapRuns(nodes,r=>{
      if(r.nor)return;
      if(ALPHA[kind]){r.s=[...r.s].map(ch=>alphaMap(kind,ch)).join('');}
      else r.sty=kind;
    });
    return {t:'g',c:nodes,style:true};
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
const FONT='<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const val=(n,v)=>`<m:${n} m:val="${xa(v)}"/>`;
function flat(nodes){const o=[];for(const n of nodes||[]){if(n.t==='g'){o.push(...flat(n.c));}else if(n.t==='fnop'){}else o.push(n);}return o;}
function runXML(s,sty,nor){
  let pr='';
  if(nor)pr='<m:nor/>';else if(sty&&sty!=='i')pr=val('sty',{p:'p',b:'b',bi:'bi'}[sty]||'p');
  return `<m:r>${pr?`<m:rPr>${pr}</m:rPr>`:''}${FONT}<m:t${/^\s|\s$/.test(s)?' xml:space="preserve"':''}>${xe(s)}</m:t></m:r>`;
}
function emit(nodes,ph){
  const L=flat(nodes);let out='';
  const E=x=>{const s=emit(x,ph);return s||(ph?runXML('□'):'');};
  const arg=(tag,x)=>`<m:${tag}>${E(x)}</m:${tag}>`;
  for(let i=0;i<L.length;i++){
    const n=L[i];
    if(n.t==='r'){
      let s=n.s;
      while(L[i+1]&&L[i+1].t==='r'&&L[i+1].sty===n.sty&&!!L[i+1].nor===!!n.nor){s+=L[i+1].s;i++;}
      if(s)out+=runXML(s,n.sty,n.nor);continue;
    }
    switch(n.t){
      case 'f':out+=`<m:f>${n.type!=='bar'?`<m:fPr>${val('type',n.type)}</m:fPr>`:''}${arg('num',n.n)}${arg('den',n.d)}</m:f>`;break;
      case 'rad':{const has=n.deg&&n.deg.length;out+=`<m:rad><m:radPr>${has?'':val('degHide','1')}</m:radPr><m:deg>${has?emit(n.deg,ph):''}</m:deg>${arg('e',n.e)}</m:rad>`;break;}
      case 'sup':out+=`<m:sSup>${arg('e',n.e)}${arg('sup',n.sup)}</m:sSup>`;break;
      case 'sub':out+=`<m:sSub>${arg('e',n.e)}${arg('sub',n.sub)}</m:sSub>`;break;
      case 'ssub':out+=`<m:sSubSup>${arg('e',n.e)}${arg('sub',n.sub)}${arg('sup',n.sup)}</m:sSubSup>`;break;
      case 'pre':out+=`<m:sPre>${arg('sub',n.sub)}${arg('sup',n.sup)}${arg('e',n.e)}</m:sPre>`;break;
      case 'nary':out+=`<m:nary><m:naryPr>${val('chr',n.chr)}${val('limLoc',n.loc)}${n.sub?'':val('subHide','1')}${n.sup?'':val('supHide','1')}</m:naryPr><m:sub>${emit(n.sub||[],ph)}</m:sub><m:sup>${emit(n.sup||[],ph)}</m:sup>${arg('e',n.e)}</m:nary>`;break;
      case 'd':out+=`<m:d><m:dPr>${val('begChr',n.beg)}${n.es.length>1?val('sepChr',n.sep):''}${val('endChr',n.end)}</m:dPr>${n.es.map(x=>`<m:e>${emit(x,ph)}</m:e>`).join('')}</m:d>`;break;
      case 'func':out+=`<m:func><m:fName>${emit(n.name,ph)}</m:fName>${arg('e',n.e)}</m:func>`;break;
      case 'fn':out+=runXML(n.name,'p');break;
      case 'limlow':out+=`<m:limLow>${arg('e',n.e)}${arg('lim',n.lim)}</m:limLow>`;break;
      case 'limupp':out+=`<m:limUpp>${arg('e',n.e)}${arg('lim',n.lim)}</m:limUpp>`;break;
      case 'acc':out+=`<m:acc><m:accPr>${val('chr',n.chr)}</m:accPr>${arg('e',n.e)}</m:acc>`;break;
      case 'bar':out+=`<m:bar><m:barPr>${val('pos',n.pos)}</m:barPr>${arg('e',n.e)}</m:bar>`;break;
      case 'gchr':out+=`<m:groupChr><m:groupChrPr>${val('chr',n.chr)}${val('pos',n.pos)}</m:groupChrPr>${arg('e',n.e)}</m:groupChr>`;break;
      case 'box':out+=`<m:borderBox>${arg('e',n.e)}</m:borderBox>`;break;
      case 'phant':out+=`<m:phant><m:phantPr>${val('show','0')}</m:phantPr>${arg('e',n.e)}</m:phant>`;break;
      case 'eqarr':out+=`<m:eqArr>${(n.rows.length?n.rows:[[]]).map(r=>`<m:e>${emit(r,ph)}</m:e>`).join('')}</m:eqArr>`;break;
      case 'mat':out+=`<m:m>${n.rows.map(r=>`<m:mr>${r.map(c=>`<m:e>${emit(c,ph)}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`;break;
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
  const rPr=kid(r,'rPr');let t='';
  for(const x of r.childNodes)if(x.nodeType===1&&x.localName==='t')t+=x.textContent;
  if(!t)return '';
  if(rPr&&on(kid(rPr,'nor')))return '\\text{'+t.replace(/[\\{}$&%#_^~]/g,m=>'\\'+m)+'}';
  const sty=rPr?mattr(kid(rPr,'sty'),'val'):null,scr=rPr?mattr(kid(rPr,'scr'),'val'):null;
  const styCmd=scr&&SCRCMD[scr]?SCRCMD[scr]:(sty==='p'?'mathrm':sty==='b'?'mathbf':sty==='bi'?'boldsymbol':null);
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
function oseq(el,c){let s='';for(const n of kids(el))s=cat(s,onode(n,c));return s;}
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
  const k={inEq:true,lossy:false};const rows=kids(a,'e').map(e=>oseq(e,k)).join(' \\\\ ');
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
function omml2tex(el){
  const c={lossy:false,inEq:false};let tex,display=false;
  if(el.localName==='oMathPara'){display=true;tex=kids(el,'oMath').map(m=>oseq(m,c)).join(' \\\\ ');}
  else tex=oseq(el,c);
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
const sy=n=>[SYM[n],'\\'+n+' ',n];
const PAL={
  s:['Struktur',[
    ['a/b','\\frac{'+CM+'}{}','Pecahan'],['√','\\sqrt{'+CM+'}','Akar kuadrat'],['ⁿ√','\\sqrt['+CM+']{}','Akar pangkat n'],
    ['xⁿ','^{'+CM+'}','Pangkat'],['xₙ','_{'+CM+'}','Indeks'],['xₙⁿ','_{'+CM+'}^{}','Indeks dan pangkat'],
    ['∫','\\int_{'+CM+'}^{}','Integral'],['∬','\\iint_{'+CM+'}','Integral ganda'],['∮','\\oint_{'+CM+'}','Integral kontur'],
    ['∑','\\sum_{'+CM+'}^{}','Sigma (jumlah)'],['∏','\\prod_{'+CM+'}^{}','Perkalian'],['lim','\\lim_{'+CM+'\\to }','Limit'],
    ['( )','\\left('+CM+'\\right)','Kurung biasa'],['[ ]','\\left['+CM+'\\right]','Kurung siku'],['{ }','\\left\\{'+CM+'\\right\\}','Kurung kurawal'],
    ['| |','\\left|'+CM+'\\right|','Nilai mutlak'],['⟨ ⟩','\\left\\langle '+CM+'\\right\\rangle','Kurung sudut'],['(n k)','\\binom{'+CM+'}{}','Binomial'],
    ['x̂','\\hat{'+CM+'}','Topi'],['x̄','\\overline{'+CM+'}','Garis atas'],['x⃗','\\vec{'+CM+'}','Vektor'],['ẋ','\\dot{'+CM+'}','Titik'],['x̃','\\tilde{'+CM+'}','Tilde'],
    ['⏟','\\underbrace{'+CM+'}_{}','Kurung bawah'],['⏞','\\overbrace{'+CM+'}^{}','Kurung atas'],
    ['sin','\\sin '+CM,'Sinus'],['cos','\\cos '+CM,'Cosinus'],['tan','\\tan '+CM,'Tangen'],['log','\\log '+CM,'Logaritma'],['ln','\\ln '+CM,'Logaritma natural'],
    ['matriks','\\begin{pmatrix} '+CM+' &  \\\\  &  \\end{pmatrix}','Matriks 2×2'],['kasus','\\begin{cases} '+CM+' &  \\\\  &  \\end{cases}','Kasus / sistem persamaan'],
    ['rata =','\\begin{aligned} '+CM+' &=  \\\\ &=  \\end{aligned}','Persamaan rata sama dengan']
  ]],
  g:['Yunani',('alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota kappa lambda mu nu xi pi rho sigma tau upsilon phi varphi chi psi omega Gamma Delta Theta Lambda Xi Pi Sigma Phi Psi Omega').split(' ').map(sy)],
  o:['Operator',('pm mp times div cdot ast circ oplus otimes wedge vee cap cup setminus neg partial nabla infty prime degree').split(' ').map(sy)],
  r:['Relasi',[['=','=','sama dengan'],['<','<','kurang dari'],['>','>','lebih dari']].concat(('neq approx equiv sim cong propto leq geq ll gg in notin subset supset subseteq supseteq forall exists emptyset perp parallel mid').split(' ').map(sy))],
  a:['Panah',('to leftarrow leftrightarrow Rightarrow Leftarrow Leftrightarrow mapsto uparrow downarrow implies iff').split(' ').map(sy)],
  l:['Lainnya',[['ℝ','\\mathbb{R}','Bilangan real'],['ℕ','\\mathbb{N}','Bilangan asli'],['ℤ','\\mathbb{Z}','Bilangan bulat'],['ℚ','\\mathbb{Q}','Bilangan rasional'],['ℂ','\\mathbb{C}','Bilangan kompleks'],
    ['abc','\\mathrm{'+CM+'}','Teks tegak'],['teks','\\text{'+CM+'}','Teks biasa'],['𝐱','\\mathbf{'+CM+'}','Tebal'],['…','\\ldots ','Titik tiga'],['⋯','\\cdots ','Titik tiga tengah'],['⋮','\\vdots ','Titik vertikal'],
    ['ℏ','\\hbar ','h-bar'],['ℓ','\\ell ','l kecil'],['∠','\\angle ','Sudut'],['∴','\\therefore ','Jadi'],['∵','\\because ','Karena'],['spasi','\\quad ','Spasi lebar'],[',','\\, ','Spasi tipis']].concat([]) ]
};
/* rumus bawaan untuk galeri di pita */
const GALLERY=[
  ['circ','Luas lingkaran','A=\\pi r^{2}'],
  ['pyth','Teorema Pythagoras','a^{2}+b^{2}=c^{2}'],
  ['abc','Rumus kuadrat (ABC)','x=\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}'],
  ['binom','Teorema binomial','(x+a)^{n}=\\sum_{k=0}^{n}\\binom{n}{k}x^{k}a^{n-k}'],
  ['taylor','Deret Taylor','e^{x}=1+\\frac{x}{1!}+\\frac{x^{2}}{2!}+\\frac{x^{3}}{3!}+\\cdots'],
  ['trig','Identitas trigonometri','\\sin\\alpha\\pm\\sin\\beta=2\\sin\\frac{1}{2}(\\alpha\\pm\\beta)\\cos\\frac{1}{2}(\\alpha\\mp\\beta)'],
  ['fourier','Deret Fourier','f(x)=a_{0}+\\sum_{n=1}^{\\infty}\\left(a_{n}\\cos\\frac{n\\pi x}{L}+b_{n}\\sin\\frac{n\\pi x}{L}\\right)'],
  ['deriv','Definisi turunan',"f'(x)=\\lim_{h\\to 0}\\frac{f(x+h)-f(x)}{h}"],
  ['integ','Teorema dasar kalkulus','\\int_{a}^{b}f(x)\\,dx=F(b)-F(a)'],
  ['normal','Distribusi normal','f(x)=\\frac{1}{\\sigma\\sqrt{2\\pi}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}'],
  ['mean','Rata-rata','\\bar{x}=\\frac{1}{n}\\sum_{i=1}^{n}x_{i}']
];

let cssDone=false;
function ensureCSS(){
  if(cssDone||typeof document==='undefined')return;cssDone=true;
  const st=document.createElement('style');
  st.textContent=`
.wd .eqdg{width:min(520px,94%);max-width:94%;padding:12px 14px}
.wd .eqdg .eqpv{min-height:56px;max-height:26vh;overflow:auto;border:1px solid #c8c8c8;background:#fafafa;display:grid;place-items:center;padding:6px 8px;margin-bottom:8px;color:#222}
.wd .eqdg .eqpv math{font-family:"Cambria Math","STIX Two Math","Noto Sans Math","Latin Modern Math",math;font-size:1.45em}
.wd .eqdg .eqpv .e{color:#a4262c;font-size:12px;text-align:left;white-space:pre-wrap}
.wd .eqdg .eqpv .h{color:#999;font-size:12px}
.wd .eqdg textarea{width:100%;font:15px/1.35 Consolas,"Courier New",monospace;border:1px solid #aaa;padding:5px 6px;resize:none;min-height:48px;display:block}
.wd .eqdg .eqnote{font-size:11px;color:#8a6d00;background:#fff8e1;border:1px solid #f0dc9a;padding:4px 6px;margin-top:6px}
.wd .eqdg .eqtabs{display:flex;gap:2px;margin:8px 0 4px;overflow-x:auto;scrollbar-width:none}
.wd .eqdg .eqtabs button{flex:none;height:28px;padding:0 10px;border:1px solid #c8c8c8;border-bottom-color:#c8c8c8;background:#f3f3f3;cursor:pointer}
.wd .eqdg .eqtabs button.on{background:#fff;color:var(--blue);border-color:var(--blue);font-weight:600}
.wd .eqdg .eqpal{display:flex;flex-wrap:wrap;gap:3px;max-height:128px;overflow:auto;padding:2px;border:1px solid #e1e1e1;background:#fff}
.wd .eqdg .eqpal button{min-width:36px;height:36px;padding:0 5px;border:1px solid #d4d4d4;background:#fbfbfb;cursor:pointer;font-size:14px;display:grid;place-items:center}
.wd .eqdg .eqpal button:hover,.wd .eqdg .eqpal button:active{background:#dbe7f7;border-color:var(--blue)}
.wd .eqdg .eqpal math{font-size:15px;font-family:"Cambria Math","STIX Two Math","Noto Sans Math",math}
.wd .eqdg .eqmode{display:flex;gap:14px;margin-top:8px;font-size:12px}
.wd .eqdg .eqmode label{display:flex;align-items:center;gap:5px;cursor:pointer}
.wd .eqdg .eqmode input{width:auto;height:auto}
.wd .page .eq.eqsel{outline:2px solid var(--blue);outline-offset:2px;background:rgba(43,87,154,.08)}
.wd .eqbar{position:absolute;z-index:25;display:flex;background:#fff;border:1px solid #8a8a8a;box-shadow:0 2px 8px rgba(0,0,0,.28)}
.wd .eqbar button{border:0;background:none;height:34px;padding:0 12px;cursor:pointer;font-size:13px;white-space:nowrap}
.wd .eqbar button+button{border-left:1px solid #ddd}
.wd .eqbar button:active,.wd .eqbar button:hover{background:#dbe7f7}
.wd .eqbar button[data-q=del]{color:#a4262c}`;
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

/* o: {host,title,okText,tex,display,note,onOk({tex,display,omml})->false untuk tetap terbuka} */
function openDialog(o){
  ensureCSS();
  const ov=document.createElement('div');ov.className='ov';
  ov.innerHTML=`<div class="dg eqdg"><h3>${xe(o.title||'Rumus')}</h3>
<div class="eqpv"></div>
<textarea class="eqtx" rows="2" spellcheck="false" autocapitalize="off" autocomplete="off" autocorrect="off" placeholder="Ketik rumus (LaTeX), mis. \\frac{a}{b}"></textarea>
${o.note?`<div class="eqnote">${xe(o.note)}</div>`:''}
<div class="eqtabs">${Object.keys(PAL).map((k,i)=>`<button type="button" data-g="${k}"${i?'':' class="on"'}>${PAL[k][0]}</button>`).join('')}</div>
<div class="eqpal"></div>
<div class="eqmode"><label><input type="radio" name="eqm" value="i"> Sebaris (dalam teks)</label><label><input type="radio" name="eqm" value="b"> Blok (di tengah)</label></div>
<div class="dbt"><button type="button" class="x">Batal</button><button type="button" class="ok">${xe(o.okText||'Sisipkan')}</button></div></div>`;
  o.host.appendChild(ov);
  const $=s=>ov.querySelector(s),ta=$('.eqtx'),pv=$('.eqpv'),pal=$('.eqpal'),okb=$('.ok');
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
  function renderPal(k){
    pal.innerHTML=PAL[k][1].map((it,i)=>{
      let g=xe(it[0]);
      if(k==='s'&&mm){try{g=mathHTML(tex2omml(it[1].replace(CM,''),{ph:true}),false);}catch(e){}}
      return `<button type="button" data-i="${i}" title="${xa(it[2]||'')}">${g}</button>`;}).join('');
    pal._k=k;
  }
  function insert(tpl){
    const s=ta.selectionStart,e=ta.selectionEnd,sel=ta.value.slice(s,e);
    let at=tpl.indexOf(CM);const has=at>=0;
    let t=has?tpl.replace(CM,sel):tpl;
    if(!has&&/\\[A-Za-z]+$/.test(t)&&!/\s$/.test(t))t+=' ';
    ta.setRangeText(t,s,e,'end');
    const pos=has?s+at+sel.length:s+t.length;
    ta.setSelectionRange(pos,pos);ta.focus({preventScroll:true});later();
  }
  renderPal('s');
  ov.addEventListener('mousedown',e=>{if(e.target.closest('.eqpal button,.eqtabs button'))e.preventDefault();});
  pal.addEventListener('click',e=>{const b=e.target.closest('button[data-i]');if(b)insert(PAL[pal._k][1][+b.dataset.i][1]);});
  $('.eqtabs').addEventListener('click',e=>{const b=e.target.closest('button[data-g]');if(!b)return;
    ov.querySelectorAll('.eqtabs button').forEach(x=>x.classList.toggle('on',x===b));renderPal(b.dataset.g);});
  ta.addEventListener('input',later);
  ov.querySelectorAll('input[name=eqm]').forEach(r=>r.addEventListener('change',update));
  const close=()=>{clearTimeout(tm);ov.remove();};
  $('.x').onclick=close;
  okb.onclick=()=>{
    update();if(!good)return;
    const d=display(),om=tex2omml(ta.value,{display:d});
    if(o.onOk({tex:ta.value,display:isPara(om),omml:om})!==false)close();
  };
  ov.addEventListener('keydown',e=>{
    if(e.key==='Escape'){e.stopPropagation();close();}
    else if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();okb.click();}
  });
  ov.addEventListener('mousedown',e=>{if(e.target===ov)e.stopPropagation();});
  update();ta.focus({preventScroll:true});ta.setSelectionRange(ta.value.length,ta.value.length);
  return {close,ta,update};
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
    const el=page.querySelector('[data-eqnew]');if(!el)return;
    el.removeAttribute('data-eqnew');
    const t=document.createElement('div');t.innerHTML=eqHTML(omml,display);
    el.innerHTML=t.firstChild?t.firstChild.innerHTML:'';
  }
  const put=(omml,display)=>{document.execCommand('insertHTML',false,shell(omml,display,!display));fill(omml,display);};
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
      document.execCommand('delete');insertNew(omml,true);return;
    }
    put(omml,display);
    refresh();stats();
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
      try{const r=omml2tex(parseXML(orig));tex=r.tex;display=r.display;lossy=r.lossy;}
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
  function showBar(eq){
    hide();armed=eq;eq.classList.add('eqsel');
    bar=document.createElement('div');bar.className='eqbar';
    bar.innerHTML=`<button type="button" data-q="edit">✎ Edit</button><button type="button" data-q="mode">${eq.classList.contains('eqd')?'Sebaris':'Blok'}</button><button type="button" data-q="del">Hapus</button>`;
    host.appendChild(bar);
    const hr=host.getBoundingClientRect(),er=eq.getBoundingClientRect(),bw=bar.offsetWidth,bh=bar.offsetHeight;
    let L=er.left-hr.left,T=er.top-hr.top-bh-6;
    if(T<30)T=er.bottom-hr.top+6;
    L=Math.max(4,Math.min(L,hr.width-bw-4));
    bar.style.left=L+'px';bar.style.top=T+'px';
    bar.addEventListener('mousedown',e=>e.preventDefault());
    bar.addEventListener('click',e=>{
      const b=e.target.closest('button[data-q]');if(!b)return;const q=b.dataset.q,t=armed;if(!t)return;
      if(q==='edit'){hide();dialog(t);}else if(q==='mode')toggleMode(t);else if(q==='del')remove(t);
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

root.WordEq={tex2omml,omml2tex,setDisplay,openDialog,attach,GALLERY,EqError,_parse:parseTex};
})(typeof window!=='undefined'?window:globalThis);
