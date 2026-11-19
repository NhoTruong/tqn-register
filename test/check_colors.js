// Rà màu trùng/gần nhau trong cùng một nhóm (tình trạng; bên). Khoảng cách CIE76 trên Lab.
const fs=require('fs'),vm=require('vm');const ctx={tqnNoAccent_:s=>s,tqnClean_:s=>s,SpreadsheetApp:{}};vm.createContext(ctx);
vm.runInContext(fs.readFileSync(__dirname+'/../apps-script/Theme.gs','utf8')+';this.O=TQN_OPT_COLORS;this.P=TQN_PARTY;this.TH=TQN_TH;',ctx);
const lab=h=>{let[r,g,b]=[1,3,5].map(i=>parseInt(h.slice(i,i+2),16)/255).map(v=>v>.04045?((v+.055)/1.055)**2.4:v/12.92);let x=(r*.4124+g*.3576+b*.1805)/.95047,y=r*.2126+g*.7152+b*.0722,z=(r*.0193+g*.1192+b*.9505)/1.08883;const f=t=>t>.008856?Math.cbrt(t):7.787*t+16/116;return[116*f(y)-16,500*(f(x)-f(y)),200*(f(y)-f(z))]};
const dE=(a,b)=>{const A=lab(a),B=lab(b);return Math.hypot(A[0]-B[0],A[1]-B[1],A[2]-B[2])};
const groups={'Tình trạng':ctx.O.filter(x=>!/tvgs|cdt|tvtk|tbcons|nha thau/.test(x[0])),'Bên':ctx.O.filter(x=>/tvgs|cdt|tvtk|tbcons/.test(x[0]))};
let bad=0;
for(const[g,L]of Object.entries(groups)){const u={};L.forEach(x=>{const k=x[1]+x[2];(u[k]=u[k]||[]).push(x[0])});const K=Object.keys(u);
 for(let i=0;i<K.length;i++)for(let j=i+1;j<K.length;j++){const d=dE(K[i].slice(0,7),K[j].slice(0,7))+dE(K[i].slice(7),K[j].slice(7));if(d<18){bad++;console.log('GẦN '+g+': '+u[K[i]].join('/')+' ~ '+u[K[j]].join('/')+' (ΔE '+d.toFixed(1)+')')}}}
const hp=[...new Set(ctx.P.map(p=>p[2]))];for(let i=0;i<hp.length;i++)for(let j=i+1;j<hp.length;j++){const d=dE(hp[i],hp[j]);if(d<15){bad++;console.log('GẦN tiêu đề bên: '+hp[i]+' ~ '+hp[j]+' ΔE '+d.toFixed(1))}}
[hp.concat([ctx.TH.head,ctx.TH.band])].forEach(L=>{});
console.log(bad?('CÒN '+bad+' cặp gần nhau'):'PASS màu: không có cặp trùng/gần nhau trong cùng nhóm');
