/* TKR Attendance — installable app for TKR Operations (Tharisa Mine).
   Data is shared through Firebase (project tkr-operations-app). Marks made with no signal stay on the device and are sent when the signal is back. */
import {initializeApp,deleteApp,getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged,createUserWithEmailAndPassword,sendPasswordResetEmail,
  initializeFirestore,persistentLocalCache,persistentMultipleTabManager,collection,doc,setDoc,getDoc,getDocs,onSnapshot,arrayUnion,arrayRemove} from "../lib/firebase.js";
const toLogin = v => { v = String(v || "").trim(); return v.includes("@") ? v : v.toLowerCase().replace(/[^a-z0-9]/g, "") + "@mgr.tkroperations.app"; };  // managers can sign in with a username
const firebaseConfig={apiKey:"AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI",authDomain:"tkr-operations-app.firebaseapp.com",projectId:"tkr-operations-app",
  storageBucket:"tkr-operations-app.firebasestorage.app",messagingSenderId:"158080247079",appId:"1:158080247079:web:7cb7786d2b7f9ee172b502"};
// Must match the Firestore rules (firestore-rules.txt).
const ADMINS=["muneer@tkroperations.com","amogelang@tkroperations.com","amogelangmabilu56@gmail.com"];
const fbApp=initializeApp(firebaseConfig);const auth=getAuth(fbApp);
let fs;try{fs=initializeFirestore(fbApp,{localCache:persistentLocalCache({tabManager:persistentMultipleTabManager()})})}catch(e){fs=initializeFirestore(fbApp,{})}
let email="",roles={capturers:[],viewers:[]};
const isAdmin=()=>ADMINS.includes(email);
const canWrite=()=>isAdmin()||roles.capturers.includes(email);
let mgrs=[];
const canView=()=>canWrite()||roles.viewers.includes(email)||mgrs.includes(email);
// Store paths used by the register, kept apart from the other TKR apps in the same Firebase project.
const PATHS=[["config/employees","config/attEmployees"],["config/settings","config/attSettings"],["cycles","attRota"],["months","attMonths"]];
const P=path=>{for(const [a,b] of PATHS)if(path===a||path.startsWith(a+"/"))return b+path.slice(a.length);return path};
const wDoc=d=>({id:d.id,exists:d.exists(),data:()=>d.data()});
const wQ=q=>({docs:q.docs.map(d=>({id:d.id,exists:true,data:()=>d.data()})),size:q.size,empty:q.empty});
const fdb={doc(path){const r=doc(fs,P(path));return{get:()=>getDoc(r).then(wDoc),set:d=>setDoc(r,d),update:d=>setDoc(r,d,{merge:true}),onSnapshot:(n,e)=>onSnapshot(r,s=>n(wDoc(s)),e)}},
  collection(path){const r=collection(fs,P(path));return{get:()=>getDocs(r).then(wQ),onSnapshot:(n,e)=>onSnapshot(r,s=>n(wQ(s)),e)}}};
let startApp=()=>{},rerender=()=>{};
(function(){
const SHEETS=["Rompad","Support Fleet","Drilling","Shared Services","Interns"];
const SHEET_ID={"Rompad":"rompad","Support Fleet":"support-fleet","Drilling":"drilling","Shared Services":"shared-services","Interns":"interns"};
// Attendance rules set by Muneer
const KIND={P:"present",MEDICAL:"present","SITE OFFICE":"present","HEAD OFFICE":"present",TRAINING:"present","SOOSAN TRAINING":"present","PIT LICENCE":"present",LATE:"present",INDUCTION:"present",
  AWOP:"absent",
  "ANNUAL LEAVE":"leave",SICK:"leave","MATERNITY LEAVE":"leave","STUDY LEAVE":"leave","FAMILY RESPONSIBILITY":"leave","FAMILY EMERGENCY":"leave",
  OFF:"off",RESIGNED:"out",ABSCONDED:"out",DISMISSED:"out",DECEASED:"out",SUSPENDED:"out"};
const CODE_ORDER=["P","LATE","MEDICAL","SITE OFFICE","HEAD OFFICE","TRAINING","SOOSAN TRAINING","PIT LICENCE","INDUCTION","AWOP","SICK","ANNUAL LEAVE","FAMILY RESPONSIBILITY","FAMILY EMERGENCY","STUDY LEAVE","MATERNITY LEAVE","OFF","SUSPENDED","RESIGNED","ABSCONDED","DISMISSED","DECEASED"];
const isTime=v=>/^\d{2}:\d{2}$/.test(v||"");
const isHrs=v=>/^\d+(\.\d+)? HRS$/.test(v||"")||/^\d{1,2}:\d{2} HRS$/.test(v||"");
const hrsOf=v=>{if(!isHrs(v))return 0;const m=/^(\d{1,2}):(\d{2})/.exec(v);return m?Number(m[1])+Number(m[2])/60:parseFloat(v)};
const hm=h=>{const t=Math.round((Number(h)||0)*60);return String(Math.floor(t/60)).padStart(2,"0")+":"+String(t%60).padStart(2,"0")};
const showMark=v=>isHrs(v)?hm(hrsOf(v)):v;
const parseHm=x=>{const m=/^\s*(\d{1,2})(?:[:h.,](\d{1,2}))?\s*$/.exec(x||"");if(!m)return NaN;const mm=m[2]==null?0:Number(m[2].length===1?m[2]+"0":m[2]);if(mm>59)return NaN;return Number(m[1])+mm/60};
const kindOf=v=>!v?"blank":(isTime(v)||isHrs(v))?"present":(KIND[v]||"present");
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pad=n=>String(n).padStart(2,"0");
const iso=d=>d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate());
const parse=s=>{const [y,m,d]=s.split("-").map(Number);return new Date(y,m-1,d)};
const addDays=(s,n)=>{const d=parse(s);d.setDate(d.getDate()+n);return iso(d)};
const daysBetween=(a,b)=>Math.round((parse(b)-parse(a))/864e5);
const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],DOW=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
const nice=s=>{const d=parse(s);return DOW[d.getDay()]+" "+d.getDate()+" "+MON[d.getMonth()]};
const short=s=>{const d=parse(s);return d.getDate()+" "+MON[d.getMonth()]};
// Cycles: the ones Muneer sets himself (config/settings.cycles) win; other dates use the standard 14th–13th, trimmed so they never overlap a set cycle
let CUSTOM=[];const _cy={};
function stdCycle(s){const d=parse(s);let y=d.getFullYear(),m=d.getMonth();if(d.getDate()<14){m--;if(m<0){m=11;y--}}return [iso(new Date(y,m,14)),iso(new Date(y,m+1,13))]}
function cycleOf(s){if(_cy[s])return _cy[s];const k=CUSTOM.find(c=>c.start<=s&&s<=c.end);let start,end;
  if(k){start=k.start;end=k.end}else{[start,end]=stdCycle(s);CUSTOM.forEach(c=>{if(c.end<s&&c.end>=start)start=addDays(c.end,1);if(c.start>s&&c.start<=end)end=addDays(c.start,-1)})}
  const dates=[];for(let x=start;x<=end;x=addDays(x,1))dates.push(x);return _cy[s]={start,end,dates,custom:!!k}}
function setCustom(list){CUSTOM=(list||[]).slice().sort((a,b)=>a.start<b.start?-1:1);for(const k in _cy)delete _cy[k]}
const pct=(p,a)=>p+a?Math.round(1000*p/(p+a))/10:null;
const fmtPct=v=>v==null?"–":v.toFixed(1)+"%";

const S={flg:{kind:"all",type:"cycle",from:null,to:null,key:""},mm:{},mExists:{},rota:{},gmeta:{},exp:{type:"daily",from:null,to:null,key:""},view:"dash",date:iso(new Date()),sheet:"All",q:"",emps:null,empVer:null,cycles:{},marks:{},markExists:{},db:null,canWrite:true,sortLow:false,unsubs:[]};
try{const sh=localStorage.getItem("tkr-att-sheet");if(sh)S.sheet=sh}catch(e){}
const $=id=>document.getElementById(id);
const main=$("main");

function toast(t){const el=document.createElement("div");el.className="toast";el.textContent=t;document.body.appendChild(el);setTimeout(()=>el.remove(),2600)}

// ---------- data ----------
// Marks are stored per calendar month (months/YYYY-MM/sheets/<sheet>), so any cycle the user sets finds its days.
const monthOf=d=>d.slice(0,7);
function monthsIn(from,to){const out=[];let m=monthOf(from);const last=monthOf(to);while(m<=last){out.push(m);let [y,mm]=m.split("-").map(Number);mm++;if(mm>12){mm=1;y++}m=y+"-"+pad(mm)}return out}
const live={};
function applyMonth(m,snap){const o={};snap.docs.forEach(d=>{S.mExists[m+"/"+d.id]=true;o[d.id]={marks:d.data().marks||{},ot:d.data().ot||{},ottype:d.data().ottype||{}}});S.mm[m]=o}
function watchCycle(){const c=cycleOf(S.date);monthsIn(c.start,c.end).forEach(m=>{if(live[m])return;
  live[m]=S.db.collection("months/"+m+"/sheets").onSnapshot(snap=>{applyMonth(m,snap);render()},err)})}
async function loadRange(from,to){await Promise.all(monthsIn(from,to).filter(m=>!live[m]).map(async m=>{const snap=await S.db.collection("months/"+m+"/sheets").get();applyMonth(m,snap)}))}
function err(e){const perm=e&&(e.code==="permission-denied"||/permission/i.test(e.message||""));main.innerHTML='<div class="empty"><b>'+(perm?"Firebase is blocking the register":"The register could not load")+'</b>'+esc(perm?"Your sign-in is not allowed to use the register yet. In the Firebase console open tkr-operations-app, then Firestore Database, then Rules, add the TKR Attendance rules and tap Publish. Then close and open the app again.":(e&&e.message||"Check your signal and open the app again."))+'</div>'}
function denied(e){return e&&(e.code==="invalid_argument"||e.code==="permission-denied")}
function writeSheet(sheet,patch,field){ // patch: {empId:{date:val}} for dates in S.date's month; field: "marks" or "ot"
  field=field||"marks";const id=SHEET_ID[sheet],m=monthOf(S.date);
  const mo=S.mm[m]=S.mm[m]||{};const sd=mo[id]=mo[id]||{marks:{},ot:{}};const store=sd[field];for(const k in patch){store[k]=Object.assign({},store[k]||{},patch[k])}
  S.db.doc("months/"+m+"/sheets/"+id).update({sheet,month:m,[field]:patch,updatedBy:email,updatedAt:Date.now()})
    .catch(e=>{if(denied(e)){S.canWrite=false;toast("You can view this register but not mark it. Ask admin for capture access.")}else toast("Not saved: "+(e&&e.message||"try again"));render()});
  render();
}
const cell=(e,d,f)=>((((S.mm[monthOf(d)]||{})[SHEET_ID[e.sheet]]||{})[f]||{})[e.id]||{})[d];
const rawMark=(e,d)=>cell(e,d,"marks")||"";
const markOf=(e,d)=>{const v=rawMark(e,d);if(!v&&e.leftDate&&d>e.leftDate)return e.leftReason||"RESIGNED";return v};
const otOf=(e,d)=>Number(cell(e,d,"ot"))||0;
const fmtH=h=>String(Math.round(h*10)/10);
function setOt(e,h){if(!S.canWrite){toast("You can view this register but not mark it.");return}writeSheet(e.sheet,{[e.id]:{[S.date]:h}},"ot")}
function setMark(e,val){if(!S.canWrite){toast("You can view this register but not mark it.");return}writeSheet(e.sheet,{[e.id]:{[S.date]:val}})}

// Shift rotas: every cycles/<id> doc holds group rotas by date; they are merged into one lookup.
function applyRotas(){S.rota={};S.gmeta={};Object.keys(S.cycles).sort().forEach(id=>{const g=(S.cycles[id]||{}).groups||{};
  for(const gid in g){S.gmeta[gid]={sheet:g[gid].sheet,section:g[gid].section,shift:g[gid].shift};S.rota[gid]=Object.assign(S.rota[gid]||{},g[gid].rota||{})}})}
function rotaFor(gid,d){return (S.rota[gid]||{})[d]||""}
function rotaGap(){const gs=Object.keys(S.rota);if(!gs.length)return 0;return cycleOf(S.date).dates.filter(d=>!gs.some(g=>S.rota[g][d])).length}
function projectRota(dates){const groups={};for(const gid in S.rota){const r=S.rota[gid];const known=Object.keys(r).filter(d=>r[d]).sort();if(known.length<4)continue;
  const base=known[0];const seq=rangeDates(base,known[known.length-1]).map(d=>r[d]||"");let per=0;
  for(let p=2;p<=seq.length/2;p++){let ok=true;for(let i=0;i+p<seq.length;i++)if(seq[i]&&seq[i+p]&&seq[i]!==seq[i+p]){ok=false;break}if(ok){per=p;break}}
  if(!per)continue;const rota={};dates.forEach(d=>{if(r[d])return;const off=((daysBetween(base,d)%per)+per)%per;let v="";for(let i=off;i<seq.length&&!v;i+=per)v=seq[i];if(v)rota[d]=v});
  if(Object.keys(rota).length)groups[gid]=Object.assign({},S.gmeta[gid],{rota})}return groups}
async function startCycle(){const c=cycleOf(S.date);const groups=projectRota(c.dates);if(!Object.keys(groups).length){toast("There is no shift pattern to carry forward");return}
  const id="r-"+c.start;const ref=S.db.doc("cycles/"+id);
  try{if(S.cycles[id])await ref.update({groups});else await ref.set({start:c.start,end:c.end,groups});toast("Shift rotas set for "+short(c.start)+" – "+short(c.end))}catch(e){toast("Not saved: "+(e.message||"try again"))}}

// Cycle dates panel
function openCyclePanel(){const c=cycleOf(S.date);
  const list=CUSTOM.length?'<ul>'+CUSTOM.map(k=>'<li>'+nice(k.start)+' '+yr(k.start)+' to '+nice(k.end)+' '+yr(k.end)+'</li>').join("")+'</ul>':'<p class="note" style="margin:0">None yet. All cycles run from the 14th to the 13th.</p>';
  $("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true" aria-labelledby="ct"><h2 id="ct">Cycle dates</h2><div class="sub">Set the start and end of the cycle that includes '+nice(S.date)+' '+yr(S.date)+'.</div>'+
   '<div class="clock"><label class="lbl" for="cy-from" style="margin:0">Start</label><input type="date" id="cy-from" value="'+c.start+'"><label class="lbl" for="cy-to" style="margin:0">End</label><input type="date" id="cy-to" value="'+c.end+'"></div>'+
   '<p class="note">Dates outside the cycles you set follow the standard 14th to 13th.</p>'+
   '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px"><button class="btn gold" id="cySave">Save cycle</button>'+(c.custom?'<button class="btn" id="cyReset">Use 14th to 13th</button>':"")+'<button class="btn" id="closeP">Cancel</button></div>'+
   '<div class="lbl">Cycles you have set</div>'+list+'</div></div>';
  const close=()=>{$("sheetHost").innerHTML=""};$("closeP").onclick=close;$("scrim").onclick=ev=>{if(ev.target.id==="scrim")close()};
  const others=CUSTOM.filter(k=>!(c.custom&&k.start===c.start&&k.end===c.end));
  const save=list=>{if(!S.canWrite){toast("You can view this register but not change it.");return}
    S.db.doc("config/settings").set({cycles:list}).then(()=>{setCustom(list);watchCycle();render();toast("Cycle saved")},e=>toast("Not saved: "+(e.message||"try again")));close()};
  $("cySave").onclick=()=>{const a=$("cy-from").value,b=$("cy-to").value;if(!a||!b){toast("Choose a start and an end date");return}
    if(b<a){toast("The end date must be after the start date");return}if(daysBetween(a,b)>92){toast("A cycle can be at most 93 days");return}
    const clash=others.find(k=>k.start<=b&&a<=k.end);if(clash){toast("These dates overlap the cycle "+short(clash.start)+" – "+short(clash.end)+". Change that cycle first.");return}
    save(others.concat([{start:a,end:b}]))};
  if(c.custom)$("cyReset").onclick=()=>save(others);
}

// ---------- helpers ----------
function activeEmps(){return (S.emps||[]).filter(e=>e.active!==false)}
function inPeriod(from){return activeEmps()}
function filtered(from){const q=S.q.trim().toLowerCase();return (from?inPeriod(from):activeEmps()).filter(e=>(S.sheet==="All"||e.sheet===S.sheet)&&(!q||(e.name+" "+e.number+" "+e.title).toLowerCase().includes(q)))}
function groupLabel(e){const parts=[];if(e.section&&e.section!==e.sheet)parts.push(e.section);if(e.shift)parts.push(e.shift);return parts.join(" · ")||e.sheet}
function bySheetGroup(list){const out=[];let cur=null,g=null;
  SHEETS.forEach(sh=>{const es=list.filter(e=>e.sheet===sh).sort((a,b)=>a.order-b.order);if(!es.length)return;cur={sheet:sh,groups:[]};out.push(cur);g=null;
    es.forEach(e=>{if(!g||g.id!==e.group){g={id:e.group,label:groupLabel(e),emps:[]};cur.groups.push(g)}g.emps.push(e)})});return out}
function pill(v){const k=kindOf(v);return '<span class="pill s-'+k+((isTime(v)||isHrs(v))?" time":"")+'">'+(v?esc(showMark(v)):"Mark")+'</span>'}
// Leave that does not lower attendance: annual, family responsibility, family emergency and study leave.
// AWOP, sick and maternity leave lower it (t.exc).
const NOT_ABSENT=["ANNUAL LEAVE","FAMILY RESPONSIBILITY","FAMILY EMERGENCY","STUDY LEAVE"];
function tally(es,dates){const t={present:0,absent:0,leave:0,off:0,out:0,blank:0,ot:0,exc:0,hrs:0,hrsDays:0,codes:{}};
  es.forEach(e=>dates.forEach(d=>{const v=markOf(e,d);const k=kindOf(v);t[k]++;if(k==="absent"||(k==="leave"&&!NOT_ABSENT.includes(v)))t.exc++;t.ot+=otOf(e,d);if(isHrs(v)){t.hrs+=hrsOf(v);t.hrsDays++}const code=isTime(v)?"CLOCKED IN":isHrs(v)?"HOURS WORKED":v;if(v)t.codes[code]=(t.codes[code]||0)+1}));return t}
function chipsBar(extra){return '<div class="bar" role="group" aria-label="Sheet">'+["All"].concat(SHEETS).map(s=>'<button class="chip" data-sheet="'+esc(s)+'" aria-pressed="'+(S.sheet===s)+'">'+esc(s)+'</button>').join("")+(extra||"")+'</div>'}

// ---------- views ----------
function render(){
  const c=cycleOf(S.date);$("date").value=S.date;$("cycleLbl").innerHTML="Cycle "+short(c.start)+" – "+short(c.end)+" "+parse(c.end).getFullYear()+" · <b>Change</b>";
  document.querySelectorAll("#tabs button").forEach(b=>b.setAttribute("aria-selected",b.dataset.v===S.view));
  if(!S.db)return;
  if(S.emps===null){main.innerHTML='<div class="empty"><b>Opening the register</b>Loading employees and this cycle\'s marks.</div>';return}
  if(S.view==="history")S.view="dash";
  if(S.view==="users"&&!isAdmin())S.view="mark";
  if(!activeEmps().length&&S.view!=="people"&&S.view!=="users"&&S.view!=="groups"){main.innerHTML=isAdmin()?'<div class="empty"><b>No employees yet</b>Load the TKR Attendance data file to bring in your employees, shift rotas and marks.<div class="bar" style="justify-content:center;margin-top:14px"><button class="btn gold" id="emptyLoad">Load data file</button></div><p class="note" id="uMsg" role="status"></p><input type="file" id="dataIn" accept=".json,application/json" hidden></div>':'<div class="empty"><b>No employees yet</b>Ask admin to load the register.</div>';
    if(isAdmin()){$("emptyLoad").onclick=()=>$("dataIn").click();$("dataIn").onchange=ev=>{const f=ev.target.files[0];if(!f)return;const rd=new FileReader();rd.onload=()=>{try{loadBackup(JSON.parse(rd.result))}catch(e){toast("That file is not a TKR Attendance data file")}};rd.readAsText(f);ev.target.value=""}}
    return}
  ({dash:viewDash,history:viewHistory,mark:viewMark,daily:viewDaily,cycle:viewCycle,export:viewExport,flags:viewFlags,people:viewPeople,groups:viewGroups,users:viewUsers}[S.view]||viewMark)();
}

/* ---------- history of new, removed, returned and moved employees ---------- */
const where=e=>(e.sheet||"")+(e.section&&e.section!==e.sheet?" · "+e.section:"");
function empEvents(oldL,newL){const O=new Map(oldL.map(e=>[e.id,e])),N=new Set(newL.map(e=>e.id)),at=Date.now(),day=iso(new Date()),ev=[];
  const base=e=>({at,by:email,id:e.id,name:e.name,number:e.number||"",sheet:e.sheet||"",title:e.title||""});
  newL.forEach(e=>{const o=O.get(e.id);
    if(!o){if(e.active!==false)ev.push(Object.assign(base(e),{type:"Added",date:day}));return}
    if(o.active!==false&&e.active===false){ev.push(Object.assign(base(e),{type:"Removed",date:e.leftDate||day,reason:e.leftReason||""}));return}
    if(o.active===false&&e.active!==false){ev.push(Object.assign(base(e),{type:"Back",date:day}));return}
    if(e.active!==false&&(o.sheet!==e.sheet||o.group!==e.group))ev.push(Object.assign(base(e),{type:"Moved",date:day,from:where(o)+(o.shift?" · "+o.shift:""),to:where(e)+(e.shift?" · "+e.shift:"")}))});
  O.forEach((o,id)=>{if(!N.has(id))ev.push(Object.assign(base(o),{type:"Deleted",date:day}))});return ev}
function allHistory(){const H=(S.hist||[]).slice();
  (S.emps||[]).filter(e=>e.active===false&&e.leftDate&&!H.some(h=>h.id===e.id&&h.type==="Removed")).forEach(e=>H.push({type:"Removed",date:e.leftDate,reason:e.leftReason||"",id:e.id,name:e.name,number:e.number||"",sheet:e.sheet||"",title:e.title||"",by:"",at:0}));
  return H.sort((a,b)=>String(b.date).localeCompare(String(a.date))||(b.at||0)-(a.at||0))}
const HTYPE={Added:"New employee",Removed:"Removed",Back:"Back on the register",Moved:"Moved",Deleted:"Deleted from the list"};
function histText(h){return h.type==="Removed"?(h.reason?h.reason.charAt(0)+h.reason.slice(1).toLowerCase():"Removed"):h.type==="Moved"?(h.from+" → "+h.to):HTYPE[h.type]||h.type}
function viewHistory(){const F=S.hf||(S.hf={type:"All",from:"",to:""});const q=S.q.trim().toLowerCase();
  const L=allHistory().filter(h=>(F.type==="All"||h.type===F.type)&&(!F.from||h.date>=F.from)&&(!F.to||h.date<=F.to)&&(S.sheet==="All"||h.sheet===S.sheet)&&(!q||(h.name+" "+h.number+" "+h.title).toLowerCase().includes(q)));
  let h=chipsBar('<input class="search" id="q" type="search" placeholder="Search name, number or job" value="'+esc(S.q)+'">');
  h+='<div class="bar">'+["All","Added","Removed","Back","Moved"].map(t=>'<button class="chip" data-htype="'+t+'" aria-pressed="'+(F.type===t)+'">'+(t==="All"?"Everything":t==="Back"?"Back on the register":t==="Added"?"New employees":t)+'</button>').join("")+'</div>';
  h+='<div class="bar" style="align-items:flex-end"><label class="lbl" style="margin:0">From <input type="date" id="hFrom" value="'+F.from+'"></label><label class="lbl" style="margin:0">To <input type="date" id="hTo" value="'+F.to+'"></label><button class="btn gold" id="xlsHist">Download Excel</button></div>';
  h+='<div class="summary">'+stat("New employees",L.filter(x=>x.type==="Added").length)+stat("Removed",L.filter(x=>x.type==="Removed").length)+stat("Back on the register",L.filter(x=>x.type==="Back").length)+stat("Moved",L.filter(x=>x.type==="Moved").length)+'</div>';
  h+=L.length?'<div class="list"><table class="tbl"><thead><tr><th>Date</th><th>Employee</th><th>Number</th><th>Sheet</th><th>What happened</th><th>Done by</th></tr></thead><tbody>'+L.map(x=>'<tr><td>'+nice(x.date)+' '+yr(x.date)+'</td><td><b>'+esc(x.name)+'</b><br><small>'+esc(x.title||"")+'</small></td><td>'+esc(x.number)+'</td><td>'+esc(x.sheet)+'</td><td><span class="pill s-'+(x.type==="Removed"||x.type==="Deleted"?"out":x.type==="Added"||x.type==="Back"?"present":"leave")+'">'+esc(HTYPE[x.type]||x.type)+'</span> '+(x.type==="Removed"||x.type==="Moved"?esc(histText(x)):"")+'</td><td><small>'+esc(x.by||"")+'</small></td></tr>').join("")+'</tbody></table></div>'
    :'<div class="empty"><b>Nothing here yet</b>From now on every new employee, removal, return and move is recorded here.</div>';
  h+='<p class="note">History is kept from 16 Nov 2026. People removed before that show with the date they were removed.</p>';
  main.innerHTML=h;
  const set=()=>{F.from=$("hFrom").value;F.to=$("hTo").value;render()};$("hFrom").onchange=set;$("hTo").onchange=set;
  $("xlsHist").onclick=()=>{if(!L.length){toast("Nothing to download");return}const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([["TKR Operations – Employee history"+(F.from||F.to?" – "+(F.from?nice(F.from)+" "+yr(F.from):"start")+" to "+(F.to?nice(F.to)+" "+yr(F.to):"today"):"")],[],["Date","Employee","Number","Sheet","Job title","What happened","Details","Done by"]].concat(L.map(x=>[x.date,x.name,x.number,x.sheet,x.title,HTYPE[x.type]||x.type,x.type==="Removed"||x.type==="Moved"?histText(x):"",x.by||""])));
    ws["!cols"]=[{wch:12},{wch:30},{wch:12},{wch:16},{wch:26},{wch:22},{wch:34},{wch:28}];XLSX.utils.book_append_sheet(wb,ws,"History");saveFile("TKR_Employee_History.xlsx",wb)};
}
/* ---------- dashboard: the important things at a glance ---------- */
function viewDash(){const c=cycleOf(S.date),upto=c.dates.filter(d=>d<=S.date),A=activeEmps().filter(e=>S.sheet==="All"||e.sheet===S.sheet);
  const tD=tally(A,[S.date]),tC=tally(A,upto),H=allHistory().filter(h=>h.date>=c.start&&h.date<=c.end&&(S.sheet==="All"||h.sheet===S.sheet));
  const per=A.map(e=>({e,t:tally([e],upto)}));const topExc=per.filter(x=>x.t.exc>0).sort((a,b)=>b.t.exc-a.t.exc).slice(0,8),topOt=per.filter(x=>x.t.ot>0).sort((a,b)=>b.t.ot-a.t.ot).slice(0,6);
  const codes=Object.entries(tC.codes).filter(([k])=>!["CLOCKED IN","HOURS WORKED"].includes(k)&&kindOf(k)!=="present"&&kindOf(k)!=="off").sort((a,b)=>b[1]-a[1]).slice(0,8);
  let h=chipsBar();
  h+='<h2 style="margin:6px 0 4px">'+nice(S.date)+' '+yr(S.date)+'</h2><p class="note" style="margin:0 0 10px">Cycle '+short(c.start)+' – '+short(c.end)+' · '+upto.length+' of '+c.dates.length+' days so far'+(S.sheet!=="All"?' · '+esc(S.sheet):'')+'</p>';
  h+='<div class="summary">'+stat("People on the register",A.length)+stat("Present today",tD.present)+stat("Absent today",tD.absent,tD.absent?"warn":"")+stat("On leave today",tD.leave)+stat("Off today",tD.off)+stat("Not marked yet today",tD.blank,tD.blank?"warn":"")+'</div>';
  h+='<div class="summary">'+stat("Cycle attendance",fmtPct(pct(tC.present,tC.exc)),"pct")+stat("Exception days this cycle",tC.exc)+stat("Overtime this cycle",(Math.round(tC.ot*10)/10)+" h")+stat("New this cycle",H.filter(x=>x.type==="Added").length)+stat("Removed this cycle",H.filter(x=>x.type==="Removed").length)+'</div>';
  if(tD.blank&&S.canWrite)h+='<div class="bar"><button class="btn gold" data-go="mark">Mark today ('+tD.blank+' not marked)</button></div>';
  if(S.sheet==="All")h+='<div class="list"><h3>Per sheet</h3><table class="tbl"><thead><tr><th>Sheet</th><th>People</th><th>Present today</th><th>Absent</th><th>Leave</th><th>Off</th><th>Not marked</th><th>Cycle attendance</th><th>Overtime (cycle)</th></tr></thead><tbody>'+SHEETS.map(sh=>{const es=A.filter(e=>e.sheet===sh);if(!es.length)return"";const t=tally(es,[S.date]),tc=tally(es,upto);
    return '<tr><td><b>'+esc(sh)+'</b></td><td>'+es.length+'</td><td>'+t.present+'</td><td>'+t.absent+'</td><td>'+t.leave+'</td><td>'+t.off+'</td><td>'+(t.blank?'<b style="color:var(--bad,#c0392b)">'+t.blank+'</b>':"0")+'</td><td>'+fmtPct(pct(tc.present,tc.exc))+'</td><td>'+(Math.round(tc.ot*10)/10)+' h</td></tr>'}).join("")+'</tbody></table></div>';
  h+='<div class="list"><h3>Most exception days this cycle</h3>'+(topExc.length?'<table class="tbl"><thead><tr><th>Employee</th><th>Sheet</th><th>Exception days</th><th>Attendance</th></tr></thead><tbody>'+topExc.map(x=>'<tr><td><b>'+esc(x.e.name)+'</b></td><td>'+esc(x.e.sheet)+'</td><td>'+x.t.exc+'</td><td>'+fmtPct(pct(x.t.present,x.t.exc))+'</td></tr>').join("")+'</tbody></table><div class="bar"><button class="btn" data-go="flags">Open Flags</button></div>':'<p class="note" style="margin:0">No AWOP, sick or maternity days this cycle.</p>')+'</div>';
  if(codes.length)h+='<div class="list"><h3>Absence and leave this cycle</h3><div class="bar">'+codes.map(([k,n])=>'<span class="pill s-'+kindOf(k)+'">'+esc(showMark(k))+' · '+n+'</span>').join("")+'</div></div>';
  if(topOt.length)h+='<div class="list"><h3>Most overtime this cycle</h3><table class="tbl"><tbody>'+topOt.map(x=>'<tr><td><b>'+esc(x.e.name)+'</b></td><td>'+esc(x.e.sheet)+'</td><td>'+(Math.round(x.t.ot*10)/10)+' h</td></tr>').join("")+'</tbody></table></div>';
  main.innerHTML=h;
  main.querySelectorAll("[data-go]").forEach(b=>b.onclick=()=>{S.view=b.dataset.go;render();scrollTo(0,0)});
}

function cycleBanner(){const gap=rotaGap();if(!gap)return"";const c=cycleOf(S.date);
  return '<div class="list" style="margin-bottom:14px"><h3>Shift rotas</h3><p class="note" style="margin:0 0 8px">'+gap+' day'+(gap===1?"":"s")+' in '+short(c.start)+' – '+short(c.end)+' have no shift rota yet. Carry the shift pattern forward to fill them.</p><button class="btn gold" id="startCycle">Carry shift pattern forward</button></div>'}

function viewMark(){
  const list=filtered();const groups=bySheetGroup(list);const t=tally(list,[S.date]);const cd=rangeDates(cycleOf(S.date).start,S.date);
  let h=cycleBanner()+chipsBar('<input class="search" id="q" type="search" placeholder="Search name, number or job" value="'+esc(S.q)+'">');
  h+='<div class="summary">'+stat("Marked",(list.length-t.blank)+" / "+list.length)+stat("Present",t.present)+stat("AWOP",t.absent)+stat("Leave",t.leave)+stat("Off",t.off)+stat("Overtime",fmtH(t.ot)+" h")+stat("Attendance",fmtPct(pct(t.present,t.exc)),"pct")+'</div>';
  if(!list.length)h+='<div class="empty"><b>No match</b>Nobody on this sheet matches the search.</div>';
  groups.forEach(s=>{h+='<h2 class="sheet-h">'+esc(s.sheet)+' <small>'+nice(S.date)+'</small></h2>';
    s.groups.forEach(g=>{const r=rotaFor(g.id,S.date);
      h+='<section class="group"><div class="ghead"><span class="gname">'+esc(g.label)+'</span>'+(S.canWrite?'<button type="button" class="rota rbtn '+esc(r)+'" data-rota="'+esc(g.id)+'" aria-label="Change the rota for '+esc(g.label)+' on '+nice(S.date)+'">'+esc(r||"Set rota")+'</button>':(r?'<span class="rota '+esc(r)+'">'+esc(r)+'</span>':""))+
      '<button class="btn sm" data-fill="'+esc(g.id)+'" data-val="P">Unmarked → P</button><button class="btn sm" data-fill="'+esc(g.id)+'" data-val="OFF">Unmarked → OFF</button></div>';
      g.emps.forEach(e=>{h+='<button class="row" data-emp="'+esc(e.id)+'"><span class="who"><span class="n">'+esc(e.name)+'</span><br><span class="t">'+esc([e.number,e.title].filter(Boolean).join(" · "))+'</span></span>'+flagBadges(e,cd)+(otOf(e,S.date)?'<span class="otb">+'+fmtH(otOf(e,S.date))+'h OT</span>':'')+pill(markOf(e,S.date))+'</button>'});
      h+='</section>'})});
  main.innerHTML=h;
}
function stat(k,v,cls){return '<div class="stat '+(cls||"")+'"><div class="k">'+k+'</div><div class="v">'+v+'</div></div>'}

// Day or night: a person works the shift their group's rota shows for that date. Groups with no rota at all (office, Monday to Friday) count as day.
const SHIFT_LBL={all:"All shifts",day:"Day shift",night:"Night shift"};
function shiftOn(e,d){const r=rotaFor(e.group,d);if(r)return r;return Object.keys(S.rota[e.group]||{}).length?"":"DAY"}
const onShift=(es,d,sh)=>sh==="day"?es.filter(e=>shiftOn(e,d)==="DAY"):sh==="night"?es.filter(e=>shiftOn(e,d)==="NIGHT"):es;
function dailyData(){S.dshift=S.dshift||"all";const es=onShift(inPeriod(S.date),S.date,S.dshift);const rows=SHEETS.map(sh=>{const list=es.filter(e=>e.sheet===sh);return {sheet:sh,n:list.length,t:tally(list,[S.date])}}).filter(r=>r.n);
  const all=tally(es,[S.date]);const pick=k=>es.filter(e=>kindOf(markOf(e,S.date))===k);return {rows,all,n:es.length,awop:pick("absent"),leave:pick("leave"),blank:pick("blank"),ot:es.filter(e=>otOf(e,S.date)>0)}}
function viewDaily(){const D=dailyData();
  let h=cycleBanner()+'<div class="bar" role="group" aria-label="Shift">'+["all","day","night"].map(k=>'<button class="chip" data-dshift="'+k+'" aria-pressed="'+(S.dshift===k)+'">'+SHIFT_LBL[k]+'</button>').join("")+'</div>'+
   (S.dshift!=="all"?'<p class="note" style="margin:0 0 10px">'+SHIFT_LBL[S.dshift]+' on '+nice(S.date)+': people whose group rota shows '+(S.dshift==="day"?"DAY":"NIGHT")+' that day'+(S.dshift==="day"?", plus groups with no rota (office and Monday to Friday)":"")+'. Groups on OFF are left out.</p>':"")+
   '<div class="bar"><button class="btn gold" id="copyDaily">Copy report for WhatsApp</button><button class="btn" id="xlsDaily">Download Excel</button><button class="btn gold" id="pdfDaily">Download PDF report</button><button class="btn gold" id="pptDaily">Download PowerPoint report</button></div>';
  h+='<div class="summary">'+stat("Employees",D.n)+stat("Present",D.all.present)+stat("AWOP",D.all.absent)+stat("Leave",D.all.leave)+stat("Off",D.all.off)+stat("Not marked",D.all.blank)+stat("Overtime",fmtH(D.all.ot)+" h")+stat("Attendance",fmtPct(pct(D.all.present,D.all.exc)),"pct")+'</div>';
  h+='<h2 class="sheet-h">Daily report'+(S.dshift!=="all"?" – "+SHIFT_LBL[S.dshift]:"")+' <small>'+nice(S.date)+'</small></h2><div class="tbl-wrap"><table><thead><tr><th>Sheet</th><th>Employees</th><th>Present</th><th>AWOP</th><th>Leave</th><th>Off</th><th>Resigned, absconded, dismissed, deceased, suspended</th><th>Not marked</th><th>Overtime</th><th>Attendance</th></tr></thead><tbody>';
  D.rows.forEach(r=>{h+=trow(r.sheet,r.n,r.t)});h+=trow("Total",D.n,D.all,"tot")+'</tbody></table></div>';
  const li=(arr,f)=>arr.length?'<ul>'+arr.map(e=>'<li>'+esc(e.name)+' <span class="note">'+esc(e.sheet)+(f?" · "+esc(f(e)):"")+'</span></li>').join("")+'</ul>':'<p class="none">None</p>';
  h+='<div class="lists"><div class="list"><h3>AWOP ('+D.awop.length+')</h3>'+li(D.awop)+'</div><div class="list"><h3>On leave ('+D.leave.length+')</h3>'+li(D.leave,e=>markOf(e,S.date))+'</div><div class="list"><h3>Not marked ('+D.blank.length+')</h3>'+li(D.blank)+'</div><div class="list"><h3>Overtime ('+fmtH(D.all.ot)+' h)</h3>'+li(D.ot,e=>fmtH(otOf(e,S.date))+' h')+'</div></div>';
  main.innerHTML=h;
}
function trow(name,n,t,cls){const p=pct(t.present,t.exc);return '<tr class="'+(cls||"")+'"><td>'+esc(name)+'</td><td>'+n+'</td><td>'+t.present+'</td><td>'+t.absent+'</td><td>'+t.leave+'</td><td>'+t.off+'</td><td>'+t.out+'</td><td>'+t.blank+'</td><td>'+hm(t.ot)+'</td><td class="pc">'+fmtPct(p)+'</td></tr>'}
function dailyText(){const D=dailyData();let s="*TKR Operations – Daily Attendance"+(S.dshift!=="all"?" – "+SHIFT_LBL[S.dshift]:"")+"*\n"+nice(S.date)+" "+parse(S.date).getFullYear()+"\n\n";
  D.rows.forEach(r=>{s+="*"+r.sheet+"*: Present "+r.t.present+", AWOP "+r.t.absent+", Leave "+r.t.leave+", Off "+r.t.off+(r.t.blank?", Not marked "+r.t.blank:"")+" – "+fmtPct(pct(r.t.present,r.t.exc))+"\n"});
  s+="\n*Total*: "+D.all.present+" present, "+D.all.absent+" AWOP, "+D.all.leave+" on leave, "+D.all.off+" off – attendance "+fmtPct(pct(D.all.present,D.all.exc))+"\n";
  if(D.awop.length)s+="\n*AWOP*\n"+D.awop.map(e=>"- "+e.name+" ("+e.sheet+")").join("\n")+"\n";
  if(D.ot.length)s+="\n*Overtime* ("+fmtH(D.all.ot)+" h)\n"+D.ot.map(e=>"- "+e.name+" ("+e.sheet+"): "+fmtH(otOf(e,S.date))+" h").join("\n")+"\n";
  if(D.leave.length)s+="\n*On leave*\n"+D.leave.map(e=>"- "+e.name+" ("+e.sheet+"): "+markOf(e,S.date)).join("\n")+"\n";
  return s}

const LEAVE_COLS=["ANNUAL LEAVE","SICK","FAMILY RESPONSIBILITY","FAMILY EMERGENCY","STUDY LEAVE","MATERNITY LEAVE"];
function cycleRows(){const c=cycleOf(S.date);const upto=c.dates.filter(d=>d<=S.date);
  return filtered(c.start).sort((a,b)=>SHEETS.indexOf(a.sheet)-SHEETS.indexOf(b.sheet)||a.order-b.order).map(e=>{const t=tally([e],upto);return {e,t,p:pct(t.present,t.exc)}})}
/* ---------- cycle report as PDF or PowerPoint ---------- */
const loadJs=src=>new Promise((res,rej)=>{if(document.querySelector('script[src="'+src+'"]'))return res();const t=document.createElement("script");t.src=src;t.onload=res;t.onerror=()=>rej(new Error("Open the app once with signal, then try again."));document.head.appendChild(t)});
async function logoData(){try{const b=await (await fetch("../icons/logo-96.png")).blob();return await new Promise(r=>{const f=new FileReader();f.onload=()=>r(f.result);f.readAsDataURL(b)})}catch(e){return null}}
function reportData(){const c=cycleOf(S.date),upto=c.dates.filter(d=>d<=S.date),rows=cycleRows(),A=rows.map(r=>r.e),all=tally(A,upto);
  const sheets=SHEETS.map(sh=>{const es=A.filter(e=>e.sheet===sh);if(!es.length)return null;const t=tally(es,upto);return {sh,n:es.length,t,p:pct(t.present,t.exc)}}).filter(Boolean);
  const codes=Object.entries(all.codes).filter(([k])=>!["CLOCKED IN","HOURS WORKED"].includes(k)&&kindOf(k)!=="present"&&kindOf(k)!=="off").sort((a,b)=>b[1]-a[1]);
  const low=rows.filter(r=>r.p!=null&&r.t.exc>0).sort((a,b)=>a.p-b.p).slice(0,10);
  const H=allHistory().filter(h=>h.date>=c.start&&h.date<=c.end&&(S.sheet==="All"||h.sheet===S.sheet)&&h.type!=="Deleted");
  const last=upto[upto.length-1]||c.start,period=nice(c.start)+" "+yr(c.start)+" – "+nice(c.end)+" "+yr(c.end);
  return {c,upto,rows,all,sheets,codes,low,H,last,period,sheetLbl:S.sheet==="All"?"All sheets":S.sheet,
    file:"TKR_Attendance_Cycle_Report_"+short(c.start).replace(" ","-")+"_to_"+short(c.end).replace(" ","-")+"_"+parse(c.end).getFullYear()+(S.sheet==="All"?"":"_"+S.sheet.replace(/\s+/g,"_"))}}
const ot1=v=>v?hm(v):"0:00",pc1=v=>v==null?"–":v.toFixed(1)+"%",cap=k=>k.charAt(0)+k.slice(1).toLowerCase();
async function pdfCycle(){const btn=$("pdfCycle");btn.disabled=true;btn.textContent="Making the PDF…";
  try{await loadJs("../lib/jspdf.umd.min.js");await loadJs("../lib/jspdf.plugin.autotable.min.js");
    const D=reportData(),logo=await logoData(),{jsPDF}=window.jspdf,P=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"}),W=297,GOLD=[212,161,42],BLACK=[17,17,17];
    const head=()=>{P.setFillColor(...BLACK);P.rect(0,0,W,24,"F");P.setFillColor(...GOLD);P.rect(0,24,W,1.4,"F");if(logo)P.addImage(logo,"PNG",10,3.5,17,17);
      P.setTextColor(255,255,255);P.setFont("helvetica","bold");P.setFontSize(14);P.text("TKR OPERATIONS – ATTENDANCE CYCLE REPORT",32,11);
      P.setTextColor(...GOLD);P.setFontSize(11);P.text("Cycle "+D.period+"   ·   "+D.sheetLbl+"   ·   days counted "+D.upto.length+" of "+D.c.dates.length+" (to "+nice(D.last)+")",32,18.5)};
    const tbl=(y,title,headRow,body,opt={})=>{P.setTextColor(...BLACK);P.setFont("helvetica","bold");P.setFontSize(11.5);P.text(title,12,y);
      P.autoTable(Object.assign({startY:y+2,head:[headRow],body,margin:{left:12,right:12,top:30},theme:"grid",styles:{fontSize:8.6,cellPadding:1.6,textColor:BLACK,lineColor:[217,211,197]},
        headStyles:{fillColor:[38,38,38],textColor:255,fontStyle:"bold"},alternateRowStyles:{fillColor:[255,248,230]},didDrawPage:()=>head()},opt));return P.lastAutoTable.finalY+8};
    head();
    // summary tiles
    const tiles=[["People",String(D.rows.length)],["Present days",String(D.all.present)],["Exception days",String(D.all.exc)],["Cycle attendance",pc1(pct(D.all.present,D.all.exc))],["Days counted",D.upto.length+" of "+D.c.dates.length],["New / removed",D.H.filter(h=>h.type==="Added").length+" / "+D.H.filter(h=>h.type==="Removed").length]];
    tiles.forEach(([k,v],i)=>{const x=12+i*45.5;P.setDrawColor(...GOLD);P.setFillColor(255,248,230);P.roundedRect(x,31,43,20,2,2,"FD");P.setTextColor(110,100,80);P.setFont("helvetica","bold");P.setFontSize(7.5);P.text(k.toUpperCase(),x+3,37);
      P.setTextColor(...BLACK);P.setFontSize(16);P.text(v,x+3,47)});
    let y=60;
    y=tbl(y,"Per sheet",["Sheet","People","Present","AWOP","Leave","Off","Not marked","Attendance"],D.sheets.map(s=>[s.sh,s.n,s.t.present,s.t.absent,s.t.leave,s.t.off,s.t.blank,pc1(s.p)]).concat([[{content:"All",styles:{fontStyle:"bold"}},D.rows.length,D.all.present,D.all.absent,D.all.leave,D.all.off,D.all.blank,pc1(pct(D.all.present,D.all.exc))]]));
    if(D.codes.length)y=tbl(y,"Absence and leave this cycle (days)",["Code","Days"],D.codes.map(([k,n])=>[cap(showMark(k)),n]),{tableWidth:110});
    y=tbl(y,"Lowest attendance this cycle",["Employee","Sheet","Present","Exception days","Attendance"],D.low.length?D.low.map(r=>[r.e.name,r.e.sheet,r.t.present,r.t.exc,pc1(r.p)]):[["No exception days this cycle","","","",""]]);
    y=tbl(y,"New and removed employees this cycle",["Date","Employee","Sheet","What happened"],D.H.length?D.H.map(h=>[nice(h.date)+" "+yr(h.date),h.name,h.sheet,(HTYPE[h.type]||h.type)+(h.type==="Removed"||h.type==="Moved"?" – "+histText(h):"")]):[["","No one added or removed this cycle","",""]]);
    P.addPage();head();
    tbl(32,"Every employee",["Employee","Sheet / group","Present","AWOP",...LEAVE_COLS.map(k=>cap(k.replace(" LEAVE","").replace("FAMILY ","Fam. "))),"Off","Not marked","Attendance"],
      D.rows.map(r=>[r.e.name,r.e.sheet+" · "+groupLabel(r.e),r.t.present,r.t.absent,...LEAVE_COLS.map(k=>r.t.codes[k]||""),r.t.off,r.t.blank,pc1(r.p)]),{styles:{fontSize:7.6,cellPadding:1.2,textColor:BLACK,lineColor:[217,211,197]}});
    const n=P.getNumberOfPages();for(let i=1;i<=n;i++){P.setPage(i);P.setFontSize(8);P.setTextColor(120);P.setFont("helvetica","normal");P.text("Attendance = present days ÷ (present + exception days). Exceptions: AWOP, sick and maternity leave.   Printed "+nice(iso(new Date()))+" "+yr(iso(new Date()))+"   ·   page "+i+" of "+n,12,204)}
    P.save(D.file+".pdf");toast("PDF report downloaded")}
  catch(e){console.warn(e);toast("The PDF couldn't be made: "+(e.message||"try again"))}
  btn.disabled=false;btn.textContent="Download PDF report"}
async function pptCycle(){const btn=$("pptCycle");btn.disabled=true;btn.textContent="Making the PowerPoint…";
  try{await loadJs("../lib/pptxgen.bundle.js");const D=reportData(),logo=await logoData(),pp=new window.PptxGenJS();pp.layout="LAYOUT_WIDE";pp.author="TKR Operations";pp.title="Attendance cycle report "+D.period;
    const G="D4A12A",K="111111",CR="FFF8E6",TXT={fontFace:"Arial",color:K};
    const base=(title)=>{const s=pp.addSlide();s.background={color:"FFFFFF"};s.addShape(pp.ShapeType.rect,{x:0,y:0,w:13.33,h:0.9,fill:{color:K}});s.addShape(pp.ShapeType.rect,{x:0,y:0.9,w:13.33,h:0.06,fill:{color:G}});
      if(logo)s.addImage({data:logo,x:0.25,y:0.1,w:0.7,h:0.7});s.addText(title,{x:1.1,y:0.08,w:9,h:0.5,fontFace:"Arial",fontSize:22,bold:true,color:"FFFFFF"});
      s.addText("Cycle "+D.period+" · "+D.sheetLbl,{x:1.1,y:0.52,w:11,h:0.32,fontFace:"Arial",fontSize:12,color:G,bold:true});
      s.addText("TKR Operations · Attendance",{x:0.3,y:7.1,w:6,h:0.3,fontFace:"Arial",fontSize:9,color:"888888"});return s};
    const T=(s,rows,o)=>s.addTable(rows,Object.assign({x:0.4,y:1.2,w:12.5,fontFace:"Arial",fontSize:12,color:K,border:{type:"solid",pt:0.5,color:"D9D3C5"},autoPage:false},o));
    const hdr=a=>a.map(t=>({text:String(t),options:{bold:true,color:"FFFFFF",fill:{color:"262626"}}}));
    // 1 title
    let s=pp.addSlide();s.background={color:"FFFFFF"};s.addShape(pp.ShapeType.rect,{x:0,y:0,w:13.33,h:0.25,fill:{color:G}});if(logo)s.addImage({data:logo,x:0.6,y:0.6,w:1.4,h:1.4});
    s.addText("ATTENDANCE CYCLE REPORT",{x:0.6,y:2.4,w:12,h:0.9,fontFace:"Arial",fontSize:40,bold:true,color:K});
    s.addText(D.period,{x:0.6,y:3.3,w:12,h:0.7,fontFace:"Arial",fontSize:26,bold:true,color:G});
    s.addText(D.sheetLbl+" · days counted "+D.upto.length+" of "+D.c.dates.length+" (to "+nice(D.last)+" "+yr(D.last)+")",{x:0.6,y:4.1,w:12,h:0.5,fontFace:"Arial",fontSize:16,color:"555555"});
    s.addText("TKR Operations · Tharisa Mine",{x:0.6,y:6.6,w:8,h:0.4,fontFace:"Arial",fontSize:12,color:"AAAAAA"});
    // 2 summary
    s=base("Summary");const tiles=[["People on the register",String(D.rows.length)],["Present days",String(D.all.present)],["Exception days",String(D.all.exc)],["Cycle attendance",pc1(pct(D.all.present,D.all.exc))],["Days counted",D.upto.length+" of "+D.c.dates.length],["New / removed",D.H.filter(h=>h.type==="Added").length+" / "+D.H.filter(h=>h.type==="Removed").length]];
    tiles.forEach(([k,v],i)=>{const x=0.5+(i%3)*4.15,y=1.4+Math.floor(i/3)*2.6;s.addShape(pp.ShapeType.roundRect,{x,y,w:3.9,h:2.2,fill:{color:CR},line:{color:G,width:1.5},rectRadius:0.1});
      s.addText(k.toUpperCase(),{x:x+0.25,y:y+0.2,w:3.5,h:0.4,fontFace:"Arial",fontSize:12,bold:true,color:"6E6450"});s.addText(v,{x:x+0.25,y:y+0.75,w:3.5,h:1,fontFace:"Arial",fontSize:40,bold:true,color:K})});
    s.addText("Attendance = present days ÷ (present + exception days). Exceptions are AWOP, sick and maternity leave. Annual, family and study leave are not counted as absent.",{x:0.5,y:6.6,w:12.3,h:0.4,fontFace:"Arial",fontSize:10,color:"666666"});
    // 3 per sheet
    s=base("Per sheet");T(s,[hdr(["Sheet","People","Present","AWOP","Leave","Off","Not marked","Attendance"])].concat(D.sheets.map(x=>[x.sh,x.n,x.t.present,x.t.absent,x.t.leave,x.t.off,x.t.blank,pc1(x.p)].map(String)),
      [["All",D.rows.length,D.all.present,D.all.absent,D.all.leave,D.all.off,D.all.blank,pc1(pct(D.all.present,D.all.exc))].map(v=>({text:String(v),options:{bold:true,fill:{color:"FFF2CC"}}}))]),{fontSize:14});
    // 4 absence and leave
    s=base("Absence and leave this cycle");
    if(D.codes.length){s.addChart(pp.ChartType.bar,[{name:"Days",labels:D.codes.slice(0,8).map(([k])=>cap(showMark(k))),values:D.codes.slice(0,8).map(([,n])=>n)}],{x:0.4,y:1.2,w:7.6,h:5.6,barDir:"bar",chartColors:[G],catAxisLabelFontSize:12,valAxisLabelFontSize:10,showValue:true,dataLabelFontSize:11,catAxisOrientation:"maxMin"});
      T(s,[hdr(["Code","Days"])].concat(D.codes.map(([k,n])=>[cap(showMark(k)),String(n)])),{x:8.4,w:4.5,fontSize:12})}
    else s.addText("No absence or leave marked this cycle.",{x:0.5,y:1.5,w:12,h:0.6,...TXT,fontSize:18});
    // 5 lowest attendance
    s=base("Lowest attendance this cycle");
    if(D.low.length)T(s,[hdr(["Employee","Sheet","Present","Exception days","Attendance"])].concat(D.low.map(r=>[r.e.name,r.e.sheet,String(r.t.present),String(r.t.exc),pc1(r.p)])),{fontSize:13});
    else s.addText("No exception days this cycle.",{x:0.5,y:1.5,w:12,h:0.6,...TXT,fontSize:18});
    // 6 joiners and leavers
    s=base("New and removed employees");
    if(D.H.length)T(s,[hdr(["Date","Employee","Sheet","What happened"])].concat(D.H.slice(0,14).map(h=>[nice(h.date)+" "+yr(h.date),h.name,h.sheet,(HTYPE[h.type]||h.type)+(h.type==="Removed"||h.type==="Moved"?" – "+histText(h):"")])),{fontSize:12});
    else s.addText("No one added or removed this cycle.",{x:0.5,y:1.5,w:12,h:0.6,...TXT,fontSize:18});
    await pp.writeFile({fileName:D.file+".pptx"});toast("PowerPoint report downloaded")}
  catch(e){console.warn(e);toast("The PowerPoint couldn't be made: "+(e.message||"try again"))}
  btn.disabled=false;btn.textContent="Download PowerPoint report"}
/* ---------- daily report as PDF or PowerPoint (no overtime) ---------- */
function dailySpec(){const D=dailyData(),day=nice(S.date)+" "+yr(S.date),shift=S.dshift!=="all"?SHIFT_LBL[S.dshift]:"All shifts";
  const row=(name,n,t,b)=>[name,n,t.present,t.absent,t.leave,t.off,t.out,t.blank,pc1(pct(t.present,t.exc))].map(v=>b?{content:String(v),bold:1}:String(v));
  const names=(arr,f)=>arr.length?arr.map(e=>[e.name,e.sheet,f?f(e):""]):[["None","",""]];
  return {title:"DAILY ATTENDANCE REPORT",sub:day+"   ·   "+shift,file:"TKR_Attendance_Daily_Report_"+S.date+(S.dshift!=="all"?"_"+S.dshift:""),
    tiles:[["Employees",D.n],["Present",D.all.present],["AWOP",D.all.absent],["Leave",D.all.leave],["Off",D.all.off],["Not marked",D.all.blank],["Attendance",pc1(pct(D.all.present,D.all.exc))]],
    sections:[{title:"Per sheet",head:["Sheet","Employees","Present","AWOP","Leave","Off","Resigned, absconded, dismissed, deceased, suspended","Not marked","Attendance"],rows:D.rows.map(r=>row(r.sheet,r.n,r.t)).concat([row("Total",D.n,D.all,1)])},
      {title:"AWOP ("+D.awop.length+")",head:["Employee","Sheet",""],rows:names(D.awop)},
      {title:"On leave ("+D.leave.length+")",head:["Employee","Sheet","Leave"],rows:names(D.leave,e=>cap(showMark(markOf(e,S.date))))},
      {title:"Not marked ("+D.blank.length+")",head:["Employee","Sheet",""],rows:names(D.blank)}]}}
async function pdfSpec(sp,btn,label){btn.disabled=true;btn.textContent="Making the PDF…";
  try{await loadJs("../lib/jspdf.umd.min.js");await loadJs("../lib/jspdf.plugin.autotable.min.js");
    const logo=await logoData(),{jsPDF}=window.jspdf,P=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"}),W=297,GOLD=[212,161,42],BLACK=[17,17,17];
    const head=()=>{P.setFillColor(...BLACK);P.rect(0,0,W,24,"F");P.setFillColor(...GOLD);P.rect(0,24,W,1.4,"F");if(logo)P.addImage(logo,"PNG",10,3.5,17,17);
      P.setTextColor(255,255,255);P.setFont("helvetica","bold");P.setFontSize(14);P.text("TKR OPERATIONS – "+sp.title,32,11);P.setTextColor(...GOLD);P.setFontSize(11);P.text(sp.sub,32,18.5)};
    head();const tw=(W-24-(sp.tiles.length-1)*2.5)/sp.tiles.length;
    sp.tiles.forEach(([k,v],i)=>{const x=12+i*(tw+2.5);P.setDrawColor(...GOLD);P.setFillColor(255,248,230);P.roundedRect(x,31,tw,20,2,2,"FD");P.setTextColor(110,100,80);P.setFont("helvetica","bold");P.setFontSize(7.5);P.text(String(k).toUpperCase(),x+3,37);P.setTextColor(...BLACK);P.setFontSize(16);P.text(String(v),x+3,47)});
    let y=60;for(const sec of sp.sections){P.setTextColor(...BLACK);P.setFont("helvetica","bold");P.setFontSize(11.5);if(y>185){P.addPage();head();y=34}P.text(sec.title,12,y);
      P.autoTable({startY:y+2,head:[sec.head],body:sec.rows.map(r=>r.map(c=>typeof c==="object"?{content:c.content,styles:{fontStyle:"bold",fillColor:[255,242,204]}}:c)),margin:{left:12,right:12,top:30},theme:"grid",
        styles:{fontSize:8.8,cellPadding:1.6,textColor:BLACK,lineColor:[217,211,197]},headStyles:{fillColor:[38,38,38],textColor:255,fontStyle:"bold"},alternateRowStyles:{fillColor:[255,248,230]},didDrawPage:()=>head()});y=P.lastAutoTable.finalY+9}
    const n=P.getNumberOfPages();for(let i=1;i<=n;i++){P.setPage(i);P.setFontSize(8);P.setTextColor(120);P.setFont("helvetica","normal");P.text("Attendance = present ÷ (present + exception days). Exceptions: AWOP, sick and maternity leave.   Printed "+nice(iso(new Date()))+" "+yr(iso(new Date()))+"   ·   page "+i+" of "+n,12,204)}
    P.save(sp.file+".pdf");toast("PDF report downloaded")}
  catch(e){console.warn(e);toast("The PDF couldn't be made: "+(e.message||"try again"))}
  btn.disabled=false;btn.textContent=label}
async function pptSpec(sp,btn,label){btn.disabled=true;btn.textContent="Making the PowerPoint…";
  try{await loadJs("../lib/pptxgen.bundle.js");const logo=await logoData(),pp=new window.PptxGenJS();pp.layout="LAYOUT_WIDE";pp.author="TKR Operations";const G="D4A12A",K="111111";
    const base=t=>{const s=pp.addSlide();s.addShape(pp.ShapeType.rect,{x:0,y:0,w:13.33,h:0.9,fill:{color:K}});s.addShape(pp.ShapeType.rect,{x:0,y:0.9,w:13.33,h:0.06,fill:{color:G}});if(logo)s.addImage({data:logo,x:0.25,y:0.1,w:0.7,h:0.7});
      s.addText(t,{x:1.1,y:0.08,w:11,h:0.5,fontFace:"Arial",fontSize:22,bold:true,color:"FFFFFF"});s.addText(sp.sub,{x:1.1,y:0.52,w:11,h:0.32,fontFace:"Arial",fontSize:12,color:G,bold:true});return s};
    let s=pp.addSlide();s.background={color:"FFFFFF"};s.addShape(pp.ShapeType.rect,{x:0,y:0,w:13.33,h:0.25,fill:{color:G}});if(logo)s.addImage({data:logo,x:0.6,y:0.6,w:1.4,h:1.4});
    s.addText(sp.title,{x:0.6,y:2.4,w:12,h:0.9,fontFace:"Arial",fontSize:40,bold:true,color:K});s.addText(sp.sub,{x:0.6,y:3.3,w:12,h:0.7,fontFace:"Arial",fontSize:24,bold:true,color:G});
    s.addText("TKR Operations · Tharisa Mine",{x:0.6,y:6.6,w:8,h:0.4,fontFace:"Arial",fontSize:12,color:"AAAAAA"});
    s=base("Summary");const per=Math.ceil(sp.tiles.length/2),tw=12.3/per;
    sp.tiles.forEach(([k,v],i)=>{const x=0.5+(i%per)*tw,y=1.4+Math.floor(i/per)*2.6;s.addShape(pp.ShapeType.roundRect,{x,y,w:tw-0.25,h:2.2,fill:{color:"FFF8E6"},line:{color:G,width:1.5},rectRadius:0.1});
      s.addText(String(k).toUpperCase(),{x:x+0.2,y:y+0.2,w:tw-0.6,h:0.4,fontFace:"Arial",fontSize:12,bold:true,color:"6E6450"});s.addText(String(v),{x:x+0.2,y:y+0.75,w:tw-0.6,h:1,fontFace:"Arial",fontSize:36,bold:true,color:K})});
    const hdr=a=>a.map(t=>({text:String(t),options:{bold:true,color:"FFFFFF",fill:{color:"262626"}}}));
    for(const sec of sp.sections){const per=sec.per||14;for(let i=0;i<Math.max(1,sec.rows.length);i+=per){s=base(sec.title+(sec.rows.length>per?" ("+(i/per+1)+")":""));
      s.addTable([hdr(sec.head)].concat(sec.rows.slice(i,i+per).map(r=>r.map(c=>typeof c==="object"?{text:c.content,options:{bold:true,fill:{color:"FFF2CC"}}}:String(c)))),{x:0.4,y:1.2,w:12.5,fontFace:"Arial",fontSize:sec.fs||13,color:K,border:{type:"solid",pt:0.5,color:"D9D3C5"},...(sec.colW?{colW:sec.colW}:{})})}}
    await pp.writeFile({fileName:sp.file+".pptx"});toast("PowerPoint report downloaded")}
  catch(e){console.warn(e);toast("The PowerPoint couldn't be made: "+(e.message||"try again"))}
  btn.disabled=false;btn.textContent=label}
function viewCycle(){const c=cycleOf(S.date);const upto=c.dates.filter(d=>d<=S.date);let rows=cycleRows();
  if(S.sortLow)rows=rows.slice().sort((a,b)=>(a.p??101)-(b.p??101));
  const all=tally(rows.map(r=>r.e),upto);
  let h=cycleBanner()+chipsBar('<input class="search" id="q" type="search" placeholder="Search name, number or job" value="'+esc(S.q)+'">');
  h+='<div class="bar"><button class="btn" id="sortLow" aria-pressed="'+S.sortLow+'">'+(S.sortLow?"Show in timesheet order":"Lowest attendance first")+'</button><button class="btn gold" id="xlsCycle">Download Excel timesheet</button><button class="btn gold" id="pdfCycle">Download PDF report</button><button class="btn gold" id="pptCycle">Download PowerPoint report</button></div>';
  h+='<div class="summary">'+stat("Days counted",upto.length+" of "+c.dates.length)+stat("Present days",all.present)+stat("Exception days",all.exc)+stat("Cycle attendance",fmtPct(pct(all.present,all.exc)),"pct")+'</div>';
  h+='<p class="note">'+short(c.start)+' to '+short(upto[upto.length-1]||c.start)+'. Attendance = present days ÷ (present days + exception days). Exceptions are AWOP, sick and maternity leave. Annual leave, family responsibility, family emergency and study leave are not counted as absent. Medical counts as present. Off days are not counted.</p>';
  h+='<div class="tbl-wrap"><table><thead><tr><th>Employee</th><th>Present</th><th>AWOP</th>'+LEAVE_COLS.map(k=>'<th>'+esc(k.replace(" LEAVE","").replace("FAMILY ","FAM. "))+'</th>').join("")+'<th>Off</th><th>Not marked</th><th>Overtime</th><th>Attendance</th></tr></thead><tbody>';
  let last="";rows.forEach(r=>{if(!S.sortLow&&r.e.sheet!==last){last=r.e.sheet;h+='<tr class="tot"><td colspan="'+(LEAVE_COLS.length+7)+'">'+esc(last)+'</td></tr>'}
    const t=r.t;h+='<tr><td>'+esc(r.e.name)+'<br><span class="note">'+esc(r.e.sheet+" · "+groupLabel(r.e))+'</span></td><td>'+t.present+'</td><td>'+t.absent+'</td>'+LEAVE_COLS.map(k=>'<td>'+(t.codes[k]||"")+'</td>').join("")+'<td>'+t.off+'</td><td>'+t.blank+'</td><td>'+(t.ot?hm(t.ot):'')+'</td><td class="'+(r.p!=null&&r.p<90?"lowpc":"pc")+'">'+fmtPct(r.p)+'</td></tr>'});
  h+='</tbody></table></div>';main.innerHTML=h;
}

let editing=null;
// Shift groups: the groups the employees are in, plus empty groups made in the app (saved with the employee list)
function allGroups(){const m=new Map();(S.emps||[]).filter(x=>x.active!==false).forEach(x=>{if(!m.has(x.group))m.set(x.group,{id:x.group,sheet:x.sheet,section:x.section,shift:x.shift})});
  (S.groupDefs||[]).forEach(g=>{if(g.removed){if(!(S.emps||[]).some(x=>x.active!==false&&x.group===g.id))m.delete(g.id)}else m.set(g.id,Object.assign(m.get(g.id)||{},g))});
  return [...m.values()].sort((a,b)=>SHEETS.indexOf(a.sheet)-SHEETS.indexOf(b.sheet)||groupFirst(a.id)-groupFirst(b.id))}
function groupFirst(id){const o=(S.emps||[]).filter(x=>x.group===id).map(x=>x.order);return o.length?Math.min(...o):1e9}
function orderFor(gid,sheet,list,skip){const same=list.filter(x=>x.group===gid&&x.id!==skip);if(same.length)return Math.max(...same.map(x=>x.order))+0.001;
  const sh=list.filter(x=>x.sheet===sheet&&x.id!==skip);return (sh.length?Math.max(...sh.map(x=>x.order)):list.length)+1}
const gName=g=>groupLabel({sheet:g.sheet,section:g.section,shift:g.shift});
let editingGroup=null;
// ---------- editing rotas: one day from the Mark tab, or the whole cycle from Shift groups ----------
function groupMeta(gid){const g=allGroups().find(x=>x.id===gid)||S.gmeta[gid]||{};return {sheet:g.sheet||"",section:g.section||"",shift:g.shift||""}}
async function saveRota(gid,days){if(!S.canWrite){toast("You can view this register but not change it.");return false}
  const by={};for(const d in days){const c=cycleOf(d);(by[c.start]=by[c.start]||{c,rota:{}}).rota[d]=days[d]}
  try{for(const k in by){const {c,rota}=by[k];await S.db.doc("cycles/r-"+c.start).update({start:c.start,end:c.end,groups:{[gid]:Object.assign(groupMeta(gid),{rota})}})}
    S.rota[gid]=Object.assign(S.rota[gid]||{},days);return true}
  catch(e){toast(denied(e)?"You can view this register but not change it.":"Not saved: "+(e.message||"try again"));return false}}
function openRotaDay(gid){const g=allGroups().find(x=>x.id===gid);if(!g)return;const cur=rotaFor(gid,S.date);
  $("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true" aria-labelledby="rdt"><h2 id="rdt">'+esc(g.sheet+" · "+gName(g))+'</h2><div class="sub">Rota for '+nice(S.date)+' '+yr(S.date)+'</div>'+
   '<div class="codes" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:12px 0">'+["DAY","NIGHT","OFF"].map(v=>'<button class="btn'+(cur===v?" gold":"")+'" data-rv="'+v+'">'+v+'</button>').join("")+'</div>'+
   '<div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn" data-rv="">Clear</button><button class="btn" id="rotaCycle">Edit the whole cycle</button><button class="btn" id="closeP">Close</button></div></div></div>';
  const close=()=>{$("sheetHost").innerHTML=""};$("closeP").onclick=close;$("scrim").onclick=ev=>{if(ev.target.id==="scrim")close()};
  $("rotaCycle").onclick=()=>openRotaCycle(gid);
  $("sheetHost").querySelectorAll("[data-rv]").forEach(b=>b.onclick=async()=>{close();if(await saveRota(gid,{[S.date]:b.dataset.rv})){toast(gName(g)+" "+(b.dataset.rv||"rota cleared")+" on "+short(S.date));render()}})}
function openRotaCycle(gid){const g=allGroups().find(x=>x.id===gid);if(!g)return;const c=cycleOf(S.date);const sel={};c.dates.forEach(d=>sel[d]=rotaFor(gid,d));
  const opt=(d)=>'<select data-rd="'+d+'">'+["","DAY","NIGHT","OFF"].map(v=>'<option value="'+v+'"'+(sel[d]===v?" selected":"")+'>'+(v||"–")+'</option>').join("")+'</select>';
  $("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true" aria-labelledby="rct" style="max-height:90vh;overflow:auto"><h2 id="rct">'+esc(g.sheet+" · "+gName(g))+'</h2><div class="sub">Rota for '+short(c.start)+' – '+short(c.end)+' '+yr(c.end)+'</div>'+
   '<p class="note" style="margin:8px 0">Fill a pattern from a start date, or change single days below.</p>'+
   '<div class="clock" style="flex-wrap:wrap"><label class="lbl" for="rp-from" style="margin:0">From</label><input type="date" id="rp-from" value="'+c.start+'" min="'+c.start+'" max="'+c.end+'">'+
   '<label class="lbl" for="rp-pat" style="margin:0">Pattern</label><input id="rp-pat" placeholder="e.g. 2 DAY, 3 NIGHT, 3 OFF" style="min-width:200px"><button class="btn" id="rpFill">Fill</button></div>'+
   '<div class="tbl-wrap" style="margin-top:10px"><table><tbody>'+c.dates.map(d=>'<tr><td>'+nice(d)+'</td><td>'+opt(d)+'</td></tr>').join("")+'</tbody></table></div>'+
   '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px"><button class="btn gold" id="rcSave">Save rota</button><button class="btn" id="closeP">Cancel</button></div></div></div>';
  const host=$("sheetHost"),close=()=>{host.innerHTML=""};$("closeP").onclick=close;$("scrim").onclick=ev=>{if(ev.target.id==="scrim")close()};
  $("rpFill").onclick=()=>{const seq=[];const parts=$("rp-pat").value.toUpperCase().split(/[,;]+/).map(x=>x.trim()).filter(Boolean);
    for(const p of parts){const m=p.match(/^(\d+)\s*(?:X\s*)?(DAY|NIGHT|OFF)S?$/)||p.match(/^(DAY|NIGHT|OFF)S?\s*(?:X\s*)?(\d+)$/);
      if(!m){toast("Write the pattern like 2 DAY, 3 NIGHT, 3 OFF");return}const n=+(m[1].match(/\d/)?m[1]:m[2]),v=m[1].match(/\d/)?m[2]:m[1];if(n<1||n>31){toast("Use 1 to 31 days in each part");return}for(let i=0;i<n;i++)seq.push(v)}
    if(!seq.length){toast("Type a pattern first");return}const from=$("rp-from").value||c.start;let i=0;
    c.dates.filter(d=>d>=from).forEach(d=>{const v=seq[i++%seq.length];host.querySelector('[data-rd="'+d+'"]').value=v});toast("Pattern filled. Check it, then tap Save rota")};
  $("rcSave").onclick=async()=>{const days={};host.querySelectorAll("[data-rd]").forEach(el=>{if(el.value!==sel[el.dataset.rd])days[el.dataset.rd]=el.value});
    if(!Object.keys(days).length){close();toast("No changes");return}close();if(await saveRota(gid,days)){toast("Rota saved for "+gName(g)+" ("+Object.keys(days).length+" day"+(Object.keys(days).length===1?"":"s")+")");render()}}}
function viewGroups(){const gs=allGroups().filter(g=>S.sheet==="All"||g.sheet===S.sheet);const eg=editingGroup||{};const w=S.canWrite;
  let h=w?'<form class="form" id="grpForm"><label>Sheet<select id="g-sheet" '+(editingGroup?"disabled":"")+'>'+SHEETS.map(x=>'<option'+((eg.sheet||(S.sheet!=="All"?S.sheet:""))===x?" selected":"")+'>'+esc(x)+'</option>').join("")+'</select></label>'+
   '<label>Section (optional)<input id="g-section" placeholder="For example Pit" value="'+esc(eg.section&&eg.section!==eg.sheet?eg.section:"")+'"></label>'+
   '<label>Group name<input id="g-shift" required placeholder="For example Shift A" value="'+esc(eg.shift||"")+'"></label>'+
   '<div style="display:flex;gap:8px;align-items:end;flex-wrap:wrap"><button class="btn gold" type="submit">'+(editingGroup?"Save name":"Add group")+'</button>'+(editingGroup?'<button class="btn" type="button" id="cancelGrp">Cancel</button>':"")+'</div></form>':"";
  h+=chipsBar("");
  if(!gs.length)h+='<div class="empty"><b>No shift groups yet</b>Add a group above, then put employees in it.</div>';
  gs.forEach(g=>{const mem=activeEmps().filter(x=>x.group===g.id).sort((a,b)=>a.order-b.order);const others=allGroups().filter(o=>o.sheet===g.sheet&&o.id!==g.id);
    h+='<section class="group"><div class="ghead"><span class="gname">'+esc(g.sheet+" · "+gName(g))+'</span><span class="note">'+mem.length+' '+(mem.length===1?"person":"people")+'</span>'+
     (w?'<span style="margin-left:auto;display:flex;gap:6px"><button class="btn sm" data-grota="'+esc(g.id)+'">Edit rota</button><button class="btn sm" data-grename="'+esc(g.id)+'">Rename</button><button class="btn sm" data-gremove="'+esc(g.id)+'">Remove</button></span>':"")+'</div>';
    h+=mem.length?'<div class="tbl-wrap"><table><tbody>'+mem.map(x=>'<tr><td>'+esc(x.name)+'<br><span class="note">'+esc([x.number,x.title].filter(Boolean).join(" · "))+'</span></td><td style="width:1%;white-space:nowrap">'+
      (w&&others.length?'<label class="note" style="display:block">Move to<select data-gmove="'+esc(x.id)+'"><option value="">Choose group</option>'+others.map(o=>'<option value="'+esc(o.id)+'">'+esc(gName(o))+'</option>').join("")+'</select></label>':"")+'</td></tr>').join("")+'</tbody></table></div>'
      :'<p class="note" style="margin:8px 0">Nobody in this group yet. Move people here from another group, or add a new employee to it on the Employees tab.</p>';
    h+='</section>'});
  h+='<p class="note">Employees can be moved to another group on the same sheet. Their marks stay with them.</p>';
  main.innerHTML=h;
  const f=$("grpForm");if(f)f.onsubmit=ev=>{ev.preventDefault();const shift=$("g-shift").value.trim(),section=$("g-section").value.trim();if(!shift){toast("Type the group name");return}
    if(editingGroup){const id=editingGroup.id,sheet=editingGroup.sheet,sec=section||sheet;editingGroup=null;
      S.groupDefs=(S.groupDefs||[]).filter(x=>x.id!==id).concat([{id,sheet,section:sec,shift}]);
      saveEmps((S.emps||[]).map(x=>x.group===id?Object.assign({},x,{section:sec,shift}):x)).then(()=>toast("Group renamed"));return}
    const sheet=$("g-sheet").value,sec=section||sheet;
    if(allGroups().some(x=>x.sheet===sheet&&gName(x).toLowerCase()===gName({sheet,section:sec,shift}).toLowerCase())){toast("That group is already on "+sheet);return}
    S.groupDefs=(S.groupDefs||[]).concat([{id:sheet+"|"+sec+"|"+shift+"|"+Date.now().toString(36),sheet,section:sec,shift}]);
    saveEmps(S.emps||[]).then(()=>toast(gName({sheet,section:sec,shift})+" added to "+sheet))};
  const c=$("cancelGrp");if(c)c.onclick=()=>{editingGroup=null;render()};
  main.querySelectorAll("[data-grota]").forEach(b=>b.onclick=()=>openRotaCycle(b.dataset.grota));
  main.querySelectorAll("[data-grename]").forEach(b=>b.onclick=()=>{editingGroup=allGroups().find(x=>x.id===b.dataset.grename);render();window.scrollTo({top:0})});
  main.querySelectorAll("[data-gremove]").forEach(b=>b.onclick=()=>{const g=allGroups().find(x=>x.id===b.dataset.gremove);const n=activeEmps().filter(x=>x.group===g.id).length;
    if(n){toast("Move the "+n+" "+(n===1?"person":"people")+" in "+gName(g)+" to another group first");return}
    if(!confirm("Remove the group "+gName(g)+" from "+g.sheet+"?"))return;S.groupDefs=(S.groupDefs||[]).filter(x=>x.id!==g.id).concat([{id:g.id,sheet:g.sheet,removed:true}]);
    saveEmps(S.emps||[]).then(()=>toast(gName(g)+" removed"))});
  main.querySelectorAll("[data-gmove]").forEach(sel=>sel.onchange=()=>{const to=allGroups().find(x=>x.id===sel.value);if(!to)return;const list=(S.emps||[]).slice();const x=list.find(y=>y.id===sel.dataset.gmove);
    const order=orderFor(to.id,to.sheet,list,x.id);saveEmps(list.map(y=>y.id===x.id?Object.assign({},y,{group:to.id,section:to.section,shift:to.shift,order}):y)).then(()=>toast(x.name+" moved to "+gName(to)))});
}
/* duplicate employees: same employee number, or the same name (words in any order, ignoring commas and spaces) */
const dupKey=n=>String(n||"").toLowerCase().replace(/[^a-z\s,]/g,"").split(/[\s,]+/).filter(Boolean).sort().join(" ");
function dupGroups(){const A=activeEmps(),par=new Map(A.map(e=>[e.id,e.id])),f=x=>{while(par.get(x)!==x)x=par.get(x);return x},u=(a,b)=>par.set(f(a),f(b));
  const by=(k,e)=>{if(!k)return;const o=by.m.get(k);if(o)u(e.id,o);else by.m.set(k,e.id)};by.m=new Map();
  A.forEach(e=>{const n=String(e.number||"").trim().toUpperCase();if(n)by("#"+n,e)});A.forEach(e=>by("n:"+dupKey(e.name),e));
  const G=new Map();A.forEach(e=>{const r=f(e.id);(G.get(r)||G.set(r,[]).get(r)).push(e)});return [...G.values()].filter(g=>g.length>1)}
function marksCount(e){let n=0;for(const m in S.mm){const sd=(S.mm[m]||{})[SHEET_ID[e.sheet]];n+=Object.values(((sd||{}).marks||{})[e.id]||{}).filter(Boolean).length}return n}
function dupPanel(){const G=dupGroups();if(!G.length)return "";
  return '<div class="list" style="margin-bottom:14px;border:2px solid var(--bad,#b3402f)"><h3>Possible duplicates ('+G.length+')</h3><p class="note" style="margin:0 0 8px">The same person is on the register more than once (same employee number or same name). Press <b>Keep this one</b> on the correct entry: the other entries are taken off, and any days marked on them are moved onto the one you keep where it has no mark. They do not show as removed in History.</p>'+
    G.map(g=>'<div style="margin:10px 0"><table class="tbl"><thead><tr><th>Employee</th><th>Number</th><th>Sheet · group</th><th>Job title</th><th>Days marked</th><th></th></tr></thead><tbody>'+g.map(e=>'<tr><td><b>'+esc(e.name)+'</b></td><td>'+esc(e.number||"")+'</td><td>'+esc(e.sheet+" · "+groupLabel(e))+'</td><td>'+esc(e.title||"")+'</td><td>'+marksCount(e)+'</td><td>'+(S.canWrite?'<button class="btn sm gold" data-keepdup="'+esc(e.id)+'" data-dups="'+esc(g.map(x=>x.id).join(","))+'">Keep this one</button>':"")+'</td></tr>').join("")+'</tbody></table></div>').join("")+'</div>'}
async function keepDup(keepId,ids){if(!S.canWrite){toast("You can view this register but not change it.");return}
  const L=S.emps||[],keep=L.find(e=>e.id===keepId),drop=L.filter(e=>ids.includes(e.id)&&e.id!==keepId);if(!keep||!drop.length)return;
  if(!confirm("Keep "+keep.name+" ("+keep.sheet+(keep.number?", "+keep.number:"")+") and take off "+drop.length+" duplicate"+(drop.length>1?"s":"")+"?\nDays marked on the duplicate are moved onto the one you keep where it has no mark."))return;
  toast("Merging…");const from=addDays(iso(new Date()),-400),to=addDays(iso(new Date()),60);try{await loadRange(from,to)}catch(e){console.warn(e)}
  let moved=0;const kid=SHEET_ID[keep.sheet];
  for(const m in S.mm){const mo=S.mm[m]||{},patch={marks:{},ot:{},ottype:{}};
    for(const d of drop){const sd=mo[SHEET_ID[d.sheet]]||{};for(const f of ["marks","ot","ottype"]){const src=(sd[f]||{})[d.id]||{},have=((mo[kid]||{})[f]||{})[keep.id]||{};
      for(const day in src){if(src[day]!==""&&src[day]!=null&&(have[day]===undefined||have[day]===""||have[day]===null)&&patch[f][day]===undefined){patch[f][day]=src[day];if(f==="marks")moved++}}}}
    const up={};for(const f of ["marks","ot","ottype"])if(Object.keys(patch[f]).length)up[f]={[keep.id]:patch[f]};
    if(Object.keys(up).length){const sdk=mo[kid]=mo[kid]||{marks:{},ot:{},ottype:{}};for(const f in up){sdk[f]=sdk[f]||{};sdk[f][keep.id]=Object.assign({},sdk[f][keep.id]||{},up[f][keep.id])}
      try{await S.db.doc("months/"+m+"/sheets/"+kid).update(Object.assign({sheet:keep.sheet,month:m,updatedBy:email,updatedAt:Date.now()},up))}catch(e){toast("Not saved: "+(e.message||"try again"));return}}}
  const aliases=[...new Set((keep.aliases||[]).concat(drop.map(d=>d.name)).filter(n=>n&&n!==keep.name))],nums=[...new Set((keep.altNumbers||[]).concat(drop.map(d=>d.number)).filter(n=>n&&n!==keep.number))];
  const list=L.filter(e=>!drop.some(d=>d.id===e.id)).map(e=>e.id===keep.id?Object.assign({},e,{aliases,altNumbers:nums,number:e.number||nums[0]||"",title:e.title||drop.map(d=>d.title).find(Boolean)||"",dept:e.dept||drop.map(d=>d.dept).find(Boolean)||""}):e);
  await saveEmps(list);toast(drop.length+" duplicate"+(drop.length>1?"s":"")+" taken off"+(moved?", "+moved+" marked days moved onto "+keep.name:"")+".")}
function viewPeople(){const list=filtered().sort((a,b)=>SHEETS.indexOf(a.sheet)-SHEETS.indexOf(b.sheet)||a.order-b.order);const e=editing||{};
  const groups=allGroups().map(g=>({group:g.id,sheet:g.sheet,section:g.section,shift:g.shift}));
  let h='<form class="form" id="empForm"><label>Sheet<select id="f-sheet" required>'+SHEETS.map(s=>'<option'+(e.sheet===s?" selected":"")+'>'+esc(s)+'</option>').join("")+'</select></label>'+
   '<label>Shift group<select id="f-group">'+groups.map(g=>'<option value="'+esc(g.group)+'" data-sheet="'+esc(g.sheet)+'"'+(e.group===g.group?" selected":"")+'>'+esc(g.sheet+" · "+groupLabel(g))+'</option>').join("")+'</select></label>'+
   '<label>Name (Surname, First names)<input id="f-name" required value="'+esc(e.name||"")+'"></label><label>Employee number<input id="f-number" value="'+esc(e.number||"")+'"></label>'+
   '<label>Department<input id="f-dept" value="'+esc(e.dept||"")+'"></label><label>Job title<input id="f-title" value="'+esc(e.title||"")+'"></label>'+
   '<div style="display:flex;gap:8px;align-items:end;flex-wrap:wrap"><button class="btn gold" type="submit">'+(editing?"Save changes":"Add employee")+'</button>'+(editing?'<button class="btn" type="button" id="cancelEdit">Cancel</button>':"")+'</div></form>';
  h+=chipsBar('<input class="search" id="q" type="search" placeholder="Search name, number or job" value="'+esc(S.q)+'">');
  h+='<div class="tbl-wrap"><table><thead><tr><th>Employee</th><th>Number</th><th>Group</th><th>Job title</th><th></th></tr></thead><tbody>'+
   list.map(x=>'<tr><td>'+esc(x.name)+'</td><td>'+esc(x.number)+'</td><td>'+esc(x.sheet+" · "+groupLabel(x))+'</td><td>'+esc(x.title)+'</td><td><button class="btn sm" data-edit="'+esc(x.id)+'">Edit</button> <button class="btn sm" data-remove="'+esc(x.id)+'">Remove</button></td></tr>').join("")+'</tbody></table></div>';
  h+='<p class="note">Removing takes the person off the register after their last working day. Their days up to then stay in the reports and Excel downloads.</p>';
  const c=cycleOf(S.date);const flagged=activeEmps().map(x=>{const d=c.dates.find(d=>["RESIGNED","ABSCONDED","DISMISSED","DECEASED"].includes(rawMark(x,d)));return d?{x,d,code:rawMark(x,d)}:null}).filter(Boolean);
  h=dupPanel()+h;
  if(flagged.length)h='<div class="list" style="margin-bottom:14px"><h3>Marked resigned, absconded, dismissed or deceased this cycle ('+flagged.length+')</h3><p class="note" style="margin:0 0 8px">These people are still on the register. Remove them to take them off the marking list.</p><ul>'+
    flagged.map(f=>'<li>'+esc(f.x.name)+' <span class="note">'+esc(f.x.sheet)+' · '+esc(f.code)+' from '+short(f.d)+'</span> <button class="btn sm" data-remove="'+esc(f.x.id)+'" data-reason="'+esc(f.code)+'" data-last="'+addDays(f.d,-1)+'">Remove</button></li>').join("")+'</ul></div>'+h;
  if(S.canWrite)h+='<div class="list" style="margin-top:18px"><h3>Update from the register Excel</h3><p class="note" style="margin:0 0 8px">Pick the register (one tab per sheet: rompad, support fleet, shared services, drilling, intern). The app puts everyone in the shift group they are under (SHIFT A, SHIFT B …), sets each group\'s DAY / NIGHT / OFF rota from the group row, and copies the marks for every date in the file. Anyone not yet on the register is added. Empty cells in the Excel do not clear marks in the app.</p><button class="btn gold" id="impReg">Choose register Excel</button> <button class="btn" id="impOld">Import old timesheets (marks only)</button><input type="file" id="regFile" accept=".xlsx,.xls,.xlsb" hidden><input type="file" id="oldFile" accept=".xlsx,.xls,.xlsb" hidden><br><span class="note">Old timesheets: only the marks are copied. Shift groups and the current register stay as they are; people who are no longer on the register are kept as removed, so past reports include them.</span><p class="note" id="regMsg" role="status">'+esc(regNote)+'</p></div>';
  if(S.canWrite)h+='<div class="list" style="margin-top:18px"><h3>Add employees from Excel</h3><p class="note" style="margin:0 0 8px">Pick an Excel list with the columns Employee, Number, Job Title and Department. People already on the register (same employee number) are skipped. Department decides the sheet: Rompad, Drilling and Shared Services go to their own sheet; Buses, Water Bowsers, Low Beds and TLB go to Support Fleet; Corporate goes to Shared Services. Each department becomes a shift group.</p><button class="btn gold" id="impEmps">Choose Excel file</button><input type="file" id="impFile" accept=".xlsx,.xls,.xlsb,.csv" hidden><p class="note" id="impMsg" role="status">'+esc(impNote)+'</p></div>';
  if(S.canWrite)h+='<div class="list" style="margin-top:18px"><h3>Import overtime from Excel</h3><p class="note" style="margin:0 0 8px">Pick the overtime report (columns: Name and Surname, Standby / Unplanned, Date, Payable Hours on Time Sheet). Each line puts the payable hours as overtime on that person\'s day, in whatever month the date falls. Lines with no hours (for example "no clockings") are skipped and listed. Importing the same file again replaces those days, it does not add them twice.</p><button class="btn gold" id="impOt">Choose overtime Excel</button><input type="file" id="otFile" accept=".xlsx,.xls,.xlsb,.csv" hidden><p class="note" id="otMsg" role="status" style="white-space:pre-line">'+esc(otNote)+'</p></div>';
  if(isAdmin()&&(S.emps||[]).length)h+='<div class="list" style="margin-top:18px"><h3>Remove all employees</h3><p class="note" style="margin:0 0 8px">Clears the whole employee list ('+(S.emps||[]).length+' people, including removed ones) so you can add your own. A copy of the data file downloads first, so you can load it back under Users if you need it. Marks already saved are not deleted.</p><button class="btn" id="clearEmps" style="border-color:var(--bad);color:var(--bad)">Remove all employees</button></div>';
  main.innerHTML=h;
  const rb=$("impReg");if(rb){rb.onclick=()=>$("regFile").click();$("regFile").onchange=ev=>{const f=ev.target.files[0];ev.target.value="";if(f)importRegister(f)}}
  const obt=$("impOld");if(obt){obt.onclick=()=>$("oldFile").click();$("oldFile").onchange=ev=>{const f=ev.target.files[0];ev.target.value="";if(f)importRegister(f,true)}}
  const ob=$("impOt");if(ob){ob.onclick=()=>$("otFile").click();$("otFile").onchange=ev=>{const f=ev.target.files[0];ev.target.value="";if(f)importOvertime(f)}}
  const ib=$("impEmps");if(ib){ib.onclick=()=>$("impFile").click();$("impFile").onchange=ev=>{const f=ev.target.files[0];ev.target.value="";if(f)importEmps(f)}}
  const ce=$("clearEmps");if(ce)ce.onclick=async()=>{const n=(S.emps||[]).length;
    if(!confirm("Remove all "+n+" employees from TKR Attendance? A copy of the data file downloads first."))return;
    if(prompt('Type REMOVE to confirm')!=="REMOVE"){toast("Nothing removed");return}
    try{await saveBackup()}catch(e){toast("The copy could not be made, so nothing was removed");return}
    await saveEmps([]);toast("All employees removed. Add your own on this tab.")};
}
// Add employees from an Excel list (Employee, Number, Job Title, Department). Skips numbers already on the register.
const DEPT_SHEET={"rompad":"Rompad","drilling":"Drilling","shared services":"Shared Services","interns":"Interns","support fleet":"Support Fleet","corporate":"Shared Services",
  "buses":"Support Fleet","bus":"Support Fleet","water bowsers":"Support Fleet","water bowser":"Support Fleet","low beds":"Support Fleet","lowbeds":"Support Fleet","low bed":"Support Fleet","tlb":"Support Fleet"};
let impNote="",regNote="",otNote="";
// ---------- update from the register Excel (same layout as the TKR timesheet) ----------
const TAB_SHEET=t=>{t=t.toLowerCase().replace(/[^a-z]/g,"");if(t.startsWith("rompad"))return "Rompad";if(t.startsWith("support"))return "Support Fleet";if(t.startsWith("drill"))return "Drilling";
  if(t.startsWith("share"))return "Shared Services";if(t.startsWith("intern"))return "Interns";return null};
const CODE_FIX={"P":"P","PRESENT":"P","SICK":"SICK","FAMILY REPONSIBILITY":"FAMILY RESPONSIBILITY","FAMILY RESPONSIBILTY":"FAMILY RESPONSIBILITY","FAMALY EMERGENCY":"FAMILY EMERGENCY","SITE OFFICE MEETING":"SITE OFFICE","SITE OFFICE":"SITE OFFICE","PIT LICENSE":"PIT LICENCE"};
const ROTA_VAL=v=>{v=String(v||"").trim().toUpperCase();return v==="DAY"||v==="NIGHT"||v==="OFF"?v:""};
const serialISO=n=>{const d=new Date(Math.round((n-25569)*864e5));return d.getUTCFullYear()+"-"+pad(d.getUTCMonth()+1)+"-"+pad(d.getUTCDate())};
const markVal=v=>{if(v===""||v==null)return"";if(typeof v==="number"){if(v>0&&v<1){const m=Math.round(v*1440);return pad(Math.floor(m/60)%24)+":"+pad(m%60)}return""}
  let t=String(v).replace(/\s+/g," ").trim().toUpperCase();if(!t)return"";if(/^\d{1,2}:\d{2}(:\d{2})?$/.test(t)){const [h,mi]=t.split(":");return pad(+h)+":"+mi}return CODE_FIX[t]||t};
const normName=n=>String(n||"").toLowerCase().replace(/\s*,\s*/g,",").replace(/\s+/g," ").trim();
function trimRef(ws){let maxR=0,maxC=0;for(const k in ws){if(k[0]==="!")continue;const c=XLSX.utils.decode_cell(k);if(ws[k].v===undefined||ws[k].v===null||ws[k].v==="")continue;if(c.r>maxR)maxR=c.r;if(c.c>maxC)maxC=c.c}
  ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR,c:maxC}})}
function parseRegister(wb){const out=[];
  wb.SheetNames.forEach(tab=>{const sheet=TAB_SHEET(tab);if(!sheet)return;const ws=wb.Sheets[tab];trimRef(ws);
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
    let dateCols=null,section=sheet,shift="",gkey="",seen={},rota=null,order=0;
    const txt=(r,i)=>String(r[i]??"").trim();
    rows.forEach(r=>{const dc=[];r.forEach((v,i)=>{if(typeof v==="number"&&v>40000&&v<60000)dc.push([i,serialISO(v)])});
      if(dc.length>=5){dateCols=dc;return}
      const b=txt(r,1),c=txt(r,2),d=txt(r,3),e=txt(r,4);if(!b||b.toLowerCase()==="employee")return;
      const isEmp=!!(c||d||e);
      if(!isEmp){const rv=dateCols?dateCols.map(([i,dt])=>[dt,ROTA_VAL(r[i])]).filter(x=>x[1]):[];
        const looksShift=/shift|monday|week|day|night/i.test(b);
        if(rv.length&&looksShift){shift=b}else{section=b.toLowerCase()===sheet.toLowerCase()?sheet:b;shift=rv.length?"":""}
        const base=sheet+"|"+section+"|"+shift;seen[base]=(seen[base]||0)+1;gkey=base+(seen[base]>1?"|"+seen[base]:"");
        rota=rv.length?Object.fromEntries(rv):null;out.push({kind:"group",sheet,section,shift,gid:gkey,rota});return}
      if(/^vacant$/i.test(b))return;
      if(!gkey){const base=sheet+"|"+section+"|"+shift;gkey=base}
      const marks={};(dateCols||[]).forEach(([i,dt])=>{const v=markVal(r[i]);if(v)marks[dt]=v});
      out.push({kind:"emp",sheet,section,shift,gid:gkey,name:b,number:c,dept:d,title:e,marks,row:order++})})});
  return out}
async function importRegister(file,old){const say=t=>{regNote=t;const m=$("regMsg");if(m)m.textContent=t};say("Reading "+file.name+"…");
  try{if(!window.XLSX)throw new Error("The Excel reader has not loaded yet. Try again.");
    const items=parseRegister(XLSX.read(await file.arrayBuffer(),{type:"array"}));const emps=items.filter(x=>x.kind==="emp");
    if(!emps.length){say("No employees found. Check that the tabs are named rompad, support fleet, shared services, drilling and intern.");return}
    let list=(S.emps||[]).map(x=>Object.assign({},x));const byNum={},byName={};
    list.forEach(x=>{[x.number].concat(x.altNumbers||[]).forEach(v=>{const n=String(v||"").trim().toUpperCase();if(n)(byNum[n]=byNum[n]||[]).push(x)});[x.name].concat(x.aliases||[]).forEach(v=>(byName[normName(v)]=byName[normName(v)]||[]).push(x))});
    const used=new Set(),added=[],moved=[],stillOff=[],marks={},rotaGroups={};let nMarks=0,seq=0;
    items.filter(x=>x.kind==="group"&&x.rota).forEach(g=>{rotaGroups[g.gid]={sheet:g.sheet,section:g.section,shift:g.shift,rota:g.rota}});
    const sheetBase=sh=>(SHEETS.indexOf(sh)+1)*10000;
    for(const r of emps){const n=r.number.toUpperCase();let x=(n&&byNum[n]&&byNum[n].length===1?byNum[n][0]:null)||((byName[normName(r.name)]||[]).find(y=>old||!used.has(y.id))||null);
      if(x&&used.has(x.id)&&!old)x=null;
      if(!x){const sur=normName(r.name).split(",")[0],first=(normName(r.name).split(",")[1]||"").split(" ")[0];
        const c=list.filter(y=>(old||!used.has(y.id))&&normName(y.name).split(",")[0]===sur&&((n&&String(y.number||"").toUpperCase().startsWith(n))||(first.length>=3&&(normName(y.name).split(",")[1]||"").startsWith(first.slice(0,4)))));
        if(c.length===1)x=c[0]}
      const g={sheet:r.sheet,group:r.gid,section:r.section,shift:r.shift,order:sheetBase(r.sheet)+(seq++)};
      if(old){   // old timesheet: copy the marks only
        if(!x){const ds=Object.keys(r.marks).sort(),res=ds.find(d=>r.marks[d]==="RESIGNED"),last=ds.filter(d=>r.marks[d]!=="RESIGNED").pop()||ds[0]||"";
          x=Object.assign({id:(r.number||r.name).replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"")+"-"+Date.now().toString(36)+seq,active:false,historic:true,name:r.name,number:r.number,dept:r.dept,title:r.title,
            leftDate:res||(last?addDays(last,1):""),leftReason:res?"Resigned":"Not on the current register (old timesheet)"},g,{order:900000+seq});list.push(x);added.push(r.name);
          byName[normName(r.name)]=[x];if(r.number)byNum[r.number.toUpperCase()]=[x]}
        else if(x.historic&&!x.active){const ds=Object.keys(r.marks).sort(),res=ds.find(d=>r.marks[d]==="RESIGNED"),last=ds.filter(d=>r.marks[d]!=="RESIGNED").pop();
          const nl=res||(last?addDays(last,1):"");if(nl&&(!x.leftDate||nl>x.leftDate)&&!(x.leftReason==="Resigned"&&!res))x.leftDate=nl;if(res)x.leftReason="Resigned"}
        used.add(x.id);const sid=SHEET_ID[x.sheet]||SHEET_ID[r.sheet];
        for(const dt in r.marks){if(x.active===false&&x.leftDate&&dt>=x.leftDate&&!x.historic)continue;const m=dt.slice(0,7);((marks[m]=marks[m]||{})[sid]=marks[m][sid]||{sheet:x.sheet||r.sheet,month:m,marks:{}}).marks[x.id]=Object.assign(marks[m][sid].marks[x.id]||{},{[dt]:r.marks[dt]});nMarks++}
        continue}
      if(x&&x.active===false){   // removed people stay removed even if the Excel still lists them; only their marks from before they left are copied
        stillOff.push(x.name);used.add(x.id);const sid=SHEET_ID[x.sheet]||SHEET_ID[r.sheet];
        for(const dt in r.marks){if(x.leftDate&&dt>=x.leftDate)continue;const m=dt.slice(0,7);((marks[m]=marks[m]||{})[sid]=marks[m][sid]||{sheet:x.sheet||r.sheet,month:m,marks:{}}).marks[x.id]=Object.assign(marks[m][sid].marks[x.id]||{},{[dt]:r.marks[dt]});nMarks++}
        continue}
      if(x){if(x.group!==g.group)moved.push(x.name);Object.assign(x,g,{active:true});if(!x.number&&r.number)x.number=r.number;if(!x.title&&r.title)x.title=r.title;if(!x.dept&&r.dept)x.dept=r.dept;delete x.leftDate;delete x.leftReason}
      else{x=Object.assign({id:(r.number||r.name).replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"")+"-"+Date.now().toString(36)+seq,active:true,name:r.name,number:r.number,dept:r.dept,title:r.title},g);list.push(x);added.push(r.name)}
      used.add(x.id);
      for(const dt in r.marks){const m=dt.slice(0,7),sid=SHEET_ID[r.sheet];((marks[m]=marks[m]||{})[sid]=marks[m][sid]||{sheet:r.sheet,month:m,marks:{}}).marks[x.id]=Object.assign(marks[m][sid].marks[x.id]||{},{[dt]:r.marks[dt]});nMarks++}}
    if(old){const ds=emps.flatMap(e=>Object.keys(e.marks)).sort();
      if(!confirm("Copy the marks from "+file.name+"?\n\n"+nMarks+" marks from "+(ds[0]?fmtD(ds[0]):"")+" to "+(ds.length?fmtD(ds.at(-1)):"")+" for "+used.size+" employees."+(added.length?"\n"+added.length+" people are no longer on the register: they are kept as removed so the reports for those dates include them.":"")+"\n\nShift groups and the current register do not change."))return say("Not imported.");
      say("Saving…");await saveEmps(list);
      for(const m in marks)for(const sid in marks[m]){const d=marks[m][sid];await S.db.doc("months/"+m+"/sheets/"+sid).update({sheet:d.sheet,month:m,marks:d.marks,updatedBy:email,updatedAt:Date.now()})}
      say("Done: "+nMarks+" marks copied for "+used.size+" employees"+(added.length?"; "+added.length+" people who are no longer on the register kept as removed ("+added.join(", ")+")":"")+".");toast("Old timesheets imported");return}
    const notIn=list.filter(x=>x.active!==false&&!used.has(x.id));notIn.forEach(x=>{x.order=900000+(x.order||0)});
    const dates=Object.values(rotaGroups).flatMap(g=>Object.keys(g.rota)).sort();
    if(!confirm("Update TKR Attendance from "+file.name+"?\n\n"+used.size+" employees in "+new Set(emps.map(x=>x.gid)).size+" shift groups ("+added.length+" new, "+moved.length+" moved to a new group)"+(stillOff.length?"\n"+stillOff.length+" removed employees are still in the Excel and stay removed":"")+"\n"+nMarks+" marks"+(dates.length?"\nRotas for "+short(dates[0])+" – "+short(dates.at(-1)):"")))return say("Nothing changed.");
    say("Saving…");S.groupDefs=(S.groupDefs||[]).filter(d=>!d.removed);
    await saveEmps(list);
    if(dates.length){const c=cycleOf(dates[0]);const id="r-"+c.start;await S.db.doc("cycles/"+id).update({start:c.start,end:c.end,groups:rotaGroups})}
    for(const m in marks)for(const sid in marks[m]){const d=marks[m][sid];await S.db.doc("months/"+m+"/sheets/"+sid).update({sheet:d.sheet,month:m,marks:d.marks,updatedBy:email,updatedAt:Date.now()})}
    say("Done: "+used.size+" employees put in their shift groups, "+nMarks+" marks copied"+(added.length?", "+added.length+" new employees added ("+added.join(", ")+")":"")+"."+(stillOff.length?" "+stillOff.length+" removed employees are still in this Excel and stay removed: "+stillOff.join(", ")+".":"")+(notIn.length?" "+notIn.length+" people on the register are not in this Excel and were left as they are: "+notIn.map(x=>x.name).join(", ")+".":""));
    toast("Register updated")}
  catch(e){say(denied(e)?"Your sign-in is not allowed to change the register.":"Could not read "+file.name+": "+(e.message||"try again"))}}
/* Overtime report: NO. | NAME AND SURNAME | STANDBY/ UNPLANNED | DATE | PAYABLE HOURS ON TIME SHEET (hours as 12:10 = 12 h 10 min) */
const otName=n=>String(n||"").toLowerCase().replace(/[^a-z]+/g," ").trim();
function xlDate(v){if(typeof v==="number"&&v>20000){const d=new Date(Date.UTC(1899,11,30)+Math.round(v)*864e5);return d.getUTCFullYear()+"-"+pad(d.getUTCMonth()+1)+"-"+pad(d.getUTCDate())}
  const t=String(v||"").trim();let m=t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(m)return m[1]+"-"+pad(+m[2])+"-"+pad(+m[3]);m=t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);if(m)return m[3]+"-"+pad(+m[2])+"-"+pad(+m[1]);return ""}
function xlHours(v){if(typeof v==="number"){if(v<=0)return 0;return v<1?Math.round(v*24*60)/60:Math.round(v*100)/100}
  const t=String(v||"").trim().toLowerCase();let m=t.match(/^(\d{1,2})[:h](\d{1,2})/);if(m)return +m[1]+(+m[2])/60;m=t.match(/^(\d+(?:[.,]\d+)?)\s*(h|hrs?|hours?)?$/);if(m)return parseFloat(m[1].replace(",","."));return 0}
async function importOvertime(file){const say=t=>{otNote=t;const m=$("otMsg");if(m)m.textContent=t};say("Reading "+file.name+"…");
  try{if(!window.XLSX)throw new Error("The Excel reader has not loaded yet. Try again.");
    const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:"",raw:true});
    const low=c=>String(c).trim().toLowerCase();
    const hi=rows.findIndex(r=>r.some(c=>/name/.test(low(c)))&&r.some(c=>low(c)==="date"));
    if(hi<0){say("No header row with Name and Date was found in "+file.name+".");return}
    const H=rows[hi].map(low),cN=H.findIndex(c=>/name/.test(c)),cD=H.indexOf("date"),cH=H.findIndex(c=>/hour/.test(c)),cT=H.findIndex(c=>/standby|unplanned|planned|type/.test(c));
    if(cH<0){say("No hours column (Payable Hours) was found in "+file.name+".");return}
    const byName=new Map();(S.emps||[]).forEach(e=>{const n=otName(e.name);byName.set(n,e);byName.set(n.split(" ").reverse().join(" "),e);const w=n.split(" ");if(w.length>2)byName.set(w[0]+" "+w[w.length-1],e)});
    const find=n=>{const k=otName(n);if(byName.has(k))return byName.get(k);const w=k.split(" ");return byName.get(w[0]+" "+w[w.length-1])||null};
    const lines=[],notFound=new Set(),noHours=[],badDate=[];
    rows.slice(hi+1).forEach(r=>{const name=String(r[cN]||"").trim();if(!name)return;const d=xlDate(r[cD]),hrs=xlHours(r[cH]),type=cT>=0?String(r[cT]||"").trim():"";
      if(!d){badDate.push(name);return}
      if(!hrs){noHours.push(name+" "+short(d)+(String(r[cH]||"").trim()?" ("+String(r[cH]).trim()+")":""));return}
      const e=find(name);if(!e){notFound.add(name);return}
      lines.push({e,d,hrs:Math.round(hrs*100)/100,type})});
    if(!lines.length){say("No overtime could be put on the register."+(notFound.size?"\nNot on the register: "+[...notFound].join(", ")+".":"")+(noHours.length?"\nNo hours: "+noHours.join(", ")+".":""));return}
    const people=new Set(lines.map(x=>x.e.id)).size,tot=lines.reduce((a,x)=>a+x.hrs,0);
    if(!confirm("Put "+lines.length+" overtime lines ("+fmtH(tot)+" h, "+people+" people) on the register?"+(notFound.size?"\n\nNot on the register, skipped: "+[...notFound].join(", "):"")+(noHours.length?"\n\nNo hours, skipped: "+noHours.length+" lines":"")))return say("Nothing imported.");
    const docs={};lines.forEach(x=>{const m=monthOf(x.d),id=SHEET_ID[x.e.sheet],k=m+"/"+id;const o=docs[k]=docs[k]||{m,id,sheet:x.e.sheet,ot:{},ottype:{}};
      (o.ot[x.e.id]=o.ot[x.e.id]||{})[x.d]=x.hrs;if(x.type)(o.ottype[x.e.id]=o.ottype[x.e.id]||{})[x.d]=x.type});
    for(const o of Object.values(docs)){await S.db.doc("months/"+o.m+"/sheets/"+o.id).update({sheet:o.sheet,month:o.m,ot:o.ot,ottype:o.ottype,updatedBy:email,updatedAt:Date.now()});
      const mo=S.mm[o.m];if(mo){const sd=mo[o.id]=mo[o.id]||{marks:{},ot:{}};for(const k in o.ot)sd.ot[k]=Object.assign({},sd.ot[k]||{},o.ot[k])}}
    say(lines.length+" overtime lines imported: "+fmtH(tot)+" h for "+people+" people."+(notFound.size?"\nNot on the register (add them, then import again): "+[...notFound].join(", ")+".":"")+(noHours.length?"\nNo hours, skipped: "+noHours.join(", ")+".":"")+(badDate.length?"\nNo date, skipped: "+badDate.join(", ")+".":""));
    toast("Overtime imported");render()}
  catch(e){say(denied(e)?"Your sign-in is not allowed to change the register.":"Could not read "+file.name+": "+(e.message||"try again"))}}
async function importEmps(file){const say=t=>{impNote=t;const m=$("impMsg");if(m)m.textContent=t};say("Reading "+file.name+"…");
  try{if(!window.XLSX)throw new Error("The Excel reader has not loaded yet. Try again.");
    const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
    const hi=rows.findIndex(r=>r.some(c=>String(c).trim().toLowerCase()==="employee")&&r.some(c=>String(c).trim().toLowerCase()==="number"));
    if(hi<0){say("No header row with Employee and Number was found in "+file.name+".");return}
    const H=rows[hi].map(c=>String(c).trim().toLowerCase()),col=n=>H.indexOf(n);
    const cE=col("employee"),cN=col("number"),cT=[col("job title"),col("title"),col("job")].find(i=>i>=0),cD=[col("department"),col("dept")].find(i=>i>=0);
    let list=(S.emps||[]).slice();const have=new Set(list.map(x=>String(x.number||"").trim().toUpperCase()).filter(Boolean));
    const added=[],skipped=[],noDept=[];
    rows.slice(hi+1).forEach(r=>{const name=String(r[cE]||"").trim();if(!name)return;const number=String(r[cN]||"").trim();
      if(number&&have.has(number.toUpperCase())){skipped.push(name);return}
      const dept=cD>=0?String(r[cD]||"").trim():"",title=cT>=0?String(r[cT]||"").trim():"";
      const sheet=DEPT_SHEET[dept.toLowerCase()]||(SHEETS.includes(dept)?dept:"Shared Services");if(!dept)noDept.push(name);
      const section=dept&&dept!==sheet?dept:sheet;const gid=sheet+"|"+section+"|";
      const id=(number||name).replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"")+"-"+Date.now().toString(36)+added.length;
      list.push({id,order:orderFor(gid,sheet,list),active:true,sheet,group:gid,section,shift:"",name,number,dept,title});if(number)have.add(number.toUpperCase());added.push(name)});
    if(!added.length){say("Nobody new to add. "+skipped.length+" already on the register.");return}
    if(!confirm("Add "+added.length+" employees to TKR Attendance?"+(skipped.length?" "+skipped.length+" already on the register will be skipped.":"")))return say("Nothing added.");
    await saveEmps(list);
    say(added.length+" employees added."+(skipped.length?" "+skipped.length+" skipped (already on the register).":"")+(noDept.length?" No department, put on Shared Services: "+noDept.join(", ")+". Edit them to change the sheet.":""));
    toast(added.length+" employees added")}
  catch(e){say("Could not read "+file.name+": "+(e.message||"try again"))}}
async function saveEmps(list){if(!S.canWrite){toast("You can view this register but not change it.");return}
  try{const hist=(S.hist||[]).concat(empEvents(S.emps||[],list)).slice(-3000);
    await S.db.doc("config/employees").set({list,groups:S.groupDefs||[],history:hist});S.emps=list;S.hist=hist;render()}catch(e){if(denied(e)){S.canWrite=false;toast("You can view this register but not change it.")}else toast("Not saved: "+(e.message||"try again"))}}

// ---------- mark panel ----------
function openPanel(e){const v=markOf(e,S.date);const now=new Date();const hv=isHrs(v)?hrsOf(v):12;
  $("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true" aria-labelledby="pt"><h2 id="pt">'+esc(e.name)+'</h2><div class="sub">'+esc([e.number,e.title,groupLabel(e)].filter(Boolean).join(" · "))+' — '+nice(S.date)+'</div>'+
   '<div class="lbl">Total hours worked today (hh:mm)</div><div class="clock"><input type="text" id="tm" inputmode="numeric" maxlength="5" placeholder="12:00" value="'+hm(hv)+'"><button class="btn gold" id="saveTime">Save hours</button></div>'+
   '<div class="lbl">Overtime hours</div><div class="clock"><input type="number" id="ot" min="0" max="24" step="0.5" inputmode="decimal" placeholder="0" value="'+(otOf(e,S.date)||"")+'"><button class="btn gold" id="saveOt">Save overtime</button></div>'+
   '<div class="lbl">Or mark as</div><div class="codes">'+CODE_ORDER.map(c=>'<button class="k-'+KIND[c]+(v===c?" cur":"")+'" data-code="'+esc(c)+'">'+esc(c)+'</button>').join("")+'</div>'+
   '<div style="display:flex;gap:8px;justify-content:space-between;flex-wrap:wrap"><button class="btn" id="clr"'+(v?"":" disabled")+'>Clear mark</button><button class="btn" id="closeP">Close</button></div></div></div>';
  const close=()=>{$("sheetHost").innerHTML=""};
  $("scrim").onclick=ev=>{if(ev.target.id==="scrim")close()};$("closeP").onclick=close;
  $("tm").oninput=ev=>{let d=ev.target.value.replace(/\D/g,"").slice(0,4);ev.target.value=d.length>2?d.slice(0,2)+":"+d.slice(2):d};
  $("saveTime").onclick=()=>{const h=parseHm($("tm").value);if(isNaN(h)||h<=0||h>24){toast("Enter hours worked like 12:00 (up to 24:00)");return}setMark(e,hm(h)+" HRS");toast(hm(h)+" hours saved for "+e.name);close()};
  $("clr").onclick=()=>{setMark(e,"");close()};
  $("saveOt").onclick=()=>{const raw=$("ot").value.trim();const h=raw===""?0:parseFloat(raw);if(isNaN(h)||h<0||h>24){toast("Enter overtime hours between 0 and 24");return}setOt(e,h?h:"");toast(h?fmtH(h)+" h overtime saved for "+e.name:"Overtime cleared for "+e.name);close()};
  $("sheetHost").querySelectorAll("[data-code]").forEach(b=>b.onclick=()=>{setMark(e,b.dataset.code);close()});
  $("tm").focus();
}
function openRemove(x,reason,last){let why=reason||"RESIGNED";
  $("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true" aria-labelledby="rt"><h2 id="rt">Remove '+esc(x.name)+'</h2><div class="sub">'+esc([x.number,x.title,x.sheet].filter(Boolean).join(" · "))+'</div>'+
   '<div class="lbl">Reason</div><div class="codes"><button data-why="RESIGNED">Resigned</button><button data-why="ABSCONDED">Absconded</button><button data-why="DISMISSED">Dismissed</button><button data-why="DECEASED">Deceased</button></div>'+
   '<div class="lbl">Last working day</div><div class="clock"><input type="date" id="lastDay" value="'+(last||S.date)+'"></div>'+
   '<p class="note">From the day after, the person is off the register and their days show as '+'<span id="whyTxt"></span>.</p>'+
   '<div style="display:flex;gap:8px;justify-content:space-between;flex-wrap:wrap"><button class="btn gold" id="doRemove">Remove from register</button><button class="btn" id="closeP">Cancel</button></div></div></div>';
  const sync=()=>{$("sheetHost").querySelectorAll("[data-why]").forEach(b=>b.classList.toggle("cur",b.dataset.why===why));$("whyTxt").textContent=why};sync();
  $("sheetHost").querySelectorAll("[data-why]").forEach(b=>b.onclick=()=>{why=b.dataset.why;sync()});
  const close=()=>{$("sheetHost").innerHTML=""};$("closeP").onclick=close;$("scrim").onclick=ev=>{if(ev.target.id==="scrim")close()};
  $("doRemove").onclick=()=>{const d=$("lastDay").value;if(!d){toast("Choose the last working day");return}
    saveEmps(S.emps.map(y=>y.id===x.id?Object.assign({},y,{active:false,leftReason:why,leftDate:d}):y)).then(()=>toast(x.name+" removed ("+why.toLowerCase()+")"));close()}}
document.addEventListener("keydown",ev=>{if(ev.key==="Escape")$("sheetHost").innerHTML=""});

// ---------- flags ----------
const FLAG_AT=3;const FLAG_CODES=[["AWOP","AWOP"],["FAMILY RESPONSIBILITY","Family responsibility"]];
const CHECK_CODES=[["AWOP","AWOP"],["SICK","Sick"],["FAMILY RESPONSIBILITY","Family responsibility"],["ANNUAL LEAVE","Annual leave"],["STUDY LEAVE","Study leave"]];
const HOT_CODES=["AWOP","SICK","FAMILY RESPONSIBILITY"];
const isFlagged=f=>FLAG_CODES.some(([c])=>f[c].length>=FLAG_AT);
function flagCounts(e,dates){const out={};CHECK_CODES.forEach(([c])=>{out[c]=dates.filter(d=>markOf(e,d)===c)});return out}
function flagBadges(e,dates){const f=flagCounts(e,dates);return FLAG_CODES.filter(([c])=>f[c].length>=FLAG_AT).map(([c,l])=>'<span class="flag" title="'+f[c].length+' '+esc(l)+' days this cycle">'+f[c].length+'× '+(c==="AWOP"?"AWOP":"FAM. RESP.")+'</span>').join("")}
const FLAG_KINDS=[["all","All types"],["AWOP","AWOP"],["SICK","Sick"],["FAMILY RESPONSIBILITY","Family responsibility"],["ANNUAL LEAVE","Annual leave"],["STUDY LEAVE","Study leave"],["3plus","3 or more only"]];
function byKind(rows,kind){if(kind==="all")return rows;if(kind==="3plus")return rows.filter(r=>HOT_CODES.some(c=>r.f[c].length>=FLAG_AT));return rows.filter(r=>r.f[kind].length)}
const kindCols=kind=>CHECK_CODES.filter(([c])=>kind==="all"||kind===c||(kind==="3plus"&&HOT_CODES.includes(c)));
function flagText(f){return FLAG_CODES.filter(([c])=>f[c].length>=FLAG_AT).map(([c])=>f[c].length+'× '+(c==='AWOP'?'AWOP':'FAM. RESP.'))}
function flagged(from,to){const dates=rangeDates(from,to);return inPeriod(from).sort((a,b)=>SHEETS.indexOf(a.sheet)-SHEETS.indexOf(b.sheet)||a.order-b.order)
  .map(e=>({e,f:flagCounts(e,dates)})).filter(r=>CHECK_CODES.some(([c])=>r.f[c].length)).sort((a,b)=>isFlagged(b.f)-isFlagged(a.f))}
function viewFlags(){const F=S.flg;if(!F.from){[F.from,F.to]=F.type==="custom"?[S.date,S.date]:preset(F.type,S.date)}
  const key=F.from+"|"+F.to;if(F.key!==key){F.key=key;F.loading=true;loadRange(F.from,F.to).then(()=>{F.loading=false;if(S.view==="flags")render()},e=>{F.loading=false;toast("Could not load those dates: "+(e.message||"try again"))})}
  let h='<div class="bar" role="group" aria-label="Check period">'+PRESETS.map(([k,l])=>'<button class="chip" data-flg="'+k+'" aria-pressed="'+(F.type===k)+'">'+l+'</button>').join("")+'</div>';
  h+='<div class="form"><label>From<input type="date" id="flg-from" value="'+F.from+'"></label><label>To<input type="date" id="flg-to" value="'+F.to+'"></label>'+
    '<div style="display:flex;align-items:end;gap:8px;flex-wrap:wrap"><button class="btn gold" id="xlsFlags"'+(F.loading?" disabled":"")+'>Download Excel</button><button class="btn gold" id="pdfFlags"'+(F.loading?" disabled":"")+'>Download PDF report</button><button class="btn gold" id="pptFlags"'+(F.loading?" disabled":"")+'>Download PowerPoint report</button></div></div>';
  h+='<div class="bar" role="group" aria-label="Flag type"><span class="lbl" style="margin:0">Flag type</span>'+FLAG_KINDS.map(([k,l])=>'<button class="chip" data-kind="'+esc(k)+'" aria-pressed="'+(F.kind===k)+'">'+l+'</button>').join("")+'</div>';
  h+='<h2 class="sheet-h">AWOP, sick, leave and flags <small>'+periodText(F.from,F.to)+'</small></h2>';
  if(F.loading){main.innerHTML=h+'<div class="empty"><b>Loading these dates</b>Fetching marks for the period.</div>';return}
  const allRows=flagged(F.from,F.to);const rows=allRows;const shown=byKind(allRows,F.kind);const cols=kindCols(F.kind);
  h+='<div class="summary">'+stat("Flagged: "+FLAG_AT+"+ AWOP",rows.filter(r=>r.f.AWOP.length>=FLAG_AT).length)+stat("Flagged: "+FLAG_AT+"+ family resp.",rows.filter(r=>r.f["FAMILY RESPONSIBILITY"].length>=FLAG_AT).length)+
    stat("Employees with AWOP",rows.filter(r=>r.f.AWOP.length).length)+stat("AWOP days",rows.reduce((a,r)=>a+r.f.AWOP.length,0))+stat("Employees sick",rows.filter(r=>r.f.SICK.length).length)+stat("Sick days",rows.reduce((a,r)=>a+r.f.SICK.length,0))+stat("Employees with family resp.",rows.filter(r=>r.f["FAMILY RESPONSIBILITY"].length).length)+stat("Family resp. days",rows.reduce((a,r)=>a+r.f["FAMILY RESPONSIBILITY"].length,0))+stat("Employees on annual leave",rows.filter(r=>r.f["ANNUAL LEAVE"].length).length)+stat("Annual leave days",rows.reduce((a,r)=>a+r.f["ANNUAL LEAVE"].length,0))+stat("Employees on study leave",rows.filter(r=>r.f["STUDY LEAVE"].length).length)+stat("Study leave days",rows.reduce((a,r)=>a+r.f["STUDY LEAVE"].length,0))+'</div>';
  h+='<p class="note">Lists every AWOP, sick, family responsibility, annual leave and study leave day in the chosen period. Annual and study leave are shown for planning and never count towards a flag. Counts of '+FLAG_AT+' or more show in red. An employee is flagged at '+FLAG_AT+' or more AWOP days, or '+FLAG_AT+' or more family responsibility days, and flagged employees are listed first.</p>';
  if(!shown.length){main.innerHTML=h+'<div class="empty"><b>Nothing to report</b>'+(F.kind==="all"?"No AWOP, sick, family responsibility, annual leave or study leave days in this period.":F.kind==="3plus"?"No employee has "+FLAG_AT+" or more AWOP, sick or family responsibility days in this period.":"No "+esc(FLAG_KINDS.find(k=>k[0]===F.kind)[1].toLowerCase())+" days in this period.")+'</div>';return}
  h+='<div class="tbl-wrap"><table><thead><tr><th>Employee</th><th style="text-align:left">Flag</th>'+cols.map(([c,l])=>'<th>'+(c==="FAMILY RESPONSIBILITY"?"Family resp.":c==="ANNUAL LEAVE"?"Annual leave":c==="STUDY LEAVE"?"Study leave":l)+'</th><th style="text-align:left">Dates</th>').join("")+'</tr></thead><tbody>'+
   shown.map(r=>'<tr><td>'+esc(r.e.name)+'<br><span class="note">'+esc([r.e.number,r.e.sheet,groupLabel(r.e)].filter(Boolean).join(" · "))+'</span></td><td style="text-align:left">'+flagText(r.f).map(t=>'<span class="flag">'+esc(t)+'</span>').join(" ")+'</td>'+
     cols.map(([c])=>{const n=r.f[c].length;const hot=n>=FLAG_AT&&HOT_CODES.includes(c);return '<td class="'+(hot?"lowpc":"")+'">'+n+'</td><td style="text-align:left;white-space:normal;min-width:110px">'+r.f[c].map(short).join(", ")+'</td>'}).join("")+'</tr>').join("")+'</tbody></table></div>';
  main.innerHTML=h}
/* flags report as PDF or PowerPoint: same period and flag type as the screen */
async function flagsReport(fmt,btn){const F=S.flg;try{await loadRange(F.from,F.to)}catch(e){toast("Could not load those dates: "+(e.message||"try again"));return}
  const rows=flagged(F.from,F.to),kind=F.kind,shown=byKind(rows,kind),cols=kindCols(kind),kl=FLAG_KINDS.find(k=>k[0]===kind)[1];
  const lbl=c=>c==="FAMILY RESPONSIBILITY"?"Family responsibility":c==="ANNUAL LEAVE"?"Annual leave":c==="STUDY LEAVE"?"Study leave":c==="SICK"?"Sick":c;
  const sum=c=>rows.reduce((a,r)=>a+r.f[c].length,0),ppl=c=>rows.filter(r=>r.f[c].length).length;
  const tiles=[["Flagged: "+FLAG_AT+"+ AWOP",rows.filter(r=>r.f.AWOP.length>=FLAG_AT).length],["Flagged: "+FLAG_AT+"+ family resp.",rows.filter(r=>r.f["FAMILY RESPONSIBILITY"].length>=FLAG_AT).length]]
    .concat(cols.map(([c])=>[(c==="FAMILY RESPONSIBILITY"?"Family resp.":lbl(c))+" days",sum(c)]));
  const fl=shown.filter(r=>flagText(r.f).length);
  const sections=[{title:"Flagged employees ("+fl.length+")",head:["Employee","Number","Sheet","Group","Flag"],rows:fl.length?fl.map(r=>[r.e.name,r.e.number||"",r.e.sheet,groupLabel(r.e),flagText(r.f).join(", ")]):[["No one flagged in this period","","","",""]]}]
    .concat(cols.filter(([c])=>shown.some(r=>r.f[c].length)).map(([c])=>{const L=shown.filter(r=>r.f[c].length).sort((a,b)=>b.f[c].length-a.f[c].length||a.e.name.localeCompare(b.e.name));
      return {title:(()=>{const d=L.reduce((a,r)=>a+r.f[c].length,0);return lbl(c)+" ("+L.length+" employee"+(L.length===1?"":"s")+", "+d+" day"+(d===1?"":"s")+")"})(),head:["Employee","Sheet","Days","Dates"],fs:11,per:12,colW:[3.4,2,0.9,6.2],
        rows:L.length?L.map(r=>{const n=r.f[c].length,hot=n>=FLAG_AT&&HOT_CODES.includes(c);return [r.e.name,r.e.sheet,hot?{content:String(n)}:String(n),r.f[c].map(short).join(", ")]}):[["None","","",""]]}}));
  const sp={title:"ATTENDANCE FLAGS REPORT",sub:periodText(F.from,F.to)+"   ·   "+kl,file:"TKR_Attendance_Flags_"+(kind==="all"?"":kl.replace(/\W+/g,"_")+"_")+(F.from===F.to?F.from:F.from+"_to_"+F.to),tiles,sections};
  if(fmt==="pdf")pdfSpec(sp,btn,"Download PDF report");else pptSpec(sp,btn,"Download PowerPoint report")}
async function xlsFlags(from,to){try{await loadRange(from,to)}catch(e){toast("Could not load those dates: "+(e.message||"try again"));return}
  const kind=S.flg.kind;const rows=byKind(flagged(from,to),kind);const cols=kindCols(kind);const kl=FLAG_KINDS.find(k=>k[0]===kind)[1];
  const aoa=[["TKR Operations – Attendance Flags"],["Period",periodText(from,to)],["Flag type",kl],["Counts of "+FLAG_AT+" or more are marked in the Flag column. Flagged at "+FLAG_AT+" or more AWOP days, or "+FLAG_AT+" or more family responsibility days."],[],
    ["Employee","Number","Sheet","Group","Job Title","Flag"].concat(...cols.map(([c,l])=>[l+" days",l+" dates"]))];
  rows.forEach(r=>aoa.push([r.e.name,r.e.number,r.e.sheet,groupLabel(r.e),r.e.title,flagText(r.f).concat(r.f.SICK.length>=FLAG_AT?[r.f.SICK.length+"\u00d7 SICK"]:[]).join(", ")].concat(...cols.map(([c])=>[r.f[c].length,r.f[c].map(short).join(", ")]))));
  if(!rows.length)aoa.push(["Nothing to report for this flag type in this period."]);
  const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=[{wch:30},{wch:11},{wch:16},{wch:22},{wch:24},{wch:22},{wch:10},{wch:30},{wch:10},{wch:30},{wch:12},{wch:30}];
  XLSX.utils.book_append_sheet(wb,ws,"Flags");saveFile("TKR Attendance Flags "+(kind==="all"?"":kl+" ")+(from===to?from:from+" to "+to)+".xlsx",wb)}

// ---------- excel ----------
// Every download goes through here: the sheets are built as rows, then written as a TKR-branded Excel (logo, black and gold).
function saveFile(name,wb){brandedSave(name,wb).catch(e=>{try{XLSX.writeFile(wb,name);toast("Downloaded without the TKR layout: "+(e.message||"the Excel styler did not load"))}catch(x){toast("Download failed: "+(x.message||"try again"))}})}
const loadScript=src=>new Promise((res,rej)=>{if(window.ExcelJS)return res();const el=document.createElement("script");el.src=src;el.onload=res;el.onerror=()=>rej(new Error("open the app once with signal"));document.head.appendChild(el)});
let logoBuf=null;
const XB={BLACK:"FF111111",HEAD:"FF262626",GROUP:"FF3A3A3A",GOLD:"FFD4A12A",CREAM:"FFFFF8E6",TOT:"FFFFF2CC",WHITE:"FFFFFFFF",GRID:"FFD9D3C5",MUTED:"FF6B6558",
  RED:"FF9C0006",REDF:"FFFDE2DF",BLUE:"FF1F4E79",BLUEF:"FFDDEBF7",GREY:"FF8A8578",NIGHT:"FF6EA7DA"};
const LEAVE_CODES=["ANNUAL LEAVE","FAMILY RESPONSIBILITY","FAMILY EMERGENCY","STUDY LEAVE"],BAD_CODES=["AWOP","SICK","MATERNITY LEAVE"];
async function brandedSave(name,wbIn){toast("Preparing the Excel…");await loadScript("../lib/exceljs.min.js");
  if(!logoBuf){try{logoBuf=await (await fetch("../icons/icon-512.png")).arrayBuffer()}catch(e){logoBuf=null}}
  const X=new ExcelJS.Workbook();X.creator="TKR Operations";X.created=new Date();
  const logo=logoBuf?X.addImage({buffer:logoBuf,extension:"png"}):null;
  const thin={style:"thin",color:{argb:XB.GRID}},box={top:thin,left:thin,bottom:thin,right:thin};
  const fill=c=>({type:"pattern",pattern:"solid",fgColor:{argb:c}});
  for(const sn of wbIn.SheetNames){const src=wbIn.Sheets[sn];const aoa=XLSX.utils.sheet_to_json(src,{header:1,raw:true,defval:""});
    const fmt={};for(const k in src){if(k[0]!=="!"&&src[k].z&&String(src[k].z).includes("%")){const c=XLSX.utils.decode_cell(k);fmt[c.r+"|"+c.c]=src[k].z}}
    const ncols=Math.max(6,...aoa.map(r=>r.length));const widths=(src["!cols"]||[]).map(c=>c&&c.wch||10);
    const ws=X.addWorksheet(sn.slice(0,31),{views:[{showGridLines:false}],pageSetup:{orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0,paperSize:9}});
    ws.columns=Array.from({length:ncols},(_,i)=>({width:Math.max(widths[i]||10,i===0?26:6)}));
    // brand band
    for(let r=1;r<=4;r++){ws.getRow(r).height=r===1?8:20;for(let c=1;c<=ncols;c++)ws.getCell(r,c).fill=fill(XB.BLACK)}
    ws.getRow(5).height=4;for(let c=1;c<=ncols;c++)ws.getCell(5,c).fill=fill(XB.GOLD);
    if(logo!==null)ws.addImage(logo,{tl:{col:0.08,row:1.08},ext:{width:56,height:56}});
    const title=String((aoa[0]||[])[0]||sn);
    const tc=ws.getCell(2,2);tc.value=title;tc.font={name:"Arial",size:15,bold:true,color:{argb:XB.WHITE}};tc.alignment={vertical:"middle"};
    const sc=ws.getCell(3,2);sc.value="TKR OPERATIONS  ·  THARISA MINE";sc.font={name:"Arial",size:9,bold:true,color:{argb:XB.GOLD}};sc.alignment={vertical:"middle"};
    const dc=ws.getCell(4,2);dc.value="Downloaded "+nice(iso(new Date()))+" "+yr(iso(new Date()))+(email?"  ·  "+email:"");dc.font={name:"Arial",size:9,color:{argb:"FFCFC8B8"}};dc.alignment={vertical:"middle"};
    // find the header row: first row after the title with 3 or more filled cells, all text
    let hi=aoa.findIndex((r,i)=>i>0&&r.filter(v=>v!==""&&v!=null).length>=3&&r.every(v=>v===""||typeof v==="string"));if(hi<0)hi=aoa.length;
    let out=7;
    for(let i=1;i<hi;i++){const r=aoa[i];if(!r||!r.some(v=>v!==""))continue;const c=ws.getCell(out,1);
      if(r.length>=2&&r[1]!==""){c.value=String(r[0]);c.font={name:"Arial",size:10,bold:true,color:{argb:XB.MUTED}};const v=ws.getCell(out,2);v.value=r.slice(1).join(" ");v.font={name:"Arial",size:10,color:{argb:XB.BLACK}}}
      else{c.value=String(r[0]);c.font={name:"Arial",size:10,italic:true,color:{argb:XB.MUTED}}}out++}
    if(hi>=aoa.length){freezeAt(ws,5);continue}
    out++;const headRowNo=out;const head=aoa[hi];const hr=ws.getRow(out);hr.height=30;
    for(let c=0;c<ncols;c++){const cell=hr.getCell(c+1);cell.value=head[c]??"";cell.font={name:"Arial",size:10,bold:true,color:{argb:XB.WHITE}};cell.fill=fill(XB.HEAD);cell.border=box;cell.alignment={vertical:"middle",horizontal:c?"center":"left",wrapText:true}}
    out++;let k=0,ended=false;
    for(let i=hi+1;i<aoa.length;i++){const r=aoa[i]||[];const filled=r.filter(v=>v!==""&&v!=null).length;
      if(!filled){ended=true;out++;continue}
      const row=ws.getRow(out);
      if(ended||(filled===1&&String(r[0]).length>60)){const c=row.getCell(1);c.value=String(r[0]);c.font={name:"Arial",size:9,italic:true,color:{argb:XB.MUTED}};out++;continue}
      const isTot=String(r[0]).trim().toLowerCase()==="total";
      const isGroup=!isTot&&r[0]!==""&&r.slice(1,4).every(v=>v===""||v==null)&&r.slice(4).every(v=>v===""||["DAY","NIGHT","OFF"].includes(v));
      row.height=isGroup?20:18;
      for(let c=0;c<ncols;c++){const v=r[c]===undefined?"":r[c];const cell=row.getCell(c+1);cell.value=v===""?null:v;cell.border=box;
        cell.font={name:"Arial",size:10,bold:isTot||isGroup||c===0&&false,color:{argb:isGroup?XB.WHITE:XB.BLACK}};
        cell.alignment={vertical:"middle",horizontal:typeof v==="number"||(c>0&&typeof v==="string"&&/^[A-Z0-9 :.%×\u00d7-]+$/.test(v)&&v.length<=22&&!/^[A-Z]-/.test(v))?"center":"left",wrapText:c<=5,shrinkToFit:c>5&&typeof v==="string"&&v.length>6};
        if(isGroup){cell.fill=fill(XB.GROUP);if(v==="DAY")cell.font={name:"Arial",size:9,bold:true,color:{argb:XB.GOLD}};else if(v==="NIGHT")cell.font={name:"Arial",size:9,bold:true,color:{argb:XB.NIGHT}};else if(v==="OFF")cell.font={name:"Arial",size:9,color:{argb:"FFBDB6A8"}}}
        else if(isTot){cell.fill=fill(XB.TOT);cell.border={top:{style:"medium",color:{argb:XB.GOLD}},left:thin,bottom:thin,right:thin}}
        else{if(k%2)cell.fill=fill(XB.CREAM);
          if(typeof v==="string"&&c>0){const u=v.toUpperCase();
            if(BAD_CODES.includes(u)){cell.fill=fill(XB.REDF);cell.font={name:"Arial",size:10,bold:true,color:{argb:XB.RED}}}
            else if(LEAVE_CODES.includes(u)){cell.fill=fill(XB.BLUEF);cell.font={name:"Arial",size:10,color:{argb:XB.BLUE}}}
            else if(u==="OFF")cell.font={name:"Arial",size:10,color:{argb:XB.GREY}};
            else if(["RESIGNED","ABSCONDED","DISMISSED","DECEASED","SUSPENDED"].includes(u))cell.font={name:"Arial",size:10,italic:true,color:{argb:XB.MUTED}}}}
        const z=fmt[i+"|"+c];if(z&&typeof v==="number"){cell.numFmt=z;if(v<0.9&&!isGroup)cell.font=Object.assign({},cell.font,{color:{argb:XB.RED},bold:true})}}
      if(!isGroup&&!isTot)k++;out++}
    freezeAt(ws,headRowNo)}
  const buf=await X.xlsx.writeBuffer();const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const {deliver}=await import("../lib/deliver.js");
  deliver(blob,name,{doneLabel:"Done – clear this phone's copy",onDone:async()=>{const fb=await import("../lib/firebase.js");const {clearPhoneCopy}=await import("../lib/deliver.js");await clearPhoneCopy(fb,[fs],toast)}})}
function freezeAt(ws,row){ws.views=[{state:"frozen",ySplit:row,xSplit:row>5?1:0,showGridLines:false}]}
const yr=d=>parse(d).getFullYear();
function rangeDates(from,to){const out=[];for(let x=from;x<=to&&out.length<400;x=addDays(x,1))out.push(x);return out}
const PRESETS=[["daily","Daily"],["weekly","Weekly"],["cycle","Cycle"],["month","Calendar month"],["custom","Custom dates"]];
function preset(type,anchor){const d=parse(anchor);
  if(type==="daily")return [anchor,anchor];
  if(type==="weekly"){const back=(d.getDay()+6)%7;const mon=addDays(anchor,-back);return [mon,addDays(mon,6)]}
  if(type==="cycle"){const c=cycleOf(anchor);return [c.start,c.end]}
  if(type==="month")return [iso(new Date(d.getFullYear(),d.getMonth(),1)),iso(new Date(d.getFullYear(),d.getMonth()+1,0))];
  return [S.exp.from||anchor,S.exp.to||anchor]}
const HPD=()=>{try{const v=Number(localStorage.getItem("attHpd"));return v>0&&v<=24?v:12}catch(e){return 12}};
const EXC=["AWOP","SICK","ANNUAL LEAVE","FAMILY RESPONSIBILITY","FAMILY EMERGENCY","STUDY LEAVE","MATERNITY LEAVE"];
function periodText(f,t){const n=rangeDates(f,t).length;return (f===t?nice(f)+" "+yr(f):nice(f)+" "+yr(f)+" to "+nice(t)+" "+yr(t))+" · "+n+(n===1?" day":" days")}
function xlsEmployees(){const act=activeEmps().slice().sort((a,b)=>SHEETS.indexOf(a.sheet)-SHEETS.indexOf(b.sheet)||a.order-b.order),gone=(S.emps||[]).filter(e=>e.active===false);
  if(!act.length&&!gone.length){toast("There are no employees to download");return}
  const head=["Sheet","Shift group","Employee","Number","Department","Job Title"],row=e=>[e.sheet,groupLabel(e),e.name,e.number||"",e.dept||"",e.title||""],cols=[{wch:16},{wch:24},{wch:32},{wch:12},{wch:18},{wch:30}];
  const wb=XLSX.utils.book_new();const add=(name,aoa,c)=>{const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=c;XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31))};
  const title="TKR Operations – Employee list – "+nice(iso(new Date()))+" "+yr(iso(new Date()));
  add("All employees",[[title],["Employees on the register: "+act.length],[],head].concat(act.map(row)),cols);
  SHEETS.forEach(sh=>{const es=act.filter(e=>e.sheet===sh);if(es.length)add(sh,[[sh+" – "+es.length+" employees"],[],head].concat(es.map(row)),cols)});
  saveFile("TKR Employee list "+iso(new Date())+".xlsx",wb)}
const fmtD=s=>{const [y,m,d]=s.split("-");return d+"/"+m+"/"+y};
function viewExport(){const E=S.exp;if(!E.from){[E.from,E.to]=preset(E.type,S.date)}
  const key=E.from+"|"+E.to;if(E.key!==key){E.key=key;E.loading=true;loadRange(E.from,E.to).then(()=>{E.loading=false;if(S.view==="export")render()},e=>{E.loading=false;toast("Could not load those dates: "+(e.message||"try again"))})}
  const dates=rangeDates(E.from,E.to);const es=inPeriod(E.from);
  let h='<div class="list" style="margin-bottom:14px"><h3>Employee list</h3><p class="note" style="margin:0 0 8px">Everyone on the register with their sheet, shift group, number, department and job title, one tab per sheet.</p><button class="btn gold" id="xlsEmps">Download employee list</button></div>';
  h+='<div class="bar" role="group" aria-label="Report period">'+PRESETS.map(([k,l])=>'<button class="chip" data-exp="'+k+'" aria-pressed="'+(E.type===k)+'">'+l+'</button>').join("")+'</div>';
  h+='<div class="form"><label>From<input type="date" id="exp-from" value="'+E.from+'"></label><label>To<input type="date" id="exp-to" value="'+E.to+'"></label>'+
    '<div style="display:flex;align-items:end"><button class="btn gold" id="xlsRange"'+(E.loading?" disabled":"")+'>Download Excel</button></div></div>';
  h+='<h2 class="sheet-h">Attendance report <small>'+periodText(E.from,E.to)+'</small></h2>';
  if(E.loading){main.innerHTML=h+'<div class="empty"><b>Loading these dates</b>Fetching marks for the period.</div>';return}
  const all=tally(es,dates);
  h+='<div class="summary">'+stat("Employees",es.length)+stat("Present days",all.present)+stat("AWOP days",all.absent)+stat("Leave & sick days",all.leave)+stat("Off days",all.off)+stat("Not marked",all.blank)+stat("Overtime",hm(all.ot))+stat("Attendance",fmtPct(pct(all.present,all.exc)),"pct")+'</div>';
  h+='<div class="tbl-wrap"><table><thead><tr><th>Sheet</th><th>Employees</th><th>Present</th><th>AWOP</th><th>Leave</th><th>Off</th><th>Resigned, absconded, dismissed, deceased, suspended</th><th>Not marked</th><th>Overtime</th><th>Attendance</th></tr></thead><tbody>';
  SHEETS.forEach(sh=>{const l=es.filter(e=>e.sheet===sh);if(l.length)h+=trow(sh,l.length,tally(l,dates))});
  h+=trow("Total",es.length,all,"tot")+'</tbody></table></div>';
  h+='<p class="note">The Excel file has a Summary tab and one tab per sheet, with each person\'s mark for every day in the period and their totals.</p>';
  main.innerHTML=h}
async function xlsRange(from,to,shift){try{await loadRange(from,to)}catch(e){toast("Could not load those dates: "+(e.message||"try again"));return}
  shift=shift||"all";const dates=rangeDates(from,to);const wb=XLSX.utils.book_new();const es=onShift(inPeriod(from),from,from===to?shift:"all");if(from!==to)shift="all";
  const sumHead=["Sheet","Employees","Present"].concat(EXC,["Off","Resigned, absconded, dismissed, deceased, suspended","Not marked","Overtime hours","Attendance %"]);
  const line=(name,n,t)=>{const p=pct(t.present,t.exc);return [name,n,t.present].concat(EXC.map(k=>t.codes[k]||0),[t.off,t.out,t.blank,hm(t.ot),p==null?"":p/100])};
  const sum=[["TKR Operations – Attendance Report"+(shift!=="all"?" – "+SHIFT_LBL[shift]:"")],["Period",periodText(from,to)],[],sumHead];
  SHEETS.forEach(sh=>{const l=es.filter(e=>e.sheet===sh);if(l.length)sum.push(line(sh,l.length,tally(l,dates)))});
  sum.push(line("Total",es.length,tally(es,dates)));
  sum.push([],["Attendance % = present days ÷ (present days + exception days). Exceptions are AWOP, sick and maternity leave. Annual leave, family responsibility, family emergency and study leave are not counted as absent. Medical counts as present. Off days are not counted."]);
  const ws0=XLSX.utils.aoa_to_sheet(sum);ws0["!cols"]=[{wch:20},{wch:10},{wch:9}].concat(EXC.map(()=>({wch:12})),[{wch:8},{wch:12},{wch:11},{wch:14},{wch:13}]);
  for(let r=4;r<sum.length;r++){const ref=XLSX.utils.encode_cell({r,c:sumHead.length-1});if(ws0[ref]&&typeof ws0[ref].v==="number")ws0[ref].z="0.0%"}
  XLSX.utils.book_append_sheet(wb,ws0,"Summary");
  bySheetGroup(es).forEach(s=>{const head=["Employee","Number","Department","Job Title"].concat(dates.map(short),["Present"],EXC,["Off","Not marked","Overtime hours","Attendance %"]);const aoa=[[s.sheet+" – "+periodText(from,to)],head];
    s.groups.forEach(g=>{aoa.push([g.label,"","",""].concat(dates.map(d=>rotaFor(g.id,d))));
      g.emps.forEach(e=>{const t=tally([e],dates);const p=pct(t.present,t.exc);aoa.push([e.name,e.number,e.dept,e.title].concat(dates.map(d=>showMark(markOf(e,d))),[t.present],EXC.map(k=>t.codes[k]||0),[t.off,t.blank,hm(t.ot),p==null?"":p/100]))})});
    const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=[{wch:30},{wch:11},{wch:16},{wch:24}].concat(dates.map(()=>({wch:7})));
    const pc=head.length-1;for(let r=2;r<aoa.length;r++){const ref=XLSX.utils.encode_cell({r,c:pc});if(ws[ref]&&typeof ws[ref].v==="number")ws[ref].z="0.0%"}
    XLSX.utils.book_append_sheet(wb,ws,s.sheet.slice(0,31))});
  const otRows=[];bySheetGroup(es).forEach(s=>s.groups.forEach(g=>g.emps.forEach(e=>dates.forEach(d=>{const h=otOf(e,d);if(h)otRows.push([e.name,e.number,s.sheet,g.label,nice(d)+" "+yr(d),cell(e,d,"ottype")||"",h])}))));
  if(otRows.length){const tot=otRows.reduce((a,r)=>a+r[6],0);const ws=XLSX.utils.aoa_to_sheet([["Overtime – "+periodText(from,to)],["Employee","Number","Sheet","Group","Date","Planned / unplanned","Hours"]].concat(otRows,[["Total","","","","","",Math.round(tot*10)/10]]));
    ws["!cols"]=[{wch:30},{wch:11},{wch:16},{wch:22},{wch:16},{wch:18},{wch:8}];XLSX.utils.book_append_sheet(wb,ws,"Overtime")}
  saveFile("TKR Attendance "+(from===to?from:from+" to "+to)+(shift!=="all"?" "+SHIFT_LBL[shift]:"")+".xlsx",wb)}

// ---------- users (admins) ----------
function viewUsers(){const row=(em,role)=>'<tr><td>'+esc(em)+'</td><td>'+(role==="capturers"?"Marks attendance":"Views reports")+'</td><td><button class="btn sm" data-rmuser="'+esc(em)+'" data-role="'+role+'">Remove</button></td></tr>';
  const n=roles.capturers.length+roles.viewers.length;
  main.innerHTML='<form class="form" id="addUser"><label>Email<input id="uEmail" type="email" required></label><label>Password<input id="uPass" type="text" minlength="6" placeholder="Leave empty if they already have one"></label>'+
   '<label>Access<select id="uRole"><option value="capturers">Mark attendance</option><option value="viewers">View reports only</option></select></label><div style="display:flex;align-items:end"><button class="btn gold" type="submit">Add user</button></div></form>'+
   '<p class="note">Admins ('+ADMINS.map(esc).join(", ")+') always have full access. People who mark attendance can also add and remove employees and set cycles.</p><p class="note" id="uMsg" role="status"></p>'+
   '<h2 class="sheet-h">Who has access <small>'+n+'</small></h2>'+(n?'<div class="tbl-wrap"><table><thead><tr><th>Email</th><th style="text-align:left">Access</th><th></th></tr></thead><tbody>'+roles.capturers.map(e=>row(e,"capturers")).join("")+roles.viewers.map(e=>row(e,"viewers")).join("")+'</tbody></table></div>':'<p class="note">Only the admins have access so far.</p>')+
   '<h2 class="sheet-h">Data file</h2><p class="note">Load a TKR Attendance data file to put its employees, shift rotas, cycles and marks into the register. Marks in the file replace the same days in the register.</p>'+
   '<div class="bar"><button class="btn" id="loadData">Load data file</button><button class="btn" id="saveData">Download data file</button><input type="file" id="dataIn" accept=".json,application/json" hidden></div>';
  $("addUser").onsubmit=async ev=>{ev.preventDefault();const em=$("uEmail").value.trim().toLowerCase(),pw=$("uPass").value,role=$("uRole").value,other=role==="capturers"?"viewers":"capturers";const msg=$("uMsg");
    if(pw&&pw.length<6){msg.textContent="The password must be at least 6 characters.";return}msg.textContent="Adding…";let note="";
    if(pw){const tmp=initializeApp(firebaseConfig,"add-"+Date.now());
      try{await createUserWithEmailAndPassword(getAuth(tmp),em,pw);note=" Their sign-in is made."}catch(err){if(err&&err.code==="auth/email-already-in-use")note=" They already had a sign-in, so their old password stays.";else{msg.textContent="The sign-in couldn't be made. Check the email and try again.";deleteApp(tmp).catch(()=>{});return}}
      deleteApp(tmp).catch(()=>{})}
    try{await setDoc(doc(fs,"config","attendance"),{[role]:arrayUnion(em),[other]:arrayRemove(em)},{merge:true});toast(em+" can now sign in."+note)}catch(err){msg.textContent="Not saved: "+(err.message||"try again")}};
  $("loadData").onclick=()=>$("dataIn").click();
  $("dataIn").onchange=ev=>{const f=ev.target.files[0];if(!f)return;const rd=new FileReader();rd.onload=()=>{try{loadBackup(JSON.parse(rd.result))}catch(e){toast("That file is not a TKR Attendance data file")}};rd.readAsText(f);ev.target.value=""};
  $("saveData").onclick=saveBackup;
}
async function loadBackup(b){if(!b||!b.employees||!b.months){toast("That file is not a TKR Attendance data file");return}
  const msg=$("uMsg")||{set textContent(v){toast(v)}};msg.textContent="Loading the data file…";
  try{await S.db.doc("config/employees").set({list:b.employees,groups:b.groups||[],history:(S.hist||[]).concat(empEvents(S.emps||[],b.employees))});
    if(b.settings)await S.db.doc("config/settings").set(b.settings);
    for(const id in (b.rota||{}))await S.db.doc("cycles/"+id).set(b.rota[id]);
    for(const m in b.months)for(const sh in b.months[m])await S.db.doc("months/"+m+"/sheets/"+sh).update(b.months[m][sh]);
    msg.textContent="Data file loaded: "+b.employees.length+" employees, "+Object.keys(b.months).length+" months of marks.";toast("Data file loaded")}
  catch(e){msg.textContent=denied(e)?"Your sign-in is not allowed to use the register yet. In the Firebase console open tkr-operations-app, then Firestore Database, then Rules, add the TKR Attendance rules and tap Publish. Then close and open the app again.":"Loading stopped: "+(e.message||"try again")}}
async function saveBackup(){const b={employees:S.emps||[],groups:S.groupDefs||[],settings:{cycles:CUSTOM},rota:S.cycles,months:{}};
  const ms=new Set();Object.keys(S.rota).forEach(g=>Object.keys(S.rota[g]).forEach(d=>ms.add(monthOf(d))));Object.keys(S.mm).forEach(m=>ms.add(m));
  for(const m of ms){if(!S.mm[m]){const snap=await S.db.collection("months/"+m+"/sheets").get();applyMonth(m,snap)}b.months[m]={};for(const sh in S.mm[m]||{})b.months[m][sh]={sheet:sh,month:m,marks:S.mm[m][sh].marks,ot:S.mm[m][sh].ot}}
  const blob=new Blob([JSON.stringify(b)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="TKR Attendance data "+iso(new Date())+".json";document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000)}
main.addEventListener("click",ev=>{const t=ev.target.closest("[data-rmuser]");if(!t)return;
  if(t.dataset.armed){setDoc(doc(fs,"config","attendance"),{[t.dataset.role]:arrayRemove(t.dataset.rmuser)},{merge:true}).then(()=>toast("Access removed"),e=>toast("Not saved: "+(e.message||"try again")))}
  else{t.dataset.armed="1";t.textContent="Tap again to remove";t.classList.add("gold")}});

// ---------- events ----------
$("tabs").onclick=ev=>{const b=ev.target.closest("button");if(!b)return;S.view=b.dataset.v;try{localStorage.setItem("tkr-att-view",S.view)}catch(e){}render()};
$("date").onchange=ev=>{if(ev.target.value){S.date=ev.target.value;if(S.db)watchCycle();render()}};
$("cycleLbl").onclick=()=>{if(S.db)openCyclePanel()};
$("prev").onclick=()=>{S.date=addDays(S.date,-1);if(S.db)watchCycle();render()};
$("next").onclick=()=>{S.date=addDays(S.date,1);if(S.db)watchCycle();render()};
main.addEventListener("input",ev=>{if(ev.target.id==="hq"){S.hq=ev.target.value;const pos=ev.target.selectionStart;render();const q=$("hq");if(q){q.focus();q.setSelectionRange(pos,pos)}return}if(ev.target.id==="q"){S.q=ev.target.value;const pos=ev.target.selectionStart;render();const q=$("q");if(q){q.focus();q.setSelectionRange(pos,pos)}}});
main.addEventListener("submit",ev=>{if(ev.target.id!=="empForm")return;ev.preventDefault();
  const f={sheet:$("f-sheet").value,group:$("f-group").value,name:$("f-name").value.trim(),number:$("f-number").value.trim(),dept:$("f-dept").value.trim(),title:$("f-title").value.trim()};
  if(!f.name){toast("Enter the employee's name");return}
  const g=allGroups().find(x=>x.id===f.group);if(!g||g.sheet!==f.sheet){f.section=f.sheet;f.shift="";f.group=f.sheet+"|"+f.sheet+"|"}else{f.section=g.section;f.shift=g.shift}
  let list=(S.emps||[]).slice();
  if(editing){const moved=editing.group!==f.group;if(moved)f.order=orderFor(f.group,f.sheet,list,editing.id);list=list.map(x=>x.id===editing.id?Object.assign({},x,f):x)}
  else{const id=(f.number||f.name).replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"")+"-"+Date.now().toString(36);
    list.push(Object.assign({id,order:orderFor(f.group,f.sheet,list),active:true},f))}
  const wasEdit=!!editing;editing=null;saveEmps(list).then(()=>toast(wasEdit?"Changes saved":f.name+" added"))});
main.addEventListener("change",ev=>{if((ev.target.id==="flg-from"||ev.target.id==="flg-to")&&ev.target.value){const F=S.flg;F.type="custom";F[ev.target.id==="flg-from"?"from":"to"]=ev.target.value;if(F.to<F.from)[F.from,F.to]=[F.to,F.from];if(rangeDates(F.from,F.to).length>=400){F.to=addDays(F.from,365);toast("A check covers at most one year")}render();return}
  if((ev.target.id==="exp-from"||ev.target.id==="exp-to")&&ev.target.value){const E=S.exp;E.type="custom";E[ev.target.id==="exp-from"?"from":"to"]=ev.target.value;if(E.to<E.from)[E.from,E.to]=[E.to,E.from];if(rangeDates(E.from,E.to).length>=400){E.to=addDays(E.from,365);toast("A report covers at most one year")}render();return}
  if(ev.target.id==="f-sheet"){const sh=ev.target.value;const sel=$("f-group");const opt=[...sel.options].find(o=>o.dataset.sheet===sh);if(opt)sel.value=opt.value}});
main.addEventListener("click",ev=>{const t=ev.target.closest("button");if(!t)return;
  if(t.dataset.sheet){S.sheet=t.dataset.sheet;try{localStorage.setItem("tkr-att-sheet",S.sheet)}catch(e){}render();return}
  if(t.dataset.rota!==undefined){openRotaDay(t.dataset.rota);return}
  if(t.dataset.emp){const e=activeEmps().find(x=>x.id===t.dataset.emp);if(e)openPanel(e);return}
  if(t.dataset.fill){if(!S.canWrite){toast("You can view this register but not mark it.");return}
    const es=filtered().filter(e=>e.group===t.dataset.fill&&!markOf(e,S.date));if(!es.length){toast("Everyone in this group is already marked");return}
    const by={};es.forEach(e=>{(by[e.sheet]=by[e.sheet]||{})[e.id]={[S.date]:t.dataset.val}});for(const sh in by)writeSheet(sh,by[sh]);toast(es.length+" marked "+t.dataset.val);return}
  if(t.dataset.edit){editing=(S.emps||[]).find(x=>x.id===t.dataset.edit);render();window.scrollTo({top:0});return}
  if(t.dataset.keepdup){t.disabled=true;keepDup(t.dataset.keepdup,t.dataset.dups.split(",")).finally(()=>{t.disabled=false});return}
  if(t.dataset.remove){const x=(S.emps||[]).find(y=>y.id===t.dataset.remove);if(x)openRemove(x,t.dataset.reason,t.dataset.last);return}
  if(t.dataset.restore){const x=(S.emps||[]).find(y=>y.id===t.dataset.restore);const {leftDate,leftReason,...rest}=x;saveEmps(S.emps.map(y=>y.id===x.id?Object.assign(rest,{active:true}):y)).then(()=>toast(x.name+" is back on the register"));return}
  if(t.id==="cancelEdit"){editing=null;render();return}
  if(t.id==="startCycle"){startCycle();return}
  if(t.id==="sortLow"){S.sortLow=!S.sortLow;render();return}
  if(t.id==="xlsCycle"){const c=cycleOf(S.date);xlsRange(c.start,c.end);return}
  if(t.dataset.dshift){S.dshift=t.dataset.dshift;render();return}
  if(t.id==="xlsDaily"){xlsRange(S.date,S.date,S.dshift);return}
  if(t.dataset.exp){S.exp.type=t.dataset.exp;if(t.dataset.exp!=="custom")[S.exp.from,S.exp.to]=preset(t.dataset.exp,S.exp.from||S.date);render();return}
  if(t.dataset.htype){(S.hf||(S.hf={type:"All",from:"",to:""})).type=t.dataset.htype;render();return}
  if(t.id==="pdfCycle"){pdfCycle();return}
  if(t.id==="pdfDaily"){pdfSpec(dailySpec(),t,"Download PDF report");return}
  if(t.id==="pptDaily"){pptSpec(dailySpec(),t,"Download PowerPoint report");return}
  if(t.id==="pptCycle"){pptCycle();return}
  if(t.id==="xlsEmps"){xlsEmployees();return}
  if(t.id==="xlsRange"){xlsRange(S.exp.from,S.exp.to);return}
  if(t.dataset.flg){const F=S.flg;F.type=t.dataset.flg;if(F.type!=="custom")[F.from,F.to]=preset(F.type,F.from||S.date);render();return}
  if(t.dataset.kind){S.flg.kind=t.dataset.kind;render();return}
  if(t.id==="xlsFlags"){xlsFlags(S.flg.from,S.flg.to);return}
  if(t.id==="pdfFlags"){flagsReport("pdf",t);return}
  if(t.id==="pptFlags"){flagsReport("ppt",t);return}
  if(t.id==="copyDaily"){const txt=dailyText();navigator.clipboard.writeText(txt).then(()=>toast("Report copied. Paste it into WhatsApp."),()=>{$("sheetHost").innerHTML='<div class="scrim" id="scrim"><div class="panel"><h2>Daily report</h2><div class="sub">Select all and copy.</div><textarea id="txt" style="width:100%;height:300px;background:var(--surface-2);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:8px;font:13px var(--mono)">'+esc(txt)+'</textarea><button class="btn" id="closeP">Close</button></div></div>';$("closeP").onclick=()=>$("sheetHost").innerHTML="";$("txt").select()});return}
});

// ---------- boot ----------
render();
startApp=function(db){if(S.db)return;S.db=db;S.canWrite=canWrite();
  db.doc("config/employees").onSnapshot(d=>{S.emps=d.exists?(d.data().list||[]):[];S.groupDefs=d.exists?(d.data().groups||[]):[];S.hist=d.exists?(d.data().history||[]):[];render()},err);
  db.collection("cycles").onSnapshot(snap=>{S.cycles={};snap.docs.forEach(d=>{S.cycles[d.id]=d.data()});applyRotas();render()},err);
  db.doc("config/settings").onSnapshot(d=>{setCustom(d.exists?d.data().cycles:[]);watchCycle();render()},err);
  watchCycle();render()};
rerender=()=>{S.canWrite=canWrite();render()};
})();

// ---------- sign-in ----------
const q=id=>document.getElementById(id);
function showSignin(on){q("signin").hidden=!on;q("main").hidden=on;q("appBar").hidden=on;q("tabs").hidden=on;q("outBtn").hidden=on}
let rolesUnsub=null;
onAuthStateChanged(auth,u=>{if(rolesUnsub){rolesUnsub();rolesUnsub=null}
  if(!u||u.isAnonymous||!u.email||u.email.endsWith("@staff.tkroperations.app")){email="";showSignin(true);return}
  email=(u.email||"").toLowerCase();showSignin(false);q("usersTab").hidden=!isAdmin();
  q("main").innerHTML='<div class="empty"><b>Opening the register</b>Checking your access.</div>';
  getDoc(doc(fs,"config","roles")).catch(()=>null).then(r=>{mgrs=r&&r.exists()?(r.data().managers||[]).map(x=>String(x).toLowerCase()):[];
  if(email!==(u.email||"").toLowerCase())return;
  rolesUnsub=onSnapshot(doc(fs,"config","attendance"),d=>{const v=d.exists()?d.data():{};roles={capturers:(v.capturers||[]).map(x=>x.toLowerCase()),viewers:(v.viewers||[]).map(x=>x.toLowerCase())};
    if(!canView()){q("appBar").hidden=true;q("tabs").hidden=true;const nm=/@mgr\.tkroperations\.app$/.test(email)?email[0].toUpperCase()+email.slice(1,email.indexOf("@")):email.replace(/[<>&]/g,"");q("main").innerHTML='<div class="empty"><b>Sorry '+nm+'</b>You don\'t have access to Attendance with this account. Please ask admin.<div style="display:flex;flex-direction:column;gap:10px;max-width:340px;margin:18px auto 0"><button id="naSwitch" type="button" style="min-height:48px;border-radius:10px;border:1px solid #d4a12a;background:#d4a12a;color:#141006;font-weight:700;font-size:16px">Sign in with a different account</button><button id="naHome" type="button" style="min-height:48px;border-radius:10px;border:1px solid #888;background:transparent;color:inherit;font-weight:600;font-size:16px">Back to the home page</button></div></div>';q("naSwitch").onclick=()=>signOut(auth).then(()=>location.reload());q("naHome").onclick=()=>{location.href="../"};return}
    q("appBar").hidden=false;q("tabs").hidden=false;startApp(fdb);rerender()},
    e=>{q("main").innerHTML='<div class="empty"><b>The register could not load</b>'+String(e&&e.message||"Check your signal and open the app again.").replace(/[<>&]/g,"")+'</div>'})})});
q("signForm").onsubmit=async e=>{e.preventDefault();const m=q("sMsg");m.textContent="Signing in…";q("sBtn").disabled=true;
  try{await signInWithEmailAndPassword(auth,toLogin(q("sEmail").value),q("sPass").value);q("sPass").value="";m.textContent=""}
  catch(err){m.textContent=err&&err.code==="auth/network-request-failed"?"No signal. Sign in once with signal; after that the app works offline.":"The email or password is wrong."}
  q("sBtn").disabled=false};
q("forgot").onclick=async()=>{const em=q("sEmail").value.trim();const m=q("sMsg");if(!em.includes("@")){m.textContent="Type your email first, then tap Forgot password.";return}
  try{await sendPasswordResetEmail(auth,em);m.textContent="A password reset email was sent to "+em+"."}catch(e){m.textContent="The reset email couldn't be sent. Ask admin."}};
q("outBtn").onclick=()=>signOut(auth).then(()=>location.reload());
let installEvt=null;
// A new version of the app was installed: load it straight away instead of on the next open.
if("serviceWorker" in navigator&&navigator.serviceWorker.controller){let reloaded=false;navigator.serviceWorker.addEventListener("controllerchange",()=>{if(reloaded)return;reloaded=true;location.reload()})}
if("serviceWorker" in navigator)navigator.serviceWorker.register("../sw.js",{scope:"../",updateViaCache:"none"}).then(r=>r.update()).catch(()=>{});
