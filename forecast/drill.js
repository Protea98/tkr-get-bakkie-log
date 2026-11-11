/* TKR drilling: metres per shift, documents by date (tally sheets, job cards, checklists, breakdown reports, supporting documents),
   breakdown report and the metres Excel. Documents are filed in SharePoint when it is connected; a light copy stays in the app. */
import { spConfigured, spInit, spReady as spSigned, spWho, spSignIn, spSignOut, spUpload, spDelete, spFolderUrl, spRoot, spPlace, spLinked, spLinkOneDrive } from "../lib/sharepoint.js";
const spReady = () => spSigned() && spLinked();

export const DOC_TYPES = ["Tally sheet", "Job card", "Checklist", "Breakdown report", "Supporting document"];
const FOLDER = { "Tally sheet": "Tally sheets", "Job card": "Job cards", "Checklist": "Checklists", "Breakdown report": "Breakdown reports", "Supporting document": "Supporting documents" };
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const CAT_RX = [
  ["Waiting for GET / drill bits", /get|\bbits?|drill bit/],
  ["Stuck rods / retrieving rods", /stuck|retriev|rods? (lost|in hole)|lost .*hole|recover/],
  ["Service and maintenance", /service|maint|250\s*hr|weekly|daily|drifter/],
  ["Ground conditions", /uneven|loose ground|surface/],
  ["Breakdowns", /breakdown|nsrt|catcher|gripper|burst|broke|broken|leak|cylinder|engine|chain|panel|compressor|dust|pipe|hose|shank|starter|stater|rod handler|suction|gauge|power|radio|loose|hydr|oil|fitting|jaw|percation|percussion|cut off|selonoid|solenoid|boilermaker/],
  ["Tramming", /tram|traming/],
  ["Blasting", /blast/],
  ["Diesel / refuelling", /diesel|refuel|refill/],
  ["Waiting for block access / bench", /access|bench|pattern|barricad/],
  ["Meetings, transport, safety and weather", /meeting|transport|safety|sheq|lightning|incident/],
  ["No operator", /operator/]];
const UNCLEAR = "Cause not clear on the shift report", NOCAUSE = "Time written, cause not written", WEEKEND = "No drilling – weekend";
export const DELAY_CATS = ["Breakdowns", "Service and maintenance", "Waiting for GET / drill bits", "Stuck rods / retrieving rods", "Tramming", "Diesel / refuelling", "Blasting",
  "Waiting for block access / bench", "No operator", "Meetings, transport, safety and weather", "Ground conditions", "Other"];
function catOf(t) {
  const tl = t.toLowerCase();
  if (/^[\d.,\s]*(hrs?|min|h)?[\d.,\s]*(min)?$/.test(tl.trim())) return NOCAUSE;
  if (/saturday|sunday/.test(tl)) return WEEKEND;
  for (const [n, rx] of CAT_RX) if (rx.test(tl)) return n;
  return UNCLEAR;
}
function hoursOf(t) {
  let tl = t.toLowerCase().replace(/,/g, ".");
  tl = tl.replace(/\b\d{3,}\s*h(?:r|rs)?\s*service/g, "service").replace(/(\d+)h(\d+)\s*hrs?/g, "$1hr $2min");
  let h = 0, hit = false, m;
  for (const x of tl.matchAll(/(\d+(?:\.\d+)?)\s*(?:hrs?|hours?|h)\b\s*(\d+)?\s*(?:min)?/g)) { h += +x[1] + (x[2] ? +x[2] / 60 : 0); hit = true; }
  if (!hit && (m = tl.match(/(\d+(?:\.\d+)?)\s*(?:hrs?|h)(\d+)\s*min/))) { h = +m[1] + +m[2] / 60; hit = true; }
  if (!hit && (m = tl.match(/(\d+)\s*min/))) { const v = +m[1]; h = v >= 1 ? v / 60 : v; hit = true; }
  if (!hit && (m = tl.match(/(\d+)\.?\s*hr\s*(\d+)\s*m\b/))) { h = +m[1] + +m[2] / 60; hit = true; }
  if (!hit && (m = tl.match(/(\d+)\s*mim/))) { h = +m[1] / 60; hit = true; }
  if (!hit && (m = tl.match(/[a-z](\d+\.\d+)$/))) { h = +m[1]; hit = true; }
  if (!hit && (m = tl.match(/\((\d+(?:\.\d+)?)\)/))) { h = +m[1]; hit = true; }
  return hit ? Math.round(h * 100) / 100 : null;
}
const stripStatus = n => String(n || "").replace(/^\s*(not working|working)\s*[–-]?\s*/i, "");
function splitDelays(text) {
  return stripStatus(text).replace(/\n/g, " ").split(/\s*·\s*|\s*\/\s+|\s+-\s+(?=\d)|;\s*/).map(p => p.replace(/^[\s.\-·:]+|[\s.\-·:]+$/g, ""))
    .filter(p => p && !/^(diesel\s*:?\s*\d+%?|breakdowns:?|delays:?)$/i.test(p) && !/^(diesel:|breakdowns:)/i.test(p) && !/^metres (as per|left)|tally sheet|chat status|chat report|\(supervisor/i.test(p));
}
const hm = h => { if (h == null) return ""; const t = Math.round(h * 60); return Math.floor(t / 60) + " h " + String(t % 60).padStart(2, "0"); };

export function makeDrill(ctx) {
  const { db, fs, $, esc, fmt, num, r1, M1, ymd, parse, addDays, toast, fail, loadLib, RIGS, isAdmin, getMetres, getEmail, render, evidencePdf } = ctx;
  const { collection, doc, setDoc, deleteDoc, getDoc, getDocs, query, where } = fs;
  const today = () => ymd(new Date());
  const dow = s => parse(s).toLocaleDateString("en-ZA", { weekday: "short" });
  const monthName = m => MONTHS[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
  const monthEnd = m => { const d = new Date(+m.slice(0, 4), +m.slice(5, 7), 0); return ymd(d); };
  const shiftMonth = (m, n) => { const d = new Date(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1); return ymd(d).slice(0, 7); };
  let month = today().slice(0, 7), bdFrom = "", bdTo = "", xFrom = "", xTo = "", tharisa = "";
  let idx = null, spCfg = null, spState = "", reindexing = false;

  /* ---------- metres: one slot per date, shift and rig ---------- */
  const statusOf = x => x.status || (/^\s*not working/i.test(x.note || "") ? "Not working" : num(x.metres) > 0 || /^\s*working/i.test(x.note || "") ? "Working" : "");
  const delaysOf = x => Array.isArray(x.delays) ? x.delays.map(d => ({ text: d.text || d.cat, cat: d.cat || catOf(d.text || ""), hours: d.h == null ? null : +d.h }))
    : splitDelays(x.note).map(t => ({ text: t, cat: catOf(t), hours: hoursOf(t) }));
  const delayText = x => Array.isArray(x.delays) ? x.delays.map(d => (d.text || d.cat) + (d.h ? ` (${hm(+d.h)})` : "")).join(" · ") : stripStatus(x.note);
  function slots(from, to) {
    const out = new Map();
    for (const x of getMetres()) { if (x.date < from || x.date > to || !RIGS.includes(x.rig)) continue;
      const k = x.date + "|" + (x.shift || "Day") + "|" + x.rig; let s = out.get(k);
      if (!s) out.set(k, s = { date: x.date, shift: x.shift || "Day", rig: x.rig, metres: 0, lines: [] });
      s.metres += num(x.metres); s.lines.push(x); }
    for (const s of out.values()) { s.metres = r1(s.metres); const L = s.lines;
      s.status = L.map(statusOf).find(v => v === "Working") || L.map(statusOf).find(Boolean) || (s.metres > 0 ? "Working" : "");
      s.supervisor = L.map(l => l.supervisor).find(Boolean) || ""; s.block = [...new Set(L.map(l => l.block).filter(Boolean))].join(", ");
      s.delays = L.flatMap(delaysOf); s.delayText = L.map(delayText).filter(Boolean).join(" · ");
      s.hrs = L.map(l => l.hrsFrom || l.hrsTo ? `${l.hrsFrom || ""}–${l.hrsTo || ""}` : "").find(Boolean) || ""; }
    return out;
  }
  const daysOf = (from, to) => { const a = []; for (let d = from; d <= to; d = addDays(d, 1)) a.push(d); return a; };

  /* ---------- documents index: one light record per day, the files themselves are read only when opened ---------- */
  async function loadIndex() {
    try { const s = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "day"))); idx = {};
      s.docs.forEach(d => { const v = d.data(); idx[v.ev_date] = { n: v.n || 0, items: v.items || null }; }); render(); reindexOld(); }
    catch (e) { console.warn(e); idx = idx || {}; }
  }
  const ensureIndex = () => { if (idx === null) { idx = {}; loadIndex(); loadSp(); } };
  const lite = (id, v) => ({ id, type: v.type || "Tally sheet", rig: v.rig || "", shift: v.shift || "", name: v.name || "", note: v.note || "", pdf: /pdf/i.test(v.mime || "") ? 1 : 0, sp: v.sp?.url || "", at: v.at || 0, by: v.by || "" });
  async function daySheets(date) {
    const s = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "sheet"), where("ev_date", "==", date)));
    return s.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => DOC_TYPES.indexOf(a.type || "Tally sheet") - DOC_TYPES.indexOf(b.type || "Tally sheet") || (a.rig + (a.shift === "Night" ? 2 : 1)).localeCompare(b.rig + (b.shift === "Night" ? 2 : 1)) || (a.at || 0) - (b.at || 0));
  }
  async function indexDay(date, list) {
    list = list || await daySheets(date); const items = list.map(x => lite(x.id, x));
    idx = idx || {}; idx[date] = { n: items.length, items };
    if (isAdmin()) { const ref = doc(db, "fcSettings", "evday_" + date);
      if (items.length) await setDoc(ref, { ev_kind: "day", ev_date: date, n: items.length, items }); else await deleteDoc(ref); }
    if (!items.length) delete idx[date];
    return list;
  }
  async function reindexOld() {            // days filed before the documents list existed: read once and list their files
    if (reindexing || !isAdmin() || !navigator.onLine) return; reindexing = true;
    try { for (const d of Object.keys(idx).filter(d => !idx[d].items).sort().reverse()) await indexDay(d); render(); } catch (e) { console.warn(e); } finally { reindexing = false; }
  }
  const docsIn = (from, to) => Object.keys(idx || {}).filter(d => d >= from && d <= to).sort().flatMap(d => (idx[d].items || []).map(x => ({ ...x, date: d })));
  const slotDocs = (date, shift, rig, type) => ((idx || {})[date]?.items || []).filter(x => x.rig === rig && x.shift === shift && (!type || x.type === type));
  const hasTally = s => { const d = (idx || {})[s.date]; return d ? (d.items ? d.items.some(x => x.type === "Tally sheet" && x.rig === s.rig && x.shift === s.shift) : null) : false; };

  /* ---------- SharePoint ---------- */
  async function loadSp() {
    try { const s = await getDoc(doc(db, "fcSettings", "sharepoint")); spCfg = s.exists() ? s.data() : {};
      if (spConfigured(spCfg)) { try { await spInit(spCfg); spState = ""; } catch (e) { console.warn(e); spState = "The Microsoft sign-in couldn't load. Check the signal."; } }
      render(); } catch (e) { console.warn(e); spCfg = {}; }
  }
  /* SharePoint: one folder per month, one folder per day and shift, the way the METRES folder is kept: "November 2026 Metres" › "01 DAY", "01 NIGHT" */
  const monthFolder = date => `Drilling/${MONTHS[+date.slice(5, 7) - 1]} ${date.slice(0, 4)} Metres`;
  const spFolder = (date, type, shift) => `${monthFolder(date)}/${date.slice(8, 10)} ${shift ? shift.toUpperCase() : "DAY AND NIGHT"}`;
  function spStrip() {
    if (!spCfg) return "";
    const P = spPlace();
    if (!spConfigured(spCfg)) return `<div class="spbar off"><b>OneDrive is not linked yet.</b> Files are kept in the app for now.${isAdmin() ? " It is set up once under Settings → OneDrive." : ""}</div>`;
    if (!spSigned()) return `<div class="spbar off"><b>${P}: not signed in on this device.</b> Files you add now stay in the app only. <button class="primary sm" data-sp="in">Sign in with your TKR Microsoft account</button> <span class="muted">${esc(spState)}</span></div>`;
    if (!spLinked()) return `<div class="spbar off"><b>Signed in as ${esc(spWho())}, but the OneDrive folder is not made yet.</b> ${isAdmin() ? `<button class="primary sm" data-sp="link">Make the TKR App folder in my OneDrive</button>` : "Ask Muneer to link it."} <span class="muted">${esc(spState)}</span></div>`;
    return `<div class="spbar on"><b>Saving to ${P}</b>${P === "OneDrive" && spCfg.owner ? ` (${esc(spCfg.owner)})` : ""} · signed in as ${esc(spWho())}. New files go to <b>${esc(spRoot())} › ${monthFolder(month + "-01").replace("/", " › ")} › 01 DAY, 01 NIGHT, 02 DAY…</b> <button class="sm" data-sp="open">Open the folder</button></div>`;
  }
  function spCard() {
    ensureIndex(); const c = spCfg || {}, redirect = new URL("../auth.html", location.href).href;
    return `<section class="card" id="spCard"><h2>OneDrive</h2>
      <p class="muted" style="margin-top:0">Every tally sheet, job card, checklist, breakdown report and supporting document added in the app is also saved in your TKR OneDrive (1 TB): <b>TKR App › Drilling › November 2026 Metres › 01 DAY, 01 NIGHT…</b>, with the Excel reports in the month folder.</p>
      <form id="fSp" class="grid" autocomplete="off">
        <label>Directory (tenant) ID <input name="tenant" value="${esc(c.tenant || "")}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"></label>
        <label>Application (client) ID <input name="clientId" value="${esc(c.clientId || "")}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"></label>
        <label>Main folder name <input name="folder" value="${esc(c.folder || "TKR App")}"></label>
        <label style="grid-column:1/-1">SharePoint site address (leave empty to use OneDrive) <input name="site" value="${esc(c.site || "")}" placeholder="empty = OneDrive"></label>
        <div style="display:flex;align-items:flex-end"><button class="primary" style="width:100%">Save</button></div>
      </form>
      <div style="margin-top:12px">${spStrip()}</div>
      ${spConfigured(c) && spReady() ? `<div class="row" style="margin-top:8px"><button data-sp="move">Send the files already in the app to ${spPlace()}</button><button class="ghost" data-sp="out">Sign out of Microsoft</button> <span id="spMsg" class="muted"></span></div>` : ""}
      <details style="margin-top:12px" ${spConfigured(c) ? "" : "open"}><summary><b>How to set it up (once, about 10 minutes)</b></summary>
        <ol style="margin:8px 0 0;padding-left:20px;line-height:1.7">
          <li>On a computer, go to <b>entra.microsoft.com</b> and sign in with your TKR email.</li>
          <li>Open <b>Applications → App registrations → New registration</b>. Name: <b>TKR Operations App</b>. Who can use it: <b>Accounts in this organizational directory only</b>.</li>
          <li>Under Redirect URI choose <b>Single-page application (SPA)</b> and paste: <b style="overflow-wrap:anywhere">${esc(redirect)}</b>. Tap Register.</li>
          <li>On the Overview page copy the <b>Application (client) ID</b> and the <b>Directory (tenant) ID</b> into the boxes above and tap Save.</li>
          <li>Open <b>API permissions → Add a permission → Microsoft Graph → Delegated</b>, tick <b>Files.ReadWrite.All</b>, Add.</li>
          <li>Back here, tap <b>Sign in with your TKR Microsoft account</b> and accept, then <b>Make the TKR App folder in my OneDrive</b>.</li>
          <li>For Amogelang to save too: in OneDrive, share the <b>TKR App</b> folder with her (Can edit). She signs in once on her phone.</li>
        </ol>
        <p class="muted">If Microsoft says you don't have permission to register an app, or asks for administrator approval, send these steps to whoever looks after TKR's Microsoft 365. It is free.</p></details></section>`;
  }
  async function linkDrive() {
    const L = await spLinkOneDrive(); spCfg = { ...spCfg, ...L, linkedAt: Date.now() };
    await setDoc(doc(db, "fcSettings", "sharepoint"), spCfg); await spInit(spCfg);
  }
  async function moveAllToSp() {
    const say = t => { const el = $("spMsg"); if (el) el.textContent = t; else toast(t); };
    let n = 0, skip = 0; const dates = Object.keys(idx || {}).sort();
    for (const d of dates) { const L = await daySheets(d); let ch = false;
      for (const x of L) { if (x.sp?.url) { skip++; continue; } const data = await fullData(x); if (!data) continue;
        say(`Sending ${fmt(d)}… ${n + 1}`); const blob = await (await fetch(data)).blob();
        const sp = await spUpload(spFolder(d, x.type || "Tally sheet", x.shift), fileName(d, x.rig, x.shift, x.type || "Tally sheet", x.name || (/pdf/.test(x.mime || "") ? "x.pdf" : "x.jpg"), x.id.slice(-6)), blob);
        await setDoc(doc(db, "fcSettings", x.id), { sp }, { merge: true }); x.sp = sp; n++; ch = true; }
      if (ch) await indexDay(d, L); }
    say(`Done: ${n} files sent to ${spPlace()}${skip ? `, ${skip} were already there` : ""}.`); render();
  }

  /* ---------- saving files ---------- */
  const dataUrl = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  async function shrink(src, max = 1400, q0 = 0.7, cap = 900000) {
    const url = typeof src === "string" ? src : URL.createObjectURL(src);
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error("image")); i.src = url; });
    const k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement("canvas"); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
    const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height); if (typeof src !== "string") URL.revokeObjectURL(url);
    let q = q0, d = c.toDataURL("image/jpeg", q); while (d.length > cap && q > 0.3) { q -= 0.1; d = c.toDataURL("image/jpeg", q); }
    return { data: d, w: c.width, h: c.height };
  }
  const fileName = (date, rig, shift, type, orig, tag) => `${rig || "All machines"} - ${type} - ${date.slice(8, 10)} ${MONTHS[+date.slice(5, 7) - 1].slice(0, 3)} ${shift ? shift.toLowerCase() : "day and night"} ${tag}.${(String(orig).match(/\.([a-z0-9]{2,5})$/i) || [, "jpg"])[1].toLowerCase()}`;
  async function fullData(x) {
    if (!x.parts) return x.data || "";
    const s = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "part"), where("parent", "==", x.id)));
    return s.docs.map(d => d.data()).sort((a, b) => a.i - b.i).map(p => p.data).join("");
  }
  const CHUNK = 700000;
  async function addFile(file, { date, rig = "", shift = "", type = "Tally sheet", note = "" }, reindex = true) {
    const isPdf = file.type === "application/pdf" || /\.pdf$/i.test(file.name), isImg = !isPdf && (/^image\//.test(file.type) || /\.(jpe?g|png|webp|gif|bmp)$/i.test(file.name));
    if (!isPdf && !isImg) throw Object.assign(new Error("type"), { user: `${file.name}: choose a photo or a PDF.` });
    const now = new Date(), tag = String(now.getHours()).padStart(2, "0") + String(now.getMinutes()).padStart(2, "0") + String(now.getSeconds()).padStart(2, "0") + "-" + Math.random().toString(36).slice(2, 4);
    const id = ("ev_" + date + "_" + (rig || "All").replace(/\s/g, "") + "_" + (shift || "NA") + "_" + type.replace(/\s/g, "") + "_" + tag).replace(/[^\w-]/g, "_");
    let sp = null, spErr = false;
    if (spReady()) { try { sp = await spUpload(spFolder(date, type, shift), fileName(date, rig, shift, type, file.name, tag), file); } catch (e) { console.warn(e); spErr = true; } }
    let data = "", parts = 0, rest = [];
    if (isImg) { try { data = (await shrink(file, sp ? 1200 : 1400)).data; } catch { throw Object.assign(new Error("img"), { user: `${file.name}: this photo type can't be read here. Send it as JPG or PDF.` }); } }
    else { const d = await dataUrl(file);
      if (d.length <= 900000) data = d;
      else if (!sp) { if (file.size > 8 * 1024 * 1024) throw Object.assign(new Error("big"), { user: `${file.name} is bigger than 8 MB. Make the PDF smaller, or link OneDrive.` });
        for (let i = 0; i < d.length; i += CHUNK) rest.push(d.slice(i, i + CHUNK)); parts = rest.length; } }
    for (let i = 0; i < rest.length; i++) await setDoc(doc(db, "fcSettings", id + "_p" + i), { ev_kind: "part", parent: id, i, data: rest[i] });
    await setDoc(doc(db, "fcSettings", id), { ev_kind: "sheet", ev_date: date, rig, shift, type, name: file.name, mime: isPdf ? "application/pdf" : "image/jpeg", note, data, parts, sp, by: getEmail(), at: Date.now() });
    if (reindex) await indexDay(date);
    return { id, sp: !!sp, spErr };
  }
  async function addFiles(files, meta) {
    let n = 0, sp = 0, err = 0; const bad = [];
    for (const f of files) { toast(`Saving ${n + 1} of ${files.length}…`); try { const r = await addFile(f, meta, false); n++; if (r.sp) sp++; if (r.spErr) err++; } catch (e) { console.warn(e); bad.push(e.user || `${f.name} didn't save.`); } }
    if (n) await indexDay(meta.date);
    toast(bad.length ? bad[0] : `${n} file${n === 1 ? "" : "s"} saved${sp ? " and filed in " + spPlace() : ""}${err ? ". " + spPlace() + " didn't take " + err + "; they are kept in the app" : ""}`);
    return n;
  }
  async function removeFile(id, date) {
    const s = await getDoc(doc(db, "fcSettings", id)); const v = s.exists() ? s.data() : {};
    if (v.parts) for (let i = 0; i < v.parts; i++) await deleteDoc(doc(db, "fcSettings", id + "_p" + i));
    await deleteDoc(doc(db, "fcSettings", id)); if (v.sp?.id && spReady()) await spDelete(v.sp.id);
    await indexDay(date);
  }
  async function openFile(id) {
    const s = await getDoc(doc(db, "fcSettings", id)); if (!s.exists()) { toast("That file is not there any more."); return; }
    const v = { id, ...s.data() }, data = await fullData(v);
    if (!data) { if (v.sp?.url) window.open(v.sp.url, "_blank", "noopener"); else toast("The file couldn't open."); return; }
    if (/pdf/.test(v.mime || "")) { const b = await (await fetch(data)).blob(); window.open(URL.createObjectURL(b), "_blank"); return; }
    const ov = document.createElement("div"); ov.className = "dv"; ov.innerHTML = `<div class="dvbar"><b>${esc(v.type || "Tally sheet")} · ${fmt(v.ev_date)} · ${esc(v.rig || "All machines")} ${esc(v.shift || "")}</b><button data-x>Close</button></div><img src="${data}" alt="">`;
    ov.onclick = e => { if (e.target.closest("[data-x]") || e.target === ov) ov.remove(); }; document.body.appendChild(ov);
  }

  /* ---------- Metres tab ---------- */
  const chip = s => s === "Working" ? '<span class="chip ok">Working</span>' : s === "Not working" ? '<span class="chip bad">Not working</span>' : "";
  function monthBar() {
    return `<div class="mbar"><button data-mnav="-1" aria-label="Month before">‹</button><b>${monthName(month)}</b><button data-mnav="1" aria-label="Next month" ${month >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
  }
  function metresView() {
    ensureIndex();
    const from = month + "-01", to = monthEnd(month) < today() ? monthEnd(month) : today(), S = slots(from, monthEnd(month)), all = [...S.values()];
    const tot = r => r1(all.filter(s => !r || s.rig === r).reduce((n, s) => n + s.metres, 0));
    const noTally = all.filter(s => s.metres > 0 && hasTally(s) === false);
    const adm = isAdmin(), days = from <= to ? daysOf(from, to).reverse() : [];
    let run = tot(); const mtd = {}; for (const d of days) { mtd[d] = run; run = r1(run - all.filter(s => s.date === d).reduce((n, s) => n + s.metres, 0)); }
    const cell = (d, sh, rig) => { const s = S.get(d + "|" + sh + "|" + rig), n = slotDocs(d, sh, rig).length, t = s ? hasTally(s) : false;
      const tag = adm ? "button" : "div";
      if (!s) return `<${tag} class="sh empty" data-shift="${d}|${sh}|${rig}"><span class="lab">${sh}</span><span class="m">–</span>${adm ? '<span class="add">+ Capture</span>' : ""}</${tag}>`;
      return `<${tag} class="sh ${s.status === "Not working" ? "down" : s.metres > 0 ? "up" : ""}" data-shift="${d}|${sh}|${rig}"><span class="lab">${sh}${s.supervisor ? " · " + esc(s.supervisor) : ""}</span>
        <span class="m">${M1(s.metres)}<small> m</small></span><span class="st">${chip(s.status)}${n ? `<span class="chip doc">${n} file${n === 1 ? "" : "s"}</span>` : ""}${s.metres > 0 && t === false ? '<span class="chip warn">No tally sheet</span>' : ""}</span>
        ${s.delayText ? `<span class="dl">${esc(s.delayText)}</span>` : ""}</${tag}>`; };
    const dayCard = d => { const dt = r1(all.filter(s => s.date === d).reduce((n, s) => n + s.metres, 0)), nd = idx?.[d]?.n || 0;
      if (!all.some(s => s.date === d)) return `<section class="day slim"><header><div><b>${dow(d)} ${fmt(d)}</b> <span class="nc">Nothing captured</span></div><div class="dt">${adm ? `<button class="sm" data-shift="new|${d}">+ Capture a shift</button>` : ""}<button class="sm ${nd ? "primary" : ""}" data-day="${d}">${nd ? `Documents (${nd})` : adm ? "+ Documents" : "No documents"}</button></div></header></section>`;
      return `<section class="day"><header><div><b>${dow(d)} ${fmt(d)}</b></div><div class="dt"><span>Day total <b>${M1(dt)} m</b></span><span>Month to date <b>${M1(mtd[d])} m</b></span>
        <button class="sm ${nd ? "primary" : ""}" data-day="${d}">${nd ? `Documents (${nd})` : adm ? "+ Documents" : "No documents"}</button></div></header>
        <div class="rigs">${RIGS.map(r => { const rt = r1(["Day", "Night"].reduce((n, sh) => n + (S.get(d + "|" + sh + "|" + r)?.metres || 0), 0));
          return `<div class="rig"><div class="rn"><b>${r}</b><span>${M1(rt)} m</span></div>${cell(d, "Day", r)}${cell(d, "Night", r)}</div>`; }).join("")}</div></section>`; };
    const first = getMetres().map(x => x.date).sort()[0] || from;
    xFrom = xFrom || from; xTo = xTo || to;
    return `<div class="mhead">${monthBar()}<div class="row">${adm ? `<button class="primary" data-shift="new">+ Capture a shift</button>` : ""}</div></div>
      <div class="tiles">${tileH("Month total", M1(tot()) + " m", `${fmt(from)} to ${fmt(to)}`, "dark")}${RIGS.map(r => tileH(r, M1(tot(r)) + " m", `Shifts working ${all.filter(s => s.rig === r && s.status === "Working").length} · not working ${all.filter(s => s.rig === r && s.status === "Not working").length}`)).join("")}
        ${tileH("Shifts without a tally sheet", String(noTally.length), noTally.length ? noTally.slice(0, 3).map(s => `${fmt(s.date).slice(0, 5)} ${s.rig.slice(4)} ${s.shift}`).join(", ") + (noTally.length > 3 ? "…" : "") : "every shift with metres has one", noTally.length ? "alert" : "")}</div>
      ${spStrip()}
      <section class="card dl-card"><h2>Download metres (Excel)</h2>
        <p class="muted" style="margin-top:0">Summary, every shift with its status and delays, daily metres with the tally sheets, machine performance, delay summary, supporting documents and comments.</p>
        <div class="row" style="align-items:flex-end"><label style="flex:1 1 140px">From <input type="date" id="xFrom" value="${xFrom}" min="${first}" max="${today()}"></label>
          <label style="flex:1 1 140px">To <input type="date" id="xTo" value="${xTo}" min="${first}" max="${today()}"></label>
          <label style="flex:1 1 180px">Tharisa's metres for these dates (if you have it) <input id="xTh" inputmode="decimal" value="${esc(tharisa)}" placeholder="optional"></label>
          <button class="primary" id="xGo" style="flex:1 1 200px;min-height:48px">Download metres Excel</button></div></section>
      ${days.map(dayCard).join("") || `<div class="card muted">No days to show for ${monthName(month)}.</div>`}
      ${adm ? `<details class="card"><summary><b>All metres lines for ${monthName(month)}</b> (delete a wrong line)</summary><div class="list" style="margin-top:8px">${getMetres().filter(x => x.date >= from && x.date <= monthEnd(month)).sort((a, b) => b.date.localeCompare(a.date) || String(a.shift).localeCompare(String(b.shift)) || a.rig.localeCompare(b.rig))
        .map(x => `<div class="it"><span>${fmt(x.date)} ${esc(x.shift || "")} · <b>${esc(x.rig)}</b> · ${M1(num(x.metres))} m ${x.note ? `<span class="muted">· ${esc(x.note)}</span>` : ""}</span>${x.builtin ? '<span class="muted">closed month</span>' : `<button class="danger" data-del="fcMetres|${esc(x.id)}">Delete</button>`}</div>`).join("") || '<div class="muted">Nothing yet.</div>'}</div></details>` : ""}`;
  }
  const tileH = (l, n, s, cls = "") => `<div class="tile ${cls}"><div class="l">${l}</div><div class="n">${n}</div>${s ? `<div class="s">${s}</div>` : ""}</div>`;

  /* ---------- capture or edit one shift for one machine ---------- */
  function shiftPanel(date, shift, rig, isNew) {
    const S = slots(date, date), s = S.get(date + "|" + shift + "|" + rig), line = s?.lines.find(l => !l.builtin), ro = !isAdmin() || (s && !line);
    const d0 = line && Array.isArray(line.delays) ? line.delays : s ? s.delays.map(d => ({ cat: DELAY_CATS.includes(d.cat) ? d.cat : "Other", h: d.hours, text: d.text })) : [];
    const files = slotDocs(date, shift, rig);
    document.getElementById("shPanel")?.remove();
    const ov = document.createElement("div"); ov.id = "shPanel"; ov.className = "scrim"; ov.setAttribute("role", "dialog");
    const delayRow = (d = {}) => { const t = d.h == null ? "" : Math.round(+d.h * 60); return `<div class="drow"><select data-k="cat">${DELAY_CATS.map(c => `<option ${c === (d.cat || "Breakdowns") ? "selected" : ""}>${c}</option>`).join("")}</select>
      <input data-k="hh" inputmode="numeric" placeholder="h" value="${t === "" ? "" : Math.floor(t / 60)}" aria-label="Hours"><input data-k="mm" inputmode="numeric" placeholder="min" value="${t === "" ? "" : t % 60}" aria-label="Minutes">
      <input data-k="text" placeholder="What happened" value="${esc(d.text || "")}"><button type="button" class="danger sm" data-drm aria-label="Remove">×</button></div>`; };
    ov.innerHTML = `<form class="panel" id="fShift" autocomplete="off">
      <h2>${ro ? "Shift" : s ? "Edit shift" : "Capture a shift"}</h2>
      ${ro && isAdmin() ? '<p class="muted" style="margin:0 0 8px">This month is closed, so the figures can\'t be changed here. You can still add documents.</p>' : ""}
      <div class="grid">
        <label>Date <input type="date" name="date" value="${date}" max="${today()}" ${isNew ? "" : "disabled"} required></label>
        <label>Shift <select name="shift" ${isNew ? "" : "disabled"}>${["Day", "Night"].map(x => `<option ${x === shift ? "selected" : ""}>${x}</option>`).join("")}</select></label>
        <label>Machine <select name="rig" ${isNew ? "" : "disabled"}>${RIGS.map(x => `<option ${x === rig ? "selected" : ""}>${x}</option>`).join("")}</select></label>
        <label>Supervisor <input name="supervisor" value="${esc(s?.supervisor || "")}" ${ro ? "disabled" : ""}></label>
        <label>Status <select name="status" ${ro ? "disabled" : ""}>${["Working", "Not working"].map(x => `<option ${x === (s?.status || "Working") ? "selected" : ""}>${x}</option>`).join("")}</select></label>
        <label>Metres drilled <input name="metres" inputmode="decimal" value="${s ? s.metres : ""}" ${ro ? "disabled" : ""}></label>
        <label>Block <input name="block" value="${esc(s?.block || "")}" ${ro ? "disabled" : ""}></label>
        <label>Machine hours at start <input name="hrsFrom" inputmode="decimal" value="${esc(line?.hrsFrom ?? "")}" ${ro ? "disabled" : ""}></label>
        <label>Machine hours at end <input name="hrsTo" inputmode="decimal" value="${esc(line?.hrsTo ?? "")}" ${ro ? "disabled" : ""}></label>
      </div>
      <h3>Breakdowns and delays</h3>
      ${ro ? `<p style="margin:0">${esc(s?.delayText || "None written")}</p>` : `<div id="dRows">${d0.map(delayRow).join("")}</div><button type="button" class="sm" id="dAdd">+ Add a breakdown or delay</button>`}
      <h3>Documents for this shift</h3>
      ${files.length ? `<div class="files">${files.map(f => `<div class="file"><span><b>${esc(f.type)}</b>${f.name ? " · " + esc(f.name) : ""}</span><button type="button" class="sm" data-open="${f.id}">View</button></div>`).join("")}</div>` : '<p class="muted" style="margin:0 0 6px">Nothing attached yet.</p>'}
      ${isAdmin() ? `<div class="upl">${DOC_TYPES.map(t => `<label class="pick"><span>${t}${t === "Tally sheet" ? "" : "s"}</span><input type="file" data-type="${t}" accept="image/*,application/pdf" multiple><em></em></label>`).join("")}</div>
      <p class="muted" style="margin:6px 0 0">Photos or PDF. ${spReady() ? "They are filed in ${spPlace()} under the month, the date and the shift." : ""}</p>` : ""}
      <div class="acts"><button type="button" class="ghost" data-close>Close</button>${isAdmin() ? `<button class="primary">${ro ? "Save documents" : "Save shift"}</button>` : ""}</div></form>`;
    document.body.appendChild(ov);
    ov.addEventListener("click", async e => {
      if (e.target === ov || e.target.closest("[data-close]")) { ov.remove(); return; }
      if (e.target.id === "dAdd") { ov.querySelector("#dRows").insertAdjacentHTML("beforeend", delayRow()); return; }
      const rm = e.target.closest("[data-drm]"); if (rm) { rm.parentElement.remove(); return; }
      const op = e.target.closest("[data-open]"); if (op) { try { await openFile(op.dataset.open); } catch (x) { console.warn(x); toast("The file couldn't open. Check the signal."); } } });
    ov.addEventListener("change", e => { if (e.target.type === "file") e.target.parentElement.querySelector("em").textContent = e.target.files.length ? e.target.files.length + " chosen" : ""; });
    ov.querySelector("form").addEventListener("submit", async e => {
      e.preventDefault(); e.stopPropagation(); const f = e.target, v = k => f.elements[k].value, btn = f.querySelector("button.primary");
      const D = v("date"), SH = v("shift"), RG = v("rig"); if (!D) return;
      const picks = [...f.querySelectorAll("input[type=file]")].flatMap(i => [...i.files].map(file => ({ file, type: i.dataset.type })));
      try {
        if (!ro) {
          const delays = [...f.querySelectorAll(".drow")].map(r => { const g = k => r.querySelector(`[data-k=${k}]`).value.trim(), hh = g("hh"), mm = g("mm");
            return { cat: g("cat"), h: hh === "" && mm === "" ? null : Math.round((num(hh) + num(mm) / 60) * 100) / 100, text: g("text") }; });
          const metres = num(v("metres")), status = v("status");
          if (status === "Working" && metres > 0 && !picks.some(p => p.type === "Tally sheet") && !slotDocs(D, SH, RG, "Tally sheet").length && !confirm("There is no tally sheet for this shift. Save without it?")) return;
          btn.disabled = true; btn.textContent = "Saving…";
          const cur = slots(D, D).get(D + "|" + SH + "|" + RG)?.lines.find(l => !l.builtin);
          if (!cur && slots(D, D).get(D + "|" + SH + "|" + RG)) { toast("That month is closed. The figures can't be changed."); btn.disabled = false; btn.textContent = "Save shift"; return; }
          const note = status + (delays.length ? " – " + delays.map(d => (d.text || d.cat) + (d.h ? ` (${hm(d.h)})` : "")).join(" · ") : "");
          await setDoc(doc(db, "fcMetres", cur?.id || `sh-${D}-${SH}-${RG.replace(/\s/g, "")}`), { date: D, shift: SH, rig: RG, metres, block: v("block").trim(), supervisor: v("supervisor").trim(), status, delays, note,
            hrsFrom: v("hrsFrom").trim(), hrsTo: v("hrsTo").trim(), by: getEmail(), createdAt: cur?.createdAt || Date.now(), editedAt: Date.now() });
        } else { btn.disabled = true; btn.textContent = "Saving…"; }
        for (const t of DOC_TYPES) { const fs2 = picks.filter(p => p.type === t).map(p => p.file); if (fs2.length) await addFiles(fs2, { date: D, rig: RG, shift: SH, type: t }); }
        if (!picks.length) toast("Shift saved");
        month = D.slice(0, 7); ov.remove(); render();
      } catch (x) { btn.disabled = false; btn.textContent = "Save shift"; fail(x); }
    });
  }

  /* ---------- Documents tab ---------- */
  let upDate = "";
  function docsView() {
    ensureIndex(); const from = month + "-01", to = monthEnd(month), list = docsIn(from, to), adm = isAdmin();
    const old = Object.keys(idx || {}).filter(d => d >= from && d <= to && !idx[d].items);
    const dates = [...new Set([...list.map(x => x.date), ...old])].sort().reverse();
    upDate = upDate || today();
    return `<div class="mhead">${monthBar()}</div>${spStrip()}
      <div class="tiles">${DOC_TYPES.map(t => tileH(FOLDER[t], String(list.filter(x => x.type === t).length), "")).join("")}</div>
      ${adm ? `<section class="card"><h2>Add documents</h2><form id="fDocs" class="grid" autocomplete="off">
        <label>Date <input type="date" name="date" value="${upDate}" max="${today()}" required></label>
        <label>Type of document <select name="type">${DOC_TYPES.map(t => `<option>${t}</option>`).join("")}</select></label>
        <label>Machine <select name="rig">${RIGS.map(r => `<option>${r}</option>`).join("")}<option value="">All machines</option></select></label>
        <label>Shift <select name="shift"><option>Day</option><option>Night</option><option value="">Whole day</option></select></label>
        <label style="grid-column:1/-1">Photos or PDF files <input type="file" name="files" accept="image/*,application/pdf" multiple required></label>
        <label style="grid-column:1/-1">Note (optional) <input name="note"></label>
        <div><button class="primary" style="width:100%">Save documents</button></div></form></section>` : ""}
      ${dates.map(d => { const L = list.filter(x => x.date === d);
        return `<section class="day"><header><div><b>${dow(d)} ${fmt(d)}</b></div><div class="dt"><span>${idx[d].n} file${idx[d].n === 1 ? "" : "s"}</span><button class="sm primary" data-day="${d}">Open the day</button></div></header>
          ${L.length ? DOC_TYPES.filter(t => L.some(x => x.type === t)).map(t => `<div class="dgrp"><div class="dg">${FOLDER[t]}</div><div class="dfiles">${L.filter(x => x.type === t).map(x => `<span class="dfile"><button class="lnk" data-open="${x.id}">${x.pdf ? "PDF" : "Photo"} · ${esc(x.rig || "All machines")}${x.shift ? " " + esc(x.shift) : ""}</button>${x.sp ? `<a href="${esc(x.sp)}" target="_blank" rel="noopener" title="Open in OneDrive">OneDrive</a>` : ""}${x.note ? `<small>${esc(x.note)}</small>` : ""}</span>`).join("")}</div></div>`).join("")
            : '<div class="muted" style="padding:10px 14px">Tally sheets. Tap Open the day to see them.</div>'}</section>`; }).join("") || `<div class="card muted">No documents for ${monthName(month)} yet.</div>`}`;
  }
  /* everything filed for one day, with the photos */
  async function showDay(date) {
    toast("Opening the documents…"); const L = await daySheets(date); if (isAdmin() || !idx[date]?.items) { try { await indexDay(date, L); } catch (e) { console.warn(e); } }
    for (const x of L) if (x.parts && !/pdf/.test(x.mime || "")) x.data = await fullData(x);
    document.getElementById("evView")?.remove(); const adm = isAdmin();
    const ov = document.createElement("div"); ov.id = "evView"; ov.className = "dayv"; ov.setAttribute("role", "dialog");
    ov.innerHTML = `<div class="dvbar"><b>${dow(date)} ${fmt(date)} · ${L.length} document${L.length === 1 ? "" : "s"}</b><span class="row">${L.some(x => (x.type || "Tally sheet") === "Tally sheet") ? `<button class="primary" data-pdf>Tally sheets PDF</button>` : ""}<button data-x>Close</button></span></div>
      <div class="dvin">${adm ? `<form class="dvadd" id="fDayAdd"><b>Add to ${fmt(date)}</b><div class="row">
          <select name="type">${DOC_TYPES.map(t => `<option>${t}</option>`).join("")}</select><select name="rig">${RIGS.map(r => `<option>${r}</option>`).join("")}<option value="">All machines</option></select>
          <select name="shift"><option>Day</option><option>Night</option><option value="">Whole day</option></select><input type="file" name="files" accept="image/*,application/pdf" multiple required><button class="primary">Save</button></div></form>` : ""}
      ${DOC_TYPES.filter(t => L.some(x => (x.type || "Tally sheet") === t)).map(t => `<h3>${FOLDER[t]}</h3>` + L.filter(x => (x.type || "Tally sheet") === t).map(x => `<figure><figcaption><span><b>${esc(x.rig || "All machines")}</b>${x.shift ? " · " + esc(x.shift) + " shift" : ""}${x.note ? " – " + esc(x.note) : ""}${x.name ? ` <small>${esc(x.name)}</small>` : ""}</span>
          <span class="row">${x.sp?.url ? `<a class="btn" href="${esc(x.sp.url)}" target="_blank" rel="noopener">${spPlace()}</a>` : ""}${/pdf/.test(x.mime || "") || !x.data ? `<button data-open="${x.id}">Open ${/pdf/.test(x.mime || "") ? "PDF" : "file"}</button>` : ""}${adm ? `<button class="danger" data-rm="${x.id}">Remove</button>` : ""}</span></figcaption>
          ${x.data && !/pdf/.test(x.mime || "") ? `<img src="${x.data}" alt="${esc(t)} ${esc(x.rig)} ${esc(x.shift)}" loading="lazy">` : ""}</figure>`).join("")).join("") || '<p class="none">Nothing filed for this day yet.</p>'}</div>`;
    document.body.appendChild(ov);
    ov.addEventListener("click", async e => { const b = e.target.closest("button"); if (!b) return;
      if (b.hasAttribute("data-x")) { ov.remove(); return; }
      if (b.hasAttribute("data-pdf")) { b.disabled = true; try { await evidencePdf(date); } catch (x) { console.warn(x); toast("The PDF couldn't be made."); } b.disabled = false; return; }
      if (b.dataset.open) { try { await openFile(b.dataset.open); } catch (x) { console.warn(x); toast("The file couldn't open."); } return; }
      if (b.dataset.rm) { if (!confirm("Remove this document from " + fmt(date) + "?" + (spReady() ? " It is also removed from " + spPlace() + "." : ""))) return; b.disabled = true;
        try { await removeFile(b.dataset.rm, date); toast("Removed"); render(); await showDay(date); } catch (x) { b.disabled = false; fail(x); } } });
    ov.querySelector("#fDayAdd")?.addEventListener("submit", async e => { e.preventDefault(); e.stopPropagation(); const f = e.target, btn = f.querySelector("button");
      btn.disabled = true; try { await addFiles([...f.elements.files.files], { date, type: f.elements.type.value, rig: f.elements.rig.value, shift: f.elements.shift.value }); render(); await showDay(date); } catch (x) { btn.disabled = false; fail(x); } });
  }

  /* ---------- Breakdown report tab ---------- */
  function bdData(from, to) {
    const all = [...slots(from, to).values()].sort((a, b) => a.date.localeCompare(b.date) || a.shift.localeCompare(b.shift) || a.rig.localeCompare(b.rig));
    const rigs = RIGS.filter(r => all.some(s => s.rig === r));
    const mach = rigs.map(r => { const L = all.filter(s => s.rig === r); return { rig: r, working: L.filter(s => s.status === "Working").length, down: L.filter(s => s.status === "Not working").length, metres: r1(L.reduce((n, s) => n + s.metres, 0)),
      hours: Math.round(L.flatMap(s => s.delays).reduce((n, d) => n + (d.hours || 0), 0) * 10) / 10 }; });
    const lines = all.flatMap(s => s.delays.map(d => ({ ...d, date: s.date, shift: s.shift, rig: s.rig })));
    const cats = [...new Set(lines.map(l => l.cat))].map(c => { const L = lines.filter(l => l.cat === c), per = rigs.map(r => Math.round(L.filter(l => l.rig === r).reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10);
      return { cat: c, per, times: rigs.map(r => L.filter(l => l.rig === r).length), total: Math.round(L.reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10 }; }).sort((a, b) => b.total - a.total || b.times.reduce((x, y) => x + y, 0) - a.times.reduce((x, y) => x + y, 0));
    return { all, rigs, mach, lines, cats, total: Math.round(lines.reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10 };
  }
  const H1 = v => (Math.round(v * 10) / 10).toFixed(1);
  function bdView() {
    ensureIndex(); bdFrom = bdFrom || month + "-01"; bdTo = bdTo || (monthEnd(month) < today() ? monthEnd(month) : today());
    const B = bdData(bdFrom, bdTo), max = Math.max(1, ...B.cats.map(c => c.total)), col = ["#d4a12a", "#3a3a3a", "#2e7d4f"];
    const reports = docsIn(bdFrom, bdTo).filter(x => x.type === "Breakdown report" || x.type === "Job card");
    return `<section class="card dl-card"><h2>Breakdown report</h2>
        <div class="row" style="align-items:flex-end"><label style="flex:1 1 150px">From <input type="date" id="bdFrom" value="${bdFrom}" max="${today()}"></label><label style="flex:1 1 150px">To <input type="date" id="bdTo" value="${bdTo}" max="${today()}"></label>
        <button class="primary" id="bdGo" style="flex:1 1 220px;min-height:48px">Download breakdown report (Excel)</button></div></section>
      <div class="tiles">${tileH("Time lost", H1(B.total) + " h", `${B.lines.length} breakdowns and delays written`, "dark")}${B.mach.map(m => tileH(m.rig, H1(m.hours) + " h lost", `Shifts working ${m.working} · not working ${m.down} · ${M1(m.metres)} m`)).join("")}
        ${tileH("Biggest cause", B.cats[0] ? `<span style="font-size:22px">${esc(B.cats[0].cat)}</span>` : "–", B.cats[0] ? H1(B.cats[0].total) + " h" : "", "")}</div>
      <div class="grid2c"><section class="card"><h2>Time lost by cause</h2>
        <div class="legend" style="margin:0 0 8px">${B.rigs.map((r, i) => `<span><i style="background:${col[i]}"></i>${r}</span>`).join("")}</div>
        ${B.cats.map(c => `<div class="hb"><div class="hl"><span>${esc(c.cat)}</span><b>${H1(c.total)} h</b></div><div class="ht">${c.per.map((v, i) => v ? `<i style="width:${v / max * 100}%;background:${col[i]}" title="${B.rigs[i]} ${H1(v)} h"></i>` : "").join("")}</div></div>`).join("") || '<p class="muted">No breakdowns or delays written for these dates.</p>'}</section>
      <section class="card"><h2>Per machine</h2><div class="tbl"><table><thead><tr><th>Machine</th><th>Shifts working</th><th>Shifts not working</th><th>Metres</th><th>Hours lost</th></tr></thead><tbody>
        ${B.mach.map(m => `<tr><td><b>${m.rig}</b></td><td class="num">${m.working}</td><td class="num">${m.down}</td><td class="num">${M1(m.metres)}</td><td class="num">${H1(m.hours)}</td></tr>`).join("")}</tbody></table></div>
        <h2 style="margin-top:16px">By cause</h2><div class="tbl"><table><thead><tr><th>Cause</th>${B.rigs.map(r => `<th>${r} h</th>`).join("")}<th>Total h</th><th>Times</th></tr></thead><tbody>
        ${B.cats.map(c => `<tr><td>${esc(c.cat)}</td>${c.per.map(v => `<td class="num">${H1(v)}</td>`).join("")}<td class="num"><b>${H1(c.total)}</b></td><td class="num">${c.times.reduce((a, b) => a + b, 0)}</td></tr>`).join("")}
        <tr class="tot"><td>Total</td>${B.mach.map(m => `<td class="num">${H1(m.hours)}</td>`).join("")}<td class="num">${H1(B.total)}</td><td class="num">${B.lines.length}</td></tr></tbody></table></div></section></div>
      ${reports.length ? `<section class="card"><h2>Breakdown reports and job cards filed</h2><div class="dfiles">${reports.map(x => `<span class="dfile"><button class="lnk" data-open="${x.id}">${fmt(x.date)} · ${esc(x.type)} · ${esc(x.rig || "All machines")} ${esc(x.shift)}</button>${x.sp ? `<a href="${esc(x.sp)}" target="_blank" rel="noopener">OneDrive</a>` : ""}</span>`).join("")}</div></section>` : ""}
      <section class="card"><h2>Every breakdown and delay</h2><div class="tbl"><table><thead><tr><th>Date</th><th>Shift</th><th>Machine</th><th>What happened</th><th>Cause</th><th>Hours</th></tr></thead><tbody>
        ${B.lines.slice().reverse().map(l => `<tr><td style="white-space:nowrap">${fmt(l.date)}</td><td>${l.shift}</td><td style="white-space:nowrap">${l.rig}</td><td>${esc(l.text)}</td><td>${esc(l.cat)}</td><td class="num">${l.hours == null ? "" : l.hours.toFixed(2)}</td></tr>`).join("") || '<tr><td colspan="6" class="muted">Nothing written.</td></tr>'}</tbody></table></div>
        <p class="muted">Hours are as written on the shift. A delay with no time written counts in Times, not in Hours.</p></section>`;
  }

  /* ---------- Excel: the drilling metres and tally sheets workbook ---------- */
  function chartPng(title, labels, series, opt = {}) {      // horizontal bars, one group per label
    const w = 1000, rowH = Math.max(26, 14 * series.length + 12), h = 86 + labels.length * rowH + 20, k = 2, c = document.createElement("canvas"); c.width = w * k; c.height = h * k; const g = c.getContext("2d"); g.scale(k, k);
    g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.fillStyle = "#111"; g.fillRect(0, 0, w, 40); g.fillStyle = "#d4a12a"; g.fillRect(0, 40, w, 3);
    g.font = "bold 17px Arial"; g.fillStyle = "#fff"; g.textBaseline = "middle"; g.fillText(title, 16, 21);
    let x0 = 16; g.font = "12px Arial"; series.forEach(s => { g.fillStyle = s.color; g.fillRect(x0, 54, 13, 13); g.fillStyle = "#1a1a1a"; g.fillText(s.name, x0 + 19, 61); x0 += g.measureText(s.name).width + 44; });
    const left = 290, right = w - 70, top = 82, max = Math.max(1, ...series.flatMap(s => s.vals));
    labels.forEach((l, i) => { const y = top + i * rowH; g.fillStyle = "#1a1a1a"; g.font = "12px Arial"; g.textAlign = "right"; g.fillText(l.length > 42 ? l.slice(0, 41) + "…" : l, left - 10, y + rowH / 2 - 3);
      const bh = Math.min(16, (rowH - 8) / series.length);
      series.forEach((s, si) => { const v = s.vals[i] || 0, bw = (right - left) * v / max, by = y + 2 + si * bh; g.fillStyle = s.color; g.fillRect(left, by, bw, bh - 2);
        if (v) { g.fillStyle = "#1a1a1a"; g.font = "bold 10px Arial"; g.textAlign = "left"; g.fillText((opt.dec === 0 ? Math.round(v) : v.toFixed(1)) + (opt.unit || ""), left + bw + 5, by + bh / 2 - 1); } });
      g.strokeStyle = "#e6e0d2"; g.beginPath(); g.moveTo(16, y + rowH - 1); g.lineTo(w - 16, y + rowH - 1); g.stroke(); });
    return { png: c.toDataURL("image/png"), w, h };
  }
  function stackPng(days, rigs, colors, val) {
    const w = 1100, h = 430, k = 2, c = document.createElement("canvas"); c.width = w * k; c.height = h * k; const g = c.getContext("2d"); g.scale(k, k);
    g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.fillStyle = "#111"; g.fillRect(0, 0, w, 44); g.fillStyle = "#d4a12a"; g.fillRect(0, 44, w, 3);
    g.font = "bold 19px Arial"; g.fillStyle = "#fff"; g.textBaseline = "middle"; g.fillText("Metres drilled per day, by machine", 18, 23);
    let x0 = 18; g.font = "13px Arial"; rigs.forEach((r, i) => { g.fillStyle = colors[i]; g.fillRect(x0, 60, 14, 14); g.fillStyle = "#1a1a1a"; g.fillText(r, x0 + 20, 67); x0 += g.measureText(r).width + 46; });
    const left = 60, right = w - 16, top = 92, bottom = h - 40, tot = days.map(d => rigs.reduce((n, r) => n + Math.max(0, val(d, r)), 0)), max = Math.max(1, ...tot), cw = (right - left) / Math.max(1, days.length);
    g.strokeStyle = "#e6e0d2"; g.textAlign = "right"; for (let t = 0; t <= 4; t++) { const y = bottom - (bottom - top) * t / 4; g.beginPath(); g.moveTo(left, y); g.lineTo(right, y); g.stroke(); g.fillStyle = "#6b6558"; g.font = "11px Arial"; g.fillText(Math.round(max * t / 4), left - 6, y); }
    g.textAlign = "center";
    days.forEach((d, i) => { let y = bottom; const x = left + cw * i + cw * 0.15, bw = cw * 0.7;
      rigs.forEach((r, ri) => { const v = Math.max(0, val(d, r)); if (!v) return; const bh = (bottom - top) * v / max; g.fillStyle = colors[ri]; g.fillRect(x, y - bh, bw, bh); y -= bh; });
      if (tot[i] && days.length <= 31) { g.fillStyle = "#1a1a1a"; g.font = "bold 10px Arial"; g.fillText(Math.round(tot[i]), x + bw / 2, y - 8); }
      if (days.length <= 16 || i % Math.ceil(days.length / 16) === 0) { g.fillStyle = "#1a1a1a"; g.font = "11px Arial"; g.fillText(d.slice(8, 10) + "/" + d.slice(5, 7), x + bw / 2, bottom + 15); } });
    return { png: c.toDataURL("image/png"), w, h };
  }
  async function exportMetres(from, to, only) {
    toast("Making the Excel file… photos are being added, keep the app open.");
    try { await loadLib("../lib/exceljs.min.js"); } catch { toast("Open the app once with signal, then try again."); return; }
    const ExcelJS = window.ExcelJS, wb = new ExcelJS.Workbook(); wb.creator = "TKR Operations"; wb.created = new Date();
    let logo = null; try { logo = wb.addImage({ buffer: await (await fetch("../icons/icon-512.png")).arrayBuffer(), extension: "png" }); } catch {}
    const X = { BLACK: "FF111111", HEAD: "FF262626", GOLD: "FFD4A12A", CREAM: "FFFFF8E6", TOT: "FFFFF2CC", WHITE: "FFFFFFFF", GRID: "FFD9D3C5", MUTED: "FF6B6558", RED: "FF9C0006", REDF: "FFFDE2DF", GREEN: "FF375623", GREENF: "FFE2F0D9" };
    const fill = x => ({ type: "pattern", pattern: "solid", fgColor: { argb: x } }), thin = { style: "thin", color: { argb: X.GRID } }, box = { top: thin, left: thin, bottom: thin, right: thin };
    const F = (o = {}) => Object.assign({ name: "Arial", size: 10, color: { argb: X.BLACK } }, o);
    const D = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); }, DEC1 = "#,##0.0";
    const period = from === to ? fmt(from) : `${fmt(from)}  to  ${fmt(to)}`;
    function sheet(name, widths, title, sub) {
      const ws = wb.addWorksheet(name, { views: [{ showGridLines: false }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
      ws.columns = widths.map(w => ({ width: w })); const n = Math.max(5, widths.length), H = [8, 22, 30, 18];
      for (let r = 1; r <= 4; r++) { ws.getRow(r).height = H[r - 1]; for (let k = 1; k <= n; k++) ws.getCell(r, k).fill = fill(X.BLACK); }
      ws.getRow(5).height = 4; for (let k = 1; k <= n; k++) ws.getCell(5, k).fill = fill(X.GOLD);
      if (logo !== null) ws.addImage(logo, { tl: { col: 0.08, row: 1.12 }, ext: { width: 62, height: 62 } });
      const put = (r, v, f) => { const x = ws.getCell(r, 2); x.value = v; x.font = f; x.alignment = { vertical: "middle" }; };
      put(2, "TKR OPERATIONS – " + title, F({ size: 13, bold: true, color: { argb: X.WHITE } }));
      put(3, "DATES:  " + period, F({ size: 20, bold: true, color: { argb: X.GOLD } }));
      put(4, sub, F({ size: 10, bold: true, color: { argb: X.WHITE } }));
      ws.views = [{ state: "frozen", ySplit: 5, showGridLines: false }]; return ws;
    }
    const head = (ws, r, labels, h = 30) => { const row = ws.getRow(r); row.height = h; labels.forEach((l, i) => { const x = row.getCell(i + 1); x.value = l; x.font = F({ bold: true, color: { argb: X.WHITE } }); x.fill = fill(X.HEAD); x.border = box; x.alignment = { vertical: "middle", horizontal: i ? "center" : "left", wrapText: true }; }); return r + 1; };
    const line = (ws, r, vals, i, o = {}) => { const row = ws.getRow(r); row.height = o.h || 18;
      vals.forEach((v, k) => { const x = row.getCell(k + 1); x.value = v === "" || v == null ? null : v; x.border = box; x.font = F({ bold: !!o.bold || k === 0 });
        x.alignment = { vertical: o.top ? "top" : "middle", horizontal: typeof v === "number" || v instanceof Date ? "center" : "left", wrapText: true };
        if (typeof v === "number") x.numFmt = (o.fmt || {})[k] || DEC1; if (v instanceof Date) x.numFmt = "dd/mm/yyyy";
        if (o.tot) { x.fill = fill(X.TOT); x.border = { ...box, top: { style: "medium", color: { argb: X.GOLD } } }; x.font = F({ bold: true }); } else if (i % 2) x.fill = fill(X.CREAM);
        const c = (o.color || {})[k]; if (c === "g") { x.fill = fill(X.GREENF); x.font = F({ bold: true, color: { argb: X.GREEN } }); } if (c === "r") { x.fill = fill(X.REDF); x.font = F({ bold: true, color: { argb: X.RED } }); } });
      return r + 1; };
    const text = (ws, r, t, o = {}) => { const x = ws.getCell(r, 1); x.value = t; x.font = F(Object.assign({ size: 9, italic: true, color: { argb: X.MUTED } }, o)); return r + 1; };
    const rowsFor = (t, per) => Math.max(1, Math.ceil(String(t || "").length / per));
    const INT = "0";

    const S = slots(from, to), all = [...S.values()], days = daysOf(from, to), B = bdData(from, to);
    const rigs = RIGS.filter(r => all.some(s => s.rig === r)); if (!rigs.length) rigs.push(...RIGS.slice(0, 2));
    const sl = (d, sh, r) => S.get(d + "|" + sh + "|" + r), dm = (d, r) => r1(["Day", "Night"].reduce((n, sh) => n + (sl(d, sh, r)?.metres || 0), 0));
    const tot = r => r1(all.filter(s => !r || s.rig === r).reduce((n, s) => n + s.metres, 0)), grand = tot();
    // the files for these dates, photos made lighter for the workbook
    const files = [];
    if (!only) for (const d of Object.keys(idx || {}).filter(d => d >= from && d <= to).sort()) { let L = []; try { L = await daySheets(d); } catch (e) { console.warn(e); }
      for (const x of L) { const f = { ...lite(x.id, x), date: d, img: null }; const isPdf = /pdf/.test(x.mime || "");
        if (!isPdf) { try { const data = x.parts ? await fullData(x) : x.data; if (data) f.img = await shrink(data, 1100, 0.6, 400000); } catch (e) { console.warn(e); } }
        files.push(f); } }
    const tally = files.filter(f => f.type === "Tally sheet");
    const img = (ws, f, col, row, bw, bh) => { const k = Math.min(bw / f.img.w, bh / f.img.h), id = wb.addImage({ base64: f.img.data, extension: "jpeg" });
      ws.addImage(id, { tl: { col: col + 0.04, row: row + 0.04 }, ext: { width: Math.round(f.img.w * k), height: Math.round(f.img.h * k) }, editAs: "oneCell" }); };
    let ws, r, i = 0;

    if (!only) {
      // 1. Summary
      ws = sheet("Summary", [24, 16, 16, 18, 22, 14], "DRILLING SUMMARY", "Tharisa Mine · " + rigs.join(", ")); r = 7;
      ws.getCell(r, 1).value = "Total metres drilled"; ws.getCell(r, 1).font = F({ size: 12, bold: true }); ws.getCell(r, 2).value = grand; ws.getCell(r, 2).numFmt = DEC1; ws.getCell(r, 2).font = F({ size: 16, bold: true, color: { argb: "FF8A6A00" } }); ws.getRow(r).height = 26; r += 2;
      r = head(ws, r, ["Machine", "Metres", "Shifts working", "Shifts not working", "Avg m per working shift"]);
      B.mach.forEach((m, i) => { r = line(ws, r, [m.rig, m.metres, m.working, m.down, m.working ? r1(m.metres / m.working) : 0], i, { fmt: { 2: INT, 3: INT } }); });
      r = line(ws, r, ["Total", grand, "", "", ""], 0, { tot: true }); r++;
      ws.getCell(r, 1).value = "Sheets in this file:"; ws.getCell(r, 1).font = F({ bold: true }); r++;
      for (const t of ["Shifts and evidence – every shift with metres, status, delays and the documents filed for it.", "Daily metres – metres per rig per day with the tally sheets next to them.", "Machine performance – each machine per day, drilled or not, with machine hours and delays.",
        "Delay summary – time lost per machine and per cause, and every delay shift by shift.", "Supporting documents – job cards, checklists, breakdown reports and other documents filed.", "Comments – shifts with metres and no tally sheet, and notes on the documents."]) { ws.getCell(r, 1).value = t; ws.getCell(r, 1).font = F(); r++; }
      r++; const sp = stackPng(days, rigs, ["#d4a12a", "#111111", "#2e7d4f"], dm); ws.addImage(wb.addImage({ base64: sp.png, extension: "png" }), { tl: { col: 0, row: r }, ext: { width: 900, height: Math.round(900 * sp.h / sp.w) } });

      // 2. TKR vs Tharisa (only when Tharisa's figure is typed in)
      const th = num(tharisa);
      if (th > 0) { ws = sheet("TKR vs Tharisa", [32, 14, 22, 18, 10], "TKR COUNT VS THARISA COUNT", "Tharisa's metres for these dates against the TKR shift register"); r = 7;
        r = head(ws, r, ["Count", "Metres", "Difference to Tharisa", "Difference %"]);
        r = line(ws, r, ["Tharisa", r1(th), "", ""], 0); r = line(ws, r, ["TKR shift register", grand, r1(grand - th), Math.round((grand - th) / th * 1000) / 10 + "%"], 1); }

      // 3. Shifts and evidence
      ws = sheet("Shifts and evidence", [11, 7, 12, ...rigs.flatMap(() => [9, 12, 34]), 40], "MACHINE PERFORMANCE PER SHIFT", "Every day and night shift: metres, status, delays and the documents filed"); r = 7;
      r = head(ws, r, ["Date", "Shift", "Supervisor", ...rigs.flatMap(x => [x + " m", x, x + " delays"]), "Tally sheets and documents filed"]); i = 0;
      for (const d of days) for (const sh of ["Day", "Night"]) { const L = rigs.map(x => sl(d, sh, x)); if (!L.some(Boolean)) continue;
        const fl = files.filter(f => f.date === d && f.shift === sh).map(f => `${f.type} · ${f.rig || "All machines"}`).join("\n"), n = Math.max(...L.map(s => rowsFor(s?.delayText, 44)), fl ? fl.split("\n").length : 1);
        r = line(ws, r, [D(d), sh, L.map(s => s?.supervisor).find(Boolean) || "", ...L.flatMap(s => s ? [s.metres, s.status, s.delayText] : ["", "", ""]), fl], i++,
          { h: Math.max(18, n * 13 + 4), color: Object.fromEntries(L.map((s, k) => [4 + k * 3, s?.status === "Working" ? "g" : s?.status === "Not working" ? "r" : ""])) }); }
      r = line(ws, r, ["Total", "", "", ...rigs.flatMap(x => [tot(x), "", ""]), "All rigs: " + M1(grand) + " m"], 0, { tot: true });

      // 4. Daily metres with the tally sheets
      ws = sheet("Daily metres", [12, 6, 10, 10, 30, 34, 34, 34, 34], "DRILLING METRES WITH TALLY SHEETS", "Metres per day per rig with the signed tally sheets"); r = 7;
      r = head(ws, r, ["Date", "Day", "Machine", "Metres", "Tally sheets attached", "Tally sheet 1", "Tally sheet 2", "Tally sheet 3", "Tally sheet 4"]); i = 0;
      for (const d of days) for (const rg of rigs) { if (!sl(d, "Day", rg) && !sl(d, "Night", rg)) continue; const T = tally.filter(f => f.date === d && f.rig === rg), m = dm(d, rg);
        const label = T.length ? T.map(f => `${f.shift || "Whole day"} shift${f.note ? " – " + f.note : ""}`).join("\n") : m > 0 ? "No tally sheet filed" : "No drilling";
        const pics = T.filter(f => f.img).slice(0, 4); line(ws, r, [D(d), dow(d), rg, m, label, "", "", "", ""], i++, { h: pics.length ? 134 : 28, top: true });
        pics.forEach((f, k) => img(ws, f, 5 + k, r - 1, 232, 172)); r++; }
      r = line(ws, r, ["Total", "", "", grand, `All rigs, ${fmt(from)} to ${fmt(to)}`, "", "", "", ""], 0, { tot: true });
      text(ws, r, "Click a tally sheet and drag a corner to read it at full size.");

      // 5. Machine performance
      ws = sheet("Machine performance", [12, 6, ...rigs.flatMap(() => [16, 9, 20, 42])], "MACHINE PERFORMANCE REGISTER", "Drill rigs " + rigs.join(", ") + " · per day"); r = 7;
      rigs.forEach((x, k) => { const c = 3 + k * 4; ws.mergeCells(r, c, r, c + 3); const cell = ws.getCell(r, c); cell.value = x; cell.font = F({ bold: true, size: 12 }); cell.fill = fill(X.GOLD); cell.alignment = { horizontal: "center", vertical: "middle" }; }); ws.getRow(r).height = 22; r++;
      r = head(ws, r, ["Date", "Day", ...rigs.flatMap(() => ["Performance", "Metres", "Machine hours", "Breakdowns / delays"])]); i = 0;
      for (const d of days) { const cols = [], color = {}; let n = 1;
        rigs.forEach((x, k) => { const a = sl(d, "Day", x), b = sl(d, "Night", x), m = dm(d, x), any = a || b;
          const perf = !any ? "" : m > 0 ? "Drilled" : "Not drilling", hrs = [a?.hrs && "D: " + a.hrs, b?.hrs && "N: " + b.hrs].filter(Boolean).join(" | "), dl = [a?.delayText && "Day: " + a.delayText, b?.delayText && "Night: " + b.delayText].filter(Boolean).join(" | ");
          color[2 + k * 4] = perf === "Drilled" ? "g" : perf ? "r" : ""; n = Math.max(n, rowsFor(dl, 56)); cols.push(perf, any ? m : "", hrs, dl); });
        if (cols.every(v => v === "")) continue; r = line(ws, r, [D(d), dow(d), ...cols], i++, { h: Math.max(18, n * 13 + 4), color }); }
      r = line(ws, r, ["Total", "", ...rigs.flatMap(x => ["", tot(x), "", ""])], 0, { tot: true }); r++;
      text(ws, r, "Green = drilled. Red = machine not drilling (breakdown, service, waiting or standing).");
    }

    // 6. Delay summary
    ws = sheet(only ? "Breakdown report" : "Delay summary", [40, 14, 14, 14, 16, 16, 16], only ? "BREAKDOWN REPORT" : "DELAY SUMMARY", "Time lost per machine and per cause, as written on the shifts"); r = 7;
    r = head(ws, r, ["Machine", "Shifts working", "Shifts not working", "Metres", "Hours lost"]);
    B.mach.forEach((m, k) => { r = line(ws, r, [m.rig, m.working, m.down, m.metres, m.hours], k, { fmt: { 1: INT, 2: INT } }); }); r++;
    r = head(ws, r, ["Delay", ...B.rigs.map(x => x + " hours"), "Total hours", ...B.rigs.map(x => "Times recorded " + x)]);
    B.cats.forEach((c, k) => { r = line(ws, r, [c.cat, ...c.per, c.total, ...c.times], k, { fmt: Object.fromEntries(c.times.map((_, j) => [B.rigs.length + 2 + j, INT])) }); });
    r = line(ws, r, ["Total", ...B.mach.map(m => m.hours), B.total, ...B.rigs.map(x => B.lines.filter(l => l.rig === x).length)], 0, { tot: true, fmt: Object.fromEntries(B.rigs.map((_, j) => [B.rigs.length + 2 + j, INT])) }); r++;
    r = text(ws, r, "Hours are as written on the shift. A delay with no time written counts in Times recorded, not in hours."); r++;
    const col = ["#d4a12a", "#111111", "#2e7d4f"];
    if (B.cats.length) { const c1 = chartPng("Hours lost per cause, by machine", B.cats.map(c => c.cat), B.rigs.map((x, k) => ({ name: x, color: col[k], vals: B.cats.map(c => c.per[k]) })), { unit: " h" });
      ws.addImage(wb.addImage({ base64: c1.png, extension: "png" }), { tl: { col: 0, row: r }, ext: { width: 820, height: Math.round(820 * c1.h / c1.w) } }); r += Math.ceil(820 * c1.h / c1.w / 20) + 2;
      const c2 = chartPng("Shifts working and not working, by machine", B.mach.map(m => m.rig), [{ name: "Working", color: "#2e7d4f", vals: B.mach.map(m => m.working) }, { name: "Not working", color: "#b3261e", vals: B.mach.map(m => m.down) }], { dec: 0 });
      ws.addImage(wb.addImage({ base64: c2.png, extension: "png" }), { tl: { col: 0, row: r }, ext: { width: 820, height: Math.round(820 * c2.h / c2.w) } }); r += Math.ceil(820 * c2.h / c2.w / 20) + 2; }
    ws.getCell(r, 1).value = "Every delay, shift by shift"; ws.getCell(r, 1).font = F({ size: 12, bold: true }); r++;
    r = head(ws, r, ["Delay (as written on the shift)", "Date", "Shift", "Machine", "Category", "Hours"]);
    B.lines.forEach((l, k) => { r = line(ws, r, [l.text, D(l.date), l.shift, l.rig, l.cat, l.hours == null ? "" : l.hours], k, { h: Math.max(18, rowsFor(l.text, 46) * 13 + 4), fmt: { 5: "0.00" } }); });

    if (!only) {
      // 7. Supporting documents
      const sup = files.filter(f => f.type !== "Tally sheet");
      ws = sheet("Supporting documents", [34, 34, 34, 34, 20], "SUPPORTING DOCUMENTS", "Job cards, checklists, breakdown reports and other documents filed for these dates"); r = 7;
      r = head(ws, r, ["Document", "Date", "Machine and shift", "File", "Note"]);
      sup.forEach((f, k) => { r = line(ws, r, [f.type, D(f.date), `${f.rig || "All machines"}${f.shift ? " · " + f.shift : ""}`, f.sp ? { text: f.name || "Open the file", hyperlink: f.sp } : f.name || (f.pdf ? "PDF in the app" : "Photo in the app"), f.note], k); });
      if (!sup.length) r = text(ws, r, "No job cards, checklists, breakdown reports or other documents filed for these dates."); r++;
      const pics = sup.filter(f => f.img);
      for (let k = 0; k < pics.length; k += 4) { const row = pics.slice(k, k + 4); ws.getRow(r).height = 18; row.forEach((f, j) => { const x = ws.getCell(r, j + 1); x.value = `${f.type} · ${fmt(f.date)} · ${f.rig || "All machines"} ${f.shift}`; x.font = F({ bold: true, size: 9 }); }); r++;
        ws.getRow(r).height = 250; row.forEach((f, j) => img(ws, f, j, r - 1, 232, 326)); r += 2; }

      // 8. Comments
      ws = sheet("Comments", [12, 11, 10, 10, 80], "COMMENTS", "Shifts with metres and no tally sheet, and notes on the documents"); r = 7;
      r = head(ws, r, ["Date", "Machine", "Shift", "Metres", "Comments"]); i = 0;
      const com = [];
      for (const s of all) if (s.metres > 0 && !tally.some(f => f.date === s.date && f.rig === s.rig && f.shift === s.shift)) com.push([s.date, s.rig, s.shift, s.metres, `${s.shift} shift: ${M1(s.metres)} m as per the shift register; no tally sheet filed for this shift.`]);
      for (const f of files) if (f.note) com.push([f.date, f.rig || "All machines", f.shift, "", `${f.type}: ${f.note}`]);
      com.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || String(a[2]).localeCompare(String(b[2])));
      com.forEach(c => { r = line(ws, r, [D(c[0]), c[1], c[2], c[3], c[4]], i++, { h: Math.max(18, rowsFor(c[4], 95) * 13 + 4) }); });
      if (!com.length) text(ws, r, "Every shift with metres has a tally sheet. No notes.");
    }
    const buf = await wb.xlsx.writeBuffer(), blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const tag = `${fmt(from).replace(/\//g, "-")}${from === to ? "" : "_to_" + fmt(to).replace(/\//g, "-")}`;
    const name = only ? `TKR_Drilling_Breakdown_Report_${tag}.xlsx` : `TKR_Drilling_Metres_and_Tally_Sheets_${tag}.xlsx`;
    if (spReady()) { try { await spUpload(`${monthFolder(to)}/Reports`, name, blob); toast(`A copy of the report is in ${spPlace()} › ${monthFolder(to).replace("/", " › ")} › Reports.`); } catch (e) { console.warn(e); } }
    const { deliver } = await import("../lib/deliver.js"); deliver(blob, name);
    return blob;
  }

  /* ---------- events (called by the app before its own) ---------- */
  async function onClick(e) {
    const b = e.target.closest("button"); if (!b) return false;
    if (b.dataset.mnav) { month = shiftMonth(month, +b.dataset.mnav); xFrom = xTo = bdFrom = bdTo = ""; render(); return true; }
    if (b.dataset.shift) { if (b.dataset.shift.startsWith("new|")) shiftPanel(b.dataset.shift.slice(4), "Day", RIGS[0], true); else if (b.dataset.shift === "new") shiftPanel(month === today().slice(0, 7) ? today() : monthEnd(month), "Day", RIGS[0], true); else { const [d, sh, rg] = b.dataset.shift.split("|"); shiftPanel(d, sh, rg, false); } return true; }
    if (b.dataset.day) { b.disabled = true; try { await showDay(b.dataset.day); } catch (x) { console.warn(x); toast("The documents couldn't open. Check the signal."); } b.disabled = false; return true; }
    if (b.dataset.open) { try { await openFile(b.dataset.open); } catch (x) { console.warn(x); toast("The file couldn't open. Check the signal."); } return true; }
    if (b.id === "xGo" || b.id === "bdGo") { const bd = b.id === "bdGo", f = $(bd ? "bdFrom" : "xFrom").value, t = $(bd ? "bdTo" : "xTo").value; if (!f || !t || f > t) { toast("Choose a From date before the To date."); return true; }
      if (!bd) tharisa = $("xTh").value.trim(); b.disabled = true; b.textContent = "Making the Excel file…"; try { await exportMetres(f, t, bd); } catch (x) { console.warn(x); toast("The Excel file couldn't be made."); } render(); return true; }
    if (b.dataset.sp) { const a = b.dataset.sp;
      try { if (a === "in") { await spSignIn(); if (!spLinked() && isAdmin() && !spCfg.site) await linkDrive(); toast(spReady() ? `Signed in. New files now go to ${spPlace()}.` : "Signed in."); }
        if (a === "link") { await linkDrive(); toast("The TKR App folder is in your OneDrive. New files now go there."); }
        if (a === "out") { await spSignOut(); }
        if (a === "open") { window.open(await spFolderUrl(monthFolder(month + "-01")), "_blank", "noopener"); }
        if (a === "move") { if (!confirm("Send every file that is in the app and not yet in " + spPlace() + "? Keep the app open until it says Done.")) return true; b.disabled = true; await moveAllToSp(); } }
      catch (x) { console.warn(x); spState = "Microsoft said no: " + String(x?.errorMessage || x?.message || x).slice(0, 140); toast(spState); }
      render(); return true; }
    return false;
  }
  function onChange(e) {
    const id = e.target.id;
    if (id === "xFrom") { xFrom = e.target.value || xFrom; return true; } if (id === "xTo") { xTo = e.target.value || xTo; return true; } if (id === "xTh") { tharisa = e.target.value.trim(); return true; }
    if (id === "bdFrom" || id === "bdTo") { if (id === "bdFrom") bdFrom = e.target.value || bdFrom; else bdTo = e.target.value || bdTo; render(); return true; }
    return false;
  }
  async function onSubmit(e) {
    const f = e.target;
    if (f.id === "fDocs") { e.preventDefault(); const v = k => f.elements[k].value, btn = f.querySelector("button"); btn.disabled = true; upDate = v("date");
      try { await addFiles([...f.elements.files.files], { date: v("date"), type: v("type"), rig: v("rig"), shift: v("shift"), note: v("note").trim() }); month = v("date").slice(0, 7); } catch (x) { fail(x); } render(); return true; }
    if (f.id === "fSp") { e.preventDefault(); const v = k => f.elements[k].value.trim();
      try { const keep = spCfg && spCfg.tenant === v("tenant") && spCfg.clientId === v("clientId") && (spCfg.folder || "TKR App") === (v("folder") || "TKR App") && !v("site") ? { driveId: spCfg.driveId || "", rootId: spCfg.rootId || "", owner: spCfg.owner || "", webUrl: spCfg.webUrl || "" } : {};
        spCfg = { tenant: v("tenant"), clientId: v("clientId"), site: v("site").replace(/\/+$/, ""), folder: v("folder") || "TKR App", by: getEmail(), at: Date.now(), ...keep };
        await setDoc(doc(db, "fcSettings", "sharepoint"), spCfg); if (spConfigured(spCfg)) await spInit(spCfg); toast(spConfigured(spCfg) ? "Saved. Now tap Sign in with your TKR Microsoft account." : "Saved. Fill in both IDs to link OneDrive."); } catch (x) { fail(x); } render(); return true; }
    return false;
  }
  return { metresView, docsView, bdView, spCard, onClick, onChange, onSubmit, indexDay, showDay, addFile, exportMetres, _bd: bdData };
}
