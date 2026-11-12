/* TKR Drilling Metres – works with no signal.
   Everything is kept on this PC or phone (IndexedDB). Every few days: Export → the Excel (same tabs as the
   metres workbook) and a ZIP with the documents in month and shift folders, plus a backup to restore anywhere. */
import { metresChart, metresChartClick } from "../lib/metres-chart.js";
import { tabBar, pageHead, matches, icon } from "../lib/ui.js";
import { OCT_2026_LINES } from "../lib/oct2026.js";
const EMBED = new URLSearchParams(location.search).has("embed");
if (EMBED) document.body.classList.add("embed");
import { buildMetresWorkbook, bdData, daysOf } from "../lib/metres-xlsx.js";
import { spConfigured, spInit, spReady, spWho, spSignIn, spSignOut, spUpload, spLinked, spLinkOneDrive, spFolderUrl } from "../lib/sharepoint.js";
/* TKR's Microsoft 365 app registration (TKR Operations App). These two IDs are not secret. */
const MS_DEFAULT = { tenant: "7897d56e-74c6-47b7-a13d-e9335a8e9c86", clientId: "c11659ac-24c0-47b2-ba78-1dbb58a3b106", folder: "TKR App" };

const $ = id => document.getElementById(id);
const app = $("app");
const pad = n => String(n).padStart(2, "0");
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parse = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return ymd(d); };
const today = () => ymd(new Date());
const fmt = s => s ? s.split("-").reverse().join("/") : "";
const dow = s => parse(s).toLocaleDateString("en-ZA", { weekday: "short" });
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const r1 = v => Math.round((v || 0) * 10) / 10;
const M1 = v => { const [i, d] = r1(v).toFixed(1).split("."); return i.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "." + d; };
const num = v => { const n = parseFloat(String(v ?? "").replace(",", ".")); return isNaN(n) ? 0 : n; };
const hm = h => { if (h == null || h === "") return ""; const t = Math.round(h * 60); return Math.floor(t / 60) + " h " + pad(t % 60); };
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthName = m => MONTHS[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
const monthEnd = m => ymd(new Date(+m.slice(0, 4), +m.slice(5, 7), 0));
const shiftMonth = (m, n) => ymd(new Date(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)).slice(0, 7);
const when = t => { if (!t) return ""; const d = new Date(t); return `${fmt(ymd(d))} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
function toast(m) { const t = $("toast"); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3500); }

const DOC_TYPES = ["Tally sheet", "Job card", "Checklist", "Breakdown report", "Supporting document"];
const PLURAL = { "Tally sheet": "tally sheets", "Job card": "job cards", "Checklist": "checklists", "Breakdown report": "breakdown reports", "Supporting document": "supporting documents" };
const DELAY_CATS = ["Breakdowns", "Service and maintenance", "Waiting for GET / drill bits", "Stuck rods / retrieving rods", "Tramming", "Diesel / refuelling", "Blasting",
  "Waiting for block access / bench", "No operator", "Meetings, transport, safety and weather", "Ground conditions", "Other"];
const BASE_RIGS = ["TDR 013", "TDR 015", "TDR 014"];
const allRigs = () => [...new Set([...BASE_RIGS, ...(cfg.extraRigs || [])])];

/* ---------- storage on this device ---------- */
let dbp = null;
function idb() {
  if (dbp) return dbp;
  dbp = new Promise((res, rej) => { const r = indexedDB.open("tkr-metres", 1);
    r.onupgradeneeded = () => { const d = r.result; d.createObjectStore("shifts", { keyPath: "id" }); const f = d.createObjectStore("files", { keyPath: "id" }); f.createIndex("shiftId", "shiftId"); d.createObjectStore("blobs", { keyPath: "id" }); d.createObjectStore("meta", { keyPath: "k" }); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  return dbp;
}
const tx = async (stores, mode, fn) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(stores, mode); let out; Promise.resolve(fn(t)).then(v => out = v); t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error || new Error("aborted")); }); };
const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const getAll = store => tx([store], "readonly", t => req(t.objectStore(store).getAll()));
const getOne = (store, id) => tx([store], "readonly", t => req(t.objectStore(store).get(id)));
const metaGet = async k => (await getOne("meta", k))?.v;
const metaSet = (k, v) => tx(["meta"], "readwrite", t => { t.objectStore("meta").put({ k, v }); });

/* ---------- state ---------- */
let shifts = [], files = [], cfg = { rigs: ["TDR 013", "TDR 015"], device: "" }, lastExport = null, ms = { ...MS_DEFAULT }, msState = "";
let tab = "home", month = today().slice(0, 7), q = "", xFrom = "", xTo = "", tharisa = "", est = null;
const rigs = () => allRigs().filter(r => cfg.rigs.includes(r));
/* October 2026 is closed in GET Forecast: its shift register is built into the app. Bring it in once. */
async function ensureOctober() {
  if (await metaGet("oct2026")) return;
  const by = new Map(); for (const x of OCT_2026_LINES) { const k = x.date + "|" + x.shift; (by.get(k) || by.set(k, []).get(k)).push(x); }
  await tx(["shifts"], "readwrite", t => { for (const [k, L] of by) if (!shifts.some(s => s.id === k)) { const s = fromForecast(L); s.savedAt = s.editedAt = 1; s.fcAt = 9e15; s.device = "GET Forecast (October closed)"; t.objectStore("shifts").put(s); } });
  await metaSet("oct2026", 1);
}
async function load() {
  try { [shifts, files] = await Promise.all([getAll("shifts"), getAll("files")]); cfg = { ...cfg, ...(await metaGet("cfg") || {}) }; lastExport = await metaGet("lastExport") || null; ms = { ...MS_DEFAULT, ...(await metaGet("ms") || {}) }; fcLast = await metaGet("fcLast") || 0; }
  catch (e) { console.warn(e); app.innerHTML = `<div class="notice warn"><b class="t">This browser can't store the metres</b>It may be in private mode. Open the app in a normal window.</div>`; return false; }
  try { est = await navigator.storage?.estimate?.(); } catch {}
  return true;
}
const shiftOf = (d, sh) => shifts.find(s => s.id === d + "|" + sh);
const filesOf = id => files.filter(f => f.shiftId === id);
const rigOn = (s, r) => s?.rigs?.[r];
const delayText = rg => (rg?.delays || []).map(d => (d.text || d.cat) + (d.h ? ` (${hm(d.h)})` : "")).join(" · ");
const hasTally = (s, r) => files.some(f => f.shiftId === s.id && f.rig === r && f.type === "Tally sheet");
/* every rig that drilled metres on a shift needs its tally sheet attached, whatever its status says */
const needsTally = s => rigs().filter(r => { const g = rigOn(s, r); return g && num(g.metres) > 0 && !hasTally(s, r); });
const anyTally = s => files.some(f => f.shiftId === s.id && f.type === "Tally sheet");
const shiftTotal = s => r1(Object.values(s.rigs || {}).reduce((n, g) => n + num(g.metres), 0));
const sortKey = s => s.date + (s.shift === "Night" ? "2" : "1");

/* ---------- pieces ---------- */
const tile = (l, n, s, cls = "", attr = "") => `<${attr ? "button type=\"button\"" : "div"} class="tile ${cls} ${attr ? "tbtn" : ""}" ${attr}><div class="l">${l}</div><div class="n">${n}</div>${s ? `<div class="s">${s}</div>` : ""}${attr ? '<span class="go">Open ›</span>' : ""}</${attr ? "button" : "div"}>`;
const monthBar = () => `<div class="mbar"><button data-mnav="-1" aria-label="Month before">‹</button><b>${monthName(month)}</b><button data-mnav="1" aria-label="Next month" ${month >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
const chip = s => s === "Working" ? '<span class="chip ok">Working</span>' : s === "Not working" ? '<span class="chip bad">Not working</span>' : "";
const mb = b => b > 1e9 ? (b / 1e9).toFixed(1) + " GB" : Math.max(0.1, b / 1e6).toFixed(1) + " MB";
function monthStats(m) {
  const from = m + "-01", end = monthEnd(m), to = end < today() ? end : today(), L = shifts.filter(s => s.date >= from && s.date <= end);
  const expected = daysOf(from, end).length * 2, sofar = from <= to ? daysOf(from, to).length * 2 : 0;
  return { from, end, to, L, expected, sofar, metres: r1(L.reduce((n, s) => n + shiftTotal(s), 0)), rig: r => r1(L.reduce((n, s) => n + num(rigOn(s, r)?.metres), 0)),
    noTally: L.flatMap(s => needsTally(s).map(r => ({ s, r }))) };
}
const changedSinceExport = () => shifts.filter(s => !lastExport || (s.editedAt || s.savedAt) > lastExport.at).length;


/* ---------- metres chart (shared with the GET Forecast dashboard) ---------- */
const chartCard = () => metresChart(shifts.flatMap(s => rigs().map(r => ({ date: s.date, rig: r, metres: num(rigOn(s, r)?.metres) }))), rigs());

/* ---------- render ---------- */
const TABS = [["home", "Metres", "metres"], ["register", "Shift register", "shifts"], ["missing", "What's missing", "missing"], ["search", "Search", "search"], ["export", "Export & backup", "import"], ["settings", "Settings", "settings"]];
function render() {
  const n = shifts.length, used = files.reduce((a, f) => a + (f.size || 0), 0);
  const w = fbUser ? fcWaiting() : 0;
  $("store").textContent = `Kept on this device · ${n} shift${n === 1 ? "" : "s"} · ${files.length} document${files.length === 1 ? "" : "s"} (${mb(used)})` + (fbUser ? (w ? ` · ${w} waiting for GET Forecast` : " · GET Forecast up to date") : "");
  let body = "";
  if (tab === "home") body = homeView(); if (tab === "register") body = registerView(); if (tab === "missing") body = missingView();
  if (tab === "search") body = searchView(); if (tab === "export") body = exportView(); if (tab === "settings") body = settingsView();
  app.innerHTML = (EMBED ? `<p class="muted" style="margin:0 0 6px;font-size:12.5px">${esc($("store").textContent)}</p>` : "") + tabBar(TABS, tab) + body;
  if (tab === "search") { const i = $("q"); i.focus(); i.setSelectionRange(i.value.length, i.value.length); }
}
function homeView() {
  const S = monthStats(month), ch = changedSinceExport(), days = lastExport ? Math.floor((Date.now() - lastExport.at) / 864e5) : null;
  const latest = shifts.slice().sort((a, b) => sortKey(b).localeCompare(sortKey(a))).slice(0, 4);
  return pageHead("Drilling metres · works with no signal", "Shift capture", monthName(month), `<button class="primary" data-cap="">+ Capture a shift</button>`,
      "Each shift is saved on this device straight away. Export every few days to take the Excel and the documents off this device.")
    + `<div class="mhead">${monthBar()}</div>`
    + `<div class="tiles">${tile("Month total", M1(S.metres) + " m", `${fmt(S.from)} to ${fmt(S.to)}`, "dark")}${rigs().map(r => tile(r, M1(S.rig(r)) + " m", `${S.L.filter(s => rigOn(s, r)?.status === "Working").length} shifts working · ${S.L.filter(s => rigOn(s, r)?.status === "Not working").length} not working`)).join("")}
      ${tile("Shifts captured", `${S.L.length} <small style="font-size:16px;color:var(--muted)">of ${S.sofar}</small>`, `${S.expected} shifts in ${MONTHS[+month.slice(5, 7) - 1]} (day and night)`, S.L.length < S.sofar ? "alert" : "good", 'data-tab="missing"')}
      ${tile("Shifts without a tally sheet", String(S.noTally.length), S.noTally.length ? "tap to see which" : "every shift with metres has one", S.noTally.length ? "alert" : "good", 'data-tab="missing"')}
      ${tile("Not exported yet", String(ch), lastExport ? `last export ${when(lastExport.at)}` : "nothing exported yet", ch && (days === null || days >= 3) ? "alert" : "", 'data-tab="export"')}</div>`
    + chartCard()
    + `<div class="sect">Last shifts captured</div><div class="res">${latest.map(shiftLine).join("") || '<div class="card muted">Nothing captured yet. Tap + Capture a shift.</div>'}</div>`;
}
function shiftLine(s) {
  const nt = needsTally(s), nf = filesOf(s.id).length;
  return `<div class="ri"><div class="rt"><b>${dow(s.date)} ${fmt(s.date)} · ${s.shift} shift${s.supervisor ? " · " + esc(s.supervisor) : ""}</b>
    <small>${rigs().map(r => { const g = rigOn(s, r); return g ? `${r}: ${M1(num(g.metres))} m ${g.status ? "(" + g.status.toLowerCase() + ")" : ""}` : ""; }).filter(Boolean).join(" · ") || "No machines filled in"} · ${nf} document${nf === 1 ? "" : "s"}${nt.length ? ` · <span style="color:var(--bad)">no tally sheet: ${nt.join(", ")}</span>` : ""}</small></div>
    <div class="row" style="flex:none"><button class="sm" data-cap="${s.id}">Edit</button><button class="sm danger" data-del="${s.id}">Delete</button></div></div>`;
}
function registerView() {
  const S = monthStats(month), R = rigs(), L = S.L.slice().sort((a, b) => sortKey(b).localeCompare(sortKey(a)));
  return pageHead("Shift register", monthName(month), "", `<button class="primary" data-cap="">+ Capture a shift</button>`, "Like the Shifts and evidence sheet. Edit to change a shift, Delete to remove it and capture it again.")
    + `<div class="mhead">${monthBar()}<span class="muted">${S.L.length} shifts · ${M1(S.metres)} m</span></div>`
    + `<div class="tbl card" style="padding:0"><table class="reg"><thead><tr><th>Date</th><th>Shift</th><th>Supervisor</th>${R.map(r => `<th>${r} m</th><th>${r}</th><th>${r} delays</th>`).join("")}<th>Shift report</th><th>Documents</th><th></th></tr></thead><tbody>
      ${L.map((s, i) => `<tr class="${i && L[i - 1].date !== s.date ? "day1" : ""}"><td>${dow(s.date)} ${fmt(s.date)}</td><td>${s.shift}</td><td>${esc(s.supervisor)}</td>
        ${R.map(r => { const g = rigOn(s, r); return g ? `<td class="num">${M1(num(g.metres))}</td><td class="${g.status === "Working" ? "ok" : g.status === "Not working" ? "bad" : ""}">${esc(g.status)}</td><td class="wrap">${esc(delayText(g))}</td>` : "<td></td><td></td><td></td>"; }).join("")}
        <td class="wrap">${esc(s.report || "")}</td><td class="wrap">${DOC_TYPES.map(t => { const n = filesOf(s.id).filter(f => f.type === t).length; return n ? `${n} ${n === 1 ? t.toLowerCase() : PLURAL[t]}` : ""; }).filter(Boolean).join(", ") || '<span style="color:var(--bad)">none</span>'}</td>
        <td><div class="act"><button data-cap="${s.id}">Edit</button><button class="danger" data-del="${s.id}">Delete</button></div></td></tr>`).join("")
      || `<tr><td colspan="${6 + R.length * 3}" class="muted">No shifts captured in ${monthName(month)}.</td></tr>`}
      ${L.length ? `<tr class="tot"><td>Total</td><td></td><td></td>${R.map(r => `<td class="num">${M1(S.rig(r))}</td><td></td><td></td>`).join("")}<td>All rigs: ${M1(S.metres)} m</td><td></td><td></td></tr>` : ""}</tbody></table></div>`;
}
function missingView() {
  const S = monthStats(month), days = daysOf(S.from, S.end), t = today();
  const state = (d, sh) => { if (d > t) return "fut"; const s = shiftOf(d, sh); if (!s) return "miss"; return anyTally(s) && !needsTally(s).length ? "ok" : "part"; };
  const cells = days.flatMap(d => ["Day", "Night"].map(sh => ({ d, sh, st: state(d, sh) })));
  const notCap = cells.filter(c => c.st === "miss"), noTally = S.noTally;
  const noDocs = S.L.filter(s => rigs().some(r => /^Working/.test(rigOn(s, r)?.status || "") || num(rigOn(s, r)?.metres) > 0)).map(s => ({ s, miss: ["Job card", "Checklist"].filter(t => !filesOf(s.id).some(f => f.type === t)) })).filter(x => x.miss.length);
  return pageHead("What's missing", monthName(month), "", "", `${S.expected} shifts this month (day and night, every day). Tap a box to capture or open that shift.`)
    + `<div class="mhead">${monthBar()}</div>`
    + `<div class="tiles">${tile("Shifts captured", `${S.L.length} of ${S.sofar}`, `so far · ${S.expected} in the month`, notCap.length ? "alert" : "good")}${tile("Not captured", String(notCap.length), "up to today", notCap.length ? "alert" : "good")}${tile("Without a tally sheet", String(noTally.length), "shifts with metres", noTally.length ? "alert" : "good")}${tile("Without a job card or checklist", String(noDocs.length), "working shifts", noDocs.length ? "alert" : "good")}</div>`
    + `<div class="legend2"><span><i style="background:var(--ok)"></i>Captured, tally sheet attached for every rig that drilled</span><span><i style="background:var(--gold)"></i>Captured, no tally sheet attached</span><span><i style="background:var(--bad)"></i>Not captured</span></div>`
    + `<div class="mgrid">${cells.map(c => `<button class="mcell ${c.st}" data-cap="${c.d}|${c.sh}" ${c.st === "fut" ? "disabled" : ""}><b>${c.d.slice(8)}</b>${c.sh}</button>`).join("")}</div>`
    + `<div class="grid2c"><section class="card"><h2>Not captured (${notCap.length})</h2><div class="res miss-list">${notCap.slice().reverse().map(c => `<div class="ri"><div class="rt"><b>${dow(c.d)} ${fmt(c.d)} · ${c.sh} shift</b></div><button class="sm primary" data-cap="${c.d}|${c.sh}">Capture</button></div>`).join("") || '<p class="muted">Every shift up to today is captured.</p>'}</div></section>
      <section class="card"><h2>No tally sheet (${noTally.length})</h2><div class="res">${noTally.map(({ s, r }) => `<div class="ri"><div class="rt"><b>${dow(s.date)} ${fmt(s.date)} · ${s.shift} · ${r}</b><small>${M1(num(rigOn(s, r).metres))} m</small></div><button class="sm" data-cap="${s.id}">Add tally sheet</button></div>`).join("") || '<p class="muted">Every shift with metres has a tally sheet.</p>'}</div>
      <h2 style="margin-top:16px">No job card or checklist (${noDocs.length})</h2><div class="res">${noDocs.map(({ s, miss }) => `<div class="ri"><div class="rt"><b>${dow(s.date)} ${fmt(s.date)} · ${s.shift}</b><small>missing: ${miss.join(", ").toLowerCase()}</small></div><button class="sm" data-cap="${s.id}">Add</button></div>`).join("") || '<p class="muted">Every working shift has a job card and a checklist.</p>'}</div></section></div>`;
}
function searchView() {
  const hits = !q.trim() ? [] : [
    ...shifts.filter(s => matches([s.date, fmt(s.date), dow(s.date), s.shift, s.supervisor, s.report, ...Object.entries(s.rigs || {}).flatMap(([r, g]) => [r, g.status, g.block, g.metres, delayText(g), ...(g.delays || []).map(d => d.cat)])].join(" "), q))
      .sort((a, b) => sortKey(b).localeCompare(sortKey(a))).map(s => ({ k: "Shift", html: shiftLine(s) })),
    ...files.filter(f => matches([f.date, fmt(f.date), f.shift, f.rig, f.type, f.name, f.note].join(" "), q)).sort((a, b) => b.date.localeCompare(a.date))
      .map(f => ({ k: f.type, html: `<div class="ri"><div class="rt"><b>${esc(f.type)} · ${fmt(f.date)} ${esc(f.shift)} · ${esc(f.rig)}</b><small>${esc(f.name)} · ${mb(f.size || 0)}</small></div><div class="row" style="flex:none"><button class="sm" data-view="${f.id}">View</button><button class="sm" data-cap="${f.shiftId}">Open the shift</button></div></div>` }))];
  return pageHead("Search", "Find a shift or a document", "", "", "Type a date (10/11), a rig (013), a supervisor, a delay (hydraulic, diesel), a block or a document (tally, job card).")
    + `<div class="searchbox"><input id="q" type="search" placeholder="Search…" value="${esc(q)}" autocomplete="off"></div>`
    + (q.trim() ? `<p class="muted">${hits.length} found</p><div class="res">${hits.map(h => h.html).join("")}</div>` : "");
}
function exportView() {
  const first = shifts.map(s => s.date).sort()[0] || today();
  xFrom = xFrom || (lastExport ? (lastExport.to < today() ? addDays(lastExport.to, 1) : today()) : month + "-01"); if (xFrom < first && !lastExport) xFrom = first; xTo = xTo || today();
  const used = files.reduce((a, f) => a + (f.size || 0), 0), q2 = est?.quota || 0;
  return pageHead("Export and backup", "Take the metres off this device", "", "", "Do this every few days. The ZIP has the Excel, every document in month and shift folders (November 2026 Metres › 01 DAY), and a backup that can be loaded on any other PC or phone.")
    + fcCard() + `<section class="card dl-card"><h2>Export</h2><div class="rangebar"><label>From <input type="date" id="xFrom" value="${xFrom}" max="${today()}"></label><label>To <input type="date" id="xTo" value="${xTo}" max="${today()}"></label>
        <label style="flex:1 1 200px">Tharisa's metres for these dates (if you have it) <input id="xTh" inputmode="decimal" value="${esc(tharisa)}" placeholder="optional"></label></div>
      <div class="openers" style="margin:0"><button class="opener" data-x="zip"><span class="oi">${icon("import")}</span><span class="ot"><b>Download everything (ZIP)</b><small>Excel + all documents in folders + backup</small></span><span class="oa">›</span></button>
      <button class="opener" data-x="xlsx"><span class="oi">${icon("docs")}</span><span class="ot"><b>Download the Excel only</b><small>Same tabs as the metres workbook, tally sheets inside</small></span><span class="oa">›</span></button></div>
      ${ms.driveId ? `<div class="openers" style="margin:10px 0 0"><button class="opener good" data-x="od"><span class="oi">${icon("import")}</span><span class="ot"><b>Send to OneDrive</b><small>${files.filter(f => !f.odAt).length} documents not sent yet · needs signal</small></span><span class="oa">›</span></button></div>` : ""}
      <p class="muted" id="xMsg" style="margin:10px 0 0">${lastExport ? `Last export: <b>${when(lastExport.at)}</b> (${fmt(lastExport.from)} to ${fmt(lastExport.to)}) · ${changedSinceExport()} shifts captured or changed since then.` : "Nothing exported from this device yet."}</p></section>
    <section class="card"><h2>Load a backup</h2><p class="muted" style="margin-top:0">Choose a ZIP exported from this page (from this or another device). Shifts and documents not on this device are added; a shift saved on both keeps the newer one.</p>
      <input type="file" id="restore" accept=".zip"><p class="muted" id="rMsg"></p></section>
    <section class="card"><h2>Space on this device</h2><p style="margin:0">${mb(used)} of documents${q2 ? ` · the browser allows up to about ${mb(q2)}` : ""}.</p>${q2 ? `<div class="sizebar"><i style="width:${Math.min(100, (est.usage || used) / q2 * 100).toFixed(1)}%"></i></div>` : ""}</section>`;
}
function settingsView() {
  return pageHead("Settings", "This device", "", "", "")
    + `<section class="card"><h2>Drill rigs on the shift form</h2><div class="seg" style="max-width:520px">${allRigs().map(r => `<label class="ok"><input type="checkbox" name="rig" value="${r}" ${cfg.rigs.includes(r) ? "checked" : ""}>${r}</label>`).join("")}</div>
      <div class="row" style="margin-top:12px;max-width:520px"><input id="newRig" placeholder="New rig, e.g. TDR 016" style="flex:1 1 200px"><button id="addRig">+ Add rig</button></div>
      <p class="muted" style="margin:6px 0 0">Ticked rigs show on the shift form, the register, the chart and the Excel.</p>
      <label style="display:block;margin-top:14px;max-width:420px">Name of this device (shows on exports) <input id="devName" value="${esc(cfg.device)}" placeholder="e.g. Control room PC"></label>
      <button class="primary" id="cfgSave" style="margin-top:12px">Save</button></section>
    <section class="card" id="msCard"><h2>OneDrive</h2>
      <p class="muted" style="margin-top:0">When TKR's Microsoft 365 admin has approved the TKR Operations App, sign in once on this device. After that, <b>Send to OneDrive</b> on the Export tab puts the documents in <b>TKR App › Drilling › November 2026 Metres › 01 DAY, 01 NIGHT…</b> and the Excel in the month folder. It needs signal only while sending.</p>
      <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr))"><label>Directory (tenant) ID <input id="msTenant" value="${esc(ms.tenant)}"></label><label>Application (client) ID <input id="msClient" value="${esc(ms.clientId)}"></label><label>Main folder <input id="msFolder" value="${esc(ms.folder || "TKR App")}"></label></div>
      <div class="row" style="margin-top:12px"><button id="msSave">Save</button>${spReady() ? `<button class="primary" data-ms="open">Open the folder</button><button class="ghost" data-ms="out">Sign out of Microsoft</button>` : `<button class="primary" data-ms="in">Sign in with the TKR Microsoft account</button>`}</div>
      <p class="muted" style="margin-bottom:0">${spReady() && spLinked() ? `Linked: signed in as <b>${esc(spWho())}</b>${ms.owner ? `, folder in ${esc(ms.owner)}'s OneDrive` : ""}.` : navigator.onLine ? "Not linked yet." : "No signal now. Linking needs signal once."} ${esc(msState)}</p></section>
    <section class="card"><h2>Good to know</h2><ul style="margin:0;padding-left:18px;line-height:1.7">
      <li>No signal is needed. Open this page once with signal so it is kept on the device, then it works offline.</li>
      <li>The shifts stay in this browser on this device only. Don't clear the browser data on this PC.</li>
      <li>Export every few days, and keep the ZIP on a flash drive, OneDrive or email it.</li></ul></section>`;
}

/* ---------- capture / edit a shift ---------- */
let form = null;   // { id, isNew, date, shift, keep: Set(fileIds kept), drop: Set }
const delayRow = (d = {}) => { const t = d.h == null || d.h === "" ? "" : Math.round(+d.h * 60), ro = d.ctrl ? ' readonly title="From the Control Room"' : "";
  return `<div class="drow${d.ctrl ? " ctrl" : ""}" ${d.ctrl ? `data-ctrl="${esc(d.ctrl)}"` : ""}><select data-k="cat" aria-label="Type of delay">${DELAY_CATS.map(c => `<option ${c === (d.cat || "Breakdowns") ? "selected" : ""}>${c}</option>`).join("")}</select>
    <input data-k="hh" inputmode="numeric" placeholder="h" value="${t === "" ? "" : Math.floor(t / 60)}" aria-label="Hours"${ro}><input data-k="mm" inputmode="numeric" placeholder="min" value="${t === "" ? "" : t % 60}" aria-label="Minutes"${ro}>
    <input data-k="text" placeholder="What happened" value="${esc(d.text || "")}">${d.ctrl ? '<span class="ctag" title="Hours and minutes come from the Control Room">CR</span>' : '<button type="button" class="danger" data-drm aria-label="Remove this delay">×</button>'}</div>`; };
async function fileChips(list) {
  return list.map(f => `<span class="fchip" data-fid="${f.id}">${f.thumb ? `<img src="${f.thumb}" alt="">` : `<span class="pdf">${/pdf/.test(f.mime) ? "PDF" : "PHOTO"}</span>`}<span>${esc(f.type)}</span><button type="button" data-view="${f.id}">View</button><button type="button" class="danger" data-frm="${f.id}">Remove</button></span>`).join("");
}
async function openForm(id, preset) {
  let s = id ? shifts.find(x => x.id === id) : null;
  let date = s?.date || preset?.date || today(), sh = s?.shift || preset?.shift || (new Date().getHours() >= 17 || new Date().getHours() < 5 ? "Night" : "Day");
  if (!s) s = shiftOf(date, sh) || null;
  form = { id: s?.id || null, date, shift: sh, drop: new Set() };
  const sups = [...new Set(shifts.map(x => x.supervisor).filter(Boolean))].sort();
  const F = s ? filesOf(s.id) : [];
  const rigCard = async r => { const g0 = rigOn(s, r) || {}, st = g0.status || "", g = s ? g0 : { ...g0, delays: ctrlDelays(date, sh, r) };
    return `<div class="rigcard" data-rig="${r}"><h3>${r}<span class="muted" style="font-size:13px;font-weight:500">${st ? "" : "leave empty if not used"}</span></h3>
      <div class="seg" style="margin-bottom:10px">${[["Working", "ok"], ["Not working", "bad"], ["", ""]].map(([v, c]) => `<label class="${c}"><input type="radio" name="st-${r}" value="${v}" ${st === v ? "checked" : ""}>${v || "Not used"}</label>`).join("")}</div>
      <div class="grid"><label>Metres drilled <input data-f="metres" inputmode="decimal" value="${g.metres ?? ""}"></label><label>Block <input data-f="block" value="${esc(g.block || "")}"></label>
        <label>Machine hours start <input data-f="hrsFrom" inputmode="decimal" value="${esc(g.hrsFrom || "")}"></label><label>Machine hours end <input data-f="hrsTo" inputmode="decimal" value="${esc(g.hrsTo || "")}"></label></div>
      <div class="sect" style="margin:14px 0 6px">Breakdowns and delays <span class="muted" style="font-weight:500;text-transform:none;letter-spacing:0">· CR = hours and minutes from the Control Room</span></div><div data-delays>${(g.delays || []).map(delayRow).join("")}</div><button type="button" class="sm" data-dadd>+ Add a breakdown or delay</button>
      <div class="sect" style="margin:14px 0 6px">Documents</div><div class="flist">${await fileChips(F.filter(f => f.rig === r))}</div>
      <div class="docs">${DOC_TYPES.map(t => `<label class="pick ${t === "Tally sheet" ? "req" : ""}">${t}<input type="file" data-type="${t}" accept="image/*,application/pdf" multiple><em class="muted" style="font-style:normal"></em></label>`).join("")}</div></div>`; };
  const cards = (await Promise.all(rigs().map(rigCard))).join("");
  document.getElementById("capForm")?.remove();
  const ov = document.createElement("div"); ov.id = "capForm"; ov.className = "sheet"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
  ov.innerHTML = `<div class="sheet-bar"><button type="button" class="back" data-close>‹ Back</button><b>${s ? "Edit shift" : "Capture a shift"}</b></div>
    <div class="sheet-body"><form id="fShift" autocomplete="off">
      ${s ? `<div class="info">You are changing the <b>${dow(s.date)} ${fmt(s.date)} ${s.shift}</b> shift, saved ${when(s.editedAt || s.savedAt)}.</div>` : ""}
      <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr))">
        <label>Date <input type="date" name="date" value="${date}" max="${today()}" required></label>
        <div><span class="muted" style="font-size:12px;letter-spacing:1px;text-transform:uppercase;font-weight:600">Shift</span><div class="seg" style="margin-top:6px">${["Day", "Night"].map(v => `<label><input type="radio" name="shift" value="${v}" ${v === sh ? "checked" : ""}>${v}</label>`).join("")}</div></div>
        <label>Supervisor <input name="supervisor" list="supList" value="${esc(s?.supervisor || "")}"><datalist id="supList">${sups.map(x => `<option value="${esc(x)}">`).join("")}</datalist></label>
      </div>
      <div class="rigcards">${cards}</div>
      <label>Shift report (as sent on WhatsApp, optional) <textarea name="report" rows="4" style="width:100%">${esc(s?.report || "")}</textarea></label>
      <div class="formbar"><button class="primary" type="submit">${s ? "Save changes" : "Save shift"}</button>${s ? `<button type="button" class="danger" data-del="${s.id}">Delete this shift</button>` : ""}<button type="button" class="ghost" data-close>Cancel</button></div>
    </form></div>`;
  document.body.appendChild(ov); document.body.classList.add("noscroll");
  ov.querySelector(".sheet-body").scrollTop = 0;
  ov.addEventListener("click", formClick); ov.addEventListener("change", formChange); ov.querySelector("form").addEventListener("submit", saveForm);
}
function closeForm() { document.getElementById("capForm")?.remove(); document.body.classList.remove("noscroll"); form = null; }
async function formClick(e) {
  const b = e.target.closest("button"); if (!b) return;
  if (b.hasAttribute("data-close")) { closeForm(); return; }
  if (b.hasAttribute("data-dadd")) { b.previousElementSibling.insertAdjacentHTML("beforeend", delayRow()); return; }
  if (b.hasAttribute("data-drm")) { b.parentElement.remove(); return; }
  if (b.dataset.frm) { const c = b.closest(".fchip"); const id = b.dataset.frm; if (form.drop.has(id)) { form.drop.delete(id); c.classList.remove("gone"); b.textContent = "Remove"; } else { form.drop.add(id); c.classList.add("gone"); b.textContent = "Keep"; } return; }
  if (b.dataset.view) { viewFile(b.dataset.view); return; }
  if (b.dataset.del) { askDelete(b.dataset.del); return; }
}
function formChange(e) {
  const t = e.target;
  if (t.type === "file") { const n = t.files.length; t.parentElement.querySelector("em").textContent = n ? `${n} chosen` : ""; t.closest(".pick").style.borderColor = n ? "var(--ok)" : ""; return; }
  if ((t.name === "date" || t.name === "shift") && !form.id) {     // that shift is already saved: open it to edit instead of saving it twice
    const f = t.form, d = f.elements.date.value, sh = f.elements.shift.value, ex = shiftOf(d, sh);
    if (ex) { toast(`${dow(d)} ${fmt(d)} ${sh} is already captured. Opening it.`); openForm(ex.id); }
  }
}
const dataUrl = b => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(b); });
async function shrink(blob, max, q = 0.7) {
  const url = URL.createObjectURL(blob);
  try { const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    const k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement("canvas"); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
    const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
    return { data: c.toDataURL("image/jpeg", q), w: c.width, h: c.height }; }
  finally { URL.revokeObjectURL(url); }
}
async function saveForm(e) {
  e.preventDefault(); const f = e.target, btn = f.querySelector("button[type=submit]");
  const date = f.elements.date.value, sh = f.elements.shift.value;
  if (!date) { toast("Choose the date."); return; }
  const id = date + "|" + sh, other = shiftOf(date, sh);
  if (other && other.id !== form.id) { toast(`${fmt(date)} ${sh} is already captured. Open it from the register to change it.`); return; }
  const rigData = {}, picks = [];
  for (const card of f.querySelectorAll(".rigcard")) {
    const r = card.dataset.rig, v = k => card.querySelector(`[data-f=${k}]`).value.trim(), st = card.querySelector(`input[name="st-${r}"]:checked`)?.value || "";
    const delays = [...card.querySelectorAll(".drow")].map(x => { const g = k => x.querySelector(`[data-k=${k}]`).value.trim(), hh = g("hh"), mm = g("mm");
      return { cat: g("cat"), h: hh === "" && mm === "" ? null : Math.round((num(hh) + num(mm) / 60) * 100) / 100, text: g("text"), ...(x.dataset.ctrl ? { ctrl: x.dataset.ctrl } : {}) }; });
    const rp = [...card.querySelectorAll("input[type=file]")].flatMap(i => [...i.files].map(file => ({ file, type: i.dataset.type, rig: r })));
    picks.push(...rp);
    const keepFiles = form.id ? filesOf(form.id).filter(x => x.rig === r && !form.drop.has(x.id)).length : 0;
    if (!st && v("metres") === "" && !delays.length && !rp.length && !keepFiles && !v("block")) continue;
    rigData[r] = { status: st || (num(v("metres")) > 0 ? "Working" : ""), metres: v("metres") === "" ? "" : num(v("metres")), block: v("block"), hrsFrom: v("hrsFrom"), hrsTo: v("hrsTo"), delays };
  }
  if (!Object.keys(rigData).length && !picks.length) { toast("Fill in at least one drill rig."); return; }
  btn.disabled = true; btn.textContent = "Saving…";
  try {
    const now = Date.now(), old = form.id ? shifts.find(x => x.id === form.id) : null;
    const rec = { id, date, shift: sh, supervisor: f.elements.supervisor.value.trim(), report: f.elements.report.value.trim(), rigs: rigData, savedAt: old?.savedAt || now, editedAt: now, device: cfg.device || "" };
    // files: make the small picture for the screen, keep the original as it was taken
    const newFiles = [];
    for (const p of picks) { const isPdf = p.file.type === "application/pdf" || /\.pdf$/i.test(p.file.name); let thumb = "";
      if (!isPdf) { try { thumb = (await shrink(p.file, 160, 0.6)).data; } catch {} }
      newFiles.push({ meta: { id: "f" + now.toString(36) + Math.random().toString(36).slice(2, 7), shiftId: id, date, shift: sh, rig: p.rig, type: p.type, name: p.file.name, mime: p.file.type || (isPdf ? "application/pdf" : "image/jpeg"), size: p.file.size, thumb, note: "", savedAt: now }, blob: p.file }); }
    const moved = old && old.id !== id ? filesOf(old.id).filter(x => !form.drop.has(x.id)) : [];
    await tx(["shifts", "files", "blobs"], "readwrite", t => {
      if (old && old.id !== id) t.objectStore("shifts").delete(old.id);
      t.objectStore("shifts").put(rec);
      for (const x of form.drop) { t.objectStore("files").delete(x); t.objectStore("blobs").delete(x); }
      for (const x of moved) t.objectStore("files").put({ ...x, shiftId: id, date, shift: sh });
      for (const n of newFiles) { t.objectStore("files").put(n.meta); t.objectStore("blobs").put({ id: n.meta.id, blob: n.blob }); }
    });
    try { navigator.storage?.persist?.(); } catch {}
    const removed = form.drop.size;
    const gone = files.filter(x => form.drop.has(x.id) && x.fcId).map(x => ({ fcId: x.fcId, date: x.date }));
    if (old && old.id !== id && old.fcAt) await metaSet("fcDel", [...((await metaGet("fcDel")) || []), { date: old.date, shift: old.shift }]);
    if (gone.length) await metaSet("fcGone", [...((await metaGet("fcGone")) || []), ...gone]);
    await load(); closeForm(); render();
    savedPopup(rec, newFiles.map(n => n.meta), removed, !!old);
    if (fbUser && navigator.onLine) fcSync(false);
  } catch (x) { console.warn(x); btn.disabled = false; btn.textContent = "Save shift"; toast("It didn't save: " + (x?.name === "QuotaExceededError" ? "this device is full. Export and free up space." : "try again.")); }
}
function popup(cls, html, on) {
  document.querySelector(".done")?.remove();
  const d = document.createElement("div"); d.className = "done " + cls; d.setAttribute("role", "alertdialog"); d.innerHTML = `<div>${html}</div>`;
  d.addEventListener("click", e => { const b = e.target.closest("button"); if (!b && e.target !== d) return; d.remove(); on?.(b?.dataset.a); });
  document.body.appendChild(d); d.querySelector("button")?.focus();
}
function savedPopup(s, added, removed, edited) {
  const nt = needsTally(s), counts = DOC_TYPES.map(t => [t, added.filter(f => f.type === t).length]).filter(x => x[1]);
  const next = s.shift === "Day" ? { date: s.date, shift: "Night" } : { date: addDays(s.date, 1), shift: "Day" };
  const nextOk = next.date <= today() && !shiftOf(next.date, next.shift);
  popup("", `<div class="big"><i>✓</i>${added.length ? "Saved, uploaded and stored" : "Saved and stored"}</div>
    <ul><li><b>${dow(s.date)} ${fmt(s.date)} · ${s.shift} shift</b>${s.supervisor ? " · " + esc(s.supervisor) : ""}${edited ? " (changed)" : ""}</li>
      ${Object.entries(s.rigs).map(([r, g]) => `<li>${r}: <b>${M1(num(g.metres))} m</b>${g.status ? " · " + g.status.toLowerCase() : ""}${g.delays.length ? ` · ${g.delays.length} delay${g.delays.length === 1 ? "" : "s"}` : ""}</li>`).join("")}
      <li>${counts.length ? "Uploaded: " + counts.map(([t, n]) => `${n} ${n === 1 ? t.toLowerCase() : PLURAL[t]}`).join(", ") : "No new documents"}${removed ? ` · ${removed} removed` : ""}</li>
      <li>Stored on this device. ${filesOf(s.id).length} document${filesOf(s.id).length === 1 ? "" : "s"} on this shift.</li>
      ${nt.length ? `<li style="color:var(--bad)">No tally sheet yet for ${nt.join(", ")}.</li>` : ""}</ul>
    <div class="acts">${nextOk ? `<button class="primary" data-a="next">Capture the next shift (${dow(next.date)} ${next.shift})</button>` : ""}<button data-a="reg">See it in the shift register</button><button data-a="ok">Done</button></div>`,
    a => { if (a === "next") openForm(null, next); if (a === "reg") { tab = "register"; month = s.date.slice(0, 7); render(); } });
}
function askDelete(id) {
  const s = shifts.find(x => x.id === id); if (!s) return; const F = filesOf(id);
  popup("del", `<div class="big"><i>!</i>Delete this shift?</div>
    <ul><li><b>${dow(s.date)} ${fmt(s.date)} · ${s.shift} shift</b></li>${Object.entries(s.rigs).map(([r, g]) => `<li>${r}: ${M1(num(g.metres))} m</li>`).join("")}<li>${F.length} document${F.length === 1 ? "" : "s"} are deleted with it</li></ul>
    <p class="muted" style="margin-top:0">It is removed from this device. You can capture it again straight after.</p>
    <div class="acts"><button class="danger" data-a="yes" style="border-color:var(--bad)">Yes, delete it</button><button data-a="no">Keep it</button></div>`,
    async a => { if (a !== "yes") return;
      await tx(["shifts", "files", "blobs"], "readwrite", t => { t.objectStore("shifts").delete(id); for (const f of F) { t.objectStore("files").delete(f.id); t.objectStore("blobs").delete(f.id); } });
      if (s.fcAt) await metaSet("fcDel", [...((await metaGet("fcDel")) || []), { date: s.date, shift: s.shift }]);
      const gone = F.filter(f => f.fcId).map(f => ({ fcId: f.fcId, date: f.date })); if (gone.length) await metaSet("fcGone", [...((await metaGet("fcGone")) || []), ...gone]);
      if (fbUser && navigator.onLine) fcSync(false);
      await load(); closeForm(); render();
      popup("del", `<div class="big"><i>✓</i>Deleted</div><p class="muted">${dow(s.date)} ${fmt(s.date)} ${s.shift} shift is removed.</p><div class="acts"><button class="primary" data-a="redo">Capture it again</button><button data-a="ok">Done</button></div>`,
        b => { if (b === "redo") openForm(null, { date: s.date, shift: s.shift }); }); });
}
async function blobOf(f) {
  const have = (await getOne("blobs", f.id))?.blob; if (have || !f.remote) return have || null;
  if (!navigator.onLine) return null;
  await fire(); const { db, doc, getDoc, getDocs, query, where, collection } = fb;
  const v = (await getDoc(doc(db, "fcSettings", f.fcId))).data(); if (!v) return null;
  let data = v.data || "";
  if (v.parts) { const s = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "part"), where("parent", "==", f.fcId))); data = s.docs.map(d => d.data()).sort((a, b) => a.i - b.i).map(p => p.data).join(""); }
  if (!data) return null;
  const b = await (await fetch(data)).blob(); let thumb = f.thumb; if (!thumb && !/pdf/.test(f.mime)) { try { thumb = (await shrink(b, 160, 0.6)).data; } catch {} }
  await tx(["files", "blobs"], "readwrite", t => { t.objectStore("blobs").put({ id: f.id, blob: b }); t.objectStore("files").put({ ...f, thumb, size: b.size }); });
  f.thumb = thumb; f.size = b.size; return b;
}
async function viewFile(id) {
  const m = files.find(f => f.id === id); if (!m) return; toast("Opening…");
  let b = null; try { b = await blobOf(m); } catch (e) { console.warn(e); }
  if (!b) { toast(m.remote ? "This photo is in GET Forecast. Open it once with signal to keep it on this device." : "That document isn't on this device."); return; }
  const url = URL.createObjectURL(b);
  if (/pdf/.test(m?.mime || "")) { window.open(url, "_blank"); return; }
  const ov = document.createElement("div"); ov.className = "photo"; ov.innerHTML = `<div class="bar"><b>${esc(m.type)} · ${fmt(m.date)} ${esc(m.shift)} · ${esc(m.rig)}</b><button class="primary" data-x>Close</button></div><img src="${url}" alt="">`;
  ov.onclick = e => { if (e.target.closest("[data-x]") || e.target === ov) { ov.remove(); URL.revokeObjectURL(url); } }; document.body.appendChild(ov);
}

/* ---------- export ---------- */
const loadLib = src => new Promise((res, rej) => { if (document.querySelector(`script[src="${src}"]`)?.dataset.ok) return res(); const s = document.createElement("script"); s.src = src; s.onload = () => { s.dataset.ok = 1; res(); }; s.onerror = rej; document.head.appendChild(s); });
function slotsFor(from, to) {
  const S = new Map();
  for (const s of shifts) { if (s.date < from || s.date > to) continue;
    for (const [r, g] of Object.entries(s.rigs || {})) S.set(`${s.date}|${s.shift}|${r}`, { date: s.date, shift: s.shift, rig: r, metres: r1(num(g.metres)), status: g.status || "", supervisor: s.supervisor || "", block: g.block || "",
      hrs: g.hrsFrom || g.hrsTo ? `${g.hrsFrom || ""}–${g.hrsTo || ""}` : "", delays: (g.delays || []).map(d => ({ text: d.text || d.cat, cat: d.cat, hours: d.h == null ? null : +d.h })), delayText: delayText(g), report: s.report || "" }); }
  return S;
}
const folderOf = f => `${MONTHS[+f.date.slice(5, 7) - 1]} ${f.date.slice(0, 4)} Metres/${f.date.slice(8, 10)} ${(f.shift || "Day").toUpperCase()}`;
const ext = f => (String(f.name).match(/\.([a-z0-9]{2,5})$/i) || [, /pdf/.test(f.mime) ? "pdf" : "jpg"])[1].toLowerCase();
async function makeExcel(from, to, say) {
  await loadLib("../lib/exceljs.min.js");
  const F = files.filter(f => f.date >= from && f.date <= to).sort((a, b) => (a.date + a.shift + a.rig).localeCompare(b.date + b.shift + b.rig));
  const list = []; let k = 0;
  for (const f of F) { say(`Adding the documents to the Excel… ${++k} of ${F.length}`); let img = null;
    if (!/pdf/.test(f.mime)) { try { const b = await blobOf(f); if (b) img = await shrink(b, 1100, 0.6); } catch {} }
    list.push({ type: f.type, date: f.date, shift: f.shift, rig: f.rig, name: f.name, note: f.note || "", pdf: /pdf/.test(f.mime), img }); }
  say("Making the Excel…");
  let logoBuf = null; try { logoBuf = await (await fetch("../icons/icon-512.png")).arrayBuffer(); } catch {}
  const blob = await buildMetresWorkbook(window.ExcelJS, { from, to, rigs: rigs(), slots: slotsFor(from, to), files: list, tharisa: num(tharisa), logoBuf });
  const tag = `${fmt(from).replace(/\//g, "-")}${from === to ? "" : "_to_" + fmt(to).replace(/\//g, "-")}`;
  return { blob, name: `TKR_Drilling_Metres_and_Tally_Sheets_${tag}.xlsx`, tag };
}
async function doExport(kind) {
  const from = $("xFrom").value, to = $("xTo").value, msg = $("xMsg"), say = t => { if (msg) msg.textContent = t; };
  if (!from || !to || from > to) { toast("Choose a From date before the To date."); return; }
  if (!shifts.some(s => s.date >= from && s.date <= to)) { toast("No shifts captured between those dates."); return; }
  document.querySelectorAll("[data-x]").forEach(b => b.disabled = true);
  if (kind === "od") { try { await sendOneDrive(from, to, say); } catch (e) { console.warn(e); say("OneDrive: " + String(e?.errorMessage || e?.message || e).slice(0, 160)); toast("Sending to OneDrive didn't work."); } document.querySelectorAll("[data-x]").forEach(b => b.disabled = false); return; }
  try {
    const x = await makeExcel(from, to, say); let out = x.blob, name = x.name;
    if (kind === "zip") {
      await loadLib("../lib/jszip.min.js"); const zip = new window.JSZip(), root = `TKR Drilling Metres ${x.tag}`;
      zip.file(`${root}/${x.name}`, x.blob);
      const F = files.filter(f => f.date >= from && f.date <= to), used = new Set(), back = [];
      let k = 0;
      for (const f of F) { say(`Putting the documents in folders… ${++k} of ${F.length}`);
        const b = await blobOf(f).catch(() => null); if (!b) continue;
        let base = `${folderOf(f)}/${f.rig} - ${f.type}`, p = `${base}.${ext(f)}`, n = 2; while (used.has(p)) p = `${base} ${n++}.${ext(f)}`; used.add(p);
        zip.file(`${root}/${p}`, b); back.push({ ...f, thumb: undefined, path: p }); }
      const S = shifts.filter(s => s.date >= from && s.date <= to);
      zip.file(`${root}/backup/tkr-metres-backup.json`, JSON.stringify({ kind: "tkr-metres-backup", v: 1, at: Date.now(), device: cfg.device || "", from, to, rigs: cfg.rigs, shifts: S, files: back }, null, 1));
      zip.file(`${root}/backup/READ ME.txt`, `TKR Drilling Metres export, ${fmt(from)} to ${fmt(to)}.\r\nTo load these shifts on another PC or phone: open the Drilling Metres page, Export & backup, Load a backup, and choose this ZIP file.\r\n`);
      say("Making the ZIP…"); out = await zip.generateAsync({ type: "blob", compression: "STORE" }); name = `${root}.zip`;
    }
    lastExport = { at: Date.now(), from, to, kind }; await metaSet("lastExport", lastExport);
    const { deliver } = await import("../lib/deliver.js"); deliver(out, name);
    xFrom = ""; xTo = ""; render(); toast("Exported. Keep the file safe.");
  } catch (e) { console.warn(e); say("The export didn't work. Try a shorter date range."); toast("The export didn't work."); }
  document.querySelectorAll("[data-x]").forEach(b => b.disabled = false);
}
async function msReady() {
  if (!spConfigured(ms)) throw new Error("Fill in the two IDs under Settings → OneDrive.");
  await spInit(ms); if (!spReady()) { await spSignIn(); }
  if (!ms.driveId) { const L = await spLinkOneDrive(); ms = { ...ms, ...L }; await metaSet("ms", ms); await spInit(ms); }
}
async function sendOneDrive(from, to, say) {
  if (!navigator.onLine) { toast("No signal. Send to OneDrive when this device has signal."); return; }
  await msReady();
  const F = files.filter(f => !f.odAt && f.date >= from && f.date <= to); let n = 0;
  for (const f of F) { say(`Sending to OneDrive… ${n + 1} of ${F.length}`); const b = (await getOne("blobs", f.id))?.blob; if (!b) continue;
    const sp = await spUpload("Drilling/" + folderOf(f), `${f.rig} - ${f.type} - ${f.id.slice(-5)}.${ext(f)}`, b);
    await tx(["files"], "readwrite", t => { t.objectStore("files").put({ ...f, odAt: Date.now(), odUrl: sp.url }); }); n++; }
  say("Sending the Excel…"); const x = await makeExcel(from, to, say);
  await spUpload(`Drilling/${MONTHS[+to.slice(5, 7) - 1]} ${to.slice(0, 4)} Metres/Reports`, x.name, x.blob);
  await load(); render(); toast(`Sent to OneDrive: ${n} documents and the Excel.`);
}
async function restore(file) {
  const m = $("rMsg"), say = t => { if (m) m.textContent = t; };
  try { await loadLib("../lib/jszip.min.js"); say("Reading the backup…");
    const zip = await window.JSZip.loadAsync(file), bk = Object.keys(zip.files).find(p => /backup\/tkr-metres-backup\.json$/.test(p));
    if (!bk) { say("This ZIP is not a Drilling Metres backup."); return; }
    const root = bk.replace(/backup\/tkr-metres-backup\.json$/, ""), data = JSON.parse(await zip.file(bk).async("string"));
    let ns = 0, nf = 0, kept = 0;
    for (const s of data.shifts || []) { const have = shifts.find(x => x.id === s.id);
      if (have && (have.editedAt || 0) >= (s.editedAt || 0)) { kept++; continue; }
      await tx(["shifts"], "readwrite", t => { t.objectStore("shifts").put(s); }); ns++; }
    const have = new Set(files.map(f => f.id));
    for (const f of data.files || []) { if (have.has(f.id)) continue; const z = zip.file(root + f.path); if (!z) continue;
      say(`Loading documents… ${nf + 1}`); const blob = new Blob([await z.async("arraybuffer")], { type: f.mime || "" });
      let thumb = ""; if (!/pdf/.test(f.mime)) { try { thumb = (await shrink(blob, 160, 0.6)).data; } catch {} }
      const { path, ...meta } = f; await tx(["files", "blobs"], "readwrite", t => { t.objectStore("files").put({ ...meta, thumb }); t.objectStore("blobs").put({ id: f.id, blob }); }); nf++; }
    await load(); render(); toast(`Backup loaded: ${ns} shifts and ${nf} documents added${kept ? `, ${kept} already here` : ""}.`);
  } catch (e) { console.warn(e); say("The backup couldn't be read."); }
}


/* ---------- link with GET Forecast (needs signal; admins send, managers can bring in) ---------- */
const FB_CFG = { apiKey: "AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI", authDomain: "tkr-operations-app.firebaseapp.com", projectId: "tkr-operations-app",
  storageBucket: "tkr-operations-app.firebasestorage.app", messagingSenderId: "158080247079", appId: "1:158080247079:web:7cb7786d2b7f9ee172b502" };
const FC_ADMINS = ["muneer@tkroperations.com", "amogelang@tkroperations.com"];
const FC_FROM = "2026-11-01";                 // October is closed in GET Forecast; the link starts with November
const toLogin = v => { v = String(v || "").trim(); return v.includes("@") ? v : v.toLowerCase().replace(/[^a-z0-9]/g, "") + "@mgr.tkroperations.app"; };
let fb = null, fbUser = null, fcBusy = false, fcMsg = "", fcLast = 0;
async function fire() {
  if (fb) return fb;
  const m = await import("../lib/firebase.js"), a = m.initializeApp(FB_CFG), auth = m.getAuth(a), db = m.initializeFirestore(a, {});
  fb = { ...m, auth, db };
  m.onAuthStateChanged(auth, u => { fbUser = u; render(); if (u && navigator.onLine) fcSync(false); });
  return fb;
}
const fcAdmin = () => !!fbUser && FC_ADMINS.includes((fbUser.email || "").toLowerCase());
const fcWaiting = () => shifts.filter(s => s.date >= FC_FROM && (!s.fcAt || s.editedAt > s.fcAt)).length + files.filter(f => f.date >= FC_FROM && !f.fcId).length;
const CAT_RX = [["Waiting for GET / drill bits", /get|\bbits?|drill bit/], ["Stuck rods / retrieving rods", /stuck|retriev|recover/], ["Service and maintenance", /service|maint|weekly|daily|drifter/],
  ["Ground conditions", /uneven|loose ground|surface/], ["Breakdowns", /breakdown|burst|broke|leak|cylinder|engine|chain|compressor|pipe|hose|shank|starter|gauge|power|radio|hydr|oil|percussion|percation/],
  ["Tramming", /tram/], ["Blasting", /blast/], ["Diesel / refuelling", /diesel|refuel|refill/], ["Waiting for block access / bench", /access|bench|pattern|barricad/],
  ["Meetings, transport, safety and weather", /meeting|transport|safety|sheq|lightning|incident/], ["No operator", /operator/]];
const catOf = t => (CAT_RX.find(([, rx]) => rx.test(String(t).toLowerCase())) || ["Other"])[0];
const hoursIn = t => { const m = String(t).toLowerCase().replace(/,/g, ".").match(/(\d+(?:\.\d+)?)\s*(?:hrs?|hours?|h)\b/); if (m) return +m[1]; const n = String(t).match(/(\d+)\s*min/i); return n ? Math.round(+n[1] / 60 * 100) / 100 : null; };
/* GET Forecast lines for one date and shift → one shift on this device */
function fromForecast(L) {
  const s0 = L[0], rigsOut = {};
  for (const x of L) { if (!x.rig || x.rig === "Not allocated") continue; const g = rigsOut[x.rig] || (rigsOut[x.rig] = { status: "", metres: 0, block: "", hrsFrom: "", hrsTo: "", delays: [] });
    g.metres = r1(num(g.metres) + num(x.metres));
    g.status = x.status || g.status || (/^\s*not working/i.test(x.note || "") ? "Not working" : num(x.metres) > 0 || /^\s*working/i.test(x.note || "") ? "Working" : "");
    g.block = g.block || x.block || ""; g.hrsFrom = g.hrsFrom || x.hrsFrom || ""; g.hrsTo = g.hrsTo || x.hrsTo || "";
    if (Array.isArray(x.delays)) g.delays.push(...x.delays.map(d => ({ cat: d.cat || "Other", h: d.h ?? null, text: d.text || "", ...(d.ctrl ? { ctrl: d.ctrl } : {}) })));
    else String(x.note || "").replace(/^\s*(not working|working)\s*[–-]?\s*/i, "").split(/\s*·\s*|;\s*/).map(t => t.trim()).filter(Boolean).forEach(t => g.delays.push({ cat: catOf(t), h: hoursIn(t), text: t })); }
  const t = Math.max(...L.map(x => x.editedAt || x.createdAt || 0));
  return { id: s0.date + "|" + (s0.shift || "Day"), date: s0.date, shift: s0.shift || "Day", supervisor: L.map(x => x.supervisor).find(Boolean) || "", report: L.map(x => x.report).find(Boolean) || "", rigs: rigsOut, savedAt: t, editedAt: t, fcAt: t, device: "GET Forecast" };
}
/* GET Forecast's documents list (no photos): every document shows here; the photo is fetched when opened or exported */
async function fcPullDocs(from, to) {
  const { db, collection, getDocs, query, where } = fb;
  const snap = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "day")));
  const days = snap.docs.map(d => d.data()).filter(v => v.ev_date >= from && v.ev_date <= to);
  const items = days.filter(v => Array.isArray(v.items)).flatMap(v => v.items.map(it => ({ ...it, date: v.ev_date })));
  for (const v of days.filter(v => !Array.isArray(v.items))) {      // older days filed before the documents list existed
    const L = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "sheet"), where("ev_date", "==", v.ev_date)));
    L.docs.forEach(d => { const x = d.data(); items.push({ id: d.id, type: x.type || "Tally sheet", rig: x.rig || "", shift: x.shift || "", name: x.name || "", note: x.note || "", pdf: /pdf/i.test(x.mime || "") ? 1 : 0, at: x.at || 0, date: v.ev_date }); }); }
  const known = new Set(files.map(f => f.fcId).filter(Boolean)), live = new Set(items.map(it => it.id));
  const add = items.filter(it => !known.has(it.id)), drop = files.filter(f => f.remote && f.date >= from && f.date <= to && !live.has(f.fcId));
  if (!add.length && !drop.length) return 0;
  await tx(["files", "blobs"], "readwrite", t => {
    for (const it of add) t.objectStore("files").put({ id: "r_" + it.id, fcId: it.id, remote: 1, shiftId: it.date + "|" + (it.shift || "Day"), date: it.date, shift: it.shift || "Day", rig: it.rig || "", type: it.type || "Tally sheet",
      name: it.name || it.type || "Document", mime: it.pdf ? "application/pdf" : "image/jpeg", size: 0, thumb: "", note: it.note || "", savedAt: it.at || 0 });
    for (const f of drop) { t.objectStore("files").delete(f.id); t.objectStore("blobs").delete(f.id); }
  });
  return add.length;
}
async function fcPull(from, to) {
  const { db, collection, getDocs, query, where } = fb;
  const snap = await getDocs(query(collection(db, "fcMetres"), where("date", ">=", from), where("date", "<=", to)));
  const by = new Map(); snap.docs.forEach(d => { const x = { id: d.id, ...d.data() }; const k = x.date + "|" + (x.shift || "Day"); (by.get(k) || by.set(k, []).get(k)).push(x); });
  let n = 0;
  for (const [k, L] of by) { const loc = shifts.find(s => s.id === k), t = Math.max(...L.map(x => x.editedAt || x.createdAt || 0));
    if (loc && ((loc.fcAt || 0) >= t || (loc.editedAt || 0) >= t)) continue;
    const s = fromForecast(L); if (!Object.keys(s.rigs).length) continue;
    if (loc) s.savedAt = loc.savedAt;
    await tx(["shifts"], "readwrite", tt => { tt.objectStore("shifts").put(s); }); n++; }
  return n;
}
const shrinkData = async (blob) => { let q = 0.7; let r = await shrink(blob, 1400, q); while (r.data.length > 900000 && q > 0.35) { q -= 0.1; r = await shrink(blob, 1400, q); } return r.data; };
async function fcIndexDay(date) {
  const { db, collection, getDocs, query, where, doc, setDoc, deleteDoc } = fb;
  const s = await getDocs(query(collection(db, "fcSettings"), where("ev_kind", "==", "sheet"), where("ev_date", "==", date)));
  const items = s.docs.map(d => { const v = d.data(); return { id: d.id, type: v.type || "Tally sheet", rig: v.rig || "", shift: v.shift || "", name: v.name || "", note: v.note || "", pdf: /pdf/i.test(v.mime || "") ? 1 : 0, sp: v.sp?.url || "", at: v.at || 0, by: v.by || "" }; });
  const ref = doc(db, "fcSettings", "evday_" + date);
  if (items.length) await setDoc(ref, { ev_kind: "day", ev_date: date, n: items.length, items }); else await deleteDoc(ref);
}
async function fcPush(say) {
  const { db, collection, getDocs, query, where, doc, setDoc, deleteDoc } = fb, me = fbUser.email || "";
  const lineId = (d, sh, r) => `sh-${d}-${sh}-${r.replace(/\s/g, "")}`;
  const remote = {}, getDay = async d => remote[d] || (remote[d] = (await getDocs(query(collection(db, "fcMetres"), where("date", "==", d)))).docs.map(x => ({ id: x.id, ...x.data() })));
  const days = new Set(); let ns = 0, nf = 0;
  // shifts deleted on this device
  for (const g of (await metaGet("fcDel")) || []) { for (const x of await getDay(g.date)) if ((x.shift || "Day") === g.shift && cfg.rigs.includes(x.rig)) await deleteDoc(doc(db, "fcMetres", x.id)); }
  await metaSet("fcDel", []);
  for (const s of shifts.filter(s => s.date >= FC_FROM && (!s.fcAt || s.editedAt > s.fcAt))) {
    say?.(`Sending shifts to GET Forecast… ${++ns}`);
    const R = await getDay(s.date);
    for (const r of cfg.rigs) { const g = s.rigs[r], mine = lineId(s.date, s.shift, r);
      for (const x of R.filter(x => (x.shift || "Day") === s.shift && x.rig === r && x.id !== mine)) await deleteDoc(doc(db, "fcMetres", x.id));
      if (!g) { if (R.some(x => x.id === mine)) await deleteDoc(doc(db, "fcMetres", mine)); continue; }
      const note = (g.status || "") + (g.delays.length ? " – " + g.delays.map(d => (d.text || d.cat) + (d.h ? ` (${hm(d.h)})` : "")).join(" · ") : "");
      await setDoc(doc(db, "fcMetres", mine), { date: s.date, shift: s.shift, rig: r, metres: num(g.metres), block: g.block || "", supervisor: s.supervisor || "", status: g.status || "", delays: g.delays, note,
        hrsFrom: g.hrsFrom || "", hrsTo: g.hrsTo || "", report: s.report || "", by: me, src: "metres", createdAt: s.savedAt || Date.now(), editedAt: s.editedAt || Date.now() }); }
    await tx(["shifts"], "readwrite", t => { t.objectStore("shifts").put({ ...s, fcAt: s.editedAt }); });
  }
  // documents removed on this device
  for (const g of (await metaGet("fcGone")) || []) { try { const v = (await fb.getDoc(doc(db, "fcSettings", g.fcId))).data(); for (let i = 0; i < (v?.parts || 0); i++) await deleteDoc(doc(db, "fcSettings", g.fcId + "_p" + i)); await deleteDoc(doc(db, "fcSettings", g.fcId)); days.add(g.date); } catch (e) { console.warn(e); } }
  await metaSet("fcGone", []);
  // new documents
  for (const f of files.filter(f => f.date >= FC_FROM && !f.fcId)) {
    say?.(`Sending documents to GET Forecast… ${++nf}`);
    const b = (await getOne("blobs", f.id))?.blob; if (!b) continue;
    const isPdf = /pdf/.test(f.mime); let data = "", parts = [];
    if (isPdf) { if (b.size > 8 * 1024 * 1024) continue; const d = await dataUrl(b); if (d.length <= 900000) data = d; else for (let i = 0; i < d.length; i += 700000) parts.push(d.slice(i, i + 700000)); }
    else { try { data = await shrinkData(b); } catch { continue; } }
    const id = ("ev_" + f.date + "_" + f.rig.replace(/\s/g, "") + "_" + (f.shift || "NA") + "_" + f.type.replace(/\s/g, "") + "_m" + f.id.slice(-6)).replace(/[^\w-]/g, "_");
    for (let i = 0; i < parts.length; i++) await setDoc(doc(db, "fcSettings", id + "_p" + i), { ev_kind: "part", parent: id, i, data: parts[i] });
    await setDoc(doc(db, "fcSettings", id), { ev_kind: "sheet", ev_date: f.date, rig: f.rig, shift: f.shift, type: f.type, name: f.name, mime: isPdf ? "application/pdf" : "image/jpeg", note: f.note || "", data, parts: parts.length, sp: null, by: me, at: f.savedAt || Date.now(), src: "metres" });
    await tx(["files"], "readwrite", t => { t.objectStore("files").put({ ...f, fcId: id }); }); days.add(f.date);
  }
  for (const d of days) { say?.("Updating the documents list…"); await fcIndexDay(d); }
}
async function fcSync(manual, say) {
  if (fcBusy || !navigator.onLine) { if (manual && !navigator.onLine) toast("No signal now. It sends by itself when there is signal."); return; }
  try { await fire(); } catch (e) { console.warn(e); return; }
  if (!fbUser) { if (manual) toast("Sign in first (Export & backup → GET Forecast)."); return; }
  fcBusy = true; fcMsg = "Checking GET Forecast…"; render();
  try {
    const t = today(), m0 = (month < t.slice(0, 7) ? month : t.slice(0, 7)) + "-01", from = m0 < FC_FROM ? FC_FROM : m0;
    const pulled = await fcPull(from, t); try { await fcPullDocs("2026-10-01", t); } catch (e) { console.warn(e); } await load(); await linkCtrl();
    if (fcAdmin()) { await fcPush(x => { fcMsg = x; const el = $("fcMsg"); if (el) el.textContent = x; }); await load(); }
    fcLast = Date.now(); await metaSet("fcLast", fcLast);
    fcMsg = `Up to date with GET Forecast (${when(fcLast)})${pulled ? ` · ${pulled} shift${pulled === 1 ? "" : "s"} brought in` : ""}.`;
    if (manual) toast(fcAdmin() ? "GET Forecast is up to date." : `Brought in ${pulled} shifts. Only an admin can send shifts to GET Forecast.`);
  } catch (e) { console.warn(e); fcMsg = e?.code === "permission-denied" ? "This account can't change GET Forecast. Sign in as an admin." : "Couldn't reach GET Forecast. It tries again when there is signal."; if (manual) toast(fcMsg); }
  fcBusy = false; render();
}
function fcCard() {
  const w = fcWaiting();
  if (!fbUser) return `<section class="card" id="fcCard"><h2>GET Forecast link</h2><p class="muted" style="margin-top:0">Sign in once on this device with your app email (an admin). After that, whenever there is signal, new shifts and their documents go to GET Forecast by themselves, and metres captured in GET Forecast come in here. ${navigator.onLine ? "" : "<b>No signal now.</b>"}</p>
    <form id="fcSign" class="grid" autocomplete="on" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))"><label>Email or username <input name="e" autocapitalize="none" autocomplete="username" required></label><label>Password <input name="p" type="password" autocomplete="current-password" required></label><div style="display:flex;align-items:flex-end"><button class="primary" style="width:100%">Sign in</button></div></form>
    <p class="muted" id="fcMsg">${esc(fcMsg)}</p></section>`;
  return `<section class="card" id="fcCard"><h2>GET Forecast link</h2>
    <p style="margin-top:0">Signed in as <b>${esc(fbUser.email)}</b>${fcAdmin() ? "" : " (can bring in, can't send)"} · ${w ? `<b>${w}</b> shifts or documents waiting to go` : "nothing waiting"}${fcLast ? ` · last checked ${when(fcLast)}` : ""}.</p>
    <div class="row"><button class="primary" id="fcNow" ${fcBusy ? "disabled" : ""}>${fcBusy ? "Working…" : "Send and bring in now"}</button><button class="ghost" id="fcOut">Sign out</button></div>
    <p class="muted" id="fcMsg">${esc(fcMsg)}</p></section>`;
}
addEventListener("online", () => { if (fbUser) fcSync(false); else render(); });

/* ---------- link with the Control Room: its breakdown, standby and maintenance times for the drill rigs
   come in as delays (hours and minutes only) on the matching shift. Read from the Control Room on this PC and,
   with signal, from the TKR app. ---------- */
const CTRL_CAT = { BD: "Breakdowns", PM: "Service and maintenance", SB: "Other" }, CTRL_NAME = { BD: "Breakdown", PM: "Planned maintenance", SB: "Standby" };
let ctrlEvs = new Map();
const ctMs = s => { if (!s) return 0; const [d, t] = s.split("T"), [y, m, dd] = d.split("-").map(Number), [h, mi] = (t || "00:00").split(":").map(Number); return new Date(y, m - 1, dd, h, mi).getTime(); };
function takeCtrl(L) { for (const r of L) { if (r.kind !== "ev") continue; const o = ctrlEvs.get(r.id); if (!o || (r.editedAt || 0) > (o.editedAt || 0)) ctrlEvs.set(r.id, r); } }
async function readCtrlLocal() {
  try { const L = await indexedDB.databases?.(); if (!L || !L.some(d => d.name === "tkr-control")) return; } catch { return; }
  const d = await new Promise((res, rej) => { const r = indexedDB.open("tkr-control"); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  try { if (!d.objectStoreNames.contains("rec")) return; takeCtrl(await req(d.transaction("rec").objectStore("rec").getAll())); } finally { d.close(); }
}
async function readCtrlRemote() {
  if (!navigator.onLine || !fbUser) return; const { db, collection, getDocs, query, where } = fb, t = today().slice(0, 7);
  const months = [...new Set([shiftMonth(t, -1), t, month])];
  takeCtrl((await getDocs(query(collection(db, "ctrl"), where("ym", "in", months)))).docs.map(x => ({ id: x.id, ...x.data() })));
}
function ctrlDelays(date, sh, rig) {
  const mc = rig.replace(/\s+/g, ""), a = ctMs(date + "T" + (sh === "Day" ? "06:00" : "18:00")), b = a + 12 * 36e5, now = Date.now();
  return [...ctrlEvs.values()].filter(e => !e.deleted && e.machine === mc && CTRL_CAT[e.type])
    .map(e => { const s0 = Math.max(a, ctMs(e.start)), e0 = Math.min(b, e.end ? ctMs(e.end) : now), min = Math.round((e0 - s0) / 6e4); return { e, min }; })
    .filter(x => x.min > 0).sort((x, y) => ctMs(x.e.start) - ctMs(y.e.start))
    .map(({ e, min }) => ({ cat: CTRL_CAT[e.type], h: Math.round(min / 60 * 100) / 100, text: `${e.reason || CTRL_NAME[e.type]} (Control Room${e.end ? "" : ", still going"})`, ctrl: e.id }));
}
async function linkCtrl() {
  try { await readCtrlLocal(); } catch (e) { console.warn(e); }
  try { await readCtrlRemote(); } catch (e) { console.warn("control room not reachable", e?.code || e); }
  if (!ctrlEvs.size) return 0;
  const now = Date.now(), changed = [];
  for (const s of shifts) { if (s.date < FC_FROM) continue; const R2 = { ...(s.rigs || {}) }; let ch = false;
    for (const r of allRigs()) { const g = R2[r], cd = ctrlDelays(s.date, s.shift, r), own = (g?.delays || []).filter(d => !d.ctrl), nd = [...own, ...cd];
      if (JSON.stringify(nd) === JSON.stringify(g?.delays || [])) continue;
      if (!g && !cd.length) continue;
      R2[r] = { ...(g || { status: "", metres: "", block: "", hrsFrom: "", hrsTo: "" }), delays: nd }; ch = true; }
    if (ch) changed.push({ ...s, rigs: R2, editedAt: now }); }
  if (changed.length) { await tx(["shifts"], "readwrite", t => { for (const s of changed) t.objectStore("shifts").put(s); }); await load(); }
  return changed.length;
}
let linking = false;
async function relink() { if (linking) return; linking = true; try { if (await linkCtrl()) render(); } finally { linking = false; } }
addEventListener("focus", relink); document.addEventListener("visibilitychange", () => { if (!document.hidden) relink(); });

/* ---------- events ---------- */
app.addEventListener("click", async e => {
  const b = e.target.closest("button"); if (!b) return;
  if (b.dataset.tab) { tab = b.dataset.tab; render(); scrollTo(0, 0); return; }
  if (b.dataset.mnav) { month = shiftMonth(month, +b.dataset.mnav); render(); return; }
  if (b.dataset.cap !== undefined) { const v = b.dataset.cap; if (!v) openForm(null); else if (shifts.some(s => s.id === v)) openForm(v); else { const [d, sh] = v.split("|"); openForm(null, { date: d, shift: sh }); } return; }
  if (b.dataset.del) { askDelete(b.dataset.del); return; }
  if (b.dataset.view) { viewFile(b.dataset.view); return; }
  if (b.dataset.x) { b.blur(); doExport(b.dataset.x); return; }
  if (b.id === "fcNow") { fcSync(true); return; }
  if (b.id === "fcOut") { await fire(); await fb.signOut(fb.auth); fcMsg = ""; render(); return; }
  if (b.id === "msSave") { ms = { ...ms, tenant: $("msTenant").value.trim(), clientId: $("msClient").value.trim(), folder: $("msFolder").value.trim() || "TKR App" }; await metaSet("ms", ms); toast("Saved"); render(); return; }
  if (b.dataset.ms) { try { if (b.dataset.ms === "in") { await msReady(); msState = ""; toast("OneDrive is linked on this device."); }
      if (b.dataset.ms === "out") { await spSignOut(); }
      if (b.dataset.ms === "open") { await spInit(ms); window.open(await spFolderUrl("Drilling"), "_blank", "noopener"); } }
    catch (x) { console.warn(x); msState = "Microsoft said: " + String(x?.errorMessage || x?.message || x).slice(0, 160); toast("OneDrive isn't linked yet."); } render(); return; }
  if (metresChartClick(b)) { render(); return; }
  if (b.id === "addRig") { let r = $("newRig").value.trim().toUpperCase().replace(/^([A-Z]+)\s*0*(\d+)$/, (m, a, d) => `${a} ${d.padStart(3, "0")}`); if (!r) { toast("Type the rig name."); return; }
    if (allRigs().includes(r)) { toast(r + " is already on the list."); return; }
    cfg = { ...cfg, extraRigs: [...(cfg.extraRigs || []), r], rigs: [...cfg.rigs, r] }; await metaSet("cfg", cfg); toast(r + " added. It shows on the shift form now."); render(); return; }
  if (b.id === "cfgSave") { const r = [...app.querySelectorAll("input[name=rig]:checked")].map(i => i.value); if (!r.length) { toast("Choose at least one rig."); return; }
    cfg = { ...cfg, rigs: r, device: $("devName").value.trim() }; await metaSet("cfg", cfg); toast("Saved"); render(); }
});
app.addEventListener("submit", async e => {
  if (e.target.id !== "fcSign") return; e.preventDefault(); const f = e.target, btn = f.querySelector("button");
  if (!navigator.onLine) { toast("Signing in needs signal."); return; }
  btn.disabled = true; fcMsg = "Signing in…";
  try { await fire(); await fb.signInWithEmailAndPassword(fb.auth, toLogin(f.elements.e.value), f.elements.p.value); fcMsg = ""; }
  catch (x) { fcMsg = String(x?.code || "").includes("network") ? "No signal. Signing in needs signal." : "Email or password is wrong."; const el = $("fcMsg"); if (el) el.textContent = fcMsg; btn.disabled = false; return; }
});
app.addEventListener("input", e => { if (e.target.id === "q") { q = e.target.value; render(); } });
app.addEventListener("change", e => {
  const id = e.target.id;
  if (id === "xFrom") xFrom = e.target.value; if (id === "xTo") xTo = e.target.value; if (id === "xTh") tharisa = e.target.value.trim();
  if (id === "restore" && e.target.files[0]) restore(e.target.files[0]);
});
addEventListener("keydown", e => { if (e.key === "Escape" && form) closeForm(); });

(async () => { if (await load()) { try { await ensureOctober(); await load(); } catch (e) { console.warn(e); } try { await linkCtrl(); } catch (e) { console.warn(e); } render(); if (navigator.onLine) fire().catch(e => console.warn(e)); if (navigator.onLine && ms.driveId) { try { await spInit(ms); render(); } catch (e) { console.warn(e); } } } })();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("../sw.js", { updateViaCache: "none" }).then(r => r.update()).catch(() => {});
