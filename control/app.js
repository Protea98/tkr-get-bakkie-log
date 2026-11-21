/* TKR Control Room – works with no signal on the control room PC.
   Controllers add themselves (name + 4-digit PIN) on this device. Breakdowns, standby, planned maintenance and
   "not in use" are logged with the exact start and end time; the documents checklist and the invoicing lists are
   ticked off. Everything is saved on this device at once, and goes to the TKR app when there is signal. */
import { tabBar, pageHead, sheet, sheetClick, afterRender, closeSheet, icon } from "../lib/ui.js";
import { bdData, chartPng } from "../lib/metres-xlsx.js";

const $ = id => document.getElementById(id);
const app = $("app");
const pad = n => String(n).padStart(2, "0");
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parse = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return ymd(d); };
const fmt = s => s ? s.split("-").reverse().join("/") : "";
const dow = s => parse(s).toLocaleDateString("en-ZA", { weekday: "short" });
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthName = m => MONTHS[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
const monthEnd = m => ymd(new Date(+m.slice(0, 4), +m.slice(5, 7), 0));
const shiftMonth = (m, n) => ymd(new Date(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)).slice(0, 7);
const nowLocal = () => { const d = new Date(Date.now()); return `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const toMs = s => { if (!s) return 0; const [d, t] = s.split("T"), [y, m, dd] = d.split("-").map(Number), [h, mi] = (t || "00:00").split(":").map(Number); return new Date(y, m - 1, dd, h, mi).getTime(); };
const hhmm = s => s ? s.slice(11, 16) : "";
const hm = h => { const t = Math.round(h * 60); return `${Math.floor(t / 60)}:${pad(t % 60)}`; };
const h2 = h => hm(h || 0);   // hours always shown as hours:minutes
const today = () => ymd(new Date(Date.now()));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
function toast(m) { const t = $("toast"); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3200); }

/* ---------- what the control room tracks (from the November control room workbook) ---------- */
const TYPES = { BD: "Breakdown", SB: "Standby (available, not operating)", PM: "Planned maintenance", NU: "Not in use" };
const DEFAULT_FLEET = [["Drill rigs", ["TDR013", "TDR014", "TDR015"]], ["Rockbreakers", ["TKRH011", "TKRH015", "TKRH017", "TKRH019", "TKRH020"]],
  ["Buses", ["TKR020", "TKR021", "TKR023", "TKR024", "TKR025", "TKR026"]], ["Lowbed", ["TKR06"]], ["Water bowser", ["TKR09"]], ["Bakkies", ["LDV011", "LDV012"]]];
const CHECK = [["Collect data", ["PTO", "Inspection reports", "SOP", "T Connect attendance", "Risk assessment", "Capture", "Pit licence expiring"]], ["Documents", ["Job cards", "Checklists", "Fleet reports", "Trip summary"]],
  ["Systems", ["Cartrack", "On / Off", "Planned maintenance"]], ["Cartrack", ["Predictive failures", "Mechanical inspections", "Operator behaviour", "Fuel consumption management system", "Geofences"]]];
const COLS = ["Rompad", "Drilling", "Buses"];
const SUPPORT = [["Drilling", ["Tally sheets / fleet report", "Job cards", "Checklists"], "Lebo / Godfrey"], ["Rompad", ["Job cards and checklists", "Fleet report"], "Refilwe"], ["Buses", ["Checklists", "Trip summary"], "Naledi"],
  ["Cementation", ["Checklist", "Trip summary"], "Zandile"], ["Pit Project", ["Job card", "Checklist", "Fleet report"], "Naledi"], ["Water cart", ["Checklist", "Trip summary", "Machine list"], "Zandile"],
  ["Lowbed", ["Checklist", "Trip summary"], "Zandile"], ["TLB operator", ["Clockings"], "Naledi"]];
const ADDITIONAL = [["HSEC compliance", "Tsholo (HSEC)"], ["Clockings", "Amogelang"], ["AWOP", "Amogelang"], ["Resignations & New Recruits", "Truly"], ["Attendance register", "Amogelang"],
  ["Update of Machine location", "Naledi"], ["Update of Checklists & Jobcards", "Muneer"], ["COF legals compliance", "Naledi"], ["Pit licence expiry dates", "Tsholo (HSEC)"], ["Leave register", "Zandi (Rom) / Riaan (Drilling)"],
  ["Planned Overtime", "Zandi (Rom) / Riaan (Drilling)"], ["Medicals", "Tsholo (HSEC)", "Everyone returning from leave must go for medicals"], ["Investigate overbookings for Oct", "Godfrey"], ["Planned Maintenance and servicing", "Grace"],
  ["Daily Machine Availability", "Grace"], ["SOPs and PTOs, Inspection Report, Tee connect attendance, attendance & risk assessment", "(Safety) Zane & Nico"], ["Pit licence", "Tsholo (HSEC)", "Everyone returning from leave must do the pit licence induction"]];

/* ---------- storage on this device ---------- */
let dbp = null;
const idb = () => dbp || (dbp = new Promise((res, rej) => { const r = indexedDB.open("tkr-control", 2);
  r.onupgradeneeded = () => { const d = r.result; if (!d.objectStoreNames.contains("rec")) { const s = d.createObjectStore("rec", { keyPath: "id" }); s.createIndex("ym", "ym"); }
    if (!d.objectStoreNames.contains("meta")) d.createObjectStore("meta", { keyPath: "k" }); if (!d.objectStoreNames.contains("files")) d.createObjectStore("files", { keyPath: "id" }); };
  r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }));
const tx = async (mode, fn) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(["rec", "meta"], mode); fn(t); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); };
const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const all = async () => { const d = await idb(); return req(d.transaction("rec").objectStore("rec").getAll()); };
const metaGet = async k => { const d = await idb(); return (await req(d.transaction("meta").objectStore("meta").get(k)))?.v; };
const metaSet = (k, v) => tx("readwrite", t => t.objectStore("meta").put({ k, v }));
const fileGet = async id => { const d = await idb(); return (await req(d.transaction("files").objectStore("files").get(id)))?.data; };
const fileSet = async (id, data) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction("files", "readwrite"); t.objectStore("files").put({ id, data }); t.oncomplete = res; t.onerror = () => rej(t.error); }); };

/* Every record: { id, kind, ym, editedAt, by, deleted?, ... } — kinds: user, ev, chk, day, log, inv, add, fleet */
let cvMemo = new Map(); const bump = () => cvMemo.clear();
let recs = [], me = null, tab = "now", date = today(), month = today().slice(0, 7), pinFor = null, addingMe = false;
const R = k => recs.filter(r => r.kind === k && !r.deleted);
/* Locking: every filled box remembers the shift (session) that filled it. Once that controller logs out,
   the box is grey and locked for everyone else; empty boxes stay open. Locked boxes change only with admin approval. */
const CTRL_ADMIN = ["muneer@tkroperations.com"];
const TRACK = new Set(["hr", "ev", "chk", "inv", "add"]);
const META_K = new Set(["id", "kind", "ym", "editedAt", "by", "synced", "fs", "deleted", "date", "at", "group", "grid", "endBy", "sid", "item", "svc", "doc"]);
const filled = v => v != null && v !== "";
const fieldsOf = r => { const m = new Map(); for (const k in r) { if (META_K.has(k)) continue; if (k === "cells") { for (const h in r.cells || {}) m.set("h" + h, r.cells[h]); } else m.set(k, r[k]); } return m; };
function stamp(rec, old) {
  if (!TRACK.has(rec.kind) || !me) return rec;
  const fs = { ...(old?.fs || {}) }, O = old ? fieldsOf(old) : new Map();
  if (old && !old.fs) for (const [k, v] of O) if (filled(v)) fs[k] = "old";
  for (const [k, v] of fieldsOf(rec)) if (JSON.stringify(v ?? "") !== JSON.stringify(O.get(k) ?? "")) fs[k] = me.sid;
  return { ...rec, fs };
}
/* A box is locked only when it belongs to an earlier shift (before the current 06:00 / 18:00 change) and was not filled by the
   controller logged in now. Everything in the current shift stays open for the controller on duty, whoever filled it. */
function fieldTime(r, f) {
  if (r.kind === "hr" && /^h\d+$/.test(f)) return hStart(r.date, +f.slice(1));
  if (r.kind === "ev") return toMs(f === "end" ? (r.end || r.start) : r.start);
  if (r.kind === "chk") return toMs(r.date + "T23:59");
  return r.editedAt || 0;
}
function lockedF(r, f) { if (!r || !me || me.admin) return false; if (r.fs?.[f] && r.fs[f] === me.sid) return false; return fieldTime(r, f) < curShift().start; }
const valOf = (r, f) => /^h\d+$/.test(f) && r?.cells ? r.cells[f.slice(1)] : r?.[f];
const lockedV = (r, f) => !!r && filled(valOf(r, f)) && lockedF(r, f);
const evLocked = e => ["machine", "type", "start"].some(f => lockedV(e, f));
/* every change since the last Save, so Discard can put things back the way they were */
let undo = [];
function journal(id, old) { if (!me || undo.some(u => u.id === id)) return; undo.push({ id, prev: old ? JSON.parse(JSON.stringify(old)) : null }); metaSet("undo", undo).catch(() => {}); }
async function clearUndo() { undo = []; await metaSet("undo", []).catch(() => {}); }
async function discardChanges() {
  const now = Date.now(), L = undo.slice().reverse(), out = [];
  for (const u of L) { const cur = recs.find(x => x.id === u.id); if (!cur && !u.prev) continue;
    out.push(u.prev ? { ...u.prev, editedAt: now, synced: 0 } : { ...cur, deleted: true, editedAt: now, synced: 0 }); }
  await tx("readwrite", t => { for (const r of out) t.objectStore("rec").put(r); });
  for (const r of out) { const i = recs.findIndex(x => x.id === r.id); if (i >= 0) recs[i] = r; else recs.push(r); } bump();
  const n = undo.filter(u => recs.find(x => x.id === u.id)?.kind !== "log").length; await clearUndo(); syncSoon(); return n;
}
async function put(r, log) {
  const old = recs.find(x => x.id === r.id); journal(r.id, old);
  const rec = { ...stamp(r, old), editedAt: Date.now(), by: me?.name || r.by || "", synced: 0 };
  const L = log ? { id: "log_" + uid(), kind: "log", ym: (r.date || r.start || today()).slice(0, 7), date: (r.date || r.start || today()).slice(0, 10), at: Date.now(), by: me?.name || "", what: log, editedAt: Date.now(), synced: 0 } : null;
  if (L) journal(L.id, null);
  await tx("readwrite", t => { t.objectStore("rec").put(rec); if (L) t.objectStore("rec").put(L); });
  const i = recs.findIndex(x => x.id === rec.id); if (i >= 0) recs[i] = rec; else recs.push(rec); if (L) recs.push(L); bump();
  syncSoon(); return rec;
}
const fleet = () => { const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0]; return f?.groups || DEFAULT_FLEET; };
const groupOf = m => (fleet().find(([, L]) => L.includes(m)) || ["Other"])[0];

/* ---------- hours: events overlapping a time window ---------- */
const evs = () => R("ev");
/* machines not on site: left out of availability, the board, the message and the documents check for that time */
const siteRec = () => recs.find(r => r.id === "site" && !r.deleted) || { id: "site", kind: "site", ym: "setup", off: {} };
const offAt = (m, t) => (siteRec().off?.[m] || []).some(p => p.from <= t && (!p.to || t < p.to));
const offNow = m => offAt(m, Date.now());
const offShift = (m, d, sh) => { const [a, b] = shiftWin(d, sh); return offAt(m, a) && offAt(m, b - 1); };
const evEnd = e => e.end ? toMs(e.end) : Date.now();
function hoursIn(machine, from, to) {
  const out = { BD: 0, SB: 0, PM: 0, NU: 0 };
  for (const e of evs()) { if (e.machine !== machine) continue; const a = Math.max(toMs(e.start), from), b = Math.min(evEnd(e), to); if (b > a) out[e.type] = (out[e.type] || 0) + (b - a) / 36e5; }
  return out;
}
/* one machine, one date (06:00 to 06:00 next day), up to now */
/* first hour of the run of BD / PM that is going on now in the sheet */
function sheetSince(m, t) { const c = curShift(), [d, sh] = c.key.split("|"), C = cellsOf(d, m), now = Date.now(); const H = HRS(sh).filter(x => hStart(d, x) <= now);
  let k = H.length - 1; while (k >= 0 && !C[H[k]]) k--; if (k < 0 || C[H[k]] !== t) return ""; while (k > 0 && C[H[k - 1]] === t) k--; return pad(H[k] % 24) + ":00"; }
/* why a machine is not working now: the reason on its breakdown or maintenance, or the shift comment in the sheet */
function reasonNow(m) { const s = statusNow(m); if (s?.reason) return s.reason.trim();
  const c = curShift(), [d, sh] = c.key.split("|"), [a] = shiftWin(d, sh), now = Date.now();
  const e = evs().filter(e => e.machine === m && toMs(e.start) <= now && evEnd(e) > a && e.reason).sort((x, y) => toMs(y.start) - toMs(x.start))[0];
  if (e) return e.reason.trim(); return (recs.find(r => r.id === `hr_${d}_${m}`)?.notes?.[sh] || "").trim(); }
function stateNow(m) { if (offNow(m)) return "OFF"; const s = statusNow(m), c = curShift(), [d, sh] = c.key.split("|"), C = cellsOf(d, m), now = Date.now();
  const h = HRS(sh).filter(x => hStart(d, x) <= now && C[x]).at(-1);
  if (s && (h == null || toMs(s.start) >= hStart(d, h))) return s.type;     // the newest of the logged time and the sheet wins
  if (h != null) return C[h] === "-" ? "?" : C[h]; if (s) return s.type;
  return lastKnown(m) || "?"; }
/* nothing marked yet this shift: the machine's last marked status (up to 2 days back), so the dashboard and the message keep showing it */
function lastKnown(m, from = Date.now()) {
  const E = evs().filter(e => e.machine === m && ["BD", "PM", "SB", "NU"].includes(e.type));
  for (let t = from - 36e5, i = 0; i < 24 * 14; i++, t -= 36e5) { const dt = new Date(t), hh = dt.getHours(), d0 = ymd(dt), d = hh < 6 ? addDays(d0, -1) : d0, h = hh < 6 ? hh + 24 : hh;
    const v = cellsOf(d, m)[h]; if (v && v !== "-") return v;
    const e = E.filter(e => toMs(e.start) < t + 36e5 && evEnd(e) > t).sort((x, y) => toMs(y.start) - toMs(x.start))[0]; if (e) return e.type; }   // up to 14 days back
  return ""; }   // nothing filled in this shift yet   // nothing marked this shift: working (clean slate)
function statusNow(m) { const t = Date.now(); return evs().filter(e => e.machine === m && toMs(e.start) <= t && (!e.end || toMs(e.end) > t)).sort((a, b) => toMs(b.start) - toMs(a.start))[0] || null; }

/* ---------- controllers on this device ---------- */
async function pinHash(name, pin) {
  const s = "tkr-ctl|" + name.toLowerCase() + "|" + pin;
  try { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, "0")).join(""); }
  catch { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) | 0; return "x" + h; }
}
const adminRec = () => R("admin").sort((a, b) => b.editedAt - a.editedAt)[0];
let adminStep = "";
const okUser = u => !u.status || u.status === "approved";   // controllers added before 20/11/2026 stay approved
const pendingUsers = () => R("user").filter(u => u.status === "pending");
function gateView() {
  const users = R("user").filter(u => u.status !== "declined").sort((a, b) => a.name.localeCompare(b.name));
  if (!adminStep && (addingMe || !users.length)) return pageHead("Control Room", "Ask to be a controller", "", "", "Type your name and choose a 4-digit PIN you will remember. The admin must approve you before you can sign in. The request sends when the PC has signal.")
    + `<form id="fAdd" class="card" style="max-width:460px" autocomplete="off"><label>Name and surname <input name="name" required autocapitalize="words"></label>
      <label style="margin-top:10px;display:block">Choose a 4-digit PIN <input name="pin" class="pin" inputmode="numeric" maxlength="4" pattern="\\d{4}" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" required></label>
      <label style="margin-top:10px;display:block">Type the PIN again <input name="pin2" class="pin" inputmode="numeric" maxlength="4" pattern="\\d{4}" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" required></label>
      <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Send to the admin</button>${users.length ? '<button type="button" class="ghost" id="cancelAdd">Back</button>' : ""}</div></form>`;
  if (adminStep === "pin") { const a = adminRec();
    return pageHead("Control Room", "Admin", "", "", "Type your admin PIN. No signal needed.")
      + `<form id="fAdmPin" class="card" style="max-width:360px" autocomplete="off"><input name="pin" class="pin" inputmode="numeric" maxlength="4" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" required>
        <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Open as admin</button><button type="button" class="ghost" id="cancelAdm">Back</button></div><p class="flag" id="pinErr"></p>
        <p class="muted" style="margin-bottom:0">Forgot it? <button type="button" class="sm ghost" id="admReset">Sign in with your TKR email again</button></p></form>`; }
  if (adminStep === "email") return pageHead("Control Room", "Admin sign-in", "", "", "Once only on this PC, with signal: sign in with your TKR admin email and password, then choose a 4-digit admin PIN. After that the PIN works with no signal.")
    + `<form id="fAdm" class="card" style="max-width:460px" autocomplete="off"><label>TKR email <input name="e" type="email" autocapitalize="none" required></label>
      <label style="margin-top:10px;display:block">Password <input name="p" type="password" required></label>
      <label style="margin-top:10px;display:block">Choose a 4-digit admin PIN <input name="pin" class="pin" inputmode="numeric" maxlength="4" pattern="\\d{4}" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" required></label>
      <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Sign in</button><button type="button" class="ghost" id="cancelAdm">Back</button></div>
      <p class="flag" id="admErr"></p><button type="button" class="ghost sm" id="admForgot" style="margin-top:6px">Forgot password? Email me a reset link</button></form>`;
  if (pinFor) { const u = users.find(x => x.id === pinFor);
    if (u && !u.pin) return pageHead("Control Room", `Hello ${esc(u.name)}`, "", "", "The admin reset your PIN. Choose a new 4-digit PIN and type it twice.")
      + `<form id="fNewPin" class="card" style="max-width:360px" autocomplete="off"><input name="pin" class="pin" inputmode="numeric" maxlength="4" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" autofocus required placeholder="New PIN">
        <input name="pin2" class="pin" inputmode="numeric" maxlength="4" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc;margin-top:10px" required placeholder="Again">
        <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Save and start</button><button type="button" class="ghost" id="cancelPin">Not me</button></div><p class="flag" id="pinErr"></p></form>`;
    return pageHead("Control Room", `Hello ${esc(u?.name || "")}`, "", "", "Type your PIN.")
      + `<form id="fPin" class="card" style="max-width:360px" autocomplete="off"><input name="pin" class="pin" inputmode="numeric" maxlength="4" type="text" autocomplete="one-time-code" style="-webkit-text-security:disc" value="" autofocus required>
        <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Start</button><button type="button" class="ghost" id="cancelPin">Not me</button></div><p class="flag" id="pinErr"></p></form>`; }
  return pageHead("Control Room", "Who is on shift?", "", `<button class="primary" id="addMe">+ Ask to be a controller</button>`, "Tap your name. New controller? Ask to be added – the admin approves it.")
    + `<div class="who-pick">${users.map(u => okUser(u) ? `<button data-who="${u.id}"><b>${esc(u.name)}</b><small>added ${fmt(ymd(new Date(u.addedAt || u.editedAt)))}</small></button>` : `<button disabled style="opacity:.6"><b>${esc(u.name)}</b><small>waiting for the admin to approve</small></button>`).join("")}</div>`
    + `<p style="margin-top:22px"><button class="ghost sm" id="adminIn">Admin</button></p>`;
}

/* ---------- views ---------- */
const TABS = [["now", "Live board and workbook", "deploy"], ["files", "Documents", "docs"], ["inv", "Invoicing", "money"], ["export", "Export & sync", "import"], ["settings", "Settings", "settings"]];
const pendingReqs = () => R("req").filter(q => q.status === "pending").length + R("user").filter(u => u.status === "pending").length;
const tabs = () => { const n = pendingReqs(); return [TABS[0], TABS[1], ["req", `Change requests${n ? ` (${n})` : ""}`, "missing"], ...TABS.slice(2)]; };
function render() {
  $("who").innerHTML = me ? `${esc(me.name)} · <button class="sm ghost" data-logout="swap" style="min-height:36px;padding:2px 12px">Change controller</button><button class="sm danger" data-logout style="min-height:36px;padding:2px 12px">Log out</button>` : "";
  if (!me && viewer) { $("who").innerHTML = `Viewing only · <button class="sm" data-unview="1" style="min-height:36px;padding:2px 12px">Controller sign-in</button>`;
    const ws = app.querySelector(".xlsheet")?.scrollTop, wl = app.querySelector(".xlsheet")?.scrollLeft; if (tab !== "files") tab = "now";
    app.innerHTML = tabBar([["now", "Dashboard", "deploy"], ["files", "Documents", "docs"]], tab) + `<div class="ro">${tab === "files" ? filesView() : nowView()}</div>`; afterRender(app);
    const w = app.querySelector(".xlsheet"); if (w && ws != null) { w.scrollTop = ws; w.scrollLeft = wl; } return; }
  if (!me) { app.innerHTML = `<div class="card viewcard"><b>Just want to see the machines?</b><span class="muted">Anyone can view the live board and the hours. Only controllers can change them.</span><button class="primary" data-view="1">View the dashboard</button></div>` + gateView(); return; }
  let body = "";
  if (tab === "now") body = nowView(); if (tab === "avail") body = availView(); if (tab === "docs") body = docsView(); if (tab === "files") body = filesView();
  if (askMail && me && !me.admin) body = mailToView();
  else if (tab === "req") body = reqView();
  if (tab === "inv") body = invView(); if (tab === "export") body = exportView(); if (tab === "settings") body = settingsView();
  const ws = app.querySelector(".xlsheet")?.scrollTop, wl = app.querySelector(".xlsheet")?.scrollLeft;
  app.innerHTML = tabBar(tabs(), tab) + body; afterRender(app);
  const w = app.querySelector(".xlsheet"); if (w && ws != null) { w.scrollTop = ws; w.scrollLeft = wl; }
}
let viewer = false;
const dateBar = () => `<div class="dnav"><button data-dnav="-1" aria-label="Day before">‹</button><input type="date" id="dPick" value="${date}" max="${today()}"><button data-dnav="1" aria-label="Next day" ${date >= today() ? "disabled" : ""}>›</button>${date !== today() ? '<button data-today class="sm">Today</button>' : ""}</div>`;
function statusBar() {
  const logs = R("log").filter(l => l.date === today()).sort((a, b) => b.at - a.at), last = logs[0];
  const dueAt = last ? last.at + 36e5 : 0, due = !last || Date.now() > dueAt, w = waiting();
  return `<div class="statusbar"><span>Last update <b>${last ? `${new Date(last.at).toTimeString().slice(0, 5)}</b> by ${esc(last.by)}` : "–</b> today"}</span>
    <span>${last ? `Next update by <b>${new Date(dueAt).toTimeString().slice(0, 5)}</b> ` : ""}<span class="${due ? "due" : "ok"}">${due ? "Update due" : "Up to date"}</span></span>
    <span>${w ? `${w} change${w === 1 ? "" : "s"} waiting for signal` : (sync.last ? `Sent to the TKR app ${new Date(sync.last).toTimeString().slice(0, 5)}` : "Saved on this PC")}</span>
    <button class="danger" data-logout style="margin-left:auto;min-height:44px">Log out · end shift</button></div>`;
}
/* Breakdowns still open from an earlier shift: the new controller confirms each one at the start of the shift */
function curShift() { const n = new Date(Date.now()), h = n.getHours(), d = ymd(n);
  if (h >= 6 && h < 18) return { key: d + "|Day", start: toMs(d + "T06:00"), label: "06:00" };
  const dd = h >= 18 ? d : addDays(d, -1); return { key: dd + "|Night", start: toMs(dd + "T18:00"), label: "18:00" }; }
/* Clean slate: every new shift starts with all machines working. Anything still open from the shift before is closed
   at the shift change (06:00 or 18:00). If a machine is still down, the new controller logs it again for the new shift. */
let closing = false;
async function closeCarried() {
  if (closing) return 0; closing = true;
  try { const c = curShift(), L = evs().filter(e => !e.end && toMs(e.start) < c.start);
    for (const e of L) await put({ ...e, end: isoOf(c.start), endBy: "Shift change" }, `${e.machine}: ${TYPES[e.type].toLowerCase()} closed at ${c.label}, end of the shift`);
    return L.length; } finally { closing = false; }
}
/* the shift that is running now, and its totals for the whole fleet */
function shiftNow() { const c = curShift(), [d, sh] = c.key.split("|"); return { d, sh, win: shiftWin(d, sh) }; }
const pcOf = t => t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 1000) / 10 + "%" : "–";
function fleetSum(L, fn) { return L.reduce((a, m) => { const x = fn(m); if (x) for (const k of ["avail", "BD", "PM", "down"]) a[k] += x[k] || 0; return a; }, { avail: 0, BD: 0, PM: 0, down: 0 }); }
/* breakdowns and maintenance in the running shift: logged times, plus hours picked in the sheet without a time */
function shiftEvents() {
  const { d, sh, win: [a, b] } = shiftNow(), now = Date.now(), out = evs().filter(e => toMs(e.start) < Math.min(b, now) && evEnd(e) > a).sort((x, y) => toMs(y.start) - toMs(x.start)).map(e => ({ e }));
  for (const m of fleet().flatMap(([, L]) => L)) { const C = cellsOf(d, m);
    for (const t of ["BD", "PM"]) { const hrs = HRS(sh).filter(h => C[h] === t && hStart(d, h) <= now && !evs().some(e => e.machine === m && e.type === t && toMs(e.start) < hStart(d, h) + 36e5 && evEnd(e) > hStart(d, h)));
      if (hrs.length) out.push({ sheet: { machine: m, type: t, hrs } }); } }
  return out;
}
function nowView() {
  const F = fleet(), ALL = F.flatMap(([, L]) => L).filter(m => !offNow(m)), nOff = F.flatMap(([, L]) => L).length - ALL.length, open = evs().filter(e => !e.end).sort((a, b) => toMs(a.start) - toMs(b.start));
  const { d: sd, sh: ssh } = shiftNow(), shT = fleetSum(ALL, m => shiftSum(m, sd, ssh)), dayT = fleetSum(ALL, m => daySum(m, sd));
  const SE = shiftEvents(), logs = R("log").filter(l => l.date === today()).sort((a, b) => b.at - a.at);
  const cnt = k => ALL.filter(m => stateNow(m) === k).length;
  const seLine = x => x.e ? evLine(x.e) : `<div class="evrow"><span class="tag ${x.sheet.type}">${x.sheet.type}</span><div><b>${esc(x.sheet.machine)}</b> · ${TYPES[x.sheet.type]}<br><span class="t">${x.sheet.hrs.map(hLabel).join(", ")} · ${x.sheet.hrs.length} h, picked in the sheet</span></div></div>`;
  return pageHead("Control room · live", "Machine status", new Date(Date.now()).toLocaleDateString("en-ZA", { weekday: "long", day: "numeric", month: "long" }) + ` · ${ssh} shift`, `<button class="primary" data-log="">+ Log breakdown</button> <button data-upload="">Upload documents</button> <button data-tab="files">Documents</button>`, "Tap a machine to log a breakdown, standby, maintenance or not in use, with the exact time. Tap it again when it is back.")
    + statusBar()
    + `<div class="tiles"><div class="tile good"><div class="l">Availability this shift</div><div class="n">${pcOf(shT)}</div><div class="s">${hm(shT.avail)} available · ${hm(shT.down)} down so far (hours:minutes)</div></div>
      <div class="tile"><div class="l">Availability today</div><div class="n">${pcOf(dayT)}</div><div class="s">from 06:00 ${fmt(sd).slice(0, 5)} · ${hm(dayT.avail)} available · ${hm(dayT.down)} down</div></div>
      <div class="tile good"><div class="l">Operating now</div><div class="n">${cnt("OP")}</div><div class="s">of ${ALL.length} machines</div></div>
      <div class="tile"><div class="l">Available, standing</div><div class="n">${cnt("SB")}</div><div class="s">available but not operating</div></div>
      <div class="tile ${cnt("BD") ? "alert" : ""}"><div class="l">Breakdowns now</div><div class="n">${cnt("BD")}</div></div>
      <div class="tile"><div class="l">In maintenance now</div><div class="n">${cnt("PM")}</div></div>
      ${nOff ? `<div class="tile"><div class="l">Not on site</div><div class="n">${nOff}</div><div class="s">left out of availability</div></div>` : ""}
      ${cnt("NU") ? `<div class="tile"><div class="l">Not in use now</div><div class="n">${cnt("NU")}</div></div>` : ""}
      ${cnt("?") ? `<div class="tile alert"><div class="l">Not updated this shift</div><div class="n">${cnt("?")}</div><div class="s">pick OP, BD or PM in the sheet</div></div>` : ""}</div>`
    + `<div class="fleet">${F.map(([g, L]) => { const gT = fleetSum(L, m => shiftSum(m, sd, ssh)); return `<section class="grp"><h3>${esc(g)}<span>${L.every(offNow) ? "not on site" : `${L.filter(m => stateNow(m) === "OP").length}/${L.filter(m => !offNow(m)).length} operating${L.some(m => stateNow(m) === "SB") ? ` · ${L.filter(m => stateNow(m) === "SB").length} standing` : ""} · ${pcOf(gT)} available this shift`}</span></h3><div class="mlist">${L.map(m => { const s0 = statusNow(m), st = stateNow(m), s = s0 && s0.type === st ? s0 : null; return `<button class="mach ${s ? s.type : st === "?" ? "UN" : st}" data-mach="${esc(m)}"><b>${esc(m)}</b><small>${st === "OFF" ? "Not on site" : s ? `${TYPES[s.type]} since ${hhmm(s.start)}${s.start.slice(0, 10) !== today() ? " " + fmt(s.start.slice(0, 10)).slice(0, 5) : ""}` : st === "OP" ? "Operating" : st === "SB" ? "Available, not operating" : st === "?" ? "Not updated this shift" : (TYPES[st] || st) + (sheetSince(m, st) ? " since " + sheetSince(m, st) : "")}</small>${st !== "OP" && st !== "?" && reasonNow(m) ? `<small class="why">${esc(reasonNow(m))}</small>` : ""}</button>`; }).join("")}</div>
        <div class="row" style="margin-top:8px;gap:6px"><button class="sm" data-addmg="${F.findIndex(x => x[0] === g)}">+ Add machine</button>${L.length ? `<button class="sm ghost" data-rmmg="${F.findIndex(x => x[0] === g)}">− Remove machine</button><button class="sm ghost" data-sitemg="${F.findIndex(x => x[0] === g)}">On site / not on site</button>` : ""}</div></section>`; }).join("")}</div>`
    + hourBoard()
    + `<div class="openers">${sheet("open", `Still down now (${open.length})`, open.map(evLine).join("") || '<p class="muted">Nothing open with a logged time.</p>', { icon: "!", count: open.length, cls: open.length ? "alert" : "good", sub: "breakdowns, standby and maintenance not closed yet" })}
      ${sheet("todayev", "This shift's breakdowns and maintenance", `<div class="res">${SE.map(seLine).join("") || '<p class="muted">Nothing on this shift yet.</p>'}</div>`, { icon: icon("bd"), count: SE.length, sub: `${ssh.toLowerCase()} shift from ${ssh === "Day" ? "06:00" : "18:00"}, from the sheet and the logged times` })}
      ${sheet("delays", `Delay summary · ${monthName(today().slice(0, 7))}`, delaySummary(today().slice(0, 7)), { icon: icon("bd"), sub: "lost time per machine and per cause this month" })}
      ${sheet("checks", "Data and documents checklist", checklist(today()), { icon: icon("docs"), sub: `${chkDone(today())} of ${CHECK.flatMap(([, i]) => i).length * COLS.length} ticked today` })}
      ${sheet("ulog", "Update log", `<div class="res">${logs.map(l => `<div class="ri"><div class="rt"><b>${new Date(l.at).toTimeString().slice(0, 5)} · ${esc(l.by)}</b><small>${esc(l.what)}</small></div></div>`).join("") || '<p class="muted">No updates yet today.</p>'}</div>`, { icon: icon("hours"), count: logs.length, sub: "every change, with the time and who did it" })}</div>`
    + workbookSection();
}
/* this shift hour by hour: one big block per hour per machine, easy to read on a phone */
function hourBoard() {
  const { d, sh } = shiftNow(), H = HRS(sh), now = Date.now(), lab = h => pad(h % 24);
  const cell = (mc, h, prev) => { const a = hStart(d, h), v = a <= now ? cellVal(d, mc, h) : "", cur = a <= now && now < a + 36e5;
    if (!v || v === "OFF") return `<div class="hb ${v || "fut"}${cur ? " now" : ""}"><i>${lab(h)}</i><b>${v || ""}</b></div>`;
    const q = hourParts(mc, d, h, v, prev), dn = q.BD + q.PM, t = dn >= q.len - 1 / 120 ? (q.BD >= q.PM ? "BD" : "PM") : dn > 0 ? "MIX" : v === "SB" ? "SB" : "OP";
    const pct = Math.round(dn / 1 * 100), dt = q.BD >= q.PM ? "BD" : "PM", lbl = t === "MIX" ? `${dt} ${hm(dn)}` : t === "OP" || t === "SB" ? t : dn < 1 - 1 / 120 && cur ? `${dt} ${hm(dn)}` : t;
    const base = v === "SB" ? "#4d4116" : "#1f4d33", dc = dt === "BD" ? "#9b2c22" : "#2a4f8a", pas = Math.round(q.len * 100), up = cur ? pas : 100;   // filled as far as the hour has gone: down part first, then working
    const fill = dn > 0 && pct >= up ? dc : base, bg = cur || t === "MIX" ? `style="background:linear-gradient(to right,${dc} 0 ${Math.min(pct, up)}%,${fill === dc ? dc : base} ${Math.min(pct, up)}% ${up}%,#202220 ${up}% 100%)"` : "";
    return `<div class="hb ${t}${t === "MIX" ? " " + dt : ""}${cur ? " now" : ""}" ${bg} title="${lab(h)}:00–${lab(h + 1)}:00 · ${dn ? `${dt === "BD" ? "breakdown" : "maintenance"} ${hm(dn)} of ${hm(q.len)}` : v === "SB" ? "available, not operating" : "operating"}"><i>${lab(h)}</i><b>${lbl}</b></div>`; };
  return `<section class="hboard"><div class="hbhead"><h2 class="sect" style="margin:0">This shift hour by hour</h2><span class="muted">${sh} shift ${sh === "Day" ? "06:00 – 18:00" : "18:00 – 06:00"} · ${fmt(d)}</span>
      <span class="hbkey"><span class="hb OP"><b>OP</b></span> Operating <span class="hb SB"><b>SB</b></span> Available, standing <span class="hb BD"><b>BD</b></span> Breakdown <span class="hb PM"><b>PM</b></span> Planned maintenance <span class="hb OFF"><b>OFF</b></span> Not on site</span></div>
    ${fleet().map(([g, L]) => `<div class="hbgrp"><h3>${esc(g)}</h3>${L.map(mc => { const st = stateNow(mc), why = st !== "OP" ? reasonNow(mc) : "";
      return `<div class="hbrow"><div class="hbm"><b>${esc(mc)}</b><small class="${st}">${st === "OP" ? "Operating" : st === "SB" ? "Available, not operating" : st === "OFF" ? "Not on site" : TYPES[st] || st}</small>${why ? `<small class="why">${esc(why)}</small>` : ""}</div><div class="hbcells">${(() => { let prev = ""; return H.map(h => { const c = cell(mc, h, prev); prev = hStart(d, h) <= now ? cellVal(d, mc, h) : ""; return c; }).join(""); })()}</div></div>`; }).join("")}</div>`).join("")}</section>`;
}
function evLine(e) {
  const dur = (evEnd(e) - toMs(e.start)) / 36e5;
  return `<div class="evrow"><span class="tag ${e.type}">${e.type}</span><div><b>${esc(e.machine)}</b> · ${TYPES[e.type]}<br><span class="t">${fmt(e.start.slice(0, 10)).slice(0, 5)} ${hhmm(e.start)} → ${e.end ? `${e.end.slice(0, 10) !== e.start.slice(0, 10) ? fmt(e.end.slice(0, 10)).slice(0, 5) + " " : ""}${hhmm(e.end)}` : "still going"} · ${hm(dur)} h</span>${e.reason ? `<br><small class="muted">${esc(e.reason)}</small>` : ""}<br><small class="muted">by ${esc(e.by)}${e.endBy && e.endBy !== e.by ? `, closed by ${esc(e.endBy)}` : ""}</small></div>
    <div class="row" style="flex-direction:column;gap:6px">${e.end ? "" : `<button class="sm primary" data-up="${e.id}">Back up now</button>`}<button class="sm" data-edit="${e.id}">Change</button><button class="sm ghost" data-mail="${e.id}">Email</button></div></div>`;
}
const chkKey = (d, item) => `chk_${d}_${item.replace(/\W+/g, "")}`;
const chkDone = d => R("chk").filter(c => c.date === d).reduce((n, c) => n + COLS.filter(k => c[k] === "✓").length, 0);
function checklist(d) {
  const get = item => recs.find(r => r.id === chkKey(d, item)) || {};
  return `<p class="muted" style="margin-top:0">Tap a box: empty → ✓ received → ✗ not received → empty. The time and your name are saved with it.</p><div class="tbl card" style="padding:0"><table class="chk"><thead><tr><th>Section</th><th>Item</th>${COLS.map(c => `<th>${c}</th>`).join("")}<th>Updated</th><th>Comment</th></tr></thead><tbody>
    ${CHECK.flatMap(([sec, items]) => items.map((it, i) => { const c = get(it); return `<tr><td>${i ? "" : `<b>${esc(sec)}</b>`}</td><td>${esc(it)}</td>${COLS.map(k => `<td><button class="tick ${c[k] === "✓" ? "y" : c[k] === "✗" ? "n" : ""}${lockedV(c, k) ? ` lk" data-lockf="${esc(c.id)}|${k}` : `" data-chk="${esc(d)}|${esc(it)}|${k}`}">${c[k] || " "}</button></td>`).join("")}
      <td class="muted" style="white-space:nowrap">${c.editedAt ? `${new Date(c.editedAt).toTimeString().slice(0, 5)} ${esc(c.by)}` : ""}</td><td>${lockedV(c, "comment") ? `<input class="lk" readonly data-lockf="${esc(c.id)}|comment" value="${esc(c.comment)}" style="min-width:160px">` : `<input data-chkc="${esc(d)}|${esc(it)}" value="${esc(c.comment || "")}" placeholder="…" style="min-width:160px">`}</td></tr>`; })).join("")}</tbody></table></div>`;
}
function availView() {
  const F = fleet(), d = date, rowsFor = L => L.map(m => [m, daySum(m, d)]);
  const pc = v => v == null ? "–" : Math.round(v * 100) + "%";
  return pageHead("Machine availability", fmt(d), dow(d), "", "From the logged times: a day runs 06:00 to 06:00. Standby counts as available; not in use is left out.")
    + `<div class="mhead">${dateBar()}</div>`
    + F.map(([g, L]) => { const rows = rowsFor(L), tot = rows.reduce((a, [, s]) => { if (s) for (const k of ["avail", "BD", "PM", "SB", "NU", "down"]) a[k] = (a[k] || 0) + s[k]; return a; }, {});
      return `<section class="card"><h2>${esc(g)}</h2><div class="tbl"><table class="avt"><thead><tr><th>Machine</th><th>Available h</th><th>Breakdown h</th><th>Planned maint. h</th><th>Standby h</th><th>Not in use h</th><th>Availability</th></tr></thead><tbody>
        ${rows.map(([m, s]) => `<tr><td><b>${esc(m)}</b></td>${s ? `<td class="num">${h2(s.avail)}</td><td class="num ${s.BD ? "hot" : ""}">${h2(s.BD)}</td><td class="num">${h2(s.PM)}</td><td class="num">${h2(s.SB)}</td><td class="num">${h2(s.NU)}</td><td class="num pct">${pc(s.pct)}</td>` : '<td colspan="6" class="muted">not started</td>'}</tr>`).join("")}
        <tr class="tot"><td>Total</td><td class="num">${h2(tot.avail || 0)}</td><td class="num">${h2(tot.BD || 0)}</td><td class="num">${h2(tot.PM || 0)}</td><td class="num">${h2(tot.SB || 0)}</td><td class="num">${h2(tot.NU || 0)}</td><td class="num pct">${pc((tot.avail || 0) + (tot.down || 0) ? tot.avail / (tot.avail + tot.down) : null)}</td></tr></tbody></table></div></section>`; }).join("")
    + delaySummary(d.slice(0, 7))
    + `<div class="openers">${sheet("month", `Month: ${monthName(d.slice(0, 7))}`, monthGrid(d.slice(0, 7)), { icon: icon("shifts"), sub: "available hours per machine per day, with month totals" })}</div>`;
}
/* delay summary for the month, per machine and per cause (like the Delay summary sheet in the metres workbook) */
const CAUSES = [["Waiting for GET / drill bits", /\bget\b|\bbits?\b|drill bit/], ["Stuck rods / retrieving rods", /stuck|retriev|recover/], ["Service and maintenance", /service|maint|250\s*hr|weekly|daily/],
  ["Tyres", /tyre|tire|puncture|flat/], ["Engine", /engine|overheat|starter|stater|battery|alternator/], ["Hydraulics", /hydr|hose|pipe|burst|leak|oil|cylinder/], ["Electrical", /electric|wiring|fuse|light|radio|power|sensor/],
  ["Brakes and steering", /brake|steer/], ["Compressor", /compressor|air/], ["Diesel / refuelling", /diesel|refuel|refill|fuel/], ["Waiting for parts or mechanic", /waiting|parts|spares|mechanic|fitter|boilermaker/],
  ["Tramming", /tram/], ["Blasting", /blast/], ["No operator", /operator|driver/]];
function causeOf(e) { const t = String(e.reason || "").toLowerCase().trim(); if (e.type === "PM" && !t) return "Planned maintenance"; if (!t) return "No reason written";
  for (const [n, rx] of CAUSES) if (rx.test(t)) return n; return e.reason.trim().replace(/^./, c => c.toUpperCase()); }
/* metres and tally-sheet delays per drill rig, shared by the Metres section (same figures as its Delay summary) */
const rigKey = r => String(r || "").replace(/\s+/g, "").toUpperCase();
function mtrSlots(from, to) { const S = new Map();
  for (const r of R("mtr")) { if (r.ym < from.slice(0, 7) || r.ym > to.slice(0, 7)) continue; let L = []; try { L = JSON.parse(r.data || "[]"); } catch {}
    for (const x of L) if (x.date >= from && x.date <= to) S.set(`${x.date}|${x.shift}|${x.rig}`, x); }
  return S; }
function dsData(from, to) {
  const a = toMs(from + "T06:00"), b = Math.min(toMs(addDays(to, 1) + "T06:00"), Date.now()), ALL = fleet().flatMap(([, L]) => L);
  const days = []; for (let x = from; x <= to && x <= today(); x = addDays(x, 1)) days.push(x);
  const per = ALL.map(mc => { const t = days.reduce((o, x) => { const z = daySum(mc, x); if (z) for (const k of ["avail", "BD", "PM", "down"]) o[k] += z[k]; return o; }, { avail: 0, BD: 0, PM: 0, down: 0 });
    const E = evs().filter(e => e.machine === mc && (e.type === "BD" || e.type === "PM") && toMs(e.start) < b && evEnd(e) > a);
    return { mc, g: groupOf(mc), ...t, nBD: E.filter(e => e.type === "BD").length, nPM: E.filter(e => e.type === "PM").length }; });
  const lines = evs().filter(e => (e.type === "BD" || e.type === "PM") && toMs(e.start) < b && evEnd(e) > a).map(e => ({ e, h: (Math.min(evEnd(e), b) - Math.max(toMs(e.start), a)) / 36e5, cause: causeOf(e) }))
    .sort((x, y) => toMs(x.e.start) - toMs(y.e.start));
  const C = new Map(); for (const l of lines) { const c = C.get(l.cause) || { cause: l.cause, h: 0, n: 0, M: new Map() }; c.h += l.h; c.n++; c.M.set(l.e.machine, (c.M.get(l.e.machine) || 0) + l.h); C.set(l.cause, c); }
  const T = per.reduce((o, x) => { for (const k of ["avail", "BD", "PM", "down", "nBD", "nPM"]) o[k] += x[k]; return o; }, { avail: 0, BD: 0, PM: 0, down: 0, nBD: 0, nPM: 0 });
  const tdr = ALL.filter(mc => /^TDR/i.test(mc)), slots = mtrSlots(from, to);
  const B = bdData(slots, from, to, [...new Set([...slots.values()].map(x => x.rig))].sort((x, y) => tdr.findIndex(t => rigKey(t) === rigKey(x)) - tdr.findIndex(t => rigKey(t) === rigKey(y))));
  return { from, to: days.at(-1) || from, days, per, lines, causes: [...C.values()].sort((x, y) => y.h - x.h), T, B, avOf: rig => per.find(p => rigKey(p.mc) === rigKey(rig)) };
}
/* the graphs, the same on screen and in the Excel */
function dsCharts(D) {
  const out = [], per = D.per.filter(x => x.avail + x.down > 0), B = D.B, col = ["#d4a12a", "#111111", "#2e7d4f"];
  if (per.length) {
    out.push(chartPng("Availability per machine (%)", per.map(x => x.mc), [{ name: "Availability %", color: "#2e7d4f", vals: per.map(x => Math.round(x.avail / (x.avail + x.down) * 1000) / 10) }], { unit: "%" }));
    const dn = per.filter(x => x.BD + x.PM > 0);
    if (dn.length) out.push(chartPng("Hours down per machine (Control Room)", dn.map(x => x.mc), [{ name: "Breakdown", color: "#b3261e", vals: dn.map(x => Math.round(x.BD * 60) / 60) }, { name: "Planned maintenance", color: "#1f4e78", vals: dn.map(x => Math.round(x.PM * 60) / 60) }], { unit: " h" })); }
  if (D.causes.length) out.push(chartPng("Hours down per cause (Control Room)", D.causes.map(c => c.cause), [{ name: "Hours", color: "#d4a12a", vals: D.causes.map(c => Math.round(c.h * 60) / 60) }], { unit: " h" }));
  if (B.cats.length) out.push(chartPng("Drill rigs: hours lost per cause (shift reports)", B.cats.map(c => c.cat), B.rigs.map((x, k) => ({ name: x, color: col[k % 3], vals: B.cats.map(c => c.per[k]) })), { unit: " h" }));
  if (B.mach.length) { out.push(chartPng("Drill rigs: metres drilled", B.mach.map(m => m.rig), [{ name: "Metres", color: "#d4a12a", vals: B.mach.map(m => m.metres) }], { unit: " m" }));
    out.push(chartPng("Drill rigs: shifts working and not working", B.mach.map(m => m.rig), [{ name: "Working", color: "#2e7d4f", vals: B.mach.map(m => m.working) }, { name: "Not working", color: "#b3261e", vals: B.mach.map(m => m.down) }], { dec: 0 })); }
  return out;
}
const pcOfT = t => t && t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 1000) / 10 + "%" : "–";
function delaySummary(m) {
  const D = dsData(m + "-01", monthEnd(m)), { per, causes, lines, T, B } = D, h1 = v => hm(v || 0);
  const drill = B.all.length ? `<h3 style="margin:16px 0 6px">Drill rigs · metres and delays from the Metres section</h3>
    <div class="tbl"><table class="avt"><thead><tr><th>Machine</th><th>Shifts working</th><th>Shifts not working</th><th>Metres</th><th>Delay time (shift reports)</th><th>Available (Control Room)</th><th>Availability</th></tr></thead><tbody>
      ${B.mach.map(x => { const av = D.avOf(x.rig); return `<tr><td><b>${esc(x.rig)}</b></td><td class="num">${x.working}</td><td class="num">${x.down}</td><td class="num">${(+x.metres || 0).toFixed(1)}</td><td class="num">${h1(x.hours)}</td><td class="num">${av ? hm(av.avail) : "–"}</td><td class="num pct">${pcOfT(av)}</td></tr>`; }).join("")}</tbody></table></div>
    <div class="tbl" style="margin-top:10px"><table class="avt"><thead><tr><th>Delay</th>${B.rigs.map(r => `<th>${esc(r)} time</th>`).join("")}<th>Total time</th>${B.rigs.map(r => `<th>Times recorded ${esc(r)}</th>`).join("")}</tr></thead><tbody>
      ${B.cats.map(c => `<tr><td>${esc(c.cat)}</td>${c.per.map(v => `<td class="num">${h1(v)}</td>`).join("")}<td class="num"><b>${h1(c.total)}</b></td>${c.times.map(v => `<td class="num">${v}</td>`).join("")}</tr>`).join("")}
      <tr class="tot"><td>Total</td>${B.rigs.map((r, i) => `<td class="num">${h1(B.cats.reduce((n, c) => n + c.per[i], 0))}</td>`).join("")}<td class="num">${h1(B.total)}</td>${B.rigs.map((r, i) => `<td class="num">${B.cats.reduce((n, c) => n + c.times[i], 0)}</td>`).join("")}</tr></tbody></table></div>
    <details style="margin-top:8px"><summary style="cursor:pointer;font-weight:600">Every delay on the shift reports (${B.lines.length})</summary><div class="tbl"><table class="avt"><thead><tr><th>Delay (as written on the shift report)</th><th>Date</th><th>Shift</th><th>Machine</th><th>Category</th><th>Time</th></tr></thead><tbody>
      ${B.lines.map(l => `<tr><td>${esc(l.text)}</td><td>${fmt(l.date)}</td><td>${l.shift}</td><td>${esc(l.rig)}</td><td>${esc(l.cat || "")}</td><td class="num">${l.hours == null ? "" : h1(l.hours)}</td></tr>`).join("")}</tbody></table></div></details>`
    : `<p class="muted" style="margin-top:14px">Drill rig metres and shift-report delays show here once the Metres section has synced this month.</p>`;
  return `<section class="card"><h2>Delay summary · ${monthName(m)}</h2><p class="muted" style="margin-top:0">${fmt(D.from)} to ${fmt(D.to)} · times are hours:minutes (0:58 = 58 minutes).</p>
    <div class="tbl"><table class="avt"><thead><tr><th>Group</th><th>Machine</th><th>Available</th><th>Breakdown</th><th>Planned maint.</th><th>Down</th><th>Availability</th><th>Times broken down</th><th>Times in maintenance</th></tr></thead><tbody>
      ${per.map(x => `<tr><td class="muted">${esc(x.g)}</td><td><b>${esc(x.mc)}</b></td><td class="num">${hm(x.avail)}</td><td class="num ${x.BD ? "hot" : ""}">${hm(x.BD)}</td><td class="num">${hm(x.PM)}</td><td class="num">${hm(x.down)}</td><td class="num pct">${pcOfT(x)}</td><td class="num">${x.nBD}</td><td class="num">${x.nPM}</td></tr>`).join("")}
      <tr class="tot"><td></td><td>Total</td><td class="num">${hm(T.avail)}</td><td class="num">${hm(T.BD)}</td><td class="num">${hm(T.PM)}</td><td class="num">${hm(T.down)}</td><td class="num pct">${pcOfT(T)}</td><td class="num">${T.nBD}</td><td class="num">${T.nPM}</td></tr></tbody></table></div>
    <div class="tbl" style="margin-top:12px"><table class="avt"><thead><tr><th>Cause</th><th>Time</th><th>Times recorded</th><th>Machines</th></tr></thead><tbody>
      ${causes.map(c => `<tr><td>${esc(c.cause)}</td><td class="num">${hm(c.h)}</td><td class="num">${c.n}</td><td>${[...c.M].sort((x, y) => y[1] - x[1]).map(([k, v]) => `${esc(k)} ${hm(v)}`).join(" · ")}</td></tr>`).join("") || '<tr><td colspan="4" class="muted">No breakdowns or maintenance logged this month.</td></tr>'}</tbody></table></div>
    <details style="margin-top:10px"><summary style="cursor:pointer;font-weight:600">Every breakdown and maintenance, one by one (${lines.length})</summary><div class="tbl"><table class="avt"><thead><tr><th>Date</th><th>Machine</th><th>Type</th><th>From</th><th>To</th><th>Time</th><th>Cause</th><th>Reason / comments</th></tr></thead><tbody>
      ${lines.map(l => `<tr><td>${fmt(l.e.start.slice(0, 10))}</td><td><b>${esc(l.e.machine)}</b></td><td>${TYPES[l.e.type]}</td><td>${hhmm(l.e.start)}</td><td>${l.e.end ? (l.e.end.slice(0, 10) !== l.e.start.slice(0, 10) ? fmt(l.e.end.slice(0, 10)).slice(0, 5) + " " : "") + hhmm(l.e.end) : "still going"}</td><td class="num">${hm(l.h)}</td><td>${esc(l.cause)}</td><td>${esc(l.e.reason || "")}</td></tr>`).join("")}</tbody></table></div></details>
    ${drill}
    <details style="margin-top:12px" open><summary style="cursor:pointer;font-weight:600">Graphs</summary><div style="display:grid;gap:12px;margin-top:8px">${(() => { try { return dsCharts(D).map(c => `<img src="${c.png}" alt="" style="width:100%;max-width:900px;height:auto;border:1px solid #e6e0d2;border-radius:6px">`).join(""); } catch { return ""; } })()}</div></details></section>`;
}
function monthGrid(m) {
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  return `<div class="card grid30" style="padding:0"><table><thead><tr><th>Machine</th>${days.map(x => `<th>${+x.slice(8)}</th>`).join("")}<th>Avail h</th><th>BD h</th><th>PM h</th><th>SB h</th><th>Avail %</th></tr></thead><tbody>
    ${fleet().map(([g, L]) => `<tr><td colspan="${days.length + 6}" style="color:var(--gold);font-weight:700">${esc(g)}</td></tr>` + L.map(mc => { const S = days.map(x => daySum(mc, x)), t = S.reduce((a, s) => { if (s) for (const k of ["avail", "BD", "PM", "SB", "down"]) a[k] += s[k]; return a; }, { avail: 0, BD: 0, PM: 0, SB: 0, down: 0 });
      return `<tr><td><b>${esc(mc)}</b></td>${S.map(s => `<td class="${s && s.BD ? "hot" : ""}">${s ? hm(s.avail) : ""}</td>`).join("")}<td>${h2(t.avail)}</td><td>${h2(t.BD)}</td><td>${h2(t.PM)}</td><td>${h2(t.SB)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : "–"}</b></td></tr>`; }).join("")).join("")}</tbody></table></div>`;
}
function docsView() {
  return pageHead("Data and documents", fmt(date), dow(date), "", "The daily checklist. It saves as you tap.") + `<div class="mhead">${dateBar()}</div>` + checklist(date);
}
const invKey = (m, svc, doc) => `inv_${m}_${(svc + "_" + doc).replace(/\W+/g, "")}`;
const addKey = (m, item) => `add_${m}_${item.replace(/\W+/g, "")}`;
function invView() {
  const m = month, mb = `<div class="mbar"><button data-mnav="-1">‹</button><b>${monthName(m)}</b><button data-mnav="1" ${m >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
  const rec = id => recs.find(r => r.id === id) || {};
  const sup = SUPPORT.map(([svc, docs, person]) => { const D = docs.map(dc => ({ dc, r: rec(invKey(m, svc, dc)) })), ready = D.every(x => x.r.done === "✓");
    return `<tr class="${ready ? "" : ""}"><td><b>${esc(svc)}</b><br><small class="muted">${esc(rec(invKey(m, svc, "_person")).person || person)}</small></td><td>${D.map(x => `<div class="row" style="gap:8px;margin:3px 0"><button class="tick ${x.r.done === "✓" ? "y" : x.r.done === "✗" ? "n" : ""}${lockedV(x.r, "done") ? ` lk" data-lockf="${esc(x.r.id)}|done` : `" data-inv="${esc(svc)}|${esc(x.dc)}`}">${x.r.done || " "}</button>${esc(x.dc)}</div>`).join("")}</td><td>${ready ? '<span class="tag OP">READY</span>' : ""}</td></tr>`; }).join("");
  const add = ADDITIONAL.map(([it, person, note]) => { const r = rec(addKey(m, it));
    return `<tr><td>${esc(it)}${note ? `<br><small class="muted">${esc(note)}</small>` : ""}</td><td><button class="tick ${r.done === "✓" ? "y" : r.done === "✗" ? "n" : ""}${lockedV(r, "done") ? ` lk" data-lockf="${esc(r.id)}|done` : `" data-add="${esc(it)}`}">${r.done || " "}</button></td><td>${esc(r.person || person)}</td>${["collector", "notes"].map(k => `<td>${lockedV(r, k) ? `<input class="lk" readonly data-lockf="${esc(r.id)}|${k}" value="${esc(r[k])}">` : `<input data-addc="${esc(it)}|${k}" value="${esc(r[k] || "")}" placeholder="${k === "collector" ? "name" : "…"}">`}</td>`).join("")}</tr>`; }).join("");
  const nDone = ADDITIONAL.filter(([it]) => rec(addKey(m, it)).done === "✓").length;
  return pageHead("Invoicing", "Supporting documents", monthName(m), "", "Tick each document when it is received. A service shows READY when all its documents are in.") + `<div class="mhead">${mb}</div>`
    + `<div class="tiles">${(() => { const r = SUPPORT.filter(([svc, docs]) => docs.every(dc => rec(invKey(m, svc, dc)).done === "✓")).length; return `<div class="tile ${r === SUPPORT.length ? "good" : ""}"><div class="l">Services ready to invoice</div><div class="n">${r} of ${SUPPORT.length}</div></div><div class="tile ${nDone === ADDITIONAL.length ? "good" : ""}"><div class="l">Additional documents</div><div class="n">${nDone} of ${ADDITIONAL.length}</div></div>`; })()}</div>`
    + `<div class="openers">${sheet("sup", `Supporting documents per service – ${monthName(m)}`, `<div class="tbl card" style="padding:0"><table><thead><tr><th>Service / machine</th><th>Supporting documents</th><th>Ready</th></tr></thead><tbody>${sup}</tbody></table></div>`, { icon: icon("docs"), sub: "per service, with the person responsible" })}
      ${sheet("add", `Additional documents – ${monthName(m)}`, `<div class="tbl card" style="padding:0"><table><thead><tr><th>Additional document</th><th>Done</th><th>Person responsible</th><th>Control room collector</th><th>Notes</th></tr></thead><tbody>${add}</tbody></table></div>`, { icon: icon("missing"), sub: `${nDone} of ${ADDITIONAL.length} received` })}</div>`;
}
function exportView() {
  const m = month, mb = `<div class="mbar"><button data-mnav="-1">‹</button><b>${monthName(m)}</b><button data-mnav="1" ${m >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
  return pageHead("Export and sync", "One file when you need it", "", "", "")
    + `<section class="card dl-card"><h2>Download Excel (date to date)</h2><div class="mhead">${mb}</div><p class="muted">The same sheets as the control room workbook: the invoicing lists, Monthly Availability and one sheet per day with the documents, the day and night shift times, the 24-hour summary and the hourly update log.</p><div class="row" style="align-items:flex-end;margin-top:10px"><label style="flex:1 1 150px">From <input type="date" id="cxFrom" value="${m}-01" max="${today()}"></label><label style="flex:1 1 150px">To <input type="date" id="cxTo" value="${monthEnd(m) < today() ? monthEnd(m) : today()}" max="${today()}"></label><button class="primary" id="xl" style="min-height:52px;flex:1 1 200px">Download Excel for these dates</button></div><p class="muted" style="margin-bottom:0">Pick a month with ‹ › or set your own From and To dates.</p></section>`
    + syncCard()
    + `<section class="card"><h2>Backup</h2><p class="muted" style="margin-top:0">Only needed if this PC never gets signal: save a backup file and load it on another PC.</p><div class="row"><button id="bk">Save a backup file</button><label class="btn" style="display:inline-flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:10px;padding:8px 14px;cursor:pointer">Load a backup <input type="file" id="rs" accept=".json" hidden></label></div></section>`;
}
function settingsView() {
  const F = fleet(), users = R("user");
  return pageHead("Settings", "Machines and controllers", "", "", "Changes here go to every control room device with the next sync. You can also add or remove a machine on the live board.")
    + `<section class="card"><h2>Machines</h2>${F.map(([g, L], gi) => `<div style="margin-bottom:12px"><b>${esc(g)}</b><div class="row" style="margin-top:6px">${L.map(m => `<span class="tag" style="font-size:13px;padding:4px 10px">${esc(m)} <button class="sm ghost" data-rmm="${gi}|${esc(m)}" style="min-height:28px;padding:0 6px">×</button></span>`).join("")}</div>
      <div class="row" style="margin-top:6px"><input data-newm="${gi}" placeholder="Add a machine to ${esc(g)}" style="max-width:260px"><button class="sm" data-addm="${gi}">+ Add</button></div></div>`).join("")}</section>
    <section class="card"><h2>Controllers</h2><div class="res">${users.filter(u => u.status !== "declined").map(u => `<div class="ri"><div class="rt"><b>${esc(u.name)}</b><small>${u.status === "pending" ? '<span style="color:var(--gold);font-weight:600">waiting for the admin to approve</span>' : "added " + fmt(ymd(new Date(u.addedAt || u.editedAt)))}</small></div><div class="row" style="gap:6px">${u.status === "pending" && me.admin ? `<button class="sm primary" data-uok="${u.id}">Approve</button><button class="sm danger" data-uno="${u.id}">Decline</button>` : `${me.admin && u.id !== me.id ? `<button class="sm" data-rpin="${u.id}">Reset PIN</button>` : ""}${u.id === me.id || me.admin ? `<button class="sm danger" data-rmu="${u.id}">${u.id === me.id ? "Remove me" : "Remove"}</button>` : ""}`}</div></div>`).join("")}</div></section>`;
}


/* ---------- the control room workbook, as in Excel: Monthly Availability, then one sheet per day ---------- */
let sheetSel = "", sheetDirty = false;
const shiftWin = (d, sh) => sh === "Day" ? [toMs(d + "T06:00"), toMs(d + "T18:00")] : [toMs(d + "T18:00"), toMs(addDays(d, 1) + "T06:00")];
const isoAt = (d, sh, t) => sh === "Day" || t >= "18:00" ? `${d}T${t}` : `${addDays(d, 1)}T${t}`;
const tOf = ms => { const x = new Date(ms); return `${pad(x.getHours())}:${pad(x.getMinutes())}`; };
function evsIn(m, d, sh) { const [a, b] = shiftWin(d, sh); return evs().filter(e => e.machine === m && toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)); }
const n2 = v => v == null ? "" : hm(v);
function sheetTabs() {
  const m = (sheetSel && sheetSel !== "month" ? sheetSel : date).slice(0, 7), days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  return `<div class="xltabs"><button data-wbm="-1" aria-label="Month before">‹</button><b>${monthName(m)}</b><button data-wbm="1" aria-label="Next month" ${m >= today().slice(0, 7) ? "disabled" : ""}>›</button>
    <button class="xt ${sheetSel === "month" ? "on" : ""}" data-sh="month">Monthly Availability</button>${days.map(x => `<button class="xt ${sheetSel === x ? "on" : ""}" data-sh="${x}" ${x > today() ? "disabled" : ""}>${x.slice(8)} ${MONTHS[+x.slice(5, 7) - 1].slice(0, 3)}</button>`).join("")}</div>`;
}
function workbookSection() {
  if (!sheetSel) sheetSel = date;
  return `<h2 class="sect" style="margin-top:26px">Control room workbook</h2>${!me ? '<p class="muted" style="margin-top:-4px"><b>Viewing only.</b> The hours are filled in by the controller on shift.</p>' : ""}`
    + `<p class="muted ctl-only" style="margin-top:-4px">The same sheets as the Excel. Pick OP, BD or PM in each hour box, like on the Excel. When you pick BD or PM the exact time is logged by itself, and again when you pick OP. At the end of the shift tap <b>Save and upload</b>.</p>`
    + `<div class="xlwin"><div class="xlsheet">${sheetSel === "month" ? monthSheet(date.slice(0, 7)) : daySheet(sheetSel)}</div>${sheetTabs()}</div>`;
}
function monthSheet(m) {
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  const T = () => ({ avail: 0, BD: 0, PM: 0, SB: 0, down: 0, per: days.map(() => 0) });
  const add = (t, S) => { S.forEach((s, i) => { if (s) { t.per[i] += s.avail; for (const k of ["avail", "BD", "PM", "SB", "down"]) t[k] += s[k]; } }); };
  const line = (lab, t, cls = "") => `<tr class="${cls}"><td colspan="2">${lab}</td>${t.per.map(v => `<td>${v ? hm(v) : ""}</td>`).join("")}<td>${n2(t.avail)}</td><td>${n2(t.BD)}</td><td>${n2(t.PM)}</td><td>${n2(t.down)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : ""}</b></td></tr>`;
  const fleetT = T(); let body = "";
  for (const [g, L] of fleet()) { const gt = T();
    L.forEach((mc, i) => { const S = days.map(x => daySum(mc, x)), t = T(); add(t, S); add(gt, S);
      body += `<tr>${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><b>${esc(mc)}</b></td>${S.map(s => `<td class="${s && s.BD ? "hot" : ""}">${s ? hm(s.avail) : ""}</td>`).join("")}<td>${n2(t.avail)}</td><td>${n2(t.BD)}</td><td>${n2(t.PM)}</td><td>${n2(t.down)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : ""}</b></td></tr>`; });
    body += line(`${esc(g)} total`, gt, "gt"); for (const k of ["avail", "BD", "PM", "SB", "down"]) fleetT[k] += gt[k]; gt.per.forEach((v, i) => fleetT.per[i] += v); }
  body += line("Fleet total", fleetT, "ft");
  return `<div class="xlhead"><b>Monthly Availability – ${monthName(m)}</b><span>TKR OPERATIONS · THARISA MINE</span><small>Available hours per machine per day (06:00 to 06:00), with month totals. Filled in by itself from the day sheets.</small></div>
    <div class="xlwrap"><table class="xl"><thead><tr><th colspan="2"></th><th colspan="${days.length}">Available time per day (hours:minutes)</th><th colspan="5">Month totals</th></tr>
    <tr><th>Group</th><th>Machine</th>${days.map(x => `<th>${+x.slice(8)}</th>`).join("")}<th>Available hrs</th><th>Breakdown hrs</th><th>Planned maint. hrs</th><th>Down hrs</th><th>Availability %</th></tr></thead><tbody>${body}</tbody></table></div>`;
}
const TSEL = v => `<option value=""></option>${Object.keys(TYPES).map(k => `<option ${k === v ? "selected" : ""}>${k}</option>`).join("")}`;
/* hourly grid, as on the Excel day sheet: one drop-down per machine per hour (OP / BD / PM).
   Hours are numbered 6..29 from 06:00 (24..29 = 00:00–06:00 the next morning). */
const HRS = sh => sh === "Day" ? [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17] : [18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29];
const hStart = (d, h) => toMs(d + "T06:00") + (h - 6) * 36e5;
const hLabel = h => `${pad(h % 24)}–${pad((h + 1) % 24)}`;
const isoOf = ms => { const x = new Date(ms); return `${ymd(x)}T${pad(x.getHours())}:${pad(x.getMinutes())}`; };
const cellsOf = (d, mc) => recs.find(r => r.id === `hr_${d}_${mc}` && !r.deleted)?.cells || {};
const downEvs = mc => evs().filter(e => e.machine === mc && (e.type === "BD" || e.type === "PM"));
/* a shift the control room worked in: every hour that has passed and has no BD or PM counts as OP (clean slate: machines are working unless marked) */
const usedCache = new Map();
function shiftUsed(d, sh) { if (curShift().key === d + "|" + sh) return true;
  const k = d + "|" + sh + "|" + recs.length; if (usedCache.has(k)) return usedCache.get(k);
  const [a, b] = shiftWin(d, sh), H = HRS(sh);
  const u = recs.some(r => !r.deleted && ((r.kind === "hr" && !r.seed && r.date === d && H.some(h => r.cells?.[h])) || (r.kind === "ev" && toMs(r.start) < b && evEnd(r) > a)));
  usedCache.set(k, u); return u; }
function cellVal(d, mc, h, C = cellsOf(d, mc)) {
  if (offAt(mc, hStart(d, h))) return hStart(d, h) <= Date.now() ? "OFF" : "";
  if (C[h] === "-" || hStart(d, h) > Date.now()) return "";   // cleared, or an hour that hasn't started: blank
  if (C[h]) return C[h];
  const a = hStart(d, h), b = a + 36e5, s0 = shiftWin(d, h < 18 ? "Day" : "Night")[0];
  if (a > Date.now()) return "";
  const e = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a && (e.end || toMs(e.start) >= s0));   // a breakdown still open from an earlier shift does not fill the new shift's boxes
  if (e.length) return e.some(x => x.type === "BD") ? "BD" : "PM";
  return "";   // nothing filled in: the box stays blank until the controller fills it
}
/* the hour before: was the machine left on breakdown or maintenance? (a timed breakdown that was closed does not carry on) */
const prevHour = (d, h) => h === 6 ? [addDays(d, -1), 29] : [d, h - 1];
function carryFrom(d, mc, h, depth = 0) {
  const k = `${d}|${mc}|${h}`; if (cvMemo.has(k)) return cvMemo.get(k); if (depth > 48) return "";
  const [d2, h2] = prevHour(d, h), a = hStart(d2, h2), b = a + 36e5; let v = "";
  if (!offAt(mc, a)) { const C = cellsOf(d2, mc);
    if (C[h2]) v = C[h2] === "BD" || C[h2] === "PM" || C[h2] === "SB" ? C[h2] : "";   // "-" (cleared) and OP stop it
    else { const E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a);
      if (E.length) { const g = E.filter(e => e.grid && evEnd(e) <= b); v = g.length === E.length ? (E.some(x => x.type === "BD") ? "BD" : "PM") : (E.some(e => evEnd(e) >= b) ? (E.some(x => x.type === "BD") ? "BD" : "PM") : ""); }
      else v = carryFrom(d2, mc, h2, depth + 1); } }
  cvMemo.set(k, v); return v; }
/* one hour box: how much of it has passed, and how much of that was breakdown or maintenance (to the minute) */
function hourParts(mc, d, hr, v, inT) {
  const x = hStart(d, hr), y = Math.min(x + 36e5, Date.now()), out = { len: 0, BD: 0, PM: 0 }; if (y <= x || !v || v === "OFF") return out;
  out.len = (y - x) / 36e5; const e = hoursIn(mc, x, y);
  if ((e.BD || 0) + (e.PM || 0) > 0) { out.BD = Math.min(e.BD || 0, out.len); out.PM = Math.min(e.PM || 0, out.len - out.BD);
    if (inT === "BD" || inT === "PM") { const f = downEvs(mc).filter(q => toMs(q.start) < y && evEnd(q) > x).map(q => toMs(q.start)).sort((p2, q2) => p2 - q2)[0];   // still down from the hour before
      if (f > x) out[inT] += (Math.min(f, y) - x) / 36e5; } }
  else if (v === "BD" || v === "PM") out[v] = out.len;
  if (out.BD + out.PM > out.len) { const k = out.len / (out.BD + out.PM); out.BD *= k; out.PM *= k; }
  return out; }
function shiftSum(mc, d, sh) {
  const [, b0] = shiftWin(d, sh), H = HRS(sh), C = cellsOf(d, mc), V = H.map(h => cellVal(d, mc, h, C)), filled = V.filter(v => v && v !== "OFF").length;
  let BD = 0, PM = 0, span = 0;
  H.forEach((hr, k) => { if (hStart(d, hr) >= b0) return; const q = hourParts(mc, d, hr, V[k], k ? V[k - 1] : ""); span += q.len; BD += q.BD; PM += q.PM; });
  const down = BD + PM;
  return { filled, BD, PM, down, avail: Math.max(0, span - down), SB: 0, NU: 0 };
}
function daySum(mc, d) {
  const x = shiftSum(mc, d, "Day"), y = shiftSum(mc, d, "Night"), filled = x.filled + y.filled, BD = x.BD + y.BD, PM = x.PM + y.PM, down = BD + PM;
  if (!filled && !down) return null; const avail = x.avail + y.avail;
  return { total: avail + down, avail, BD, PM, SB: 0, NU: 0, down, pct: avail + down > 0 ? avail / (avail + down) : null };
}
/* hour meter: opening and closing hours per machine per shift */
const meterOf = (d, mc, sh) => recs.find(r => r.id === `hr_${d}_${mc}` && !r.deleted)?.meter?.[sh] || {};
const prevShift = (d, sh) => sh === "Day" ? [addDays(d, -1), "Night"] : [d, "Day"];
function lastClose(d, mc, sh) { let [d2, s2] = prevShift(d, sh); for (let i = 0; i < 60; i++) { const m = meterOf(d2, mc, s2); if (m.close != null && m.close !== "") return +m.close; if (m.open != null && m.open !== "") return null; [d2, s2] = prevShift(d2, s2); } return null; }
const meterHrs = m => m.open !== undefined && m.open !== "" && m.close !== undefined && m.close !== "" && +m.close >= +m.open ? +m.close - +m.open : null;
const mtxt = v => v == null || v === "" ? "" : String(Math.round(+v * 10) / 10).replace(".", ",");
async function setMeter(mc, d, sh, k, raw) {
  const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id && !r.deleted) || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells: {} };
  const v = String(raw).trim().replace(/\s/g, "").replace(",", "."); if (v !== "" && !/^\d+(\.\d+)?$/.test(v)) { toast("Type the hour meter reading as a number, e.g. 12345,6"); return false; }
  const m = { ...(rec.meter?.[sh] || {}), [k]: v === "" ? "" : +v };
  if (m.open !== undefined && m.open !== "" && m.close !== undefined && m.close !== "" && +m.close < +m.open) { toast("The closing hours can't be lower than the opening hours."); return false; }
  await put({ ...rec, deleted: false, meter: { ...(rec.meter || {}), [sh]: m } }, `${mc} ${sh.toLowerCase()} shift ${k === "open" ? "opening" : "closing"} hours: ${v ? mtxt(v) : "cleared"}`); return true; }
function shiftTable(d, sh) {
  const [a, b] = shiftWin(d, sh), now = Date.now(), fut = a > now, H = HRS(sh);
  const tot = () => ({ avail: 0, BD: 0, PM: 0, down: 0 }), shT = tot(); let body = "";
  const opt = v => ["", "OP", "SB", "BD", "PM"].map(o => `<option ${o === v ? "selected" : ""}>${o}</option>`).join("");
  for (const [g, L] of fleet()) { const gt = tot(), opCount = H.map(() => 0);
    L.forEach((mc, i) => { const hr = recs.find(r => r.id === `hr_${d}_${mc}` && !r.deleted), C = cellsOf(d, mc), E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)), s = shiftSum(mc, d, sh);
      const reason = [...new Set([...E.map(e => e.reason), hr?.notes?.[sh]].filter(Boolean))].join("; ");
      const cells = H.map((h, k) => { const v = cellVal(d, mc, h, C); if (v === "OP") opCount[k]++;
        if (C[h] && lockedV(hr, "h" + h)) return `<td class="hc lk" data-lockcell="${esc(mc)}|${h}" title="Filled on an earlier shift (${esc(hr.by)}). Tap to ask the admin to change it."><span class="lkv ${v}">${v || "–"} <small>🔒</small></span></td>`;
        const carry = !C[h] && v && v !== "OP", auto = !C[h] && v === "OP";
        return `<td class="hc"><button type="button" data-hcell="${esc(mc)}|${h}" data-v="${v}" class="hbtn ty ${v}${carry ? " carry" : ""}${auto ? " auto" : ""}" title="${carry ? "Carried on from a breakdown that has not been closed. Tap and pick OP if it was working." : "Tap to pick OP, BD or PM. Tap the same one again to clear it."}">${v || "–"}</button></td>`; }).join("");
      const tm = (e, end) => end ? (e.end ? (toMs(e.end) <= b ? hhmm(e.end) : "") : "") : (toMs(e.start) < a ? tOf(a) : hhmm(e.start));
      body += `<tr data-mach="${esc(mc)}">${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><button type="button" class="mbtn" data-mdocs="${esc(mc)}" title="Tap to see and upload the documents for ${esc(mc)}">${esc(mc)}</button></td>${(() => { const m = meterOf(d, mc, sh), lc = lastClose(d, mc, sh), mh = meterHrs(m);
          return `<td class="mt"><input data-meter="${esc(mc)}|open" inputmode="decimal" value="${mtxt(m.open)}" placeholder="${lc != null ? mtxt(lc) : "opening"}" title="Hour meter at the start of the shift${lc != null ? ". Last closing: " + mtxt(lc) : ""}"></td><td class="mt"><input data-meter="${esc(mc)}|close" inputmode="decimal" value="${mtxt(m.close)}" placeholder="closing" title="Hour meter at the end of the shift"></td><td>${mh != null ? hm(mh) : ""}</td>`; })()}${cells}
        <td class="tm">${E.length ? E.map(e => `<span data-edit="${e.id}" class="${evLocked(e) ? "lk" : ""}" title="Tap to correct">${tm(e)}</span>`).join("<br>") : `<span data-newev="${esc(mc)}" title="Log a breakdown with exact times" style="color:#999">+ time</span>`}</td><td class="tm">${E.map(e => `<span data-edit="${e.id}" class="${lockedV(e, "end") ? "lk" : ""}" title="Tap to correct">${tm(e, 1) || "…"}</span>`).join("<br>")}</td>
        <td>${(() => { const lr = E.find(e => lockedV(e, "reason")); return lr ? `<input readonly class="rsn lk" data-edit="${lr.id}" value="${esc(reason)}">` : `<input data-reason="${esc(mc)}" value="${esc(reason)}" class="rsn" placeholder="${E.length ? "" : "comment"}">`; })()}</td><td>${s.filled ? n2(s.avail) : ""}</td><td class="${s.BD ? "hot" : ""}">${s.filled || s.BD ? n2(s.BD) : ""}</td><td>${s.filled || s.PM ? n2(s.PM) : ""}</td><td>${s.filled || s.down ? n2(s.down) : ""}</td></tr>`;
      for (const k of ["avail", "BD", "PM", "down"]) gt[k] += s[k]; });
    body += `<tr class="gt"><td colspan="2">${esc(g)} total</td><td></td><td></td><td>${(() => { const t = L.reduce((x, mc) => x + (meterHrs(meterOf(d, mc, sh)) || 0), 0); return t ? hm(t) : ""; })()}</td>${opCount.map(n => `<td>${n}/${L.length}</td>`).join("")}<td></td><td></td><td></td><td>${n2(gt.avail)}</td><td>${n2(gt.BD)}</td><td>${n2(gt.PM)}</td><td>${n2(gt.down)}</td></tr>`;
    for (const k of ["avail", "BD", "PM", "down"]) shT[k] += gt[k]; }
  body += `<tr class="ft"><td colspan="${H.length + 8}">${sh} shift total</td><td>${n2(shT.avail)}</td><td>${n2(shT.BD)}</td><td>${n2(shT.PM)}</td><td>${n2(shT.down)}</td></tr>`;
  const sv = R("log").filter(l => l.date === d && l.shiftSave === sh).sort((x, y) => y.at - x.at)[0];
  return `<div class="xlband">Machine availability – ${sh} shift (${sh === "Day" ? "06:00 to 18:00" : "18:00 to 06:00"}) &nbsp; OP = Operating · SB = Available, standing (not operating) · BD = Breakdown · PM = Planned maintenance · the exact time is logged when you pick BD or PM, and when you pick OP again · grey = filled on an earlier shift (tap it to ask the admin)</div>
    <div class="xlwrap"><table class="xl av" data-shift="${sh}"><thead><tr><th>Group</th><th>Machine</th><th>Opening hrs</th><th>Closing hrs</th><th>Machine hrs</th>${H.map(h => `<th>${hLabel(h)}</th>`).join("")}<th>Down at</th><th>Back up</th><th>Reason / comments</th><th>Avail hrs</th><th>BD hrs</th><th>PM hrs</th><th>Down hrs</th></tr></thead><tbody>${body}</tbody></table></div>
    <div class="savebar"><button type="button" data-saveonly="${sh}">Save</button><button type="button" class="primary" data-saveshift="${sh}">Save and upload ${sh.toLowerCase()} shift</button><button type="button" data-msg="${sh}">Generate message</button>${undo.length ? `<button type="button" class="danger" data-discard="1" title="Put everything back the way it was at the last Save">Discard changes</button>` : ""}<span class="muted">${sv ? `Saved ${new Date(sv.at).toTimeString().slice(0, 5)}` : ""}</span></div>`;
}
/* one hour box changed: keep it, and log the exact time for a breakdown or planned maintenance */
/* Clear: the box goes blank (stored as "-"), the grid time in it is taken away and nothing carries into it */
async function clearRest(d, mc, h) {
  const a = hStart(d, h), b = a + 36e5;
  for (const e of downEvs(mc).filter(e => e.grid && toMs(e.start) < b && evEnd(e) > a)) { const s0 = toMs(e.start), e0 = evEnd(e);
    if (s0 >= a && e0 <= b) await put({ ...e, deleted: true }); else if (s0 >= a) await put({ ...e, start: isoOf(b) }); else if (e0 <= b) await put({ ...e, end: isoOf(a) });
    else { await put({ ...e, end: isoOf(a) }); await put({ ...e, id: "ev_" + uid(), start: isoOf(b) }); } }
  bump(); }
async function setCell(d, mc, h, val) {
  if (val && hStart(d, h) > Date.now()) { toast("That hour hasn't started yet. Fill it in when it starts."); return; }   // no boxes filled ahead of time
  const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id) || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells: {} };
  const a = hStart(d, h), b = a + 36e5, now = Date.now(), current = now >= a && now < b, cells = { ...rec.cells }; if (val) cells[h] = val; else if (a > now) delete cells[h]; else cells[h] = "-";   // a box ahead of time cleared: as if never touched
  await put({ ...rec, deleted: false, cells }, `${mc} ${hLabel(h)}: ${val || "cleared"}`);
  if (a > now) return;
  if (!val) await clearRest(d, mc, h);
  if (!val && current) {   // cleared in the hour that is running: take away the breakdown or maintenance picked by mistake
    for (const e of downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a && !lockedV(e, "start"))) {
      if (toMs(e.start) >= a) await put({ ...e, deleted: true }, `${mc}: ${TYPES[e.type].toLowerCase()} from ${hhmm(e.start)} taken away`); else await put({ ...e, end: isoOf(a), endBy: me.name }); }
    return; }
  if (!current) for (const e of downEvs(mc).filter(e => (e.grid || !val) && e.type !== val && toMs(e.start) < b && evEnd(e) > a && !(!e.grid && lockedV(e, "start")))) {   // an hour box changed: take that hour off the old grid time
    const s0 = toMs(e.start), e0 = evEnd(e);
    if (s0 >= a && e0 <= b) await put({ ...e, deleted: true }); else if (s0 >= a) await put({ ...e, start: isoOf(b) }); else if (e0 <= b) await put({ ...e, end: isoOf(a) });
    else { await put({ ...e, end: isoOf(a) }); await put({ ...e, id: "ev_" + uid(), start: isoOf(b) }); } }
  const open = downEvs(mc).filter(e => !e.end && toMs(e.start) <= now).sort((x, y) => toMs(y.start) - toMs(x.start))[0];
  if (val === "BD" || val === "PM") {
    if (current) { if (open && open.type === val) return; if (open) await put({ ...open, end: nowLocal(), endBy: me.name });
      await put({ id: "ev_" + uid(), kind: "ev", ym: d.slice(0, 7), machine: mc, group: groupOf(mc), type: val, start: nowLocal(), end: "", reason: "" }, `${mc}: ${TYPES[val]} from ${hhmm(nowLocal())}`);
      toast(`${mc}: ${TYPES[val].toLowerCase()} logged at ${hhmm(nowLocal())}.`);
      if (val === "BD" && mailList().length) { const ne = evs().filter(x => x.machine === mc && !x.end).sort((x, y) => toMs(y.start) - toMs(x.start))[0]; if (ne) openEv(mailForm(ne), "Email this breakdown"); } return; }
    const E = downEvs(mc).filter(e => e.type === val);
    if (E.some(e => toMs(e.start) < b && evEnd(e) > a)) return;                                   // that hour is already inside a logged time
    const before = E.find(e => e.end && toMs(e.end) === a), after = E.find(e => toMs(e.start) === b);
    if (before && !lockedV(before, "end")) await put({ ...before, end: isoOf(b) }); else if (after && !lockedV(after, "start")) await put({ ...after, start: isoOf(a) });
    else await put({ id: "ev_" + uid(), kind: "ev", ym: d.slice(0, 7), machine: mc, group: groupOf(mc), type: val, start: isoOf(a), end: isoOf(b), reason: "", grid: 1 });
    toast(`${mc}: tap Down at or Back up to type the exact time.`); return;
  }
  if (val === "OP" || val === "SB") {
    if (current && open) { await put({ ...open, end: nowLocal(), endBy: me.name }, `${mc} back up ${hhmm(nowLocal())}`); toast(`${mc}: back up logged at ${hhmm(nowLocal())}.`); return; }
    if (open && toMs(open.start) < a) { await put({ ...open, end: isoOf(a), endBy: me.name }); toast(`${mc}: back up at ${tOf(a)}. Tap Back up to type the exact time.`); }
  }
}
async function setReason(mc, d, sh, reason) {
  const [a, b] = shiftWin(d, sh), E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a);
  for (const e of E) if ((e.reason || "") !== reason && !lockedV(e, "reason")) await put({ ...e, reason }, `${mc}: ${reason}`);
  if (!E.length) { const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id) || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells: {} };   // no breakdown on this shift: keep it as a comment
    if ((rec.notes?.[sh] || "") !== reason) await put({ ...rec, deleted: false, notes: { ...(rec.notes || {}), [sh]: reason } }, `${mc} ${sh.toLowerCase()} shift comment: ${reason || "cleared"}`); }
}
function shiftMessage(d, sh) {   // simple: every machine, its status, available or not
  const [a, b] = shiftWin(d, sh), day = recs.find(r => r.id === "day_" + d) || {}, ctl = (sh === "Day" ? day.dayCtl : day.nightCtl) || me?.name || "", now = Date.now();
  const cur = now >= a && now < b, T = cur ? now : Math.min(b, now) - 1, hr = HRS(sh).filter(h => hStart(d, h) <= T).at(-1);
  const stateAt = mc => cur ? stateNow(mc) : offAt(mc, T) ? "OFF" : (hr != null && cellVal(d, mc, hr)) || lastKnown(mc, T + 1) || "?";
  const why = mc => { const n = (recs.find(r => r.id === `hr_${d}_${mc}` && !r.deleted)?.notes?.[sh] || "").trim();   // the comment typed for this shift, else the reason on its latest breakdown
    const e = evs().filter(e => e.machine === mc && e.reason && toMs(e.start) <= T && (evEnd(e) > a || !e.end)).sort((x, y) => toMs(y.start) - toMs(x.start))[0];
    if (n || e?.reason) return n || e.reason.trim();
    let [d2, s2] = [d, sh]; for (let i = 0; i < 4; i++) { [d2, s2] = prevShift(d2, s2); const [x2, y2] = shiftWin(d2, s2), c = (recs.find(r => r.id === `hr_${d2}_${mc}` && !r.deleted)?.notes?.[s2] || "").trim() || (evs().filter(e => e.machine === mc && e.reason && toMs(e.start) < y2 && evEnd(e) > x2).sort((p, q) => toMs(q.start) - toMs(p.start))[0]?.reason || "").trim(); if (c) return c; }   // nothing yet this shift: the last comment
    return ""; };
  let avail = 0, total = 0; const body = [];
  for (const [g, M] of fleet()) { body.push(`*${g}*`);
    for (const mc of M) { const st = stateAt(mc), w = why(mc);
      if (st === "OFF") { body.push(`${mc} – Not on site`); continue; } total++;
      if (st === "OP") { avail++; body.push(`${mc} – Operating – available${w ? " – " + w : ""}`); }   // only OP and SB count as available – the same as the dashboard
      else if (st === "SB") { avail++; body.push(`${mc} – Standby – available, not operating${w ? ": " + w : ""}`); }
      else body.push(`${mc} – ${st === "BD" ? "Breakdown" : st === "PM" ? "Planned maintenance" : st === "NU" ? "Not in use" : "Status not captured yet"}${w ? ": " + w : ""}${st === "?" ? "" : " – not available"}`); }
    body.push(""); }
  return [`*TKR Control Room – ${sh} shift ${dow(d)} ${fmt(d)}*`, `Controller: ${ctl}`, `Status at ${tOf(cur ? now : Math.min(b, now))}`, `*Available: ${avail} of ${total} machines*${(() => { const n = fleet().flatMap(([, M]) => M).filter(mc => stateAt(mc) === "OP").length; return ` (${n} operating)`; })()}`, "", ...body].join("\n").trim();
}
function showMessage(text) {
  document.querySelector(".ccdone")?.remove(); const d = document.createElement("div"); d.className = "ccdone"; d.setAttribute("role", "dialog");
  d.innerHTML = `<div style="max-width:620px"><div class="big">Shift message</div><textarea class="msg" style="width:100%;min-height:300px;font-size:14px">${esc(text)}</textarea>
    <div class="row" style="margin-top:8px"><button class="primary" data-m="copy" style="flex:1">Copy</button><button data-m="wa" style="flex:1">Open WhatsApp</button><button class="ghost" data-m="x" style="flex:1">Close</button></div></div>`;
  d.onclick = async e => { const b = e.target.closest("button"); if (!b) { if (e.target === d) d.remove(); return; } const t = d.querySelector("textarea").value;
    if (b.dataset.m === "copy") { try { await navigator.clipboard.writeText(t); toast("Copied. Paste it in WhatsApp."); } catch { d.querySelector("textarea").select(); document.execCommand("copy"); toast("Copied."); } }
    if (b.dataset.m === "wa") window.open("https://wa.me/?text=" + encodeURIComponent(t), "_blank", "noopener");
    if (b.dataset.m === "x") d.remove(); };
  document.body.appendChild(d);
}
function daySheet(d) {
  const day = recs.find(r => r.id === "day_" + d && !r.deleted) || {}, logs = R("log").filter(l => l.date === d).sort((x, y) => x.at - y.at), last = logs[logs.length - 1];
  const get = item => recs.find(r => r.id === chkKey(d, item)) || {}, opt = v => ["", "✓", "✗"].map(o => `<option ${o === (v || "") ? "selected" : ""}>${o}</option>`).join("");
  const hours = []; for (let h = 6; h < 30; h++) hours.push(h % 24);
  const hourLog = hours.map(h => { const st = toMs((h < 6 ? addDays(d, 1) : d) + `T${pad(h)}:00`), L = logs.filter(l => l.at >= st && l.at < st + 36e5); return [h, L]; });
  return `<form id="fDay" data-date="${d}" autocomplete="off"><div class="xlhead"><b>Control Room Capture – ${parse(d).getDate()} ${MONTHS[+d.slice(5, 7) - 1]} ${d.slice(0, 4)}</b><span>TKR OPERATIONS · THARISA MINE</span><small>Day shift 06:00 to 18:00 · Night shift 18:00 to 06:00</small></div>
    <div class="xlinfo"><label>Date <input value="${fmt(d)}" readonly></label><label>Day controller <input name="dayCtl" value="${esc(day.dayCtl || "")}" placeholder="${esc((me?.name || ""))}"></label><label>Night controller <input name="nightCtl" value="${esc(day.nightCtl || "")}" placeholder="${esc((me?.name || ""))}"></label>
      <label>Last update <input value="${last ? new Date(last.at).toTimeString().slice(0, 5) : ""}" readonly></label><label>Next update <input value="${last ? new Date(last.at + 36e5).toTimeString().slice(0, 5) : ""}" readonly></label>
      <label>Status <input value="${d !== today() ? "" : !last ? "Not started" : Date.now() - last.at > 36e5 ? "Update due" : "Up to date"}" readonly></label></div>
    ${shiftTable(d, "Day")}${shiftTable(d, "Night")}
    <div class="xlband">24-hour availability summary per machine</div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Group</th><th>Machine</th><th>Available hrs</th><th>Breakdown hrs</th><th>Planned maint. hrs</th><th>Down hrs</th><th>Total hrs</th><th>Availability %</th></tr></thead><tbody>
      ${fleet().map(([g, L]) => L.map((mc, i) => { const s = daySum(mc, d); return `<tr>${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><b>${esc(mc)}</b></td>${s ? `<td>${n2(s.avail)}</td><td class="${s.BD ? "hot" : ""}">${n2(s.BD)}</td><td>${n2(s.PM)}</td><td>${n2(s.down)}</td><td>${n2(s.avail + s.down)}</td><td><b>${s.pct == null ? "" : Math.round(s.pct * 100) + "%"}</b></td>` : '<td colspan="6"></td>'}</tr>`; }).join("")).join("")}</tbody></table></div>
    <div class="xlband">Hourly update log <small>(fills in by itself every time something is saved)</small></div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Hour</th><th>Updated</th><th>Time done</th><th>Controller</th><th>Notes</th></tr></thead><tbody>
      ${hourLog.map(([h, L]) => `<tr><td>${pad(h)}:00–${pad((h + 1) % 24)}:00</td><td class="${L.length ? "okc" : ""}">${L.length ? "✓" : ""}</td><td>${L[0] ? new Date(L[0].at).toTimeString().slice(0, 5) : ""}</td><td>${esc([...new Set(L.map(l => l.by))].join(", "))}</td><td class="nt">${esc(L.map(l => l.what).slice(0, 3).join(" · "))}${L.length > 3 ? ` · +${L.length - 3} more` : ""}</td></tr>`).join("")}
      <tr class="gt"><td>Hours updated</td><td colspan="4">${hourLog.filter(([, L]) => L.length).length} of 24</td></tr></tbody></table></div>
    <div class="xlband">Data and documents</div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Section</th><th>Item</th>${COLS.map(c => `<th>${c}</th>`).join("")}<th>Time updated</th><th>Updated by</th><th>Comments</th></tr></thead><tbody>
      ${CHECK.map(([sec, items]) => items.map((it, i) => { const c = get(it); return `<tr data-item="${esc(it)}">${i ? "" : `<td rowspan="${items.length}" class="grpc">${esc(sec)}</td>`}<td class="lt">${esc(it)}</td>${COLS.map(k => lockedV(c, k) ? `<td class="lk" data-lockf="${esc(c.id)}|${k}"><span class="lkv">${c[k]}</span></td>` : `<td><select data-c="${k}" class="tk">${opt(c[k])}</select></td>`).join("")}<td>${c.editedAt ? new Date(c.editedAt).toTimeString().slice(0, 5) : ""}</td><td>${esc(c.by || "")}</td><td>${lockedV(c, "comment") ? `<input class="lk" readonly data-lockf="${esc(c.id)}|comment" value="${esc(c.comment)}">` : `<input data-c="comment" value="${esc(c.comment || "")}">`}</td></tr>`; }).join("")).join("")}
      <tr class="gt"><td colspan="2">Captured</td>${COLS.map(k => `<td>${CHECK.flatMap(([, i]) => i).filter(it => get(it)[k] === "✓").length} of ${CHECK.flatMap(([, i]) => i).length}</td>`).join("")}<td colspan="3"></td></tr></tbody></table></div></form>`;
}
async function saveShift(sh, upload = true) {
  const f = $("fDay"); if (!f) return; const d = f.dataset.date; setTimeout(clearUndo, 0);
  const dayR = recs.find(r => r.id === "day_" + d) || { id: "day_" + d, kind: "day", ym: d.slice(0, 7), date: d };
  const ctl = { dayCtl: f.elements.dayCtl.value.trim() || (sh === "Day" ? me.name : dayR.dayCtl || ""), nightCtl: f.elements.nightCtl.value.trim() || (sh === "Night" ? me.name : dayR.nightCtl || "") };
  if (ctl.dayCtl !== (dayR.dayCtl || "") || ctl.nightCtl !== (dayR.nightCtl || "")) await put({ ...dayR, ...ctl }, `Controllers: day ${ctl.dayCtl || "–"}, night ${ctl.nightCtl || "–"}`);
  const n = R("log").filter(l => l.date === d).length;
  await put({ id: "log_" + uid(), kind: "log", ym: d.slice(0, 7), date: d, at: Date.now(), shiftSave: sh, what: upload ? `Saved and uploaded the ${sh.toLowerCase()} shift sheet` : `Saved the ${sh.toLowerCase()} shift sheet` });
  if (!upload) { render(); popupMsg("Saved", `${sh} shift, ${dow(d)} ${fmt(d)}: saved on this PC. Tap Save and upload at the end of the shift.`); return; }
  let up = false; if (navigator.onLine) { await syncNow(false); up = !waiting(); }
  render();
  popupMsg(up ? "Saved and uploaded" : "Saved on this PC", `${sh} shift, ${dow(d)} ${fmt(d)}: ${n} updates on this sheet.` + (up ? " Everyone in the TKR app can see it now." : " It uploads by itself as soon as this PC has signal."));
}
function popupMsg(title, text) {
  document.querySelector(".ccdone")?.remove(); const d = document.createElement("div"); d.className = "ccdone"; d.setAttribute("role", "alertdialog");
  d.innerHTML = `<div><div class="big"><i>✓</i>${esc(title)}</div><p>${esc(text)}</p><button class="primary">OK</button></div>`; d.onclick = e => { if (e.target.closest("button") || e.target === d) d.remove(); };
  document.body.appendChild(d); d.querySelector("button").focus();
}

/* ---------- log / change an event ---------- */
function evForm(e = {}, machine = "") {
  const M = fleet().flatMap(([, L]) => L), isNew = !e.id, open = machine ? statusNow(machine) : null;
  if (isNew && open) { e = open; }                                 // the machine is already down: open that event to close or change it
  const st = e.start || nowLocal(), en = e.end || "", lk = !!e.id && !me.admin && ["machine", "type", "start", "end", "reason"].some(f => lockedV(e, f));
  return `<form id="fEv" class="card" autocomplete="off" style="max-width:640px;margin:0 auto">
    ${lk ? `<div class="notice warn" style="margin-bottom:10px">Logged by <b>${esc(e.by)}</b> on an earlier shift, so it is locked.${e.end ? "" : " You can still log <b>Back up</b>."} Anything left empty you can fill in. A change to what was already filled in goes to the admin for approval.</div>` : ""}
    ${!isNew || open ? `<div class="info">${esc(e.machine)} is logged as <b>${TYPES[e.type]}</b> since ${fmt(st.slice(0, 10))} ${hhmm(st)}.${e.end ? "" : " Tap <b>Back up now</b> when it works again."}</div>` : ""}
    <input type="hidden" name="eid" value="${esc(e.id || "")}">
    <label>Machine <select name="machine" required>${M.map(m => `<option ${m === (e.machine || machine) ? "selected" : ""}>${esc(m)}</option>`).join("")}</select></label>
    <div style="margin-top:12px"><span class="muted" style="font-size:12px;letter-spacing:1px;text-transform:uppercase;font-weight:600">What happened</span><div class="seg" style="margin-top:6px">${Object.entries(TYPES).filter(([k]) => k !== "SB" || e.type === "SB").map(([k, t]) => `<label><input type="radio" name="type" value="${k}" ${k === (e.type || "BD") ? "checked" : ""}>${t}</label>`).join("")}</div></div>
    <div class="grid" style="margin-top:12px"><label>Started (exact time) <input type="datetime-local" name="start" value="${st}" max="${nowLocal()}" required></label>
      <label>Ended (empty = still going) <input type="datetime-local" name="end" value="${en}" max="${nowLocal()}"></label></div>
    <label style="display:block;margin-top:12px">Reason / what happened <input name="reason" value="${esc(e.reason || "")}" placeholder="e.g. hydraulic pipe burst"></label>
    ${lk ? `<label style="display:block;margin-top:12px">Why must it change? (for the admin) <input name="why" placeholder="e.g. wrong time typed"></label>` : ""}
    <div class="formbar" style="display:flex;gap:8px;flex-wrap:wrap;margin-top:16px">${e.id && !e.end ? `<button type="button" class="primary" data-up="${e.id}" style="flex:1 1 200px;min-height:52px">Back up now (${new Date(Date.now()).toTimeString().slice(0, 5)})</button>` : ""}
      <button class="${e.id && !e.end ? "" : "primary"}" style="flex:1 1 200px;min-height:52px">Save</button>${e.id ? `<button type="button" class="danger" data-delev="${e.id}">${lk && evLocked(e) ? "Ask admin to delete" : "Delete"}</button>` : ""}</div></form>`;
}
function openEv(html, title) {
  document.getElementById("evPanel")?.remove();
  const ov = document.createElement("div"); ov.id = "evPanel"; ov.className = "sheet"; ov.setAttribute("role", "dialog");
  ov.innerHTML = `<div class="sheet-bar"><button type="button" class="back" data-x>‹ Back</button><b>${title}</b></div><div class="sheet-body">${html}</div>`;
  document.body.appendChild(ov); document.body.classList.add("noscroll");
  ov.addEventListener("click", clickH); ov.addEventListener("submit", submitH); ov.addEventListener("change", changeH);
}
const closeEv = () => { document.getElementById("evPanel")?.remove(); if (!document.querySelector(".sheet:not([hidden])")) document.body.classList.remove("noscroll"); };
async function backUp(id) {
  const e = recs.find(r => r.id === id); if (!e) return;
  await put({ ...e, end: nowLocal(), endBy: me.name }, `${e.machine} back up (${TYPES[e.type]} from ${hhmm(e.start)})`);
  closeEv(); render(); toast(`${e.machine} is back up at ${new Date(Date.now()).toTimeString().slice(0, 5)}. Saved.`);
}

/* ---------- sessions: log in, log out ---------- */
let admApp = null;
async function startSession(u) { const sid = "s_" + uid(), start = Date.now(); me = { ...u, sid, start }; await metaSet("me", { id: u.id, admin: !!u.admin, sid, start, until: start + (u.admin ? 12 : 14) * 36e5 }); }
async function logout(swap) {
  if (!me) return;
  if (!me.admin && !confirm(`${swap ? "Change controller: l" : "L"}og out ${me.name} and end your shift?\n\nYour entries are saved and turn grey (locked). Hour boxes you left empty stay open for the next controller. To change a locked entry later, ask the admin.`)) return;
  await put({ id: "log_" + uid(), kind: "log", ym: today().slice(0, 7), date: today(), at: Date.now(), what: `${me.name} logged out${me.admin ? "" : " – end of shift"}` });
  me = null; pinFor = null; adminStep = ""; addingMe = false; closeEv(); closeSheet(app); await metaSet("me", null); await clearUndo(); tab = "now"; sheetSel = ""; render(); scrollTo(0, 0);
  toast(swap ? "Logged out. The next controller taps their name, or Add yourself." : "Logged out. Everything is saved on this PC."); syncNow(false);
}

/* ---------- change requests: locked entries change only once the admin approves ---------- */
const evTxt = e => `${e.machine} ${TYPES[e.type] || ""} ${e.start ? fmt(e.start.slice(0, 10)).slice(0, 5) + " " + hhmm(e.start) : ""}–${e.end ? hhmm(e.end) : "still going"}${e.reason ? " · " + e.reason : ""}`;
const FLAB = { done: "done", comment: "comment", collector: "collector", notes: "notes" };
function reqTarget(t) {
  if (t.type === "cell") return { what: `${t.mc} · ${hLabel(t.h)} · ${dow(t.d)} ${fmt(t.d)}`, now: cellsOf(t.d, t.mc)[t.h] || "", opts: ["", "OP", "SB", "BD", "PM"] };
  const r = recs.find(x => x.id === t.id) || {}, f = t.field;
  const what = r.kind === "chk" ? `${r.item} · ${FLAB[f] || f} · ${fmt(r.date)}` : r.kind === "inv" ? `${r.svc} – ${r.doc} · ${monthName(r.ym)}` : `${r.item} · ${FLAB[f] || f} · ${monthName(r.ym || today().slice(0, 7))}`;
  return { what, now: r[f] || "", opts: [...COLS, "done"].includes(f) ? ["", "✓", "✗"] : null };
}
function reqForm(t) {
  const x = reqTarget(t), r = t.type === "cell" ? recs.find(q => q.id === `hr_${t.d}_${t.mc}`) : recs.find(q => q.id === t.id);
  return `<form id="fReq" class="card" autocomplete="off" style="max-width:520px;margin:0 auto"><input type="hidden" name="t" value="${esc(JSON.stringify(t))}">
    <div class="notice warn" style="margin-bottom:12px">Filled in by <b>${esc(r?.by || "an earlier controller")}</b> on an earlier shift, so it is locked. The admin approves or declines the change.</div>
    <p style="margin-top:0"><b>${esc(x.what)}</b><br>Now: <b>${esc(x.now || "empty")}</b></p>
    <label>Change to ${x.opts ? `<select name="to">${x.opts.map(o => `<option value="${o}" ${o === x.now ? "selected" : ""}>${o || "(empty)"}</option>`).join("")}</select>` : `<input name="to" value="${esc(x.now)}">`}</label>
    <label style="display:block;margin-top:12px">Why must it change? <input name="why" required placeholder="e.g. picked the wrong hour"></label>
    <button class="primary" style="width:100%;min-height:52px;margin-top:14px">Send to the admin</button></form>`;
}
async function addReq(o) {
  await put({ id: "req_" + uid(), kind: "req", ym: today().slice(0, 7), date: today(), status: "pending", askedBy: me.name, askedAt: Date.now(), ...o }, `Asked the admin: ${o.what}: ${o.fromTxt} → ${o.toTxt}`);
  closeEv(); render(); popupMsg("Sent to the admin", "The admin sees the request under Change requests. Once it is approved the change shows here by itself (when the PC has signal).");
}
async function decide(id, ok) {
  const q = recs.find(r => r.id === id); if (!q || q.status !== "pending") return;
  if (ok) { const t = q.target;
    if (t.type === "cell") await setCell(t.d, t.mc, t.h, q.to);
    else if (t.type === "ev") { const e = recs.find(r => r.id === t.id); if (e) await put(q.del ? { ...e, deleted: true } : { ...e, ...q.to, group: groupOf(q.to.machine) }, `Admin approved: ${q.what} → ${q.toTxt}`); }
    else { const r = recs.find(x => x.id === t.id); if (r) await put({ ...r, [t.field]: q.to }, `Admin approved: ${q.what} → ${q.toTxt}`); } }
  await put({ ...q, status: ok ? "approved" : "declined", decidedBy: me.name, decidedAt: Date.now() });
  render(); toast(ok ? "Approved. The change is made." : "Declined.");
}
function reqView() {
  const PU = pendingUsers(), puCard = !PU.length ? "" : `<section class="card" style="border:2px solid var(--gold)"><h2>New controllers waiting (${PU.length})</h2><div class="res">${PU.map(u => `<div class="ri"><div class="rt"><b>${esc(u.name)}</b><small>asked ${fmt(ymd(new Date(u.addedAt)))} ${tOf(u.addedAt)}</small></div>${me.admin ? `<div class="row" style="gap:6px"><button class="sm primary" data-uok="${u.id}">Approve</button><button class="sm danger" data-uno="${u.id}">Decline</button></div>` : '<span class="muted">waiting for the admin</span>'}</div>`).join("")}</div></section>`;
  const Q = R("req").sort((a, b) => (a.status === "pending" ? 0 : 1) - (b.status === "pending" ? 0 : 1) || b.askedAt - a.askedAt), P = Q.filter(q => q.status === "pending");
  const when = t => `${fmt(ymd(new Date(t))).slice(0, 5)} ${new Date(t).toTimeString().slice(0, 5)}`;
  return pageHead("Control room", "Change requests", `${P.length} waiting`, "", me.admin ? "Changes asked for entries from an earlier shift. Approve makes the change; decline leaves it as it is." : "Grey boxes were filled in on an earlier shift and are locked. Tap a grey box to ask the admin to change it.")
    + puCard + `<div class="res">${Q.map(q => `<div class="ri" style="flex-wrap:wrap"><div class="rt"><b>${esc(q.what)}</b><small>${esc(q.fromTxt)} → <b>${esc(q.toTxt)}</b></small><small>Why: ${esc(q.why)}</small><small class="muted">asked by ${esc(q.askedBy)} ${when(q.askedAt)}${q.decidedBy ? ` · ${q.status} by ${esc(q.decidedBy)} ${when(q.decidedAt)}` : ""}</small></div>
      ${q.status === "pending" ? (me.admin ? `<div class="row" style="gap:6px"><button class="sm primary" data-appr="${q.id}">Approve</button><button class="sm danger" data-decl="${q.id}">Decline</button></div>` : '<span class="tag SB">WAITING</span>') : `<span class="tag ${q.status === "approved" ? "OP" : "BD"}">${q.status.toUpperCase()}</span>`}</div>`).join("") || '<p class="muted">No change requests.</p>'}</div>`;
}

/* ---------- add / remove a machine ---------- */
async function addMachine(gi, raw) {
  const v = String(raw || "").trim().toUpperCase().replace(/\s+/g, ""); if (!v) { toast("Type the machine number."); return false; }
  const G = fleet().map(([g, L]) => [g, [...L]]); if (G.some(([, L]) => L.includes(v))) { toast(v + " is already on the list."); return false; }
  G[gi][1].push(v); await put({ id: "fleet", kind: "fleet", ym: "setup", groups: G }, `Added machine ${v} to ${G[gi][0]}`); toast(v + " added."); return true;
}
async function removeMachine(m) { const G = fleet().map(([g, L]) => [g, L.filter(x => x !== m)]); await put({ id: "fleet", kind: "fleet", ym: "setup", groups: G }, `Removed machine ${m}`); }
function siteForm(gi) {
  const [g, L] = fleet()[gi], c = curShift(), [d, sh] = c.key.split("|"), st = toMs(d + (sh === "Day" ? "T06:00" : "T18:00"));
  return `<form id="fSite" class="card" autocomplete="off" style="max-width:520px;margin:0 auto"><p class="muted" style="margin-top:0">A machine that is not on site is left out of the availability, the board, the shift message and the documents check, from the time you pick. Its hours stay as they were.</p>
    <div class="res">${L.map(m => { const off = offNow(m); return `<div class="ri"><div class="rt"><b>${esc(m)}</b><small>${off ? '<span style="color:var(--gold);font-weight:600">Not on site</span>' : "On site"}</small></div><button type="button" class="sm ${off ? "primary" : ""}" data-site="${esc(m)}|${off ? "on" : "off"}">${off ? "Back on site" : "Not on site"}</button></div>`; }).join("")}</div>
    <label style="margin-top:12px">From <input type="datetime-local" name="from" value="${isoOf(st)}" max="${nowLocal()}"></label>
    <p class="muted" style="margin:6px 0 0">Starts at the beginning of this shift. Pick an earlier date if the machine left site before today.</p></form>`;
}
async function setSite(m, off, fromMs) {
  const r = structuredClone(siteRec()), P = r.off[m] = r.off[m] || [];
  if (off) { if (offAt(m, fromMs)) return; P.push({ from: fromMs, to: null }); }
  else { const o = P.find(p => !p.to); if (o) o.to = Math.max(o.from, fromMs); }
  await put(r, `${m}: ${off ? "not on site" : "back on site"} from ${fmt(ymd(new Date(fromMs)))} ${tOf(fromMs)}`); }
function machForm(gi, mode) {
  const [g, L] = fleet()[gi];
  return `<form id="fMach" data-gi="${gi}" data-mode="${mode}" class="card" autocomplete="off" style="max-width:440px;margin:0 auto">${mode === "add"
    ? `<label>New machine in ${esc(g)} <input name="m" required autocapitalize="characters" placeholder="e.g. ${esc(L[0] || "TDR020")}"></label><button class="primary" style="width:100%;min-height:52px;margin-top:14px">Add machine</button>`
    : `<label>Machine to take off ${esc(g)} <select name="m">${L.map(m => `<option>${esc(m)}</option>`).join("")}</select></label><p class="muted">Its logged times are kept.</p><button class="danger" style="width:100%;min-height:52px">Remove machine</button>`}</form>`;
}

/* ---------- breakdown email: optional, the controller chooses who gets it ---------- */
let askMail = false;
const cleanMails = v => [...new Set(String(v || "").split(/[\s,;]+/).map(x => x.trim().toLowerCase()).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)))].join(", ");
const mailList = () => [...new Set([...(me?.mailTo || "").split(", "), ...R("user").flatMap(u => (u.mailTo || "").split(", "))].filter(Boolean))];
function mailToView() {
  return pageHead("Control Room", `Welcome, ${esc(me.name.split(" ")[0])}`, "", "", "Optional: who must get an email when you log a breakdown? Leave it empty if nobody.")
    + `<form id="fMailTo" class="card" style="max-width:560px" autocomplete="off"><label>Email addresses (separate with a comma) <textarea name="to" rows="2" placeholder="e.g. workshop@tkroperations.com, godfrey@tkroperations.com">${esc(me.mailTo || "")}</textarea></label>
      ${mailList().length ? `<p class="muted">Used on this PC before: ${esc(mailList().join(", "))}</p>` : ""}
      <div class="row" style="margin-top:12px"><button class="primary" style="flex:1;min-height:50px">Save and start the shift</button><button type="button" class="ghost" id="skipMail">Skip</button></div></form>`;
}
function mailForm(e) {
  const mine = (me?.mailTo || "").split(", ").filter(Boolean), L = mailList();
  const subj = `Breakdown: ${e.machine} (${groupOf(e.machine)}) down at ${hhmm(e.start)} on ${fmt(e.start.slice(0, 10))}`;
  const body = `Machine: ${e.machine} (${groupOf(e.machine)})\nWhat: ${TYPES[e.type]}\nDown at: ${fmt(e.start.slice(0, 10))} ${hhmm(e.start)}\nBack up: ${e.end ? fmt(e.end.slice(0, 10)) + " " + hhmm(e.end) : "still down"}\nReason: ${e.reason || ""}\n\nLogged by ${me?.name || e.by || ""}, TKR Control Room`;
  return `<form id="fMail" class="card" autocomplete="off" style="max-width:600px;margin:0 auto"><p class="muted" style="margin-top:0">Optional. Tick who must get it, then tap Open email. Your email program opens with it written; it sends when the PC has signal.</p>
    <div class="res">${L.map(a => `<label class="ri" style="cursor:pointer"><input type="checkbox" name="to" value="${esc(a)}" ${mine.includes(a) ? "checked" : ""} style="width:22px;height:22px"> <span>${esc(a)}</span></label>`).join("") || '<p class="muted">No addresses yet.</p>'}</div>
    <label style="display:block;margin-top:10px">Other address (optional) <input name="extra" type="email" placeholder="name@company.com"></label>
    <label style="display:block;margin-top:10px">Subject <input name="subj" value="${esc(subj)}"></label>
    <label style="display:block;margin-top:10px">Message <textarea name="body" rows="7">${esc(body)}</textarea></label>
    <div class="row" style="margin-top:12px"><button class="primary" style="flex:1;min-height:50px">Open email</button><button type="button" class="ghost" data-x>Not now</button></div></form>`;
}

/* ---------- documents per machine per shift (job cards, checklists…), kept on this PC and sent when there is signal ---------- */
const DOC_TYPES = ["Fleet report", "Job card", "Checklist", "Tally sheet", "Trip summary", "Breakdown report", "Other"];
const NEEDS = { "Drill rigs": ["Checklist", "Tally sheet"], "Rockbreakers": ["Checklist"], "Buses": ["Checklist", "Trip summary"], "Lowbed": ["Checklist", "Trip summary"], "Water bowser": ["Checklist", "Trip summary"], "Bakkies": ["Checklist"] };
const docsOf = (m, d, sh) => R("doc").concat(R("mdoc")).filter(x => x.machine === m && (!d || (x.date === d && x.shift === sh)));   // mdoc: drill rig documents from the Metres section
function shiftsSoFar(m0) { const L = [], c = curShift(), end = m0 === today().slice(0, 7) ? today() : monthEnd(m0);
  for (let d = m0 + "-01"; d <= end; d = addDays(d, 1)) for (const sh of ["Day", "Night"]) { if (shiftWin(d, sh)[0] > Date.now()) continue; L.push([d, sh]); } return L; }
function needFor() { return ["Checklist"]; }        // the minimum per machine per shift; any other document can be added
function docMissing(m, m0 = month) { return shiftsSoFar(m0).flatMap(([d, sh]) => needFor(m, d, sh).filter(t => !docsOf(m, d, sh).some(x => x.type === t)).map(t => ({ d, sh, t }))); }
function docsPanel(m, sel = "") {
  const end = monthEnd(month), cells = [], now = Date.now();
  for (let d = month + "-01"; d <= end; d = addDays(d, 1)) for (const sh of ["Day", "Night"]) { const fut = shiftWin(d, sh)[0] > now, have = docsOf(m, d, sh);
    cells.push({ d, sh, fut, have, st: fut ? "fut" : have.some(x => x.type === "Checklist") ? "ok" : have.length ? "part" : "miss" }); }
  const past = cells.filter(c => !c.fut), withCl = past.filter(c => c.st === "ok").length, other = docsOf(m).filter(x => x.date.startsWith(month) && x.type !== "Checklist").length;
  const pick = sel ? cells.find(c => `${c.d}|${c.sh}` === sel) : past.at(-1), all = docsOf(m).filter(x => x.kind === "doc" && x.date.startsWith(month)).length;
  const tileH = (l, n, sub, cls) => `<div class="tile ${cls}"><div class="l">${l}</div><div class="n">${n}</div><div class="s">${sub}</div></div>`;
  return `<div style="max-width:1100px;margin:0 auto"><div class="row" style="justify-content:space-between;align-items:flex-end;margin-bottom:10px"><div><span class="muted" style="font-size:12px;letter-spacing:1px;text-transform:uppercase">${esc(groupOf(m))} · documents${/^TDR/.test(m) ? " · linked with the Metres section" : ""}</span><h2 style="margin:2px 0 0">${esc(m)} · ${monthName(month)}</h2></div>
      <div class="row" style="gap:6px"><button class="primary" data-upload="${esc(m)}">Upload documents</button><button data-dzip="${esc(m)}">Download all (${all})</button></div></div>
    <div class="tiles">${tileH("Shifts with a checklist", `${withCl} of ${past.length}`, `so far · ${cells.length} in the month`, withCl < past.length ? "alert" : "good")}${tileH("No checklist", String(past.length - withCl), "up to now", past.length - withCl ? "alert" : "good")}${tileH("Other documents", String(other), "job cards, trip summaries and others", "")}</div>
    <div class="legend2"><span><i style="background:var(--ok)"></i>Checklist uploaded</span><span><i style="background:var(--gold)"></i>Other documents, no checklist</span><span><i style="background:var(--bad)"></i>No checklist</span></div>
    <div class="mgrid">${cells.map(c => `<button class="mcell ${c.st}" data-dshift="${esc(m)}|${c.d}|${c.sh}" ${c.fut ? "disabled" : ""} style="${pick && pick.d === c.d && pick.sh === c.sh ? "outline:3px solid var(--gold);outline-offset:1px" : ""}"><b>${c.d.slice(8)}</b>${c.sh}</button>`).join("")}</div>
    ${pick ? `<section class="card"><div class="row" style="justify-content:space-between;align-items:center"><h3 style="margin:0">${dow(pick.d)} ${fmt(pick.d)} · ${pick.sh} shift</h3>
        <div class="row" style="gap:6px">${pick.have.some(x => x.type === "Checklist") ? "" : `<button class="sm primary" data-upload="${esc(m)}|${pick.d}|${pick.sh}|Checklist">+ Checklist</button>`}<button class="sm" data-upload="${esc(m)}|${pick.d}|${pick.sh}|Job card">+ Other document</button></div></div>
      <div class="res" style="margin-top:10px">${pick.have.map(x => x.kind === "mdoc" ? `<div class="ri"><div class="rt"><b>${esc(x.type)}</b><small>${esc(x.name)} · in the Metres section (open it there)</small></div><span class="tag OP">METRES</span></div>`
        : `<div class="ri"><div class="rt"><b>${esc(x.type)}</b><small>${esc(x.name)} · by ${esc(x.by || "")}</small></div><div class="row" style="gap:6px"><button class="sm" data-dview="${x.id}">Open</button><button class="sm" data-ddl="${x.id}">Download</button><button class="sm ghost" data-ddel="${x.id}" aria-label="Delete">×</button></div></div>`).join("")
        || `<p style="color:var(--bad);font-weight:600;margin:0">No documents for this shift yet. A checklist is the minimum.</p>`}</div></section>` : ""}</div>`;
}
function docForm(m = "", d = "", sh = "", t = "") {
  const c = curShift(), M = fleet().flatMap(([, L]) => L); d = d || c.key.split("|")[0]; sh = sh || c.key.split("|")[1];
  return `<form id="fDoc" class="card" autocomplete="off" style="max-width:600px;margin:0 auto"><p class="muted" style="margin-top:0">Any file: photos, PDF, Excel, Word, PowerPoint and others, up to 15 MB each. They are saved on this PC straight away and go to the TKR app by themselves when there is signal.</p>
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px">
    <label>Machine <select name="m"><option value="">General (no one machine)</option>${M.map(x => `<option ${x === m ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
    <label>Document <select name="t">${DOC_TYPES.map(x => `<option ${x === (t || "Checklist") ? "selected" : ""}>${x}</option>`).join("")}</select></label>
    <label>Date <input type="date" name="d" value="${d}" max="${today()}"></label>
    <label>Shift <select name="sh"><option ${sh === "Day" ? "selected" : ""}>Day</option><option ${sh === "Night" ? "selected" : ""}>Night</option></select></label></div>
    <label class="pick req" style="margin-top:12px;display:flex;flex-direction:column;gap:6px">Choose files <input type="file" name="f" multiple required></label>
    <button class="primary" style="width:100%;min-height:52px;margin-top:14px">Save documents</button></form>`;
}
const dataUrlOf = b => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(b); });
async function shrinkImg(file) {
  const url = URL.createObjectURL(file); try { const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
    let q = 0.75, max = 1600, out = ""; for (let k = 0; k < 6; k++) { const sc = Math.min(1, max / Math.max(img.width, img.height)), cv = document.createElement("canvas"); cv.width = Math.round(img.width * sc); cv.height = Math.round(img.height * sc);
      const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0, cv.width, cv.height); out = cv.toDataURL("image/jpeg", q); if (out.length <= 900000) break; q -= 0.1; max = Math.round(max * 0.85); }
    return out; } finally { URL.revokeObjectURL(url); }
}
/* same document for the same machine, day, shift and type uploaded again: show what is different and let them choose */
function askDup(old, file, sizeNew) {
  return new Promise(res => { document.querySelector(".ccdone")?.remove(); const d = document.createElement("div"); d.className = "ccdone"; d.setAttribute("role", "alertdialog");
    const diff = [old.name !== file.name ? `File name: <b>${esc(old.name)}</b> → <b>${esc(file.name)}</b>` : `Same file name: <b>${esc(file.name)}</b>`,
      `Size: ${sizeTxt(old.size)} → ${sizeTxt(sizeNew)}${Math.abs((old.size || 0) - sizeNew) < 1024 ? " (about the same – it may be the same document)" : ""}`,
      `First one uploaded by ${esc(old.by || "")} on ${fmt(ymd(new Date(old.editedAt || Date.now())))} ${tOf(old.editedAt || Date.now())}`];
    d.innerHTML = `<div style="max-width:460px;text-align:left"><div class="big" style="font-size:20px">Already uploaded</div><p style="margin:6px 0 10px"><b>${esc(old.machine || "General")} · ${esc(old.type)}</b> for ${fmt(old.date)} ${esc(old.shift)} shift is already there. What changed:</p>
      <ul style="margin:0 0 14px;padding-left:18px;line-height:1.6">${diff.map(x => `<li>${x}</li>`).join("")}</ul>
      <div style="display:grid;gap:8px"><button class="primary" data-a="replace">Replace the old one with the new one</button><button data-a="both">Keep both</button><button class="ghost" data-a="skip">Don't upload this file</button></div></div>`;
    d.onclick = e => { const b = e.target.closest("button[data-a]"); if (!b) return; d.remove(); res(b.dataset.a); };
    document.body.appendChild(d); }); }
async function saveDocs(f) {
  const files = [...f.elements.f.files], m = f.elements.m.value, t = f.elements.t.value, d = f.elements.d.value, sh = f.elements.sh.value; if (!files.length) return;
  const btn = f.querySelector("button.primary"); btn.disabled = true; btn.textContent = "Saving…"; let n = 0;
  for (const file of files) { try { const photo = /^image\/(jpeg|png|webp|bmp)$/.test(file.type);
      if (!photo && file.size > 15 * 1024 * 1024) { toast(file.name + " is bigger than 15 MB."); continue; }
      let data = photo ? await shrinkImg(file).catch(() => "") : ""; const mime = data ? "image/jpeg" : (file.type || mimeOf(file.name));
      if (!data) data = await dataUrlOf(new Blob([file], { type: mime }));
      const sizeNew = Math.round(data.length * 0.75), old = R("doc").filter(x => x.machine === m && x.date === d && x.shift === sh && x.type === t).sort((a, b) => (b.editedAt || 0) - (a.editedAt || 0))[0];
      let replace = null;
      if (old) { const a = await askDup(old, file, sizeNew); if (a === "skip") continue; if (a === "replace") replace = old; }
      const fid = uid(), parts = Math.max(1, Math.ceil(data.length / 700000));
      for (let i = 0; i < parts; i++) await fileSet(`${fid}_${i}`, data.slice(i * 700000, (i + 1) * 700000));
      await put({ id: "doc_" + fid, kind: "doc", ym: d.slice(0, 7), date: d, shift: sh, machine: m, group: m ? groupOf(m) : "", type: t, name: data.startsWith("data:image/jpeg") && photo && !/\.jpe?g$/i.test(file.name) ? file.name.replace(/\.[^.]+$/, "") + ".jpg" : file.name, mime, fid, parts, size: sizeNew, ...(replace ? { replaces: replace.name } : {}) }, `${replace ? "Replaced document" : "Document"}: ${t} for ${m || "general"} (${fmt(d).slice(0, 5)} ${sh})`); n++; if (replace) await put({ ...replace, deleted: true }); }
    catch (e) { console.warn(e); toast(file.name + " couldn't be read."); } }
  if (m && tab !== "files") openEv(docsPanel(m, d + "|" + sh), `${m} · documents`); else closeEv(); render(); toast(`${n} document${n === 1 ? "" : "s"} saved${navigator.onLine ? "" : " on this PC. They send when there is signal"}.`);
}
async function docData(r) { const parts = [];
  for (let i = 0; i < (r.parts || 1); i++) { let p = await fileGet(`${r.fid}_${i}`);
    if (!p && navigator.onLine) { try { await fire(); const x = await fb.getDoc(fb.doc(fb.db, "ctrl", `file_${r.fid}_${i}`)); p = x.exists() ? x.data().data : ""; if (p) await fileSet(`${r.fid}_${i}`, p); } catch (e) { console.warn(e); } }
    if (!p) return null; parts.push(p); } return parts.join(""); }
const docName = r => `${r.machine || "General"}_${r.date}_${r.shift}_${r.type.replace(/\s+/g, "_")}_${r.name}`.replace(/[\\/:*?"<>|]+/g, "-");
const MIMES = { pdf: "application/pdf", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xls: "application/vnd.ms-excel", xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12", csv: "text/csv",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", doc: "application/msword", pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation", ppt: "application/vnd.ms-powerpoint",
  txt: "text/plain", jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", heic: "image/heic", zip: "application/zip", mp4: "video/mp4" };
const mimeOf = n => MIMES[(n.split(".").pop() || "").toLowerCase()] || "application/octet-stream";
const canOpen = r => /^image\/(jpeg|png|webp|gif)$|pdf/.test(r.mime || "") || /\.(pdf|jpe?g|png)$/i.test(r.name || "");
const sizeTxt = b => !b ? "" : b < 1048576 ? Math.max(1, Math.round(b / 1024)) + " KB" : (b / 1048576).toFixed(1).replace(".", ",") + " MB";
/* the Documents tab: every uploaded file, per machine, to look up and download */
let fMach = "", fType = "", fFind = "";
function fileList() { const q = fFind.toLowerCase();
  return R("doc").filter(x => x.date?.startsWith(month) && (!fMach || (fMach === "-" ? !x.machine : x.machine === fMach)) && (!fType || x.type === fType) && (!q || `${x.name} ${x.machine} ${x.type} ${x.by}`.toLowerCase().includes(q)))
    .sort((a, b) => b.date.localeCompare(a.date) || (b.shift || "").localeCompare(a.shift || "") || (b.editedAt || 0) - (a.editedAt || 0)); }
let fSub = "files", fPick = "";
const NEED_ALL = ["Checklist", "Job card", "Fleet report", "Tally sheet", "Trip summary"];
const needList = () => { const r = R("dneed").sort((a, b) => b.editedAt - a.editedAt)[0]; return r?.list || ["Checklist", "Job card", "Fleet report"]; };
const hasDoc = (L, t, m) => L.some(x => x.type === t || (t === "Fleet report" && /^TDR/.test(m) && x.type === "Tally sheet"));   // drill rigs: the tally sheet is the fleet report
function docGaps(m0, onlyM = "") {   // every finished shift of the month, every machine that worked in it: which needed documents are not uploaded
  const need = needList(), now = Date.now(), end = m0 === today().slice(0, 7) ? today() : monthEnd(m0), M = fleet().flatMap(([, L]) => L).filter(m => !onlyM || m === onlyM), out = [];
  for (let d = m0 + "-01"; d <= end; d = addDays(d, 1)) for (const sh of ["Day", "Night"]) { const fin = shiftWin(d, sh)[1] <= now;
    const rows = !fin ? [] : M.filter(m => !offShift(m, d, sh)).map(m => { const L = docsOf(m, d, sh); return { m, have: L, miss: need.filter(t => !hasDoc(L, t, m)) }; });
    out.push({ d, sh, fut: !fin, rows, st: !fin ? "fut" : !rows.length ? "none" : rows.every(r => !r.miss.length) ? "ok" : rows.some(r => r.have.length) ? "part" : "miss" }); }
  return out; }
function missingDocsView() {
  const need = needList(), G = docGaps(month, fMach && fMach !== "-" ? fMach : ""), R2 = G.flatMap(g => g.rows.map(r => ({ ...r, d: g.d, sh: g.sh }))), gaps = R2.filter(r => r.miss.length);
  const tile = (l, n, sub, cls = "") => `<div class="tile ${cls}"><div class="l">${l}</div><div class="n">${n}</div><div class="s">${sub}</div></div>`;
  const cntT = t => gaps.filter(r => r.miss.includes(t)).length, pick = G.find(g => `${g.d}|${g.sh}` === fPick) || null;
  const byM = new Map(); for (const r of gaps) byM.set(r.m, (byM.get(r.m) || 0) + r.miss.length);
  const M = fleet();
  return `<section class="card fbar"><div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;align-items:end">
      <label>Machine <select id="fMach"><option value="">All machines</option>${M.map(([g, ms]) => `<optgroup label="${esc(g)}">${ms.map(m => `<option value="${esc(m)}" ${fMach === m ? "selected" : ""}>${esc(m)}</option>`).join("")}</optgroup>`).join("")}</select></label>
      <div><span class="muted" style="font-size:12px;letter-spacing:1px;text-transform:uppercase">Needed for every machine, every shift${me ? " · tap to change" : ""}</span><div class="row" style="gap:6px;margin-top:6px">${NEED_ALL.map(t => `<button class="sm${need.includes(t) ? " primary" : ""}" data-need="${t}" ${me ? "" : "disabled"}>${need.includes(t) ? "✓ " : ""}${t}</button>`).join("")}</div></div></div></section>`
    + `<div class="tiles">${tile("Machine shifts complete", `${R2.length - gaps.length} of ${R2.length}`, "every machine, every finished shift", gaps.length ? "alert" : "good")}${need.map(t => tile("No " + t.toLowerCase(), String(cntT(t)), "machine shifts", cntT(t) ? "alert" : "good")).join("")}</div>`
    + `<div class="legend2"><span><i style="background:var(--ok)"></i>Everything uploaded</span><span><i style="background:var(--gold)"></i>Some documents, not all</span><span><i style="background:var(--bad)"></i>Nothing uploaded</span></div>`
    + `<div class="mgrid">${G.map(g => `<button class="mcell ${g.st === "none" ? "" : g.st}" data-fpick="${g.d}|${g.sh}" ${g.fut ? "disabled" : ""} style="${pick === g ? "outline:3px solid var(--gold);outline-offset:1px" : ""}"><b>${g.d.slice(8)}</b>${g.sh}${g.rows.length && g.st !== "ok" ? `<small style="display:block;color:var(--bad)">${g.rows.filter(r => r.miss.length).length} short</small>` : ""}</button>`).join("")}</div>`
    + (pick ? `<section class="card"><div class="row" style="justify-content:space-between;align-items:center"><h3 style="margin:0">${dow(pick.d)} ${fmt(pick.d)} · ${pick.sh} shift</h3><button class="sm ghost" data-fpick="">Close</button></div>
        <div class="res" style="margin-top:10px">${pick.rows.map(r => `<div class="ri"><div class="rt"><b>${esc(r.m)}</b><small>${r.miss.length ? `<span style="color:var(--bad);font-weight:600">missing: ${r.miss.join(", ").toLowerCase()}</span>` : '<span style="color:var(--ok);font-weight:600">complete</span>'}${r.have.length ? ` · uploaded: ${[...new Set(r.have.map(x => x.type))].join(", ").toLowerCase()}` : ""}</small></div>
          ${me && r.miss.length ? `<button class="sm primary" data-upload="${esc(r.m)}|${pick.d}|${pick.sh}|${r.miss[0]}">+ ${esc(r.miss[0])}</button>` : ""}</div>`).join("") || '<p class="muted">No machines on the list.</p>'}</div></section>` : "")
    + `<div class="grid2c"><section class="card"><h2>Missing per machine</h2><div class="res">${[...byM].sort((a, b) => b[1] - a[1]).map(([m, n]) => `<div class="ri"><div class="rt"><b>${esc(m)}</b><small>${n} document${n === 1 ? "" : "s"} missing</small></div><button class="sm" data-fm="${esc(m)}">Show</button></div>`).join("") || '<p class="muted">Nothing missing this month.</p>'}</div></section>
      <section class="card"><details ${fMach ? "open" : ""}><summary style="cursor:pointer"><h2 style="display:inline">Every missing shift (${gaps.length})</h2> <span class="muted">tap to open</span></summary><div class="res miss-list" style="margin-top:10px">${gaps.slice().reverse().slice(0, 120).map(r => `<div class="ri"><div class="rt"><b>${dow(r.d)} ${fmt(r.d)} · ${r.sh} · ${esc(r.m)}</b><small>missing: ${r.miss.join(", ").toLowerCase()}</small></div>${me ? `<button class="sm" data-upload="${esc(r.m)}|${r.d}|${r.sh}|${r.miss[0]}">Upload</button>` : ""}</div>`).join("") || '<p class="muted">Every machine has its documents for every shift it worked.</p>'}${gaps.length > 120 ? `<p class="muted">Showing the latest 120. Pick a machine to see all of its gaps.</p>` : ""}</div></details></section></div>`;
}
function filesView() {
  const M = fleet(), L = fileList(), mon = R("doc").filter(x => x.date?.startsWith(month)), cnt = t => mon.filter(x => x.type === t).length;
  const mb = `<div class="mbar"><button data-mnav="-1">‹</button><b>${monthName(month)}</b><button data-mnav="1" ${month >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
  const tile = (l, n, s) => `<div class="tile"><div class="l">${l}</div><div class="n">${n}</div><div class="s">${s}</div></div>`;
  const per = M.flatMap(([, ms]) => ms).map(m => [m, mon.filter(x => x.machine === m).length]);
  return pageHead("Control room", "Documents", "Fleet reports, job cards, checklists and any other file, kept for each machine", me ? `<button class="primary" data-upload="${esc(fMach && fMach !== "-" ? fMach : "")}">+ Upload files</button>` : "",
      me ? "Upload any file for a machine. Everyone who opens the Control Room can find it here and download it." : "Find a file and tap Download. Only controllers can upload.")
    + `<div class="mhead">${mb}</div>`
    + `<div class="row" style="gap:6px;margin:0 0 12px"><button class="${fSub === "files" ? "primary" : ""}" data-fsub="files">Files</button><button class="${fSub === "miss" ? "primary" : ""}" data-fsub="miss">What's missing${(() => { const n = docGaps(month).flatMap(g => g.rows).filter(r => r.miss.length).length; return n ? ` (${n})` : ""; })()}</button></div>`
    + (fSub === "miss" ? missingDocsView() : `<div class="tiles">${tile("Files this month", mon.length, monthName(month))}${tile("Fleet reports", cnt("Fleet report"), "this month")}${tile("Job cards", cnt("Job card"), "this month")}${tile("Checklists", cnt("Checklist"), "this month")}</div>`
    + `<section class="card fbar"><div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;align-items:end">
      <label>Machine <select id="fMach"><option value="">All machines</option>${M.map(([g, ms]) => `<optgroup label="${esc(g)}">${ms.map(m => `<option value="${esc(m)}" ${fMach === m ? "selected" : ""}>${esc(m)}</option>`).join("")}</optgroup>`).join("")}<option value="-" ${fMach === "-" ? "selected" : ""}>General (no one machine)</option></select></label>
      <label>Document <select id="fType"><option value="">All documents</option>${DOC_TYPES.map(t => `<option ${fType === t ? "selected" : ""}>${t}</option>`).join("")}</select></label>
      <label>Search <input id="fFind" type="search" value="${esc(fFind)}" placeholder="File name or who uploaded"></label>
      <button data-fzip="1" ${L.length ? "" : "disabled"}>Download these (${L.length})</button></div>
      <div class="mchips" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:12px">${per.map(([m, n]) => `<button class="sm${fMach === m ? " primary" : ""}" data-fm="${esc(m)}">${esc(m)} <b>${n}</b></button>`).join("")}</div></section>`
    + `<div class="res">${L.map(x => `<div class="ri"><div class="rt"><b>${esc(x.machine || "General")} · ${esc(x.type)}</b><small>${esc(x.name)}${x.size ? " · " + sizeTxt(x.size) : ""}</small><small>${dow(x.date)} ${fmt(x.date)} · ${esc(x.shift || "")} shift · by ${esc(x.by || "")}</small></div>
        <div class="row" style="gap:6px;flex:none">${canOpen(x) ? `<button class="sm" data-dview="${x.id}">Open</button>` : ""}<button class="sm primary" data-ddl="${x.id}">Download</button>${me ? `<button class="sm ghost" data-ddel="${x.id}" aria-label="Delete">×</button>` : ""}</div></div>`).join("")
      || `<p class="muted" style="text-align:center;padding:24px 0">No files for ${fMach ? esc(fMach === "-" ? "general" : fMach) : "any machine"}${fType ? " · " + esc(fType.toLowerCase()) : ""} in ${monthName(month)}.</p>`}</div>`);
}
async function downloadList(L, root) {
  if (L.length === 1) return downloadDoc(L[0].id);
  await new Promise((res, rej) => { if (window.JSZip) return res(); const s = document.createElement("script"); s.src = "../lib/jszip.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
  const zip = new window.JSZip(); let n = 0, skip = 0;
  for (const r of L) { const d = await docData(r); if (!d) { skip++; continue; } zip.file(`${root}/${r.machine || "General"}/${docName(r)}`, d.split(",")[1], { base64: true }); n++; }
  if (!n) { toast("These files aren't on this device yet. Try again with signal."); return; }
  const { deliver } = await import("../lib/deliver.js"); deliver(await zip.generateAsync({ type: "blob" }), `${root.replace(/\s+/g, "_")}.zip`);
  toast(`${n} file${n === 1 ? "" : "s"} downloaded${skip ? ` · ${skip} not on this device yet` : ""}.`); }
async function downloadDoc(id) { const r = recs.find(x => x.id === id); if (!r) return; const d = await docData(r); if (!d) { toast(navigator.onLine ? "This document hasn't reached the TKR app yet." : "This document is on another PC. Try again with signal."); return; }
  const { deliver } = await import("../lib/deliver.js"); deliver(await (await fetch(d)).blob(), docName(r)); }
async function downloadAll(m) {
  const L = docsOf(m).filter(x => x.kind === "doc" && x.date.startsWith(month)); if (!L.length) { toast("No documents for " + m + " this month yet."); return; }
  await new Promise((res, rej) => { if (window.JSZip) return res(); const s = document.createElement("script"); s.src = "../lib/jszip.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
  const zip = new window.JSZip(), root = `${m} documents ${monthName(month)}`; let n = 0, skip = 0;
  for (const r of L) { const d = await docData(r); if (!d) { skip++; continue; } zip.file(`${root}/${r.date} ${r.shift}/${docName(r)}`, d.split(",")[1], { base64: true }); n++; }
  if (!n) { toast("These documents are on another PC. Try again with signal."); return; }
  const { deliver } = await import("../lib/deliver.js"); deliver(await zip.generateAsync({ type: "blob" }), `${root.replace(/\s+/g, "_")}.zip`);
  toast(`${n} document${n === 1 ? "" : "s"} downloaded${skip ? ` · ${skip} not on this PC yet` : ""}.`); }
async function viewDoc(id) {
  const r = recs.find(x => x.id === id); if (!r) return; if (!canOpen(r)) return downloadDoc(id); let parts = [];
  for (let i = 0; i < (r.parts || 1); i++) { let p = await fileGet(`${r.fid}_${i}`);
    if (!p) { if (!navigator.onLine) { toast("This document is on another PC. Open it when there is signal."); return; }
      try { await fire(); const x = await fb.getDoc(fb.doc(fb.db, "ctrl", `file_${r.fid}_${i}`)); p = x.exists() ? x.data().data : ""; if (p) await fileSet(`${r.fid}_${i}`, p); } catch (e) { console.warn(e); } }
    if (!p) { toast("This document hasn't reached the TKR app yet."); return; } parts.push(p); }
  const blob = await (await fetch(parts.join(""))).blob(), url = URL.createObjectURL(blob), w = window.open(url, "_blank"); if (!w) location.href = url;
}

/* ---------- events ---------- */
function cellMenu(btn) {
  document.getElementById("hpop")?.remove(); const v = btn.classList.contains("auto") || btn.classList.contains("carry") ? "" : btn.dataset.v, r = btn.getBoundingClientRect(), pop = document.createElement("div"); pop.id = "hpop"; pop.className = "hpop";
  const LB = { OP: "OP · operating", SB: "SB · available, standing", BD: "BD · breakdown", PM: "PM · maintenance" };
  pop.innerHTML = ["OP", "SB", "BD", "PM"].map(o => `<button type="button" class="ty ${o}${o === v ? " cur" : ""}" data-hpick="${o}">${LB[o]}${o === v ? " ✓" : ""}</button>`).join("") + `<button type="button" class="clr" data-hpick="">Clear</button>`
    + (v ? `<small>Tap ${v} again to clear it</small>` : "");
  pop.dataset.cell = btn.dataset.hcell; pop.dataset.cur = v; document.body.appendChild(pop);
  const w = pop.offsetWidth, h = pop.offsetHeight; pop.style.left = scrollX + Math.max(6, Math.min(innerWidth - w - 6, r.left + r.width / 2 - w / 2)) + "px"; pop.style.top = scrollY + (r.bottom + h + 8 > innerHeight ? r.top - h - 4 : r.bottom + 4) + "px";
}
document.addEventListener("click", async ev => {
  const pk = ev.target.closest("[data-hpick]"), pop = document.getElementById("hpop");
  if (pk && pop) { const [mc, h] = pop.dataset.cell.split("|"), cur = pop.dataset.cur || "", val = pk.dataset.hpick === cur ? "" : pk.dataset.hpick; pop.remove();
    if (val === cur && pk.dataset.hpick !== "") return; await setCell($("fDay").dataset.date, mc, +h, val); if (!val) toast(`${mc} ${hLabel(+h)} cleared.`); render(); return; }
  if (pop && !ev.target.closest("#hpop") && !ev.target.closest("[data-hcell]")) pop.remove();
}, true);
async function clickH(ev) {
  const vb = ev.target.closest("[data-view],[data-unview]");
  if (vb) { viewer = !!vb.dataset.view; if (!viewer) { pinFor = null; adminStep = ""; } render(); if (viewer) syncNow(false).then(() => { if (viewer && !me) render(); }); return; }
  if (!me) { if (viewer) { if (sheetClick(ev, app)) return; const b = ev.target.closest("[data-sh],[data-wbm],[data-tab],[data-ddl],[data-dview],[data-dzip],[data-fzip],[data-mnav],[data-fsub],[data-fpick],[data-fm]"); if (!b) return; } }   // viewing only: nothing can be changed
  const hb = ev.target.closest("[data-hcell]"); if (hb) { if (!me) return; cellMenu(hb); return; }
  if (sheetClick(ev, app)) return;
  const ne = ev.target.closest("[data-newev]"); if (ne) { openEv(evForm({}, ne.dataset.newev), "Log breakdown"); return; }
  const lc = ev.target.closest("[data-lockcell]"); if (lc) { const [mc, h] = lc.dataset.lockcell.split("|"); openEv(reqForm({ type: "cell", d: $("fDay").dataset.date, mc, h: +h }), "Ask the admin to change it"); return; }
  const lf = ev.target.closest("[data-lockf]"); if (lf) { const [id, field] = lf.dataset.lockf.split("|"); openEv(reqForm({ type: "field", id, field }), "Ask the admin to change it"); return; }
  const cell = ev.target.closest("[data-edit]:not(button)"); if (cell) { const e = recs.find(r => r.id === cell.dataset.edit); if (e) openEv(evForm(e), `${e.machine} · correct the times`); return; }
  const b = ev.target.closest("button"); if (!b) return;
  if (b.hasAttribute("data-x")) { closeEv(); return; }
  if (b.id === "admForgot") { const e = document.querySelector("#fAdm [name=e]")?.value.trim() || CTRL_ADMIN[0];
    try { const m = await import("../lib/firebase.js"); admApp = admApp || m.initializeApp(FB_CFG, "ctrlAdmin"); await m.sendPasswordResetEmail(m.getAuth(admApp), e); popupMsg("Reset link sent", `Check ${e} for an email from the TKR app (look in junk too). Choose a new password, then sign in here.`); }
    catch (x) { console.warn(x); toast("The reset link couldn't be sent. Check the email and signal."); } return; }
  if (b.dataset.rpin) { if (!me?.admin) return; const u = recs.find(r => r.id === b.dataset.rpin); if (!u || !confirm(`Reset ${u.name}'s PIN? Next time they tap their name they choose a new PIN.`)) return;
    await put({ ...u, pin: "", status: "approved" }, `Admin reset ${u.name}'s PIN`); syncSoon(); render(); toast(`${u.name} chooses a new PIN at the next sign-in. It reaches their PC when it has signal.`); return; }
  if (b.dataset.uok || b.dataset.uno) { if (!me?.admin) return; const u = recs.find(r => r.id === (b.dataset.uok || b.dataset.uno)); if (!u) return; const ok = !!b.dataset.uok;
    if (!ok && !confirm(`Decline ${u.name}? They won't be able to sign in.`)) return;
    await put({ ...u, status: ok ? "approved" : "declined", decidedBy: me.name, decidedAt: Date.now() }, `Admin ${ok ? "approved" : "declined"} ${u.name} as a controller`); render(); toast(ok ? `${u.name} can sign in now.` : "Declined."); return; }
  if (b.id === "addMe") { addingMe = true; render(); return; }
  if (b.id === "cancelAdd") { addingMe = false; render(); return; }
  if (b.id === "cancelPin") { pinFor = null; render(); return; }
  if (b.dataset.who) { pinFor = b.dataset.who; render(); setTimeout(() => app.querySelector("input[name=pin]")?.focus(), 50); return; }
  if (b.hasAttribute("data-logout")) { logout(b.dataset.logout === "swap"); return; }
  if (b.id === "adminIn") { adminStep = adminRec() ? "pin" : "email"; render(); return; }
  if (b.id === "cancelAdm") { adminStep = ""; render(); return; }
  if (b.id === "admReset") { adminStep = "email"; render(); return; }
  if (b.dataset.appr || b.dataset.decl) { b.disabled = true; await decide(b.dataset.appr || b.dataset.decl, !!b.dataset.appr); return; }
  if (b.id === "skipMail") { askMail = false; render(); return; }
  if (b.dataset.mail) { const e = recs.find(r => r.id === b.dataset.mail); if (e) openEv(mailForm(e), "Email this breakdown"); return; }
  if (b.dataset.mdocs) { openEv(docsPanel(b.dataset.mdocs), `${b.dataset.mdocs} · documents`); return; }
  if (b.dataset.upload !== undefined) { const [m, d, sh, t] = b.dataset.upload.split("|"); openEv(docForm(m, d, sh, t), "Upload documents"); return; }
  if (b.dataset.dshift) { const [m, d, sh] = b.dataset.dshift.split("|"); openEv(docsPanel(m, d + "|" + sh), `${m} · documents`); return; }
  if (b.dataset.dview) { await viewDoc(b.dataset.dview); return; }
  if (b.dataset.ddl) { await downloadDoc(b.dataset.ddl); return; }
  if (b.dataset.fsub) { fSub = b.dataset.fsub; fPick = ""; render(); return; }
  if (b.dataset.fpick !== undefined) { fPick = fPick === b.dataset.fpick ? "" : b.dataset.fpick; render(); return; }
  if (b.dataset.need) { if (!me) return; const L = needList(), t = b.dataset.need, nx = L.includes(t) ? L.filter(x => x !== t) : [...L, t]; if (!nx.length) { toast("Keep at least one document."); return; }
    await put({ id: "dneed", kind: "dneed", ym: "setup", list: NEED_ALL.filter(x => nx.includes(x)) }, `Documents needed per shift: ${nx.join(", ")}`); render(); return; }
  if (b.dataset.fm) { fMach = fMach === b.dataset.fm ? "" : b.dataset.fm; render(); return; }
  if (b.dataset.fzip) { b.disabled = true; try { await downloadList(fileList(), `Control Room documents ${monthName(month)}${fMach ? " " + (fMach === "-" ? "General" : fMach) : ""}${fType ? " " + fType : ""}`); } catch (e) { console.warn(e); toast("The download couldn't be made."); } b.disabled = false; return; }
  if (b.dataset.dzip) { b.disabled = true; try { await downloadAll(b.dataset.dzip); } catch (e) { console.warn(e); toast("The download couldn't be made."); } b.disabled = false; return; }
  if (b.dataset.ddel) { const r = recs.find(x => x.id === b.dataset.ddel); if (r && confirm(`Delete ${r.name}?`)) { await put({ ...r, deleted: true }, `Deleted document ${r.type} ${r.machine}`); if (r.machine && document.getElementById("evPanel") && tab !== "files") openEv(docsPanel(r.machine, r.date + "|" + r.shift), `${r.machine} · documents`); else closeEv(); render(); } return; }
  if (b.dataset.addmg) { openEv(machForm(+b.dataset.addmg, "add"), "Add a machine"); return; }
  if (b.dataset.sitemg) { openEv(siteForm(+b.dataset.sitemg), "On site / not on site"); return; }
  if (b.dataset.site) { const [m, k] = b.dataset.site.split("|"), f = document.querySelector("#fSite [name=from]"), ms = toMs(f?.value) || Date.now();
    if (ms > Date.now()) { toast("Pick a time that has passed."); return; } await setSite(m, k === "off", ms); const gi = fleet().findIndex(([, L]) => L.includes(m)); openEv(siteForm(gi), "On site / not on site"); render(); toast(`${m} ${k === "off" ? "marked not on site" : "back on site"}.`); return; }
  if (b.dataset.rmmg) { openEv(machForm(+b.dataset.rmmg, "rm"), "Remove a machine"); return; }
  if (b.dataset.tab) { tab = b.dataset.tab; closeSheet(app); render(); scrollTo(0, 0); return; }
  if (b.dataset.dnav) { date = addDays(date, +b.dataset.dnav); render(); return; }
  if (b.hasAttribute("data-today")) { date = today(); render(); return; }
  if (b.dataset.mnav) { month = shiftMonth(month, +b.dataset.mnav); render(); return; }
  if (b.dataset.sh) { if (sheetDirty && !confirm("This sheet has changes that are not saved. Leave without saving?")) return; sheetDirty = false; sheetSel = b.dataset.sh; if (sheetSel !== "month") date = sheetSel; render(); document.querySelector(".xltabs")?.scrollIntoView({ block: "start" }); return; }
  if (b.dataset.wbm) { if (sheetDirty && !confirm("This sheet has changes that are not saved. Leave without saving?")) return; sheetDirty = false; const m = shiftMonth((sheetSel === "month" ? date : sheetSel).slice(0, 7), +b.dataset.wbm), last = monthEnd(m) < today() ? monthEnd(m) : today(); date = last; sheetSel = sheetSel === "month" ? "month" : last; render(); return; }
  if (b.dataset.msg) { showMessage(shiftMessage($("fDay").dataset.date, b.dataset.msg)); return; }
  if (b.dataset.saveonly) { b.disabled = true; try { await saveShift(b.dataset.saveonly, false); } catch (x) { console.warn(x); toast("It didn't save. Try again."); } finally { b.disabled = false; } return; }
  if (b.dataset.saveshift) { b.disabled = true; b.textContent = "Saving…"; try { await saveShift(b.dataset.saveshift); } catch (x) { console.warn(x); toast("It didn't save. Try again."); } finally { b.disabled = false; } return; }
  if (b.dataset.mach !== undefined) { const m = b.dataset.mach, open = statusNow(m); openEv(evForm({}, m), open ? `${m} · ${TYPES[open.type]}` : `${m} · log what happened`); return; }
  if (b.dataset.log !== undefined) { openEv(evForm(), "Log breakdown"); return; }
  if (b.dataset.edit) { const e = recs.find(r => r.id === b.dataset.edit); openEv(evForm(e), `${e.machine} · change`); return; }
  if (b.dataset.up) { backUp(b.dataset.up); return; }
  if (b.dataset.delev && !me.admin && evLocked(recs.find(r => r.id === b.dataset.delev))) { const e = recs.find(r => r.id === b.dataset.delev), why = document.querySelector("#fEv [name=why]")?.value.trim();
    if (!why) { toast("Type why it must be deleted, then tap again."); document.querySelector("#fEv [name=why]")?.focus(); return; }
    await addReq({ target: { type: "ev", id: e.id }, del: true, what: evTxt(e), fromTxt: evTxt(e), toTxt: "Delete it", to: null, why }); return; }
  if (b.dataset.delev) { const e = recs.find(r => r.id === b.dataset.delev); if (!confirm(`Delete this ${TYPES[e.type].toLowerCase()} of ${e.machine}?`)) return; await put({ ...e, deleted: true }, `Deleted ${e.machine} ${TYPES[e.type]} from ${hhmm(e.start)}`); closeEv(); render(); toast("Deleted"); return; }
  if (b.dataset.chk) { const [d, it, k] = b.dataset.chk.split("|"), id = chkKey(d, it), c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it };
    const v = c[k] === "✓" ? "✗" : c[k] === "✗" ? "" : "✓"; await put({ ...c, [k]: v }, `${it} (${k}): ${v || "cleared"}`); const keep = app.querySelector(".sheet:not([hidden]) .sheet-body")?.scrollTop; render(); if (keep != null) { const sb = app.querySelector(".sheet:not([hidden]) .sheet-body"); if (sb) sb.scrollTop = keep; } return; }
  if (b.dataset.inv) { const [svc, dc] = b.dataset.inv.split("|"), id = invKey(month, svc, dc), c = recs.find(r => r.id === id) || { id, kind: "inv", ym: month, svc, doc: dc };
    const v = c.done === "✓" ? "✗" : c.done === "✗" ? "" : "✓"; await put({ ...c, done: v }, `${svc} – ${dc}: ${v || "cleared"}`); render(); return; }
  if (b.dataset.add) { const it = b.dataset.add, id = addKey(month, it), c = recs.find(r => r.id === id) || { id, kind: "add", ym: month, item: it };
    const v = c.done === "✓" ? "✗" : c.done === "✗" ? "" : "✓"; await put({ ...c, done: v }, `${it}: ${v || "cleared"}`); render(); return; }
  if (b.dataset.addm) { const gi = +b.dataset.addm, inp = app.querySelector(`[data-newm="${gi}"]`); if (await addMachine(gi, inp.value)) render(); return; }
  if (b.dataset.rmm) { const [, m] = b.dataset.rmm.split("|"); if (!confirm(`Take ${m} off the list? Its logged times stay.`)) return; await removeMachine(m); render(); return; }
  if (b.dataset.rmu) { if (!confirm("Remove yourself from this PC?")) return; const u = recs.find(r => r.id === b.dataset.rmu); await put({ ...u, deleted: true }); me = null; await metaSet("me", null); render(); return; }
  if (b.dataset.discard) { const n = undo.length; if (!n) return; if (!confirm("Discard everything changed since the last Save? The hour boxes, times, reasons and ticks go back to how they were.")) return;
    b.disabled = true; await discardChanges(); sheetDirty = false; render(); toast("Changes discarded. Everything is back to the last Save."); return; }
  if (b.id === "xl") { const f = $("cxFrom")?.value || month + "-01", t = $("cxTo")?.value || monthEnd(month); if (t < f) { toast("The To date must be after the From date."); return; }
    b.disabled = true; try { await exportXlsx(f, t); } catch (x) { console.warn(x); toast("The Excel file couldn't be made."); } b.disabled = false; return; }
  if (b.id === "bk") { const blob = new Blob([JSON.stringify({ kind: "tkr-control-backup", at: Date.now(), recs })], { type: "application/json" }); const { deliver } = await import("../lib/deliver.js"); deliver(blob, `TKR_Control_Room_backup_${today()}.json`); return; }
  if (b.id === "syncNow") { syncNow(true); return; }
  if (b.id === "syncOut") { await fire(); await fb.signOut(fb.auth); render(); return; }
}
async function submitH(ev) {
  const f = ev.target; ev.preventDefault();
  if (f.id === "fAdd") { const name = f.elements.name.value.trim().replace(/\s+/g, " "), p = f.elements.pin.value, p2 = f.elements.pin2.value;
    if (name.length < 3) { toast("Type your name and surname."); return; } if (!/^\d{4}$/.test(p)) { toast("The PIN must be 4 numbers."); return; } if (p !== p2) { toast("The two PINs are not the same."); return; }
    if (R("user").some(u => u.name.toLowerCase() === name.toLowerCase())) { toast("That name is already on this PC. Tap it to sign in."); addingMe = false; render(); return; }
    await put({ id: "user_" + uid(), kind: "user", ym: "setup", name, pin: await pinHash(name, p), addedAt: Date.now(), status: "pending" }, `${name} asked to be a controller`);
    addingMe = false; render(); syncSoon(); popupMsg("Sent to the admin", `${name}, the admin must approve you first. Once approved, tap your name and type your PIN.${navigator.onLine ? "" : " The request sends when this PC has signal."}`); return; }
  if (f.id === "fNewPin") { const u = recs.find(r => r.id === pinFor), p = f.elements.pin.value.trim(), p2 = f.elements.pin2.value.trim();
    if (!/^\d{4}$/.test(p)) { $("pinErr").textContent = "The PIN must be 4 numbers."; return; } if (p !== p2) { $("pinErr").textContent = "The two PINs are not the same."; return; }
    if (!u || u.pin) return; await startSession(u); await put({ ...u, pin: await pinHash(u.name, p) }, `${u.name} chose a new PIN`); pinFor = null; askMail = true; render(); toast("New PIN saved."); return; }
  if (f.id === "fPin") { const u = recs.find(r => r.id === pinFor); if (u && !okUser(u)) { $("pinErr").textContent = "The admin hasn't approved you yet."; return; }
    const pv = f.elements.pin.value.replace(/\D/g, "").slice(-4), same = x => R("user").filter(y => y.name.trim().toLowerCase() === x.name.trim().toLowerCase() && okUser(y));
    let hit = null; if (u) for (const y of same(u)) if (y.pin && y.pin === await pinHash(y.name, pv)) { hit = y; break; }   // the same name added on two PCs: either PIN works
    if (u && !hit && navigator.onLine) { $("pinErr").textContent = "Checking with the TKR app…"; try { await syncNow(false); } catch {} for (const y of same(u)) if (y.pin && y.pin === await pinHash(y.name, pv)) { hit = y; break; } }   // this PC may have an old copy
    if (hit) { await startSession(hit); pinFor = null; askMail = true; render(); }
    else { $("pinErr").textContent = "Wrong PIN. Try again."; f.elements.pin.value = ""; f.elements.pin.focus(); } return; }
  if (f.id === "fAdm") { if (!navigator.onLine) { toast("The first admin sign-in needs signal."); return; } const p = f.elements.pin.value; if (!/^\d{4}$/.test(p)) { toast("The PIN must be 4 numbers."); return; }
    try { const m = await import("../lib/firebase.js"); admApp = admApp || m.initializeApp(FB_CFG, "ctrlAdmin"); const auth = m.getAuth(admApp);
      const c = await m.signInWithEmailAndPassword(auth, toLogin(f.elements.e.value), f.elements.p.value), em = String(c?.user?.email || "").toLowerCase(); await m.signOut(auth).catch(() => {});
      if (!CTRL_ADMIN.includes(em)) { toast("Only the Control Room admin can sign in here."); return; }
      const a = await put({ id: "admin_" + em.replace(/\W+/g, ""), kind: "admin", ym: "setup", email: em, name: "Muneer (admin)", pin: await pinHash("admin|" + em, p) });
      await startSession({ id: a.id, name: a.name, admin: true }); adminStep = ""; tab = "req"; render(); toast("Signed in as admin. Next time your PIN is enough."); }
    catch (x) { console.warn(x); const c = String(x?.code || ""), msg = c.includes("too-many-requests") ? "Too many tries – Firebase has paused sign-in for a few minutes. Wait, or tap Forgot password." : c.includes("network") ? "No signal reached the TKR app. Try again with signal." : c.includes("invalid-email") ? "That email isn't right. Use muneer@tkroperations.com." : "The email or password is wrong. Use muneer@tkroperations.com and your TKR app password, or tap Forgot password.";
      if ($("admErr")) $("admErr").textContent = msg; toast(msg); } return; }
  if (f.id === "fAdmPin") { const a = adminRec(); if (a && a.pin === await pinHash("admin|" + a.email, f.elements.pin.value)) { await startSession({ id: a.id, name: a.name, admin: true }); adminStep = ""; tab = "req"; render(); }
    else { $("pinErr").textContent = "Wrong PIN. Try again."; f.elements.pin.value = ""; f.elements.pin.focus(); } return; }
  if (f.id === "fMailTo") { const v = cleanMails(f.elements.to.value); const u = recs.find(r => r.id === me.id); if (u) await put({ ...u, mailTo: v }); me.mailTo = v; askMail = false; render(); toast(v ? "Saved. Breakdowns can be emailed to " + v + "." : "Saved."); return; }
  if (f.id === "fMail") { const to = [...f.querySelectorAll("input[name=to]:checked")].map(x => x.value).concat(cleanMails(f.elements.extra.value).split(", ").filter(Boolean));
    if (!to.length) { toast("Tick or type at least one email address."); return; }
    location.href = `mailto:${encodeURIComponent([...new Set(to)].join(","))}?subject=${encodeURIComponent(f.elements.subj.value)}&body=${encodeURIComponent(f.elements.body.value)}`; closeEv(); return; }
  if (f.id === "fDoc") { await saveDocs(f); return; }
  if (f.id === "fReq") { const t = JSON.parse(f.elements.t.value), x = reqTarget(t), to = f.elements.to.value.trim(), why = f.elements.why.value.trim();
    if (to === (x.now || "")) { toast("That is the same as now."); return; } if (!why) { toast("Say why it must change."); return; }
    await addReq({ target: t, what: x.what, fromTxt: x.now || "empty", toTxt: to || "empty", to, why }); return; }
  if (f.id === "fMach") { const gi = +f.dataset.gi, v = f.elements.m.value;
    if (f.dataset.mode === "add") { if (await addMachine(gi, v)) { closeEv(); render(); } } else { if (!confirm(`Take ${v} off the list? Its logged times are kept.`)) return; await removeMachine(v); closeEv(); render(); toast(v + " removed."); } return; }
  if (f.id === "fEv") { const v = k => f.elements[k].value, s = v("start"), e = v("end");
    if (e && toMs(e) <= toMs(s)) { toast("The end time must be after the start time."); return; }
    const old = v("eid") ? recs.find(r => r.id === v("eid")) : null;
    if (old && !me.admin) { const prop = { machine: v("machine"), type: f.querySelector("input[name=type]:checked").value, start: s, end: e, reason: v("reason").trim() };
      if (Object.keys(prop).some(k => (old[k] || "") !== (prop[k] || "") && lockedV(old, k))) { const why = f.elements.why?.value.trim();
        if (!why) { toast("Say why it must change, for the admin."); f.elements.why?.focus(); return; }
        await addReq({ target: { type: "ev", id: old.id }, what: evTxt(old), fromTxt: evTxt(old), toTxt: evTxt(prop), to: prop, why }); return; } }
    const rec = { ...(old || {}), id: old?.id || "ev_" + uid(), kind: "ev", ym: s.slice(0, 7), machine: v("machine"), group: groupOf(v("machine")), type: f.querySelector("input[name=type]:checked").value, start: s, end: e, reason: v("reason").trim(), endBy: e && !old?.end ? me.name : old?.endBy || "" };
    await put(rec, `${old ? "Changed" : "Logged"} ${rec.machine}: ${TYPES[rec.type]} ${hhmm(s)}${e ? "–" + hhmm(e) : " (still going)"}${rec.reason ? " · " + rec.reason : ""}`);
    closeEv(); render(); toast(`Saved: ${rec.machine} ${TYPES[rec.type].toLowerCase()} from ${hhmm(s)}${e ? " to " + hhmm(e) : ""}.`);
    if (!old && rec.type === "BD" && mailList().length) openEv(mailForm(recs.find(x => x.id === rec.id) || rec), "Email this breakdown"); return; }
  if (f.id === "fSync") { if (!navigator.onLine) { toast("Signing in needs signal."); return; } try { await fire(); await fb.signInWithEmailAndPassword(fb.auth, toLogin(f.elements.e.value), f.elements.p.value); toast("Signed in. The control room now sends to the TKR app by itself."); } catch { toast("Email or password is wrong."); } return; }
}
async function changeH(ev) {
  if (ev.target.id === "fMach") { fMach = ev.target.value; render(); return; }
  if (ev.target.id === "fType") { fType = ev.target.value; render(); return; }
  if (!me) return;
  const t = ev.target;
  if (t.dataset.cell !== undefined) { const [mc, h] = t.dataset.cell.split("|"); await setCell($("fDay").dataset.date, mc, +h, t.value); render(); app.querySelector(`select[data-cell="${CSS.escape(t.dataset.cell)}"]`)?.focus({ preventScroll: true }); return; }
  if (t.dataset.meter) { const f = $("fDay"), tr = t.closest("table.av"), [mc, k] = t.dataset.meter.split("|"); if (await setMeter(mc, f.dataset.date, tr.dataset.shift, k, t.value)) { toast("Saved"); render(); } else render(); return; }
  if (t.dataset.reason !== undefined) { const f = $("fDay"), tr = t.closest("table.av"); await setReason(t.dataset.reason, f.dataset.date, tr.dataset.shift, t.value.trim()); toast("Saved"); return; }
  if (t.closest && t.closest("#fDay") && t.closest("tr[data-item]")) { const f = $("fDay"), d = f.dataset.date, tr = t.closest("tr[data-item]"), it = tr.dataset.item, id = chkKey(d, it);
    const c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it }; await put({ ...c, [t.dataset.c]: t.value.trim() }, `${it} (${t.dataset.c}): ${t.value.trim() || "cleared"}`);
    if (t.tagName === "SELECT") { const keep = scrollY, ws = app.querySelector(".xlsheet")?.scrollTop; render(); scrollTo(0, keep); const w = app.querySelector(".xlsheet"); if (w && ws != null) w.scrollTop = ws; } else toast("Saved"); return; }
  if (t.closest && t.closest("#fDay")) return;
  if (t.id === "dPick") { date = t.value || today(); render(); return; }
  if (t.dataset.chkc) { const [d, it] = t.dataset.chkc.split("|"), id = chkKey(d, it), c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it }; await put({ ...c, comment: t.value.trim() }, `${it}: comment`); toast("Saved"); return; }
  if (t.dataset.addc) { const [it, k] = t.dataset.addc.split("|"), id = addKey(month, it), c = recs.find(r => r.id === id) || { id, kind: "add", ym: month, item: it }; await put({ ...c, [k]: t.value.trim() }, `${it}: ${k}`); toast("Saved"); return; }
  if (t.id === "rs" && t.files[0]) { try { const d = JSON.parse(await t.files[0].text()); let n = 0; for (const r of d.recs || []) { const have = recs.find(x => x.id === r.id); if (!have || (r.editedAt || 0) > (have.editedAt || 0)) { await tx("readwrite", tt => tt.objectStore("rec").put({ ...r, synced: 0 })); n++; } } recs = await all(); bump(); render(); toast(`Backup loaded: ${n} changes added.`); } catch { toast("That file is not a Control Room backup."); } }
}
app.addEventListener("click", clickH); app.addEventListener("submit", submitH); app.addEventListener("change", changeH);
app.addEventListener("input", ev => { if (ev.target.id !== "fFind") return; fFind = ev.target.value; clearTimeout(fFind.t); const pos = ev.target.selectionStart; render(); const x = $("fFind"); if (x) { x.focus(); x.setSelectionRange(pos, pos); } }); $("who").addEventListener("click", clickH);

/* ---------- sync with the TKR app when there is signal ---------- */
const FB_CFG = { apiKey: "AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI", authDomain: "tkr-operations-app.firebaseapp.com", projectId: "tkr-operations-app",
  storageBucket: "tkr-operations-app.firebasestorage.app", messagingSenderId: "158080247079", appId: "1:158080247079:web:7cb7786d2b7f9ee172b502" };
const toLogin = v => { v = String(v || "").trim(); return v.includes("@") ? v : v.toLowerCase().replace(/[^a-z0-9]/g, "") + "@mgr.tkroperations.app"; };
let fb = null, fbUser = null; const sync = { busy: false, last: 0, msg: "" };
const waiting = () => recs.filter(r => !r.synced || r.synced < r.editedAt).length;
async function fire() {
  if (fb) return fb; const m = await import("../lib/firebase.js"), a = m.initializeApp(FB_CFG), auth = m.getAuth(a);
  fb = { ...m, auth, db: m.initializeFirestore(a, {}) };
  m.onAuthStateChanged(auth, u => { fbUser = u; if (u) syncNow(false); else if (navigator.onLine) m.signInAnonymously(auth).catch(() => {}); if (tab === "export") render(); });
  return fb;
}
let syncT = 0; const syncSoon = () => { clearTimeout(syncT); syncT = setTimeout(() => syncNow(false), 4000); };
async function syncNow(manual) {
  if (sync.busy || !navigator.onLine) { if (manual && !navigator.onLine) toast("No signal now. It sends by itself when there is signal."); return; }
  try { await fire(); } catch { return; }
  if (!fbUser) { try { fbUser = (await fb.signInAnonymously(fb.auth)).user; } catch (e) { console.warn(e); if (manual) toast("Couldn't reach the TKR app. It tries again by itself."); return; } }
  sync.busy = true;
  try { const { db, collection, doc, setDoc, getDocs, query, where } = fb;
    for (const r of recs.filter(r => r.kind === "doc" && !r.up && !r.deleted)) { let ok = true;
      for (let i = 0; i < (r.parts || 1); i++) { const data = await fileGet(`${r.fid}_${i}`); if (!data) { ok = false; break; } await setDoc(doc(db, "ctrl", `file_${r.fid}_${i}`), { kind: "file", ym: "files", data }); }
      if (ok) { r.up = 1; r.editedAt = Date.now(); r.synced = 0; await tx("readwrite", t => t.objectStore("rec").put(r)); } }
    for (const r of recs.filter(r => !r.synced || r.synced < r.editedAt)) { const { synced, pin, ...out } = r; await setDoc(doc(db, "ctrl", r.id), r.kind === "user" ? { ...out, pin } : out); r.synced = r.editedAt; await tx("readwrite", t => t.objectStore("rec").put(r)); }
    const yms = ["setup", shiftMonth(today().slice(0, 7), -1), today().slice(0, 7)]; if (!yms.includes(month)) yms.push(month);
    const snap = await getDocs(query(collection(db, "ctrl"), where("ym", "in", yms)));
    let n = 0; for (const d of snap.docs) { const r = d.data(), have = recs.find(x => x.id === r.id);
      if (!have || (r.editedAt || 0) > (have.editedAt || 0)) { const nr = { ...r, synced: r.editedAt }; await tx("readwrite", t => t.objectStore("rec").put(nr)); n++; } }
    recs = await all(); bump(); if (me) { await backfillRigs(); await busUpdate(); await seedFleetReports(); await clearAhead(); await purgeFuture(); } sync.last = Date.now(); await metaSet("syncLast", sync.last); liveWatch(yms); sync.msg = `Sent and checked ${new Date(Date.now()).toTimeString().slice(0, 5)}${n ? ` · ${n} changes from other devices` : ""}.`;
    if (manual) toast("The TKR app is up to date.");
  } catch (e) { console.warn(e); sync.msg = e?.code === "permission-denied" ? "The TKR app's Firebase rules don't allow the Control Room yet (Muneer publishes them once)." : "Couldn't reach the TKR app. It tries again by itself."; if (manual) toast(sync.msg); }
  sync.busy = false; if (me && !sheetDirty && !document.activeElement?.closest?.("#fDay") && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render();
}
/* once: the drill rigs (not on site) are on breakdown for every hour from 1 November 2026, as Muneer asked on 20/11/2026 */
async function backfillRigs() {
  const KEY = "seed_rigsbd_2026-11"; if (recs.some(r => r.id === KEY)) return;
  const RIGS = ["TDR013", "TDR014", "TDR015"], now = Date.now(), keep = undo.slice(); let n = 0;
  for (let d = "2026-11-01"; d <= today(); d = addDays(d, 1)) for (const mc of RIGS) {
    const id = `hr_${d}_${mc}`, old = recs.find(r => r.id === id && !r.deleted), cells = { ...(old?.cells || {}) }; let ch = false;
    if (d > "2026-11-20") break;   // only up to 20/11, when this was asked for
    for (let h = 6; h <= 29; h++) { if (hStart(d, h) > now) break; if (cells[h] !== "BD" && cells[h] !== "PM") { cells[h] = "BD"; ch = true; } }
    if (!ch) continue; n++;
    await put(old ? { ...old, cells } : { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells, seed: 1 }); }
  await put({ id: KEY, kind: "seed", ym: "setup", at: now }, `Drill rigs TDR013, TDR014 and TDR015 marked as breakdown for every hour from 01/11/2026 (not on site)`);
  undo = keep; await metaSet("undo", undo).catch(() => {}); bump(); }   // not part of "Discard changes"
/* once: the rockbreaker fleet reports from 01/11/2026 (date, shift, machine, start hr, stop hr, machine hrs, O operational / B breakdown / P service / S standby).
   Each shift: the machine hours as OP, the rest of the 12 hours as SB (operational), BD (breakdown) or PM (service). Boxes and hour meters already filled are kept. */
const FLEET_RB = [["2026-11-01","Day","TKRH011",8008,8018,10,"O"],["2026-11-01","Day","TKRH012",7820,7820,0,"P"],["2026-11-01","Day","TKRH015",7064,7064,0,"B"],["2026-11-01","Day","TKRH017",13069,13076,7,"B"],["2026-11-01","Day","TKRH019",4444,4444,0,"B"],["2026-11-01","Day","TKRH020",10530,10540,10,"O"],["2026-11-01","Day","TKRH021",728,738,10,"O"],["2026-11-01","Night","TKRH011",8018,8025,7,"O"],["2026-11-01","Night","TKRH012",7820,7820,0,"P"],["2026-11-01","Night","TKRH015",7064,7064,0,"B"],["2026-11-01","Night","TKRH017",13076,13076,0,"B"],["2026-11-01","Night","TKRH019",4444,4444,0,"B"],["2026-11-01","Night","TKRH020",10540,10550,10,"O"],["2026-11-01","Night","TKRH021",738,739,1,"O"],["2026-11-02","Day","TKRH011",8026,8035,9,"O"],["2026-11-02","Day","TKRH012",7820,7820,0,"P"],["2026-11-02","Day","TKRH015",7064,7064,0,"B"],["2026-11-02","Day","TKRH017",13076,13078,2,"O"],["2026-11-02","Day","TKRH019",4444,4444,0,"B"],["2026-11-02","Day","TKRH020",10540,10555,5,"O"],["2026-11-02","Day","TKRH021",738,742,4,"O"],["2026-11-02","Night","TKRH011",8035,8044,9,"O"],["2026-11-02","Night","TKRH012",null,null,0,"P"],["2026-11-02","Night","TKRH015",null,null,0,"B"],["2026-11-02","Night","TKRH017",13078,13083,5,"O"],["2026-11-02","Night","TKRH019",null,null,0,"B"],["2026-11-02","Night","TKRH020",10556,10566,10,"O"],["2026-11-02","Night","TKRH021",742,751,9,"O"],["2026-11-03","Day","TKRH011",8044,8054,10,"O"],["2026-11-03","Day","TKRH012",7820,7820,0,"P"],["2026-11-03","Day","TKRH015",7064,7064,0,"B"],["2026-11-03","Day","TKRH017",13083,13088,5,"O"],["2026-11-03","Day","TKRH019",4444,4444,0,"B"],["2026-11-03","Day","TKRH020",10564,10574,10,"O"],["2026-11-03","Day","TKRH021",751,756,5,"O"],["2026-11-03","Night","TKRH011",8054,8064,10,"O"],["2026-11-03","Night","TKRH012",7820,7820,0,"P"],["2026-11-03","Night","TKRH015",7064,7064,0,"B"],["2026-11-03","Night","TKRH017",13088,13095,7,"O"],["2026-11-03","Night","TKRH019",4444,4444,0,"B"],["2026-11-03","Night","TKRH020",10574,10584,10,"O"],["2026-11-03","Night","TKRH021",756,767,10,"O"],["2026-11-04","Day","TKRH011",8064,8072,8,"O"],["2026-11-04","Day","TKRH012",7820,7820,0,"P"],["2026-11-04","Day","TKRH015",7064,7064,0,"B"],["2026-11-04","Day","TKRH017",13095,13102,7,"O"],["2026-11-04","Day","TKRH019",4444,4444,0,"B"],["2026-11-04","Day","TKRH020",10584,10590,6,"O"],["2026-11-04","Day","TKRH021",767,774,7,"O"],["2026-11-04","Night","TKRH011",8072,8082,10,"O"],["2026-11-04","Night","TKRH012",7820,7820,0,"P"],["2026-11-04","Night","TKRH015",7064,7064,0,"B"],["2026-11-04","Night","TKRH017",13102,13110,8,"O"],["2026-11-04","Night","TKRH019",4444,4444,0,"B"],["2026-11-04","Night","TKRH020",10590,10601,11,"O"],["2026-11-04","Night","TKRH021",774,784,10,"O"],["2026-11-05","Day","TKRH011",8083,8090,7,"O"],["2026-11-05","Day","TKRH012",7820,7820,0,"P"],["2026-11-05","Day","TKRH015",7064,7064,0,"B"],["2026-11-05","Day","TKRH017",13110,13116,6,"O"],["2026-11-05","Day","TKRH019",4444,4444,0,"B"],["2026-11-05","Day","TKRH020",10600,10604,4,"O"],["2026-11-05","Day","TKRH021",785,794,9,"O"],["2026-11-05","Night","TKRH011",8091,8102,11,"O"],["2026-11-05","Night","TKRH012",7820,7820,0,"P"],["2026-11-05","Night","TKRH015",7064,7064,0,"B"],["2026-11-05","Night","TKRH017",13116,13122,6,"O"],["2026-11-05","Night","TKRH019",4444,4444,0,"B"],["2026-11-05","Night","TKRH020",10605,10614,9,"O"],["2026-11-05","Night","TKRH021",795,805,10,"O"],["2026-11-06","Day","TKRH011",8102,8110,8,"O"],["2026-11-06","Day","TKRH012",7820,7820,0,"P"],["2026-11-06","Day","TKRH015",7064,7064,0,"B"],["2026-11-06","Day","TKRH017",13122,13127,5,"O"],["2026-11-06","Day","TKRH019",4444,4444,0,"B"],["2026-11-06","Day","TKRH020",10614,10622,8,"O"],["2026-11-06","Day","TKRH021",806,809,3,"O"],["2026-11-06","Night","TKRH011",8111,8116,5,"O"],["2026-11-06","Night","TKRH012",7820,7820,0,"P"],["2026-11-06","Night","TKRH015",7064,7064,0,"B"],["2026-11-06","Night","TKRH017",13129,13134,5,"O"],["2026-11-06","Night","TKRH019",4444,4444,0,"B"],["2026-11-06","Night","TKRH020",10624,10635,11,"O"],["2026-11-06","Night","TKRH021",810,813,3,"O"],["2026-11-07","Day","TKRH011",null,null,0,"P"],["2026-11-07","Day","TKRH012",null,null,0,"P"],["2026-11-07","Day","TKRH015",null,null,0,"B"],["2026-11-07","Day","TKRH017",13134,13139,5,"O"],["2026-11-07","Day","TKRH019",null,null,0,"B"],["2026-11-07","Day","TKRH020",10635,10644,9,"O"],["2026-11-07","Day","TKRH021",814,823,9,"O"],["2026-11-07","Night","TKRH011",8116,8116,0,"O"],["2026-11-07","Night","TKRH012",7820,7820,0,"P"],["2026-11-07","Night","TKRH015",7066,7072,7,"O"],["2026-11-07","Night","TKRH017",13129,13134,5,"O"],["2026-11-07","Night","TKRH019",4444,4444,0,"B"],["2026-11-07","Night","TKRH020",10640,10655,6,"O"],["2026-11-07","Night","TKRH021",823,833,10,"O"],["2026-11-08","Day","TKRH011",null,null,0,"P"],["2026-11-08","Day","TKRH012",null,null,0,"P"],["2026-11-08","Day","TKRH015",7073,7084,11,"O"],["2026-11-08","Day","TKRH017",13145,13151,6,"O"],["2026-11-08","Day","TKRH019",null,null,0,"B"],["2026-11-08","Day","TKRH020",10650,10656,6,"O"],["2026-11-08","Day","TKRH021",833,839,6,"O"],["2026-11-08","Night","TKRH011",8116,8116,0,"P"],["2026-11-08","Night","TKRH012",7820,7820,0,"P"],["2026-11-08","Night","TKRH015",7085,7092,7,"O"],["2026-11-08","Night","TKRH017",13151,13151,0,"B"],["2026-11-08","Night","TKRH019",4444,4444,0,"B"],["2026-11-08","Night","TKRH020",10652,10662,6,"O"],["2026-11-08","Night","TKRH021",839,849,10,"O"],["2026-11-09","Day","TKRH011",8116,8116,0,"P"],["2026-11-09","Day","TKRH012",7820,7820,0,"P"],["2026-11-09","Day","TKRH015",7092,7099,7,"O"],["2026-11-09","Day","TKRH017",13151,13155,4,"O"],["2026-11-09","Day","TKRH019",4444,4444,0,"P"],["2026-11-09","Day","TKRH020",10662,10673,11,"O"],["2026-11-09","Day","TKRH021",849,859,9,"O"],["2026-11-09","Night","TKRH011",null,null,0,"P"],["2026-11-09","Night","TKRH012",null,null,0,"P"],["2026-11-09","Night","TKRH015",7099,7108,9,"O"],["2026-11-09","Night","TKRH017",13155,13160,5,"O"],["2026-11-09","Night","TKRH019",null,null,0,"P"],["2026-11-09","Night","TKRH020",10674,10683,9,"O"],["2026-11-09","Night","TKRH021",860,869,9,"O"],["2026-11-10","Day","TKRH011",8116,8116,0,"O"],["2026-11-10","Day","TKRH012",7820,7820,0,"P"],["2026-11-10","Day","TKRH015",7109,7116,7,"O"],["2026-11-10","Day","TKRH017",13160,13165,5,"O"],["2026-11-10","Day","TKRH019",4444,4444,0,"P"],["2026-11-10","Day","TKRH020",10683,10694,10,"O"],["2026-11-10","Day","TKRH021",869,879,10,"O"],["2026-11-11","Day","TKRH011",8126,8135,9,"O"],["2026-11-11","Day","TKRH012",7820,7820,0,"P"],["2026-11-11","Day","TKRH015",7117,7120,3,"O"],["2026-11-11","Day","TKRH017",13173,13178,5,"O"],["2026-11-11","Day","TKRH019",4444,4444,0,"O"],["2026-11-11","Day","TKRH020",10702,10708,6,"O"],["2026-11-11","Day","TKRH021",885,893,8,"O"],["2026-11-11","Night","TKRH011",8135,8135,0,"O"],["2026-11-11","Night","TKRH012",7820,7820,0,"P"],["2026-11-11","Night","TKRH015",7120,7130,10,"O"],["2026-11-11","Night","TKRH017",13178,13179,1,"B"],["2026-11-11","Night","TKRH019",4444,4452,8,"O"],["2026-11-11","Night","TKRH020",10708,10718,10,"O"],["2026-11-11","Night","TKRH021",893,903,10,"O"],["2026-11-12","Day","TKRH011",8135,8135,0,"O"],["2026-11-12","Day","TKRH012",7820,7820,0,"P"],["2026-11-12","Day","TKRH015",7130,7140,10,"O"],["2026-11-12","Day","TKRH017",13179,13183,4,"B"],["2026-11-12","Day","TKRH019",4452,4458,6,"O"],["2026-11-12","Day","TKRH020",10720,10729,9,"O"],["2026-11-12","Day","TKRH021",904,911,7,"O"],["2026-11-12","Night","TKRH011",8135,8135,0,"O"],["2026-11-12","Night","TKRH012",7820,7820,0,"P"],["2026-11-12","Night","TKRH015",7140,7151,10,"O"],["2026-11-12","Night","TKRH017",13183,13190,7,"O"],["2026-11-12","Night","TKRH019",4458,4469,9,"O"],["2026-11-12","Night","TKRH020",10730,10740,10,"O"],["2026-11-12","Night","TKRH021",911,921,10,"O"],["2026-11-13","Day","TKRH011",8135,8136,1,"O"],["2026-11-13","Day","TKRH012",7820,7820,0,"P"],["2026-11-13","Day","TKRH015",7151,7158,7,"O"],["2026-11-13","Day","TKRH017",13190,13192,2,"O"],["2026-11-13","Day","TKRH019",4464,4466,2,"O"],["2026-11-13","Day","TKRH020",10740,10748,8,"O"],["2026-11-13","Day","TKRH021",922,930,8,"O"],["2026-11-13","Night","TKRH011",8137,8146,9,"O"],["2026-11-13","Night","TKRH012",7820,7820,0,"P"],["2026-11-13","Night","TKRH015",7158,7158,0,"O"],["2026-11-13","Night","TKRH017",13193,13199,6,"O"],["2026-11-13","Night","TKRH019",4466,4475,9,"O"],["2026-11-13","Night","TKRH020",10749,10759,10,"O"],["2026-11-13","Night","TKRH021",931,941,10,"O"],["2026-11-14","Day","TKRH011",8146,8153,7,"O"],["2026-11-14","Day","TKRH012",7820,7820,0,"P"],["2026-11-14","Day","TKRH015",7158,7158,0,"O"],["2026-11-14","Day","TKRH017",13199,13207,8,"O"],["2026-11-14","Day","TKRH019",4475,4481,6,"O"],["2026-11-14","Day","TKRH020",10760,10767,7,"O"],["2026-11-14","Day","TKRH021",941,950,9,"O"],["2026-11-14","Night","TKRH011",8153,8162,9,"O"],["2026-11-14","Night","TKRH012",7820,7820,0,"P"],["2026-11-14","Night","TKRH015",7158,7158,0,"O"],["2026-11-14","Night","TKRH017",13207,13212,5,"O"],["2026-11-14","Night","TKRH019",4482,4485,3,"O"],["2026-11-14","Night","TKRH020",10768,10779,11,"O"],["2026-11-14","Night","TKRH021",951,961,10,"O"],["2026-11-15","Day","TKRH011",8162,8172,10,"O"],["2026-11-15","Day","TKRH015",7158,7159,1,"S"],["2026-11-15","Day","TKRH017",13212,13220,8,"O"],["2026-11-15","Day","TKRH019",4485,4491,6,"O"],["2026-11-15","Day","TKRH020",10779,10783,4,"O"],["2026-11-15","Day","TKRH021",961,970,9,"O"],["2026-11-15","Night","TKRH011",8172,8178,6,"O"],["2026-11-15","Night","TKRH015",7158,7158,0,"O"],["2026-11-15","Night","TKRH017",13220,13224,4,"O"],["2026-11-15","Night","TKRH019",4491,4491,0,"O"],["2026-11-15","Night","TKRH020",10783,10789,6,"O"],["2026-11-15","Night","TKRH021",970,980,10,"O"],["2026-11-16","Day","TKRH011",8178,8187,9,"O"],["2026-11-16","Day","TKRH017",13225,13232,7,"O"],["2026-11-16","Day","TKRH019",4491,4493,2,"O"],["2026-11-16","Day","TKRH020",10790,10800,10,"O"],["2026-11-16","Day","TKRH021",981,985,4,"O"],["2026-11-16","Night","TKRH011",8189,8195,6,"O"],["2026-11-16","Night","TKRH012",null,null,0,"P"],["2026-11-16","Night","TKRH015",7159,7166,7,"O"],["2026-11-16","Night","TKRH017",13232,13240,4,"O"],["2026-11-16","Night","TKRH019",4493,4493,0,"O"],["2026-11-16","Night","TKRH020",10800,10800,0,"B"],["2026-11-16","Night","TKRH021",985,995,9,"O"],["2026-11-17","Day","TKRH011",8195,8202,7,"O"],["2026-11-17","Day","TKRH012",7820,7820,0,"P"],["2026-11-17","Day","TKRH015",7166,7173,7,"O"],["2026-11-17","Day","TKRH017",13240,13243,3,"B"],["2026-11-17","Day","TKRH019",4493,4493,0,"S"],["2026-11-17","Day","TKRH020",10800,10803,3,"O"],["2026-11-17","Day","TKRH021",994,1001,7,"O"],["2026-11-17","Night","TKRH011",8203,8211,8,"O"],["2026-11-17","Night","TKRH015",7174,7184,10,"O"],["2026-11-17","Night","TKRH017",13244,13244,0,"B"],["2026-11-17","Night","TKRH019",4493,4494,1,"B"],["2026-11-17","Night","TKRH020",null,null,0,"S"],["2026-11-17","Night","TKRH021",1001,1011,10,"O"],["2026-11-18","Day","TKRH011",8211,8218,7,"O"],["2026-11-18","Day","TKRH012",7820,7820,0,"P"],["2026-11-18","Day","TKRH015",7184,7193,9,"O"],["2026-11-18","Day","TKRH017",13244,13248,4,"O"],["2026-11-18","Day","TKRH019",4494,4494,0,"O"],["2026-11-18","Day","TKRH020",10803,10805,2,"B"],["2026-11-18","Day","TKRH021",1012,1021,10,"O"],["2026-11-18","Night","TKRH011",8218,8223,5,"O"],["2026-11-18","Night","TKRH015",7193,7203,10,"O"],["2026-11-18","Night","TKRH017",13248,13254,6,"O"],["2026-11-18","Night","TKRH019",4495,4496,1,"B"],["2026-11-18","Night","TKRH020",null,null,0,"S"],["2026-11-18","Night","TKRH021",1021,null,5,"O"],["2026-11-19","Day","TKRH011",8221,8228,7,"O"],["2026-11-19","Day","TKRH015",7203,7203,0,"O"],["2026-11-19","Day","TKRH017",13254,13260,6,"O"],["2026-11-19","Day","TKRH019",4496,4496,0,"B"],["2026-11-19","Day","TKRH020",10806,10816,10,"O"],["2026-11-19","Day","TKRH021",1027,1036,9,"O"],["2026-11-19","Night","TKRH011",8228,8238,10,"O"],["2026-11-19","Night","TKRH015",7205,7215,10,"O"],["2026-11-19","Night","TKRH017",13260,13267,7,"O"],["2026-11-19","Night","TKRH019",4496,4496,0,"B"],["2026-11-19","Night","TKRH020",10816,10826,10,"B"],["2026-11-19","Night","TKRH021",1036,1048,10,"O"],["2026-11-20","Day","TKRH011",8239,8246,7,"O"],["2026-11-20","Day","TKRH015",7216,7224,8,"O"],["2026-11-20","Day","TKRH017",13267,13273,6,"O"],["2026-11-20","Day","TKRH019",4496,4496,0,"B"],["2026-11-20","Day","TKRH020",10826,10836,10,"O"],["2026-11-20","Day","TKRH021",1047,1055,8,"O"],["2026-11-20","Night","TKRH011",8246,8256,10,"O"],["2026-11-20","Night","TKRH015",7224,7233,9,"B"],["2026-11-20","Night","TKRH017",13273,13275,2,"O"],["2026-11-20","Night","TKRH019",4496,4496,0,"B"],["2026-11-20","Night","TKRH020",10836,10846,10,"O"],["2026-11-20","Night","TKRH021",1055,1056,10,"O"]];
async function seedFleetReports() {
  const KEY = "seed_fleet_rb_2026-11_v2"; if (recs.some(r => r.id === KEY)) return;   // v2: the Rompad supervisor openings and closings (typed) first, the fleet report photos where the chat has none
  const keep = undo.slice(), have = new Set(fleet().flatMap(([, L]) => L)), now = Date.now(); let n = 0;
  for (const [d, sh, mc, a, b, w, st] of FLEET_RB) { if (!have.has(mc)) continue;
    const id = `hr_${d}_${mc}`, old = recs.find(r => r.id === id && !r.deleted), cells = { ...(old?.cells || {}) }, H = HRS(sh), rest = st === "B" ? "BD" : st === "P" ? "PM" : "SB";
    const over = d < "2026-11-18";   // before the Control Room was in use: the reports decide; from 18/11 only empty boxes are filled
    let ch = false; H.forEach((h, k) => { if (hStart(d, h) > now) return; if (!over && cells[h] && cells[h] !== "-") return; const v = k < Math.min(w, 12) ? "OP" : rest; if (cells[h] !== v) { cells[h] = v; ch = true; } });
    const m0 = old?.meter?.[sh] || {}, m = { ...m0 }; if ((over || m.open === undefined || m.open === "") && a != null && m.open !== a) { m.open = a; ch = true; } if ((over || m.close === undefined || m.close === "") && b != null && (a == null || b >= a) && m.close !== b) { m.close = b; ch = true; }
    if (!ch) continue; n++;
    await put({ ...(old || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc }), deleted: false, cells, meter: { ...(old?.meter || {}), [sh]: m } }); }
  await put({ id: KEY, kind: "seed", ym: "setup", at: now }, `Rockbreaker openings and closings 01/11/2026 to 20/11/2026 captured from the Rompad supervisor reports (${n} machine shifts)`);
  undo = keep; await metaSet("undo", undo).catch(() => {}); bump(); }
/* once: boxes that were filled before their hour had started (20 and 21/11/2026) are emptied */
async function clearAhead() {
  const KEY = "seed_clear_ahead_2026-11-21"; if (recs.some(r => r.id === KEY)) return;
  const keep = undo.slice(), logs = R("log"); let n = 0;
  for (const r of R("hr").filter(r => r.date >= "2026-11-20" && r.date <= "2026-11-21")) {
    const cells = { ...(r.cells || {}) }; let ch = false;
    for (const k of Object.keys(cells)) { const h = +k, a = hStart(r.date, h), tag = `${r.machine} ${hLabel(h)}:`;
      const L = logs.filter(l => l.date === r.date && (l.what || "").startsWith(tag)).sort((x, y) => y.at - x.at)[0];
      if (L ? L.at < a - 5 * 6e4 : (r.date === "2026-11-21" && !r.seed)) { delete cells[k]; ch = true; n++; } }
    if (ch) await put({ ...r, cells }); }
  await put({ id: KEY, kind: "seed", ym: "setup", at: Date.now() }, n ? `${n} hour boxes that were filled before their hour started were emptied` : "");
  undo = keep; await metaSet("undo", undo).catch(() => {}); bump(); }
/* every sync: a box saved for an hour that hasn't started yet is removed (nothing may be filled ahead of time) */
async function purgeFuture() { const now = Date.now(), keep = undo.slice(); let n = 0;
  for (const r of R("hr").filter(r => r.date >= today() || r.date === addDays(today(), -1))) { const cells = { ...(r.cells || {}) }; let ch = false;
    for (const k of Object.keys(cells)) if (hStart(r.date, +k) > now) { delete cells[k]; ch = true; n++; }
    if (ch) await put({ ...r, cells }); }
  if (n) { undo = keep; await metaSet("undo", undo).catch(() => {}); bump(); } }
/* once: the bus status Muneer gave at 11:42 on 20/11/2026 */
async function busUpdate() {
  const KEY = "seed_bus_2026-11-20_1142"; if (recs.some(r => r.id === KEY)) return;
  const AT = "2026-11-20T11:42", atMs = toMs(AT), keep = undo.slice(), d = "2026-11-20", sh = "Day";
  if (Date.now() < atMs) return;
  const S = [["TKR020", "BD", "Brake lining"], ["TKR021", "OP"], ["TKR023", "BD", "Isuzu engine overhaul"], ["TKR024", "BD", "Drag link needs to be repaired"], ["TKR025", "OP"], ["TKR026", "OP"]];
  for (const [mc, st, why] of S) {
    const open = downEvs(mc).filter(e => !e.end && toMs(e.start) <= atMs);
    if (st === "OP") { for (const e of open) await put({ ...e, end: AT, endBy: me.name });
      const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id && !r.deleted) || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells: {} };
      await put({ ...rec, deleted: false, cells: { ...rec.cells, 11: "OP" } }); continue; }
    const bd = open.find(e => e.type === "BD");
    if (bd) await put({ ...bd, reason: why });
    else if (cellVal(d, mc, 11) === "BD") { const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id && !r.deleted); if (rec) await put({ ...rec, notes: { ...(rec.notes || {}), [sh]: why } }); }
    else await put({ id: "ev_" + uid(), kind: "ev", ym: d.slice(0, 7), machine: mc, group: groupOf(mc), type: "BD", start: AT, end: "", reason: why }); }
  await put({ id: KEY, kind: "seed", ym: "setup", at: Date.now() }, "Bus status at 11:42: TKR020 breakdown (brake lining), TKR021 working, TKR023 breakdown (Isuzu engine overhaul), TKR024 breakdown (drag link needs to be repaired), TKR025 and TKR026 working");
  undo = keep; await metaSet("undo", undo).catch(() => {}); bump(); }
/* live: changes made on any other PC or phone show here within seconds, dashboards included */
let liveKey = "", liveOff = null, liveT = 0;
const canRedraw = () => me && !sheetDirty && !document.activeElement?.closest?.("#fDay") && !document.getElementById("evPanel") && !document.getElementById("hpop") && !document.querySelector(".sheet:not([hidden])");
function liveWatch(yms) {
  const k = yms.join(","); if (k === liveKey && liveOff) return; liveOff?.(); liveKey = k;
  const { db, collection, query, where, onSnapshot } = fb;
  liveOff = onSnapshot(query(collection(db, "ctrl"), where("ym", "in", yms)), async snap => {
    let n = 0; const docs = snap.docChanges ? snap.docChanges().filter(c => c.type !== "removed").map(c => c.doc) : snap.docs;
    for (const d of docs) { const r = d.data(), have = recs.find(x => x.id === r.id);
      if (r.id && (!have || (r.editedAt || 0) > (have.editedAt || 0))) { await tx("readwrite", t => t.objectStore("rec").put({ ...r, synced: r.editedAt })); n++; } }
    if (!n) return; recs = await all(); bump();
    const go = () => { if (canRedraw()) render(); else { clearTimeout(liveT); liveT = setTimeout(go, 3000); } }; go();
  }, e => { console.warn("live updates stopped", e); liveOff = null; liveKey = ""; });
}
function syncCard() {
  return `<section class="card"><h2>Send to the TKR app</h2><p style="margin-top:0">Everything is sent by itself whenever this PC has signal, and changes from the other PCs come in the same way. No sign-in needed. ${waiting() ? `<b>${waiting()}</b> changes waiting on this PC.` : "Nothing waiting."}${sync.last ? ` Last sent ${fmt(ymd(new Date(sync.last)))} ${new Date(sync.last).toTimeString().slice(0, 5)}.` : ""} ${esc(sync.msg)}</p><div class="row"><button class="primary" id="syncNow">Send and check now</button></div></section>`;
}
function syncCardOld() {
  if (!fbUser) return `<section class="card"><h2>Send to the TKR app</h2><p class="muted" style="margin-top:0">Sign in once on this PC with the control room app account (needs signal once). After that, everything goes to the TKR app by itself whenever there is signal, so managers see it live. Without signal it waits on this PC.</p>
    <form id="fSync" class="grid" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))"><label>Email or username <input name="e" autocapitalize="none" required></label><label>Password <input name="p" type="password" required></label><div style="display:flex;align-items:flex-end"><button class="primary" style="width:100%">Sign in</button></div></form></section>`;
  return `<section class="card"><h2>Send to the TKR app</h2><p style="margin-top:0">Signed in as <b>${esc(fbUser.email)}</b> · ${waiting() ? `<b>${waiting()}</b> changes waiting` : "nothing waiting"}. ${esc(sync.msg)}</p><div class="row"><button class="primary" id="syncNow">Send and check now</button><button class="ghost" id="syncOut">Sign out</button></div></section>`;
}
addEventListener("online", () => syncNow(false));
setInterval(() => { closeCarried().then(n => { if (n && me) render(); }); if (navigator.onLine && !sync.busy) syncNow(false).then(() => { if (me && (tab === "export" || tab === "req")) render(); }); }, 60000);   // keep every PC up to date
setInterval(() => { if (me && tab === "now" && !sheetDirty && !document.activeElement?.closest?.("#fDay") && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render(); }, 60000);   // live times on the board

/* ---------- Excel for one month: the same sheets as the control room workbook ---------- */
async function exportXlsx(from, to) {
  const m = from.slice(0, 7), period = from === m + "-01" && to === monthEnd(m) ? monthName(m) : `${fmt(from)} to ${fmt(to)}`;
  toast("Making the Excel file…");
  await new Promise((res, rej) => { if (window.ExcelJS) return res(); const s = document.createElement("script"); s.src = "../lib/exceljs.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
  const wb = new window.ExcelJS.Workbook(); wb.creator = "TKR Operations";
  let logo = null; try { logo = wb.addImage({ buffer: await (await fetch("../icons/icon-512.png")).arrayBuffer(), extension: "png" }); } catch {}
  const fill = c => ({ type: "pattern", pattern: "solid", fgColor: { argb: c } }), thin = { style: "thin", color: { argb: "FFD9D3C5" } }, box = { top: thin, left: thin, bottom: thin, right: thin };
  const F = (o = {}) => Object.assign({ name: "Arial", size: 10 }, o);
  const COLR = { OP: ["FFE2F0D9", "FF375623"], BD: ["FFFDE2DF", "FF9C0006"], SB: ["FFFFF2CC", "FF7F6000"], PM: ["FFDDEBF7", "FF1F4E78"], NU: ["FFEDEDED", "FF595959"], "✓": ["FFE2F0D9", "FF375623"], "✗": ["FFFDE2DF", "FF9C0006"] };
  const top = (ws, n, t1, t2, t3) => { for (let r = 1; r <= 4; r++) { ws.getRow(r).height = [8, 22, 20, 18][r - 1]; for (let c = 1; c <= n; c++) ws.getCell(r, c).fill = fill("FF111111"); }
    for (let c = 1; c <= n; c++) ws.getCell(5, c).fill = fill("FFD4A12A"); ws.getRow(5).height = 4; if (logo !== null) ws.addImage(logo, { tl: { col: 0.08, row: 1.1 }, ext: { width: 54, height: 54 } });
    ws.getCell(2, 2).value = t1; ws.getCell(2, 2).font = F({ size: 15, bold: true, color: { argb: "FFFFFFFF" } }); ws.getCell(3, 2).value = t2; ws.getCell(3, 2).font = F({ size: 9, bold: true, color: { argb: "FFD4A12A" } });
    ws.getCell(4, 2).value = t3; ws.getCell(4, 2).font = F({ size: 9, color: { argb: "FFCFC8B8" } }); };
  const band = (ws, r, n, t) => { for (let c = 1; c <= n; c++) ws.getCell(r, c).fill = fill("FFD4A12A"); ws.getCell(r, 1).value = t; ws.getCell(r, 1).font = F({ bold: true }); ws.getRow(r).height = 20; };
  const head = (ws, r, L) => { L.forEach((t, i) => { const c = ws.getCell(r, i + 1); c.value = t; c.font = F({ bold: true, color: { argb: "FFFFFFFF" } }); c.fill = fill("FF262626"); c.border = box; c.alignment = { wrapText: true, vertical: "middle", horizontal: i < 2 ? "left" : "center" }; }); ws.getRow(r).height = 28; };
  const row = (ws, r, L, o = {}) => { L.forEach((v, i) => { const c = ws.getCell(r, i + 1); c.value = v === "" ? null : v; c.border = box; c.font = F({ bold: !!o.bold || (o.boldCol || []).includes(i) });
    c.alignment = { vertical: "middle", wrapText: !!o.wrap, horizontal: i < 2 || (o.left || []).includes(i) ? "left" : "center" };
    if (typeof v === "number") { if ((o.pct || []).includes(i)) c.numFmt = "0.0%"; else if ((o.fmt || {})[i]) c.numFmt = o.fmt[i]; else { c.value = v / 24; c.numFmt = "[h]:mm"; } }   // hours as hours:minutes
    if (o.fill) c.fill = fill(o.fill);
    const cc = COLR[v]; if (cc && !o.fill) { c.fill = fill(cc[0]); c.font = F({ bold: true, color: { argb: cc[1] } }); } }); };
  const days = []; for (let x = from; x <= to && (x <= today() || x === from); x = addDays(x, 1)) days.push(x);   // days still to come are left out
  const rec = id => recs.find(x => x.id === id && !x.deleted) || {};
  // 1. supporting documents per service
  let ws = wb.addWorksheet(monthName(m), { views: [{ showGridLines: false }] }); ws.columns = [20, 30, 9, 22, 16].map(w => ({ width: w }));
  top(ws, 5, `Invoicing Process – ${monthName(m)}`, "TKR OPERATIONS  ·  THARISA MINE", "Supporting documents per machine and the person we get them from");
  head(ws, 6, ["Service / machine", "Supporting documents", "Done", "Person responsible", "Ready to invoice"]); let r = 7;
  for (const [svc, docs, person] of SUPPORT) { const ready = docs.every(dc => rec(invKey(m, svc, dc)).done === "✓"); docs.forEach((dc, i) => row(ws, r++, [i ? "" : svc, dc, rec(invKey(m, svc, dc)).done || "", i ? "" : person, i ? "" : ready ? "Ready" : ""], { boldCol: [0] })); }
  // 2. additional documents
  ws = wb.addWorksheet("Additional documents", { views: [{ showGridLines: false }] }); ws.columns = [44, 8, 26, 24, 44].map(w => ({ width: w }));
  top(ws, 5, `Additional Documents for Invoicing – ${monthName(m)}`, "TKR OPERATIONS  ·  THARISA MINE", "Needed every month, together with the supporting documents per machine");
  head(ws, 6, ["Additional documents", "Done", "Person responsible", "Control Room Collector", "Notes"]); r = 7;
  for (const [it, person, note] of ADDITIONAL) { const c = rec(addKey(m, it)); row(ws, r++, [it, c.done || "", c.person || person, c.collector || "", c.notes || note || ""], { wrap: true, left: [4] }); }
  const nDone = ADDITIONAL.filter(([it]) => rec(addKey(m, it)).done === "✓").length; row(ws, r, ["All additional documents received", nDone === ADDITIONAL.length ? "Ready" : `${nDone} of ${ADDITIONAL.length}`, "", "", ""], { bold: true, fill: "FFFFF2CC" });
  // 3. monthly availability
  ws = wb.addWorksheet("Monthly Availability", { views: [{ state: "frozen", xSplit: 2, ySplit: 7, showGridLines: false }] }); ws.columns = [14, 11, ...days.map(() => 6.4), 11, 11, 11, 11, 11].map(w => ({ width: w }));
  const nc = 2 + days.length + 5; top(ws, nc, `Monthly Availability – ${period}`, "TKR OPERATIONS  ·  THARISA MINE", "Available hours per machine per day (06:00 to 06:00), with month totals");
  head(ws, 7, ["Group", "Machine", ...days.map(x => days.some(y => y.slice(0, 7) !== m) ? x.slice(8) + "/" + x.slice(5, 7) : +x.slice(8)), "Available (h:mm)", "Breakdown (h:mm)", "Planned maint. (h:mm)", "Down (h:mm)", "Availability %"]); r = 8;
  const pct = (a, d) => a + d ? a / (a + d) : "";
  const fleetT = { per: days.map(() => 0), avail: 0, BD: 0, PM: 0, SB: 0, down: 0 };
  for (const [g, L] of fleet()) { const gt = { per: days.map(() => 0), avail: 0, BD: 0, PM: 0, SB: 0, down: 0 };
    for (const mc of L) { const S = days.map(x => daySum(mc, x)), t = { avail: 0, BD: 0, PM: 0, SB: 0, down: 0 }; S.forEach((s, i) => { if (s) { gt.per[i] += s.avail; for (const k in t) { t[k] += s[k]; gt[k] += s[k]; } } });
      row(ws, r, [g, mc, ...S.map(s => s ? Math.round(s.avail * 100) / 100 : ""), t.avail, t.BD, t.PM, t.down, pct(t.avail, t.down)], { boldCol: [1] , pct: [nc - 1] }); ws.getCell(r, nc).numFmt = "0%"; r++; }
    row(ws, r, [`${g} total`, "", ...gt.per.map(v => Math.round(v * 100) / 100), gt.avail, gt.BD, gt.PM, gt.down, pct(gt.avail, gt.down)], { bold: true, fill: "FFFFF2CC" , pct: [nc - 1] }); ws.getCell(r, nc).numFmt = "0%"; r++;
    gt.per.forEach((v, i) => fleetT.per[i] += v); for (const k of ["avail", "BD", "PM", "SB", "down"]) fleetT[k] += gt[k]; }
  row(ws, r, ["Fleet total", "", ...fleetT.per.map(v => Math.round(v * 100) / 100), fleetT.avail, fleetT.BD, fleetT.PM, fleetT.down, pct(fleetT.avail, fleetT.down)], { bold: true, fill: "FFD9D9D9" , pct: [nc - 1] }); ws.getCell(r, nc).numFmt = "0%";
  // hour meters: opening and closing hours per machine per shift, and the machine hours they give
  { ws = wb.addWorksheet("Machine hours", { views: [{ state: "frozen", xSplit: 2, ySplit: 7, showGridLines: false }] }); ws.columns = [14, 11, ...days.map(() => 6.4), 13, 13].map(w => ({ width: w }));
    const nm = 2 + days.length + 2, mh = (mc, x) => ["Day", "Night"].reduce((t, sh) => { const v = meterHrs(meterOf(x, mc, sh)); return v == null ? t : (t ?? 0) + v; }, null);
    top(ws, nm, `Machine hours – ${period}`, "TKR OPERATIONS  ·  THARISA MINE", "From the opening and closing hours (hour meter) per shift, day and night together");
    head(ws, 7, ["Group", "Machine", ...days.map(x => days.some(y => y.slice(0, 7) !== m) ? x.slice(8) + "/" + x.slice(5, 7) : +x.slice(8)), "Machine (h:mm)", "Available (h:mm)"]); r = 8;
    const ft = days.map(() => 0); let fT = 0, fA = 0;
    for (const [g, L] of fleet()) { const gt = days.map(() => 0); let gT = 0, gA = 0;
      for (const mc of L) { const V = days.map(x => mh(mc, x)), T = V.reduce((t, v) => t + (v || 0), 0), A = days.reduce((t, x) => t + (daySum(mc, x)?.avail || 0), 0);
        V.forEach((v, i) => { gt[i] += v || 0; }); gT += T; gA += A;
        row(ws, r++, [g, mc, ...V.map(v => v == null ? "" : v), T || "", A || ""], { boldCol: [1] }); }
      row(ws, r++, [`${g} total`, "", ...gt.map(v => v || ""), gT || "", gA || ""], { bold: true, fill: "FFFFF2CC" }); gt.forEach((v, i) => ft[i] += v); fT += gT; fA += gA; }
    row(ws, r++, ["Fleet total", "", ...ft.map(v => v || ""), fT || "", fA || ""], { bold: true, fill: "FFD9D9D9" });
    r++; band(ws, r++, nm, "Hour meter readings per shift");
    head(ws, r++, ["Date", "Shift", "Group", "Machine", "Opening hrs", "Closing hrs", "Machine (h:mm)", "Captured by"]);
    for (const x of days) for (const sh of ["Day", "Night"]) for (const [g, L] of fleet()) for (const mc of L) { const M = meterOf(x, mc, sh); if ((M.open === undefined || M.open === "") && (M.close === undefined || M.close === "")) continue;
      row(ws, r++, [fmt(x), sh, g, mc, M.open === "" || M.open === undefined ? "" : +M.open, M.close === "" || M.close === undefined ? "" : +M.close, meterHrs(M) ?? "", recs.find(q => q.id === `hr_${x}_${mc}` && !q.deleted)?.by || ""], { boldCol: [3], fmt: { 4: "0.0", 5: "0.0" } }); } }
  // 4. delay summary per machine (Control Room), and the drill rigs as in the Metres Delay summary
  { const D = dsData(from, to), B = D.B; ws = wb.addWorksheet("Delay summary", { views: [{ showGridLines: false }] }); const nw = Math.max(9, 2 + B.rigs.length * 2);
    ws.columns = [34, 14, 13, 13, 13, 13, 13, 13, 13, 13, 13].slice(0, Math.max(nw, 9)).map(w => ({ width: w }));
    top(ws, nw, `Delay summary – ${period}`, "TKR OPERATIONS  ·  THARISA MINE", `Dates: ${fmt(D.from)} to ${fmt(D.to)}  ·  times are hours:minutes`); r = 7;
    band(ws, r++, nw, "Availability per machine (Control Room)"); head(ws, r++, ["Machine", "Group", "Available (h:mm)", "Breakdown (h:mm)", "Planned maint. (h:mm)", "Down (h:mm)", "Availability %", "Times broken down", "Times in maintenance"]);
    for (const x of D.per) { row(ws, r, [x.mc, x.g, x.avail, x.BD, x.PM, x.down, pct(x.avail, x.down), x.nBD, x.nPM], { boldCol: [0], fmt: { 7: "0", 8: "0" } , pct: [7 - 1] }); ws.getCell(r, 7).numFmt = "0.0%"; r++; }
    row(ws, r, ["Total", "", D.T.avail, D.T.BD, D.T.PM, D.T.down, pct(D.T.avail, D.T.down), D.T.nBD, D.T.nPM], { bold: true, fill: "FFFFF2CC", fmt: { 7: "0", 8: "0" } , pct: [7 - 1] }); ws.getCell(r, 7).numFmt = "0.0%"; r += 2;
    band(ws, r++, nw, "Breakdowns and maintenance per cause (Control Room)"); head(ws, r++, ["Cause", "Time (h:mm)", "Times recorded", "Machines"]);
    for (const c of D.causes) row(ws, r++, [c.cause, c.h, c.n, [...c.M].sort((x, y) => y[1] - x[1]).map(([k, v]) => `${k} ${hm(v)}`).join(" · ")], { left: [3], fmt: { 2: "0" } });
    if (!D.causes.length) row(ws, r++, ["No breakdowns or maintenance logged", "", "", ""]);
    if (B.all.length) { r++; band(ws, r++, nw, "Drill rigs – metres and delays from the shift reports (Metres section)"); head(ws, r++, ["Machine", "Shifts working", "Shifts not working", "Metres", "Delay time (h:mm)", "Available (h:mm) (Control Room)", "Availability %"]);
      for (const x of B.mach) { const av = D.avOf(x.rig); row(ws, r, [x.rig, x.working, x.down, x.metres, x.hours, av ? av.avail : "", av ? pct(av.avail, av.down) : ""], { boldCol: [0], fmt: { 1: "0", 2: "0", 3: "#,##0.0", 4: "0.0" } , pct: [7 - 1] }); ws.getCell(r, 7).numFmt = "0.0%"; r++; }
      r++; head(ws, r++, ["Delay", ...B.rigs.map(x => x + " (h:mm)"), "Total (h:mm)", ...B.rigs.map(x => "Times recorded " + x)]);
      for (const c of B.cats) row(ws, r++, [c.cat, ...c.per, c.total, ...c.times], { fmt: Object.fromEntries(B.rigs.map((_, i) => [B.rigs.length + 2 + i, "0"])) });
      row(ws, r++, ["Total", ...B.rigs.map((_, i) => B.cats.reduce((n, c) => n + c.per[i], 0)), B.total, ...B.rigs.map((_, i) => B.cats.reduce((n, c) => n + c.times[i], 0))], { bold: true, fill: "FFFFF2CC", fmt: Object.fromEntries(B.rigs.map((_, i) => [B.rigs.length + 2 + i, "0"])) });
      r++; band(ws, r++, nw, "Every delay, shift by shift"); head(ws, r++, ["Delay (as written on the shift report)", "Date", "Shift", "Machine", "Category", "Time (h:mm)"]);
      for (const l of B.lines) row(ws, r++, [l.text, fmt(l.date), l.shift, l.rig, l.cat || "", l.hours == null ? "" : l.hours], { left: [4] }); }
    r++; band(ws, r++, nw, "Every breakdown and maintenance (Control Room)"); head(ws, r++, ["Reason / comments", "Date", "Machine", "Type", "From", "To", "Time (h:mm)", "Cause"]);
    for (const l of D.lines) row(ws, r++, [l.e.reason || "", fmt(l.e.start.slice(0, 10)), l.e.machine, TYPES[l.e.type], hhmm(l.e.start), l.e.end ? hhmm(l.e.end) : "still going", l.h, l.cause], { left: [7] });
    const G = dsCharts(D); if (G.length) { const gs = wb.addWorksheet("Delay graphs", { views: [{ showGridLines: false }] }); gs.columns = Array(10).fill({ width: 12 });
      top(gs, 10, `Delay graphs – ${period}`, "TKR OPERATIONS  ·  THARISA MINE", "Availability, breakdowns and delays per machine"); let gr = 6;
      for (const c of G) { const w = 860, h = Math.round(w * c.h / c.w); gs.addImage(wb.addImage({ base64: c.png, extension: "png" }), { tl: { col: 0, row: gr }, ext: { width: w, height: h } }); gr += Math.ceil(h / 20) + 2; } } }
  // 5. one sheet per day, like the control room capture sheets
  for (const d of days) {
    ws = wb.addWorksheet(`${d.slice(8)} ${MONTHS[+d.slice(5, 7) - 1].slice(0, 3)}`, { views: [{ showGridLines: false }] });
    ws.columns = [16, 12, ...Array(12).fill(7.5), 12, 12, 30, 9, 9, 9, 9].map(w => ({ width: w }));
    top(ws, 21, `Control Room Capture – ${parse(d).getDate()} ${MONTHS[+d.slice(5, 7) - 1]} ${d.slice(0, 4)}`, "TKR OPERATIONS  ·  THARISA MINE", "Day shift 06:00 to 18:00  ·  Night shift 18:00 to 06:00");
    const day = rec("day_" + d), logs = R("log").filter(l => l.date === d).sort((a, b) => a.at - b.at), last = logs[logs.length - 1];
    row(ws, 7, ["Date", fmt(d), "Day controller", "", "", day.dayCtl || "", "", "", "Night controller", "", "", day.nightCtl || ""], { boldCol: [0, 2, 8] });
    row(ws, 8, ["Last update", last ? new Date(last.at).toTimeString().slice(0, 5) : "", "Next update", "", "", last ? new Date(last.at + 36e5).toTimeString().slice(0, 5) : "", "", "", "Updates", "", "", `${new Set(logs.map(l => new Date(l.at).getHours())).size} of 24 hours`], { boldCol: [0, 2, 8] });
    r = 10; band(ws, r++, 21, "Data and documents"); head(ws, r++, ["Section", "Item", ...COLS, "Time updated", "Updated by", "Comments"]);
    for (const [sec, items] of CHECK) items.forEach((it, i) => { const c = rec(chkKey(d, it)); row(ws, r++, [i ? "" : sec, it, ...COLS.map(k => c[k] || ""), c.editedAt ? new Date(c.editedAt).toTimeString().slice(0, 5) : "", c.by || "", c.comment || ""], { boldCol: [0], left: [7] }); });
    for (const sh of ["Day", "Night"]) {
      r++; band(ws, r++, 24, `Machine availability – ${sh} shift (${sh === "Day" ? "06:00 to 18:00" : "18:00 to 06:00"})   OP = Operating  ·  SB = Available, standing (not operating)  ·  BD = Breakdown  ·  PM = Planned maintenance`);
      head(ws, r++, ["Group", "Machine", "Opening hrs", "Closing hrs", "Machine hrs (h:mm)", ...HRS(sh).map(hLabel), "Down at", "Back up", "Reason / comments", "Avail (h:mm)", "BD (h:mm)", "PM (h:mm)", "Down (h:mm)"]);
      const [a, b] = shiftWin(d, sh);
      for (const [g, L] of fleet()) L.forEach((mc, i) => { const E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)), s = shiftSum(mc, d, sh), C = cellsOf(d, mc);
        const M3 = meterOf(d, mc, sh);
        row(ws, r++, [i ? "" : g, mc, M3.open === undefined || M3.open === "" ? "" : +M3.open, M3.close === undefined || M3.close === "" ? "" : +M3.close, meterHrs(M3) ?? "", ...HRS(sh).map(h => cellVal(d, mc, h, C)), E.map(e => toMs(e.start) < a ? tOf(a) : hhmm(e.start)).join(", "), E.map(e => e.end && toMs(e.end) <= b ? hhmm(e.end) : "").filter(Boolean).join(", "),
          [...new Set([...E.map(e => e.reason), recs.find(x => x.id === `hr_${d}_${mc}` && !x.deleted)?.notes?.[sh]].filter(Boolean))].join("; "),
          s.filled ? s.avail : "", s.filled || s.BD ? s.BD : "", s.filled || s.PM ? s.PM : "", s.filled || s.down ? s.down : ""], { boldCol: [0, 1], left: [19], fmt: { 2: "0.0", 3: "0.0" } }); });
    }
    r++; band(ws, r++, 21, "24-hour availability summary per machine"); head(ws, r++, ["Group", "Machine", "Available (h:mm)", "Breakdown (h:mm)", "Planned maint. (h:mm)", "Down (h:mm)", "Total (h:mm)", "Availability %"]);
    for (const [g, L] of fleet()) L.forEach((mc, i) => { const s = daySum(mc, d); row(ws, r, [i ? "" : g, mc, ...(s ? [s.avail, s.BD, s.PM, s.down, s.avail + s.down, s.pct ?? ""] : ["", "", "", "", "", ""])], { boldCol: [0, 1] , pct: [8 - 1] }); ws.getCell(r, 8).numFmt = "0%"; r++; });
    r++; band(ws, r++, 21, "Hourly update log"); head(ws, r++, ["Hour", "Updated", "Time done", "Controller", "Notes"]);
    for (let h = 6; h < 30; h++) { const hh = h % 24, st = toMs((h >= 24 ? addDays(d, 1) : d) + `T${pad(hh)}:00`), L = logs.filter(l => l.at >= st && l.at < st + 36e5);
      row(ws, r++, [`${pad(hh)}:00–${pad((hh + 1) % 24)}:00`, L.length ? "✓" : "", L[0] ? new Date(L[0].at).toTimeString().slice(0, 5) : "", [...new Set(L.map(l => l.by))].join(", "), L.map(l => l.what).slice(0, 4).join(" · ")], { left: [4] }); }
  }
  const blob = new Blob([await wb.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const { deliver } = await import("../lib/deliver.js"); deliver(blob, `TKR_Control_Room_${period === monthName(m) ? monthName(m).replace(" ", "_") : fmt(from).replace(/\//g, "-") + "_to_" + fmt(to).replace(/\//g, "-")}.xlsx`);
}

/* ---------- start ---------- */
/* this PC's clock: if it is wrong, the app uses the real time from the TKR app's server instead and says so */
let clockSkew = 0; const realNow = Date.now.bind(Date);
async function checkClock() { try { const r = await fetch(location.href.split("#")[0] + (location.href.includes("?") ? "&" : "?") + "t=" + realNow(), { method: "HEAD", cache: "no-store" }), h = r.headers.get("date"); if (!h) return;
  const sk = new Date(h).getTime() + 500 - realNow(); if (Math.abs(sk) < 3 * 6e4) { clockSkew = 0; Date.now = realNow; document.getElementById("clockWarn")?.remove(); return; }
  clockSkew = sk; Date.now = () => realNow() + clockSkew;
  if (!document.getElementById("clockWarn")) { const w = document.createElement("div"); w.id = "clockWarn"; w.className = "notice warn"; w.style.cssText = "margin:8px auto;max-width:1100px";
    w.innerHTML = `<b class="t">This PC's clock is wrong by ${hm(Math.abs(sk) / 36e5)} (${sk > 0 ? "behind" : "ahead"})</b>The Control Room is using the correct time from the TKR app. Fix the date and time in Windows: Settings → Time & language → Set time automatically.`; app.before(w); } } catch {} }
(async () => {
  await checkClock(); setInterval(checkClock, 10 * 6e4);
  try { recs = await all(); bump(); } catch (e) { app.innerHTML = '<div class="notice warn"><b class="t">This browser can\'t store the control room data</b>Open the app in a normal (not private) window.</div>'; return; }
  sync.last = await metaGet("syncLast") || 0;
  if (!(await metaGet("fix113"))) {               // once: drill rigs are TDR013, TDR014, TDR015 only; Muneer's example controller login removed
    const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0];
    if (f) { const G = f.groups.map(([g, L]) => [g, g === "Drill rigs" ? L.filter(m => !["TDR016", "TDR017", "TDR018", "TDR019"].includes(m)) : L]); if (JSON.stringify(G) !== JSON.stringify(f.groups)) await put({ ...f, groups: G }); }
    for (const u of R("user").filter(u => /^muneer\b/i.test(u.name))) await put({ ...u, deleted: true });
    await metaSet("fix113", 1); }
  if (!(await metaGet("fix126"))) {               // once: the water bowser (TKR08) joins the control room list
    const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0];
    if (f && !f.groups.some(([g, L]) => g === "Water bowser" || L.includes("TKR09"))) { const G = f.groups.map(([g, L]) => [g, [...L]]), i = G.findIndex(([g]) => g === "Lowbed"); G.splice(i < 0 ? G.length : i + 1, 0, ["Water bowser", ["TKR09"]]); await put({ ...f, groups: G }); }
    await metaSet("fix126", 1); }
  if (!(await metaGet("fix136"))) {               // once: TKRH012 taken off the control room list (Muneer, 16/11)
    const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0];
    if (f && f.groups.some(([, L]) => L.includes("TKRH012"))) await put({ ...f, groups: f.groups.map(([g, L]) => [g, L.filter(m => m !== "TKRH012")]) }, "Removed machine TKRH012");
    await metaSet("fix136", 1); }
  if (!(await metaGet("fix130"))) {               // once: the water bowser is TKR09, not TKR08
    const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0];
    if (f && f.groups.some(([, L]) => L.includes("TKR08"))) await put({ ...f, groups: f.groups.map(([g, L]) => [g, L.map(m => m === "TKR08" ? "TKR09" : m)]) });
    await metaSet("fix130", 1); }
  try { await closeCarried(); } catch (e) { console.warn(e); }
  try { undo = (await metaGet("undo")) || []; } catch { undo = []; }
  const s = await metaGet("me");
  if (s && s.until > Date.now()) { const u = s.admin ? R("admin").find(a => a.id === s.id) : R("user").find(u => u.id === s.id);
    if (u) { me = { ...u, admin: !!s.admin, sid: s.sid || "s_" + uid(), start: s.start || s.until - 14 * 36e5 }; if (!s.sid) await metaSet("me", { ...s, sid: me.sid, start: me.start }); } }
  try { navigator.storage?.persist?.(); } catch {}
  render(); if (navigator.onLine) fire().catch(() => {});
})();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("../sw.js", { updateViaCache: "none" }).then(r => r.update()).catch(() => {});
/* keep this PC on the newest version: check for an update each time the Control Room opens, and reload once when it arrives */
if ("serviceWorker" in navigator) { if (navigator.serviceWorker.controller) { let r = false; navigator.serviceWorker.addEventListener("controllerchange", () => { if (r) return; r = true; location.reload(); }); }
  navigator.serviceWorker.register("../sw.js", { scope: "../", updateViaCache: "none" }).then(r => r.update()).catch(() => {}); }
