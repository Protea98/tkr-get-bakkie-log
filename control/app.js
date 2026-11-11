/* TKR Control Room – works with no signal on the control room PC.
   Controllers add themselves (name + 4-digit PIN) on this device. Breakdowns, standby, planned maintenance and
   "not in use" are logged with the exact start and end time; the documents checklist and the invoicing lists are
   ticked off. Everything is saved on this device at once, and goes to the TKR app when there is signal. */
import { tabBar, pageHead, sheet, sheetClick, afterRender, closeSheet, icon } from "../lib/ui.js";

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
const nowLocal = () => { const d = new Date(); return `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const toMs = s => { if (!s) return 0; const [d, t] = s.split("T"), [y, m, dd] = d.split("-").map(Number), [h, mi] = (t || "00:00").split(":").map(Number); return new Date(y, m - 1, dd, h, mi).getTime(); };
const hhmm = s => s ? s.slice(11, 16) : "";
const hm = h => { const t = Math.round(h * 60); return `${Math.floor(t / 60)}:${pad(t % 60)}`; };
const h2 = h => (Math.round(h * 100) / 100).toFixed(2);
const today = () => ymd(new Date());
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
function toast(m) { const t = $("toast"); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3200); }

/* ---------- what the control room tracks (from the November control room workbook) ---------- */
const TYPES = { BD: "Breakdown", SB: "Standby", PM: "Planned maintenance", NU: "Not in use" };
const DEFAULT_FLEET = [["Drill rigs", ["TDR013", "TDR014", "TDR015", "TDR016", "TDR017", "TDR018", "TDR019"]], ["Rockbreakers", ["TKRH011", "TKRH012", "TKRH015", "TKRH017", "TKRH019", "TKRH020"]],
  ["Buses", ["TKR020", "TKR021", "TKR023", "TKR024", "TKR025", "TKR026"]], ["Lowbed", ["TKR06"]], ["Bakkies", ["LDV011", "LDV012"]]];
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
const idb = () => dbp || (dbp = new Promise((res, rej) => { const r = indexedDB.open("tkr-control", 1);
  r.onupgradeneeded = () => { const d = r.result; const s = d.createObjectStore("rec", { keyPath: "id" }); s.createIndex("ym", "ym"); d.createObjectStore("meta", { keyPath: "k" }); };
  r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }));
const tx = async (mode, fn) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(["rec", "meta"], mode); fn(t); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); };
const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const all = async () => { const d = await idb(); return req(d.transaction("rec").objectStore("rec").getAll()); };
const metaGet = async k => { const d = await idb(); return (await req(d.transaction("meta").objectStore("meta").get(k)))?.v; };
const metaSet = (k, v) => tx("readwrite", t => t.objectStore("meta").put({ k, v }));

/* Every record: { id, kind, ym, editedAt, by, deleted?, ... } — kinds: user, ev, chk, day, log, inv, add, fleet */
let recs = [], me = null, tab = "now", date = today(), month = today().slice(0, 7), pinFor = null, addingMe = false;
const R = k => recs.filter(r => r.kind === k && !r.deleted);
async function put(r, log) {
  const rec = { ...r, editedAt: Date.now(), by: me?.name || r.by || "", synced: 0 };
  const L = log ? { id: "log_" + uid(), kind: "log", ym: (r.date || r.start || today()).slice(0, 7), date: (r.date || r.start || today()).slice(0, 10), at: Date.now(), by: me?.name || "", what: log, editedAt: Date.now(), synced: 0 } : null;
  await tx("readwrite", t => { t.objectStore("rec").put(rec); if (L) t.objectStore("rec").put(L); });
  const i = recs.findIndex(x => x.id === rec.id); if (i >= 0) recs[i] = rec; else recs.push(rec); if (L) recs.push(L);
  syncSoon(); return rec;
}
const fleet = () => { const f = R("fleet").sort((a, b) => b.editedAt - a.editedAt)[0]; return f?.groups || DEFAULT_FLEET; };
const groupOf = m => (fleet().find(([, L]) => L.includes(m)) || ["Other"])[0];

/* ---------- hours: events overlapping a time window ---------- */
const evs = () => R("ev");
const evEnd = e => e.end ? toMs(e.end) : Date.now();
function hoursIn(machine, from, to) {
  const out = { BD: 0, SB: 0, PM: 0, NU: 0 };
  for (const e of evs()) { if (e.machine !== machine) continue; const a = Math.max(toMs(e.start), from), b = Math.min(evEnd(e), to); if (b > a) out[e.type] = (out[e.type] || 0) + (b - a) / 36e5; }
  return out;
}
/* one machine, one date (06:00 to 06:00 next day), up to now */
function daySum(machine, d) {
  const a = toMs(d + "T06:00"), b = Math.min(toMs(addDays(d, 1) + "T06:00"), Date.now()); if (b <= a) return null;
  const h = hoursIn(machine, a, b), total = (b - a) / 36e5, down = h.BD + h.PM, avail = Math.max(0, total - down - h.NU);
  return { total, ...h, down, avail, pct: avail + down > 0 ? avail / (avail + down) : null };
}
function statusNow(m) { const t = Date.now(); return evs().filter(e => e.machine === m && toMs(e.start) <= t && (!e.end || toMs(e.end) > t)).sort((a, b) => toMs(b.start) - toMs(a.start))[0] || null; }

/* ---------- controllers on this device ---------- */
async function pinHash(name, pin) {
  const s = "tkr-ctl|" + name.toLowerCase() + "|" + pin;
  try { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, "0")).join(""); }
  catch { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) | 0; return "x" + h; }
}
function gateView() {
  const users = R("user").sort((a, b) => a.name.localeCompare(b.name));
  if (addingMe || !users.length) return pageHead("Control Room", "Add yourself", "", "", "No wifi needed. Type your name and choose a 4-digit PIN you will remember.")
    + `<form id="fAdd" class="card" style="max-width:460px" autocomplete="off"><label>Name and surname <input name="name" required autocapitalize="words"></label>
      <label style="margin-top:10px;display:block">Choose a 4-digit PIN <input name="pin" class="pin" inputmode="numeric" maxlength="4" pattern="\\d{4}" type="password" required></label>
      <label style="margin-top:10px;display:block">Type the PIN again <input name="pin2" class="pin" inputmode="numeric" maxlength="4" pattern="\\d{4}" type="password" required></label>
      <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Add me</button>${users.length ? '<button type="button" class="ghost" id="cancelAdd">Back</button>' : ""}</div></form>`;
  if (pinFor) { const u = users.find(x => x.id === pinFor);
    return pageHead("Control Room", `Hello ${esc(u?.name || "")}`, "", "", "Type your PIN.")
      + `<form id="fPin" class="card" style="max-width:360px" autocomplete="off"><input name="pin" class="pin" inputmode="numeric" maxlength="4" type="password" autofocus required>
        <div class="row" style="margin-top:14px"><button class="primary" style="flex:1">Start</button><button type="button" class="ghost" id="cancelPin">Not me</button></div><p class="flag" id="pinErr"></p></form>`; }
  return pageHead("Control Room", "Who is on shift?", "", `<button class="primary" id="addMe">+ Add yourself</button>`, "Tap your name. New controller? Tap Add yourself.")
    + `<div class="who-pick">${users.map(u => `<button data-who="${u.id}"><b>${esc(u.name)}</b><small>added ${fmt(ymd(new Date(u.addedAt || u.editedAt)))}</small></button>`).join("")}</div>`;
}

/* ---------- views ---------- */
const TABS = [["now", "Live board", "deploy"], ["avail", "Availability", "hours"], ["docs", "Documents", "docs"], ["inv", "Invoicing", "money"], ["export", "Export & sync", "import"], ["settings", "Settings", "settings"]];
function render() {
  $("who").innerHTML = me ? `${esc(me.name)} · <button class="sm ghost" id="swap" style="min-height:32px;padding:2px 10px">Change controller</button>` : "";
  if (!me) { app.innerHTML = gateView(); return; }
  let body = "";
  if (tab === "now") body = nowView(); if (tab === "avail") body = availView(); if (tab === "docs") body = docsView();
  if (tab === "inv") body = invView(); if (tab === "export") body = exportView(); if (tab === "settings") body = settingsView();
  app.innerHTML = tabBar(TABS, tab) + body; afterRender(app);
}
const dateBar = () => `<div class="dnav"><button data-dnav="-1" aria-label="Day before">‹</button><input type="date" id="dPick" value="${date}" max="${today()}"><button data-dnav="1" aria-label="Next day" ${date >= today() ? "disabled" : ""}>›</button>${date !== today() ? '<button data-today class="sm">Today</button>' : ""}</div>`;
function statusBar() {
  const logs = R("log").filter(l => l.date === today()).sort((a, b) => b.at - a.at), last = logs[0];
  const dueAt = last ? last.at + 36e5 : 0, due = !last || Date.now() > dueAt, w = waiting();
  return `<div class="statusbar"><span>Last update <b>${last ? `${new Date(last.at).toTimeString().slice(0, 5)}</b> by ${esc(last.by)}` : "–</b> today"}</span>
    <span>${last ? `Next update by <b>${new Date(dueAt).toTimeString().slice(0, 5)}</b> ` : ""}<span class="${due ? "due" : "ok"}">${due ? "Update due" : "Up to date"}</span></span>
    <span>${w ? `${w} change${w === 1 ? "" : "s"} waiting for signal` : (sync.last ? `Sent to the TKR app ${new Date(sync.last).toTimeString().slice(0, 5)}` : "Saved on this PC")}</span></div>`;
}
function nowView() {
  const F = fleet(), open = evs().filter(e => !e.end).sort((a, b) => toMs(a.start) - toMs(b.start));
  const todays = evs().filter(e => e.start.slice(0, 10) === today() || !e.end || (e.end || "").slice(0, 10) === today()).sort((a, b) => toMs(b.start) - toMs(a.start));
  const logs = R("log").filter(l => l.date === today()).sort((a, b) => b.at - a.at);
  return pageHead("Control room · live", "Machine status", new Date().toLocaleDateString("en-ZA", { weekday: "long", day: "numeric", month: "long" }), `<button class="primary" data-log="">+ Log breakdown or standby</button>`, "Tap a machine to log a breakdown, standby, maintenance or not in use, with the exact time. Tap it again when it is back.")
    + statusBar()
    + `<div class="tiles">${[["BD", "Breakdowns now"], ["SB", "On standby now"], ["PM", "In maintenance now"], ["NU", "Not in use now"]].map(([k, l]) => { const n = F.flatMap(([, L]) => L).filter(m => statusNow(m)?.type === k).length; return `<div class="tile ${k === "BD" && n ? "alert" : ""}"><div class="l">${l}</div><div class="n">${n}</div></div>`; }).join("")}
      <div class="tile good"><div class="l">Working now</div><div class="n">${F.flatMap(([, L]) => L).filter(m => !statusNow(m)).length}</div><div class="s">of ${F.flatMap(([, L]) => L).length} machines</div></div></div>`
    + `<div class="fleet">${F.map(([g, L]) => `<section class="grp"><h3>${esc(g)}<span>${L.filter(m => !statusNow(m)).length}/${L.length} working</span></h3><div class="mlist">${L.map(m => { const s = statusNow(m);
        return `<button class="mach ${s ? s.type : ""}" data-mach="${esc(m)}"><b>${esc(m)}</b><small>${s ? `${TYPES[s.type]} since ${hhmm(s.start)}${s.start.slice(0, 10) !== today() ? " " + fmt(s.start.slice(0, 10)).slice(0, 5) : ""}` : "Working"}</small></button>`; }).join("")}</div></section>`).join("")}</div>`
    + `<div class="openers">${sheet("open", `Still down now (${open.length})`, open.map(evLine).join("") || '<p class="muted">Every machine is working.</p>', { icon: "!", count: open.length, cls: open.length ? "alert" : "good", sub: "breakdowns, standby and maintenance not closed yet" })}
      ${sheet("todayev", "Today's breakdowns and standby", `<div class="res">${todays.map(evLine).join("") || '<p class="muted">Nothing logged today.</p>'}</div>`, { icon: icon("bd"), count: todays.length, sub: "every event with its exact times" })}
      ${sheet("checks", "Data and documents checklist", checklist(today()), { icon: icon("docs"), sub: `${chkDone(today())} of ${CHECK.flatMap(([, i]) => i).length * COLS.length} ticked today` })}
      ${sheet("ulog", "Update log", `<div class="res">${logs.map(l => `<div class="ri"><div class="rt"><b>${new Date(l.at).toTimeString().slice(0, 5)} · ${esc(l.by)}</b><small>${esc(l.what)}</small></div></div>`).join("") || '<p class="muted">No updates yet today.</p>'}</div>`, { icon: icon("hours"), count: logs.length, sub: "every change, with the time and who did it" })}</div>`;
}
function evLine(e) {
  const dur = (evEnd(e) - toMs(e.start)) / 36e5;
  return `<div class="evrow"><span class="tag ${e.type}">${e.type}</span><div><b>${esc(e.machine)}</b> · ${TYPES[e.type]}<br><span class="t">${fmt(e.start.slice(0, 10)).slice(0, 5)} ${hhmm(e.start)} → ${e.end ? `${e.end.slice(0, 10) !== e.start.slice(0, 10) ? fmt(e.end.slice(0, 10)).slice(0, 5) + " " : ""}${hhmm(e.end)}` : "still going"} · ${hm(dur)} h</span>${e.reason ? `<br><small class="muted">${esc(e.reason)}</small>` : ""}<br><small class="muted">by ${esc(e.by)}${e.endBy && e.endBy !== e.by ? `, closed by ${esc(e.endBy)}` : ""}</small></div>
    <div class="row" style="flex-direction:column;gap:6px">${e.end ? "" : `<button class="sm primary" data-up="${e.id}">Back up now</button>`}<button class="sm" data-edit="${e.id}">Change</button></div></div>`;
}
const chkKey = (d, item) => `chk_${d}_${item.replace(/\W+/g, "")}`;
const chkDone = d => R("chk").filter(c => c.date === d).reduce((n, c) => n + COLS.filter(k => c[k] === "✓").length, 0);
function checklist(d) {
  const get = item => recs.find(r => r.id === chkKey(d, item)) || {};
  return `<p class="muted" style="margin-top:0">Tap a box: empty → ✓ received → ✗ not received → empty. The time and your name are saved with it.</p><div class="tbl card" style="padding:0"><table class="chk"><thead><tr><th>Section</th><th>Item</th>${COLS.map(c => `<th>${c}</th>`).join("")}<th>Updated</th><th>Comment</th></tr></thead><tbody>
    ${CHECK.flatMap(([sec, items]) => items.map((it, i) => { const c = get(it); return `<tr><td>${i ? "" : `<b>${esc(sec)}</b>`}</td><td>${esc(it)}</td>${COLS.map(k => `<td><button class="tick ${c[k] === "✓" ? "y" : c[k] === "✗" ? "n" : ""}" data-chk="${esc(d)}|${esc(it)}|${k}">${c[k] || " "}</button></td>`).join("")}
      <td class="muted" style="white-space:nowrap">${c.editedAt ? `${new Date(c.editedAt).toTimeString().slice(0, 5)} ${esc(c.by)}` : ""}</td><td><input data-chkc="${esc(d)}|${esc(it)}" value="${esc(c.comment || "")}" placeholder="…" style="min-width:160px"></td></tr>`; })).join("")}</tbody></table></div>`;
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
    + `<div class="openers">${sheet("month", `Month: ${monthName(d.slice(0, 7))}`, monthGrid(d.slice(0, 7)), { icon: icon("shifts"), sub: "available hours per machine per day, with month totals" })}</div>`;
}
function monthGrid(m) {
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  return `<div class="card grid30" style="padding:0"><table><thead><tr><th>Machine</th>${days.map(x => `<th>${+x.slice(8)}</th>`).join("")}<th>Avail h</th><th>BD h</th><th>PM h</th><th>SB h</th><th>Avail %</th></tr></thead><tbody>
    ${fleet().map(([g, L]) => `<tr><td colspan="${days.length + 6}" style="color:var(--gold);font-weight:700">${esc(g)}</td></tr>` + L.map(mc => { const S = days.map(x => daySum(mc, x)), t = S.reduce((a, s) => { if (s) for (const k of ["avail", "BD", "PM", "SB", "down"]) a[k] += s[k]; return a; }, { avail: 0, BD: 0, PM: 0, SB: 0, down: 0 });
      return `<tr><td><b>${esc(mc)}</b></td>${S.map(s => `<td class="${s && s.BD ? "hot" : ""}">${s ? Math.round(s.avail * 10) / 10 : ""}</td>`).join("")}<td>${h2(t.avail)}</td><td>${h2(t.BD)}</td><td>${h2(t.PM)}</td><td>${h2(t.SB)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : "–"}</b></td></tr>`; }).join("")).join("")}</tbody></table></div>`;
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
    return `<tr class="${ready ? "" : ""}"><td><b>${esc(svc)}</b><br><small class="muted">${esc(rec(invKey(m, svc, "_person")).person || person)}</small></td><td>${D.map(x => `<div class="row" style="gap:8px;margin:3px 0"><button class="tick ${x.r.done === "✓" ? "y" : x.r.done === "✗" ? "n" : ""}" data-inv="${esc(svc)}|${esc(x.dc)}">${x.r.done || " "}</button>${esc(x.dc)}</div>`).join("")}</td><td>${ready ? '<span class="tag OP">READY</span>' : ""}</td></tr>`; }).join("");
  const add = ADDITIONAL.map(([it, person, note]) => { const r = rec(addKey(m, it));
    return `<tr><td>${esc(it)}${note ? `<br><small class="muted">${esc(note)}</small>` : ""}</td><td><button class="tick ${r.done === "✓" ? "y" : r.done === "✗" ? "n" : ""}" data-add="${esc(it)}">${r.done || " "}</button></td><td>${esc(r.person || person)}</td><td><input data-addc="${esc(it)}|collector" value="${esc(r.collector || "")}" placeholder="name"></td><td><input data-addc="${esc(it)}|notes" value="${esc(r.notes || "")}" placeholder="…"></td></tr>`; }).join("");
  const nDone = ADDITIONAL.filter(([it]) => rec(addKey(m, it)).done === "✓").length;
  return pageHead("Invoicing", "Supporting documents", monthName(m), "", "Tick each document when it is received. A service shows READY when all its documents are in.") + `<div class="mhead">${mb}</div>`
    + `<div class="tiles">${(() => { const r = SUPPORT.filter(([svc, docs]) => docs.every(dc => rec(invKey(m, svc, dc)).done === "✓")).length; return `<div class="tile ${r === SUPPORT.length ? "good" : ""}"><div class="l">Services ready to invoice</div><div class="n">${r} of ${SUPPORT.length}</div></div><div class="tile ${nDone === ADDITIONAL.length ? "good" : ""}"><div class="l">Additional documents</div><div class="n">${nDone} of ${ADDITIONAL.length}</div></div>`; })()}</div>`
    + `<div class="openers">${sheet("sup", `Supporting documents per service – ${monthName(m)}`, `<div class="tbl card" style="padding:0"><table><thead><tr><th>Service / machine</th><th>Supporting documents</th><th>Ready</th></tr></thead><tbody>${sup}</tbody></table></div>`, { icon: icon("docs"), sub: "per service, with the person responsible" })}
      ${sheet("add", `Additional documents – ${monthName(m)}`, `<div class="tbl card" style="padding:0"><table><thead><tr><th>Additional document</th><th>Done</th><th>Person responsible</th><th>Control room collector</th><th>Notes</th></tr></thead><tbody>${add}</tbody></table></div>`, { icon: icon("missing"), sub: `${nDone} of ${ADDITIONAL.length} received` })}</div>`;
}
function exportView() {
  const m = month, mb = `<div class="mbar"><button data-mnav="-1">‹</button><b>${monthName(m)}</b><button data-mnav="1" ${m >= today().slice(0, 7) ? "disabled" : ""}>›</button></div>`;
  return pageHead("Export and sync", "One file when you need it", "", "", "")
    + `<section class="card dl-card"><h2>Download the month (Excel)</h2><div class="mhead">${mb}</div><p class="muted">Monthly availability, every breakdown and standby with exact times, the daily checklists, the supporting documents and the additional documents, in one file.</p><button class="primary" id="xl" style="min-height:52px">Download ${monthName(m)} Excel</button></section>`
    + syncCard()
    + `<section class="card"><h2>Backup</h2><p class="muted" style="margin-top:0">Only needed if this PC never gets signal: save a backup file and load it on another PC.</p><div class="row"><button id="bk">Save a backup file</button><label class="btn" style="display:inline-flex;align-items:center;gap:8px;border:1px solid var(--line);border-radius:10px;padding:8px 14px;cursor:pointer">Load a backup <input type="file" id="rs" accept=".json" hidden></label></div></section>`;
}
function settingsView() {
  const F = fleet(), users = R("user");
  return pageHead("Settings", "Machines and controllers", "", "", "Changes here go to every control room device with the next sync.")
    + `<section class="card"><h2>Machines</h2>${F.map(([g, L], gi) => `<div style="margin-bottom:12px"><b>${esc(g)}</b><div class="row" style="margin-top:6px">${L.map(m => `<span class="tag" style="font-size:13px;padding:4px 10px">${esc(m)} <button class="sm ghost" data-rmm="${gi}|${esc(m)}" style="min-height:28px;padding:0 6px">×</button></span>`).join("")}</div>
      <div class="row" style="margin-top:6px"><input data-newm="${gi}" placeholder="Add a machine to ${esc(g)}" style="max-width:260px"><button class="sm" data-addm="${gi}">+ Add</button></div></div>`).join("")}</section>
    <section class="card"><h2>Controllers on this PC</h2><div class="res">${users.map(u => `<div class="ri"><div class="rt"><b>${esc(u.name)}</b><small>added ${fmt(ymd(new Date(u.addedAt || u.editedAt)))}</small></div>${u.id === me.id ? `<button class="sm danger" data-rmu="${u.id}">Remove me</button>` : ""}</div>`).join("")}</div></section>`;
}

/* ---------- log / change an event ---------- */
function evForm(e = {}, machine = "") {
  const M = fleet().flatMap(([, L]) => L), isNew = !e.id, open = machine ? statusNow(machine) : null;
  if (isNew && open) { e = open; }                                 // the machine is already down: open that event to close or change it
  const st = e.start || nowLocal(), en = e.end || "";
  return `<form id="fEv" class="card" autocomplete="off" style="max-width:640px;margin:0 auto">
    ${!isNew || open ? `<div class="info">${esc(e.machine)} is logged as <b>${TYPES[e.type]}</b> since ${fmt(st.slice(0, 10))} ${hhmm(st)}.${e.end ? "" : " Tap <b>Back up now</b> when it works again."}</div>` : ""}
    <input type="hidden" name="eid" value="${esc(e.id || "")}">
    <label>Machine <select name="machine" required>${M.map(m => `<option ${m === (e.machine || machine) ? "selected" : ""}>${esc(m)}</option>`).join("")}</select></label>
    <div style="margin-top:12px"><span class="muted" style="font-size:12px;letter-spacing:1px;text-transform:uppercase;font-weight:600">What happened</span><div class="seg" style="margin-top:6px">${Object.entries(TYPES).map(([k, t]) => `<label><input type="radio" name="type" value="${k}" ${k === (e.type || "BD") ? "checked" : ""}>${t}</label>`).join("")}</div></div>
    <div class="grid" style="margin-top:12px"><label>Started (exact time) <input type="datetime-local" name="start" value="${st}" max="${nowLocal()}" required></label>
      <label>Ended (empty = still going) <input type="datetime-local" name="end" value="${en}" max="${nowLocal()}"></label></div>
    <label style="display:block;margin-top:12px">Reason / what happened <input name="reason" value="${esc(e.reason || "")}" placeholder="e.g. hydraulic pipe burst"></label>
    <div class="formbar" style="display:flex;gap:8px;flex-wrap:wrap;margin-top:16px">${e.id && !e.end ? `<button type="button" class="primary" data-up="${e.id}" style="flex:1 1 200px;min-height:52px">Back up now (${new Date().toTimeString().slice(0, 5)})</button>` : ""}
      <button class="${e.id && !e.end ? "" : "primary"}" style="flex:1 1 200px;min-height:52px">Save</button>${e.id ? `<button type="button" class="danger" data-delev="${e.id}">Delete</button>` : ""}</div></form>`;
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
  closeEv(); render(); toast(`${e.machine} is back up at ${new Date().toTimeString().slice(0, 5)}. Saved.`);
}

/* ---------- events ---------- */
async function clickH(ev) {
  if (sheetClick(ev, app)) return;
  const b = ev.target.closest("button"); if (!b) return;
  if (b.hasAttribute("data-x")) { closeEv(); return; }
  if (b.id === "addMe") { addingMe = true; render(); return; }
  if (b.id === "cancelAdd") { addingMe = false; render(); return; }
  if (b.id === "cancelPin") { pinFor = null; render(); return; }
  if (b.dataset.who) { pinFor = b.dataset.who; render(); setTimeout(() => app.querySelector("input[name=pin]")?.focus(), 50); return; }
  if (b.id === "swap") { me = null; pinFor = null; closeSheet(app); await metaSet("me", null); render(); return; }
  if (b.dataset.tab) { tab = b.dataset.tab; closeSheet(app); render(); scrollTo(0, 0); return; }
  if (b.dataset.dnav) { date = addDays(date, +b.dataset.dnav); render(); return; }
  if (b.hasAttribute("data-today")) { date = today(); render(); return; }
  if (b.dataset.mnav) { month = shiftMonth(month, +b.dataset.mnav); render(); return; }
  if (b.dataset.mach !== undefined) { const m = b.dataset.mach, open = statusNow(m); openEv(evForm({}, m), open ? `${m} · ${TYPES[open.type]}` : `${m} · log what happened`); return; }
  if (b.dataset.log !== undefined) { openEv(evForm(), "Log breakdown or standby"); return; }
  if (b.dataset.edit) { const e = recs.find(r => r.id === b.dataset.edit); openEv(evForm(e), `${e.machine} · change`); return; }
  if (b.dataset.up) { backUp(b.dataset.up); return; }
  if (b.dataset.delev) { const e = recs.find(r => r.id === b.dataset.delev); if (!confirm(`Delete this ${TYPES[e.type].toLowerCase()} of ${e.machine}?`)) return; await put({ ...e, deleted: true }, `Deleted ${e.machine} ${TYPES[e.type]} from ${hhmm(e.start)}`); closeEv(); render(); toast("Deleted"); return; }
  if (b.dataset.chk) { const [d, it, k] = b.dataset.chk.split("|"), id = chkKey(d, it), c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it };
    const v = c[k] === "✓" ? "✗" : c[k] === "✗" ? "" : "✓"; await put({ ...c, [k]: v }, `${it} (${k}): ${v || "cleared"}`); const keep = app.querySelector(".sheet:not([hidden]) .sheet-body")?.scrollTop; render(); if (keep != null) { const sb = app.querySelector(".sheet:not([hidden]) .sheet-body"); if (sb) sb.scrollTop = keep; } return; }
  if (b.dataset.inv) { const [svc, dc] = b.dataset.inv.split("|"), id = invKey(month, svc, dc), c = recs.find(r => r.id === id) || { id, kind: "inv", ym: month, svc, doc: dc };
    const v = c.done === "✓" ? "✗" : c.done === "✗" ? "" : "✓"; await put({ ...c, done: v }, `${svc} – ${dc}: ${v || "cleared"}`); render(); return; }
  if (b.dataset.add) { const it = b.dataset.add, id = addKey(month, it), c = recs.find(r => r.id === id) || { id, kind: "add", ym: month, item: it };
    const v = c.done === "✓" ? "✗" : c.done === "✗" ? "" : "✓"; await put({ ...c, done: v }, `${it}: ${v || "cleared"}`); render(); return; }
  if (b.dataset.addm) { const gi = +b.dataset.addm, inp = app.querySelector(`[data-newm="${gi}"]`), v = inp.value.trim().toUpperCase().replace(/\s+/g, ""); if (!v) return;
    const G = fleet().map(([g, L]) => [g, [...L]]); if (G.some(([, L]) => L.includes(v))) { toast(v + " is already on the list."); return; }
    G[gi][1].push(v); await put({ id: "fleet", kind: "fleet", ym: "setup", groups: G }, `Added machine ${v}`); render(); toast(v + " added."); return; }
  if (b.dataset.rmm) { const [gi, m] = b.dataset.rmm.split("|"); if (!confirm(`Take ${m} off the list? Its logged times stay.`)) return; const G = fleet().map(([g, L]) => [g, L.filter(x => x !== m)]); await put({ id: "fleet", kind: "fleet", ym: "setup", groups: G }, `Removed machine ${m}`); render(); return; }
  if (b.dataset.rmu) { if (!confirm("Remove yourself from this PC?")) return; const u = recs.find(r => r.id === b.dataset.rmu); await put({ ...u, deleted: true }); me = null; await metaSet("me", null); render(); return; }
  if (b.id === "xl") { b.disabled = true; try { await exportXlsx(month); } catch (x) { console.warn(x); toast("The Excel file couldn't be made."); } b.disabled = false; return; }
  if (b.id === "bk") { const blob = new Blob([JSON.stringify({ kind: "tkr-control-backup", at: Date.now(), recs })], { type: "application/json" }); const { deliver } = await import("../lib/deliver.js"); deliver(blob, `TKR_Control_Room_backup_${today()}.json`); return; }
  if (b.id === "syncNow") { syncNow(true); return; }
  if (b.id === "syncOut") { await fire(); await fb.signOut(fb.auth); render(); return; }
}
async function submitH(ev) {
  const f = ev.target; ev.preventDefault();
  if (f.id === "fAdd") { const name = f.elements.name.value.trim().replace(/\s+/g, " "), p = f.elements.pin.value, p2 = f.elements.pin2.value;
    if (name.length < 3) { toast("Type your name and surname."); return; } if (!/^\d{4}$/.test(p)) { toast("The PIN must be 4 numbers."); return; } if (p !== p2) { toast("The two PINs are not the same."); return; }
    if (R("user").some(u => u.name.toLowerCase() === name.toLowerCase())) { toast("That name is already on this PC. Tap it to sign in."); addingMe = false; render(); return; }
    const u = await put({ id: "user_" + uid(), kind: "user", ym: "setup", name, pin: await pinHash(name, p), addedAt: Date.now() });
    me = u; addingMe = false; await metaSet("me", { id: u.id, until: Date.now() + 14 * 36e5 }); tab = "now"; render(); toast(`Welcome ${name}. You are on shift.`); return; }
  if (f.id === "fPin") { const u = recs.find(r => r.id === pinFor); if (u && u.pin === await pinHash(u.name, f.elements.pin.value)) { me = u; pinFor = null; await metaSet("me", { id: u.id, until: Date.now() + 14 * 36e5 }); render(); }
    else { $("pinErr").textContent = "Wrong PIN. Try again."; f.elements.pin.value = ""; f.elements.pin.focus(); } return; }
  if (f.id === "fEv") { const v = k => f.elements[k].value, s = v("start"), e = v("end");
    if (e && toMs(e) <= toMs(s)) { toast("The end time must be after the start time."); return; }
    const old = v("eid") ? recs.find(r => r.id === v("eid")) : null;
    const rec = { ...(old || {}), id: old?.id || "ev_" + uid(), kind: "ev", ym: s.slice(0, 7), machine: v("machine"), group: groupOf(v("machine")), type: f.querySelector("input[name=type]:checked").value, start: s, end: e, reason: v("reason").trim(), endBy: e && !old?.end ? me.name : old?.endBy || "" };
    await put(rec, `${old ? "Changed" : "Logged"} ${rec.machine}: ${TYPES[rec.type]} ${hhmm(s)}${e ? "–" + hhmm(e) : " (still going)"}${rec.reason ? " · " + rec.reason : ""}`);
    closeEv(); render(); toast(`Saved: ${rec.machine} ${TYPES[rec.type].toLowerCase()} from ${hhmm(s)}${e ? " to " + hhmm(e) : ""}.`); return; }
  if (f.id === "fSync") { if (!navigator.onLine) { toast("Signing in needs signal."); return; } try { await fire(); await fb.signInWithEmailAndPassword(fb.auth, toLogin(f.elements.e.value), f.elements.p.value); toast("Signed in. The control room now sends to the TKR app by itself."); } catch { toast("Email or password is wrong."); } return; }
}
async function changeH(ev) {
  const t = ev.target;
  if (t.id === "dPick") { date = t.value || today(); render(); return; }
  if (t.dataset.chkc) { const [d, it] = t.dataset.chkc.split("|"), id = chkKey(d, it), c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it }; await put({ ...c, comment: t.value.trim() }, `${it}: comment`); toast("Saved"); return; }
  if (t.dataset.addc) { const [it, k] = t.dataset.addc.split("|"), id = addKey(month, it), c = recs.find(r => r.id === id) || { id, kind: "add", ym: month, item: it }; await put({ ...c, [k]: t.value.trim() }, `${it}: ${k}`); toast("Saved"); return; }
  if (t.id === "rs" && t.files[0]) { try { const d = JSON.parse(await t.files[0].text()); let n = 0; for (const r of d.recs || []) { const have = recs.find(x => x.id === r.id); if (!have || (r.editedAt || 0) > (have.editedAt || 0)) { await tx("readwrite", tt => tt.objectStore("rec").put({ ...r, synced: 0 })); n++; } } recs = await all(); render(); toast(`Backup loaded: ${n} changes added.`); } catch { toast("That file is not a Control Room backup."); } }
}
app.addEventListener("click", clickH); app.addEventListener("submit", submitH); app.addEventListener("change", changeH);

/* ---------- sync with the TKR app when there is signal ---------- */
const FB_CFG = { apiKey: "AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI", authDomain: "tkr-operations-app.firebaseapp.com", projectId: "tkr-operations-app",
  storageBucket: "tkr-operations-app.firebasestorage.app", messagingSenderId: "158080247079", appId: "1:158080247079:web:7cb7786d2b7f9ee172b502" };
const toLogin = v => { v = String(v || "").trim(); return v.includes("@") ? v : v.toLowerCase().replace(/[^a-z0-9]/g, "") + "@mgr.tkroperations.app"; };
let fb = null, fbUser = null; const sync = { busy: false, last: 0, msg: "" };
const waiting = () => recs.filter(r => !r.synced || r.synced < r.editedAt).length;
async function fire() {
  if (fb) return fb; const m = await import("../lib/firebase.js"), a = m.initializeApp(FB_CFG), auth = m.getAuth(a);
  fb = { ...m, auth, db: m.initializeFirestore(a, {}) };
  m.onAuthStateChanged(auth, u => { fbUser = u; if (u) syncNow(false); if (tab === "export") render(); });
  return fb;
}
let syncT = 0; const syncSoon = () => { clearTimeout(syncT); syncT = setTimeout(() => syncNow(false), 4000); };
async function syncNow(manual) {
  if (sync.busy || !navigator.onLine) { if (manual && !navigator.onLine) toast("No signal now. It sends by itself when there is signal."); return; }
  try { await fire(); } catch { return; } if (!fbUser) { if (manual) toast("Sign in under Export & sync first."); return; }
  sync.busy = true;
  try { const { db, collection, doc, setDoc, getDocs, query, where } = fb;
    for (const r of recs.filter(r => !r.synced || r.synced < r.editedAt)) { const { synced, pin, ...out } = r; await setDoc(doc(db, "ctrl", r.id), r.kind === "user" ? { ...out, pin } : out); r.synced = r.editedAt; await tx("readwrite", t => t.objectStore("rec").put(r)); }
    const yms = ["setup", shiftMonth(today().slice(0, 7), -1), today().slice(0, 7)]; if (!yms.includes(month)) yms.push(month);
    const snap = await getDocs(query(collection(db, "ctrl"), where("ym", "in", yms)));
    let n = 0; for (const d of snap.docs) { const r = d.data(), have = recs.find(x => x.id === r.id);
      if (!have || (r.editedAt || 0) > (have.editedAt || 0)) { const nr = { ...r, synced: r.editedAt }; await tx("readwrite", t => t.objectStore("rec").put(nr)); n++; } }
    recs = await all(); sync.last = Date.now(); await metaSet("syncLast", sync.last); sync.msg = `Sent and checked ${new Date().toTimeString().slice(0, 5)}${n ? ` · ${n} changes from other devices` : ""}.`;
    if (manual) toast("The TKR app is up to date.");
  } catch (e) { console.warn(e); sync.msg = e?.code === "permission-denied" ? "This sign-in can't save Control Room data. Ask Muneer to add it." : "Couldn't reach the TKR app. It tries again by itself."; if (manual) toast(sync.msg); }
  sync.busy = false; if (me && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render();
}
function syncCard() {
  if (!fbUser) return `<section class="card"><h2>Send to the TKR app</h2><p class="muted" style="margin-top:0">Sign in once on this PC with the control room app account (needs signal once). After that, everything goes to the TKR app by itself whenever there is signal, so managers see it live. Without signal it waits on this PC.</p>
    <form id="fSync" class="grid" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))"><label>Email or username <input name="e" autocapitalize="none" required></label><label>Password <input name="p" type="password" required></label><div style="display:flex;align-items:flex-end"><button class="primary" style="width:100%">Sign in</button></div></form></section>`;
  return `<section class="card"><h2>Send to the TKR app</h2><p style="margin-top:0">Signed in as <b>${esc(fbUser.email)}</b> · ${waiting() ? `<b>${waiting()}</b> changes waiting` : "nothing waiting"}. ${esc(sync.msg)}</p><div class="row"><button class="primary" id="syncNow">Send and check now</button><button class="ghost" id="syncOut">Sign out</button></div></section>`;
}
addEventListener("online", () => syncNow(false));
setInterval(() => { if (me && tab === "now" && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render(); }, 60000);   // live times on the board

/* ---------- Excel for one month ---------- */
async function exportXlsx(m) {
  toast("Making the Excel file…");
  await new Promise((res, rej) => { if (window.ExcelJS) return res(); const s = document.createElement("script"); s.src = "../lib/exceljs.min.js"; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
  const wb = new window.ExcelJS.Workbook(); wb.creator = "TKR Operations";
  let logo = null; try { logo = wb.addImage({ buffer: await (await fetch("../icons/icon-512.png")).arrayBuffer(), extension: "png" }); } catch {}
  const fill = c => ({ type: "pattern", pattern: "solid", fgColor: { argb: c } }), thin = { style: "thin", color: { argb: "FFD9D3C5" } }, box = { top: thin, left: thin, bottom: thin, right: thin };
  const F = (o = {}) => Object.assign({ name: "Arial", size: 10 }, o);
  const sheetH = (name, widths, title, sub) => { const ws = wb.addWorksheet(name, { views: [{ state: "frozen", ySplit: 5, showGridLines: false }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = widths.map(w => ({ width: w })); for (let r = 1; r <= 4; r++) { ws.getRow(r).height = [8, 22, 28, 18][r - 1]; for (let c = 1; c <= Math.max(6, widths.length); c++) ws.getCell(r, c).fill = fill("FF111111"); }
    for (let c = 1; c <= Math.max(6, widths.length); c++) ws.getCell(5, c).fill = fill("FFD4A12A"); ws.getRow(5).height = 4;
    if (logo !== null) ws.addImage(logo, { tl: { col: 0.08, row: 1.1 }, ext: { width: 58, height: 58 } });
    ws.getCell(2, 2).value = "TKR OPERATIONS – " + title; ws.getCell(2, 2).font = F({ size: 13, bold: true, color: { argb: "FFFFFFFF" } });
    ws.getCell(3, 2).value = monthName(m); ws.getCell(3, 2).font = F({ size: 18, bold: true, color: { argb: "FFD4A12A" } });
    ws.getCell(4, 2).value = sub; ws.getCell(4, 2).font = F({ bold: true, color: { argb: "FFFFFFFF" } }); return ws; };
  const head = (ws, r, L) => { L.forEach((t, i) => { const c = ws.getCell(r, i + 1); c.value = t; c.font = F({ bold: true, color: { argb: "FFFFFFFF" } }); c.fill = fill("FF262626"); c.border = box; c.alignment = { wrapText: true, vertical: "middle", horizontal: i < 2 ? "left" : "center" }; }); ws.getRow(r).height = 28; };
  const row = (ws, r, L, o = {}) => { L.forEach((v, i) => { const c = ws.getCell(r, i + 1); c.value = v; c.border = box; c.font = F({ bold: !!o.bold }); c.alignment = { vertical: "middle", wrapText: true, horizontal: typeof v === "number" ? "center" : "left" }; if (typeof v === "number") c.numFmt = o.int ? "0" : "0.00"; if (o.fill) c.fill = fill(o.fill); }); };
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  // Monthly availability
  let ws = sheetH("Monthly Availability", [14, 11, ...days.map(() => 5.2), 10, 10, 10, 10, 10, 10, 10], "MONTHLY AVAILABILITY", "Available hours per machine per day (06:00 to 06:00), from the exact times logged");
  head(ws, 7, ["Group", "Machine", ...days.map(x => +x.slice(8)), "Available hrs", "Breakdown hrs", "Planned maint. hrs", "Standby hrs", "Not in use hrs", "Down hrs", "Availability %"]); let r = 8;
  for (const [g, L] of fleet()) { for (const mc of L) { const S = days.map(x => daySum(mc, x)), t = S.reduce((a, s) => { if (s) for (const k of ["avail", "BD", "PM", "SB", "NU", "down"]) a[k] += s[k]; return a; }, { avail: 0, BD: 0, PM: 0, SB: 0, NU: 0, down: 0 });
      row(ws, r, [g, mc, ...S.map(s => s ? Math.round(s.avail * 100) / 100 : null), t.avail, t.BD, t.PM, t.SB, t.NU, t.down, t.avail + t.down ? t.avail / (t.avail + t.down) : null]); ws.getCell(r, days.length + 9).numFmt = "0%"; r++; } }
  // Breakdown log
  ws = sheetH("Breakdown log", [12, 12, 20, 16, 16, 11, 40, 16, 16], "BREAKDOWN AND STANDBY LOG", "Every breakdown, standby, planned maintenance and not-in-use with the exact start and end time");
  head(ws, 7, ["Group", "Machine", "Type", "Started", "Ended", "Hours", "Reason", "Logged by", "Closed by"]); r = 8;
  const mf = toMs(m + "-01T00:00"), mt = toMs(addDays(monthEnd(m), 1) + "T06:00");
  for (const e of evs().filter(e => toMs(e.start) < mt && evEnd(e) > mf).sort((a, b) => toMs(a.start) - toMs(b.start))) {
    row(ws, r++, [e.group || groupOf(e.machine), e.machine, TYPES[e.type], `${fmt(e.start.slice(0, 10))} ${hhmm(e.start)}`, e.end ? `${fmt(e.end.slice(0, 10))} ${hhmm(e.end)}` : "still going", Math.round((evEnd(e) - toMs(e.start)) / 36e5 * 100) / 100, e.reason || "", e.by || "", e.endBy || ""]); }
  // Daily checklist
  ws = sheetH("Daily checklist", [12, 16, 34, 10, 10, 10, 14, 30], "DATA AND DOCUMENTS", "Daily control room checklist");
  head(ws, 7, ["Date", "Section", "Item", ...COLS, "Updated", "Comment"]); r = 8;
  for (const d of days) for (const [sec, items] of CHECK) for (const it of items) { const c = recs.find(x => x.id === chkKey(d, it) && !x.deleted); if (!c) continue;
    row(ws, r++, [fmt(d), sec, it, ...COLS.map(k => c[k] || ""), `${new Date(c.editedAt).toTimeString().slice(0, 5)} ${c.by || ""}`, c.comment || ""]); }
  // Invoicing
  ws = sheetH("Supporting documents", [20, 30, 9, 22, 14], "INVOICING PROCESS", "Supporting documents per machine and the person we get them from");
  head(ws, 7, ["Service / machine", "Supporting documents", "Done", "Person responsible", "Ready to invoice"]); r = 8;
  for (const [svc, docs, person] of SUPPORT) { const ready = docs.every(dc => recs.find(x => x.id === invKey(m, svc, dc))?.done === "✓");
    docs.forEach((dc, i) => row(ws, r++, [i ? "" : svc, dc, recs.find(x => x.id === invKey(m, svc, dc))?.done || "", i ? "" : person, i ? "" : ready ? "Ready" : ""], ready ? { fill: "FFE2F0D9" } : {})); }
  ws = sheetH("Additional documents", [40, 8, 26, 22, 40], "ADDITIONAL DOCUMENTS FOR INVOICING", "Needed every month");
  head(ws, 7, ["Additional documents", "Done", "Person responsible", "Control room collector", "Notes"]); r = 8;
  for (const [it, person, note] of ADDITIONAL) { const c = recs.find(x => x.id === addKey(m, it)) || {}; row(ws, r++, [it, c.done || "", c.person || person, c.collector || "", c.notes || note || ""]); }
  const blob = new Blob([await wb.xlsx.writeBuffer()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const { deliver } = await import("../lib/deliver.js"); deliver(blob, `TKR_Control_Room_${monthName(m).replace(" ", "_")}.xlsx`);
}

/* ---------- start ---------- */
(async () => {
  try { recs = await all(); } catch (e) { app.innerHTML = '<div class="notice warn"><b class="t">This browser can\'t store the control room data</b>Open the app in a normal (not private) window.</div>'; return; }
  sync.last = await metaGet("syncLast") || 0;
  const s = await metaGet("me"); if (s && s.until > Date.now()) me = R("user").find(u => u.id === s.id) || null;
  try { navigator.storage?.persist?.(); } catch {}
  render(); if (navigator.onLine) fire().catch(() => {});
})();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("../sw.js", { updateViaCache: "none" }).then(r => r.update()).catch(() => {});
