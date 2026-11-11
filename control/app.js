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
const TABS = [["now", "Live board and workbook", "deploy"], ["inv", "Invoicing", "money"], ["export", "Export & sync", "import"], ["settings", "Settings", "settings"]];
function render() {
  $("who").innerHTML = me ? `${esc(me.name)} · <button class="sm ghost" id="swap" style="min-height:32px;padding:2px 10px">Change controller</button>` : "";
  if (!me) { app.innerHTML = gateView(); return; }
  let body = "";
  if (tab === "now") body = nowView(); if (tab === "avail") body = availView(); if (tab === "docs") body = docsView();
  if (tab === "inv") body = invView(); if (tab === "export") body = exportView(); if (tab === "settings") body = settingsView();
  const ws = app.querySelector(".xlsheet")?.scrollTop, wl = app.querySelector(".xlsheet")?.scrollLeft;
  app.innerHTML = tabBar(TABS, tab) + body; afterRender(app);
  const w = app.querySelector(".xlsheet"); if (w && ws != null) { w.scrollTop = ws; w.scrollLeft = wl; }
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
      ${sheet("ulog", "Update log", `<div class="res">${logs.map(l => `<div class="ri"><div class="rt"><b>${new Date(l.at).toTimeString().slice(0, 5)} · ${esc(l.by)}</b><small>${esc(l.what)}</small></div></div>`).join("") || '<p class="muted">No updates yet today.</p>'}</div>`, { icon: icon("hours"), count: logs.length, sub: "every change, with the time and who did it" })}</div>` + workbookSection();
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
    + `<section class="card dl-card"><h2>Download the month (Excel)</h2><div class="mhead">${mb}</div><p class="muted">The same sheets as the control room workbook: the invoicing lists, Monthly Availability and one sheet per day with the documents, the day and night shift times, the 24-hour summary and the hourly update log.</p><button class="primary" id="xl" style="min-height:52px">Download ${monthName(m)} Excel</button></section>`
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


/* ---------- the control room workbook, as in Excel: Monthly Availability, then one sheet per day ---------- */
let sheetSel = "", sheetDirty = false;
const shiftWin = (d, sh) => sh === "Day" ? [toMs(d + "T06:00"), toMs(d + "T18:00")] : [toMs(d + "T18:00"), toMs(addDays(d, 1) + "T06:00")];
const isoAt = (d, sh, t) => sh === "Day" || t >= "18:00" ? `${d}T${t}` : `${addDays(d, 1)}T${t}`;
const tOf = ms => { const x = new Date(ms); return `${pad(x.getHours())}:${pad(x.getMinutes())}`; };
function evsIn(m, d, sh) { const [a, b] = shiftWin(d, sh); return evs().filter(e => e.machine === m && toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)); }
const n2 = v => v == null ? "" : (Math.round(v * 100) / 100).toFixed(2);
function sheetTabs() {
  const m = (sheetSel && sheetSel !== "month" ? sheetSel : date).slice(0, 7), days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  return `<div class="xltabs"><button data-wbm="-1" aria-label="Month before">‹</button><b>${monthName(m)}</b><button data-wbm="1" aria-label="Next month" ${m >= today().slice(0, 7) ? "disabled" : ""}>›</button>
    <button class="xt ${sheetSel === "month" ? "on" : ""}" data-sh="month">Monthly Availability</button>${days.map(x => `<button class="xt ${sheetSel === x ? "on" : ""}" data-sh="${x}" ${x > today() ? "disabled" : ""}>${x.slice(8)} ${MONTHS[+x.slice(5, 7) - 1].slice(0, 3)}</button>`).join("")}</div>`;
}
function workbookSection() {
  if (!sheetSel) sheetSel = date;
  return `<h2 class="sect" style="margin-top:26px">Control room workbook</h2><p class="muted" style="margin-top:-4px">The same sheets as the Excel. Pick OP, BD or PM in each hour box, like on the Excel. When you pick BD or PM the exact time is logged by itself, and again when you pick OP. At the end of the shift tap <b>Save and upload</b>.</p>`
    + `<div class="xlwin"><div class="xlsheet">${sheetSel === "month" ? monthSheet(date.slice(0, 7)) : daySheet(sheetSel)}</div>${sheetTabs()}</div>`;
}
function monthSheet(m) {
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
  const T = () => ({ avail: 0, BD: 0, PM: 0, SB: 0, down: 0, per: days.map(() => 0) });
  const add = (t, S) => { S.forEach((s, i) => { if (s) { t.per[i] += s.avail; for (const k of ["avail", "BD", "PM", "SB", "down"]) t[k] += s[k]; } }); };
  const line = (lab, t, cls = "") => `<tr class="${cls}"><td colspan="2">${lab}</td>${t.per.map(v => `<td>${v ? Math.round(v * 10) / 10 : ""}</td>`).join("")}<td>${n2(t.avail)}</td><td>${n2(t.BD)}</td><td>${n2(t.PM)}</td><td>${n2(t.down)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : ""}</b></td></tr>`;
  const fleetT = T(); let body = "";
  for (const [g, L] of fleet()) { const gt = T();
    L.forEach((mc, i) => { const S = days.map(x => daySum(mc, x)), t = T(); add(t, S); add(gt, S);
      body += `<tr>${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><b>${esc(mc)}</b></td>${S.map(s => `<td class="${s && s.BD ? "hot" : ""}">${s ? Math.round(s.avail * 10) / 10 : ""}</td>`).join("")}<td>${n2(t.avail)}</td><td>${n2(t.BD)}</td><td>${n2(t.PM)}</td><td>${n2(t.down)}</td><td><b>${t.avail + t.down ? Math.round(t.avail / (t.avail + t.down) * 100) + "%" : ""}</b></td></tr>`; });
    body += line(`${esc(g)} total`, gt, "gt"); for (const k of ["avail", "BD", "PM", "SB", "down"]) fleetT[k] += gt[k]; gt.per.forEach((v, i) => fleetT.per[i] += v); }
  body += line("Fleet total", fleetT, "ft");
  return `<div class="xlhead"><b>Monthly Availability – ${monthName(m)}</b><span>TKR OPERATIONS · THARISA MINE</span><small>Available hours per machine per day (06:00 to 06:00), with month totals. Filled in by itself from the day sheets.</small></div>
    <div class="xlwrap"><table class="xl"><thead><tr><th colspan="2"></th><th colspan="${days.length}">Available hours per day</th><th colspan="5">Month totals</th></tr>
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
function cellVal(d, mc, h, C = cellsOf(d, mc)) {
  if (C[h]) return C[h];
  const a = hStart(d, h), b = a + 36e5, e = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a);
  return e.some(x => x.type === "BD") ? "BD" : e.length ? "PM" : "";
}
function shiftSum(mc, d, sh) {
  const [a, b0] = shiftWin(d, sh), b = Math.min(b0, Date.now()), C = cellsOf(d, mc), V = HRS(sh).map(h => cellVal(d, mc, h, C)), filled = V.filter(Boolean).length;
  const h = b > a ? hoursIn(mc, a, b) : { BD: 0, PM: 0 }, BD = h.BD || 0, PM = h.PM || 0, down = BD + PM;
  return { filled, BD, PM, down, avail: Math.max(0, filled - down), SB: 0, NU: 0 };
}
function daySum(mc, d) {
  const x = shiftSum(mc, d, "Day"), y = shiftSum(mc, d, "Night"), filled = x.filled + y.filled, BD = x.BD + y.BD, PM = x.PM + y.PM, down = BD + PM;
  if (!filled && !down) return null; const avail = x.avail + y.avail;
  return { total: avail + down, avail, BD, PM, SB: 0, NU: 0, down, pct: avail + down > 0 ? avail / (avail + down) : null };
}
function shiftTable(d, sh) {
  const [a, b] = shiftWin(d, sh), now = Date.now(), fut = a > now, H = HRS(sh);
  const tot = () => ({ avail: 0, BD: 0, PM: 0, down: 0 }), shT = tot(); let body = "";
  const opt = v => ["", "OP", "BD", "PM"].map(o => `<option ${o === v ? "selected" : ""}>${o}</option>`).join("");
  for (const [g, L] of fleet()) { const gt = tot(), opCount = H.map(() => 0);
    L.forEach((mc, i) => { const C = cellsOf(d, mc), E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)), s = shiftSum(mc, d, sh);
      const reason = [...new Set(E.map(e => e.reason).filter(Boolean))].join("; ");
      const cells = H.map((h, k) => { const v = cellVal(d, mc, h, C); if (v === "OP") opCount[k]++; return `<td class="hc"><select data-cell="${esc(mc)}|${h}" class="ty ${v}" ${hStart(d, h) > now ? "disabled" : ""}>${opt(v)}</select></td>`; }).join("");
      const tm = (e, end) => end ? (e.end ? (toMs(e.end) <= b ? hhmm(e.end) : "") : "") : (toMs(e.start) < a ? tOf(a) : hhmm(e.start));
      body += `<tr data-mach="${esc(mc)}">${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><b>${esc(mc)}</b></td>${cells}
        <td class="tm">${E.map(e => `<span data-edit="${e.id}" title="Tap to correct">${tm(e)}</span>`).join("<br>")}</td><td class="tm">${E.map(e => `<span data-edit="${e.id}" title="Tap to correct">${tm(e, 1) || "…"}</span>`).join("<br>")}</td>
        <td><input data-reason="${esc(mc)}" value="${esc(reason)}" class="rsn" ${E.length ? "" : "disabled"}></td><td>${s.filled ? n2(s.avail) : ""}</td><td class="${s.BD ? "hot" : ""}">${s.filled || s.BD ? n2(s.BD) : ""}</td><td>${s.filled || s.PM ? n2(s.PM) : ""}</td><td>${s.filled || s.down ? n2(s.down) : ""}</td></tr>`;
      for (const k of ["avail", "BD", "PM", "down"]) gt[k] += s[k]; });
    body += `<tr class="gt"><td colspan="2">${esc(g)} total</td>${opCount.map(n => `<td>${n}/${L.length}</td>`).join("")}<td></td><td></td><td></td><td>${n2(gt.avail)}</td><td>${n2(gt.BD)}</td><td>${n2(gt.PM)}</td><td>${n2(gt.down)}</td></tr>`;
    for (const k of ["avail", "BD", "PM", "down"]) shT[k] += gt[k]; }
  body += `<tr class="ft"><td colspan="${H.length + 5}">${sh} shift total</td><td>${n2(shT.avail)}</td><td>${n2(shT.BD)}</td><td>${n2(shT.PM)}</td><td>${n2(shT.down)}</td></tr>`;
  const sv = R("log").filter(l => l.date === d && l.shiftSave === sh).sort((x, y) => y.at - x.at)[0];
  return `<div class="xlband">Machine availability – ${sh} shift (${sh === "Day" ? "06:00 to 18:00" : "18:00 to 06:00"}) &nbsp; OP = Operational / standby · BD = Breakdown · PM = Planned maintenance · the exact time is logged when you pick BD or PM, and when you pick OP again</div>
    <div class="xlwrap"><table class="xl av" data-shift="${sh}"><thead><tr><th>Group</th><th>Machine</th>${H.map(h => `<th>${hLabel(h)}</th>`).join("")}<th>Down at</th><th>Back up</th><th>Reason / comments</th><th>Avail hrs</th><th>BD hrs</th><th>PM hrs</th><th>Down hrs</th></tr></thead><tbody>${body}</tbody></table></div>
    <div class="savebar"><button type="button" class="primary" data-saveshift="${sh}" ${fut ? "disabled" : ""}>Save and upload ${sh.toLowerCase()} shift</button><button type="button" data-msg="${sh}" ${fut ? "disabled" : ""}>Generate message</button><span class="muted">${sv ? `Saved ${new Date(sv.at).toTimeString().slice(0, 5)}` : fut ? "This shift has not started yet." : ""}</span></div>`;
}
/* one hour box changed: keep it, and log the exact time for a breakdown or planned maintenance */
async function setCell(d, mc, h, val) {
  const id = `hr_${d}_${mc}`, rec = recs.find(r => r.id === id) || { id, kind: "hr", ym: d.slice(0, 7), date: d, machine: mc, cells: {} };
  const a = hStart(d, h), b = a + 36e5, now = Date.now(), current = now >= a && now < b, cells = { ...rec.cells }; if (val) cells[h] = val; else delete cells[h];
  await put({ ...rec, deleted: false, cells }, `${mc} ${hLabel(h)}: ${val || "cleared"}`);
  const open = downEvs(mc).filter(e => !e.end && toMs(e.start) <= now).sort((x, y) => toMs(y.start) - toMs(x.start))[0];
  if (val === "BD" || val === "PM") {
    if (current) { if (open && open.type === val) return; if (open) await put({ ...open, end: nowLocal(), endBy: me.name });
      await put({ id: "ev_" + uid(), kind: "ev", ym: d.slice(0, 7), machine: mc, group: groupOf(mc), type: val, start: nowLocal(), end: "", reason: "" }, `${mc}: ${TYPES[val]} from ${hhmm(nowLocal())}`);
      toast(`${mc}: ${TYPES[val].toLowerCase()} logged at ${hhmm(nowLocal())}.`); return; }
    const E = downEvs(mc).filter(e => e.type === val);
    if (E.some(e => toMs(e.start) < b && evEnd(e) > a)) return;                                   // that hour is already inside a logged time
    const before = E.find(e => e.end && toMs(e.end) === a), after = E.find(e => toMs(e.start) === b);
    if (before) await put({ ...before, end: isoOf(b) }); else if (after) await put({ ...after, start: isoOf(a) });
    else await put({ id: "ev_" + uid(), kind: "ev", ym: d.slice(0, 7), machine: mc, group: groupOf(mc), type: val, start: isoOf(a), end: isoOf(b), reason: "", grid: 1 });
    toast(`${mc}: tap Down at or Back up to type the exact time.`); return;
  }
  if (val === "OP") {
    if (current && open) { await put({ ...open, end: nowLocal(), endBy: me.name }, `${mc} back up ${hhmm(nowLocal())}`); toast(`${mc}: back up logged at ${hhmm(nowLocal())}.`); return; }
    if (open && toMs(open.start) < a) { await put({ ...open, end: isoOf(a), endBy: me.name }); toast(`${mc}: back up at ${tOf(a)}. Tap Back up to type the exact time.`); }
  }
}
async function setReason(mc, d, sh, reason) {
  const [a, b] = shiftWin(d, sh), E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a);
  for (const e of E) if ((e.reason || "") !== reason) await put({ ...e, reason }, `${mc}: ${reason}`);
}
function shiftMessage(d, sh) {
  const [a, b] = shiftWin(d, sh), day = recs.find(r => r.id === "day_" + d) || {}, ctl = (sh === "Day" ? day.dayCtl : day.nightCtl) || me.name;
  const L = [`*TKR Control Room – ${sh} shift ${dow(d)} ${fmt(d)}* (${sh === "Day" ? "06:00–18:00" : "18:00–06:00"})`, `Controller: ${ctl}`, ""];
  for (const [g, M] of fleet()) {
    const rows = M.map(mc => [mc, evsIn(mc, d, sh).filter(e => e.type === "BD" || e.type === "PM"), shiftSum(mc, d, sh)]), S = rows.reduce((t, [, , s]) => { if (s) { t.avail += s.avail; t.down += s.down; } return t; }, { avail: 0, down: 0 });
    L.push(`*${g}* – availability ${S.avail + S.down ? Math.round(S.avail / (S.avail + S.down) * 100) + "%" : "–"}`);
    const down = rows.filter(([, E]) => E.length);
    if (!down.length) L.push("All working"); else for (const [mc, E] of down) L.push(`${mc}: ` + E.map(e => `${TYPES[e.type]} ${toMs(e.start) < a ? tOf(a) : hhmm(e.start)}–${e.end && toMs(e.end) <= b ? hhmm(e.end) : e.end ? tOf(b) : "still going"} (${hm((Math.min(evEnd(e), b, Date.now()) - Math.max(toMs(e.start), a)) / 36e5)} h)${e.reason ? " – " + e.reason : ""}`).join("; "));
    L.push("");
  }
  const items = CHECK.flatMap(([, i]) => i), got = k => items.filter(it => (recs.find(r => r.id === chkKey(d, it)) || {})[k] === "✓").length;
  L.push(`*Documents received:* ${COLS.map(k => `${k} ${got(k)}/${items.length}`).join(" · ")}`);
  const missing = items.filter(it => COLS.some(k => (recs.find(r => r.id === chkKey(d, it)) || {})[k] === "✗")); if (missing.length) L.push(`Not received: ${missing.join(", ")}`);
  return L.join("\n");
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
    <div class="xlinfo"><label>Date <input value="${fmt(d)}" readonly></label><label>Day controller <input name="dayCtl" value="${esc(day.dayCtl || "")}" placeholder="${esc(me.name)}"></label><label>Night controller <input name="nightCtl" value="${esc(day.nightCtl || "")}" placeholder="${esc(me.name)}"></label>
      <label>Last update <input value="${last ? new Date(last.at).toTimeString().slice(0, 5) : ""}" readonly></label><label>Next update <input value="${last ? new Date(last.at + 36e5).toTimeString().slice(0, 5) : ""}" readonly></label>
      <label>Status <input value="${d !== today() ? "" : !last ? "Not started" : Date.now() - last.at > 36e5 ? "Update due" : "Up to date"}" readonly></label></div>
    <div class="xlband">Data and documents</div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Section</th><th>Item</th>${COLS.map(c => `<th>${c}</th>`).join("")}<th>Time updated</th><th>Updated by</th><th>Comments</th></tr></thead><tbody>
      ${CHECK.map(([sec, items]) => items.map((it, i) => { const c = get(it); return `<tr data-item="${esc(it)}">${i ? "" : `<td rowspan="${items.length}" class="grpc">${esc(sec)}</td>`}<td class="lt">${esc(it)}</td>${COLS.map(k => `<td><select data-c="${k}" class="tk">${opt(c[k])}</select></td>`).join("")}<td>${c.editedAt ? new Date(c.editedAt).toTimeString().slice(0, 5) : ""}</td><td>${esc(c.by || "")}</td><td><input data-c="comment" value="${esc(c.comment || "")}"></td></tr>`; }).join("")).join("")}
      <tr class="gt"><td colspan="2">Captured</td>${COLS.map(k => `<td>${CHECK.flatMap(([, i]) => i).filter(it => get(it)[k] === "✓").length} of ${CHECK.flatMap(([, i]) => i).length}</td>`).join("")}<td colspan="3"></td></tr></tbody></table></div>
    ${shiftTable(d, "Day")}${shiftTable(d, "Night")}
    <div class="xlband">24-hour availability summary per machine</div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Group</th><th>Machine</th><th>Available hrs</th><th>Breakdown hrs</th><th>Planned maint. hrs</th><th>Down hrs</th><th>Total hrs</th><th>Availability %</th></tr></thead><tbody>
      ${fleet().map(([g, L]) => L.map((mc, i) => { const s = daySum(mc, d); return `<tr>${i ? "" : `<td rowspan="${L.length}" class="grpc">${esc(g)}</td>`}<td class="lt"><b>${esc(mc)}</b></td>${s ? `<td>${n2(s.avail)}</td><td class="${s.BD ? "hot" : ""}">${n2(s.BD)}</td><td>${n2(s.PM)}</td><td>${n2(s.down)}</td><td>${n2(s.avail + s.down)}</td><td><b>${s.pct == null ? "" : Math.round(s.pct * 100) + "%"}</b></td>` : '<td colspan="6"></td>'}</tr>`; }).join("")).join("")}</tbody></table></div>
    <div class="xlband">Hourly update log <small>(fills in by itself every time something is saved)</small></div>
    <div class="xlwrap"><table class="xl"><thead><tr><th>Hour</th><th>Updated</th><th>Time done</th><th>Controller</th><th>Notes</th></tr></thead><tbody>
      ${hourLog.map(([h, L]) => `<tr><td>${pad(h)}:00–${pad((h + 1) % 24)}:00</td><td class="${L.length ? "okc" : ""}">${L.length ? "✓" : ""}</td><td>${L[0] ? new Date(L[0].at).toTimeString().slice(0, 5) : ""}</td><td>${esc([...new Set(L.map(l => l.by))].join(", "))}</td><td class="nt">${esc(L.map(l => l.what).slice(0, 3).join(" · "))}${L.length > 3 ? ` · +${L.length - 3} more` : ""}</td></tr>`).join("")}
      <tr class="gt"><td>Hours updated</td><td colspan="4">${hourLog.filter(([, L]) => L.length).length} of 24</td></tr></tbody></table></div></form>`;
}
async function saveShift(sh) {
  const f = $("fDay"); if (!f) return; const d = f.dataset.date;
  const dayR = recs.find(r => r.id === "day_" + d) || { id: "day_" + d, kind: "day", ym: d.slice(0, 7), date: d };
  const ctl = { dayCtl: f.elements.dayCtl.value.trim() || (sh === "Day" ? me.name : dayR.dayCtl || ""), nightCtl: f.elements.nightCtl.value.trim() || (sh === "Night" ? me.name : dayR.nightCtl || "") };
  if (ctl.dayCtl !== (dayR.dayCtl || "") || ctl.nightCtl !== (dayR.nightCtl || "")) await put({ ...dayR, ...ctl }, `Controllers: day ${ctl.dayCtl || "–"}, night ${ctl.nightCtl || "–"}`);
  const n = R("log").filter(l => l.date === d).length;
  await put({ id: "log_" + uid(), kind: "log", ym: d.slice(0, 7), date: d, at: Date.now(), shiftSave: sh, what: `Saved and uploaded the ${sh.toLowerCase()} shift sheet` });
  let up = false; if (navigator.onLine && fbUser) { await syncNow(false); up = !waiting(); }
  render();
  popupMsg(up ? "Saved and uploaded" : "Saved on this PC", `${sh} shift, ${dow(d)} ${fmt(d)}: ${n} updates on this sheet.` + (up ? " Everyone in the TKR app can see it now." : fbUser ? " It uploads by itself as soon as this PC has signal." : " It uploads once someone signs in under Export & sync (needs signal once)."));
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
  const cell = ev.target.closest("[data-edit]:not(button)"); if (cell) { const e = recs.find(r => r.id === cell.dataset.edit); if (e) openEv(evForm(e), `${e.machine} · correct the times`); return; }
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
  if (b.dataset.sh) { if (sheetDirty && !confirm("This sheet has changes that are not saved. Leave without saving?")) return; sheetDirty = false; sheetSel = b.dataset.sh; if (sheetSel !== "month") date = sheetSel; render(); document.querySelector(".xltabs")?.scrollIntoView({ block: "start" }); return; }
  if (b.dataset.wbm) { if (sheetDirty && !confirm("This sheet has changes that are not saved. Leave without saving?")) return; sheetDirty = false; const m = shiftMonth((sheetSel === "month" ? date : sheetSel).slice(0, 7), +b.dataset.wbm), last = monthEnd(m) < today() ? monthEnd(m) : today(); date = last; sheetSel = sheetSel === "month" ? "month" : last; render(); return; }
  if (b.dataset.msg) { showMessage(shiftMessage($("fDay").dataset.date, b.dataset.msg)); return; }
  if (b.dataset.saveshift) { b.disabled = true; b.textContent = "Saving…"; try { await saveShift(b.dataset.saveshift); } catch (x) { console.warn(x); toast("It didn't save. Try again."); } finally { b.disabled = false; } return; }
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
  if (t.dataset.cell !== undefined) { const [mc, h] = t.dataset.cell.split("|"); await setCell($("fDay").dataset.date, mc, +h, t.value); render(); return; }
  if (t.dataset.reason !== undefined) { const f = $("fDay"), tr = t.closest("table.av"); await setReason(t.dataset.reason, f.dataset.date, tr.dataset.shift, t.value.trim()); toast("Saved"); return; }
  if (t.closest && t.closest("#fDay") && t.closest("tr[data-item]")) { const f = $("fDay"), d = f.dataset.date, tr = t.closest("tr[data-item]"), it = tr.dataset.item, id = chkKey(d, it);
    const c = recs.find(r => r.id === id) || { id, kind: "chk", ym: d.slice(0, 7), date: d, item: it }; await put({ ...c, [t.dataset.c]: t.value.trim() }, `${it} (${t.dataset.c}): ${t.value.trim() || "cleared"}`);
    if (t.tagName === "SELECT") { const keep = scrollY, ws = app.querySelector(".xlsheet")?.scrollTop; render(); scrollTo(0, keep); const w = app.querySelector(".xlsheet"); if (w && ws != null) w.scrollTop = ws; } else toast("Saved"); return; }
  if (t.closest && t.closest("#fDay")) return;
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
  sync.busy = false; if (me && !sheetDirty && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render();
}
function syncCard() {
  if (!fbUser) return `<section class="card"><h2>Send to the TKR app</h2><p class="muted" style="margin-top:0">Sign in once on this PC with the control room app account (needs signal once). After that, everything goes to the TKR app by itself whenever there is signal, so managers see it live. Without signal it waits on this PC.</p>
    <form id="fSync" class="grid" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))"><label>Email or username <input name="e" autocapitalize="none" required></label><label>Password <input name="p" type="password" required></label><div style="display:flex;align-items:flex-end"><button class="primary" style="width:100%">Sign in</button></div></form></section>`;
  return `<section class="card"><h2>Send to the TKR app</h2><p style="margin-top:0">Signed in as <b>${esc(fbUser.email)}</b> · ${waiting() ? `<b>${waiting()}</b> changes waiting` : "nothing waiting"}. ${esc(sync.msg)}</p><div class="row"><button class="primary" id="syncNow">Send and check now</button><button class="ghost" id="syncOut">Sign out</button></div></section>`;
}
addEventListener("online", () => syncNow(false));
setInterval(() => { if (me && tab === "now" && !sheetDirty && !document.activeElement?.closest?.("#fDay") && !document.getElementById("evPanel") && !document.querySelector(".sheet:not([hidden])")) render(); }, 60000);   // live times on the board

/* ---------- Excel for one month: the same sheets as the control room workbook ---------- */
async function exportXlsx(m) {
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
    c.alignment = { vertical: "middle", wrapText: !!o.wrap, horizontal: i < 2 || (o.left || []).includes(i) ? "left" : "center" }; if (typeof v === "number") c.numFmt = (o.fmt || {})[i] || "0.00"; if (o.fill) c.fill = fill(o.fill);
    const cc = COLR[v]; if (cc && !o.fill) { c.fill = fill(cc[0]); c.font = F({ bold: true, color: { argb: cc[1] } }); } }); };
  const days = []; for (let x = m + "-01"; x <= monthEnd(m); x = addDays(x, 1)) days.push(x);
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
  ws = wb.addWorksheet("Monthly Availability", { views: [{ state: "frozen", xSplit: 2, ySplit: 7, showGridLines: false }] }); ws.columns = [14, 11, ...days.map(() => 5.4), 11, 11, 11, 11, 11].map(w => ({ width: w }));
  const nc = 2 + days.length + 5; top(ws, nc, `Monthly Availability – ${monthName(m)}`, "TKR OPERATIONS  ·  THARISA MINE", "Available hours per machine per day (06:00 to 06:00), with month totals");
  head(ws, 7, ["Group", "Machine", ...days.map(x => +x.slice(8)), "Available hrs", "Breakdown hrs", "Planned maint. hrs", "Down hrs", "Availability %"]); r = 8;
  const pct = (a, d) => a + d ? a / (a + d) : "";
  const fleetT = { per: days.map(() => 0), avail: 0, BD: 0, PM: 0, SB: 0, down: 0 };
  for (const [g, L] of fleet()) { const gt = { per: days.map(() => 0), avail: 0, BD: 0, PM: 0, SB: 0, down: 0 };
    for (const mc of L) { const S = days.map(x => daySum(mc, x)), t = { avail: 0, BD: 0, PM: 0, SB: 0, down: 0 }; S.forEach((s, i) => { if (s) { gt.per[i] += s.avail; for (const k in t) { t[k] += s[k]; gt[k] += s[k]; } } });
      row(ws, r, [g, mc, ...S.map(s => s ? Math.round(s.avail * 100) / 100 : ""), t.avail, t.BD, t.PM, t.down, pct(t.avail, t.down)], { boldCol: [1] }); ws.getCell(r, nc).numFmt = "0%"; r++; }
    row(ws, r, [`${g} total`, "", ...gt.per.map(v => Math.round(v * 100) / 100), gt.avail, gt.BD, gt.PM, gt.down, pct(gt.avail, gt.down)], { bold: true, fill: "FFFFF2CC" }); ws.getCell(r, nc).numFmt = "0%"; r++;
    gt.per.forEach((v, i) => fleetT.per[i] += v); for (const k of ["avail", "BD", "PM", "SB", "down"]) fleetT[k] += gt[k]; }
  row(ws, r, ["Fleet total", "", ...fleetT.per.map(v => Math.round(v * 100) / 100), fleetT.avail, fleetT.BD, fleetT.PM, fleetT.down, pct(fleetT.avail, fleetT.down)], { bold: true, fill: "FFD9D9D9" }); ws.getCell(r, nc).numFmt = "0%";
  // 4. one sheet per day, like the control room capture sheets
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
      r++; band(ws, r++, 21, `Machine availability – ${sh} shift (${sh === "Day" ? "06:00 to 18:00" : "18:00 to 06:00"})   OP = Operational / standby  ·  BD = Breakdown  ·  PM = Planned maintenance`);
      head(ws, r++, ["Group", "Machine", ...HRS(sh).map(hLabel), "Down at", "Back up", "Reason / comments", "Avail hrs", "BD hrs", "PM hrs", "Down hrs"]);
      const [a, b] = shiftWin(d, sh);
      for (const [g, L] of fleet()) L.forEach((mc, i) => { const E = downEvs(mc).filter(e => toMs(e.start) < b && evEnd(e) > a).sort((x, y) => toMs(x.start) - toMs(y.start)), s = shiftSum(mc, d, sh), C = cellsOf(d, mc);
        row(ws, r++, [i ? "" : g, mc, ...HRS(sh).map(h => cellVal(d, mc, h, C)), E.map(e => toMs(e.start) < a ? tOf(a) : hhmm(e.start)).join(", "), E.map(e => e.end && toMs(e.end) <= b ? hhmm(e.end) : "").filter(Boolean).join(", "),
          [...new Set(E.map(e => e.reason).filter(Boolean))].join("; "), s.filled ? s.avail : "", s.filled || s.BD ? s.BD : "", s.filled || s.PM ? s.PM : "", s.filled || s.down ? s.down : ""], { boldCol: [0, 1], left: [16] }); });
    }
    r++; band(ws, r++, 21, "24-hour availability summary per machine"); head(ws, r++, ["Group", "Machine", "Available hrs", "Breakdown hrs", "Planned maint. hrs", "Down hrs", "Total hrs", "Availability %"]);
    for (const [g, L] of fleet()) L.forEach((mc, i) => { const s = daySum(mc, d); row(ws, r, [i ? "" : g, mc, ...(s ? [s.avail, s.BD, s.PM, s.down, s.avail + s.down, s.pct ?? ""] : ["", "", "", "", "", ""])], { boldCol: [0, 1] }); ws.getCell(r, 8).numFmt = "0%"; r++; });
    r++; band(ws, r++, 21, "Hourly update log"); head(ws, r++, ["Hour", "Updated", "Time done", "Controller", "Notes"]);
    for (let h = 6; h < 30; h++) { const hh = h % 24, st = toMs((h >= 24 ? addDays(d, 1) : d) + `T${pad(hh)}:00`), L = logs.filter(l => l.at >= st && l.at < st + 36e5);
      row(ws, r++, [`${pad(hh)}:00–${pad((hh + 1) % 24)}:00`, L.length ? "✓" : "", L[0] ? new Date(L[0].at).toTimeString().slice(0, 5) : "", [...new Set(L.map(l => l.by))].join(", "), L.map(l => l.what).slice(0, 4).join(" · ")], { left: [4] }); }
  }
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
