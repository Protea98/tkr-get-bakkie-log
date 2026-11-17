/* TKR Training Matrix — every employee's training expiry dates, the trainings expiring each month with booking and progress,
   and Excel downloads in the same layout as the Training Matrix workbook. Shared online through Firebase (collection "training"). */
import {
  initializeApp, deleteApp, getAuth, signInWithEmailAndPassword, createUserWithEmailAndPassword, arrayUnion, arrayRemove, signOut, onAuthStateChanged, sendPasswordResetEmail,
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, setDoc, deleteDoc, onSnapshot,
} from "../lib/firebase.js";

const firebaseConfig = {
  apiKey: "AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI", authDomain: "tkr-operations-app.firebaseapp.com", projectId: "tkr-operations-app",
  storageBucket: "tkr-operations-app.firebasestorage.app", messagingSenderId: "158080247079", appId: "1:158080247079:web:7cb7786d2b7f9ee172b502",
};
// Must match the Firestore rules (firestore-rules.txt, Training Matrix block).
const ADMINS = ["muneer@tkroperations.com", "amogelang@tkroperations.com", "amogelangmabilu56@gmail.com"];
const ACCESS_ADMINS = ["muneer@tkroperations.com", "amogelang@tkroperations.com", "amogelangmabilu56@gmail.com"];   // Muneer and Amogelang give or remove access
const toLogin = v => { v = String(v || "").trim(); return v.includes("@") ? v : v.toLowerCase().replace(/[^a-z0-9]/g, "") + "@mgr.tkroperations.app"; };

const fbApp = initializeApp(firebaseConfig);
const auth = getAuth(fbApp);
let db;
try { db = initializeFirestore(fbApp, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) }); }
catch { db = initializeFirestore(fbApp, {}); }

/* ---------- helpers ---------- */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pad = n => String(n).padStart(2, "0");
const today = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const isIso = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));
const fmtD = s => isIso(s) ? s.slice(8) + "/" + s.slice(5, 7) + "/" + s.slice(0, 4) : String(s || "");
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const ym = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const addM = (m, k) => { const [y, mo] = m.split("-").map(Number); return ym(new Date(y, mo - 1 + k, 1)); };
const mName = m => MONTHS[+m.slice(5) - 1] + " " + m.slice(0, 4);
const slug = s => String(s || "").toLowerCase().normalize("NFKD").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 70);
const normName = s => String(s || "").replace(/\s+/g, " ").trim();
const empId = name => "e_" + slug(name);
const bookId = (name, course, month) => "b_" + slug(name) + "__" + slug(course) + "__" + month;
function toast(m) { const t = $("toast"); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3200); }
function fail(e) { console.warn(e); toast(e?.code === "permission-denied" ? "You are not allowed to do that." : "Not saved: " + (e?.message || "try again")); }

const PROGRESS = ["No booking required", "No bookings yet", "Booked", "Completed", "Resigned"];
const STATUS = { ex: "Expired", m0: "Expires this month", m1: "Expires next month", ok: "Valid", mi: "Mandatory – missing" };

/* ---------- state ---------- */
let email = "", tmAcc = { editors: [], viewers: [] }, view = "dash", loaded = false, unsub = null, accUnsub = null;
let meta = { courses: [], mandatory: [] }, emps = [], books = {};
let fQ = "", fOcc = "", fShow = "", month = ym();
const canEdit = () => ADMINS.includes(email) || tmAcc.editors.includes(email);
const isAccessAdmin = () => ACCESS_ADMINS.includes(email);
const ref = id => doc(db, "training", id);
const active = () => emps.filter(e => !e.removed).sort((a, b) => a.name.localeCompare(b.name));

function status(d, course) {
  if (!d) return meta.mandatory.includes(course) ? "mi" : "";
  if (!isIso(d)) return "";
  const t = today(), m = ym();
  if (d < t) return "ex";
  if (d.slice(0, 7) === m) return "m0";
  if (d.slice(0, 7) === addM(m, 1)) return "m1";
  return "ok";
}
const keyHtml = () => `<div class="key">${["ex", "m0", "m1", "ok"].concat(meta.mandatory.length ? ["mi"] : []).map(k => `<span><i class="c-${k}"></i>${k === "m0" ? "Expires " + mName(ym()) : k === "m1" ? "Expires " + mName(addM(ym(), 1)) : STATUS[k]}</span>`).join("")}</div>`;

/* rows for one month: every training on the matrix that expires in that month, plus the booking lines kept for that month */
function monthRows(m) {
  const out = new Map();
  for (const e of active()) for (const c of meta.courses) {
    const d = (e.dates || {})[c];
    if (isIso(d) && d.slice(0, 7) === m) { const id = bookId(e.name, c, m); out.set(id, Object.assign({ id, month: m, name: e.name, pos: e.occ || "", course: c, expiry: d, cert: "", provider: "", booking: "", progress: "" }, books[id] || {})); }
  }
  for (const b of Object.values(books)) if (b.month === m && !out.has(b.id)) out.set(b.id, b);
  return [...out.values()].sort((a, b) => a.name.localeCompare(b.name) || a.course.localeCompare(b.course));
}
const progCls = p => p === "Completed" ? "done" : p === "No bookings yet" ? "none" : p === "Booked" ? "bk" : "";

/* ---------- modal ---------- */
function modal(title, sub, body, acts) {
  $("modal").innerHTML = `<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true"><h2>${esc(title)}</h2>${sub ? `<p class="note" style="margin:0 0 12px">${sub}</p>` : ""}${body}<div class="acts">${acts}</div></div></div>`;
  $("scrim").onclick = e => { if (e.target.id === "scrim") closeModal(); };
}
const closeModal = () => { $("modal").innerHTML = ""; render(); };

/* ================= Dashboard ================= */
function viewDash() {
  const A = active(), cnt = { ex: 0, m0: 0, m1: 0, ok: 0, mi: 0 }, per = {}, expired = [];
  for (const c of meta.courses) per[c] = { n: 0, ex: 0, m0: 0, m1: 0, ok: 0, mi: 0 };
  for (const e of A) for (const c of meta.courses) {
    const d = (e.dates || {})[c], s = status(d, c); if (!s) continue;
    cnt[s]++; per[c][s]++; if (d) per[c].n++;
    if (s === "ex") expired.push([e.name, e.occ, c, d]);
  }
  const kpi = (n, l, s, cls) => `<div class="kpi"${cls ? ` style="border-top-color:var(--${cls})"` : ""}><span>${l}</span><b>${n}</b>${s ? `<small>${s}</small>` : ""}</div>`;
  let h = '<div class="kpis">' + kpi(A.length, "Employees") + kpi(cnt.ex, "Expired", "trainings past their expiry date", "bad") +
    kpi(cnt.m0, "Still to expire this month", "from today to the end of " + mName(ym())) + kpi(cnt.m1, "Expire next month", mName(addM(ym(), 1))) + kpi(cnt.ok, "Valid", "", "ok") +
    (meta.mandatory.length ? kpi(cnt.mi, "Mandatory missing", "", "bad") : "") + '</div>';
  const months = Array.from({ length: 6 }, (_, i) => addM(ym(), i));
  h += `<div class="card"><h2>Expiring in the next six months</h2><div class="tbl"><table><thead><tr><th>Month</th><th class="num">Trainings expiring</th><th class="num">Completed</th><th class="num">Booked</th><th class="num">No bookings yet</th><th class="num">No progress captured</th><th></th></tr></thead><tbody>` +
    months.map(m => { const r = monthRows(m), c = p => r.filter(x => x.progress === p).length;
      return `<tr><td>${mName(m)}</td><td class="num">${r.length}</td><td class="num">${c("Completed")}</td><td class="num">${c("Booked")}</td><td class="num">${c("No bookings yet")}</td><td class="num">${c("")}</td><td><button class="sm" data-gom="${m}">Open</button></td></tr>`; }).join("") + '</tbody></table></div></div>';
  const used = meta.courses.filter(c => per[c].n || per[c].mi);
  h += `<div class="card"><h2>Per training</h2>${keyHtml()}<div class="tbl"><table><thead><tr><th>Training</th><th class="num">Employees with it</th><th class="num">Valid</th><th class="num">Expires ${MONTHS[+ym().slice(5) - 1]}</th><th class="num">Expires ${MONTHS[+addM(ym(), 1).slice(5) - 1]}</th><th class="num">Expired</th>${meta.mandatory.length ? '<th class="num">Mandatory missing</th>' : ""}</tr></thead><tbody>` +
    (used.length ? used.map(c => `<tr><td>${esc(c)}</td><td class="num">${per[c].n}</td><td class="num">${per[c].ok}</td><td class="num">${per[c].m0}</td><td class="num">${per[c].m1}</td><td class="num">${per[c].ex ? `<b class="flag">${per[c].ex}</b>` : 0}</td>${meta.mandatory.length ? `<td class="num">${per[c].mi}</td>` : ""}</tr>`).join("") : `<tr><td colspan="7" class="note">No training dates yet.${canEdit() ? " Import the Training Matrix Excel on the Settings tab." : ""}</td></tr>`) + '</tbody></table></div></div>';
  expired.sort((a, b) => a[3].localeCompare(b[3]));
  h += `<div class="card"><h2>Expired trainings <span class="note">${expired.length}</span></h2>${expired.length ? `<div class="tbl"><table><thead><tr><th>Employee</th><th>Occupation</th><th>Training</th><th>Expired on</th></tr></thead><tbody>${expired.map(r => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td><td>${esc(r[2])}</td><td class="flag">${fmtD(r[3])}</td></tr>`).join("")}</tbody></table></div>` : '<p class="note">None.</p>'}</div>`;
  $("main").innerHTML = h;
  document.querySelectorAll("[data-gom]").forEach(b => b.onclick = () => { month = b.dataset.gom; view = "month"; render(); });
}

/* ================= Training Matrix ================= */
function viewMatrix() {
  const occs = [...new Set(active().map(e => e.occ).filter(Boolean))].sort();
  let list = active().filter(e => (!fQ || e.name.toLowerCase().includes(fQ.toLowerCase())) && (!fOcc || e.occ === fOcc));
  if (fShow) list = list.filter(e => meta.courses.some(c => { const s = status((e.dates || {})[c], c); return fShow === "ex" ? s === "ex" || s === "mi" : ["ex", "m0", "m1", "mi"].includes(s); }));
  const ed = canEdit();
  let h = `<div class="bar"><input type="search" id="fQ" placeholder="Search an employee" value="${esc(fQ)}">
    <select id="fOcc"><option value="">All occupations</option>${occs.map(o => `<option${o === fOcc ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>
    <select id="fShow"><option value="">Everyone</option><option value="due"${fShow === "due" ? " selected" : ""}>Expired or expiring soon</option><option value="ex"${fShow === "ex" ? " selected" : ""}>Expired only</option></select>
    ${canEdit() ? '<button id="imX">Import Excel</button>' : ""}<button id="xAll" class="primary">Download Excel</button></div>${keyHtml()}
    <p class="note" style="margin:-4px 0 8px">${list.length} of ${active().length} employees.${ed ? " Tap a date to change it, or a name to change the employee." : ""}</p>`;
  h += `<div class="mx"><table><thead><tr><th class="nm">Employee</th><th>Occupation</th>${meta.courses.map(c => `<th class="vt">${esc(c)}</th>`).join("")}</tr></thead><tbody>` +
    list.map(e => `<tr><td class="nm"${ed ? ` data-emp="${esc(e.id)}" style="cursor:pointer"` : ""}>${esc(e.name)}</td><td class="oc">${esc(e.occ)}</td>` +
      meta.courses.map((c, i) => { const d = (e.dates || {})[c], s = status(d, c); return `<td class="d${s ? " c-" + s : ""}${ed ? " ed" : ""}"${ed ? ` data-cell="${esc(e.id)}|${i}"` : ""} title="${esc(c)}${s ? " – " + STATUS[s] : ""}">${d ? esc(fmtD(d)) : ""}</td>`; }).join("") + "</tr>").join("") +
    `</tbody></table></div>`;
  if (!active().length) h = `<div class="card note">No employees yet.${ed ? ' <button class="primary" id="imX">Import Excel</button>' : ""}</div>`;
  $("main").innerHTML = h;
  if (!active().length) { if ($("imX")) $("imX").onclick = importPick; return; }
  $("fQ").oninput = e => { fQ = e.target.value; const p = e.target.selectionStart; viewMatrix(); const q = $("fQ"); q.focus(); q.setSelectionRange(p, p); };
  $("fOcc").onchange = e => { fOcc = e.target.value; viewMatrix(); };
  $("fShow").onchange = e => { fShow = e.target.value; viewMatrix(); };
  $("xAll").onclick = () => downloadAll();
  if ($("imX")) $("imX").onclick = importPick;
  document.querySelectorAll("[data-cell]").forEach(td => td.onclick = () => { const [id, i] = td.dataset.cell.split("|"); cellForm(emps.find(e => e.id === id), meta.courses[+i]); });
  document.querySelectorAll("[data-emp]").forEach(td => td.onclick = () => empForm(emps.find(e => e.id === td.dataset.emp)));
}
function cellForm(e, c) {
  const d = (e.dates || {})[c] || "";
  modal(c, esc(e.name) + (e.occ ? " · " + esc(e.occ) : ""), `<div class="form"><label>Expiry date<input type="date" id="cD" value="${isIso(d) ? d : ""}"></label></div><p class="note">Leave it empty when the employee does not have this training.</p>`,
    `<button id="mX">Cancel</button><button class="primary" id="mS">Save</button>`);
  $("mX").onclick = closeModal;
  $("mS").onclick = async () => {
    const v = $("cD").value, dates = Object.assign({}, e.dates || {}); if (v) dates[c] = v; else delete dates[c];
    try { await setDoc(ref(e.id), { dates, updatedAt: Date.now(), updatedBy: email }, { merge: true }); e.dates = dates; toast("Saved"); closeModal(); } catch (er) { fail(er); }
  };
}
function empForm(e) {
  const isNew = !e; e = e || { name: "", occ: "", dates: {} };
  modal(isNew ? "Add an employee" : e.name, "", `<div class="form"><label>Employee (Surname, Name)<input id="eN" value="${esc(e.name)}"${isNew ? "" : " disabled"}></label><label>Occupation<input id="eO" value="${esc(e.occ)}" list="occList"></label></div>
    <datalist id="occList">${[...new Set(emps.map(x => x.occ).filter(Boolean))].sort().map(o => `<option value="${esc(o)}">`).join("")}</datalist>`,
    `${isNew ? "" : '<button class="danger" id="mR" style="margin-right:auto">Remove from the matrix</button>'}<button id="mX">Cancel</button><button class="primary" id="mS">Save</button>`);
  $("mX").onclick = closeModal;
  $("mS").onclick = async () => {
    const name = normName($("eN").value), occ = $("eO").value.trim(); if (!name) { toast("Type the employee's name."); return; }
    const id = isNew ? empId(name) : e.id, old = emps.find(x => x.id === id);
    if (isNew && old && !old.removed) { toast(name + " is already on the matrix."); return; }
    try { await setDoc(ref(id), { kind: "emp", name, occ, removed: false, dates: (old || e).dates || {}, updatedAt: Date.now(), updatedBy: email }, { merge: true }); toast("Saved"); closeModal(); } catch (er) { fail(er); }
  };
  if (!isNew) $("mR").onclick = async () => { if (!confirm("Remove " + e.name + " from the Training Matrix? Their dates are kept and they will not come back when the Excel is imported again.")) return;
    try { await setDoc(ref(e.id), { removed: true, removedAt: Date.now(), removedBy: email }, { merge: true }); toast("Removed"); closeModal(); } catch (er) { fail(er); } };
}

/* ================= Expiring per month ================= */
function viewMonth() {
  const rows = monthRows(month), ed = canEdit();
  let h = `<div class="bar"><div class="mnav"><button class="sm" id="mP" aria-label="Previous month">‹</button><b>${mName(month)}</b><button class="sm" id="mN" aria-label="Next month">›</button></div>
    <input type="month" id="mPick" value="${month}" style="width:auto"><span style="flex:1"></span><button id="xM" class="primary">Download ${mName(month)}</button></div>`;
  h += `<div class="tbl"><table><thead><tr><th class="num">No.</th><th>Employee</th><th>Position</th><th>Training Required</th><th>Expiry Date</th><th>Certificate</th><th>Provider</th><th>Booking date</th><th>Current progress</th>${ed ? "<th></th>" : ""}</tr></thead><tbody>` +
    (rows.length ? rows.map((r, i) => `<tr><td class="num">${i + 1}</td><td>${esc(r.name)}</td><td>${esc(r.pos)}</td><td>${esc(r.course)}</td><td class="${isIso(r.expiry) ? "c-" + (status(r.expiry, r.course) || "ok") : ""}">${esc(fmtD(r.expiry))}</td><td>${esc(r.cert)}</td><td>${esc(r.provider)}</td><td>${esc(fmtD(r.booking))}</td><td class="prog ${progCls(r.progress)}">${esc(r.progress)}${r.newExpiry ? `<div class="note">New expiry ${fmtD(r.newExpiry)}</div>` : ""}</td>${ed ? `<td><button class="sm" data-bk="${esc(r.id)}">Update</button></td>` : ""}</tr>`).join("")
      : `<tr><td colspan="10" class="note">No trainings expire in ${mName(month)}.</td></tr>`) + `</tbody></table></div>`;
  $("main").innerHTML = h;
  $("mP").onclick = () => { month = addM(month, -1); viewMonth(); };
  $("mN").onclick = () => { month = addM(month, 1); viewMonth(); };
  $("mPick").onchange = e => { if (e.target.value) { month = e.target.value; viewMonth(); } };
  $("xM").onclick = () => downloadXlsx(`TKR Training ${mName(month)}.xlsx`, [month], false);
  document.querySelectorAll("[data-bk]").forEach(b => b.onclick = () => bookForm(rows.find(r => r.id === b.dataset.bk)));
}
function bookForm(r) {
  const provs = [...new Set(Object.values(books).map(b => b.provider).filter(Boolean))].sort();
  modal(r.course, `${esc(r.name)} · expires ${esc(fmtD(r.expiry))}`, `<div class="form">
    <label>Certificate<select id="bC"><option value=""></option>${["Yes", "No"].map(v => `<option${v === r.cert ? " selected" : ""}>${v}</option>`).join("")}</select></label>
    <label>Provider<input id="bP" value="${esc(r.provider)}" list="provList"></label>
    <label>Booking date<input type="date" id="bB" value="${isIso(r.booking) ? r.booking : ""}"><span class="note" style="display:flex;gap:6px;align-items:center;margin-top:4px"><input type="checkbox" id="bNR" style="width:auto;min-height:0"${/no booking required/i.test(r.booking) ? " checked" : ""}> No booking required</span></label>
    <label>Current progress<select id="bG"><option value=""></option>${PROGRESS.map(v => `<option${v === r.progress ? " selected" : ""}>${v}</option>`).join("")}</select></label>
    <label class="full">New expiry date<input type="date" id="bE" value="${isIso(r.newExpiry) ? r.newExpiry : ""}"></label></div>
    <datalist id="provList">${provs.map(p => `<option value="${esc(p)}">`).join("")}</datalist>
    <p class="note">When the training is done, type the new expiry date. It goes onto the Training Matrix for ${esc(r.name)}.</p>`,
    `<button id="mX">Cancel</button><button class="primary" id="mS">Save</button>`);
  $("mX").onclick = closeModal;
  $("mS").onclick = async () => {
    const rec = { kind: "book", month: r.month, name: r.name, pos: r.pos || "", course: r.course, expiry: r.expiry || "", cert: $("bC").value, provider: $("bP").value.trim(),
      booking: $("bNR").checked ? "No booking required" : $("bB").value, progress: $("bG").value, newExpiry: $("bE").value, updatedAt: Date.now(), updatedBy: email };
    try {
      await setDoc(ref(r.id), rec);
      const e = emps.find(x => x.id === empId(r.name));
      if (rec.newExpiry && e && meta.courses.includes(r.course) && ((e.dates || {})[r.course] || "") < rec.newExpiry) await setDoc(ref(e.id), { dates: Object.assign({}, e.dates, { [r.course]: rec.newExpiry }), updatedAt: Date.now(), updatedBy: email }, { merge: true });
      toast(rec.newExpiry ? "Saved. The matrix now shows " + fmtD(rec.newExpiry) + "." : "Saved"); closeModal();
    } catch (er) { fail(er); }
  };
}

/* ================= Settings (editors) ================= */
function viewSet() {
  const removed = emps.filter(e => e.removed).sort((a, b) => a.name.localeCompare(b.name));
  $("main").innerHTML = `<div class="card"><h2>Import the Training Matrix Excel</h2>
      <p class="note">Choose the workbook in the same layout (sheet "Training Matrix" with Employee, Occupation and one column per training, and the monthly sheets such as "November 2026"). Dates in the file replace the dates in the app. Employees removed here are not added back.</p>
      <div class="bar"><input type="file" id="imF" accept=".xlsx,.xls" style="width:auto"><button class="primary" id="imGo" disabled>Import</button></div><p class="note" id="imMsg" role="status"></p></div>
    <div class="card"><h2>Employees</h2><div class="bar"><button class="primary" id="addE">Add an employee</button><span class="note">To change or remove an employee, tap their name on the Training Matrix.</span></div>
      ${removed.length ? `<p class="note">Removed: ${removed.map(e => `${esc(e.name)} <button class="sm" data-back="${esc(e.id)}">Put back</button>`).join(" · ")}</p>` : ""}</div>
    <div class="card"><h2>Trainings <span class="note">${meta.courses.length}</span></h2><p class="note">Tick the trainings that are mandatory for everyone: an employee without that date shows in red as missing.</p>
      <div class="chk">${meta.courses.map((c, i) => `<label><input type="checkbox" data-man="${i}"${meta.mandatory.includes(c) ? " checked" : ""}> ${esc(c)}</label>`).join("")}</div>
      <form class="bar" id="addC" style="margin-top:12px"><input id="cName" placeholder="New training column" required style="flex:1 1 260px;width:auto"><button class="primary">Add training</button></form></div>`;
  let parsed = null;
  $("imF").onchange = async e => {
    const f = e.target.files[0]; parsed = null; $("imGo").disabled = true; if (!f) return;
    try { parsed = parseWorkbook(await f.arrayBuffer()); $("imMsg").textContent = foundMsg(parsed); $("imGo").disabled = false; }
    catch (er) { $("imMsg").textContent = "This file could not be read: " + er.message; }
  };
  $("imGo").onclick = async () => { if (!parsed) return; $("imGo").disabled = true; $("imMsg").textContent = "Importing…"; try { const n = await importParsed(parsed); $("imMsg").textContent = importDone(n); toast("Imported"); } catch (er) { fail(er); $("imMsg").textContent = ""; $("imGo").disabled = false; } };
  $("addE").onclick = () => empForm(null);
  document.querySelectorAll("[data-back]").forEach(b => b.onclick = async () => { try { await setDoc(ref(b.dataset.back), { removed: false }, { merge: true }); toast("Put back"); } catch (er) { fail(er); } });
  document.querySelectorAll("[data-man]").forEach(cb => cb.onchange = async () => {
    const c = meta.courses[+cb.dataset.man], man = meta.mandatory.filter(x => x !== c); if (cb.checked) man.push(c);
    try { await setDoc(ref("meta"), { mandatory: man }, { merge: true }); } catch (er) { fail(er); } });
  $("addC").onsubmit = async ev => { ev.preventDefault(); const c = $("cName").value.trim(); if (!c) return; if (meta.courses.some(x => x.toLowerCase() === c.toLowerCase())) { toast("That training is already there."); return; }
    try { await setDoc(ref("meta"), { courses: meta.courses.concat([c]) }, { merge: true }); toast("Training added"); } catch (er) { fail(er); } };
}

/* ---------- reading the workbook ---------- */
function cellDate(v) {
  if (v === "" || v == null) return "";
  if (typeof v === "number" && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  const t = String(v).trim(), m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) return `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
  if (isIso(t)) return t;
  return t;
}
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array" }), clean = s => String(s ?? "").replace(/\s+/g, " ").trim();
  const mSheet = wb.SheetNames.find(n => /training\s*matrix/i.test(n)) || wb.SheetNames[0];
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[mSheet], { header: 1, raw: true, defval: "" });
  const hr = aoa.findIndex(r => r.some(v => /^employee$/i.test(clean(v)))); if (hr < 0) throw new Error("no Employee heading on the Training Matrix sheet");
  const head = aoa[hr].map(clean), cN = head.findIndex(v => /^employee$/i.test(v)), cO = head.findIndex(v => /^occupation|^position/i.test(v));
  const cols = []; head.forEach((v, i) => { if (v && i > Math.max(cN, cO)) cols.push([i, v]); });
  const courses = cols.map(c => c[1]), people = [];
  for (const r of aoa.slice(hr + 1)) { const name = normName(r[cN]); if (/^key$/i.test(name)) break; if (!name) continue; const dates = {}; for (const [i, c] of cols) { const d = cellDate(r[i]); if (d) dates[c] = d; } people.push({ name, occ: cO >= 0 ? clean(r[cO]) : "", dates }); }
  const nk = s => s.toLowerCase().replace(/[^a-z]/g, ""), courseOf = t => { const k = nk(t); return courses.find(c => nk(c) === k) || courses.find(c => k && (nk(c).startsWith(k) || k.startsWith(nk(c)))) || clean(t); };
  const books = [], sheets = [];
  for (const sn of wb.SheetNames) {
    const mm = sn.trim().match(/^([A-Za-z]+)\s+(\d{4})$/), mi = mm ? MONTHS.findIndex(x => x.toLowerCase() === mm[1].toLowerCase()) : -1; if (mi < 0) continue;
    const m = `${mm[2]}-${pad(mi + 1)}`, a = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: "" });
    const h = a.findIndex(r => r.some(v => /^employee$/i.test(clean(v)))); if (h < 0) continue; sheets.push(sn.trim());
    const H = a[h].map(clean), col = re => H.findIndex(v => re.test(v));
    const ix = { name: col(/^employee$/i), pos: col(/^position|^occupation/i), course: col(/^training/i), exp: col(/^expiry/i), cert: col(/^certificate/i), prov: col(/^provider/i), book: col(/^booking/i), prog: col(/^current|prog/i) };
    for (const r of a.slice(h + 1)) {
      const name = normName(r[ix.name]); if (!name) continue;
      const g = k => ix[k] >= 0 ? r[ix[k]] : "";
      let booking = cellDate(g("book")), progress = clean(g("prog"));
      if (booking && !isIso(booking) && /bookings? yet/i.test(booking)) { if (!progress) progress = "No bookings yet"; booking = ""; }
      if (/^complete$/i.test(progress)) progress = "Completed";
      books.push({ month: m, name, pos: clean(g("pos")), course: courseOf(g("course")), expiry: cellDate(g("exp")), cert: clean(g("cert")), provider: clean(g("prov")), booking, progress });
    }
  }
  return { courses, people, books, sheets };
}
async function importParsed(p) {
  const courses = meta.courses.slice(); for (const c of p.courses) if (!courses.some(x => x.toLowerCase() === c.toLowerCase())) courses.push(c);
  const writes = [() => setDoc(ref("meta"), { courses, mandatory: meta.mandatory || [] }, { merge: true })];
  let added = 0, changed = 0, dates = 0, skip = 0, book = 0;
  const seen = new Set();
  for (const x of p.people) {
    const id = empId(x.name), old = emps.find(e => e.id === id); if (seen.has(id)) continue; seen.add(id);
    if (old && old.removed) { skip++; continue; }
    // the file is the master for the trainings in it: its dates replace the app's, and a date taken out of the file is taken out here
    const nd = {}; for (const [c, d] of Object.entries(old?.dates || {})) if (!p.courses.includes(c)) nd[c] = d; Object.assign(nd, x.dates);
    const od = old?.dates || {}, diff = new Set(Object.keys(od).concat(Object.keys(nd))); let n = 0; diff.forEach(c => { if ((od[c] || "") !== (nd[c] || "")) n++; });
    const occ = x.occ || old?.occ || "";
    if (old && !n && occ === (old.occ || "")) continue;
    if (old) { changed++; dates += n; } else added++;
    writes.push(() => setDoc(ref(id), { kind: "emp", name: x.name, occ, removed: false, dates: nd, updatedAt: Date.now(), updatedBy: email }));
  }
  for (const b of p.books) {
    const id = bookId(b.name, b.course, b.month), cur = books[id] || {};
    if (["pos", "expiry", "cert", "provider", "booking", "progress"].every(k => (cur[k] || "") === (b[k] || "")) && cur.course === b.course) continue;
    book++; writes.push(() => setDoc(ref(id), Object.assign({ kind: "book" }, cur, b, { updatedAt: Date.now(), updatedBy: email })));
  }
  for (let i = 0; i < writes.length; i += 40) await Promise.all(writes.slice(i, i + 40).map(f => f()));
  return { added, changed, dates, book, skip };
}
const importDone = n => (n.added || n.changed || n.book) ? `Updated: ${n.added} new employee(s), ${n.dates} date(s) changed for ${n.changed} employee(s), ${n.book} monthly line(s) added or changed.${n.skip ? ` ${n.skip} removed employee(s) skipped.` : ""}` : "Nothing changed: the app already matches this file." + (n.skip ? ` ${n.skip} removed employee(s) skipped.` : "");
const foundMsg = p => `Found ${p.people.length} employees, ${p.courses.length} trainings and ${p.books.length} lines on ${p.sheets.length} monthly sheet(s)${p.sheets.length ? " (" + p.sheets.join(", ") + ")" : ""}.`;
function importPick() {
  const inp = document.createElement("input"); inp.type = "file"; inp.accept = ".xlsx,.xls";
  inp.onchange = async () => {
    const f = inp.files[0]; if (!f) return; let p;
    try { p = parseWorkbook(await f.arrayBuffer()); } catch (er) { toast("This file could not be read: " + er.message); return; }
    modal("Import " + f.name, foundMsg(p), '<p class="note">The dates in the file replace the dates in the app. Employees removed in the app are not added back.</p><p class="note" id="imRes" role="status"></p>',
      '<button id="mX">Cancel</button><button class="primary" id="mS">Import</button>');
    $("mX").onclick = closeModal;
    $("mS").onclick = async () => { $("mS").disabled = true; $("imRes").textContent = "Importing…";
      try { const n = await importParsed(p); $("imRes").textContent = importDone(n); $("mS").hidden = true; $("mX").textContent = "Close"; toast("Imported"); } catch (er) { fail(er); $("imRes").textContent = ""; $("mS").disabled = false; } };
  };
  inp.click();
}

/* ---------- Excel downloads (same layout as the Training Matrix workbook) ---------- */
const loadScript = src => new Promise((res, rej) => { if (window.ExcelJS) return res(); const el = document.createElement("script"); el.src = src; el.onload = res; el.onerror = () => rej(new Error("open the app once with signal")); document.head.appendChild(el); });
let logoBuf = null;
function downloadAll() {
  const ms = new Set(Object.values(books).map(b => b.month)); for (let i = 0; i < 4; i++) ms.add(addM(ym(), i));
  downloadXlsx(`TKR Training Matrix ${today()}.xlsx`, [...ms].filter(m => m && monthRows(m).length).sort(), true);
}
async function downloadXlsx(fname, months, withMatrix) {
  try { await brandedXlsx(fname, months, withMatrix); }
  catch (e) {
    console.warn(e); const wb = XLSX.utils.book_new();
    if (withMatrix) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["TKR Training Matrix"], [], ["No.", "Employee", "Occupation"].concat(meta.courses)].concat(active().map((e, i) => [i + 1, e.name, e.occ].concat(meta.courses.map(c => fmtD((e.dates || {})[c])))))), "Training Matrix");
    for (const m of months) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[mName(m)], ["No.", "Employee", "Position", "Training Required", "Expiry Date", "Certificate", "Provider", "Booking date", "Current progress"]].concat(monthRows(m).map((r, i) => [i + 1, r.name, r.pos, r.course, fmtD(r.expiry), r.cert, r.provider, fmtD(r.booking), r.progress]))), mName(m));
    XLSX.writeFile(wb, fname); toast("Downloaded without the TKR layout: " + (e.message || "try again"));
  }
}
async function brandedXlsx(fname, months, withMatrix) {
  toast("Preparing the Excel…"); await loadScript("../lib/exceljs.min.js");
  if (!logoBuf) { try { logoBuf = await (await fetch("../icons/icon-512.png")).arrayBuffer(); } catch { logoBuf = null; } }
  const FILL = { ex: "FFF3C7C0", m0: "FFF6C98E", m1: "FFFBE79B", ok: "FFC9E8D2", mi: "FFFF0000" };
  const C = { BLACK: "FF111111", HEAD: "FF262626", GOLD: "FFD4A12A", WHITE: "FFFFFFFF", GRID: "FFD9D3C5", MUTED: "FF6B6558" };
  const fill = c => ({ type: "pattern", pattern: "solid", fgColor: { argb: c } }), thin = { style: "thin", color: { argb: C.GRID } }, box = { top: thin, left: thin, bottom: thin, right: thin };
  const X = new ExcelJS.Workbook(); X.creator = "TKR Operations"; X.created = new Date();
  const logo = logoBuf ? X.addImage({ buffer: logoBuf, extension: "png" }) : null;
  const band = (ws, n, title, sub) => {
    for (let r = 1; r <= 3; r++) { ws.getRow(r).height = r === 1 ? 26 : 18; for (let c = 1; c <= n; c++) ws.getCell(r, c).fill = fill(C.BLACK); }
    for (let c = 1; c <= n; c++) ws.getCell(4, c).fill = fill(C.GOLD); ws.getRow(4).height = 4;
    if (logo !== null) ws.addImage(logo, { tl: { col: 0.1, row: 0.15 }, ext: { width: 52, height: 52 } });
    const t = ws.getCell(1, 3); t.value = title; t.font = { name: "Arial", size: 15, bold: true, color: { argb: C.WHITE } }; t.alignment = { vertical: "middle" };
    const s = ws.getCell(2, 3); s.value = sub; s.font = { name: "Arial", size: 9, bold: true, color: { argb: C.GOLD } };
    const d = ws.getCell(3, 3); d.value = "Downloaded " + fmtD(today()); d.font = { name: "Arial", size: 9, color: { argb: "FFCFC8B8" } };
  };
  const headRow = (ws, r, head, vert) => { const row = ws.getRow(r); row.height = vert ? 150 : 30;
    head.forEach((h, i) => { if (h == null) return; const c = row.getCell(i + 1); c.value = h; c.font = { name: "Arial", size: 9, bold: true, color: { argb: C.WHITE } }; c.fill = fill(C.HEAD); c.border = box;
      c.alignment = vert && i >= 3 ? { textRotation: 90, vertical: "bottom", horizontal: "center", wrapText: true } : { vertical: "middle", wrapText: true, horizontal: i < 3 ? "left" : "center" }; }); };
  if (withMatrix) {
    const n = 3 + meta.courses.length, ws = X.addWorksheet("Training Matrix", { views: [{ state: "frozen", xSplit: 3, ySplit: 5, showGridLines: false }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
    ws.columns = [{ width: 6 }, { width: 28 }, { width: 24 }].concat(meta.courses.map(() => ({ width: 11 })));
    band(ws, n, "TKR Training Matrix", "TKR OPERATIONS  ·  TRAINING"); headRow(ws, 5, ["No.", "Employee", "Occupation"].concat(meta.courses), true);
    active().forEach((e, k) => {
      const row = ws.getRow(6 + k); row.getCell(1).value = k + 1; row.getCell(2).value = e.name; row.getCell(3).value = e.occ || "";
      meta.courses.forEach((c, j) => { const d = (e.dates || {})[c], s = status(d, c), cell = row.getCell(4 + j);
        if (d) { if (isIso(d)) { cell.value = new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8))); cell.numFmt = "dd/mm/yyyy"; } else cell.value = d; }
        if (s) cell.fill = fill(FILL[s]); cell.alignment = { horizontal: "center" }; });
      for (let c = 1; c <= n; c++) { const cell = row.getCell(c); cell.border = box; cell.font = { name: "Arial", size: 9, bold: c === 2 }; }
    });
    let r = 6 + active().length + 1; ws.getCell(r, 2).value = "Key"; ws.getCell(r, 2).font = { name: "Arial", bold: true, size: 9 };
    for (const k of ["ex", "m0", "m1", "ok"].concat(meta.mandatory.length ? ["mi"] : [])) { r++; ws.getCell(r, 2).fill = fill(FILL[k]); ws.getCell(r, 2).border = box; ws.getCell(r, 3).value = k === "m0" ? "Expires " + mName(ym()) : k === "m1" ? "Expires " + mName(addM(ym(), 1)) : STATUS[k]; ws.getCell(r, 3).font = { name: "Arial", size: 9 }; }
  }
  const MH = ["No.", "Employee", "Position", "Training Required", "Expiry Date", "Certificate", "Provider", "Booking date", "Current progress"];
  for (const m of months) {
    const ws = X.addWorksheet(mName(m), { views: [{ state: "frozen", ySplit: 5, showGridLines: false }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
    ws.columns = [6, 28, 24, 34, 13, 12, 16, 14, 22].map(w => ({ width: w }));
    band(ws, MH.length, "Trainings expiring " + mName(m), "TKR OPERATIONS  ·  TRAINING MATRIX"); headRow(ws, 5, MH, false);
    monthRows(m).forEach((x, k) => {
      const row = ws.getRow(6 + k);
      [k + 1, x.name, x.pos, x.course, fmtD(x.expiry), x.cert, x.provider, fmtD(x.booking), x.progress].forEach((v, i) => { const c = row.getCell(i + 1); c.value = v === "" ? null : v; c.border = box; c.font = { name: "Arial", size: 10, bold: i === 1 }; c.alignment = { vertical: "middle", wrapText: i === 3, horizontal: i === 0 || i === 4 || i === 5 || i === 7 ? "center" : "left" }; });
      const s = isIso(x.expiry) ? status(x.expiry, x.course) : ""; if (s) row.getCell(5).fill = fill(FILL[s]);
      if (x.progress === "Completed") row.getCell(9).font = { name: "Arial", size: 10, bold: true, color: { argb: "FF2E7D4F" } };
      if (x.progress === "No bookings yet") row.getCell(9).font = { name: "Arial", size: 10, bold: true, color: { argb: "FFB3402F" } };
    });
    if (!monthRows(m).length) ws.getCell(6, 2).value = "No trainings expire in " + mName(m);
  }
  if (!X.worksheets.length) X.addWorksheet("Training").getCell(1, 1).value = "Nothing to download";
  const buf = await X.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a"); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 4000);
  toast("Downloaded " + fname);
}

/* ================= shell ================= */
/* ================= Access (Muneer and Amogelang) ================= */
function viewAccess() {
  const row = (em, role) => `<tr><td>${esc(em)}</td><td>${role === "editors" ? "Full access (add and change)" : "View only"}</td><td><button class="sm danger" data-rm="${esc(em)}" data-role="${role}">Remove</button></td></tr>`;
  const n = tmAcc.editors.length + tmAcc.viewers.length;
  $("main").innerHTML = `<div class="card"><h2>Give access to the Training Matrix</h2><form class="form" id="addUser"><label>Email or username<input id="uEmail" required></label><label>Password<input id="uPass" type="text" minlength="6" placeholder="Leave empty if they already sign in to the TKR app"></label>
    <label>Access<select id="uRole"><option value="editors">Full access (add and change)</option><option value="viewers">View only</option></select></label><div style="display:flex;align-items:end"><button class="primary">Add</button></div></form>
    <p class="note" id="uMsg" role="status">Only Muneer and Amogelang can give or remove access. They always have full access.</p></div>
    <div class="card"><h2>Who has access <span class="note">${n}</span></h2>${n ? `<div class="tbl"><table><thead><tr><th>Email</th><th>Access</th><th></th></tr></thead><tbody>${tmAcc.editors.map(e => row(e, "editors")).join("")}${tmAcc.viewers.map(e => row(e, "viewers")).join("")}</tbody></table></div>` : '<p class="note">Nobody else has access yet.</p>'}</div>`;
  $("addUser").onsubmit = async ev => {
    ev.preventDefault(); const em = toLogin($("uEmail").value).toLowerCase(), pw = $("uPass").value, role = $("uRole").value, other = role === "editors" ? "viewers" : "editors", msg = $("uMsg");
    if (pw && pw.length < 6) { msg.textContent = "The password must be at least 6 characters."; return; }
    msg.textContent = "Adding…"; let note = "";
    if (pw) { const tmp = initializeApp(firebaseConfig, "tm-add-" + Date.now());
      try { await createUserWithEmailAndPassword(getAuth(tmp), em, pw); note = " Their sign-in was made."; }
      catch (er) { if (er?.code === "auth/email-already-in-use") note = " They already had a sign-in, so their password stays the same."; else { msg.textContent = "The sign-in couldn't be made. Check the email and try again."; deleteApp(tmp).catch(() => {}); return; } }
      deleteApp(tmp).catch(() => {}); }
    try { await setDoc(doc(db, "config", "trainingAccess"), { [role]: arrayUnion(em), [other]: arrayRemove(em) }, { merge: true }); toast(em + " now has access." + note); } catch (er) { fail(er); msg.textContent = ""; }
  };
  document.querySelectorAll("[data-rm]").forEach(b => b.onclick = async () => { if (!confirm("Remove access for " + b.dataset.rm + "?")) return; try { await setDoc(doc(db, "config", "trainingAccess"), { [b.dataset.role]: arrayRemove(b.dataset.rm) }, { merge: true }); toast("Access removed"); } catch (er) { fail(er); } });
}

const VIEWS = { dash: viewDash, matrix: viewMatrix, month: viewMonth, set: viewSet, access: viewAccess };
function render() {
  if ((view === "set" && !canEdit()) || (view === "access" && !isAccessAdmin())) view = "dash";
  $("accTab").hidden = !isAccessAdmin();
  document.querySelectorAll("#tabs button").forEach(b => b.setAttribute("aria-selected", b.dataset.v === view));
  $("setTab").hidden = !canEdit();
  if (!loaded) { $("main").innerHTML = '<div class="card note">Loading the training matrix…</div>'; return; }
  VIEWS[view]();
}
$("tabs").onclick = e => { const b = e.target.closest("button"); if (!b) return; view = b.dataset.v; try { localStorage.setItem("tkr-tm-view", view); } catch {} render(); };
try { const v = localStorage.getItem("tkr-tm-view"); if (v && VIEWS[v]) view = v; } catch {}

function show(which) { ["signin", "noaccess", "app"].forEach(id => $(id).hidden = id !== which); $("outBtn").hidden = which === "signin"; }
function startData() {
  if (unsub) unsub();
  unsub = onSnapshot(collection(db, "training"), s => {
    const E = [], B = {}; let M = null;
    s.docs.forEach(d => { const v = Object.assign({ id: d.id }, d.data()); if (d.id === "meta") M = v; else if (v.kind === "emp") E.push(v); else if (v.kind === "book") B[d.id] = v; });
    meta = { courses: M?.courses || [], mandatory: M?.mandatory || [] }; emps = E; books = B; loaded = true; show("app");
    if (!$("modal").innerHTML && !(view === "matrix" && document.activeElement?.id === "fQ")) render();
  }, e => {
    console.warn(e); unsub = null;
    if (e.code === "permission-denied" && isAccessAdmin()) { loaded = true; show("app"); view = "access"; render(); $("main").insertAdjacentHTML("afterbegin", '<div class="card notice">The Training Matrix rules are not published in Firebase yet, so the records cannot load. You can still give access here.</div>'); return; }
    if (e.code === "permission-denied") { $("naEmail").textContent = email; show("noaccess"); if (ADMINS.includes(email)) $("naEmail").textContent = email + " (the Training Matrix rules still need to be published in Firebase)"; }
    else $("main").innerHTML = '<div class="card note">Could not load the training matrix: ' + esc(e.message) + '</div>';
  });
}
onAuthStateChanged(auth, u => {
  if (accUnsub) { accUnsub(); accUnsub = null; } if (unsub) { unsub(); unsub = null; } loaded = false;
  if (!u || !u.email) { email = ""; $("who").textContent = ""; show("signin"); return; }
  email = u.email.toLowerCase(); $("who").textContent = email.replace("@mgr.tkroperations.app", "");
  show("app"); render();
  accUnsub = onSnapshot(doc(db, "config", "trainingAccess"), d => { const v = d.exists() ? d.data() : {}; tmAcc = { editors: (v.editors || []).map(x => x.toLowerCase()), viewers: (v.viewers || []).map(x => x.toLowerCase()) }; if (loaded && !$("modal").innerHTML) render(); }, () => {});
  startData();
});
$("signForm").onsubmit = async e => {
  e.preventDefault(); const m = $("sMsg"); m.textContent = "Signing in…"; $("sBtn").disabled = true;
  try { await signInWithEmailAndPassword(auth, toLogin($("sEmail").value), $("sPass").value); m.textContent = ""; }
  catch { m.textContent = "Wrong email or password."; } finally { $("sBtn").disabled = false; }
};
$("sReset").onclick = async () => { const em = toLogin($("sEmail").value), m = $("sMsg"); if (!em.includes("@") || em.endsWith("@mgr.tkroperations.app")) { m.textContent = "Type your email first, then press Forgot password."; return; }
  try { await sendPasswordResetEmail(auth, em); m.textContent = "A password reset email was sent to " + em + "."; } catch { m.textContent = "The reset email couldn't be sent. Ask Amogelang."; } };
$("outBtn").onclick = $("naOut").onclick = () => signOut(auth);
document.addEventListener("keydown", e => { if (e.key === "Escape" && $("modal").innerHTML) closeModal(); });
if ("serviceWorker" in navigator) navigator.serviceWorker.register("../sw.js", { updateViaCache: "none" }).catch(() => {});
