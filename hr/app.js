/* TKR HR — employee management, employee relations, performance, employment equity and training.
   Shared online through Firebase. Access is given by Amogelang (HR admin) on the Access tab. */
import {
  initializeApp, deleteApp, getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged, createUserWithEmailAndPassword, sendPasswordResetEmail,
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, setDoc, getDoc, deleteDoc, onSnapshot, arrayUnion, arrayRemove,
} from "../lib/firebase.js";

const firebaseConfig = {
  apiKey: "AIzaSyALzOBLObrONGVD1bS0Tz02aNKhQ40TRoI", authDomain: "tkr-operations-app.firebaseapp.com", projectId: "tkr-operations-app",
  storageBucket: "tkr-operations-app.firebasestorage.app", messagingSenderId: "158080247079", appId: "1:158080247079:web:7cb7786d2b7f9ee172b502",
};
// Must match the Firestore rules (firestore-rules.txt, HR block).
const HR_ADMINS = ["amogelang@tkroperations.com", "amogelangmabilu56@gmail.com"];
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
const fmtD = s => { if (!s) return ""; const [y, m, d] = String(s).split("-"); return d ? `${d}/${m}/${y}` : s; };
const kb = n => n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round((n || 0) / 1024)) + " KB";
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
function toast(m) { const t = $("toast"); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, 3200); }
function fail(e) { console.warn(e); toast(e?.code === "permission-denied" ? "You are not allowed to do that." : "Not saved: " + (e?.message || "try again")); }

/* ---------- lists ---------- */
const GENDERS = ["Male", "Female"];
const RACES = ["African", "Coloured", "Indian", "White"];
const LEVELS = [
  "Top management",
  "Senior management",
  "Professionally qualified and experienced specialists and mid-management",
  "Skilled technical and academically qualified workers, junior management, supervisors, foremen and superintendents",
  "Semi-skilled and discretionary decision making",
  "Unskilled and defined decision making",
];
const MGMT = LEVELS.slice(0, 3);
const WARN_TYPES = ["Verbal warning", "Written warning", "Final written warning", "Suspension", "Disciplinary hearing", "Dismissal"];
const PDP_STATUS = ["Not started", "In progress", "Completed"];
const TERM_REASONS = ["Resigned", "Dismissed", "Contract ended", "Retired", "Absconded", "Deceased", "Other"];

/* ---------- state ---------- */
let email = "", access = { editors: [], viewers: [] }, view = "dash";
const D = { employees: [], warnings: [], perf: [], interns: [], bursaries: [] };
const loaded = {};
let unsubs = [];
const isHrAdmin = () => HR_ADMINS.includes(email);
const canEdit = () => isHrAdmin() || access.editors.includes(email);
const canView = () => canEdit() || access.viewers.includes(email);
const col = name => collection(db, "hr", name, "items");
const ref = (name, id) => doc(db, "hr", name, "items", id);
const empById = id => D.employees.find(e => e.id === id);
const empName = id => empById(id)?.name || "";
const activeEmps = () => D.employees.filter(e => e.status !== "Terminated").sort((a, b) => a.name.localeCompare(b.name));
const empOptions = (sel, list = activeEmps()) => '<option value="">Choose the employee</option>' + list.map(e => `<option value="${esc(e.id)}"${e.id === sel ? " selected" : ""}>${esc(e.name)}${e.number ? " (" + esc(e.number) + ")" : ""}</option>`).join("");
const opts = (list, sel, blank) => (blank ? `<option value="">${esc(blank)}</option>` : "") + list.map(v => `<option${v === sel ? " selected" : ""}>${esc(v)}</option>`).join("");

/* ---------- files (PDF or photo), kept in Firestore in parts ---------- */
const fileB64 = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = rej; r.readAsDataURL(f); });
async function shrinkImg(file, max = 1600) {
  const url = URL.createObjectURL(file); const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; });
  const k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement("canvas"); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
  const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
  return c.toDataURL("image/jpeg", 0.75).split(",")[1];
}
async function putFile(file) {
  let b64, mime = file.type || "";
  if (mime.startsWith("image/")) { b64 = await shrinkImg(file); mime = "image/jpeg"; }
  else if (mime === "application/pdf" || /\.pdf$/i.test(file.name)) { if (file.size > 8 * 1048576) throw new Error("the PDF is bigger than 8 MB"); b64 = await fileB64(file); mime = "application/pdf"; }
  else throw new Error("choose a PDF or a photo");
  const id = "f_" + uid(), CH = 700000, n = Math.ceil(b64.length / CH);
  for (let i = 0; i < n; i++) await setDoc(ref("files", id + "_" + i), { i, data: b64.slice(i * CH, (i + 1) * CH) });
  return { id, name: file.name, size: file.size, mime, n, at: Date.now(), by: email };
}
async function openFile(f) {
  const w = window.open("", "_blank");
  try {
    toast("Opening " + f.name + "…"); const parts = [];
    for (let i = 0; i < f.n; i++) { const d = await getDoc(ref("files", f.id + "_" + i)); if (!d.exists()) throw new Error("part of the file is missing"); parts.push(d.data().data); }
    const bin = atob(parts.join("")), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([u], { type: f.mime }));
    if (w && !w.closed) w.location.href = url; else { const a = document.createElement("a"); a.href = url; a.download = f.name; document.body.appendChild(a); a.click(); a.remove(); }
  } catch (e) { if (w) w.close(); toast("Could not open: " + (e.message || "try again")); }
}
async function dropFile(f) { for (let i = 0; i < f.n; i++) await deleteDoc(ref("files", f.id + "_" + i)).catch(() => {}); }
const fileList = (files, owner) => (files || []).length ? '<div class="files">' + files.map((f, i) => `<div class="file"><span>${esc(f.name)} <span class="note">${kb(f.size)}${f.at ? " · " + new Date(f.at).toLocaleDateString("en-ZA") : ""}</span></span><span><button class="sm" data-open="${owner}|${i}">Open</button></span></div>`).join("") + '</div>' : '<span class="note">No files</span>';

/* ---------- saving ---------- */
async function save(name, rec) { rec.id = rec.id || name.slice(0, 3) + "_" + uid(); rec.updatedAt = Date.now(); rec.updatedBy = email; await setDoc(ref(name, rec.id), rec); return rec; }
async function remove(name, rec) { for (const f of rec.files || []) await dropFile(f); await deleteDoc(ref(name, rec.id)); }

/* ---------- Excel (TKR branded) ---------- */
const loadScript = src => new Promise((res, rej) => { if (window.ExcelJS) return res(); const el = document.createElement("script"); el.src = src; el.onload = res; el.onerror = () => rej(new Error("open the app once with signal")); document.head.appendChild(el); });
let logoBuf = null;
async function brandedXlsx(fname, sheets) {
  toast("Preparing the Excel…"); await loadScript("../lib/exceljs.min.js");
  if (!logoBuf) { try { logoBuf = await (await fetch("../icons/icon-512.png")).arrayBuffer(); } catch { logoBuf = null; } }
  const C = { BLACK: "FF111111", HEAD: "FF262626", GOLD: "FFD4A12A", CREAM: "FFFFF8E6", TOT: "FFFFF2CC", WHITE: "FFFFFFFF", GRID: "FFD9D3C5", MUTED: "FF6B6558" };
  const fill = c => ({ type: "pattern", pattern: "solid", fgColor: { argb: c } }), thin = { style: "thin", color: { argb: C.GRID } }, box = { top: thin, left: thin, bottom: thin, right: thin };
  const X = new ExcelJS.Workbook(); X.creator = "TKR Operations"; X.created = new Date();
  const logo = logoBuf ? X.addImage({ buffer: logoBuf, extension: "png" }) : null;
  for (const [name, aoaIn, widths] of sheets) {
    let aoa = aoaIn.slice(), title = name;
    if (aoa.length > 1 && aoa[0].filter(v => v !== "" && v != null).length <= 2 && (!aoa[1] || !aoa[1].length)) { title = String(aoa[0][0]); aoa = aoa.slice(2); }
    const head = aoa[0] || [], rows = aoa.slice(1), n = Math.max(4, head.length);
    const ws = X.addWorksheet(name.slice(0, 31), { views: [{ showGridLines: false, state: "frozen", ySplit: 7 }], pageSetup: { orientation: n > 6 ? "landscape" : "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 } });
    ws.columns = Array.from({ length: n }, (_, i) => ({ width: Math.max((widths || [])[i] || 12, i === 0 ? 18 : 8) }));
    for (let r = 1; r <= 4; r++) { ws.getRow(r).height = r === 1 ? 8 : 20; for (let c = 1; c <= n; c++) ws.getCell(r, c).fill = fill(C.BLACK); }
    ws.getRow(5).height = 4; for (let c = 1; c <= n; c++) ws.getCell(5, c).fill = fill(C.GOLD);
    if (logo !== null) ws.addImage(logo, { tl: { col: 0.08, row: 1.08 }, ext: { width: 56, height: 56 } });
    const put = (r, v, font) => { const c = ws.getCell(r, 2); c.value = v; c.font = Object.assign({ name: "Arial" }, font); c.alignment = { vertical: "middle" }; };
    put(2, title, { size: 15, bold: true, color: { argb: C.WHITE } });
    put(3, "TKR OPERATIONS  ·  HUMAN RESOURCES", { size: 9, bold: true, color: { argb: C.GOLD } });
    put(4, "Downloaded " + fmtD(today()) + (email ? "  ·  " + email.replace("@mgr.tkroperations.app", "") : "") + "  ·  Confidential", { size: 9, color: { argb: "FFCFC8B8" } });
    const hr = ws.getRow(7); hr.height = 30;
    head.forEach((h, i) => { const c = hr.getCell(i + 1); c.value = h; c.font = { name: "Arial", size: 10, bold: true, color: { argb: C.WHITE } }; c.fill = fill(C.HEAD); c.border = box; c.alignment = { vertical: "middle", wrapText: true, horizontal: i === 0 ? "left" : "center" }; });
    rows.forEach((r, k) => {
      const row = ws.getRow(8 + k), tot = String(r[0] || "") === "Total";
      for (let i = 0; i < n; i++) {
        const v = r[i] ?? "", c = row.getCell(i + 1); c.value = v === "" ? null : v; c.border = box;
        c.font = { name: "Arial", size: 10, bold: tot || i === 0 }; c.alignment = { vertical: "middle", wrapText: typeof v === "string" && v.length > 28, horizontal: typeof v === "number" ? "center" : "left" };
        if (tot) { c.fill = fill(C.TOT); c.border = { top: { style: "medium", color: { argb: C.GOLD } }, left: thin, bottom: thin, right: thin }; } else if (k % 2) c.fill = fill(C.CREAM);
      }
    });
    const end = 8 + rows.length + 1, nc = ws.getCell(end, 1);
    nc.value = rows.length ? rows.filter(r => String(r[0] || "") !== "Total").length + " row(s)" : "No records"; nc.font = { name: "Arial", size: 9, italic: true, color: { argb: C.MUTED } };
  }
  const buf = await X.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a"); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 4000);
  toast("Downloaded " + fname);
}
function downloadXlsx(fname, sheets) {
  brandedXlsx(fname, sheets).catch(e => {
    console.warn(e); const wb = XLSX.utils.book_new();
    sheets.forEach(([title, aoa, widths]) => { const ws = XLSX.utils.aoa_to_sheet(aoa); if (widths) ws["!cols"] = widths.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, title.slice(0, 31)); });
    XLSX.writeFile(wb, fname); toast("Downloaded without the TKR layout: " + (e.message || "try again"));
  });
}

/* ---------- modal ---------- */
function modal(title, sub, body, acts) {
  $("modal").innerHTML = `<div class="scrim" id="scrim"><div class="panel" role="dialog" aria-modal="true"><h2>${esc(title)}</h2>${sub ? `<p class="note" style="margin:0 0 12px">${sub}</p>` : ""}${body}<div class="acts">${acts}</div></div></div>`;
  $("scrim").onclick = e => { if (e.target.id === "scrim") closeModal(); };
}
const closeModal = () => { $("modal").innerHTML = ""; if (typeof render === "function") render(); };

/* ================= Dashboard ================= */
function viewDash() {
  const act = D.employees.filter(e => e.status !== "Terminated"), term = D.employees.filter(e => e.status === "Terminated");
  const sa = act.filter(e => e.foreign !== "Yes" && e.race);
  const hdsa = sa.filter(e => e.race !== "White").length, white = sa.filter(e => e.race === "White").length;
  const pct = n => sa.length ? Math.round(n / sa.length * 1000) / 10 + "%" : "0%";
  const fem = act.filter(e => e.gender === "Female").length;
  const mgmt = act.filter(e => MGMT.includes(e.level)), femMgmt = mgmt.filter(e => e.gender === "Female").length;
  const dis = act.filter(e => e.disability === "Yes").length;
  const kpi = (n, l, s) => `<div class="kpi"><span>${l}</span><b>${n}</b>${s ? `<small>${s}</small>` : ""}</div>`;
  let h = '<div class="kpis">' +
    kpi(act.length, "Active employees") +
    kpi(term.length, "Terminated employees") +
    kpi(fem, "Females", act.length ? Math.round(fem / act.length * 1000) / 10 + "% of active employees" : "") +
    kpi(pct(hdsa), "HDSA", `${hdsa} HDSA against ${white} white (${pct(white)})`) +
    kpi(femMgmt, "Females in management", `of ${mgmt.length} in management`) +
    kpi(dis, "Disabled employees") + '</div>';
  const depts = {}; act.forEach(e => { const k = e.department || "No department"; depts[k] = depts[k] || { n: 0, f: 0 }; depts[k].n++; if (e.gender === "Female") depts[k].f++; });
  const dk = Object.keys(depts).sort();
  h += '<div class="card"><h2>Active employees per department</h2>' + (dk.length ? '<div class="tbl"><table><thead><tr><th>Department</th><th class="num">Employees</th><th class="num">Females</th></tr></thead><tbody>' +
    dk.map(k => `<tr><td>${esc(k)}</td><td class="num">${depts[k].n}</td><td class="num">${depts[k].f}</td></tr>`).join("") + '</tbody></table></div>' : '<p class="note">No employees yet. Add them on the Employees tab.</p>') + '</div>';
  h += '<p class="note">HDSA = African, Coloured and Indian South African employees, shown against white employees (foreign nationals left out). Management = top, senior and mid-management occupational levels.</p>';
  $("main").innerHTML = h;
}

/* ================= Employees ================= */
const EMP_FIELDS = [
  ["number", "Employee number"], ["name", "Full name"], ["idNumber", "ID number"], ["gender", "Gender", GENDERS], ["race", "Race", RACES],
  ["foreign", "Foreign national", ["No", "Yes"]], ["disability", "Disability", ["No", "Yes"]], ["occupation", "Occupation (job title)"], ["level", "Occupational level", LEVELS],
  ["department", "Department"], ["supervisor", "Supervisor"], ["startDate", "Start date", "date"], ["phone", "Phone"],
];
let empQ = "", empShow = "Active";
function viewEmps() {
  const ed = canEdit();
  let h = '<div class="bar"><input type="search" id="eq" placeholder="Search name, number, occupation" value="' + esc(empQ) + '">' +
    `<select id="eshow">${opts(["Active", "Terminated", "All"], empShow)}</select>` +
    (ed ? '<button class="primary" id="eAdd">Add employee</button><button id="eImport">Import from Excel</button><button id="eAtt">Bring in from Attendance</button>' : "") +
    '<button id="eXls">Download Excel</button><input type="file" id="eFile" accept=".xlsx,.xls,.csv" hidden></div><div id="eList"></div>';
  $("main").innerHTML = h;
  const draw = () => {
    const q = empQ.trim().toLowerCase();
    const list = D.employees.filter(e => (empShow === "All" || (empShow === "Active" ? e.status !== "Terminated" : e.status === "Terminated")) &&
      (!q || [e.name, e.number, e.occupation, e.department, e.supervisor].join(" ").toLowerCase().includes(q)))
      .sort((a, b) => (a.department || "~").localeCompare(b.department || "~") || (a.supervisor || "~").localeCompare(b.supervisor || "~") || a.name.localeCompare(b.name));
    if (!list.length) { $("eList").innerHTML = '<div class="card note">No employees to show.' + (ed && !D.employees.length ? " Add them one by one, import them from Excel or bring them in from Attendance." : "") + '</div>'; return; }
    let rows = "", dep = null, sup = null;
    list.forEach(e => {
      if ((e.department || "No department") !== dep) { dep = e.department || "No department"; sup = null; rows += `<tr class="grp"><td colspan="6">${esc(dep)} <span class="note">${list.filter(x => (x.department || "No department") === dep).length}</span></td></tr>`; }
      if ((e.supervisor || "No supervisor") !== sup) { sup = e.supervisor || "No supervisor"; rows += `<tr class="sub"><td colspan="6">Supervisor: ${esc(sup)}</td></tr>`; }
      rows += `<tr><td><b>${esc(e.name)}</b><br><span class="note">${esc(e.number || "")}</span></td><td>${esc(e.occupation || "")}</td><td>${esc(e.gender || "")}${e.race ? " · " + esc(e.race) : ""}${e.disability === "Yes" ? " · Disabled" : ""}</td><td>${fmtD(e.startDate)}</td>` +
        `<td>${e.status === "Terminated" ? `<span class="st off">Terminated</span><br><span class="note">${fmtD(e.termDate)} ${esc(e.termReason || "")}</span>` : '<span class="st on">Active</span>'}</td><td style="white-space:nowrap"><button class="sm" data-eopen="${esc(e.id)}">${ed ? "Edit" : "View"}</button></td></tr>`;
    });
    $("eList").innerHTML = `<div class="tbl"><table><thead><tr><th>Employee</th><th>Occupation</th><th>Gender · race</th><th>Start date</th><th>Status</th><th></th></tr></thead><tbody>${rows}</tbody></table></div><p class="note">${list.length} employee${list.length === 1 ? "" : "s"} · sorted by department and supervisor</p>`;
    $("eList").querySelectorAll("[data-eopen]").forEach(b => b.onclick = () => empForm(empById(b.dataset.eopen)));
  };
  draw();
  $("eq").oninput = e => { empQ = e.target.value; draw(); };
  $("eshow").onchange = e => { empShow = e.target.value; draw(); };
  $("eXls").onclick = () => {
    const head = EMP_FIELDS.map(f => f[1]).concat(["Status", "Termination date", "Termination reason"]);
    const rows = D.employees.slice().sort((a, b) => (a.department || "").localeCompare(b.department || "") || (a.supervisor || "").localeCompare(b.supervisor || "") || a.name.localeCompare(b.name))
      .map(e => EMP_FIELDS.map(([k, , t]) => t === "date" ? fmtD(e[k]) : (e[k] || "")).concat([e.status || "Active", fmtD(e.termDate), e.termReason || ""]));
    downloadXlsx("TKR HR Employee list " + today() + ".xlsx", [["Employees", [["Employee list"], [], head].concat(rows), head.map(h => Math.max(12, h.length + 2))]]);
  };
  if (!ed) return;
  $("eAdd").onclick = () => empForm(null);
  $("eImport").onclick = () => $("eFile").click();
  $("eFile").onchange = ev => { const f = ev.target.files[0]; ev.target.value = ""; if (f) importEmps(f); };
  $("eAtt").onclick = fromAttendance;
}
function empForm(e) {
  const ed = canEdit(), x = e || { status: "Active", foreign: "No", disability: "No" };
  const fld = ([k, l, t]) => Array.isArray(t) ? `<label>${l}<select data-k="${k}"${ed ? "" : " disabled"}>${opts(t, x[k], "Choose")}</select></label>`
    : `<label>${l}<input data-k="${k}" ${t === "date" ? 'type="date"' : ""} value="${esc(x[k] || "")}"${ed ? "" : " disabled"}></label>`;
  let body = '<div class="form">' + EMP_FIELDS.map(fld).join("") + '</div>';
  if (e && (e.history || []).length) body += '<h3>Occupation changes</h3><ul class="hist">' + e.history.map(hh => `<li>${fmtD(hh.date)}: ${esc(hh.from || "")} → ${esc(hh.to || "")}</li>`).join("") + '</ul>';
  if (e && e.status === "Terminated") body += `<p class="note">Terminated on ${fmtD(e.termDate)} · ${esc(e.termReason || "")}${e.termNote ? " · " + esc(e.termNote) : ""}</p>`;
  const acts = (ed ? (e ? (e.status === "Terminated" ? '<button id="mReinstate">Reinstate</button>' : '<button class="danger" id="mTerm">Terminate</button>') + '<button class="danger" id="mDel">Delete</button>' : "") + '<button class="primary" id="mSave">Save</button>' : "") + '<button id="mClose">Close</button>';
  modal(e ? e.name : "Add employee", e ? esc([e.number, e.occupation, e.department].filter(Boolean).join(" · ")) : "", body, acts);
  $("mClose").onclick = closeModal;
  if (!ed) return;
  $("mSave").onclick = async () => {
    const rec = Object.assign({}, x); document.querySelectorAll("#modal [data-k]").forEach(i => rec[i.dataset.k] = i.value.trim());
    if (!rec.name) { toast("Enter the employee's full name"); return; }
    if (rec.number && D.employees.some(o => o.id !== rec.id && (o.number || "").toUpperCase() === rec.number.toUpperCase())) { toast("Another employee already has number " + rec.number); return; }
    if (e && (e.occupation || "") !== rec.occupation) rec.history = (e.history || []).concat([{ date: today(), from: e.occupation || "", to: rec.occupation, by: email }]);
    try { await save("employees", rec); toast(e ? "Saved" : rec.name + " added"); closeModal(); } catch (er) { fail(er); }
  };
  if (!e) return;
  if ($("mTerm")) $("mTerm").onclick = () => {
    modal("Terminate " + e.name, "", `<div class="form"><label>Termination date<input type="date" id="tDate" value="${today()}"></label><label>Reason<select id="tWhy">${opts(TERM_REASONS)}</select></label><label class="full">Note (optional)<input id="tNote" maxlength="200"></label></div>`,
      '<button id="tGo" class="danger">Terminate</button><button id="tBack">Cancel</button>');
    $("tBack").onclick = closeModal;
    $("tGo").onclick = async () => { try { await save("employees", Object.assign({}, e, { status: "Terminated", termDate: $("tDate").value, termReason: $("tWhy").value, termNote: $("tNote").value.trim() })); toast(e.name + " terminated"); closeModal(); } catch (er) { fail(er); } };
  };
  if ($("mReinstate")) $("mReinstate").onclick = async () => { try { await save("employees", Object.assign({}, e, { status: "Active", termDate: "", termReason: "", termNote: "" })); toast(e.name + " reinstated"); closeModal(); } catch (er) { fail(er); } };
  $("mDel").onclick = async () => { if (!confirm("Delete " + e.name + " completely? Use Terminate to keep their record.")) return; try { await remove("employees", e); toast("Deleted"); closeModal(); } catch (er) { fail(er); } };
}
const normKey = s => String(s || "").toLowerCase().replace(/[^a-z]/g, "");
async function importEmps(file) {
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false, dateNF: "yyyy-mm-dd" });
    const map = {}; EMP_FIELDS.forEach(([k, l]) => { map[normKey(l)] = k; map[normKey(k)] = k; });
    Object.assign(map, { name: "name", surname: "surname", employeeno: "number", empno: "number", number: "number", jobtitle: "occupation", title: "occupation", position: "occupation", dept: "department", sex: "gender", startdate: "startDate", dateengaged: "startDate" });
    let added = 0, updated = 0;
    for (const r of rows) {
      const rec = {}; for (const c in r) { const k = map[normKey(c)]; if (k && String(r[c]).trim()) rec[k] = String(r[c]).trim(); }
      if (rec.surname) { rec.name = ((rec.name || "") + " " + rec.surname).trim(); delete rec.surname; }
      if (!rec.name) continue;
      if (rec.startDate && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(rec.startDate)) { const [d, m, y] = rec.startDate.split("/"); rec.startDate = `${y}-${pad(m)}-${pad(d)}`; }
      const old = D.employees.find(o => (rec.number && (o.number || "").toUpperCase() === rec.number.toUpperCase()) || o.name.toLowerCase() === rec.name.toLowerCase());
      await save("employees", Object.assign({ status: "Active", foreign: "No", disability: "No" }, old || {}, rec)); old ? updated++ : added++;
    }
    toast(`${added} added, ${updated} updated from ${file.name}`);
  } catch (e) { fail(e); }
}
async function fromAttendance() {
  try {
    const d = await getDoc(doc(db, "config", "attEmployees"));
    const list = (d.exists() ? d.data().list || [] : []).filter(e => e.active !== false);
    if (!list.length) { toast("No employees found in Attendance"); return; }
    if (!confirm(`Bring in ${list.length} active employees from Attendance? People already in HR (same number or name) are skipped.`)) return;
    let n = 0;
    for (const a of list) {
      if (D.employees.some(o => (a.number && (o.number || "").toUpperCase() === String(a.number).toUpperCase()) || o.name.toLowerCase() === String(a.name || "").toLowerCase())) continue;
      await save("employees", { name: a.name || "", number: a.number || "", occupation: a.title || "", department: a.dept || a.sheet || "", status: "Active", foreign: "No", disability: "No" }); n++;
    }
    toast(n + " employees brought in from Attendance. Fill in gender, race, level and supervisor for each.");
  } catch (e) { e?.code === "permission-denied" ? toast("You need Attendance access to bring employees in from there") : fail(e); }
}

/* ================= Employee relations ================= */
let erFrom = "", erTo = "", erType = "All";
function viewEr() {
  const ed = canEdit();
  let h = '<div class="bar">' + (ed ? '<button class="primary" id="wAdd">Add warning</button>' : "") +
    `<label style="flex-direction:row;align-items:center;gap:6px">From <input type="date" id="wFrom" value="${erFrom}" style="width:auto"></label><label style="flex-direction:row;align-items:center;gap:6px">To <input type="date" id="wTo" value="${erTo}" style="width:auto"></label>` +
    `<select id="wType">${opts(["All"].concat(WARN_TYPES), erType)}</select><button id="wXls">Download report</button></div><div id="wList"></div>`;
  $("main").innerHTML = h;
  const list = () => D.warnings.filter(w => (!erFrom || w.date >= erFrom) && (!erTo || w.date <= erTo) && (erType === "All" || w.type === erType)).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const draw = () => {
    const l = list();
    $("wList").innerHTML = l.length ? `<div class="tbl"><table><thead><tr><th>Date</th><th>Employee</th><th>Warning</th><th>Offence</th><th>Action taken</th><th>Issued by</th><th></th></tr></thead><tbody>` +
      l.map(w => { const e = empById(w.empId); return `<tr><td>${fmtD(w.date)}</td><td><b>${esc(e?.name || w.empName || "")}</b><br><span class="note">${esc(e?.department || "")}</span></td><td>${esc(w.type)}${w.validUntil ? `<br><span class="note">valid until ${fmtD(w.validUntil)}</span>` : ""}</td><td>${esc(w.offence || "")}</td><td>${esc(w.action || "")}</td><td>${esc(w.issuedBy || "")}</td><td><button class="sm" data-wopen="${esc(w.id)}">${ed ? "Edit" : "View"}</button></td></tr>`; }).join("") +
      `</tbody></table></div><p class="note">${l.length} record${l.length === 1 ? "" : "s"}</p>` : '<div class="card note">No warnings recorded for this selection.</div>';
    $("wList").querySelectorAll("[data-wopen]").forEach(b => b.onclick = () => warnForm(D.warnings.find(w => w.id === b.dataset.wopen)));
  };
  draw();
  $("wFrom").onchange = e => { erFrom = e.target.value; draw(); };
  $("wTo").onchange = e => { erTo = e.target.value; draw(); };
  $("wType").onchange = e => { erType = e.target.value; draw(); };
  $("wXls").onclick = () => {
    const head = ["Date", "Employee", "Employee number", "Department", "Supervisor", "Warning", "Offence", "Action taken", "Issued by", "Valid until", "Notes"];
    const rows = list().map(w => { const e = empById(w.empId) || {}; return [fmtD(w.date), e.name || w.empName || "", e.number || "", e.department || "", e.supervisor || "", w.type, w.offence || "", w.action || "", w.issuedBy || "", fmtD(w.validUntil), w.notes || ""]; });
    const by = {}; list().forEach(w => by[w.type] = (by[w.type] || 0) + 1);
    downloadXlsx(`TKR HR Employee relations report ${today()}.xlsx`, [["Warnings and actions", [[`Employee relations report${erFrom || erTo ? ` – ${fmtD(erFrom) || "start"} to ${fmtD(erTo) || "today"}` : ""}`], [], head].concat(rows), [12, 26, 14, 18, 18, 20, 34, 34, 18, 12, 30]],
      ["Summary", [["Warnings summary"], [], ["Warning", "Number"]].concat(WARN_TYPES.map(t => [t, by[t] || 0])).concat([["Total", rows.length]]), [24, 10]]]);
  };
  if (ed) $("wAdd").onclick = () => warnForm(null);
}
function warnForm(w) {
  const ed = canEdit(), x = w || { date: today(), type: WARN_TYPES[0] };
  const dis = ed ? "" : " disabled";
  const body = `<div class="form"><label class="full">Employee<select id="fEmp"${dis}>${empOptions(x.empId, D.employees.slice().sort((a, b) => a.name.localeCompare(b.name)))}</select></label>
    <label>Date issued<input type="date" id="fDate" value="${esc(x.date || "")}"${dis}></label><label>Warning<select id="fType"${dis}>${opts(WARN_TYPES, x.type)}</select></label>
    <label>Valid until (optional)<input type="date" id="fValid" value="${esc(x.validUntil || "")}"${dis}></label><label>Issued by<input id="fBy" value="${esc(x.issuedBy || "")}"${dis}></label>
    <label class="full">Offence<textarea id="fOff" rows="2"${dis}>${esc(x.offence || "")}</textarea></label>
    <label class="full">Action taken<textarea id="fAct" rows="2"${dis}>${esc(x.action || "")}</textarea></label>
    <label class="full">Notes<textarea id="fNotes" rows="2"${dis}>${esc(x.notes || "")}</textarea></label>
    <div class="full"><div class="note" style="margin-bottom:6px">Signed warning (PDF or photo)</div>${fileList(x.files, "w")}${ed ? '<input type="file" id="fFile" accept="application/pdf,image/*" style="margin-top:8px">' : ""}</div></div>`;
  modal(w ? "Warning" : "Add warning", "", body, (ed ? (w ? '<button class="danger" id="mDel">Delete</button>' : "") + '<button class="primary" id="mSave">Save</button>' : "") + '<button id="mClose">Close</button>');
  $("mClose").onclick = closeModal; wireOpen(x);
  if (!ed) return;
  $("mSave").onclick = async () => {
    const rec = Object.assign({}, x, { empId: $("fEmp").value, date: $("fDate").value, type: $("fType").value, validUntil: $("fValid").value, issuedBy: $("fBy").value.trim(), offence: $("fOff").value.trim(), action: $("fAct").value.trim(), notes: $("fNotes").value.trim() });
    if (!rec.empId) { toast("Choose the employee"); return; } rec.empName = empName(rec.empId);
    const b = $("mSave"); b.disabled = true; b.textContent = "Saving…";
    try { const f = $("fFile").files[0]; if (f) rec.files = (rec.files || []).concat([await putFile(f)]); await save("warnings", rec); toast("Saved"); closeModal(); } catch (er) { fail(er); b.disabled = false; b.textContent = "Save"; }
  };
  if (w) $("mDel").onclick = async () => { if (!confirm("Delete this warning?")) return; try { await remove("warnings", w); toast("Deleted"); closeModal(); } catch (er) { fail(er); } };
}
function wireOpen(rec) { document.querySelectorAll("#modal [data-open]").forEach(b => b.onclick = () => { const i = +b.dataset.open.split("|")[1]; openFile(rec.files[i]); }); }

/* ================= Performance ================= */
let perfKind = "Review";
function viewPerf() {
  const ed = canEdit();
  let h = `<div class="bar"><div class="tabs" style="margin:0;border:0">${["Review", "PDP"].map(k => `<button data-pk="${k}" aria-selected="${perfKind === k}">${k === "Review" ? "Completed reviews" : "Personal development plans"}</button>`).join("")}</div>` +
    (ed ? `<button class="primary" id="pAdd">${perfKind === "Review" ? "Upload review" : "Add development plan"}</button>` : "") + '<button id="pXls">Download Excel</button></div>';
  const l = D.perf.filter(p => p.kind === perfKind).sort((a, b) => (b.date || b.start || "").localeCompare(a.date || a.start || ""));
  if (perfKind === "Review") h += l.length ? `<div class="tbl"><table><thead><tr><th>Employee</th><th>Review period</th><th>Review date</th><th>Reviewer</th><th>Files</th><th></th></tr></thead><tbody>` +
    l.map(p => `<tr><td><b>${esc(empName(p.empId) || p.empName)}</b></td><td>${esc(p.period || "")}</td><td>${fmtD(p.date)}</td><td>${esc(p.reviewer || "")}</td><td>${(p.files || []).length}</td><td><button class="sm" data-popen="${esc(p.id)}">${ed ? "Edit" : "View"}</button></td></tr>`).join("") + '</tbody></table></div>' : '<div class="card note">No completed reviews uploaded yet.</div>';
  else h += l.length ? `<div class="tbl"><table><thead><tr><th>Employee</th><th>Development area</th><th>Start</th><th>Due</th><th>Status</th><th></th></tr></thead><tbody>` +
    l.map(p => { const late = p.status !== "Completed" && p.due && p.due < today(); return `<tr><td><b>${esc(empName(p.empId) || p.empName)}</b></td><td>${esc(p.goal || "")}</td><td>${fmtD(p.start)}</td><td>${fmtD(p.due)}${late ? ' <span class="st off">Overdue</span>' : ""}</td><td><span class="st ${p.status === "Completed" ? "on" : p.status === "In progress" ? "mid" : ""}">${esc(p.status || "Not started")}</span></td><td><button class="sm" data-popen="${esc(p.id)}">${ed ? "Edit" : "View"}</button></td></tr>`; }).join("") + '</tbody></table></div>' : '<div class="card note">No personal development plans yet.</div>';
  $("main").innerHTML = h;
  document.querySelectorAll("[data-pk]").forEach(b => b.onclick = () => { perfKind = b.dataset.pk; viewPerf(); });
  document.querySelectorAll("[data-popen]").forEach(b => b.onclick = () => perfForm(D.perf.find(p => p.id === b.dataset.popen)));
  if (ed) $("pAdd").onclick = () => perfForm(null);
  $("pXls").onclick = () => {
    const rv = D.perf.filter(p => p.kind === "Review"), pd = D.perf.filter(p => p.kind === "PDP");
    downloadXlsx(`TKR HR Performance ${today()}.xlsx`, [
      ["Completed reviews", [["Completed performance reviews"], [], ["Employee", "Department", "Review period", "Review date", "Reviewer", "Comments", "Files"]].concat(rv.map(p => [empName(p.empId) || p.empName, empById(p.empId)?.department || "", p.period || "", fmtD(p.date), p.reviewer || "", p.notes || "", (p.files || []).map(f => f.name).join(", ")])), [26, 18, 18, 12, 18, 34, 30]],
      ["Development plans", [["Personal development plans"], [], ["Employee", "Department", "Development area", "Actions", "Start", "Due", "Status", "Files"]].concat(pd.map(p => [empName(p.empId) || p.empName, empById(p.empId)?.department || "", p.goal || "", p.notes || "", fmtD(p.start), fmtD(p.due), p.status || "", (p.files || []).map(f => f.name).join(", ")])), [26, 18, 30, 34, 12, 12, 12, 30]]]);
  };
}
function perfForm(p) {
  const ed = canEdit(), kind = p ? p.kind : perfKind, x = p || (kind === "Review" ? { kind, date: today() } : { kind, start: today(), status: "Not started" }), dis = ed ? "" : " disabled";
  const top = `<label class="full">Employee<select id="fEmp"${dis}>${empOptions(x.empId, D.employees.slice().sort((a, b) => a.name.localeCompare(b.name)))}</select></label>`;
  const mid = kind === "Review"
    ? `<label>Review period<input id="fPeriod" placeholder="e.g. Jan – Jun 2026" value="${esc(x.period || "")}"${dis}></label><label>Review date<input type="date" id="fDate" value="${esc(x.date || "")}"${dis}></label><label>Reviewer<input id="fRev" value="${esc(x.reviewer || "")}"${dis}></label><label class="full">Comments<textarea id="fNotes" rows="2"${dis}>${esc(x.notes || "")}</textarea></label>`
    : `<label class="full">Development area<input id="fGoal" value="${esc(x.goal || "")}"${dis}></label><label>Start date<input type="date" id="fStart" value="${esc(x.start || "")}"${dis}></label><label>Due date<input type="date" id="fDue" value="${esc(x.due || "")}"${dis}></label><label>Status<select id="fStatus"${dis}>${opts(PDP_STATUS, x.status)}</select></label><label class="full">Actions and support<textarea id="fNotes" rows="2"${dis}>${esc(x.notes || "")}</textarea></label>`;
  const files = `<div class="full"><div class="note" style="margin-bottom:6px">${kind === "Review" ? "Completed review" : "Development plan"} (PDF or photo)</div>${fileList(x.files, "p")}${ed ? '<input type="file" id="fFile" accept="application/pdf,image/*" style="margin-top:8px">' : ""}</div>`;
  modal(kind === "Review" ? (p ? "Completed review" : "Upload completed review") : (p ? "Personal development plan" : "Add personal development plan"), "", `<div class="form">${top}${mid}${files}</div>`,
    (ed ? (p ? '<button class="danger" id="mDel">Delete</button>' : "") + '<button class="primary" id="mSave">Save</button>' : "") + '<button id="mClose">Close</button>');
  $("mClose").onclick = closeModal; wireOpen(x);
  if (!ed) return;
  $("mSave").onclick = async () => {
    const rec = Object.assign({}, x, { empId: $("fEmp").value, notes: $("fNotes").value.trim() });
    if (kind === "Review") Object.assign(rec, { period: $("fPeriod").value.trim(), date: $("fDate").value, reviewer: $("fRev").value.trim() });
    else Object.assign(rec, { goal: $("fGoal").value.trim(), start: $("fStart").value, due: $("fDue").value, status: $("fStatus").value });
    if (!rec.empId) { toast("Choose the employee"); return; } rec.empName = empName(rec.empId);
    if (kind === "PDP" && (!rec.start || !rec.due)) { toast("Enter the start and due dates"); return; }
    const f = $("fFile").files[0]; if (kind === "Review" && !f && !(rec.files || []).length) { toast("Choose the completed review to upload"); return; }
    const b = $("mSave"); b.disabled = true; b.textContent = "Saving…";
    try { if (f) rec.files = (rec.files || []).concat([await putFile(f)]); await save("perf", rec); toast("Saved"); closeModal(); } catch (er) { fail(er); b.disabled = false; b.textContent = "Save"; }
  };
  if (p) $("mDel").onclick = async () => { if (!confirm("Delete this record and its files?")) return; try { await remove("perf", p); toast("Deleted"); closeModal(); } catch (er) { fail(er); } };
}

/* ================= Employment equity ================= */
function eeTable(list) {
  const cols = [["Male", "African"], ["Male", "Coloured"], ["Male", "Indian"], ["Male", "White"], ["Female", "African"], ["Female", "Coloured"], ["Female", "Indian"], ["Female", "White"]];
  const cnt = (lv, g, r) => list.filter(e => e.level === lv && e.foreign !== "Yes" && e.gender === g && e.race === r).length;
  const fn = (lv, g) => list.filter(e => e.level === lv && e.foreign === "Yes" && e.gender === g).length;
  const rows = LEVELS.map(lv => { const r = cols.map(([g, ra]) => cnt(lv, g, ra)).concat([fn(lv, "Male"), fn(lv, "Female")]); return [lv].concat(r, [r.reduce((a, b) => a + b, 0)]); });
  const tot = ["Total"].concat(rows[0].slice(1).map((_, i) => rows.reduce((a, r) => a + r[i + 1], 0)));
  const unset = list.filter(e => !LEVELS.includes(e.level) || !e.gender || (e.foreign !== "Yes" && !e.race)).length;
  return { head: ["Occupational level", "Male A", "Male C", "Male I", "Male W", "Female A", "Female C", "Female I", "Female W", "Foreign male", "Foreign female", "Total"], rows, tot, unset };
}
function viewEe() {
  const act = D.employees.filter(e => e.status !== "Terminated");
  const all = eeTable(act), dis = eeTable(act.filter(e => e.disability === "Yes"));
  const tbl = (t, title) => `<div class="card"><h2>${title}</h2><div class="tbl ee"><table><thead><tr>${t.head.map(h => `<th>${h}</th>`).join("")}</tr></thead><tbody>${t.rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}<tr class="tot">${t.tot.map(c => `<td>${c}</td>`).join("")}</tr></tbody></table></div></div>`;
  $("main").innerHTML = '<div class="bar"><button id="eeXls">Download workforce profile</button></div>' + tbl(all, "Workforce profile – all employees") + tbl(dis, "Workforce profile – employees with disabilities") +
    (all.unset ? `<p class="note">${all.unset} active employee${all.unset === 1 ? " is" : "s are"} not counted because their occupational level, gender or race is not filled in on the Employees tab.</p>` : "") +
    '<p class="note">A = African, C = Coloured, I = Indian, W = White. Active employees only.</p>';
  $("eeXls").onclick = () => downloadXlsx(`TKR Workforce profile ${today()}.xlsx`, [
    ["All employees", [["Workforce profile – all employees"], [], all.head].concat(all.rows, [all.tot]), [60, 8, 8, 8, 8, 9, 9, 9, 9, 12, 13, 8]],
    ["Disabled employees", [["Workforce profile – employees with disabilities"], [], dis.head].concat(dis.rows, [dis.tot]), [60, 8, 8, 8, 8, 9, 9, 9, 9, 12, 13, 8]]]);
}

/* ================= Training ================= */
let trKind = "interns";
const TR = {
  interns: { title: "Interns", one: "intern", fields: [["name", "Full name"], ["idNumber", "ID number"], ["phone", "Phone"], ["institution", "Institution"], ["qualification", "Qualification / programme"], ["department", "Department"], ["mentor", "Mentor"], ["start", "Start date", "date"], ["end", "End date", "date"]],
    cols: [["name", "Intern"], ["qualification", "Programme"], ["department", "Department"], ["mentor", "Mentor"], ["start", "Start", "date"], ["end", "End", "date"]], files: "Job cards" },
  bursaries: { title: "Bursaries", one: "student", fields: [["name", "Student name"], ["idNumber", "ID number"], ["phone", "Phone"], ["email", "Email"], ["institution", "Institution"], ["course", "Course"], ["year", "Year of study"], ["start", "Bursary start", "date"], ["end", "Bursary end", "date"], ["notes", "Notes"]],
    cols: [["name", "Student"], ["institution", "Institution"], ["course", "Course"], ["year", "Year"], ["start", "Start", "date"], ["end", "End", "date"]], files: "Documents" },
};
function viewTrain() {
  const ed = canEdit(), T = TR[trKind], l = D[trKind].slice().sort((a, b) => (a.name || "").localeCompare(b.name || ""));
  let h = `<div class="bar"><div class="tabs" style="margin:0;border:0">${Object.keys(TR).map(k => `<button data-tk="${k}" aria-selected="${trKind === k}">${TR[k].title}</button>`).join("")}</div>` +
    (ed ? `<button class="primary" id="tAdd">Add ${T.one}</button>` : "") + '<button id="tXls">Download Excel</button></div>';
  h += l.length ? `<div class="tbl"><table><thead><tr>${T.cols.map(c => `<th>${c[1]}</th>`).join("")}<th>${T.files}</th><th></th></tr></thead><tbody>` +
    l.map(r => `<tr>${T.cols.map(([k, , t], i) => `<td>${i ? "" : "<b>"}${t === "date" ? fmtD(r[k]) : esc(r[k] || "")}${i ? "" : "</b>"}</td>`).join("")}<td>${(r.files || []).length}</td><td><button class="sm" data-topen="${esc(r.id)}">${ed ? "Edit" : "View"}</button></td></tr>`).join("") + '</tbody></table></div>'
    : `<div class="card note">No ${T.title.toLowerCase()} yet.</div>`;
  $("main").innerHTML = h;
  document.querySelectorAll("[data-tk]").forEach(b => b.onclick = () => { trKind = b.dataset.tk; viewTrain(); });
  document.querySelectorAll("[data-topen]").forEach(b => b.onclick = () => trForm(D[trKind].find(r => r.id === b.dataset.topen)));
  if (ed) $("tAdd").onclick = () => trForm(null);
  $("tXls").onclick = () => downloadXlsx(`TKR HR ${T.title} ${today()}.xlsx`, [[T.title, [[T.title], [], T.fields.map(f => f[1]).concat([T.files])].concat(l.map(r => T.fields.map(([k, , t]) => t === "date" ? fmtD(r[k]) : (r[k] || "")).concat([(r.files || []).map(f => f.name).join(", ")]))), T.fields.map(f => Math.max(14, f[1].length + 2)).concat([30])]]);
}
function trForm(r) {
  const ed = canEdit(), T = TR[trKind], x = r || {}, dis = ed ? "" : " disabled";
  const body = '<div class="form">' + T.fields.map(([k, l, t]) => `<label${k === "notes" ? ' class="full"' : ""}>${l}<input data-k="${k}" ${t === "date" ? 'type="date"' : ""} value="${esc(x[k] || "")}"${dis}></label>`).join("") +
    `<div class="full"><div class="note" style="margin-bottom:6px">${T.files} (PDF or photo)</div>${fileList(x.files, "t")}${ed ? '<input type="file" id="fFile" accept="application/pdf,image/*" multiple style="margin-top:8px">' : ""}</div></div>`;
  modal(r ? r.name : "Add " + T.one, "", body, (ed ? (r ? '<button class="danger" id="mDel">Delete</button>' : "") + '<button class="primary" id="mSave">Save</button>' : "") + '<button id="mClose">Close</button>');
  $("mClose").onclick = closeModal; wireOpen(x);
  if (!ed) return;
  $("mSave").onclick = async () => {
    const rec = Object.assign({}, x); document.querySelectorAll("#modal [data-k]").forEach(i => rec[i.dataset.k] = i.value.trim());
    if (!rec.name) { toast("Enter the name"); return; }
    const b = $("mSave"); b.disabled = true; b.textContent = "Saving…";
    try { for (const f of $("fFile").files) rec.files = (rec.files || []).concat([await putFile(f)]); await save(trKind, rec); toast("Saved"); closeModal(); } catch (er) { fail(er); b.disabled = false; b.textContent = "Save"; }
  };
  if (r) $("mDel").onclick = async () => { if (!confirm("Delete " + r.name + " and their files?")) return; try { await remove(trKind, r); toast("Deleted"); closeModal(); } catch (er) { fail(er); } };
}

/* ================= Access (HR admin only) ================= */
function viewAccess() {
  const row = (em, role) => `<tr><td>${esc(em)}</td><td>${role === "editors" ? "Full HR access" : "View only"}</td><td><button class="sm danger" data-rm="${esc(em)}" data-role="${role}">Remove</button></td></tr>`;
  const n = access.editors.length + access.viewers.length;
  $("main").innerHTML = `<div class="card"><h2>Give HR access</h2><form class="form" id="addUser"><label>Email or username<input id="uEmail" required></label><label>Password<input id="uPass" type="text" minlength="6" placeholder="Leave empty if they already sign in to the TKR app"></label>
    <label>Access<select id="uRole"><option value="editors">Full HR access (add and change)</option><option value="viewers">View only</option></select></label><div style="display:flex;align-items:end"><button class="primary">Add</button></div></form>
    <p class="note" id="uMsg" role="status">Only you (${HR_ADMINS.map(esc).join(", ")}) can give or remove HR access.</p></div>
    <div class="card"><h2>Who has HR access <span class="note">${n}</span></h2>${n ? `<div class="tbl"><table><thead><tr><th>Email</th><th>Access</th><th></th></tr></thead><tbody>${access.editors.map(e => row(e, "editors")).join("")}${access.viewers.map(e => row(e, "viewers")).join("")}</tbody></table></div>` : '<p class="note">Nobody else has HR access yet.</p>'}</div>`;
  $("addUser").onsubmit = async ev => {
    ev.preventDefault(); const em = toLogin($("uEmail").value).toLowerCase(), pw = $("uPass").value, role = $("uRole").value, other = role === "editors" ? "viewers" : "editors", msg = $("uMsg");
    if (pw && pw.length < 6) { msg.textContent = "The password must be at least 6 characters."; return; }
    msg.textContent = "Adding…"; let note = "";
    if (pw) { const tmp = initializeApp(firebaseConfig, "hr-add-" + Date.now());
      try { await createUserWithEmailAndPassword(getAuth(tmp), em, pw); note = " Their sign-in was made."; }
      catch (er) { if (er?.code === "auth/email-already-in-use") note = " They already had a sign-in, so their password stays the same."; else { msg.textContent = "The sign-in couldn't be made. Check the email and try again."; deleteApp(tmp).catch(() => {}); return; } }
      deleteApp(tmp).catch(() => {}); }
    try { await setDoc(doc(db, "config", "hrAccess"), { [role]: arrayUnion(em), [other]: arrayRemove(em) }, { merge: true }); toast(em + " now has HR access." + note); } catch (er) { fail(er); msg.textContent = ""; }
  };
  document.querySelectorAll("[data-rm]").forEach(b => b.onclick = async () => { if (!confirm("Remove HR access for " + b.dataset.rm + "?")) return; try { await setDoc(doc(db, "config", "hrAccess"), { [b.dataset.role]: arrayRemove(b.dataset.rm) }, { merge: true }); toast("Access removed"); } catch (er) { fail(er); } });
}

/* ================= shell ================= */
const VIEWS = { dash: viewDash, emps: viewEmps, er: viewEr, perf: viewPerf, ee: viewEe, train: viewTrain, access: viewAccess };
function render() {
  if (view === "access" && !isHrAdmin()) view = "dash";
  document.querySelectorAll("#tabs button").forEach(b => b.setAttribute("aria-selected", b.dataset.v === view));
  $("accessTab").hidden = !isHrAdmin();
  if (!loaded.employees) { $("main").innerHTML = '<div class="card note">Loading HR records…</div>'; return; }
  VIEWS[view]();
}
$("tabs").onclick = e => { const b = e.target.closest("button"); if (!b) return; view = b.dataset.v; try { localStorage.setItem("tkr-hr-view", view); } catch {} render(); };
try { const v = localStorage.getItem("tkr-hr-view"); if (v && VIEWS[v]) view = v; } catch {}

function show(which) { ["signin", "noaccess", "app"].forEach(id => $(id).hidden = id !== which); $("outBtn").hidden = which === "signin"; }
let accUnsub = null;
function startData() {
  unsubs.forEach(u => u()); unsubs = [];
  Object.keys(D).forEach(name => unsubs.push(onSnapshot(col(name), s => {
    D[name] = s.docs.map(d => Object.assign({ id: d.id }, d.data())); loaded[name] = true;
    if (!$("modal").innerHTML) render();
  }, e => { console.warn(name, e); if (name === "employees") { $("main").innerHTML = '<div class="card note">Could not load HR records: ' + esc(e.code === "permission-denied" ? "the HR access rules are not set up yet. Ask Muneer to publish them in Firebase." : e.message) + '</div>'; } })));
}
onAuthStateChanged(auth, u => {
  if (accUnsub) { accUnsub(); accUnsub = null; } unsubs.forEach(x => x()); unsubs = [];
  if (!u) { email = ""; $("who").textContent = ""; show("signin"); return; }
  email = (u.email || "").toLowerCase(); $("who").textContent = email.replace("@mgr.tkroperations.app", "");
  accUnsub = onSnapshot(doc(db, "config", "hrAccess"), d => {
    const v = d.exists() ? d.data() : {}; access = { editors: (v.editors || []).map(x => x.toLowerCase()), viewers: (v.viewers || []).map(x => x.toLowerCase()) };
    if (!canView()) { $("naEmail").textContent = email; unsubs.forEach(x => x()); unsubs = []; show("noaccess"); return; }
    if (!unsubs.length) startData(); show("app"); render();
  }, () => { access = { editors: [], viewers: [] }; if (isHrAdmin()) { if (!unsubs.length) startData(); show("app"); render(); } else { $("naEmail").textContent = email; show("noaccess"); } });
});
$("signForm").onsubmit = async e => {
  e.preventDefault(); const m = $("sMsg"); m.textContent = "Signing in…"; $("sBtn").disabled = true;
  try { await signInWithEmailAndPassword(auth, toLogin($("sEmail").value), $("sPass").value); m.textContent = ""; }
  catch { m.textContent = "Wrong email or password."; } finally { $("sBtn").disabled = false; }
};
$("sReset").onclick = async () => { const em = toLogin($("sEmail").value), m = $("sMsg"); if (!em.includes("@") || em.endsWith("@mgr.tkroperations.app")) { m.textContent = "Type your email first, then press Forgot password."; return; }
  try { await sendPasswordResetEmail(auth, em); m.textContent = "A password reset email was sent to " + em + "."; } catch { m.textContent = "The reset email couldn't be sent. Ask Amogelang."; } };
$("outBtn").onclick = $("naOut").onclick = () => signOut(auth);
document.addEventListener("keydown", e => { if (e.key === "Escape") closeModal(); });
if ("serviceWorker" in navigator) navigator.serviceWorker.register("../sw.js", { updateViaCache: "none" }).catch(() => {});
