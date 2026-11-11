/* The TKR drilling metres workbook (same tabs as "TKR Drilling Metres and Tally Sheets"):
   Summary · TKR vs Tharisa · Shifts and evidence · Daily metres (with tally sheets) · Machine performance ·
   Delay summary (with charts) · Supporting documents · Comments.
   Input:
     slots  Map "yyyy-mm-dd|Day|TDR 013" → { metres, status, supervisor, block, hrs, delays:[{text,cat,hours}], delayText, report }
     files  [{ type, date, shift, rig, name, note, pdf, link, img:{data,w,h} }]  (img = a lighter JPEG data URL, or null) */
const pad = n => String(n).padStart(2, "0");
const parse = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return ymd(d); };
const fmt = s => s ? s.split("-").reverse().join("/") : "";
const dow = s => parse(s).toLocaleDateString("en-ZA", { weekday: "short" });
const r1 = v => Math.round((v || 0) * 10) / 10;
const M1 = v => { const [i, d] = r1(v).toFixed(1).split("."); return i.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "." + d; };
export const daysOf = (from, to) => { const a = []; for (let d = from; d <= to; d = addDays(d, 1)) a.push(d); return a; };

export function bdData(slots, from, to, RIGS) {
  const all = [...slots.values()].filter(s => s.date >= from && s.date <= to).sort((a, b) => a.date.localeCompare(b.date) || a.shift.localeCompare(b.shift) || a.rig.localeCompare(b.rig));
  const rigs = RIGS.filter(r => all.some(s => s.rig === r));
  const mach = rigs.map(r => { const L = all.filter(s => s.rig === r); return { rig: r, working: L.filter(s => s.status === "Working").length, down: L.filter(s => s.status === "Not working").length, metres: r1(L.reduce((n, s) => n + s.metres, 0)),
    hours: Math.round(L.flatMap(s => s.delays).reduce((n, d) => n + (d.hours || 0), 0) * 10) / 10 }; });
  const lines = all.flatMap(s => s.delays.map(d => ({ ...d, date: s.date, shift: s.shift, rig: s.rig })));
  const cats = [...new Set(lines.map(l => l.cat))].map(c => { const L = lines.filter(l => l.cat === c), per = rigs.map(r => Math.round(L.filter(l => l.rig === r).reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10);
    return { cat: c, per, times: rigs.map(r => L.filter(l => l.rig === r).length), total: Math.round(L.reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10 }; }).sort((a, b) => b.total - a.total);
  return { all, rigs, mach, lines, cats, total: Math.round(lines.reduce((n, l) => n + (l.hours || 0), 0) * 10) / 10 };
}

function chartPng(title, labels, series, opt = {}) {
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

export async function buildMetresWorkbook(ExcelJS, { from, to, rigs: RIGS, slots: S, files = [], tharisa = 0, only = false, logoBuf = null }) {
  const wb = new ExcelJS.Workbook(); wb.creator = "TKR Operations"; wb.created = new Date();
  let logo = null; if (logoBuf) try { logo = wb.addImage({ buffer: logoBuf, extension: "png" }); } catch {}
  const X = { BLACK: "FF111111", HEAD: "FF262626", GOLD: "FFD4A12A", CREAM: "FFFFF8E6", TOT: "FFFFF2CC", WHITE: "FFFFFFFF", GRID: "FFD9D3C5", MUTED: "FF6B6558", RED: "FF9C0006", REDF: "FFFDE2DF", GREEN: "FF375623", GREENF: "FFE2F0D9" };
  const fill = x => ({ type: "pattern", pattern: "solid", fgColor: { argb: x } }), thin = { style: "thin", color: { argb: X.GRID } }, box = { top: thin, left: thin, bottom: thin, right: thin };
  const F = (o = {}) => Object.assign({ name: "Arial", size: 10, color: { argb: X.BLACK } }, o);
  const D = s => { const [y, m, d] = s.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); }, DEC1 = "#,##0.0", INT = "0";
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

  const days = daysOf(from, to), B = bdData(S, from, to, RIGS), all = B.all;
  const rigs = RIGS.filter(r => all.some(s => s.rig === r)); if (!rigs.length) rigs.push(...RIGS.slice(0, 2));
  const sl = (d, sh, r) => S.get(d + "|" + sh + "|" + r), dm = (d, r) => r1(["Day", "Night"].reduce((n, sh) => n + (sl(d, sh, r)?.metres || 0), 0));
  const tot = r => r1(all.filter(s => !r || s.rig === r).reduce((n, s) => n + s.metres, 0)), grand = tot();
  const tally = files.filter(f => f.type === "Tally sheet");
  const img = (ws, f, col, row, bw, bh) => { const k = Math.min(bw / f.img.w, bh / f.img.h), id = wb.addImage({ base64: f.img.data, extension: "jpeg" });
    ws.addImage(id, { tl: { col: col + 0.04, row: row + 0.04 }, ext: { width: Math.round(f.img.w * k), height: Math.round(f.img.h * k) }, editAs: "oneCell" }); };
  let ws, r, i = 0;

  if (!only) {
    ws = sheet("Summary", [24, 16, 16, 18, 22, 14], "DRILLING SUMMARY", "Tharisa Mine · " + rigs.join(", ")); r = 7;
    ws.getCell(r, 1).value = "Total metres drilled"; ws.getCell(r, 1).font = F({ size: 12, bold: true }); ws.getCell(r, 2).value = grand; ws.getCell(r, 2).numFmt = DEC1; ws.getCell(r, 2).font = F({ size: 16, bold: true, color: { argb: "FF8A6A00" } }); ws.getRow(r).height = 26; r += 2;
    r = head(ws, r, ["Machine", "Metres", "Shifts working", "Shifts not working", "Avg m per working shift"]);
    B.mach.forEach((m, k) => { r = line(ws, r, [m.rig, m.metres, m.working, m.down, m.working ? r1(m.metres / m.working) : 0], k, { fmt: { 2: INT, 3: INT } }); });
    r = line(ws, r, ["Total", grand, "", "", ""], 0, { tot: true }); r++;
    const exp = days.length * 2, cap = days.reduce((n, d) => n + ["Day", "Night"].filter(sh => rigs.some(x => sl(d, sh, x))).length, 0);
    r = line(ws, r, ["Shifts captured", cap, `of ${exp} shifts`, "", ""], 0, { fmt: { 1: INT } }); r++;
    ws.getCell(r, 1).value = "Sheets in this file:"; ws.getCell(r, 1).font = F({ bold: true }); r++;
    for (const t of ["Shifts and evidence – every shift with metres, status, delays, the shift report and the documents filed for it.", "Daily metres – metres per rig per day with the tally sheets next to them.", "Machine performance – each machine per day, drilled or not, with machine hours and delays.",
      "Delay summary – time lost per machine and per cause, and every delay shift by shift.", "Supporting documents – job cards, checklists, breakdown reports and other documents filed.", "Comments – shifts with metres and no tally sheet, and notes on the documents."]) { ws.getCell(r, 1).value = t; ws.getCell(r, 1).font = F(); r++; }
    r++; const sp = stackPng(days, rigs, ["#d4a12a", "#111111", "#2e7d4f"], dm); ws.addImage(wb.addImage({ base64: sp.png, extension: "png" }), { tl: { col: 0, row: r }, ext: { width: 900, height: Math.round(900 * sp.h / sp.w) } });

    const th = +tharisa || 0;
    if (th > 0) { ws = sheet("TKR vs Tharisa", [32, 14, 22, 18, 10], "TKR COUNT VS THARISA COUNT", "Tharisa's metres for these dates against the TKR shift register"); r = 7;
      r = head(ws, r, ["Count", "Metres", "Difference to Tharisa", "Difference %"]);
      r = line(ws, r, ["Tharisa", r1(th), "", ""], 0); r = line(ws, r, ["TKR shift register", grand, r1(grand - th), Math.round((grand - th) / th * 1000) / 10 + "%"], 1); }

    ws = sheet("Shifts and evidence", [11, 7, 12, ...rigs.flatMap(() => [9, 12, 34]), 40, 34], "MACHINE PERFORMANCE PER SHIFT", "Every day and night shift: metres, status, delays, shift report and the documents filed"); r = 7;
    r = head(ws, r, ["Date", "Shift", "Supervisor", ...rigs.flatMap(x => [x + " m", x, x + " delays"]), "Shift report", "Tally sheets and documents filed"]); i = 0;
    for (const d of days) for (const sh of ["Day", "Night"]) { const L = rigs.map(x => sl(d, sh, x)); if (!L.some(Boolean)) continue;
      const rep = L.map(s => s?.report).find(Boolean) || "";
      const fl = files.filter(f => f.date === d && f.shift === sh).map(f => `${f.type} · ${f.rig || "All machines"}`).join("\n"), n = Math.max(...L.map(s => rowsFor(s?.delayText, 44)), rowsFor(rep, 44), fl ? fl.split("\n").length : 1);
      r = line(ws, r, [D(d), sh, L.map(s => s?.supervisor).find(Boolean) || "", ...L.flatMap(s => s ? [s.metres, s.status, s.delayText] : ["", "", ""]), rep, fl], i++,
        { h: Math.min(400, Math.max(18, n * 13 + 4)), color: Object.fromEntries(L.map((s, k) => [4 + k * 3, s?.status === "Working" ? "g" : s?.status === "Not working" ? "r" : ""])) }); }
    r = line(ws, r, ["Total", "", "", ...rigs.flatMap(x => [tot(x), "", ""]), "", "All rigs: " + M1(grand) + " m"], 0, { tot: true });

    ws = sheet("Daily metres", [12, 6, 10, 10, 30, 34, 34, 34, 34], "DRILLING METRES WITH TALLY SHEETS", "Metres per day per rig with the signed tally sheets"); r = 7;
    r = head(ws, r, ["Date", "Day", "Machine", "Metres", "Tally sheets attached", "Tally sheet 1", "Tally sheet 2", "Tally sheet 3", "Tally sheet 4"]); i = 0;
    for (const d of days) for (const rg of rigs) { if (!sl(d, "Day", rg) && !sl(d, "Night", rg)) continue; const T = tally.filter(f => f.date === d && f.rig === rg), m = dm(d, rg);
      const label = T.length ? T.map(f => `${f.shift || "Whole day"} shift${f.note ? " – " + f.note : ""}`).join("\n") : m > 0 ? "No tally sheet filed" : "No drilling";
      const pics = T.filter(f => f.img).slice(0, 4); line(ws, r, [D(d), dow(d), rg, m, label, "", "", "", ""], i++, { h: pics.length ? 134 : 28, top: true });
      pics.forEach((f, k) => img(ws, f, 5 + k, r - 1, 232, 172)); r++; }
    r = line(ws, r, ["Total", "", "", grand, `All rigs, ${fmt(from)} to ${fmt(to)}`, "", "", "", ""], 0, { tot: true });
    text(ws, r, "Click a tally sheet and drag a corner to read it at full size. The full photos are in the documents ZIP.");

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
    const sup = files.filter(f => f.type !== "Tally sheet");
    ws = sheet("Supporting documents", [34, 34, 34, 34, 20], "SUPPORTING DOCUMENTS", "Job cards, checklists, breakdown reports and other documents filed for these dates"); r = 7;
    r = head(ws, r, ["Document", "Date", "Machine and shift", "File", "Note"]);
    sup.forEach((f, k) => { r = line(ws, r, [f.type, D(f.date), `${f.rig || "All machines"}${f.shift ? " · " + f.shift : ""}`, f.link ? { text: f.name || "Open the file", hyperlink: f.link } : f.name || (f.pdf ? "PDF" : "Photo"), f.note], k); });
    if (!sup.length) r = text(ws, r, "No job cards, checklists, breakdown reports or other documents filed for these dates."); r++;
    const pics = sup.filter(f => f.img);
    for (let k = 0; k < pics.length; k += 4) { const row = pics.slice(k, k + 4); ws.getRow(r).height = 18; row.forEach((f, j) => { const x = ws.getCell(r, j + 1); x.value = `${f.type} · ${fmt(f.date)} · ${f.rig || "All machines"} ${f.shift || ""}`; x.font = F({ bold: true, size: 9 }); }); r++;
      ws.getRow(r).height = 250; row.forEach((f, j) => img(ws, f, j, r - 1, 232, 326)); r += 2; }

    ws = sheet("Comments", [12, 11, 10, 10, 80], "COMMENTS", "Shifts with metres and no tally sheet, and notes on the documents"); r = 7;
    r = head(ws, r, ["Date", "Machine", "Shift", "Metres", "Comments"]); i = 0;
    const com = [];
    for (const s of all) if (s.metres > 0 && !tally.some(f => f.date === s.date && f.rig === s.rig && f.shift === s.shift)) com.push([s.date, s.rig, s.shift, s.metres, `${s.shift} shift: ${M1(s.metres)} m as per the shift register; no tally sheet filed for this shift.`]);
    for (const f of files) if (f.note) com.push([f.date, f.rig || "All machines", f.shift || "", "", `${f.type}: ${f.note}`]);
    com.sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || String(a[2]).localeCompare(String(b[2])));
    com.forEach(c => { r = line(ws, r, [D(c[0]), c[1], c[2], c[3], c[4]], i++, { h: Math.max(18, rowsFor(c[4], 95) * 13 + 4) }); });
    if (!com.length) text(ws, r, "Every shift with metres has a tally sheet. No notes.");
  }
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
