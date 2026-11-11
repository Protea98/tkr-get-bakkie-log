/* Metres chart shared by the Metres page and the GET Forecast dashboard: Line / Bar / Waterfall,
   per day / week / month / year (D1 W1 MN Y1, like the timeframes on a trading chart). One set of data: fcMetres. */
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
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthName = m => MONTHS[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
const shiftMonth = (m, n) => ymd(new Date(+m.slice(0, 4), +m.slice(5, 7) - 1 + n, 1)).slice(0, 7);
const num = v => { const n = parseFloat(String(v ?? "").replace(/\s/g, "").replace(",", ".")); return isFinite(n) ? n : 0; };
/* ---------- metres chart: Day / Week / Month / Year, like the timeframes on a trading chart ---------- */
const TF = [["D", "Day"], ["W", "Week"], ["M", "Month"], ["Y", "Year"]];
const SER_COL = ["#efc75e", "#6fb3ff", "#8fd1a6", "#ec8f7b", "#c39bff", "#ffb86b"];
export const ch = (() => { try { return { tf: "D", off: 0, hide: [], ...JSON.parse(localStorage.getItem("tkr-metres-chart") || "{}") }; } catch { return { tf: "D", off: 0, hide: [] }; } })();
export const saveCh = () => { try { localStorage.setItem("tkr-metres-chart", JSON.stringify({ tf: ch.tf, hide: ch.hide, kind: ch.kind })); } catch {} };
const monday = d => { const x = parse(d), w = (x.getDay() + 6) % 7; x.setDate(x.getDate() - w); return ymd(x); };
function chartBuckets(E, R) {
  const t = today(), keyOf = { D: d => d, W: monday, M: d => d.slice(0, 7), Y: d => d.slice(0, 4) }[ch.tf], keys = [];
  if (ch.tf === "D") { const end = addDays(t, -ch.off * 31); for (let i = 30; i >= 0; i--) keys.push(addDays(end, -i)); }
  if (ch.tf === "W") { const end = addDays(monday(t), -ch.off * 16 * 7); for (let i = 15; i >= 0; i--) keys.push(addDays(end, -i * 7)); }
  if (ch.tf === "M") { const end = shiftMonth(t.slice(0, 7), -ch.off * 12); for (let i = 11; i >= 0; i--) keys.push(shiftMonth(end, -i)); }
  if (ch.tf === "Y") { const first = Math.min(+t.slice(0, 4), ...E.map(e => +String(e.date).slice(0, 4)).filter(Boolean)); for (let y = first; y <= +t.slice(0, 4); y++) keys.push(String(y)); }
  const val = new Map(keys.map(k => [k, Object.fromEntries(R.map(r => [r, 0]))]));
  for (const e of E) { if (!e.date) continue; const k = keyOf(e.date), v = val.get(k); if (!v || !(e.rig in v)) continue; v[e.rig] += num(e.metres); }
  const label = k => ch.tf === "D" ? `${k.slice(8)}/${k.slice(5, 7)}` : ch.tf === "W" ? `wk ${k.slice(8)}/${k.slice(5, 7)}` : ch.tf === "M" ? MONTHS[+k.slice(5, 7) - 1].slice(0, 3) + " " + k.slice(2, 4) : k;
  const title = k => ch.tf === "D" ? `${dow(k)} ${fmt(k)}` : ch.tf === "W" ? `Week of ${fmt(k)} to ${fmt(addDays(k, 6))}` : ch.tf === "M" ? monthName(k + "-01") : k;
  return { keys, R, rows: keys.map(k => ({ k, l: label(k), t: title(k), v: val.get(k), tot: r1(R.reduce((n, r) => n + val.get(k)[r], 0)) })) };
}
/* E: [{date, rig, metres}], R: the rigs to show, in order */
export function metresChart(E, R, title = "Metres drilled") {
  const { rows } = chartBuckets(E, R), series = [["Total", rows.map(x => x.tot)], ...R.map(r => [r, rows.map(x => r1(x.v[r]))])].map(([n, v], i) => ({ n, v, c: SER_COL[i % SER_COL.length], off: ch.hide.includes(n) }));
  const kind = ch.kind || "line", on = series.filter(s => !s.off), rigOn2 = on.filter(s => s.n !== "Total"), stackS = rigOn2.length ? rigOn2 : on;
  const stackOf = i => r1(stackS.reduce((a, s) => a + s.v[i], 0)), n = rows.length, slots = kind === "wf" ? n + 1 : n;
  const cum = []; rows.reduce((a, _, i) => (cum[i] = r1(a + stackOf(i))), 0); const final = n ? cum[n - 1] : 0;
  const max = Math.max(10, ...(kind === "line" ? on.flatMap(s => s.v) : kind === "bar" ? rows.map((_, i) => stackOf(i)) : [final])) * (kind === "line" ? 1 : 1.1);
  const W = 1000, H = 300, l = 52, rr = 14, top = 14, bot = 34, sw = (W - l - rr) / Math.max(1, slots);
  const x = i => l + sw * (i + 0.5), y = v => top + (H - top - bot) * (1 - v / max), bw = Math.max(2, Math.min(46, sw * 0.68));
  const step = Math.pow(10, Math.floor(Math.log10(max / 4))), tick = Math.ceil(max / 4 / step) * step;
  let g = ""; for (let v = 0; v <= max; v += tick) g += `<line x1="${l}" x2="${W - rr}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)"/><text x="${l - 8}" y="${y(v) + 4}" text-anchor="end">${Math.round(v)}</text>`;
  const every = Math.ceil(n / (ch.tf === "D" ? 11 : 12));
  rows.forEach((r, i) => { if (i % every === 0 || i === n - 1) g += `<text x="${x(i)}" y="${H - 12}" text-anchor="middle">${r.l}</text>`; });
  if (kind === "line") {
    for (const s of on) { g += `<polyline fill="none" stroke="${s.c}" stroke-width="${s.n === "Total" ? 3 : 2}" stroke-linejoin="round" points="${s.v.map((v, i) => `${x(i)},${y(v)}`).join(" ")}"/>`;
      if (n <= 40) g += s.v.map((v, i) => `<circle cx="${x(i)}" cy="${y(v)}" r="${s.n === "Total" ? 3.5 : 2.5}" fill="${s.c}"/>`).join(""); }
  } else if (kind === "bar") {
    rows.forEach((r, i) => { let base = 0; for (const s of stackS) { const v = s.v[i]; if (!v) continue; g += `<rect x="${x(i) - bw / 2}" y="${y(base + v)}" width="${bw}" height="${y(base) - y(base + v)}" fill="${s.c}" rx="2"/>`; base += v; }
      if (base && n <= 31) g += `<text x="${x(i)}" y="${y(base) - 5}" text-anchor="middle" style="fill:var(--fg);font-size:11px">${Math.round(base)}</text>`; });
  } else {            // waterfall: each period adds on top of the running total; the last bar is the total
    rows.forEach((r, i) => { let base = i ? cum[i - 1] : 0; const start = base;
      for (const s of stackS) { const v = s.v[i]; if (!v) continue; g += `<rect x="${x(i) - bw / 2}" y="${y(base + v)}" width="${bw}" height="${y(base) - y(base + v)}" fill="${s.c}" opacity=".9"/>`; base += v; }
      if (i < n - 1) g += `<line x1="${x(i) + bw / 2}" x2="${x(i + 1) - bw / 2}" y1="${y(cum[i])}" y2="${y(cum[i])}" stroke="var(--muted)" stroke-dasharray="3 3"/>`;
      if (base > start && n <= 16) g += `<text x="${x(i)}" y="${y(base) - 5}" text-anchor="middle" style="fill:var(--fg);font-size:11px">+${Math.round(base - start)}</text>`; });
    g += `<line x1="${x(n - 1) + bw / 2}" x2="${x(n) - bw / 2}" y1="${y(final)}" y2="${y(final)}" stroke="var(--muted)" stroke-dasharray="3 3"/>`;
    g += `<rect x="${x(n) - bw / 2}" y="${y(final)}" width="${bw}" height="${y(0) - y(final)}" fill="#efc75e" rx="2"/><text x="${x(n)}" y="${y(final) - 6}" text-anchor="middle" style="fill:var(--gold);font-size:12px;font-weight:700">${M1(final)}</text><text x="${x(n)}" y="${H - 12}" text-anchor="middle" style="fill:var(--gold)">Total</text>`;
  }
  rows.forEach((r, i) => { g += `<rect class="hov" x="${x(i) - sw / 2}" y="${top}" width="${sw}" height="${H - top - bot}" fill="transparent"><title>${esc(r.t)}\nTotal: ${M1(r.tot)} m${R.map(k => `\n${k}: ${M1(r.v[k])} m`).join("")}${kind === "wf" ? `\nRunning total: ${M1(cum[i])} m` : ""}</title></rect>`; });
  const tot = r1(rows.reduce((a, r) => a + r.tot, 0)), withM = rows.filter(r => r.tot > 0), avg = withM.length ? r1(tot / withM.length) : 0, best = rows.reduce((b, r) => r.tot > (b?.tot || 0) ? r : b, null);
  const unit = { D: "day", W: "week", M: "month", Y: "year" }[ch.tf];
  return `<section class="card chartc"><div class="chead"><h2>${esc(title)}</h2>
      <div class="tfs" role="group" aria-label="Timeframe">${TF.map(([k, t]) => `<button data-tf="${k}" aria-pressed="${ch.tf === k}" title="Per ${t.toLowerCase()}">${k === "D" ? "D1" : k === "W" ? "W1" : k === "M" ? "MN" : "Y1"}<span>${t}</span></button>`).join("")}</div>
      <div class="tfs" role="group" aria-label="Chart type">${[["line", "Line"], ["bar", "Bar"], ["wf", "Waterfall"]].map(([k, t]) => `<button data-kind="${k}" aria-pressed="${(ch.kind || "line") === k}">${k === "line" ? "╱╲" : k === "bar" ? "▮▮" : "▁▃▆"}<span>${t}</span></button>`).join("")}</div>
      ${ch.tf !== "Y" ? `<div class="tfs"><button data-coff="1" aria-label="Earlier">‹</button><button data-coff="-1" aria-label="Later" ${ch.off ? "" : "disabled"}>›</button></div>` : ""}</div>
    <div class="cstats"><span>${ch.tf === "W" ? `${fmt(rows[0].k)} – ${fmt(addDays(rows[n - 1].k, 6))}` : ch.tf === "D" ? `${fmt(rows[0].k)} – ${fmt(rows[n - 1].k)}` : `${rows[0].t} – ${rows[n - 1].t}`}</span><span>Total <b>${M1(tot)} m</b></span><span>Average per ${unit} drilled <b>${M1(avg)} m</b></span>${best && best.tot ? `<span>Best ${unit} <b>${esc(best.t.replace(/^Week of /, "wk "))} · ${M1(best.tot)} m</b></span>` : ""}</div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Metres drilled per ${unit}">${g}</svg>
    <div class="legend">${series.map(s => `<button class="ser" data-ser="${esc(s.n)}" aria-pressed="${!s.off}"><i style="background:${s.c}"></i>${esc(s.n)}</button>`).join("")}</div></section>`;
}

/* the chart's buttons; returns true when it handled the tap */
export function metresChartClick(b) {
  if (!b || !b.closest(".chartc")) return false;
  if (b.dataset.kind) { ch.kind = b.dataset.kind; saveCh(); return true; }
  if (b.dataset.tf) { ch.tf = b.dataset.tf; ch.off = 0; saveCh(); return true; }
  if (b.dataset.coff) { ch.off = Math.max(0, ch.off + +b.dataset.coff); return true; }
  if (b.dataset.ser) { const n = b.dataset.ser; ch.hide = ch.hide.includes(n) ? ch.hide.filter(x => x !== n) : [...ch.hide, n]; saveCh(); return true; }
  return false;
}
