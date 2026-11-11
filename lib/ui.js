/* Shared screen pieces for the TKR app: buttons that open a section on its own screen, page headings, tab bar. */
let openId = "";
const e = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* A button (or a tile) that opens `inner` on its own screen.
   opt: { sub, count, icon, cls, tile: { l, n, s, cls } } */
export function sheet(id, title, inner, opt = {}) {
  const btn = opt.tile
    ? `<button type="button" class="tile tbtn ${opt.tile.cls || ""}" data-sheet="${id}"><div class="l">${opt.tile.l}</div><div class="n">${opt.tile.n}</div>${opt.tile.s ? `<div class="s">${opt.tile.s}</div>` : ""}<span class="go" aria-hidden="true">Open ›</span></button>`
    : `<button type="button" class="opener ${opt.cls || ""}" data-sheet="${id}"><span class="oi" aria-hidden="true">${opt.icon || "›"}</span><span class="ot"><b>${title}</b>${opt.sub ? `<small>${opt.sub}</small>` : ""}</span>${opt.count != null && opt.count !== "" ? `<span class="oc">${opt.count}</span>` : ""}<span class="oa" aria-hidden="true">›</span></button>`;
  return btn + `<div class="sheet" id="sheet-${id}" ${openId === id ? "" : "hidden"} role="dialog" aria-modal="true" aria-label="${e(title)}">
    <div class="sheet-bar"><button type="button" class="back" data-sheet-close aria-label="Back">‹ Back</button><b>${title}</b></div>
    <div class="sheet-body">${inner}</div></div>`;
}
/* Call after every render: keeps the open screen open, and stops the page behind it from scrolling. */
export function afterRender(root) {
  if (openId && !root.querySelector("#sheet-" + CSS.escape(openId))) openId = "";
  document.body.classList.toggle("noscroll", !!openId);
}
/* Call first in the click handler. Returns true when it handled the click. */
export function sheetClick(ev, root) {
  const o = ev.target.closest("[data-sheet]");
  if (o) { openId = o.dataset.sheet; const el = root.querySelector("#sheet-" + CSS.escape(openId)); if (el) { el.hidden = false; el.querySelector(".sheet-body").scrollTop = 0; } document.body.classList.add("noscroll"); return true; }
  if (ev.target.closest("[data-sheet-close]")) { closeSheet(root); return true; }
  return false;
}
export function closeSheet(root) { if (openId) { const el = root.querySelector("#sheet-" + CSS.escape(openId)); if (el) el.hidden = true; } openId = ""; document.body.classList.remove("noscroll"); }
export const isOpen = () => !!openId;
addEventListener("keydown", ev => { if (ev.key === "Escape" && openId) { const el = document.getElementById("sheet-" + openId); if (el) el.hidden = true; openId = ""; document.body.classList.remove("noscroll"); } });
/* Android back button closes the open screen instead of leaving the page */
addEventListener("popstate", () => { if (openId) { const el = document.getElementById("sheet-" + openId); if (el) el.hidden = true; openId = ""; document.body.classList.remove("noscroll"); } });

/* Page heading like the mock: small eyebrow, big title, light second line, action buttons on the right */
export const pageHead = (eyebrow, title, light, actions = "", note = "") => `<div class="phead"><div><div class="eyebrow">${eyebrow}</div><h2 class="ptitle">${title}${light ? `<span>${light}</span>` : ""}</h2>${note ? `<p class="pnote">${note}</p>` : ""}</div>${actions ? `<div class="pact">${actions}</div>` : ""}</div>`;

/* Simple line icons for the tab bar */
const P = {
  dash: '<path d="M3 13h8V3H3zM13 21h8V11h-8zM3 21h8v-6H3zM13 3v6h8V3z"/>',
  metres: '<path d="M12 2v14M8 12l4 4 4-4"/><path d="M5 22h14"/>',
  docs: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h6"/>',
  bd: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4z"/>',
  missing: '<circle cx="12" cy="12" r="9"/><path d="M12 7v6M12 16.5v.5"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
  issues: '<path d="M3 7h13v10H3zM16 10h3l2 3v4h-5"/><circle cx="7" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/>',
  stock: '<path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  import: '<path d="M12 3v12M7 10l5 5 5-5"/><path d="M5 21h14"/>',
  deploy: '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>',
  hours: '<circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2M9 2h6"/>',
  money: '<path d="M3 17l6-6 4 4 8-8"/><path d="M14 7h7v7"/>',
  maint: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4z"/>',
  shifts: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4M16 2v4M3 10h18"/>',
  bus: '<rect x="4" y="3" width="16" height="15" rx="2"/><path d="M4 11h16M8 21v-3M16 21v-3"/><circle cx="8" cy="14.5" r="1"/><circle cx="16" cy="14.5" r="1"/>',
};
export const icon = k => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[k] || P.dash}</svg>`;
export const tabBar = (tabs, cur) => `<nav class="ftabs" role="tablist">${tabs.map(([k, t, ic]) => `<button role="tab" aria-selected="${cur === k}" data-tab="${k}">${icon(ic || k)}<span>${t}</span></button>`).join("")}</nav>`;
/* Text search: every word typed must be somewhere in the line */
export const matches = (text, q) => { const t = String(text || "").toLowerCase(); return String(q || "").toLowerCase().split(/\s+/).filter(Boolean).every(w => t.includes(w)); };
