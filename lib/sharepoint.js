/* SharePoint filing for the TKR app.
   Files are sent to the TKR SharePoint site with Microsoft Graph, signed in with the person's own Microsoft 365 account.
   Set up once under GET Forecast → Settings → SharePoint (tenant ID, application ID, site address). Nothing here is secret. */
let pca = null, cfg = null, account = null, drive = null, loading = null;
const SCOPES = ["Sites.ReadWrite.All"];
const G = "https://graph.microsoft.com/v1.0";
const loadScript = src => new Promise((res, rej) => { if (window.msal) return res(); const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); });

export const spConfigured = c => !!(c && c.tenant && c.clientId && c.site);
export const spReady = () => !!(pca && account);
export const spWho = () => account?.username || "";
export const spRoot = () => (cfg?.folder || "TKR App").replace(/^\/+|\/+$/g, "");

/* Call when the settings are known. Signs in silently if this phone or PC signed in before. */
export async function spInit(config) {
  if (!spConfigured(config)) { pca = null; cfg = null; account = null; drive = null; return false; }
  if (cfg && cfg.tenant === config.tenant && cfg.clientId === config.clientId && cfg.site === config.site && pca) { cfg = config; return spReady(); }
  cfg = config; drive = null;
  loading = (async () => {
    await loadScript(new URL("./msal-browser.min.js", import.meta.url).href);
    pca = new window.msal.PublicClientApplication({
      auth: { clientId: cfg.clientId.trim(), authority: "https://login.microsoftonline.com/" + cfg.tenant.trim(), redirectUri: new URL("../auth.html", import.meta.url).href },
      cache: { cacheLocation: "localStorage" },
    });
    await pca.initialize();
    account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
  })();
  await loading; return spReady();
}
export async function spSignIn() {
  if (!pca) throw new Error("SharePoint is not set up yet");
  const r = await pca.loginPopup({ scopes: SCOPES, prompt: "select_account" });
  account = r.account; pca.setActiveAccount(account); drive = null; await site(); return account.username;
}
export async function spSignOut() { if (pca && account) { try { await pca.clearCache({ account }); } catch {} } account = null; drive = null; }
async function token() {
  if (!pca || !account) throw new Error("Not signed in to SharePoint");
  try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch { const r = await pca.acquireTokenPopup({ scopes: SCOPES, account }); return r.accessToken; }
}
async function graph(url, opt = {}) {
  const t = await token();
  const res = await fetch(url.startsWith("http") ? url : G + url, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
  if (!res.ok) { const x = await res.text().catch(() => ""); throw new Error("SharePoint " + res.status + " " + x.slice(0, 200)); }
  return res.status === 204 ? null : res.json();
}
/* The site's Documents library */
async function site() {
  if (drive) return drive;
  const u = new URL(cfg.site.trim()); const path = u.pathname.replace(/\/+$/, "");
  const s = await graph(`/sites/${u.hostname}:${path || "/"}`);
  const d = await graph(`/sites/${s.id}/drive`);
  drive = { siteId: s.id, driveId: d.id, webUrl: d.webUrl }; return drive;
}
const clean = s => String(s).replace(/["*:<>?\\|#%~&{}]/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
const enc = p => p.split("/").map(x => encodeURIComponent(clean(x))).join("/");

/* Sends one file. folder is a path under the app's root folder, for example "Drilling/2026/11 November/2026-11-10/Tally sheets".
   Folders that don't exist yet are made by SharePoint. Returns { id, url, path }. */
export async function spUpload(folder, name, blob) {
  const d = await site(); const path = `${spRoot()}/${folder}/${name}`.replace(/\/+/g, "/");
  const base = `/drives/${d.driveId}/root:/${enc(path)}`;
  let item;
  if (blob.size <= 4 * 1024 * 1024) item = await graph(base + ":/content?@microsoft.graph.conflictBehavior=replace", { method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob });
  else {
    const ses = await graph(base + ":/createUploadSession", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) });
    const STEP = 3276800; let pos = 0, out = null;                       // 10 × 320 KiB, as Graph asks
    while (pos < blob.size) { const end = Math.min(pos + STEP, blob.size);
      const res = await fetch(ses.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${pos}-${end - 1}/${blob.size}` }, body: blob.slice(pos, end) });
      if (!res.ok) throw new Error("SharePoint upload " + res.status); out = await res.json().catch(() => null); pos = end; }
    item = out;
  }
  return { id: item?.id || "", url: item?.webUrl || "", path };
}
export async function spDelete(id) { if (!id) return; const d = await site(); try { await graph(`/drives/${d.driveId}/items/${id}`, { method: "DELETE" }); } catch (e) { console.warn(e); } }
/* Address of a folder in the browser, to open it in SharePoint */
export async function spFolderUrl(folder) { const d = await site(); try { return (await graph(`/drives/${d.driveId}/root:/${enc(`${spRoot()}/${folder}`.replace(/\/+$/, ""))}`)).webUrl; } catch { return d.webUrl; } }
