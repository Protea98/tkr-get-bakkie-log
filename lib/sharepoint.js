/* Filing documents in Microsoft 365 for the TKR app: a folder in a OneDrive (1 TB with a TKR Microsoft 365 account),
   or a SharePoint site's Documents. Microsoft Graph, signed in with the person's own TKR Microsoft account.
   Set up once under GET Forecast → Settings → OneDrive / SharePoint. Nothing here is secret. */
let pca = null, cfg = null, account = null, drive = null, loading = null;
const G = "https://graph.microsoft.com/v1.0";
const loadScript = src => new Promise((res, rej) => { if (window.msal) return res(); const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
const isSite = () => !!(cfg && cfg.site);
const scopes = () => isSite() ? ["Sites.ReadWrite.All"] : ["Files.ReadWrite.All"];

export const spConfigured = c => !!(c && c.tenant && c.clientId);
export const spReady = () => !!(pca && account);
export const spWho = () => account?.username || "";
export const spRoot = () => (cfg?.folder || "TKR App").replace(/^\/+|\/+$/g, "");
export const spPlace = () => isSite() ? "SharePoint" : "OneDrive";
export const spLinked = () => isSite() || !!(cfg && cfg.driveId && cfg.rootId);

/* Call when the settings are known. Signs in silently if this phone or PC signed in before. */
export async function spInit(config) {
  if (!spConfigured(config)) { pca = null; cfg = null; account = null; drive = null; return false; }
  const same = cfg && pca && cfg.tenant === config.tenant && cfg.clientId === config.clientId;
  cfg = config; drive = null;
  if (same) return spReady();
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
  if (!pca) throw new Error("Not set up yet");
  const r = await pca.loginPopup({ scopes: scopes(), prompt: "select_account" });
  account = r.account; pca.setActiveAccount(account); drive = null; return account.username;
}
export async function spSignOut() { if (pca && account) { try { await pca.clearCache({ account }); } catch {} } account = null; drive = null; }
async function token() {
  if (!pca || !account) throw new Error("Not signed in to Microsoft");
  try { return (await pca.acquireTokenSilent({ scopes: scopes(), account })).accessToken; }
  catch { const r = await pca.acquireTokenPopup({ scopes: scopes(), account }); return r.accessToken; }
}
async function graph(url, opt = {}) {
  const t = await token();
  const res = await fetch(url.startsWith("http") ? url : G + url, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
  if (!res.ok) { const x = await res.text().catch(() => ""); throw new Error("Microsoft " + res.status + " " + x.slice(0, 200)); }
  return res.status === 204 ? null : res.json();
}
const clean = s => String(s).replace(/["*:<>?\\|#%~&{}]/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
const enc = p => p.split("/").filter(Boolean).map(x => encodeURIComponent(clean(x))).join("/");

/* OneDrive: the person who links it (an admin) gets a "TKR App" folder in their own OneDrive.
   Everyone else files into that same folder once the owner has shared it with them in OneDrive (Can edit). */
export async function spLinkOneDrive() {
  const d = await graph("/me/drive");
  let f; try { f = await graph(`/me/drive/root:/${enc(spRoot())}`); }
  catch { f = await graph("/me/drive/root/children", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: clean(spRoot()), folder: {}, "@microsoft.graph.conflictBehavior": "fail" }) }); }
  return { driveId: d.id, rootId: f.id, owner: account?.username || "", webUrl: f.webUrl || "" };
}
async function where() {
  if (drive) return drive;
  if (isSite()) {
    const u = new URL(cfg.site.trim()); const path = u.pathname.replace(/\/+$/, "");
    const s = await graph(`/sites/${u.hostname}:${path || "/"}`), d = await graph(`/sites/${s.id}/drive`);
    drive = { driveId: d.id, rootId: "", webUrl: d.webUrl };
  } else {
    if (!cfg.driveId || !cfg.rootId) throw new Error("OneDrive folder not linked yet");
    drive = { driveId: cfg.driveId, rootId: cfg.rootId, webUrl: cfg.webUrl || "" };
  }
  return drive;
}
/* the address of a path under the app's main folder */
const at = (d, rel) => d.rootId ? `/drives/${d.driveId}/items/${d.rootId}:/${enc(rel)}` : `/drives/${d.driveId}/root:/${enc(spRoot() + "/" + rel)}`;

/* Sends one file. folder is a path under the main folder, for example "Drilling/November 2026 Metres/01 DAY".
   Folders that don't exist yet are made by Microsoft. Returns { id, url, path }. */
export async function spUpload(folder, name, blob) {
  const d = await where(); const rel = `${folder}/${name}`.replace(/\/+/g, "/"), base = at(d, rel);
  let item;
  if (blob.size <= 4 * 1024 * 1024) item = await graph(base + ":/content?@microsoft.graph.conflictBehavior=replace", { method: "PUT", headers: { "Content-Type": blob.type || "application/octet-stream" }, body: blob });
  else {
    const ses = await graph(base + ":/createUploadSession", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) });
    const STEP = 3276800; let pos = 0, out = null;                       // 10 × 320 KiB, as Graph asks
    while (pos < blob.size) { const end = Math.min(pos + STEP, blob.size);
      const res = await fetch(ses.uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${pos}-${end - 1}/${blob.size}` }, body: blob.slice(pos, end) });
      if (!res.ok) throw new Error("Upload " + res.status); out = await res.json().catch(() => null); pos = end; }
    item = out;
  }
  return { id: item?.id || "", url: item?.webUrl || "", path: `${spRoot()}/${rel}` };
}
export async function spDelete(id) { if (!id) return; const d = await where(); try { await graph(`/drives/${d.driveId}/items/${id}`, { method: "DELETE" }); } catch (e) { console.warn(e); } }
/* Address of a folder in the browser */
export async function spFolderUrl(folder) { const d = await where(); try { return (await graph(at(d, folder.replace(/\/+$/, "")))).webUrl; } catch { return d.webUrl; } }
