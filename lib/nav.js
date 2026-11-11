/* On every section page: a Back button at the top, and a clear pop-up when signing in fails or hangs. */
(() => {
  const css = `
  .tkr-back{flex:none;display:inline-flex;align-items:center;gap:4px;min-height:44px;padding:8px 14px 8px 10px;border-radius:10px;border:1px solid #5a554b;background:transparent;color:#f2efe8;font:600 15px/1 Roboto,"Segoe UI",Arial,sans-serif;text-decoration:none;cursor:pointer}
  .tkr-back:hover{border-color:#d4a12a;color:#d4a12a}
  .tkr-back b{font-size:22px;line-height:1;margin-top:-2px}
  .tkr-pop{position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.6);display:flex;align-items:center;justify-content:center;padding:16px}
  .tkr-pop>div{background:#1a1a1a;color:#f2efe8;border:1px solid #d4a12a;border-radius:14px;padding:20px;max-width:400px;width:100%;box-shadow:0 12px 40px rgba(0,0,0,.5);font-family:Roboto,"Segoe UI",Arial,sans-serif}
  .tkr-pop h3{margin:0 0 8px;font-size:20px;color:#d4a12a}
  .tkr-pop p{margin:0 0 16px;color:#cfc8b8;font-size:15px;line-height:1.45}
  .tkr-pop .acts{display:flex;flex-direction:column;gap:8px}
  .tkr-pop button{min-height:48px;border-radius:10px;font:600 16px Roboto,"Segoe UI",Arial,sans-serif;border:1px solid #5a554b;background:#262626;color:#f2efe8;cursor:pointer}
  .tkr-pop button.go{background:#d4a12a;border-color:#d4a12a;color:#141006}`;
  const st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  const goBack = () => {
    let same = false; try { same = document.referrer && new URL(document.referrer).origin === location.origin && document.referrer !== location.href; } catch {}
    if (same && history.length > 1) history.back(); else location.href = "../";
  };
  function addBack() {
    const logo = document.querySelector(".logolink");
    if (!logo || document.querySelector(".tkr-back")) return;
    const b = document.createElement("button"); b.type = "button"; b.className = "tkr-back"; b.setAttribute("aria-label", "Back");
    b.innerHTML = "<b>‹</b> Back"; b.onclick = goBack;
    logo.parentNode.insertBefore(b, logo);
  }

  /* ---------- sign-in pop-up ---------- */
  const IDS = ["sErr", "sMsg", "gErr", "exportErr", "lErr"];
  const FORGOT = ["forgotBtn", "sReset", "forgot", "gForgot"];
  const BAD = /wrong|not set up|didn't work|couldn't|no signal|too many|not an admin|only an admin|not approved/i;
  const BUSY = /^(signing in|checking|registering)/i;
  let pop = null, busyT = null;
  function close() { if (pop) { pop.remove(); pop = null; } }
  function show(title, text, near) {
    close();
    const forgot = FORGOT.map(id => document.getElementById(id)).find(b => b && b.offsetParent !== null);
    pop = document.createElement("div"); pop.className = "tkr-pop"; pop.setAttribute("role", "alertdialog");
    pop.innerHTML = `<div><h3></h3><p></p><div class="acts"><button class="go" data-a="again">Try again</button>${forgot && /wrong/i.test(text) ? '<button data-a="forgot">Forgot password</button>' : ""}<button data-a="reload">Start the sign-in again</button><button data-a="home">Back to the home page</button></div></div>`;
    pop.querySelector("h3").textContent = title; pop.querySelector("p").textContent = text;
    pop.onclick = e => {
      const a = e.target.closest("button")?.dataset.a; if (!a && e.target !== pop) return;
      close();
      const form = near?.closest("form") || document;
      if (a === "again" || !a) { form.querySelectorAll('input[type=password],input[inputmode=numeric][maxlength="4"]').forEach(i => i.value = ""); (form.querySelector('input[type=password]') || form.querySelector("input"))?.focus(); }
      if (a === "forgot") forgot?.click();
      if (a === "reload") location.reload();
      if (a === "home") location.href = "../";
    };
    document.body.appendChild(pop); pop.querySelector("button.go").focus();
    form_enable(near);
  }
  function form_enable(near) { const f = near?.closest("form"); if (f) f.querySelectorAll("button").forEach(b => b.disabled = false); }
  function watch(el) {
    if (el.dataset.tkrWatch) return; el.dataset.tkrWatch = 1;
    new MutationObserver(() => {
      const t = (el.textContent || "").trim(); clearTimeout(busyT);
      if (!t) return;
      if (BUSY.test(t)) { busyT = setTimeout(() => { if (BUSY.test((el.textContent || "").trim()) && el.isConnected) show("No answer yet", "Signing in is taking too long. Check the signal, then try again.", el); }, 20000); return; }
      if (BAD.test(t) && el.offsetParent !== null) show(/signal/i.test(t) ? "No signal" : "Couldn't sign in", t, el);
    }).observe(el, { childList: true, characterData: true, subtree: true });
  }
  function scan() { IDS.forEach(id => { const el = document.getElementById(id); if (el) watch(el); }); }
  const start = () => { addBack(); scan(); new MutationObserver(scan).observe(document.body, { childList: true, subtree: true }); };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();
  addEventListener("keydown", e => { if (e.key === "Escape") close(); });
})();
