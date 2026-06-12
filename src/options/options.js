import { getConfig, setConfig, DEFAULT_CONFIG } from "../lib/storage.js";
import { discoverEwsUrl } from "../autodiscover/autodiscover.js";

const $ = (id) => document.getElementById(id);

const fields = [
  "authMode",
  "ewsUrl",
  "owaUrl",
  "exchangeVersion",
  "addressBookName",
  "clientId",
  "tenant",
  "scopes",
  "username",
  "password",
  "email",
];

function setMsg(text, kind = "") {
  const el = $("msg");
  el.textContent = text;
  el.className = `msg ${kind}`;
}

function applyAuthModeVisibility() {
  const mode = $("authMode").value;
  const isEwsOauth = mode === "oauth" || mode === "oauth-interactive";
  const isOwa = mode === "owa";
  const isBasic = mode === "basic";
  const needsOwnApp = mode === "oauth-interactive";

  $("oauthSection").classList.toggle("hidden", !isEwsOauth);
  $("owaSection").classList.toggle("hidden", !isOwa);
  $("basicSection").classList.toggle("hidden", !isBasic);
  $("redirectBox").classList.toggle("hidden", !needsOwnApp);
  // In OWA mode the address-book names come from the accounts, not this field.
  $("addressBookNameRow").classList.toggle("hidden", isOwa);
  if (needsOwnApp) showRedirectUrl();
  if (isOwa) renderAccounts();
}

// ---------------------------------------------------------------------------
// OWA accounts list (one address book per signed-in mailbox)
// ---------------------------------------------------------------------------
function fmtTimeLeft(s) {
  if (s == null) return "unknown expiry";
  if (s <= 0) return "expired";
  return `${Math.floor(s / 3600)}h ${Math.floor((s % 3600) / 60)}m left`;
}

async function renderAccounts() {
  const box = $("accountList");
  if (!box) return;
  let accounts = [];
  let statusById = {};
  try {
    const listResp = await messenger.runtime.sendMessage({ type: "listAccounts" });
    if (listResp.ok) accounts = listResp.result.accounts || [];
    const statResp = await messenger.runtime.sendMessage({ type: "owaTokenStatus" });
    if (statResp.ok) {
      for (const s of statResp.result.accounts || []) statusById[s.id] = s;
    }
  } catch (_) {/* ignore */}

  box.textContent = "";
  if (!accounts.length) {
    const empty = document.createElement("p");
    empty.style.cssText = "color:var(--muted);font-size:13px";
    empty.textContent = "No accounts yet. Click “Add account (sign in)” above.";
    box.appendChild(empty);
    return;
  }

  for (const acct of accounts) {
    const st = statusById[acct.id] || {};
    const row = document.createElement("div");
    row.className = "account-row";

    const name = document.createElement("input");
    name.type = "text";
    name.value = acct.label;
    name.title = "Address book name";
    name.className = "account-name";

    const status = document.createElement("span");
    status.className = "account-status";
    status.textContent = st.hasToken
      ? `● ${acct.anchor || ""} — token OK, ${fmtTimeLeft(st.expiresInSeconds)}`
      : `○ ${acct.anchor || ""} — no token (sign in)`;
    status.style.color = st.hasToken ? "var(--ok)" : "var(--err)";

    const rename = document.createElement("button");
    rename.type = "button";
    rename.textContent = "Rename";
    rename.addEventListener("click", async () => {
      const resp = await messenger.runtime.sendMessage({
        type: "renameAccount", id: acct.id, label: name.value.trim(),
      });
      setMsg(resp.ok
        ? "Renamed. Restart Thunderbird to see the new address-book name."
        : `Rename failed: ${resp.error}`, resp.ok ? "ok" : "err");
    });

    const sync = document.createElement("button");
    sync.type = "button";
    sync.textContent = "Sync contacts";
    sync.addEventListener("click", () => syncContacts(acct.id));

    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Remove";
    remove.addEventListener("click", async () => {
      const resp = await messenger.runtime.sendMessage({ type: "removeAccount", id: acct.id });
      if (resp.ok) {
        setMsg(`Removed “${acct.label}”.`, "ok");
        renderAccounts();
      } else {
        setMsg(`Remove failed: ${resp.error}`, "err");
      }
    });

    const top = document.createElement("div");
    top.className = "row";
    top.append(name, rename);
    const actions = document.createElement("div");
    actions.className = "row";
    actions.append(sync, remove);

    row.append(top, status, actions);
    box.appendChild(row);
  }
}

async function showRedirectUrl() {
  try {
    const resp = await messenger.runtime.sendMessage({ type: "getRedirectUrl" });
    if (resp.ok) $("redirectUrl").textContent = resp.result.redirectUrl;
  } catch (_) {/* ignore */}
}

async function load() {
  const cfg = await getConfig();
  for (const id of fields) {
    if ($(id)) $(id).value = cfg[id] ?? "";
  }
  applyAuthModeVisibility();
}

async function save() {
  const patch = {};
  for (const id of fields) {
    if ($(id)) patch[id] = $(id).value.trim();
  }
  await setConfig(patch);
  setMsg("Saved. (Address-book name changes apply after restarting Thunderbird.)", "ok");
}

// ---------------------------------------------------------------------------
// Autodiscover
// ---------------------------------------------------------------------------
async function autoDiscover() {
  const email = $("email").value.trim();
  if (!email || !email.includes("@")) {
    setMsg("Enter your email address first.", "err");
    return;
  }

  $("discoverBtn").disabled = true;
  setMsg("Detecting server…");

  try {
    const domain = email.split("@")[1];

    // Request optional host permissions for the target domain so the extension
    // can reach autodiscover.{domain} (already covered for outlook.office365.com).
    const extraOrigins = [
      `https://autodiscover.${domain}/*`,
      `https://${domain}/*`,
    ];
    try {
      await browser.permissions.request({ origins: extraOrigins });
    } catch (_) {
      // User denied or API unavailable — proceed anyway; O365 path needs no
      // extra permission and is already covered by host_permissions.
    }

    const { ewsUrl, method } = await discoverEwsUrl(email);

    $("ewsUrl").value = ewsUrl;

    // Auto-set auth mode based on which path found the URL.
    if (method === "autodiscover-v2-o365" || ewsUrl.includes("office365.com")) {
      $("authMode").value = "oauth";
    } else {
      $("authMode").value = "basic";
      // Pre-fill username with the full email (works for UPN-style logins).
      if (!$("username").value) $("username").value = email;
    }
    applyAuthModeVisibility();

    await save();
    setMsg(`✓ Found: ${ewsUrl}`, "ok");
  } catch (err) {
    setMsg(`Auto-detect failed: ${err.message}`, "err");
  } finally {
    $("discoverBtn").disabled = false;
  }
}

// ---------------------------------------------------------------------------
// Sign in (OAuth device code)
// ---------------------------------------------------------------------------
async function signIn() {
  setMsg("");
  $("signInBtn").disabled = true;
  try {
    await save();

    const mode = $("authMode").value;
    if (mode === "oauth-interactive") {
      // One await — the browser sign-in window/tab handles everything.
      setMsg("Opening Microsoft sign-in…");
      const resp = await messenger.runtime.sendMessage({ type: "startInteractiveSignIn" });
      if (!resp.ok) throw new Error(resp.error);
      setMsg("✓ Signed in successfully.", "ok");
      $("signInBtn").disabled = false;
      return;
    }

    // Device code flow.
    const resp = await messenger.runtime.sendMessage({ type: "startSignIn" });
    if (!resp.ok) throw new Error(resp.error);

    const { userCode, verificationUri } = resp.result;
    $("deviceCode").textContent = userCode;
    const link = $("deviceUri");
    link.href = verificationUri;
    link.textContent = verificationUri;
    $("deviceBox").classList.remove("hidden");
    $("signInStatus").textContent = "Waiting for you to sign in…";
  } catch (err) {
    setMsg(`Sign-in failed: ${err.message}`, "err");
    $("signInBtn").disabled = false;
  }
}

// Background notifies us when polling resolves.
messenger.runtime.onMessage.addListener((msg) => {
  if (msg?.type !== "signInComplete") return;
  $("signInBtn").disabled = false;
  if (msg.ok) {
    $("signInStatus").textContent = "✓ Signed in. GAL is ready.";
    setMsg("Signed in successfully.", "ok");
  } else {
    $("signInStatus").textContent = `✗ ${msg.error}`;
    setMsg(`Sign-in failed: ${msg.error}`, "err");
  }
});

// ---------------------------------------------------------------------------
// Test & sync
// ---------------------------------------------------------------------------
async function owaSignIn() {
  setMsg("");
  $("owaSignInBtn").disabled = true;
  setMsg("Opening webmail sign-in… complete the login in the new tab.");
  try {
    await save();
    const resp = await messenger.runtime.sendMessage({ type: "startOwaSignIn" });
    if (!resp.ok) throw new Error(resp.error);
    const label = resp.result?.account?.label;
    setMsg(
      `✓ Signed in${label ? ` to “${label}”` : ""}. A new address book will appear (restart Thunderbird if it doesn't show immediately).`,
      "ok"
    );
    renderAccounts();
  } catch (err) {
    setMsg(`OWA sign-in failed: ${err.message}`, "err");
  } finally {
    $("owaSignInBtn").disabled = false;
  }
}

async function owaInspect() {
  $("owaInspectBtn").disabled = true;
  try {
    const resp = await messenger.runtime.sendMessage({ type: "owaTokenStatus" });
    if (!resp.ok) throw new Error(resp.error);
    const accounts = resp.result.accounts || [];
    await renderAccounts();
    const withToken = accounts.filter((a) => a.hasToken).length;
    if (accounts.length) {
      setMsg(`${withToken}/${accounts.length} account(s) have a live token. Auto-refreshes in the background.`, "ok");
    } else {
      setMsg("No accounts yet — Add account (sign in) and keep the Outlook tab open.", "err");
    }
  } catch (err) {
    setMsg(`Status check failed: ${err.message}`, "err");
  } finally {
    $("owaInspectBtn").disabled = false;
  }
}

async function syncContacts(accountId) {
  setMsg("Syncing your personal contacts…");
  try {
    const resp = await messenger.runtime.sendMessage({ type: "syncPersonalContacts", accountId });
    if (!resp.ok) throw new Error(resp.error);
    setMsg(`✓ Synced ${resp.result.count} contacts into "${resp.result.bookName}".`, "ok");
  } catch (err) {
    setMsg(`Contacts sync failed: ${err.message}`, "err");
  }
}

async function test() {
  setMsg("Testing…");
  try {
    await save();
    const resp = await messenger.runtime.sendMessage({ type: "testConnection" });
    if (!resp.ok) throw new Error(resp.error);
    setMsg("✓ Connection OK — GAL is reachable.", "ok");
  } catch (err) {
    setMsg(`Test failed: ${err.message}`, "err");
  }
}

async function syncFull() {
  setMsg("Syncing full GAL (this can take a while)…");
  try {
    const resp = await messenger.runtime.sendMessage({ type: "enumerateGal" });
    if (!resp.ok) throw new Error(resp.error);
    setMsg(`✓ Retrieved ${resp.result.count} directory entries.`, "ok");
  } catch (err) {
    setMsg(`Sync failed: ${err.message}`, "err");
  }
}

// ---------------------------------------------------------------------------
// Wire up
// ---------------------------------------------------------------------------
async function diagSearch() {
  const query = $("diagQuery").value.trim() || "a";
  $("diagBtn").disabled = true;
  $("diagResult").classList.add("hidden");
  setMsg("Searching…");

  try {
    await save();
    const resp = await messenger.runtime.sendMessage({ type: "searchGal", query });
    const result = $("diagResult");
    const summary = $("diagSummary");
    const xmlPre = $("diagXml");

    if (resp.ok) {
      summary.textContent = `✓ Got ${resp.result.count} contact(s) for "${query}".`;
      summary.style.color = resp.result.count > 0 ? "var(--ok)" : "var(--err)";
      if (resp.result.contacts?.length > 0) {
        summary.textContent += ` First: ${resp.result.contacts[0].DisplayName} <${resp.result.contacts[0].PrimaryEmail}>`;
      }
    } else {
      summary.textContent = `✗ Error: ${resp.error}`;
      summary.style.color = "var(--err)";
    }

    xmlPre.textContent = formatXml(resp.result?.rawXml || resp.rawXml || "(no XML)");
    result.classList.remove("hidden");
    setMsg("");
  } catch (err) {
    setMsg(`Diagnostics failed: ${err.message}`, "err");
  } finally {
    $("diagBtn").disabled = false;
  }
}

function formatXml(xml) {
  if (!xml || xml === "(no XML)") return xml;
  try {
    // Pretty-print by adding newlines after tags
    return xml
      .replace(/></g, ">\n<")
      .replace(/(<[^/][^>]*>)\n(<\/)/g, "$1$2")
      .slice(0, 8000);
  } catch {
    return xml.slice(0, 8000);
  }
}

$("authMode").addEventListener("change", applyAuthModeVisibility);
$("discoverBtn").addEventListener("click", autoDiscover);
$("saveBtn").addEventListener("click", save);
$("signInBtn").addEventListener("click", signIn);
$("owaSignInBtn").addEventListener("click", owaSignIn);
$("owaInspectBtn").addEventListener("click", owaInspect);
$("testBtn").addEventListener("click", test);
$("syncBtn").addEventListener("click", syncFull);
$("syncContactsBtn").addEventListener("click", () => syncContacts());
$("diagBtn").addEventListener("click", diagSearch);

$("ewsUrl").placeholder = DEFAULT_CONFIG.ewsUrl;
$("clientId").placeholder = DEFAULT_CONFIG.clientId;
$("scopes").placeholder = DEFAULT_CONFIG.scopes;

load();
