// OWA backend for the NEW Outlook (Monarch) — the "free Owl alternative" path
// for locked tenants where EWS+OAuth third-party apps are blocked.
//
// The new Outlook authenticates its internal APIs with a Bearer token minted by
// the first-party "One Outlook Web" client (pre-consented in the tenant — the
// user is already using it). We can't mint that token ourselves (every clean
// OAuth flow is blocked by the tenant), so instead we OBSERVE it:
//
//   1. The user signs in to webmail in a normal tab (passes Conditional Access).
//   2. A webRequest listener (registered in background.js) captures the
//      `Authorization: Bearer …` header from the page's own API calls.
//   3. We reuse that token to call the GAL/people API ourselves. Host
//      permission lets the extension call it cross-origin without CORS issues.
//
// While the Outlook tab stays open it keeps refreshing the token, so ours stays
// fresh. This is the user's own token for the user's own directory access.

import { getConfig } from "../lib/storage.js";

const TOKEN_KEY = "owaToken";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// In-memory cache (fast path); also persisted so it survives background restarts.
let cachedToken = null;

/**
 * Called by the background webRequest listener for every OWA request that
 * carries an Authorization header. Stores the freshest Bearer token.
 * @param {string} authValue  full header value, e.g. "Bearer eyJ…"
 * @param {string|null} anchorMailbox  X-AnchorMailbox value, for routing
 */
export function captureAuthHeader(authValue, anchorMailbox) {
  if (!authValue || !authValue.startsWith("Bearer ")) return;
  // Skip if unchanged, to avoid hammering storage on every request.
  if (cachedToken && cachedToken.auth === authValue) {
    if (anchorMailbox && !cachedToken.anchor) cachedToken.anchor = anchorMailbox;
    return;
  }
  cachedToken = { auth: authValue, anchor: anchorMailbox || cachedToken?.anchor || null, ts: Date.now() };
  messenger.storage.local.set({ [TOKEN_KEY]: cachedToken }).catch(() => {});
  console.debug("[Directorium] Captured fresh OWA bearer token", cachedToken.anchor ? `(anchor ${cachedToken.anchor})` : "");
}

async function getCapturedToken() {
  if (cachedToken) return cachedToken;
  const stored = await messenger.storage.local.get(TOKEN_KEY);
  cachedToken = stored[TOKEN_KEY] || null;
  return cachedToken;
}

/** Epoch ms when the JWT access token expires, or null. */
function parseExp(authValue) {
  try {
    const jwt = authValue.replace(/^Bearer\s+/i, "");
    const payload = JSON.parse(
      atob(jwt.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );
    return payload.exp ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

/**
 * Silently refresh the token: open the mailbox in a background (inactive) tab,
 * wait until a newer token is captured from its traffic, then close it. Works
 * without user interaction as long as the Microsoft SSO session is still valid.
 * @returns {Promise<{ok:boolean, refreshed:boolean}>}
 */
export async function silentRefreshToken() {
  const cfg = await getConfig();
  const before = (await getCapturedToken())?.ts || 0;

  const tab = await browser.tabs.create({ url: `${cfg.owaUrl}/mail/`, active: false });
  const tabId = tab.id;
  const deadline = Date.now() + 45_000;
  try {
    while (Date.now() < deadline) {
      await sleep(1500);
      const tok = await getCapturedToken();
      if (tok && tok.ts > before) {
        console.debug("[Directorium] OWA token silently refreshed.");
        return { ok: true, refreshed: true };
      }
    }
    console.debug("[Directorium] Silent refresh got no token (SSO session likely expired).");
    return { ok: false, refreshed: false };
  } finally {
    browser.tabs.remove(tabId).catch(() => {});
  }
}

/** Refresh proactively if the token is missing, old, or expiring soon. */
export async function maybeRefresh() {
  const cfg = await getConfig();
  if (cfg.authMode !== "owa") return;
  const tok = await getCapturedToken();
  const ageMs = tok ? Date.now() - tok.ts : Infinity;
  const exp = tok ? parseExp(tok.auth) : null;
  const expiringSoon = exp ? exp - Date.now() < 2 * 3600 * 1000 : true;
  if (!tok || ageMs > 6 * 3600 * 1000 || expiringSoon) {
    await silentRefreshToken();
  }
}

/**
 * Interactive sign-in: open webmail and KEEP THE TAB OPEN. We resolve as soon
 * as a Bearer token has been captured from the page's traffic.
 */
export async function signInOwa() {
  const cfg = await getConfig();
  const tab = await browser.tabs.create({ url: `${cfg.owaUrl}/mail/`, active: true });
  const tabId = tab.id;
  const deadline = Date.now() + 300_000;

  while (Date.now() < deadline) {
    await sleep(2000);

    // If the user closed the tab, stop.
    try {
      await browser.tabs.get(tabId);
    } catch {
      throw new Error("Sign-in tab was closed before a token could be captured.");
    }

    const tok = await getCapturedToken();
    // A token captured AFTER we opened the tab means this session is live.
    if (tok && tok.ts >= Date.now() - 300_000) {
      console.debug("[Directorium] OWA sign-in: token captured, leaving tab open for refresh.");
      // Leave the tab open so the page keeps the token fresh.
      return { ok: true, keepTabOpen: true };
    }
  }
  throw new Error(
    "No OWA token was captured. Make sure you completed sign-in and the mailbox loaded."
  );
}

/**
 * FindPeople JSON request.
 * @param folderId "directory" (GAL) or "contacts" (the user's own contacts)
 */
function findPeopleBody(query, offset, max, folderId = "directory") {
  return {
    __type: "FindPeopleJsonRequest:#Exchange",
    Header: { __type: "JsonRequestHeaders:#Exchange", RequestServerVersion: "Exchange2013" },
    Body: {
      __type: "FindPeopleRequest:#Exchange",
      IndexedPageItemView: { __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: offset, MaxEntriesReturned: max },
      QueryString: query || null,
      SearchPeopleSuggestionIndex: false,
      ParentFolderId: {
        __type: "TargetFolderId:#Exchange",
        BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: folderId },
      },
      PersonaShape: { __type: "PersonaResponseShape:#Exchange", BaseShape: "Default" },
    },
  };
}

function extractPersonas(json) {
  return json?.Body?.ResultSet || json?.ResultSet || json?.value || json?.results || [];
}

/** True for a real SMTP address (has @, not a LegacyExchangeDN "/o=…"). */
function isSmtp(addr) {
  return !!addr && addr.includes("@") && !addr.startsWith("/");
}

/**
 * Pick the SMTP address out of a persona. EX/LegacyExchangeDN contacts keep the
 * real address in `OriginalDisplayName`; some only have a SIP `ImAddress`.
 */
function bestSmtpFromEntry(e) {
  if (!e || typeof e !== "object") return isSmtp(e) ? e : "";
  if (isSmtp(e.EmailAddress)) return e.EmailAddress;
  if (isSmtp(e.OriginalDisplayName)) return e.OriginalDisplayName; // EX contacts
  return "";
}

function personaSmtp(p) {
  const list = Array.isArray(p.EmailAddresses) ? p.EmailAddresses : [];
  for (const e of list) {
    const v = bestSmtpFromEntry(e);
    if (v) return v;
  }
  const single = bestSmtpFromEntry(p.EmailAddress);
  if (single) return single;
  // Last resort: a SIP IM address often equals the SMTP address.
  if (typeof p.ImAddress === "string") {
    const sip = p.ImAddress.replace(/^SIP:/i, "");
    if (isSmtp(sip)) return sip;
  }
  return "";
}

function personaToContact(p) {
  const firstEmail = personaSmtp(p);
  const name = p.DisplayName || firstEmail;
  if (!name && !firstEmail) return null;

  const c = { DisplayName: name };
  if (firstEmail) c.PrimaryEmail = firstEmail;
  if (p.GivenName) c.FirstName = p.GivenName;
  if (p.Surname) c.LastName = p.Surname;
  if (p.Title) c.JobTitle = p.Title;
  if (p.CompanyName) c.Company = p.CompanyName;
  const dept = Array.isArray(p.Departments) ? p.Departments[0] : p.Department;
  if (dept) c.Department = dept;
  const phone =
    (Array.isArray(p.Phones) ? p.Phones[0]?.Number : null) ||
    (Array.isArray(p.PhoneNumbers) ? p.PhoneNumbers[0]?.Number : null);
  if (phone) c.WorkPhone = phone;
  return c;
}

/** Call OWA service.svc FindPeople with the captured Bearer token. */
async function findPeople(query, { folderId = "directory", offset = 0, max = 100, allowRefresh = true } = {}) {
  const cfg = await getConfig();
  const tok = await getCapturedToken();
  if (!tok?.auth) {
    throw new Error(
      "No OWA token yet. Click Sign in (OWA) and keep the Outlook tab open, then retry."
    );
  }

  const resp = await fetch(`${cfg.owaUrl}/owa/service.svc?action=FindPeople`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      Accept: "application/json",
      Authorization: tok.auth,
      Action: "FindPeople",
      "X-AnchorMailbox": tok.anchor || "",
      "X-RoutingParameter-SessionKey": tok.anchor || "",
      "X-Requested-With": "XMLHttpRequest",
    },
    body: JSON.stringify(findPeopleBody(query, offset, max, folderId)),
  });

  const text = await resp.text();
  console.debug(`[Directorium] OWA FindPeople "${query}" HTTP ${resp.status}:\n${text.slice(0, 1500)}`);

  if (resp.status === 401) {
    if (allowRefresh) {
      // Token went stale — try one silent refresh, then retry once.
      const r = await silentRefreshToken();
      if (r.refreshed) return findPeople(query, { folderId, offset, max, allowRefresh: false });
    }
    throw new Error("OWA token expired and silent refresh failed — Sign in (OWA) again.");
  }
  if (!resp.ok) {
    throw new Error(`OWA HTTP ${resp.status}: ${text.slice(0, 300)}`);
  }
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error("OWA returned non-JSON (session/endpoint issue).");
  }
  return { json, raw: text };
}

export async function searchGal(query) {
  if (!query || !query.trim()) return { contacts: [], rawXml: "" };
  const { json, raw } = await findPeople(query.trim());
  const contacts = extractPersonas(json).map(personaToContact).filter(Boolean);
  console.debug(`[Directorium] OWA "${query}": ${contacts.length} contacts`);
  return { contacts, rawXml: raw };
}

export async function enumerateGal(onBatch) {
  const prefixes = "abcdefghijklmnopqrstuvwxyz0123456789".split("");
  const seen = new Set();
  const all = [];
  let lastRaw = "";
  for (const prefix of prefixes) {
    try {
      const { contacts, rawXml } = await searchGal(prefix);
      lastRaw = rawXml;
      const batch = [];
      for (const c of contacts) {
        const key = c.PrimaryEmail?.toLowerCase() || c.DisplayName?.toLowerCase();
        if (key && !seen.has(key)) { seen.add(key); batch.push(c); }
      }
      all.push(...batch);
      if (onBatch && batch.length) onBatch(batch, all.length);
    } catch (err) {
      console.warn(`[Directorium] OWA enum "${prefix}" failed:`, err.message);
    }
  }
  return { contacts: all, rawXml: lastRaw };
}

export async function testConnection() {
  const { contacts, rawXml } = await searchGal("a");
  return { ok: true, contactCount: contacts.length, rawXml };
}

/**
 * Fetch ALL of the user's personal contacts (their own Contacts folder),
 * paging through FindPeople. Unlike the GAL, this set is small enough to sync
 * into a real, browsable Thunderbird address book.
 * @returns {Promise<Array<Record<string,string>>>}
 */
export async function fetchPersonalContacts() {
  const pageSize = 100;
  let offset = 0;
  const all = [];
  for (;;) {
    const { json } = await findPeople("", { folderId: "contacts", offset, max: pageSize });
    const personas = extractPersonas(json);
    for (const p of personas) {
      const c = personaToContact(p);
      if (c) all.push(c);
    }
    if (personas.length < pageSize) break;
    offset += pageSize;
    if (offset > 10000) break; // safety
  }
  console.debug(`[Directorium] Fetched ${all.length} personal contacts`);
  return all;
}

/** Diagnostic: do we currently hold a captured token, how old, and time left? */
export async function tokenStatus() {
  const tok = await getCapturedToken();
  if (!tok) return { hasToken: false };
  const exp = parseExp(tok.auth);
  return {
    hasToken: true,
    ageSeconds: Math.round((Date.now() - tok.ts) / 1000),
    expiresInSeconds: exp ? Math.round((exp - Date.now()) / 1000) : null,
    anchor: tok.anchor,
  };
}
