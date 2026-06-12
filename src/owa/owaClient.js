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
// MULTI-ACCOUNT: each captured token is filed under an "account" keyed by its
// mailbox (X-AnchorMailbox / UPN). Signing into a second mailbox creates a
// second account with its own token and its own address book. The account model
// lives in lib/storage.js; this module operates on a given account.

import { getConfig, getAccounts, upsertOwaAccount } from "../lib/storage.js";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Dedupe in-memory: last auth header seen per anchor, so we don't write storage
// on every single request the Outlook tab makes.
const lastSeenByAnchor = new Map();

/**
 * Called by the background webRequest listener for every OWA request that
 * carries an Authorization header. Files the freshest Bearer token under its
 * account (creating the account on first sight).
 * @param {string} authValue  full header value, e.g. "Bearer eyJ…"
 * @param {string|null} anchorMailbox  X-AnchorMailbox value, identifies account
 * @returns {Promise<{account: object, isNew: boolean}|null>}
 */
export async function captureAuthHeader(authValue, anchorMailbox) {
  if (!authValue || !authValue.startsWith("Bearer ")) return null;
  const key = (anchorMailbox || "").toLowerCase();
  if (lastSeenByAnchor.get(key) === authValue) return null; // unchanged
  lastSeenByAnchor.set(key, authValue);
  const res = await upsertOwaAccount(authValue, anchorMailbox);
  console.debug(
    `[Directorium] Captured OWA token for "${res.account.label}"`,
    res.isNew ? "(new account)" : "",
    res.account.anchor ? `anchor ${res.account.anchor}` : ""
  );
  return res;
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

/** Re-read an account from storage by id (to pick up a refreshed token). */
async function reloadAccount(id) {
  return (await getAccounts()).find((a) => a.id === id) || null;
}

/**
 * Silently refresh an account's token: open its mailbox in a background
 * (inactive) tab, wait until a newer token is captured from its traffic, then
 * close it. Works without user interaction while the Microsoft SSO session for
 * that mailbox is still valid.
 *
 * Note: in a shared browser session the opened tab loads whichever mailbox the
 * cookies resolve to; tokens are still routed to the correct account by anchor,
 * so a refresh may land on a different account than requested. The per-account
 * token stays correct either way.
 * @returns {Promise<{ok:boolean, refreshed:boolean, account:object}>}
 */
export async function silentRefreshToken(account) {
  const before = account.token?.ts || 0;
  const owaUrl = account.owaUrl || (await getConfig()).owaUrl;
  const tab = await browser.tabs.create({ url: `${owaUrl}/mail/`, active: false });
  const tabId = tab.id;
  const deadline = Date.now() + 45_000;
  try {
    while (Date.now() < deadline) {
      await sleep(1500);
      const fresh = await reloadAccount(account.id);
      if (fresh && (fresh.token?.ts || 0) > before) {
        console.debug(`[Directorium] Token silently refreshed for "${fresh.label}".`);
        return { ok: true, refreshed: true, account: fresh };
      }
    }
    console.debug("[Directorium] Silent refresh got no token (SSO session likely expired).");
    return { ok: false, refreshed: false, account };
  } finally {
    browser.tabs.remove(tabId).catch(() => {});
  }
}

/** Refresh every account whose token is missing, old, or expiring soon. */
export async function maybeRefresh() {
  const cfg = await getConfig();
  if (cfg.authMode !== "owa") return;
  for (const account of await getAccounts()) {
    const tok = account.token;
    const ageMs = tok ? Date.now() - tok.ts : Infinity;
    const exp = tok ? parseExp(tok.auth) : null;
    const expiringSoon = exp ? exp - Date.now() < 2 * 3600 * 1000 : true;
    if (!tok || ageMs > 6 * 3600 * 1000 || expiringSoon) {
      await silentRefreshToken(account);
    }
  }
}

/**
 * Interactive sign-in to ADD an account: open webmail and KEEP THE TAB OPEN. We
 * resolve as soon as a token is captured into a new or updated account.
 * @returns {Promise<{ok:boolean, keepTabOpen:boolean, account:object}>}
 */
export async function signInOwa() {
  const cfg = await getConfig();
  const before = new Map(
    (await getAccounts()).map((a) => [a.id, a.token?.ts || 0])
  );

  const tab = await browser.tabs.create({ url: `${cfg.owaUrl}/mail/`, active: true });
  const tabId = tab.id;
  const deadline = Date.now() + 300_000;

  while (Date.now() < deadline) {
    await sleep(2000);

    try {
      await browser.tabs.get(tabId);
    } catch {
      throw new Error("Sign-in tab was closed before a token could be captured.");
    }

    const accounts = await getAccounts();
    const fresh = accounts.find(
      (a) => (a.token?.ts || 0) > (before.get(a.id) || 0)
    );
    if (fresh && fresh.token?.ts >= Date.now() - 300_000) {
      console.debug(`[Directorium] OWA sign-in captured account "${fresh.label}".`);
      // Leave the tab open so the page keeps the token fresh.
      return {
        ok: true,
        keepTabOpen: true,
        account: { id: fresh.id, label: fresh.label, anchor: fresh.anchor },
      };
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

/** Call OWA service.svc FindPeople with a specific account's Bearer token. */
async function findPeople(account, query, { folderId = "directory", offset = 0, max = 100, allowRefresh = true } = {}) {
  const tok = account.token;
  if (!tok?.auth) {
    throw new Error(
      `No token for "${account.label}" yet. Sign in (OWA) and keep the Outlook tab open, then retry.`
    );
  }
  const owaUrl = account.owaUrl || (await getConfig()).owaUrl;
  const anchor = tok.anchor || account.anchor || "";

  const resp = await fetch(`${owaUrl}/owa/service.svc?action=FindPeople`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      Accept: "application/json",
      Authorization: tok.auth,
      Action: "FindPeople",
      "X-AnchorMailbox": anchor,
      "X-RoutingParameter-SessionKey": anchor,
      "X-Requested-With": "XMLHttpRequest",
    },
    body: JSON.stringify(findPeopleBody(query, offset, max, folderId)),
  });

  const text = await resp.text();
  console.debug(`[Directorium] OWA FindPeople "${query}" [${account.label}] HTTP ${resp.status}:\n${text.slice(0, 1500)}`);

  if (resp.status === 401) {
    if (allowRefresh) {
      // Token went stale — try one silent refresh, then retry once.
      const r = await silentRefreshToken(account);
      if (r.refreshed) {
        return findPeople(r.account, query, { folderId, offset, max, allowRefresh: false });
      }
    }
    throw new Error(`OWA token for "${account.label}" expired and silent refresh failed — Sign in (OWA) again.`);
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

export async function searchGal(query, account) {
  if (!account) throw new Error("No OWA account configured — Sign in (OWA) first.");
  if (!query || !query.trim()) return { contacts: [], rawXml: "" };
  const { json, raw } = await findPeople(account, query.trim());
  const contacts = extractPersonas(json).map(personaToContact).filter(Boolean);
  console.debug(`[Directorium] OWA "${query}" [${account.label}]: ${contacts.length} contacts`);
  return { contacts, rawXml: raw };
}

export async function enumerateGal(account, onBatch) {
  const prefixes = "abcdefghijklmnopqrstuvwxyz0123456789".split("");
  const seen = new Set();
  const all = [];
  let lastRaw = "";
  for (const prefix of prefixes) {
    try {
      const { contacts, rawXml } = await searchGal(prefix, account);
      lastRaw = rawXml;
      const batch = [];
      for (const c of contacts) {
        const key = c.PrimaryEmail?.toLowerCase() || c.DisplayName?.toLowerCase();
        if (key && !seen.has(key)) { seen.add(key); batch.push(c); }
      }
      all.push(...batch);
      if (onBatch && batch.length) onBatch(batch, all.length);
    } catch (err) {
      console.warn(`[Directorium] OWA enum "${prefix}" [${account.label}] failed:`, err.message);
    }
  }
  return { contacts: all, rawXml: lastRaw };
}

export async function testConnection(account) {
  const { contacts, rawXml } = await searchGal("a", account);
  return { ok: true, contactCount: contacts.length, rawXml };
}

/**
 * Fetch ALL of one account's personal contacts (their own Contacts folder),
 * paging through FindPeople. Unlike the GAL, this set is small enough to sync
 * into a real, browsable Thunderbird address book.
 * @returns {Promise<Array<Record<string,string>>>}
 */
export async function fetchPersonalContacts(account) {
  const pageSize = 100;
  let offset = 0;
  const all = [];
  for (;;) {
    const { json } = await findPeople(account, "", { folderId: "contacts", offset, max: pageSize });
    const personas = extractPersonas(json);
    for (const p of personas) {
      const c = personaToContact(p);
      if (c) all.push(c);
    }
    if (personas.length < pageSize) break;
    offset += pageSize;
    if (offset > 10000) break; // safety
  }
  console.debug(`[Directorium] Fetched ${all.length} personal contacts for "${account.label}"`);
  return all;
}

/** Diagnostic status for a single account (held token, age, time left). */
function statusFor(account) {
  const tok = account.token;
  if (!tok) {
    return { id: account.id, label: account.label, anchor: account.anchor, hasToken: false };
  }
  const exp = parseExp(tok.auth);
  return {
    id: account.id,
    label: account.label,
    anchor: account.anchor,
    hasToken: true,
    ageSeconds: Math.round((Date.now() - tok.ts) / 1000),
    expiresInSeconds: exp ? Math.round((exp - Date.now()) / 1000) : null,
  };
}

/** Diagnostic: status of every account (used by the options page). */
export async function tokenStatus() {
  const accounts = await getAccounts();
  return { accounts: accounts.map(statusFor) };
}
