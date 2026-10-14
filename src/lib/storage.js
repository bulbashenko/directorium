// Centralised access to extension settings and cached OAuth tokens.
// Settings live in storage.local so they survive restarts but never sync to
// any Mozilla account (tokens must stay on-device).

/** @typedef {"owa" | "oauth" | "oauth-interactive" | "basic"} AuthMode */

export const DEFAULT_CONFIG = {
  // "owa" = sign in through Outlook on the web and reuse its first-party token —
  // works on locked-down Office 365 tenants without admin approval.
  // "oauth-interactive" needs your own Azure app; "oauth" (device code) works on
  // lenient tenants; "basic" is for on-premise Exchange.
  authMode: /** @type {AuthMode} */ ("owa"),

  // EWS endpoint. For Office 365 this is always the same host.
  ewsUrl: "https://outlook.office365.com/EWS/Exchange.asmx",
  exchangeVersion: "Exchange2013_SP1",

  // OWA base origin, used by the "owa" auth mode (cookie session + internal API).
  owaUrl: "https://outlook.office365.com",

  // The address book that Directorium registers for the EWS/Basic backends.
  // OWA accounts each register their own book named after the account's org
  // (see the accounts model below) rather than using this single name.
  addressBookName: "Exchange GAL",

  // ---- OAuth (Office 365) ----
  // Default to the public Microsoft Office desktop client id. It is a
  // pre-consented first-party app, so signing in does NOT require a tenant
  // admin to approve a new application. This is the "works like Owl" path.
  clientId: "d3590ed6-52b3-4102-aeff-aad2292ab01c",
  // "common" works for any tenant. "organizations" excludes personal accounts.
  tenant: "common",
  // EWS delegated scope + offline_access so we receive a refresh token.
  scopes: "https://outlook.office365.com/EWS.AccessAsUser.All offline_access",

  // Saved email (used for Autodiscover and username pre-fill).
  email: "",

  // ---- Basic auth (on-premise Exchange) ----
  username: "",
  // Stored locally only; never transmitted anywhere except the EWS server.
  password: "",
};

const CONFIG_KEY = "config";
const TOKEN_KEY = "tokens";
const ACCOUNTS_KEY = "accounts";
// Synchronously-readable mirror of the book list. The MV3 background is a
// background *page* (not a service worker), so localStorage is available and is
// shared across all extension pages (background + options) on the same origin.
// We need the book names synchronously at background startup to register the
// address-book providers before the first onSearchRequest fires.
const BOOKS_MIRROR_KEY = "directorium_books";

/** @returns {Promise<typeof DEFAULT_CONFIG>} */
export async function getConfig() {
  const stored = await messenger.storage.local.get(CONFIG_KEY);
  return { ...DEFAULT_CONFIG, ...(stored[CONFIG_KEY] || {}) };
}

/** @param {Partial<typeof DEFAULT_CONFIG>} patch */
export async function setConfig(patch) {
  const current = await getConfig();
  const next = { ...current, ...patch };
  await messenger.storage.local.set({ [CONFIG_KEY]: next });
  // The EWS book name / authMode affect the desired book set.
  await refreshBooksMirror();
  return next;
}

/**
 * @typedef {Object} TokenSet
 * @property {string} accessToken
 * @property {string} [refreshToken]
 * @property {number} expiresAt  Epoch ms when the access token expires.
 */

/** @returns {Promise<TokenSet|null>} */
export async function getTokens() {
  const stored = await messenger.storage.local.get(TOKEN_KEY);
  return stored[TOKEN_KEY] || null;
}

/** @param {TokenSet|null} tokens */
export async function setTokens(tokens) {
  if (tokens === null) {
    await messenger.storage.local.remove(TOKEN_KEY);
  } else {
    await messenger.storage.local.set({ [TOKEN_KEY]: tokens });
  }
}

// ---------------------------------------------------------------------------
// OWA accounts (multi-account)
// ---------------------------------------------------------------------------
// Each OWA sign-in becomes its own "account" with its own captured Bearer token
// and its own Thunderbird address book, named after the account's organization
// (derived from the signed-in email's domain, and user-editable). Signing into
// a second mailbox adds a second account/book alongside the first.
//
// @typedef {Object} OwaAccount
// @property {string} id        Stable internal id.
// @property {string} label     Address-book name shown in Thunderbird.
// @property {string|null} user    "<tid>/<oid>" from the token: the account key.
// @property {string|null} email   Mailbox email from the token (for labels/UI).
// @property {string|null} anchor  X-AnchorMailbox used for routing (email,
//                                 "OID:<oid>@<tid>" or "PUID:<puid>@<tid>").
// @property {string|null} tenant  Azure tenant id (tid claim).
// @property {string} owaUrl    OWA origin this account lives on.
// @property {{auth:string, anchor:string|null, ts:number}|null} token

/** Decode a JWT access token's payload, or null. */
function parseJwt(authValue) {
  try {
    const jwt = String(authValue).replace(/^Bearer\s+/i, "");
    return JSON.parse(
      atob(jwt.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );
  } catch {
    return null;
  }
}

const isEmail = (s) => /^[^\s@:]+@[^\s@]+\.[^\s@]+$/.test(s || "");

/**
 * Who a captured token belongs to, from its JWT claims. `user` (tenant + object
 * id) is the stable account key: the new Outlook sends X-AnchorMailbox in
 * several formats for the same mailbox ("OID:<oid>@<tid>", "PUID:<puid>@<tid>",
 * an SMTP address), so the anchor alone can't identify an account.
 */
function tokenIdentity(authValue) {
  const p = parseJwt(authValue) || {};
  const email =
    [p.upn, p.preferred_username, p.email, p.smtp, p.unique_name].find(isEmail) || "";
  return {
    user: p.tid && p.oid ? `${p.tid}/${p.oid}`.toLowerCase() : null,
    tenant: p.tid || null,
    oid: p.oid || null,
    puid: p.puid || null,
    email,
  };
}

/** Email in an X-AnchorMailbox value ("user@x.com" or "SMTP:user@x.com"), or "". */
function anchorEmail(anchor) {
  const a = String(anchor || "").replace(/^smtp:/i, "");
  return isEmail(a) ? a : "";
}

/**
 * Whether an X-AnchorMailbox value points at the token's own mailbox (and not,
 * say, a shared mailbox the user has open). Unknown formats count as a match
 * only when the token carries no identity to compare against.
 */
function anchorIsOwn(anchor, id) {
  if (!anchor) return false;
  const m = /^(oid|puid):([^@]+)@/i.exec(anchor);
  if (m) {
    const want = m[1].toLowerCase() === "oid" ? id.oid : id.puid;
    return !want || m[2].toLowerCase() === String(want).toLowerCase();
  }
  const email = anchorEmail(anchor);
  return !id.email || !email || email.toLowerCase() === id.email.toLowerCase();
}

/**
 * Default address-book label for an account: the email domain (e.g.
 * "contoso.com"), trimmed of the noisy ".onmicrosoft.com" suffix. Falls back to
 * "Exchange GAL" if no email can be determined.
 */
function labelForEmail(email) {
  if (!email) return "Exchange GAL";
  const domain = (email.split("@")[1] || "").replace(/\.onmicrosoft\.com$/i, "");
  return domain || "Exchange GAL";
}

/** Labels assigned automatically when no email was known (pre-0.3.1 fallback). */
const isFallbackLabel = (label) => /^Exchange GAL( \(\d+\))?$/.test(label || "");

/** `label`, or a variant not used by any of `others` (email, then "(2)", "(3)"…). */
function uniqueLabel(label, email, others) {
  const taken = new Set(others.map((a) => a.label));
  if (!taken.has(label)) return label;
  if (email && !taken.has(email)) return email;
  for (let n = 2; ; n++) {
    if (!taken.has(`${label} (${n})`)) return `${label} (${n})`;
  }
}

/** @returns {Promise<OwaAccount[]>} */
export async function getAccounts() {
  const stored = await messenger.storage.local.get(ACCOUNTS_KEY);
  return stored[ACCOUNTS_KEY] || [];
}

async function writeAccounts(accounts) {
  await messenger.storage.local.set({ [ACCOUNTS_KEY]: accounts });
  await refreshBooksMirror(accounts);
  return accounts;
}

/**
 * Recompute the desired book set (one per OWA account, plus one for the EWS
 * backend when a non-OWA mode is active) and write the synchronous mirror.
 * @returns {Promise<Array<{id:string, label:string, kind:"owa"|"ews"}>>}
 */
export async function refreshBooksMirror(accounts) {
  const list = accounts || (await getAccounts());
  const cfg = await getConfig();
  const books = list.map((a) => ({ id: a.id, label: a.label, kind: "owa" }));
  if (cfg.authMode !== "owa") {
    books.push({ id: "ews", label: cfg.addressBookName || "Exchange GAL", kind: "ews" });
  }
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(BOOKS_MIRROR_KEY, JSON.stringify(books));
    }
  } catch {
    /* localStorage unavailable — reconcile from storage.local will still run */
  }
  return books;
}

/** Synchronous read of the book mirror, for top-level startup registration. */
export function readBooksMirror() {
  try {
    if (typeof localStorage !== "undefined") {
      return JSON.parse(localStorage.getItem(BOOKS_MIRROR_KEY) || "[]");
    }
  } catch {
    /* fall through */
  }
  return [];
}

/** The account a token identity belongs to: by user key, else by email/anchor. */
function findAccountFor(accounts, id, anchor) {
  const lc = (s) => String(s || "").toLowerCase();
  if (id.user) {
    const byUser = accounts.find((a) => a.user === id.user);
    if (byUser) return byUser;
  }
  const email = lc(id.email || anchorEmail(anchor));
  return (
    (email && accounts.find((a) => lc(a.email) === email || lc(anchorEmail(a.anchor)) === email)) ||
    (!id.user && anchor && accounts.find((a) => lc(a.anchor) === lc(anchor))) ||
    null
  );
}

/**
 * Insert or update the OWA account that owns a freshly captured token, matched
 * by the token's user (tenant + object id), falling back to its email. Returns
 * the account and whether it was newly created (so the caller can register a
 * new provider).
 * @returns {Promise<{account: OwaAccount, isNew: boolean}>}
 */
export async function upsertOwaAccount(authValue, anchor) {
  const accounts = await getAccounts();
  const id = tokenIdentity(authValue);
  let account = findAccountFor(accounts, id, anchor);

  // Route with the captured anchor only if it is this user's own mailbox;
  // otherwise keep what we had, or build one from the token.
  const ownAnchor =
    (anchorIsOwn(anchor, id) && anchor) ||
    account?.anchor ||
    id.email ||
    (id.oid && id.tenant ? `OID:${id.oid}@${id.tenant}` : null);
  const token = { auth: authValue, anchor: ownAnchor, ts: Date.now() };

  let isNew = false;
  if (account) {
    account.token = token;
    account.anchor = ownAnchor;
    if (id.user) account.user = id.user;
    if (id.email) account.email = id.email;
    if (id.tenant) account.tenant = id.tenant;
  } else {
    isNew = true;
    const email = id.email || anchorEmail(anchor);
    account = {
      id: `acct_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`,
      // Two mailboxes in the same tenant would share a label: fall back to the
      // full email address.
      label: uniqueLabel(labelForEmail(email), email, accounts),
      user: id.user,
      email: email || null,
      anchor: ownAnchor,
      tenant: id.tenant,
      owaUrl: (await getConfig()).owaUrl,
      token,
    };
    accounts.push(account);
  }
  await writeAccounts(accounts);
  return { account, isNew };
}

/**
 * Repair accounts created by earlier versions, which keyed accounts by the raw
 * X-AnchorMailbox: fill in user/email from each stored token, merge duplicates
 * of the same user (keeping the newest token), and replace "Exchange GAL"
 * fallback labels with the organization's domain. Runs once per startup.
 * @returns {Promise<boolean>} whether anything changed
 */
export async function normalizeAccounts() {
  const accounts = await getAccounts();
  const merged = [];
  let changed = false;

  for (const a of accounts) {
    const id = a.token?.auth ? tokenIdentity(a.token.auth) : null;
    if (id?.user && a.user !== id.user) { a.user = id.user; changed = true; }
    if (id?.email && a.email !== id.email) { a.email = id.email; changed = true; }

    const twin = a.user && merged.find((m) => m.user === a.user);
    if (twin) {
      if ((a.token?.ts || 0) > (twin.token?.ts || 0)) twin.token = a.token;
      if (isFallbackLabel(twin.label) && !isFallbackLabel(a.label)) twin.label = a.label;
      changed = true;
      continue;
    }
    merged.push(a);
  }

  for (const a of merged) {
    if (isFallbackLabel(a.label) && a.email) {
      const others = merged.filter((m) => m !== a);
      a.label = uniqueLabel(labelForEmail(a.email), a.email, others);
      changed = true;
    }
  }

  if (changed) await writeAccounts(merged);
  return changed;
}

/** Patch the in-memory token of an existing account (after a silent refresh). */
export async function setAccountToken(id, token) {
  const accounts = await getAccounts();
  const account = accounts.find((a) => a.id === id);
  if (!account) return null;
  account.token = token;
  await writeAccounts(accounts);
  return account;
}

/** Rename an account's address book. Takes effect after a Thunderbird restart. */
export async function updateAccountLabel(id, label) {
  const accounts = await getAccounts();
  const account = accounts.find((a) => a.id === id);
  if (!account) return null;
  account.label = (label || "").trim() || account.label;
  await writeAccounts(accounts);
  return account;
}

/** Remove an account (and, by reconcile, its address-book provider). */
export async function removeAccount(id) {
  const accounts = (await getAccounts()).filter((a) => a.id !== id);
  await writeAccounts(accounts);
  return accounts;
}
