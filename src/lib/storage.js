// Centralised access to extension settings and cached OAuth tokens.
// Settings live in storage.local so they survive restarts but never sync to
// any Mozilla account (tokens must stay on-device).

/** @typedef {"oauth" | "basic"} AuthMode */

export const DEFAULT_CONFIG = {
  // "oauth-interactive" = interactive sign-in with your own (verified or
  // admin-approved) Azure app — the realistic Office 365 path on managed
  // tenants. "oauth" (device code) works on lenient tenants without an app.
  authMode: /** @type {AuthMode} */ ("oauth-interactive"),

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
// @property {string|null} anchor  X-AnchorMailbox / UPN (the mailbox email).
// @property {string|null} tenant  Azure tenant id (tid claim), for matching.
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

/** Best-effort mailbox email for a captured token (anchor or a JWT claim). */
function emailFromToken(authValue, anchor) {
  const p = parseJwt(authValue) || {};
  const cand =
    anchor || p.upn || p.unique_name || p.preferred_username || p.email || "";
  return /\S+@\S+\.\S+/.test(cand) ? cand : "";
}

/** Azure tenant id from the token, or null. */
function tenantFromToken(authValue) {
  return parseJwt(authValue)?.tid || null;
}

/**
 * Default address-book label for a newly captured account: the email domain
 * (e.g. "contoso.com"), trimmed of the noisy ".onmicrosoft.com" suffix. Falls
 * back to "Exchange GAL" if no email can be determined.
 */
export function deriveAccountLabel(authValue, anchor) {
  const email = emailFromToken(authValue, anchor);
  if (!email) return "Exchange GAL";
  const domain = (email.split("@")[1] || email).replace(/\.onmicrosoft\.com$/i, "");
  return domain || "Exchange GAL";
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

/**
 * Insert or update the OWA account that owns a freshly captured token. Matched
 * by anchor mailbox (preferred) or tenant id. Returns the account and whether
 * it was newly created (so the caller can register a new provider).
 * @returns {Promise<{account: OwaAccount, isNew: boolean}>}
 */
export async function upsertOwaAccount(authValue, anchor) {
  const accounts = await getAccounts();
  const email = emailFromToken(authValue, anchor);
  const key = (anchor || email || "").toLowerCase();
  const tenant = tenantFromToken(authValue);

  let account =
    (key && accounts.find((a) => (a.anchor || "").toLowerCase() === key)) ||
    (!key && tenant && accounts.find((a) => a.tenant === tenant)) ||
    null;

  const token = {
    auth: authValue,
    anchor: anchor || account?.token?.anchor || account?.anchor || null,
    ts: Date.now(),
  };

  let isNew = false;
  if (account) {
    account.token = token;
    if (!account.anchor && email) account.anchor = email;
    if (!account.tenant && tenant) account.tenant = tenant;
  } else {
    isNew = true;
    let label = deriveAccountLabel(authValue, anchor);
    // Disambiguate if another account already claims this label (e.g. two
    // mailboxes in the same tenant): fall back to the full email address.
    if (accounts.some((a) => a.label === label)) label = email || `${label} (2)`;
    account = {
      id: `acct_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`,
      label,
      anchor: anchor || email || null,
      tenant,
      owaUrl: (await getConfig()).owaUrl,
      token,
    };
    accounts.push(account);
  }
  await writeAccounts(accounts);
  return { account, isNew };
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
