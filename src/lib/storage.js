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

  // The address book that Directorium registers inside Thunderbird.
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
