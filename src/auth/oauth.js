// OAuth2 for Office 365 using the Device Code flow.
//
// Why device code instead of the usual auth-code/launchWebAuthFlow:
//   * It works with a *borrowed* public first-party client_id (e.g. the
//     Microsoft Office desktop app). Such apps are already trusted in the
//     tenant, so no Azure admin has to approve a brand-new application.
//   * It needs no redirect_uri. launchWebAuthFlow can only catch a redirect
//     back to the extension's own URL, which a borrowed client_id will never
//     be registered for. Device code sidesteps that entirely.
//
// Flow:
//   1. POST /devicecode  -> user_code + verification_uri + device_code
//   2. User opens https://microsoft.com/devicelogin, types the code, signs in.
//   3. We poll /token until the user finishes -> access_token + refresh_token.
//   4. Later we silently refresh with the refresh_token.

import { getConfig, getTokens, setTokens } from "../lib/storage.js";

const AUTHORITY = "https://login.microsoftonline.com";

function endpoints(tenant) {
  const base = `${AUTHORITY}/${encodeURIComponent(tenant)}/oauth2/v2.0`;
  return {
    deviceCode: `${base}/devicecode`,
    token: `${base}/token`,
  };
}

/**
 * Step 1: ask Microsoft for a device + user code.
 * @returns {Promise<{deviceCode:string,userCode:string,verificationUri:string,interval:number,expiresIn:number,message:string}>}
 */
export async function startDeviceCode() {
  const cfg = await getConfig();
  const { deviceCode } = endpoints(cfg.tenant);

  const body = new URLSearchParams({
    client_id: cfg.clientId,
    scope: cfg.scopes,
  });

  const resp = await fetch(deviceCode, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await resp.json();
  if (!resp.ok) {
    throw new Error(
      `Device code request failed: ${json.error || resp.status} ${json.error_description || ""}`
    );
  }

  return {
    deviceCode: json.device_code,
    userCode: json.user_code,
    verificationUri: json.verification_uri,
    interval: json.interval || 5,
    expiresIn: json.expires_in || 900,
    message: json.message,
  };
}

/**
 * Step 3: poll the token endpoint until the user completes sign-in.
 * Resolves with a stored TokenSet, or throws on timeout/denial.
 * @param {string} deviceCode
 * @param {number} interval seconds between polls
 * @param {number} expiresIn seconds before the device code dies
 */
export async function pollForToken(deviceCode, interval, expiresIn) {
  const cfg = await getConfig();
  const { token } = endpoints(cfg.tenant);
  const deadline = Date.now() + expiresIn * 1000;
  let waitMs = interval * 1000;

  while (Date.now() < deadline) {
    await sleep(waitMs);

    const body = new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: cfg.clientId,
      device_code: deviceCode,
    });

    const resp = await fetch(token, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });
    const json = await resp.json();

    if (resp.ok) {
      return persistTokenResponse(json);
    }

    switch (json.error) {
      case "authorization_pending":
        break; // keep polling at the normal interval
      case "slow_down":
        waitMs += 5000; // Microsoft asked us to back off
        break;
      case "expired_token":
      case "authorization_declined":
      case "access_denied":
        throw new Error(`Sign-in not completed: ${json.error}`);
      default:
        throw new Error(
          `Token poll failed: ${json.error} ${json.error_description || ""}`
        );
    }
  }
  throw new Error("Device code expired before sign-in completed.");
}

/**
 * Return a valid access token, refreshing silently if needed.
 * @returns {Promise<string>}
 */
export async function getAccessToken() {
  const tokens = await getTokens();
  if (!tokens) {
    throw new Error("Not signed in. Open Directorium settings and sign in.");
  }
  // Refresh a minute early to avoid edge-of-expiry failures.
  if (Date.now() < tokens.expiresAt - 60_000) {
    return tokens.accessToken;
  }
  if (!tokens.refreshToken) {
    throw new Error("Session expired and no refresh token. Please sign in again.");
  }
  const refreshed = await refresh(tokens.refreshToken);
  return refreshed.accessToken;
}

/** @param {string} refreshToken */
async function refresh(refreshToken) {
  const cfg = await getConfig();
  const { token } = endpoints(cfg.tenant);

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: cfg.clientId,
    refresh_token: refreshToken,
    scope: cfg.scopes,
  });

  const resp = await fetch(token, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const json = await resp.json();
  if (!resp.ok) {
    await setTokens(null); // force a fresh interactive sign-in next time
    throw new Error(`Token refresh failed: ${json.error || resp.status}`);
  }
  return persistTokenResponse(json, refreshToken);
}

/** Normalise a Microsoft token response and store it. */
async function persistTokenResponse(json, fallbackRefresh) {
  const tokenSet = {
    accessToken: json.access_token,
    refreshToken: json.refresh_token || fallbackRefresh,
    expiresAt: Date.now() + (json.expires_in || 3600) * 1000,
  };
  await setTokens(tokenSet);
  return tokenSet;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
