// OAuth2 Authorization Code flow with PKCE, via an interactive browser window.
//
// This is the "works like Owl" path for tenants whose Conditional Access blocks
// the device code flow (a common control). An interactive sign-in window passes
// those policies because it's an ordinary browser login.
//
// Requirement: the redirect_uri must be registered on the Azure app. A borrowed
// first-party client_id won't have our extension's redirect URL, so this flow
// needs your OWN (free) multi-tenant app registration. See README for the
// 5-minute setup. The redirect URL to register is shown in Directorium settings
// (browser.identity.getRedirectURL()).

import { getConfig, setTokens } from "../lib/storage.js";

const AUTHORITY = "https://login.microsoftonline.com";

/** The redirect URL Thunderbird expects launchWebAuthFlow to land on. */
export function getRedirectUrl() {
  return browser.identity.getRedirectURL();
}

// ---- PKCE helpers ----
function base64UrlEncode(bytes) {
  let str = "";
  for (const b of new Uint8Array(bytes)) str += String.fromCharCode(b);
  return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function randomVerifier() {
  const bytes = new Uint8Array(64);
  crypto.getRandomValues(bytes);
  return base64UrlEncode(bytes);
}

async function challengeFrom(verifier) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return base64UrlEncode(digest);
}

/**
 * Run the interactive sign-in. Resolves once tokens are stored.
 * @returns {Promise<{ok: true}>}
 */
export async function signInInteractive() {
  const cfg = await getConfig();
  const redirectUri = getRedirectUrl();
  const verifier = randomVerifier();
  const challenge = await challengeFrom(verifier);
  const state = randomVerifier().slice(0, 16);

  const authUrl =
    `${AUTHORITY}/${encodeURIComponent(cfg.tenant)}/oauth2/v2.0/authorize?` +
    new URLSearchParams({
      client_id: cfg.clientId,
      response_type: "code",
      redirect_uri: redirectUri,
      response_mode: "query",
      scope: cfg.scopes,
      code_challenge: challenge,
      code_challenge_method: "S256",
      state,
      prompt: "select_account",
    }).toString();

  // Opens a window, lets the user sign in, resolves with the redirect URL.
  const redirectResponse = await browser.identity.launchWebAuthFlow({
    url: authUrl,
    interactive: true,
  });

  const url = new URL(redirectResponse);
  const returnedState = url.searchParams.get("state");
  const code = url.searchParams.get("code");
  const error = url.searchParams.get("error");

  if (error) {
    throw new Error(
      `${error}: ${url.searchParams.get("error_description") || "sign-in failed"}`
    );
  }
  if (!code) throw new Error("No authorization code returned.");
  if (returnedState !== state) throw new Error("State mismatch — aborting for safety.");

  await exchangeCode(cfg, code, verifier, redirectUri);
  return { ok: true };
}

async function exchangeCode(cfg, code, verifier, redirectUri) {
  const tokenUrl = `${AUTHORITY}/${encodeURIComponent(cfg.tenant)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: cfg.clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
    scope: cfg.scopes,
  });

  const resp = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const json = await resp.json();
  if (!resp.ok) {
    throw new Error(
      `Token exchange failed: ${json.error} ${json.error_description || ""}`
    );
  }

  await setTokens({
    accessToken: json.access_token,
    refreshToken: json.refresh_token,
    expiresAt: Date.now() + (json.expires_in || 3600) * 1000,
  });
}
