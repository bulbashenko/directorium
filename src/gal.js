// Dispatches GAL operations to the backend selected by authMode:
//   "owa"  → OWA internal JSON API (works like Owl on locked tenants),
//            multi-account: each account is queried with its own token
//   else   → EWS SOAP (OAuth or Basic), single account
//
// When an `account` is supplied (e.g. by a per-account address-book provider),
// the OWA backend uses exactly that account. Otherwise we fall back to the
// first OWA account (for one-off diagnostics / RPC callers).

import { getConfig, getAccounts } from "./lib/storage.js";
import * as ews from "./ews/ewsClient.js";
import * as owa from "./owa/owaClient.js";

async function firstOwaAccount() {
  const accounts = await getAccounts();
  if (!accounts.length) {
    throw new Error("No OWA account yet — click Sign in (OWA) and keep the Outlook tab open.");
  }
  return accounts[0];
}

export async function searchGal(query, account) {
  const cfg = await getConfig();
  if (cfg.authMode === "owa") {
    return owa.searchGal(query, account || (await firstOwaAccount()));
  }
  return ews.searchGal(query);
}

export async function enumerateGal(account, onBatch) {
  const cfg = await getConfig();
  if (cfg.authMode === "owa") {
    return owa.enumerateGal(account || (await firstOwaAccount()), onBatch);
  }
  return ews.enumerateGal(onBatch);
}

export async function testConnection(account) {
  const cfg = await getConfig();
  if (cfg.authMode === "owa") {
    return owa.testConnection(account || (await firstOwaAccount()));
  }
  return ews.testConnection();
}
