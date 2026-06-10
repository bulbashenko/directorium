// Dispatches GAL operations to the backend selected by authMode:
//   "owa"  → OWA internal JSON API (works like Owl on locked tenants)
//   else   → EWS SOAP (OAuth or Basic)

import { getConfig } from "./lib/storage.js";
import * as ews from "./ews/ewsClient.js";
import * as owa from "./owa/owaClient.js";

async function backend() {
  const cfg = await getConfig();
  return cfg.authMode === "owa" ? owa : ews;
}

export async function searchGal(query) {
  return (await backend()).searchGal(query);
}

export async function enumerateGal(onBatch) {
  return (await backend()).enumerateGal(onBatch);
}

export async function testConnection() {
  return (await backend()).testConnection();
}
