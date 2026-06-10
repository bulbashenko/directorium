// High-level EWS client: applies auth, posts SOAP, returns parsed contacts.

import { getConfig } from "../lib/storage.js";
import { getAccessToken } from "../auth/oauth.js";
import {
  resolveNamesRequest,
  findPeopleRequest,
  parseResponse,
  NS,
} from "./soap.js";
import { resolutionToContact, personaToContact } from "../contacts/mapper.js";

/** Build the Authorization header for the configured auth mode. */
async function authHeader(cfg) {
  if (cfg.authMode === "oauth") {
    const token = await getAccessToken();
    return `Bearer ${token}`;
  }
  if (!cfg.username) {
    throw new Error("Basic auth selected but no username is configured.");
  }
  const raw = `${cfg.username}:${cfg.password}`;
  const b64 = btoa(unescape(encodeURIComponent(raw)));
  return `Basic ${b64}`;
}

/**
 * POST a SOAP envelope to the EWS endpoint and return the parsed Document.
 * Also returns the raw XML for diagnostics.
 */
async function post(soapBody, soapAction) {
  const cfg = await getConfig();
  const resp = await fetch(cfg.ewsUrl, {
    method: "POST",
    headers: {
      "Content-Type": "text/xml; charset=utf-8",
      Authorization: await authHeader(cfg),
      SOAPAction: `"http://schemas.microsoft.com/exchange/services/2006/messages/${soapAction}"`,
    },
    body: soapBody,
  });

  const xmlText = await resp.text();
  console.debug(`[Directorium] EWS ${soapAction} HTTP ${resp.status}:\n${xmlText.slice(0, 2000)}`);

  if (resp.status === 401) {
    throw new Error("EWS returned 401 Unauthorized — check your username/password or sign in again.");
  }
  if (resp.status === 403) {
    throw new Error("EWS returned 403 Forbidden — your account may not have EWS access. Contact your IT admin.");
  }
  if (!resp.ok) {
    throw new Error(`EWS HTTP ${resp.status}: ${xmlText.slice(0, 300)}`);
  }
  return { doc: parseResponse(xmlText), rawXml: xmlText };
}

/**
 * Search the GAL for a term (incremental search as the user types).
 * @param {string} query
 * @returns {Promise<{contacts: Array<Record<string,string>>, rawXml: string}>}
 */
export async function searchGal(query) {
  if (!query || query.trim().length < 1) return { contacts: [], rawXml: "" };
  const cfg = await getConfig();
  const { doc, rawXml } = await post(
    resolveNamesRequest(cfg.exchangeVersion, query.trim()),
    "ResolveNames"
  );

  assertResponseSuccess(doc, "ResolveNamesResponseMessage", {
    allowNoResults: true,
  });

  const contacts = [];
  for (const res of doc.getElementsByTagNameNS(NS.t, "Resolution")) {
    const c = resolutionToContact(res);
    if (c) contacts.push(c);
  }

  console.debug(`[Directorium] ResolveNames "${query}": ${contacts.length} contacts`, contacts);
  return { contacts, rawXml };
}

/**
 * Enumerate the entire GAL.
 *
 * First tries FindPeople with the "directory" folder (Exchange 2013+).
 * If that returns 0 results (common on Exchange Online due to tenant policy),
 * falls back to alphabet enumeration: queries ResolveNames for every letter
 * a–z plus digits 0–9, deduplicates by email. Each prefix yields up to 100
 * results, so this reliably covers large directories.
 */
export async function enumerateGal(onBatch) {
  // Try FindPeople first.
  try {
    const result = await findPeopleGal(onBatch);
    if (result.contacts.length > 0) return result;
    console.debug("[Directorium] FindPeople returned 0 — falling back to alphabet enumeration");
  } catch (err) {
    console.debug("[Directorium] FindPeople failed, falling back to alphabet enumeration:", err.message);
  }

  return alphabetEnumerateGal(onBatch);
}

async function findPeopleGal(onBatch) {
  const cfg = await getConfig();
  const pageSize = 100;
  let offset = 0;
  const all = [];
  let lastRawXml = "";

  for (;;) {
    const { doc, rawXml } = await post(
      findPeopleRequest(cfg.exchangeVersion, "", offset, pageSize),
      "FindPeople"
    );
    lastRawXml = rawXml;

    const faultMsg = doc.getElementsByTagNameNS(NS.m, "FindPeopleResponseMessage")[0];
    if (faultMsg?.getAttribute("ResponseClass") === "Error") {
      const code = faultMsg.getElementsByTagNameNS(NS.m, "ResponseCode")[0]?.textContent;
      const txt  = faultMsg.getElementsByTagNameNS(NS.m, "MessageText")[0]?.textContent;
      throw new EwsError(`FindPeople (${code}): ${txt}`, rawXml);
    }

    const personas = doc.getElementsByTagNameNS(NS.t, "Persona");
    const batch = [];
    for (const p of personas) {
      const c = personaToContact(p);
      if (c) batch.push(c);
    }
    all.push(...batch);
    if (onBatch) onBatch(batch, all.length);

    if (personas.length < pageSize) break;
    offset += pageSize;
    if (offset > 50000) break;
  }
  return { contacts: all, rawXml: lastRawXml };
}

/**
 * Enumerate the GAL by querying every letter/digit prefix via ResolveNames.
 * Deduplicates by PrimaryEmail so overlapping results don't inflate the count.
 */
async function alphabetEnumerateGal(onBatch) {
  const prefixes = "abcdefghijklmnopqrstuvwxyz0123456789".split("");
  const seen = new Set();
  const all = [];
  let lastRawXml = "";

  for (const prefix of prefixes) {
    try {
      const { contacts, rawXml } = await searchGal(prefix);
      lastRawXml = rawXml;
      const batch = [];
      for (const c of contacts) {
        const key = c.PrimaryEmail?.toLowerCase() || c.DisplayName?.toLowerCase();
        if (key && !seen.has(key)) {
          seen.add(key);
          batch.push(c);
        }
      }
      all.push(...batch);
      if (onBatch && batch.length > 0) onBatch(batch, all.length);
    } catch (err) {
      console.warn(`[Directorium] alphabet enum prefix "${prefix}" failed:`, err.message);
    }
  }
  return { contacts: all, rawXml: lastRawXml };
}

/** Lightweight connectivity + auth check. Returns raw XML for diagnostics. */
export async function testConnection() {
  const { contacts, rawXml } = await searchGal("a");
  return { ok: true, contactCount: contacts.length, rawXml };
}

function assertResponseSuccess(doc, messageLocalName, { allowNoResults } = {}) {
  const msg = doc.getElementsByTagNameNS(NS.m, messageLocalName)[0];
  if (!msg) return;
  const responseClass = msg.getAttribute("ResponseClass");
  if (responseClass === "Success") return;

  const code = doc.getElementsByTagNameNS(NS.m, "ResponseCode")[0]?.textContent;
  if (allowNoResults && (code === "ErrorNameResolutionNoResults" || code === "ErrorNameResolutionMultipleResults")) return;

  const text = doc.getElementsByTagNameNS(NS.m, "MessageText")[0]?.textContent;
  throw new Error(`EWS error (${code || responseClass}): ${text || "no detail"}`);
}

export class EwsError extends Error {
  constructor(message, rawXml) {
    super(message);
    this.rawXml = rawXml;
  }
}
