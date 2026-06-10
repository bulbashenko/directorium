// Exchange Autodiscover: finds the EWS URL for any email address.
//
// Tries in order:
//   1. Autodiscover v2 (REST/JSON) via Office 365 endpoint — no auth needed,
//      covers any O365-hosted domain including university tenants.
//   2. Autodiscover v2 at autodiscover.{domain} — covers on-prem Exchange 2013+.
//   3. Autodiscover v2 at {domain} directly — some orgs skip the subdomain.
//   4. Autodiscover v1 (SOAP/XML) at autodiscover.{domain} — Exchange 2010/2007
//      fallback; returns EwsUrl without requiring authentication.

const TIMEOUT_MS = 6000;

/**
 * Discover the EWS URL for the given email address.
 * @param {string} email  full email, e.g. "user@uniba.sk"
 * @returns {Promise<{ewsUrl: string, method: string}>}
 * @throws if no endpoint could be found
 */
export async function discoverEwsUrl(email) {
  const domain = email.split("@")[1]?.toLowerCase();
  if (!domain) throw new Error("Invalid email address.");

  // 1. O365 redirect service — works for any domain hosted on Exchange Online.
  try {
    const url = `https://outlook.office365.com/autodiscover/autodiscover.json?Email=${encodeURIComponent(email)}&Protocol=EWS`;
    const ews = await fetchJsonV2(url);
    if (ews) return { ewsUrl: ews, method: "autodiscover-v2-o365" };
  } catch (_) {/* try next */}

  // 2. Autodiscover v2 at the standard subdomain.
  try {
    const url = `https://autodiscover.${domain}/autodiscover/autodiscover.json?Email=${encodeURIComponent(email)}&Protocol=EWS`;
    const ews = await fetchJsonV2(url);
    if (ews) return { ewsUrl: ews, method: "autodiscover-v2-subdomain" };
  } catch (_) {/* try next */}

  // 3. Autodiscover v2 at the apex domain.
  try {
    const url = `https://${domain}/autodiscover/autodiscover.json?Email=${encodeURIComponent(email)}&Protocol=EWS`;
    const ews = await fetchJsonV2(url);
    if (ews) return { ewsUrl: ews, method: "autodiscover-v2-apex" };
  } catch (_) {/* try next */}

  // 4. Autodiscover v1 SOAP — unauthenticated POST returns the EwsUrl in XML.
  try {
    const url = `https://autodiscover.${domain}/autodiscover/autodiscover.xml`;
    const ews = await fetchSoapV1(url, email);
    if (ews) return { ewsUrl: ews, method: "autodiscover-v1" };
  } catch (_) {/* try next */}

  throw new Error(
    `Could not auto-detect the EWS endpoint for "${domain}". ` +
    `Try entering the URL manually (ask your IT helpdesk for the EWS address).`
  );
}

/** Fetch an Autodiscover v2 JSON endpoint and return the EWS URL or null. */
async function fetchJsonV2(url) {
  const resp = await fetchWithTimeout(url, {
    headers: { Accept: "application/json" },
  });
  if (!resp.ok) return null;
  const json = await resp.json();
  // Microsoft's response: { "Protocol": "EWS", "Url": "https://..." }
  if (json?.Protocol === "EWS" && json?.Url) return json.Url;
  // Some servers return an array.
  if (Array.isArray(json)) {
    const entry = json.find((e) => e.Protocol === "EWS");
    if (entry?.Url) return entry.Url;
  }
  return null;
}

/** POST an Autodiscover v1 SOAP envelope and parse the EwsUrl out of the XML. */
async function fetchSoapV1(url, email) {
  const body = `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
  <Request>
    <EMailAddress>${email}</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
  </Request>
</Autodiscover>`;

  const resp = await fetchWithTimeout(url, {
    method: "POST",
    headers: {
      "Content-Type": "text/xml; charset=utf-8",
    },
    body,
  });

  if (!resp.ok) return null;
  const xml = await resp.text();
  const doc = new DOMParser().parseFromString(xml, "text/xml");

  // The EWS URL lives inside <EwsUrl> inside a <Protocol> block.
  const ewsEl = doc.querySelector("EwsUrl");
  return ewsEl?.textContent?.trim() || null;
}

function fetchWithTimeout(url, opts = {}) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  return fetch(url, { ...opts, signal: ctrl.signal }).finally(() =>
    clearTimeout(timer)
  );
}
