// Builders for the EWS SOAP envelopes we need, plus small XML helpers.
// EWS is SOAP 1.1 / XML over HTTPS. We hand-build the few requests we use
// rather than pull in a SOAP library — it keeps the extension tiny and the
// requests are short and stable.

const NS = {
  soap: "http://schemas.xmlsoap.org/soap/envelope/",
  t: "http://schemas.microsoft.com/exchange/services/2006/types",
  m: "http://schemas.microsoft.com/exchange/services/2006/messages",
};

/** Escape a string for safe inclusion in XML text/attribute content. */
export function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function envelope(version, bodyXml) {
  return (
    `<?xml version="1.0" encoding="utf-8"?>` +
    `<soap:Envelope xmlns:soap="${NS.soap}" xmlns:t="${NS.t}" xmlns:m="${NS.m}">` +
    `<soap:Header><t:RequestServerVersion Version="${xmlEscape(version)}"/></soap:Header>` +
    `<soap:Body>${bodyXml}</soap:Body>` +
    `</soap:Envelope>`
  );
}

/**
 * ResolveNames against the GAL (Active Directory). Best for incremental search
 * as the user types — fuzzy matches names, aliases and email addresses.
 * Capped by Exchange at 100 results.
 * @param {string} version  e.g. "Exchange2013_SP1"
 * @param {string} query    the text the user is searching for
 */
export function resolveNamesRequest(version, query) {
  const body =
    `<m:ResolveNames ReturnFullContactData="true" SearchScope="ActiveDirectory">` +
    `<m:UnresolvedEntry>${xmlEscape(query)}</m:UnresolvedEntry>` +
    `</m:ResolveNames>`;
  return envelope(version, body);
}

/**
 * FindPeople over the directory folder — supports paging, used to enumerate
 * the entire GAL (phase 2). An empty query string returns everything.
 * @param {string} version
 * @param {string} query     optional filter; "" enumerates all
 * @param {number} offset    paging offset
 * @param {number} pageSize  max entries to return (Exchange caps near 100)
 */
export function findPeopleRequest(version, query, offset, pageSize) {
  const body =
    `<m:FindPeople>` +
    `<m:IndexedPageItemView MaxEntriesReturned="${pageSize}" BasePoint="Beginning" Offset="${offset}"/>` +
    `<m:ParentFolderId><t:DistinguishedFolderId Id="directory"/></m:ParentFolderId>` +
    (query ? `<m:QueryString>${xmlEscape(query)}</m:QueryString>` : ``) +
    `</m:FindPeople>`;
  return envelope(version, body);
}

/**
 * Parse an EWS XML response string into a Document.
 * Throws on a SOAP fault or a non-success ResponseClass.
 * @param {string} xmlText
 * @returns {Document}
 */
export function parseResponse(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "text/xml");

  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error(`Malformed EWS response: ${parseError.textContent}`);
  }

  // SOAP-level fault.
  const fault = doc.getElementsByTagNameNS(NS.soap, "Fault")[0];
  if (fault) {
    const reason =
      fault.querySelector("faultstring")?.textContent || "Unknown SOAP fault";
    throw new Error(`EWS SOAP fault: ${reason}`);
  }

  return doc;
}

export { NS };
