// Build vCard 4.0 strings from our flat contact objects.
//
// In Manifest V3, addressBooks.provider.onSearchRequest must return each result
// as a vCard *string* — Thunderbird runs vCardToAbCard() on it directly. (The
// older docs showing {DisplayName, PrimaryEmail} objects apply to MV2 only.)

/** Escape a value for a vCard text field (RFC 6350 §3.4). */
function esc(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

/** Escape a single component of a structured field (N, ORG) — keep field ';'. */
function escComponent(value) {
  return esc(value);
}

/**
 * Convert one flat contact object (from the mapper) into a vCard 4.0 string.
 * @param {Record<string,string>} c
 * @returns {string}
 */
export function contactToVCard(c) {
  const lines = ["BEGIN:VCARD", "VERSION:4.0"];

  // FN is mandatory in vCard 4.0.
  const fn = c.DisplayName || c.PrimaryEmail || "Unknown";
  lines.push(`FN:${esc(fn)}`);

  // Structured name: Family;Given;Additional;Prefix;Suffix
  if (c.LastName || c.FirstName) {
    lines.push(
      `N:${escComponent(c.LastName || "")};${escComponent(c.FirstName || "")};;;`
    );
  }

  if (c.NickName) lines.push(`NICKNAME:${esc(c.NickName)}`);

  // Emails. The first is the preferred one.
  if (c.PrimaryEmail) lines.push(`EMAIL;PREF=1;TYPE=work:${esc(c.PrimaryEmail)}`);
  if (c.SecondEmail) lines.push(`EMAIL;TYPE=home:${esc(c.SecondEmail)}`);

  // Phones.
  if (c.WorkPhone) lines.push(`TEL;TYPE="work,voice":${esc(c.WorkPhone)}`);
  if (c.CellularNumber) lines.push(`TEL;TYPE="cell,voice":${esc(c.CellularNumber)}`);
  if (c.HomePhone) lines.push(`TEL;TYPE="home,voice":${esc(c.HomePhone)}`);
  if (c.FaxNumber) lines.push(`TEL;TYPE="work,fax":${esc(c.FaxNumber)}`);
  if (c.PagerNumber) lines.push(`TEL;TYPE=pager:${esc(c.PagerNumber)}`);

  // Job info.
  if (c.JobTitle) lines.push(`TITLE:${esc(c.JobTitle)}`);
  if (c.Company || c.Department) {
    lines.push(
      `ORG:${escComponent(c.Company || "")}` +
        (c.Department ? `;${escComponent(c.Department)}` : "")
    );
  }

  // A stable UID keeps Thunderbird from merging distinct people.
  const uid = c.PrimaryEmail || fn;
  lines.push(`UID:${esc(uid)}`);

  lines.push("END:VCARD");
  return lines.join("\r\n");
}
