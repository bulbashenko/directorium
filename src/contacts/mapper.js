import { NS } from "../ews/soap.js";

// GUIDs appear as mailbox names for system mailboxes and DLs — not useful as display names.
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const isGuid = (s) => GUID_RE.test(s?.trim() || "");

// Exchange internal routing addresses (not SMTP) — not shown to the user.
const isSmtpEmail = (s) => s && s.includes("@") && !s.startsWith("/o=");

function text(parent, localName, ns = NS.t) {
  const el = parent.getElementsByTagNameNS(ns, localName)[0];
  return el ? el.textContent.trim() : "";
}

function keyedEntries(parent, dictLocalName) {
  const result = {};
  const dict = parent.getElementsByTagNameNS(NS.t, dictLocalName)[0];
  if (!dict) return result;
  for (const entry of dict.getElementsByTagNameNS(NS.t, "Entry")) {
    const key = entry.getAttribute("Key");
    const value = entry.textContent.trim();
    if (key && value) result[key] = value;
  }
  return result;
}

/**
 * Convert one EWS <t:Resolution> (from ResolveNames) into ContactProperties.
 * Returns null for system/resource mailboxes with no usable display name.
 */
export function resolutionToContact(resolution) {
  const mailbox = resolution.getElementsByTagNameNS(NS.t, "Mailbox")[0];
  const contact = resolution.getElementsByTagNameNS(NS.t, "Contact")[0];

  // Prefer the Contact element's email list (SMTP) over the raw Mailbox address
  // which may be an internal EX routing address.
  const rawMailboxEmail = mailbox ? text(mailbox, "EmailAddress") : "";
  const primaryEmail =
    (isSmtpEmail(rawMailboxEmail) ? rawMailboxEmail : "") ||
    (contact ? firstEmail(contact) : "");

  // Contact.DisplayName > Mailbox.Name (skip GUIDs) > email
  const rawName =
    (contact && text(contact, "DisplayName")) ||
    (mailbox && !isGuid(text(mailbox, "Name")) ? text(mailbox, "Name") : "");
  const displayName = rawName || primaryEmail;

  // Skip entries that have no meaningful identity at all.
  if (!displayName && !primaryEmail) return null;

  const props = { DisplayName: displayName };
  if (primaryEmail) props.PrimaryEmail = primaryEmail;
  if (contact) Object.assign(props, contactFields(contact));
  return props;
}

function contactFields(contact) {
  const props = {};
  const set = (k, localName) => {
    const v = text(contact, localName);
    if (v) props[k] = v;
  };

  set("FirstName", "GivenName");
  set("LastName", "Surname");
  set("JobTitle", "JobTitle");
  set("Company", "CompanyName");
  set("Department", "Department");
  set("NickName", "Nickname");

  const phones = keyedEntries(contact, "PhoneNumbers");
  if (phones.BusinessPhone) props.WorkPhone = phones.BusinessPhone;
  if (phones.HomePhone) props.HomePhone = phones.HomePhone;
  if (phones.MobilePhone) props.CellularNumber = phones.MobilePhone;
  if (phones.BusinessFax) props.FaxNumber = phones.BusinessFax;
  if (phones.Pager) props.PagerNumber = phones.Pager;

  const emails = keyedEntries(contact, "EmailAddresses");
  // EmailAddress1 is already PrimaryEmail; grab a second one if present.
  if (emails.EmailAddress2 && isSmtpEmail(emails.EmailAddress2)) {
    props.SecondEmail = emails.EmailAddress2;
  }

  return props;
}

function firstEmail(contact) {
  const emails = keyedEntries(contact, "EmailAddresses");
  for (const k of ["EmailAddress1", "EmailAddress2", "EmailAddress3"]) {
    if (emails[k] && isSmtpEmail(emails[k])) return emails[k];
  }
  return "";
}

/**
 * Convert one FindPeople <t:Persona> into ContactProperties.
 */
export function personaToContact(persona) {
  const displayName = text(persona, "DisplayName");

  // Persona email lives in EmailAddresses/t:EmailAddress/t:EmailAddress
  let primaryEmail = "";
  const emailsNode = persona.getElementsByTagNameNS(NS.t, "EmailAddresses")[0];
  if (emailsNode) {
    for (const addr of emailsNode.getElementsByTagNameNS(NS.t, "EmailAddress")) {
      const candidate = text(addr, "EmailAddress");
      if (isSmtpEmail(candidate)) { primaryEmail = candidate; break; }
    }
  }

  if (isGuid(displayName) && !primaryEmail) return null;
  if (!displayName && !primaryEmail) return null;

  const props = { DisplayName: isGuid(displayName) ? primaryEmail : (displayName || primaryEmail) };
  if (primaryEmail) props.PrimaryEmail = primaryEmail;

  const set = (k, localName) => { const v = text(persona, localName); if (v) props[k] = v; };
  set("FirstName", "GivenName");
  set("LastName", "Surname");
  set("JobTitle", "Title");
  set("Company", "CompanyName");
  set("Department", "Department");

  return props;
}
