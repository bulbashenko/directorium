// Sync the user's personal Exchange contacts into a real, browsable Thunderbird
// address book (not just a search-only one like the GAL). Personal contacts are
// few enough to fully cache locally via the public addressBooks/contacts API.
//
// One-way for now: Exchange → Thunderbird. Re-running replaces the book's
// contents (simple and reliable). Two-way (push local edits back) is a future
// step.

import { getConfig } from "../lib/storage.js";
import { fetchPersonalContacts } from "../owa/owaClient.js";
import { contactToVCard } from "./vcard.js";

const BOOK_NAME = "Exchange Contacts";

// In Manifest V3 the contact API is namespaced under addressBooks
// (messenger.addressBooks.contacts), not the top-level messenger.contacts (MV2).
const contactsApi = messenger.addressBooks.contacts;

/** Find our address book by name, or create it. Returns its id. */
async function getOrCreateBook() {
  const books = await messenger.addressBooks.list();
  const existing = books.find((b) => b.name === BOOK_NAME);
  if (existing) return existing.id;
  return messenger.addressBooks.create({ name: BOOK_NAME });
}

/** Remove every contact currently in the book (for a clean re-sync). */
async function clearBook(bookId) {
  const contacts = await contactsApi.list(bookId);
  for (const c of contacts) {
    await contactsApi.delete(c.id).catch(() => {});
  }
}

/**
 * Pull all personal contacts from Exchange and (re)populate the local book.
 * @returns {Promise<{count:number, bookName:string}>}
 */
export async function syncPersonalContacts() {
  const cfg = await getConfig();
  if (cfg.authMode !== "owa") {
    throw new Error("Personal contacts sync currently requires the OWA web login mode.");
  }

  const contacts = await fetchPersonalContacts();
  const bookId = await getOrCreateBook();
  await clearBook(bookId);

  let created = 0;
  for (const c of contacts) {
    try {
      // MV3: create(parentId, vCardString) — vCard is a positional string.
      await contactsApi.create(bookId, contactToVCard(c));
      created++;
    } catch (err) {
      console.warn("[Directorium] contact create failed:", err.message, c.DisplayName);
    }
  }

  console.debug(`[Directorium] Synced ${created}/${contacts.length} personal contacts into "${BOOK_NAME}"`);
  return { count: created, bookName: BOOK_NAME };
}
