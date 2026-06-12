import { searchGal, enumerateGal, testConnection } from "./gal.js";
import { startDeviceCode, pollForToken } from "./auth/oauth.js";
import { signInInteractive, getRedirectUrl } from "./auth/authcode.js";
import { signInOwa, captureAuthHeader, tokenStatus, maybeRefresh } from "./owa/owaClient.js";
import { contactToVCard } from "./contacts/vcard.js";
import { syncPersonalContacts } from "./contacts/sync.js";
import * as ews from "./ews/ewsClient.js";
import * as owa from "./owa/owaClient.js";
import {
  getAccounts,
  upsertOwaAccount,
  updateAccountLabel,
  removeAccount,
  refreshBooksMirror,
  readBooksMirror,
} from "./lib/storage.js";

// ---------------------------------------------------------------------------
// GAL search address books (one provider per account)
// ---------------------------------------------------------------------------
// CRITICAL: in Manifest V3 Thunderbird terminates idle background pages. A
// provider listener must be added *synchronously at the top level* during
// startup so it registers as a persistent listener — otherwise the background
// wakes for a search but the listener isn't wired and results never render
// (the symptom: empty/grey contact cards).
//
// Book names are dynamic (one per signed-in account, named after that account's
// org). Storage is async, but the Thunderbird MV3 background is a background
// *page*, so a synchronous localStorage mirror of the book list is available at
// top level — we register from it immediately, then reconcile against the
// authoritative storage.local set asynchronously.

/** bookId → { fn, label, kind } currently registered with Thunderbird. */
const providerRegistry = new Map();

async function onSearch(book, searchString) {
  console.debug(
    `[Directorium] onSearchRequest [${book.label}] — searchString=${JSON.stringify(searchString)}`
  );
  try {
    // Thunderbird sends an empty searchString when the book is first opened
    // (to list "all"); for a directory we only answer real searches.
    const term = (searchString || "").trim();
    if (!term) return { isCompleteResult: true, results: [] };

    let contacts;
    if (book.kind === "ews") {
      ({ contacts } = await ews.searchGal(term));
    } else {
      const account = (await getAccounts()).find((a) => a.id === book.id);
      if (!account) {
        console.warn(`[Directorium] No account for book ${book.id} (${book.label}).`);
        return { isCompleteResult: false, results: [] };
      }
      ({ contacts } = await owa.searchGal(term, account));
    }
    // MV3 provider results must be vCard *strings*, not property objects.
    const results = contacts.map(contactToVCard);
    console.debug(`[Directorium] [${book.label}] returning ${results.length} vCards.`);
    return { isCompleteResult: true, results };
  } catch (err) {
    console.error(`[Directorium] GAL search failed for "${book.label}":`, err);
    return { isCompleteResult: false, results: [] };
  }
}

/** Register (or re-register, if the name changed) a provider for one book. */
function registerProvider(book) {
  const existing = providerRegistry.get(book.id);
  if (existing && existing.label === book.label && existing.kind === book.kind) return;
  if (existing) {
    messenger.addressBooks.provider.onSearchRequest.removeListener(existing.fn);
  }
  const fn = (node, searchString) => onSearch(book, searchString);
  messenger.addressBooks.provider.onSearchRequest.addListener(fn, {
    addressBookName: book.label,
    isSecure: true,
  });
  providerRegistry.set(book.id, { fn, label: book.label, kind: book.kind });
  console.debug(`[Directorium] Registered address book "${book.label}" (${book.kind}).`);
}

function unregisterProvider(id) {
  const entry = providerRegistry.get(id);
  if (!entry) return;
  messenger.addressBooks.provider.onSearchRequest.removeListener(entry.fn);
  providerRegistry.delete(id);
  console.debug(`[Directorium] Unregistered address book "${entry.label}".`);
}

/** Reconcile registered providers against the authoritative storage.local set. */
async function reconcileProviders() {
  const books = await refreshBooksMirror();
  const want = new Map(books.map((b) => [b.id, b]));
  for (const id of [...providerRegistry.keys()]) {
    if (!want.has(id)) unregisterProvider(id);
  }
  for (const book of books) registerProvider(book);
}

// Synchronous top-level registration from the mirror (survives background
// termination — this is the path that handles a cold wake-for-search).
for (const book of readBooksMirror()) registerProvider(book);

// ---------------------------------------------------------------------------
// OWA token capture (for the "owa" auth mode — the free Owl alternative)
// ---------------------------------------------------------------------------
// Observe the Bearer token each signed-in Outlook tab uses for its own API
// calls, file it under its account, and reuse it for that account's GAL
// queries. Registered synchronously at the top level so it survives background
// termination. A brand-new account triggers a provider (re)registration.

messenger.webRequest.onSendHeaders.addListener(
  (details) => {
    let auth = null;
    let anchor = null;
    for (const h of details.requestHeaders || []) {
      const n = h.name.toLowerCase();
      if (n === "authorization") auth = h.value;
      else if (n === "x-anchormailbox") anchor = h.value;
    }
    if (!auth) return;
    captureAuthHeader(auth, anchor)
      .then((res) => {
        if (res?.isNew) reconcileProviders().catch(() => {});
      })
      .catch((e) => console.warn("[Directorium] capture failed:", e));
  },
  {
    urls: [
      "https://outlook.office365.com/*",
      "https://outlook.office.com/*",
      "https://outlook.cloud.microsoft/*",
    ],
  },
  ["requestHeaders"]
);

// Keep OWA tokens fresh in the background: a periodic alarm silently reloads
// each mailbox in a background tab to capture a new token before the old one
// expires.
messenger.alarms.create("owaRefresh", { periodInMinutes: 60 });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "owaRefresh") {
    maybeRefresh().catch((e) => console.warn("[Directorium] owaRefresh failed:", e));
  }
});

// ---------------------------------------------------------------------------
// One-time async startup: migrate the legacy single token, reconcile providers
// against storage, then do an initial refresh check.
// ---------------------------------------------------------------------------
(async () => {
  try {
    await migrateLegacyToken();
    await reconcileProviders();
  } catch (e) {
    console.warn("[Directorium] startup reconcile failed:", e);
  }
  maybeRefresh().catch(() => {});
})();

/** Migrate the pre-multi-account single OWA token into an account, once. */
async function migrateLegacyToken() {
  if ((await getAccounts()).length) return;
  const stored = await messenger.storage.local.get("owaToken");
  const tok = stored.owaToken;
  if (tok?.auth) {
    await upsertOwaAccount(tok.auth, tok.anchor);
    await messenger.storage.local.remove("owaToken");
    console.debug("[Directorium] Migrated legacy OWA token into an account.");
  }
}

// ---------------------------------------------------------------------------
// RPC from the options page
// ---------------------------------------------------------------------------

messenger.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  handleMessage(msg).then(
    (result) => sendResponse({ ok: true, result }),
    (error) => sendResponse({
      ok: false,
      error: String(error?.message || error),
      rawXml: error?.rawXml || null,
    })
  );
  return true;
});

async function handleMessage(msg) {
  switch (msg?.type) {
    case "startSignIn":
      return beginSignIn();

    case "startInteractiveSignIn":
      return signInInteractive();

    case "startOwaSignIn": {
      const res = await signInOwa();
      // A fresh sign-in may have created an account; make sure its book exists.
      await reconcileProviders();
      return res;
    }

    case "owaTokenStatus":
      return tokenStatus();

    case "listAccounts":
      return { accounts: await getAccounts() };

    case "renameAccount": {
      const account = await updateAccountLabel(msg.id, msg.label);
      await reconcileProviders();
      return { account };
    }

    case "removeAccount": {
      await removeAccount(msg.id);
      await reconcileProviders();
      return { ok: true };
    }

    case "getRedirectUrl":
      return { redirectUrl: getRedirectUrl() };

    case "testConnection":
      return testConnection(await accountFor(msg.accountId));

    case "enumerateGal": {
      const { contacts, rawXml } = await enumerateGal(await accountFor(msg.accountId));
      return { count: contacts.length, rawXml };
    }

    case "searchGal": {
      const { contacts, rawXml } = await searchGal(msg.query || "a", await accountFor(msg.accountId));
      return { count: contacts.length, contacts, rawXml };
    }

    case "syncPersonalContacts":
      return syncPersonalContacts(msg.accountId);

    default:
      throw new Error(`Unknown message type: ${msg?.type}`);
  }
}

/** Resolve an account by id for RPC callers; undefined lets gal.js pick. */
async function accountFor(accountId) {
  if (!accountId) return undefined;
  return (await getAccounts()).find((a) => a.id === accountId);
}

async function beginSignIn() {
  const device = await startDeviceCode();
  pollForToken(device.deviceCode, device.interval, device.expiresIn).then(
    () => messenger.runtime.sendMessage({ type: "signInComplete", ok: true }),
    (err) => messenger.runtime.sendMessage({
      type: "signInComplete",
      ok: false,
      error: String(err?.message || err),
    })
  );
  return {
    userCode: device.userCode,
    verificationUri: device.verificationUri,
    message: device.message,
    expiresIn: device.expiresIn,
  };
}
