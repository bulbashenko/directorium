import { searchGal, enumerateGal, testConnection } from "./gal.js";
import { startDeviceCode, pollForToken } from "./auth/oauth.js";
import { signInInteractive, getRedirectUrl } from "./auth/authcode.js";
import { signInOwa, captureAuthHeader, tokenStatus, maybeRefresh } from "./owa/owaClient.js";
import { contactToVCard } from "./contacts/vcard.js";
import { syncPersonalContacts } from "./contacts/sync.js";

// ---------------------------------------------------------------------------
// GAL search address book
// ---------------------------------------------------------------------------
// CRITICAL: in Manifest V3 Thunderbird terminates idle background pages. A
// provider listener must be added *synchronously at the top level* during
// startup so it registers as a persistent listener — otherwise the background
// wakes for a search but the listener isn't wired and results never render
// (the symptom: empty/grey contact cards).
//
// Because storage is async we can't read a custom book name here, so the name
// is a fixed constant. The search handler itself may run async.

const ADDRESS_BOOK_NAME = "Exchange GAL";

async function onSearch(node, searchString, query) {
  console.debug(
    `[Directorium] onSearchRequest fired — searchString=${JSON.stringify(searchString)} query=${JSON.stringify(query)}`
  );
  try {
    // Thunderbird sends an empty searchString when the book is first opened
    // (to list "all"); for a directory we only answer real searches.
    const term = (searchString || "").trim();
    if (!term) {
      return { isCompleteResult: true, results: [] };
    }

    const { contacts } = await searchGal(term);
    // MV3 provider results must be vCard *strings*, not property objects.
    const results = contacts.map(contactToVCard);
    console.debug(
      `[Directorium] onSearchRequest returning ${results.length} vCards. Sample:\n${results[0]}`
    );
    return { isCompleteResult: true, results };
  } catch (err) {
    console.error("[Directorium] GAL search failed:", err);
    return { isCompleteResult: false, results: [] };
  }
}

messenger.addressBooks.provider.onSearchRequest.addListener(onSearch, {
  addressBookName: ADDRESS_BOOK_NAME,
  isSecure: true,
});

// ---------------------------------------------------------------------------
// OWA token capture (for the "owa" auth mode — the free Owl alternative)
// ---------------------------------------------------------------------------
// Observe the Bearer token the signed-in Outlook tab uses for its own API
// calls, and reuse it for our GAL queries. Registered synchronously at the top
// level so it survives background termination.

messenger.webRequest.onSendHeaders.addListener(
  (details) => {
    let auth = null;
    let anchor = null;
    for (const h of details.requestHeaders || []) {
      const n = h.name.toLowerCase();
      if (n === "authorization") auth = h.value;
      else if (n === "x-anchormailbox") anchor = h.value;
    }
    if (auth) captureAuthHeader(auth, anchor);
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

// Keep the OWA token fresh in the background: a periodic alarm silently reloads
// the mailbox in a background tab to capture a new token before the old expires.
messenger.alarms.create("owaRefresh", { periodInMinutes: 60 });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "owaRefresh") {
    maybeRefresh().catch((e) => console.warn("[Directorium] owaRefresh failed:", e));
  }
});
// Also check once shortly after startup.
maybeRefresh().catch(() => {});

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

    case "startOwaSignIn":
      return signInOwa();

    case "owaTokenStatus":
      return tokenStatus();

    case "getRedirectUrl":
      return { redirectUrl: getRedirectUrl() };

    case "testConnection":
      return testConnection();

    case "enumerateGal": {
      const { contacts, rawXml } = await enumerateGal();
      return { count: contacts.length, rawXml };
    }

    case "searchGal": {
      const { contacts, rawXml } = await searchGal(msg.query || "a");
      return { count: contacts.length, contacts, rawXml };
    }

    case "syncPersonalContacts":
      return syncPersonalContacts();

    default:
      throw new Error(`Unknown message type: ${msg?.type}`);
  }
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
