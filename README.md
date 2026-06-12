# Directorium

A **free** Thunderbird add-on that brings the Microsoft Exchange / Office 365
**Global Address List (GAL)** and your **personal contacts** into Thunderbird —
**even on locked-down tenants** where third-party apps are blocked.

A free alternative to the paid **Owl for Exchange**, focused on the directory.

📖 **New here? Read the [User Guide](GUIDE.md).**

> *directorium* (Latin): a directory or guide — also the Church's liturgical
> Ordo. Fitting for an address directory (and a future calendar).

## Why it exists

Thunderbird 145+ added native Exchange **email** for free — but **not** the
address book / GAL, and it isn't on the near-term roadmap. Directorium fills
that gap, and goes further: it works on tenants whose Conditional Access blocks
device-code and third-party EWS OAuth apps (where most free tools fail).

## Features

- **GAL search** — type a name in any recipient field or search the
  organization's address book; results come live from Exchange. Each book is
  **named after the organization** it belongs to (e.g. `contoso.com`), not a
  generic "Exchange GAL", and you can rename it.
- **Multiple accounts** — sign into more than one mailbox (e.g. work + a second
  tenant) and each gets its **own named GAL** that coexist side by side.
- **Personal contacts sync** — one-way copy of your Exchange Contacts folder
  into a normal, browsable Thunderbird address book ("Exchange Contacts").
- **Three ways to connect:**
  - **OWA web login** *(recommended for Office 365, works like Owl)* — sign in
    to webmail once; Directorium reuses the first-party token your Outlook tab
    already uses. **No admin approval, no app registration**, and it passes the
    Conditional Access policies that block other tools.
  - **EWS + OAuth** (device code or your own Azure app) — for lenient tenants.
  - **Basic auth** — for on-premise Exchange.
- **Background token auto-refresh** — once signed in (OWA mode), the token is
  refreshed silently; you rarely need to sign in again.

## Install (development / from source)

Thunderbird requires signing for permanent installation, so for now load it
temporarily (or sign it via [addons.thunderbird.net](https://addons.thunderbird.net)):

1. **Tools → Developer Tools → Debug Add-ons** → **Load Temporary Add-on…**
2. Pick `manifest.json` in this folder.
3. Open **Directorium settings** (Add-ons Manager → Directorium → Options).

A packaged `.xpi` is attached to each GitHub Release.

## Configure (Office 365, locked tenant — the Owl-style path)

1. Account type → **OWA web login**.
2. Click **Sign in (OWA)**, complete the normal Microsoft login, wait for the
   mailbox to load, and **leave that Outlook tab open**.
3. **Refresh status** should show the account with a live token.
4. Search the address book named after your organization, and/or click
   **Sync contacts**. To add another mailbox, click **Add account (sign in)**.

## How it works

```
You sign in to Outlook on the web (first-party login → passes Conditional Access)
        │
        ▼
webRequest captures the Bearer token from the page's own API calls
        │
        ▼
Directorium calls Outlook's own people API (/owa/service.svc FindPeople)
        │
        ├──► GAL search   → addressBooks.provider results (vCard)
        └──► Contacts sync → a real local address book
```

The token is your own, used for your own directory access, and never leaves your
machine. See [src/owa/owaClient.js](src/owa/owaClient.js).

| File | Role |
|------|------|
| `src/background.js` | GAL provider, OWA token capture (`webRequest`), auto-refresh, RPC |
| `src/owa/owaClient.js` | OWA token reuse + FindPeople (GAL + personal contacts) |
| `src/ews/` | EWS SOAP backend (ResolveNames / FindPeople) for OAuth/Basic modes |
| `src/auth/` | OAuth device-code and interactive (PKCE) flows |
| `src/contacts/` | EWS/persona → vCard mapping, personal contacts sync |
| `src/options/` | Settings + diagnostics UI |

## Roadmap

- [x] GAL search
- [x] Personal contacts sync (one-way)
- [ ] Two-way contact sync (push local edits)
- [ ] **Calendar** (via a Thunderbird Experiment API + calendar provider)

> Mail is intentionally out of scope — Thunderbird 145+ does Exchange email
> natively. A WebExtension can't provide a mail account anyway.

## Security

Directorium stores its captured token only in the extension's local storage on
your device. It talks only to your Exchange/Office 365 server and Microsoft
login endpoints. No telemetry, no third-party servers.

## License

[MPL-2.0](LICENSE).
