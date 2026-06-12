# Directorium — User Guide

Directorium brings your Microsoft Exchange / Office 365 **Global Address List
(GAL)** and **personal contacts** into Thunderbird — including on locked-down
corporate/university tenants where other free tools fail.

- **GAL search** — start typing a name in any **To:/Cc:** field, or search your
  organization's address book, and matching colleagues appear with email,
  phone, job title, company and department. The book is **named after your
  organization** (e.g. `contoso.com`), and you can rename it.
- **Multiple accounts** — sign into more than one mailbox and each gets its own
  named GAL (and its own contacts book), side by side.
- **Contacts sync** — copy your Exchange *Contacts* folder into a normal,
  browsable Thunderbird address book.

It's a free alternative to *Owl for Exchange* for the directory. (Email itself
is handled natively by Thunderbird 145+, so Directorium doesn't do mail.)

---

## 1. Requirements

- **Thunderbird 128 or newer.**
- A **Microsoft Exchange / Office 365** account (work or university), or an
  on-premise Exchange server.

---

## 2. Install

### From the packaged add-on (.xpi)
1. Download the latest `directorium-x.y.z.xpi` from the
   [Releases page](https://github.com/bulbashenko/directorium/releases).
2. Thunderbird → **Tools → Add-ons and Themes**.
3. Gear icon ⚙️ → **Install Add-on From File…** → choose the `.xpi`.

> The `.xpi` is unsigned, so Thunderbird may only keep it until restart. For a
> permanent install, load it temporarily (below) or install the signed version
> from addons.thunderbird.net once published.

### Temporary (for development / testing)
1. **Tools → Developer Tools → Debug Add-ons**.
2. **Load Temporary Add-on…** → select `manifest.json` from the source folder.

---

## 3. Set up — Office 365 (recommended)

This mode works even when your organisation blocks third-party apps.

1. Open **Add-ons → Directorium → Options** (or Preferences).
2. **Account type** → **Office 365 — OWA web login**.
3. Leave **OWA URL** as `https://outlook.office365.com`.
4. Click **Add account (sign in)**. A tab opens to Microsoft webmail.
5. Sign in with your normal account and password (and MFA if asked).
6. **Wait for your mailbox to load, then leave that tab open.**
7. Click **Refresh status** — your account should appear under **Accounts**
   with a live token, and a new address book named after your organization.

That's it. Directorium now uses the same secure session your browser does.

> **More than one mailbox?** Click **Add account (sign in)** again and sign in
> with the other account. Each mailbox gets its own named address book. Use the
> **Rename** button to give a book a friendlier name (takes effect after a
> Thunderbird restart), or **Remove** to drop an account.

### Set up — On-premise Exchange
1. **Account type** → **On-premise Exchange**.
2. Enter your email and click **Auto-detect** (or set the **EWS URL** manually,
   e.g. `https://mail.yourcompany.com/EWS/Exchange.asmx`).
3. Enter your username (`DOMAIN\user` or your email) and password.

### Set up — EWS over OAuth (lenient tenants / your own Azure app)
Use **Office 365 — EWS device code** for tenants that allow it, or
**EWS interactive, your own app** if you registered an Azure application.
(Most Office 365 users should just use **OWA web login** above.)

---

## 4. Daily use

### Search the directory (GAL)
- In the **address book** (Tools → Address Book), click the book named after
  your organization and type in the search box.
- Or just start typing a name in the **To:** field when composing — matches
  appear automatically.

### Sync your personal contacts
- Options → **Sync contacts** (next to the account), or the **Sync my contacts**
  button for the first account.
- A new address book **Contacts — <your org>** is created/updated with that
  account's Exchange Contacts folder. Re-run it any time to refresh.

---

## 5. Keeping it working

- In **OWA web login** mode, keep an Outlook web tab open occasionally — the
  add-on refreshes its token silently in the background (about hourly) for as
  long as your Microsoft session is valid (usually days).
- If a search ever fails with an expired-session error, just click
  **Add account (sign in)** again with that mailbox.

---

## 6. Troubleshooting

| Symptom | Fix |
|--------|-----|
| **Contacts show as empty/grey cards** | Update to the latest version; reload the add-on. |
| **"No OWA token yet"** | Click **Add account (sign in)** and leave the Outlook tab open; then **Refresh status**. |
| **Search returns an "expired" error** | Click **Add account (sign in)** again with that mailbox. |
| **Renamed book still shows the old name** | Restart Thunderbird — address-book name changes apply on restart. |
| **Error 53003 on device-code sign-in** | Your tenant blocks device code — use **OWA web login** instead. |
| **"Need admin approval" on your own app** | Your tenant blocks third-party apps — use **OWA web login** instead. |
| **Validator warning: "Invalid permission addressBooks"** | Harmless — it's a Firefox-validator quirk; Thunderbird needs and supports this permission. |
| **Auto-detect can't find the server** | Ask IT for the EWS URL and enter it manually. |

Detailed logs: **Debug Add-ons → Inspect** (under Directorium) opens the
background console, where lines are tagged `[Directorium]`.

---

## 7. Privacy & security

Directorium talks only to your Exchange / Microsoft 365 server and Microsoft's
sign-in endpoints. In OWA mode it reuses **your own** session token, kept only
in Thunderbird's local storage on your device. There is **no telemetry** and
**no third-party server**. Source code: <https://github.com/bulbashenko/directorium>.

---

## 8. FAQ

**Why does it open Outlook on the web to sign in?**
Many organisations block third-party mail apps and the "device code" login. The
normal webmail login is a trusted first-party flow that those policies allow, so
this is the only reliable way to reach the directory without an administrator.

**Is this like Owl for Exchange?**
For the address book / GAL and contacts — yes, and free. Owl also does mail and
calendar via deep Thunderbird integration; Directorium focuses on the directory.

**What about email?**
Use Thunderbird's built-in Exchange support (Thunderbird 145+) to add your
mailbox. Add-ons can't provide mail accounts.

**Calendar?**
On the roadmap.
