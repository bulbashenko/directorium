# Privacy Policy — Directorium for Exchange

_Last updated: 2026-06-10_

Directorium is a Thunderbird add-on that connects **directly** to your own
Microsoft Exchange / Office 365 (or on-premise Exchange) server to provide the
Global Address List (GAL) and your contacts. Your privacy model is simple:

## What Directorium does NOT do
- It does **not** collect, transmit, or share any of your data with the
  developer or any third party.
- It contains **no analytics, no telemetry, and no tracking**.
- It uses **no servers of its own.** There is no backend.

## What data Directorium handles, and where it stays
All of the following is stored **only** in Thunderbird's local extension
storage on your own device, and is never sent anywhere except your own
Exchange / Microsoft 365 server and Microsoft's sign-in endpoints:

- **Connection settings** you enter (server URL, account type, email).
- **Credentials**: for Basic auth (on-premise), the username/password you enter
  are stored locally and sent only to your Exchange server.
- **Access token**: in "OWA web login" mode, Directorium reuses the OAuth
  bearer token that your own signed-in Outlook-on-the-web session already uses.
  This is *your* token, for *your* access to *your* directory. It is read from
  your local browser session, kept in local extension storage, and sent only
  back to your Microsoft 365 server to look up directory entries.
- **Directory / contact results** returned by your server are shown in
  Thunderbird and, for the contacts sync, saved into a local Thunderbird
  address book on your device.

## Network connections
Directorium communicates only with:
- your Exchange / Microsoft 365 server (e.g. `outlook.office365.com`, or your
  organisation's server), and
- Microsoft's sign-in endpoints (`login.microsoftonline.com`)
for authentication and directory/contact lookups that you initiate.

## Permissions
Each requested permission is used solely to deliver the features above
(directory provider, reusing your session token, opening the sign-in tab,
background token refresh, OAuth). Details are in the README and the source code.

## Source code
Directorium is open source: <https://github.com/bulbashenko/directorium>

## Contact
Questions: open an issue at the repository above.
