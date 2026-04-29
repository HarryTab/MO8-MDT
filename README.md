# MO8 MDT Starter

Starter intranet/MDT-style system for a Roblox roleplay Roads Policing Unit.

This project uses:

- A static frontend in `frontend/`
- A Google Apps Script backend in `apps-script/Code.gs`
- A Google Sheet as the database
- Google Drive links for documents and training guides

The system is intended for roleplay/community administration only. Store Roblox usernames, Discord IDs, ranks, callsigns, training status, and moderation/admin records. Do not store real-world sensitive personal data.

## Project Structure

```text
apps-script/
  Code.gs
frontend/
  index.html
  styles.css
  app.js
docs/
  sheet-schema.md
  setup.md
```

## Quick Start

1. Create a new Google Sheet.
2. Open `Extensions > Apps Script`.
3. Copy the contents of `apps-script/Code.gs` into the Apps Script editor.
4. In Apps Script, run `setupSpreadsheet()` once.
5. Add your first admin user by running `createInitialAdmin()`.
6. Deploy the Apps Script as a web app.
7. Copy the web app URL.
8. Open `frontend/app.js` and replace `YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` with your URL.
9. Open `frontend/index.html` in your browser or host the `frontend/` folder on GitHub Pages/Netlify/Vercel.

More detailed steps are in `docs/setup.md`. GitHub Pages hosting steps are in `docs/github-pages.md`.

## Default Permission Model

The starter includes these roles:

- `Sergeant`
- `Inspector`
- `Chief Inspector`
- `Command`

Permissions are stored in the `Permissions` sheet tab and checked server-side by Apps Script.

## Security Notes

- Passwords are never stored in plain text.
- The backend stores password hashes with per-user salts.
- Session tokens are stored as hashes in the Sheet.
- Every protected backend action checks the active session and required permission.
- Sensitive actions are written to `AuditLog`.

This is still a lightweight community system, not a high-security enterprise application.
