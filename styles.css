# Setup Guide

## 1. Create the Google Sheet

Create a new Google Sheet called something like `MO8 MDT Database`.

Open:

```text
Extensions > Apps Script
```

Paste in the code from:

```text
apps-script/Code.gs
```

## 2. Create the Database Tabs

In Apps Script, select `setupSpreadsheet` and click **Run**.

Google will ask you to authorise the script. Accept the permissions for your own Sheet.

This creates the database tabs and default permissions.

## 3. Create the First Admin User

In `Code.gs`, find:

```js
function createInitialAdmin() {
```

Edit these values before running it:

```js
const robloxUsername = 'YourRobloxUsername';
const discordId = 'YourDiscordID';
const temporaryPassword = 'ChangeMe123!';
```

Then run `createInitialAdmin`.

After logging in for the first time, replace the temporary password by creating a proper password change feature or manually re-running `setUserPassword_` from Apps Script.

The initial admin is created with:

```text
Rank: Commissioner
Role: Command
```

Rank is the community/policing rank shown on profiles. Role is the system permission group used by the MDT.

## 4. Deploy the Apps Script API

In Apps Script:

```text
Deploy > New deployment > Web app
```

Use:

```text
Execute as: Me
Who has access: Anyone
```

Copy the web app URL.

The app still requires your custom login token before allowing protected actions.

## 5. Connect the Frontend

Open:

```text
frontend/app.js
```

Replace:

```js
YOUR_APPS_SCRIPT_WEB_APP_URL_HERE
```

with your Apps Script web app URL.

## 6. Test Locally

Open:

```text
frontend/index.html
```

Log in using the Roblox username and temporary password from `createInitialAdmin`.

## 7. Host for Free

Good free options:

- GitHub Pages
- Netlify
- Vercel
- Firebase Hosting

You only need to host the contents of the `frontend/` folder.

## Important Security Notes

- Keep the Google Sheet private.
- Do not publish the Sheet publicly.
- Do not share the Apps Script editor with normal officers.
- Use strong passwords for Sergeant+ accounts.
- Remove suspended users by setting their `Status` to `Suspended`.
- Keep `AuditLog` append-only where possible.
