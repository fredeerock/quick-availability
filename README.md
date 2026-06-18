# Availability Composer (GitHub Pages + Google Calendar)

This is a static website that:

- Signs into Google Calendar (recommended)
- Can also sign into Microsoft (Outlook)
- Works without Microsoft app registration by importing an Outlook .ics calendar export
- Reads your busy events from your calendar
- Generates copy-ready meeting options based on:
  - Meeting length (default 60 minutes)
  - Date span (next N days)
  - Number of options to return
  - Daily time window
- Spreads options across days/times to improve odds the other person can make one

## 1) Create Google OAuth client (recommended)

This path is usually easier than Entra for school environments.

1. Open Google Cloud Console.
2. Create/select a project.
3. Enable **Google Calendar API**.
4. Go to **APIs & Services -> OAuth consent screen** and configure app details.
5. Go to **Credentials -> Create Credentials -> OAuth client ID**.
6. Application type: **Web application**.
7. Add Authorized JavaScript origins:
   - `http://localhost:5500`
   - `https://YOUR_GITHUB_USERNAME.github.io`
8. Add Authorized redirect URIs:
   - `https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME/`
   - `http://localhost:5500/`
9. Copy the OAuth **Client ID**.

Calendar scope used by this app:

- `https://www.googleapis.com/auth/calendar.readonly`

## 2) Configure the website

1. Open your deployed site URL (or local static server URL).
2. Paste your Google OAuth Client ID.
3. Click **Sign in with Google**.
4. Set meeting options and click **Find my times**.

Settings are saved in browser local storage.

## 3) Microsoft Outlook mode (optional)

Optional. Skip this entire section if your school does not allow app registrations and use the .ics workflow below.

1. Open the Azure portal -> Microsoft Entra ID -> App registrations -> New registration.
2. Name it anything (for example: `Availability Composer`).
3. Supported account types:
   - Choose `Accounts in any organizational directory and personal Microsoft accounts` if you want both work and personal Outlook accounts.
4. Create the app.
5. Copy the `Application (client) ID`.

### Configure SPA redirect URI

In your app registration:

1. Go to `Authentication`.
2. Add a platform -> `Single-page application`.
3. Add redirect URIs:
   - Your GitHub Pages URL, for example: `https://YOUR_GITHUB_USERNAME.github.io/REPO_NAME/`
   - Optional local test URL: `http://localhost:5500/`
4. Save.

### API permissions

In `API permissions`:

1. Add delegated Microsoft Graph permissions:
   - `User.Read`
   - `Calendars.Read`
2. If your organization requires admin consent, request/grant it.

## No-admin .ics workflow (school-friendly)

If your tenant blocks app registration, use this path:

1. In Outlook, export your calendar as `.ics` (or use a published calendar `.ics` file if your school allows it).
2. In this app, use the `.ics` section to paste text or upload the file.
3. Set meeting length, date span, and number of options.
4. Click **Find my times** and copy the output.

Notes:

- `.ics` mode needs periodic refresh (re-upload) as your calendar changes.
- Live sign-in mode is automatic and always current, but requires app registration.

## 4) Deploy to GitHub Pages

1. Push these files to a GitHub repository.
2. In GitHub repo settings -> Pages:
   - Source: `Deploy from a branch`
   - Branch: `main` (or `master`) and `/ (root)`
3. Visit the generated Pages URL.

## Notes

- This app is fully client-side. No backend is required.
- Google mode requests read-only calendar scope.
- Outlook mode requests `Calendars.Read`.
- If sign-in fails, the redirect URI in Azure must exactly match your live page URL.
