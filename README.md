# ReportGen

Pulls Google Analytics 4 data (and optionally Google Sheets SEO data) for a client and two user-chosen reporting periods, and generates a downloadable PowerPoint (`.pptx`) report: traffic overview, SEO performance overview, organic search performance, ecommerce funnel, top landing pages, and an optional paid media slide.

## Setup

1. Install dependencies:

   ```
   npm install
   ```

2. Fill in `.env` in the project root:

   ```
   GOOGLE_CLIENT_ID=your-oauth-client-id
   GOOGLE_CLIENT_SECRET=your-oauth-client-secret
   GOOGLE_REDIRECT_URI=http://localhost:3000/auth/callback
   PORT=3000
   SESSION_SECRET=any-random-string
   GA4_PROPERTY_ID=your-ga4-property-id
   GOOGLE_SHEET_ID=your-google-sheet-id
   ```

   `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, and `GOOGLE_REDIRECT_URI` are required — the server refuses to boot without them. These come from a Google Cloud project with the **Google Analytics Data API** and **Google Sheets API** enabled and an OAuth 2.0 Client ID (Web application) configured with `http://localhost:3000/auth/callback` as an authorized redirect URI.

   While the OAuth consent screen is in "Testing" mode, add every Google account that will sign in under **Test users**, or sign-in will fail with `403 access_denied`.

   `SESSION_SECRET` is optional — if missing, the server generates one and appends it to `.env` on first boot. `GA4_PROPERTY_ID`/`GOOGLE_SHEET_ID` are also optional; they're only used once, to seed a "Default Client" in `clients.json` the very first time the app runs with no clients configured.

3. Start the app:

   ```
   npm run dev
   ```

4. Open http://localhost:3000

## Using it

1. Click **Sign in with Google** and complete the consent screen. Sign-in is per browser session — multiple people can sign in with different Google accounts against the same running server without clobbering each other's credentials (tokens are stored locally in `tokens.json`, keyed by email).
2. Add one or more clients (name + GA4 property ID, and optionally a Google Sheet ID for SEO data) via the client selector — managed through `GET/POST /api/clients` and `DELETE /api/clients/:id`. There's no edit endpoint; delete and re-add to change a client's IDs.
3. Pick a client, a **Current Period** and a **Comparison Period** (independent month/year pickers — they don't need to be adjacent months), optionally expand **Paid Media** to enter Google Ads / Meta Ads numbers by hand, and click **Generate Report**. The `.pptx` downloads automatically once data has been fetched and the slides are built. Any non-blocking warnings (e.g. missing GA4 or SEO data for a period) are shown alongside the download.

## Notes

- There is no test suite, linter, or build tooling configured.
- `tokens.json` and `clients.json` are gitignored and hold real credentials/IDs — never commit them.
- See `CLAUDE.md` for a detailed architecture walkthrough (request flow, auth/token handling, GA4/Sheets data layers, and PPTX generation).
