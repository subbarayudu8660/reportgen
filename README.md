# ReportGen

Pulls Google Analytics 4 data for a selected month and generates a downloadable PowerPoint (`.pptx`) report: traffic overview, organic search performance, ecommerce funnel, and top landing pages.

## Setup

1. Install dependencies:

   ```
   npm install
   ```

2. Fill in `.env` in the project root:

   ```
   GOOGLE_CLIENT_ID=your-oauth-client-id
   GOOGLE_CLIENT_SECRET=your-oauth-client-secret
   GA4_PROPERTY_ID=your-ga4-property-id
   GOOGLE_REDIRECT_URI=http://localhost:3000/auth/callback
   PORT=3000
   ```

   These come from a Google Cloud project with the **Google Analytics Data API** enabled and an OAuth 2.0 Client ID (Web application) configured with `http://localhost:3000/auth/callback` as an authorized redirect URI.

   While the OAuth consent screen is in "Testing" mode, add the Google account you'll sign in with under **Test users**, or sign-in will fail with `403 access_denied`.

3. Start the app:

   ```
   npm run dev
   ```

4. Open http://localhost:3000

## Using it

1. Click **Sign in with Google** and complete the consent screen (only needed once — the refresh token is saved locally to `tokens.json`).
2. Pick a month from the dropdown and click **Generate Report**. The `.pptx` downloads automatically once GA4 data has been fetched and the slides are built.
