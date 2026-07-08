# Folio & Compass — Google Sheets Backend Setup

This version stores every client's data as rows in a Google Sheet you own —
completely free, no card needed, and you can open the Sheet any time to see
the raw data. Takes about 10 minutes.

## 1. Create the Google Sheet
1. Go to https://sheets.google.com, create a **Blank spreadsheet**.
2. Rename it (e.g. "EYM Financial — Client Data").
3. Create two tabs (right-click the bottom tab bar > rename, or "+" to add):
   - **Users** — in row 1, add headers: `Email` | `PasswordHash` | `IsAdvisor`
   - **Clients** — in row 1, add headers: `Email` | `Name` | `Phone` | `PAN` | `DOB` | `RiskScore` | `RiskDate` | `SIPsJSON` | `UpdatedAt`

## 2. Add the backend script
1. In the Sheet, go to **Extensions > Apps Script**.
2. Delete any starter code in the editor, and paste in the entire contents of
   `Code.gs` from this folder.
3. Click the **Save** icon (or Ctrl/Cmd+S).

## 3. Deploy it as a Web App
1. Click **Deploy > New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Settings:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**. The first time, Google will ask you to authorize the
   script — click through the "Advanced" / "Go to (unsafe)" prompt (this
   warning appears because it's your own unpublished script, not because
   it's actually unsafe).
5. Copy the **Web app URL** you're given — it looks like:
   `https://script.google.com/macros/s/AKfycb.../exec`

## 4. Paste the URL into the app
Open `index.html`, find this line near the top of the `<script>` section:
```js
const APPS_SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
```
Replace it with the URL you copied.

## 5. Push to GitHub / re-deploy to Pages
Upload the updated files to your GitHub repo (same as before). GitHub Pages
will redeploy automatically within a minute or two.

## 6. Create your advisor account
1. Open the live site, click **Create Account**, sign up with your own email.
2. Open the Google Sheet, go to the **Users** tab, find your row, and change
   `IsAdvisor` from `false`/blank to `TRUE`.
3. Reload the app and log in again — you'll now see the **Advisor Dashboard**
   tab listing every client.

## How it works
- Each investor's data lives in one row of the **Clients** tab.
- Passwords are never stored in plain text — only a SHA-256 hash, in the
  **Users** tab.
- A **Funds** tab is created automatically the first time someone searches
  for a mutual fund — it caches AMFI's full daily scheme list (fetched
  server-side, since browsers are blocked by AMFI from fetching it directly)
  and refreshes itself automatically once it's more than 24 hours old.
- The app calls your Apps Script Web App URL to sign up, log in, save, and
  load data — no separate server needed; Google runs the script for you.

## Updating the script later
Whenever you paste in updated `Code.gs` (like this fund-search version),
saving alone is **not** enough — the live Web App URL keeps running the old
version until you redeploy:
1. **Deploy > Manage deployments**.
2. Click the pencil/edit icon on your existing deployment.
3. Under **Version**, choose **New version**.
4. Click **Deploy**.
The URL stays the same, so you don't need to update `index.html` again.

## Honest limitations
- **This is a lightweight setup, not enterprise-grade security.** There's no
  session/token system — the app re-sends the email+password with each save
  (over HTTPS). That's fine for a small advisory practice with trusted
  clients, but if you're handling a large number of accounts or want stronger
  guarantees, Supabase or Firebase (which have real auth systems) are a
  better fit — happy to switch you over if this ever outgrows Sheets.
- **Row limits:** Google Sheets comfortably handles thousands of rows, so
  this will not be a bottleneck for typical client counts.
- **Apps Script quotas:** the free tier allows generous daily execution time
  or requests for personal/small business use — far beyond what a solo
  advisory practice would use.
- **Editing the Sheet directly is fine** for viewing, but don't hand-edit the
  `SIPsJSON` column unless you're comfortable with valid JSON — a typo there
  will break that client's SIP list until fixed.
