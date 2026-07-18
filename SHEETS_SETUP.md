# Folio & Compass — Google Sheets Backend Setup

## What changed in this version (v14)
- **New CAS format supported:** the PDF importer (Admin Console and each investor's own
  "Update my portfolio from a CAS statement") now also reads MF Central's newer
  **"Consolidated Account Summary"** layout (filename usually starts `cas_summary_report_…`),
  which lists Folio No./Demat Client Id, Scheme Details, Invested Value, Balance Units, NAV
  Date, NAV, Market Value and Gain/Loss — no ISIN or registrar name per row, unlike the
  CAMS/KFintech detailed statement the importer already handled. It's tried automatically as
  a fallback whenever the original format doesn't match, so nothing needs to be configured —
  just upload the PDF. Both SoA and Demat holdings pages are read; zero-balance/zero-value
  schemes (funds fully redeemed in the past but still listed) are skipped automatically.
- **New: live NAV Performance chart per scheme.** On the Dashboard, expanding any scheme's
  row (Drill down by "Scheme") now shows a second chart below the Invested-vs-Current one: an
  actual NAV history line chart for that fund, fetched live from the free public
  **api.mfapi.in** historical-NAV API (proxied through the Apps Script backend, same as the
  SIP NAV-history simulation already used), with 1M/6M/1Y/3Y/All period toggles. For CAS-
  uploaded holdings (which don't carry an AMFI scheme code the way a SIP added via the fund
  search box does), the scheme is first matched by name against AMFI's fund list — if no
  confident match is found, the chart area explains that instead of showing nothing.
- Because of the above, this version adds one new backend action (`getNavHistory`) — **you
  must redeploy** (see "Updating the script later" below) for it to work; otherwise you'll
  see "Unknown action" if the NAV chart tries to load.

## What changed in v13
- **Speed:** the Dashboard used to make one backend call *per SIP* just to open — an
  investor with 6 SIPs paid for 6 network round-trips. It's now one batched call
  (`computeSipAccumulationBatch`). Separately, the AMFI fund list used to occasionally force
  a live page load to wait 10-20s for a full re-fetch when it went stale; a live request now
  always serves whatever's cached instantly, and you can enable a background daily refresh
  (Admin Console → AMFI Fund List → "Enable automatic daily refresh") so it never goes stale
  in the first place.
- **Security is unchanged on purpose.** Salted PBKDF2 password hashing, session tokens, and
  the login lockout weren't the source of the slowness — they're cheap (a hash only runs once
  at login; every other check is a fast lookup) — so removing them wouldn't have fixed speed
  and would put every investor's password and CAS/PAN data at real risk. See the note at the
  top of `Code.gs` for specifics.
- **New:** a small NIFTY 50 / Bank NIFTY / DOW / China index ticker on the login screen and
  Dashboard (`getMarketIndices`, cached 5 min, no login required to view).
- **New:** Dashboard "Drill down by" — Scheme (as before), Mutual Fund (AMC), Category, or
  Domestic/International, each expandable back down to individual schemes. Groupings are
  inferred from scheme names as a best-effort guide, not an official categorization.
- **New:** a "Flush all holdings" button (investor's own Invest tab, and Admin Console per
  investor) that clears CAS-uploaded holdings only — SIPs, risk profile and personal details
  are untouched — so a fresh CAS can be uploaded onto a clean slate. Both require a confirm
  dialog and cannot be undone.

This version stores every client's data as rows in a Google Sheet you own —
completely free, no card needed, and you can open the Sheet any time to see
the raw data. Takes about 10 minutes.

## 1. Create the Google Sheet
1. Go to https://sheets.google.com, create a **Blank spreadsheet**.
2. Rename it (e.g. "EYM Financial — Client Data").
3. Create two tabs (right-click the bottom tab bar > rename, or "+" to add):
   - **Users** — in row 1, add headers: `Email` | `PasswordHash` | `IsAdvisor` | `Blocked`
   - **Clients** — in row 1, add headers: `Email` | `Name` | `Phone` | `PAN` | `DOB` | `RiskScore` | `RiskDate` | `SIPsJSON` | `UpdatedAt` | `HoldingsJSON`

Three more tabs — **Funds**, **Messages**, and **Requests** — are created automatically
the first time they're needed, so you don't need to add them by hand.

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

## 6. Create your super-user (advisor / admin) account
1. Open the live site, click **Create Account**, sign up with your own email.
2. Open the Google Sheet, go to the **Users** tab, find your row, and change
   `IsAdvisor` from `false`/blank to `TRUE`.
3. Reload the app and log in again — you'll now see three extra tabs:
   - **Admin Console** — create, edit, or permanently delete any user account;
     pick any investor from the dropdown, then add a SIP on their behalf,
     upload their CAS holdings, reset their password, grant/revoke super-user
     rights, block/unblock their login, flush all of their data (portfolio,
     SIPs, messages, requests) at once, and post a banner message to one or
     all investors.
   - **AUM** — firm-wide total invested / present value, with drill-down by
     investor and by fund.
   - **Purchase Requests** — every "New Investment / Additional Purchase"
     request investors submit from their own **Invest** tab, with a status
     dropdown (Pending / Contacted / Completed / Rejected).

## Confirming your deployment is actually up to date
"Unknown action" errors almost always mean Code.gs was edited and saved, but the
live Web App is still running an older version — Apps Script does NOT auto-update
a deployment when you save; you must explicitly create a new version.

The Admin Console now shows a **live "Backend deployment" banner** at the top —
it calls the deployed script and reports back its version string. After you
redeploy, refresh the app and check that banner matches the version noted in
Code.gs's header comment. If it's stale or unreachable, redeploy:
**Deploy > Manage deployments > pencil icon > Version: New version > Deploy.**

## Self-service CAS upload
Every investor now has an "Update my portfolio from a CAS statement" section on
their own **Invest** tab — same PDF importer as the Admin Console, scoped to
their own account. They don't need advisor help to keep their holdings current.

## Notes on the CAS upload
The Admin Console now has a real PDF importer: upload the investor's Consolidated
Account Statement (from MF Central, CAMS, or KFintech) and it auto-fills the
holdings table below (folio, scheme, units, NAV, invested value) by reading the
PDF's actual text positions — no manual typing needed. Scheme names may come
through slightly abbreviated (wrapped continuation lines are dropped); edit any
row before confirming. If the PDF is password-protected, enter the password
(often the investor's PAN in capitals) in the field provided. "Replace"
overwrites their holdings list; "Append" adds to what's already there.

## Notes on present value accuracy
For SIPs added by picking a fund from the AMFI search box, the app now looks up
that scheme's actual historical NAV (via the free public API at api.mfapi.in)
and simulates every installment for real — so "present value" is genuinely
"total units actually bought on each SIP date × today's NAV," not a manually
maintained units field. You'll see "✓ units from actual NAV history" on those
rows. SIPs added before this feature, or where a fund was typed in manually
without picking a search result, fall back to the old estimate (manually
entered units) and are labeled as such. CAS-imported holdings always use the
exact units/value from the statement, since that's already an exact snapshot.

Every time the Dashboard is opened (and every ~25s while it's on screen), the
app re-runs this simulation up to *today's* date, so units/value/gain keep
climbing on their own as each SIP installment date passes — nobody has to go
back in and edit anything.

### Entering a SIP that's already been running for a while
Two ways to add a SIP, both on the same form:
- **Pure auto mode** — enter Start date, amount, frequency and leave the "As
  of date / Invested till then / Market value as of then" fields blank. The
  engine derives the *entire* history from Start date to today off real NAVs.
  Best for funds with clean, complete NAV history going all the way back.
- **Hybrid / as-of mode** — if you already know what the SIP was worth on a
  given date (e.g. from the client's last CAS or statement), fill in "As of
  date", "Invested till that date (₹)" and "Market value as of that date
  (₹)" instead. The app converts that market value into units at that date's
  real NAV, then simulates only the *new* installments from the day after
  that date through today, adding them on top. This avoids re-deriving a long
  or messy purchase history and still keeps updating automatically as new SIP
  dates pass. Rows using this mode are labeled "✓ opening value as of
  <date> + N auto-SIP installment(s) since".

The plain "Total invested till date (₹)" field (without an As-of date) is a
manual override for the *invested* figure only — it does not change how units
are computed, and is unrelated to the as-of snapshot fields.

The Dashboard's "Future Projection" now also shows explicit Year 10 and Year
20 rupee figures (not just the chart), using the standard continuing-SIP
future value formula at an assumed 12% p.a. — clearly labeled as illustrative,
not a guarantee.

## A note on "instant" updates
Because this is a serverless Sheets backend (no live database connection),
an investor's own browser tab polls for updates automatically every ~25
seconds while they're on the Dashboard, and always re-fetches on login — so
a SIP or CAS you add on their behalf will appear for them within moments,
without them needing to do anything.

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
