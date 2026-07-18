/**
 * Folio & Compass — Google Sheets backend.
 * Paste this into Extensions > Apps Script in your Google Sheet, then deploy as a Web App.
 * See SHEETS_SETUP.md for step-by-step instructions.
 *
 * IMPORTANT — after pasting this version in for the first time, run runPbkdf2SelfTest_()
 * once from the Apps Script editor (function dropdown, top toolbar > select it > Run) to
 * confirm password hashing works correctly on your account before deploying. Check View >
 * Logs — it should say "PBKDF2 self-test passed." (it throws a visible error instead if
 * something's wrong). Existing users' passwords keep working either way — they're
 * transparently upgraded to the stronger hash the next time they log in.
 *
 * Sheets used (created automatically if missing, except Users/Clients which you create once):
 *   "Users"    columns: Email | PasswordHash | IsAdvisor | Blocked
 *   "Clients"  columns: Email | Name | Phone | PAN | DOB | RiskScore | RiskDate | SIPsJSON | UpdatedAt | HoldingsJSON
 *   "Funds"            SchemeCode | SchemeName | NAV | LastRefreshed            (auto-created)
 *   "Messages"         MessageId | Target | Message | CreatedAt | Active         (auto-created)
 *   "Requests"         RequestId | Email | Name | Type | FundName | Amount | Notes | Status | CreatedAt (auto-created)
 *
 * "IsAdvisor" = TRUE makes that account a super-user / admin: they can add SIPs or CAS
 * holdings on behalf of any investor, change any investor's password, block/unblock logins,
 * create/edit/delete any user account, flush a user's data, post banner messages, view
 * firm-wide AUM, and action New Purchase Requests.
 *
 * Auth model: login/signup are the only two actions that ever see a raw password. Every
 * other action authenticates with a session token (issued at login/signup) instead — the
 * browser never stores or resends the password itself. A session stays valid until the
 * user logs out or an admin blocks that account. Failed logins on an account lock it out
 * for 15 minutes after 5 attempts.
 *
 * v13 note: the slowness reported in this version was NOT coming from the password hashing
 * or session checks (those run once at login, or a cheap in-memory lookup on every other
 * call) — it was coming from (a) the Dashboard making one network round-trip per SIP instead
 * of one batched call, and (b) a stale AMFI fund list occasionally forcing a live page load
 * to wait 10-20s for a full re-fetch. Both are fixed below (computeSipAccumulationBatch,
 * fundsNeedRefresh/setupDailyFundRefreshTrigger) with the auth/session/hashing model left
 * exactly as-is, since weakening it wouldn't have fixed the speed problem and would put
 * every investor's login and CAS/PAN data at needless risk.
 */

// Bump this whenever you paste in new code. The Admin Console shows this live so you can
// always confirm the deployed Web App is actually running what you just pasted — "Unknown
// action" errors almost always mean you edited Code.gs but didn't create a NEW deployment
// version (Deploy > Manage deployments > pencil icon > Version: New version > Deploy).
const SCRIPT_VERSION = 'v14-2026-07-18-navhistory-cas-v2';

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'version') result = { ok: true, version: SCRIPT_VERSION };
    else if (action === 'signup') result = signup(body.email, body.password);
    else if (action === 'login') result = login(body.email, body.password);
    else if (action === 'logout') result = logout(body.token);
    else if (action === 'whoami') result = whoami(body.token);
    else if (action === 'saveData') result = saveData(body);
    else if (action === 'loadData') result = loadData(body.email, body.token);
    else if (action === 'listClients') result = listClients(body.email, body.token);
    else if (action === 'searchFunds') result = searchFunds(body.query);
    else if (action === 'refreshFunds') result = requireAdminAction(body, refreshFundList);
    else if (action === 'lookupNavs') result = lookupNavs(body.names || []);
    else if (action === 'computeSipAccumulation') result = computeSipAccumulation(body);
    else if (action === 'computeSipAccumulationBatch') result = computeSipAccumulationBatch(body);
    else if (action === 'getNavHistory') result = getNavHistory(body);
    else if (action === 'getNews') result = getNews();
    else if (action === 'getMarketIndices') result = getMarketIndices();
    else if (action === 'setupReminderTrigger') result = requireAdminAction(body, setupDailyReminderTrigger);
    else if (action === 'setupFundRefreshTrigger') result = requireAdminAction(body, setupDailyFundRefreshTrigger);
    else if (action === 'getMessages') result = getMessages(body.email, body.token);
    else if (action === 'getLoginMessages') result = getLoginMessages();
    else if (action === 'submitPurchaseRequest') result = submitPurchaseRequest(body);
    else if (action === 'uploadCas') result = uploadOwnCas(body);
    else if (action === 'flushHoldings') result = flushHoldings(body);
    // ---- super-user / admin actions ----
    else if (action === 'adminAddSip') result = adminAddSip(body);
    else if (action === 'adminUpdateSip') result = adminUpdateSip(body);
    else if (action === 'adminDeleteSip') result = adminDeleteSip(body);
    else if (action === 'adminUploadCas') result = adminUploadCas(body);
    else if (action === 'adminFlushHoldings') result = adminFlushHoldings(body);
    else if (action === 'adminChangePassword') result = adminChangePassword(body);
    else if (action === 'adminSetBlocked') result = adminSetBlocked(body);
    else if (action === 'adminCreateUser') result = adminCreateUser(body);
    else if (action === 'adminUpdateUser') result = adminUpdateUser(body);
    else if (action === 'adminDeleteUser') result = adminDeleteUser(body);
    else if (action === 'adminFlushUserData') result = adminFlushUserData(body);
    else if (action === 'adminBroadcast') result = adminBroadcast(body);
    else if (action === 'adminDeleteMessage') result = adminDeleteMessage(body);
    else if (action === 'adminListMessages') result = adminListMessages(body);
    else if (action === 'adminListRequests') result = adminListRequests(body);
    else if (action === 'adminUpdateRequestStatus') result = adminUpdateRequestStatus(body);
    else result = { ok: false, error: 'Unknown action: ' + action + ' (this deployment is ' + SCRIPT_VERSION + ' — if this action should exist, you likely need to redeploy a new version)' };
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUsersSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
}
function getClientsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
}
function getFundsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Funds');
  if (!sheet) {
    sheet = ss.insertSheet('Funds');
    sheet.appendRow(['SchemeCode', 'SchemeName', 'NAV', 'LastRefreshed']);
  }
  return sheet;
}
function getMessagesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Messages');
  if (!sheet) {
    sheet = ss.insertSheet('Messages');
    sheet.appendRow(['MessageId', 'Target', 'Message', 'CreatedAt', 'Active']);
  }
  return sheet;
}
function getRequestsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
    sheet.appendRow(['RequestId', 'Email', 'Name', 'Type', 'FundName', 'Amount', 'Notes', 'Status', 'CreatedAt']);
  }
  return sheet;
}

/**
 * Fetches AMFI's full daily NAV list (https://www.amfiindia.com/spages/NAVAll.txt),
 * parses every scheme row, and rewrites the "Funds" tab with the latest data.
 */
function refreshFundList() {
  const url = 'https://www.amfiindia.com/spages/NAVAll.txt';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'Could not reach AMFI (HTTP ' + res.getResponseCode() + ')' };
  }
  const lines = res.getContentText().split('\n');
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line.indexOf(';') === -1) continue;
    const parts = line.split(';');
    if (parts.length < 5) continue;
    const code = parts[0].trim();
    const name = parts[3].trim();
    const navRaw = parts[4].trim();
    if (!code || !name || isNaN(Number(code))) continue;
    const nav = isNaN(Number(navRaw)) ? '' : Number(navRaw);
    rows.push([code, name, nav]);
  }
  const sheet = getFundsSheet();
  sheet.clearContents();
  const now = new Date();
  sheet.appendRow(['SchemeCode', 'SchemeName', 'NAV', 'LastRefreshed']);
  if (rows.length > 0) {
    const withTimestamp = rows.map(r => [r[0], r[1], r[2], now]);
    sheet.getRange(2, 1, withTimestamp.length, 4).setValues(withTimestamp);
  }
  return { ok: true, count: rows.length };
}

/**
 * Computes the TRUE present value of a SIP: simulates every installment against the
 * scheme's actual historical NAV on that date (via mfapi.in, a free public historical
 * NAV source for Indian mutual funds), and sums the real units bought each time.
 * This replaces guessing from a manually-entered "units held" number.
 * Cached per scheme+SIP terms for 6 hours since historical NAV only changes once a day.
 */
function advanceDate_(d, frequency) {
  if (frequency === 'Daily') d.setDate(d.getDate() + 1);
  else if (frequency === 'Weekly') d.setDate(d.getDate() + 7);
  else if (frequency === 'Quarterly') d.setMonth(d.getMonth() + 3);
  else d.setMonth(d.getMonth() + 1); // Monthly, and default
  return d;
}

/**
 * Two supported modes:
 *  1) Pure auto mode — give it schemeCode/startDate/frequency/amount and it walks every
 *     installment date from startDate to today, buying at that day's REAL historical NAV
 *     (via mfapi.in), and sums the units. This is the "system calculates everything" path.
 *  2) Hybrid / as-of mode — for a SIP that was already running before the investor started
 *     using this tool, they can instead supply asOfDate + asOfInvested (₹ invested till that
 *     date) + asOfMarketValue (₹ it was worth on that date, e.g. from a CAS/statement). The
 *     as-of market value is converted into units at that date's real NAV, and the engine then
 *     only simulates NEW installments from the day after asOfDate through today — so the
 *     portfolio keeps updating itself automatically as each new SIP date passes, without
 *     needing to re-derive the entire historical purchase trail.
 */
// Public single-SIP entry point — kept for backward compatibility with any existing callers.
// Internally just delegates to the shared core so there's one implementation to maintain.
function computeSipAccumulation(body) {
  return computeSipAccumulationCore_(body);
}

/**
 * Batched sibling of computeSipAccumulation: takes body.sips = [{ id, schemeCode, startDate,
 * frequency, amount, asOfDate, asOfInvested, asOfMarketValue }, ...] and returns results for
 * ALL of them in a single request/response cycle.
 *
 * PERF NOTE: this exists because the old Dashboard code called computeSipAccumulation once
 * PER SIP — an investor with 6 SIPs paid for 6 separate network round-trips to this Web App
 * just to open their own Dashboard, each with its own HTTP overhead on top of the actual
 * work. That's a real, measurable slowdown, and it has nothing to do with login/password
 * security — it's simply how many times the browser has to talk to Apps Script. Batching
 * into one call removes that overhead entirely (the per-scheme historical-NAV lookups are
 * still individually cached exactly as before).
 */
function computeSipAccumulationBatch(body) {
  const sips = body.sips || [];
  const results = {};
  sips.forEach(s => {
    const key = s.id || (s.schemeCode + '_' + s.startDate);
    try {
      results[key] = computeSipAccumulationCore_(s);
    } catch (e) {
      results[key] = { ok: false, error: e.message };
    }
  });
  return { ok: true, results: results };
}

function computeSipAccumulationCore_(body) {
  const schemeCode = body.schemeCode;
  const startDate = body.startDate;
  const frequency = body.frequency;
  const amount = Number(body.amount || 0);
  if (!schemeCode || !startDate || !amount) return { ok: false, error: 'Missing schemeCode/startDate/amount' };

  const asOfDate = body.asOfDate || null;
  const asOfInvestedRaw = body.asOfInvested;
  const asOfMarketValueRaw = body.asOfMarketValue;
  const asOfInvested = (asOfInvestedRaw !== null && asOfInvestedRaw !== undefined && asOfInvestedRaw !== '') ? Number(asOfInvestedRaw) : null;
  const asOfMarketValue = (asOfMarketValueRaw !== null && asOfMarketValueRaw !== undefined && asOfMarketValueRaw !== '') ? Number(asOfMarketValueRaw) : null;
  const useAsOf = !!(asOfDate && asOfMarketValue !== null && !isNaN(asOfMarketValue) && !isNaN(new Date(asOfDate).getTime()));

  const cacheKey = 'sipacc_' + schemeCode + '_' + startDate + '_' + frequency + '_' + amount + '_' +
    (useAsOf ? ('asof_' + asOfDate + '_' + asOfInvested + '_' + asOfMarketValue + '_') : '') +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Etc/UTC', 'yyyy-MM-dd');
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let history;
  try {
    const res = UrlFetchApp.fetch('https://api.mfapi.in/mf/' + schemeCode, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, error: 'Historical NAV service unavailable (HTTP ' + res.getResponseCode() + ')' };
    history = JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, error: 'Could not reach historical NAV service: ' + e.message };
  }
  const series = (history.data || []).map(d => ({ date: parseMfapiDate_(d.date), nav: Number(d.nav) }))
    .filter(d => d.date && !isNaN(d.nav))
    .sort((a, b) => a.date - b.date); // ascending
  if (series.length === 0) return { ok: false, error: 'No historical NAV data for this scheme' };

  const navOnOrBefore = (targetDate) => {
    // series is ascending; find the latest entry on or before targetDate (markets closed weekends/holidays)
    let lo = 0, hi = series.length - 1, ans = null;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (series[mid].date <= targetDate) { ans = series[mid]; lo = mid + 1; }
      else hi = mid - 1;
    }
    return ans ? ans.nav : series[0].nav;
  };

  const today = new Date(); today.setHours(0, 0, 0, 0);

  let openingUnits = 0, openingInvested = 0, simulateFrom;
  if (useAsOf) {
    const asOf = new Date(asOfDate);
    asOf.setHours(0, 0, 0, 0);
    const navAsOf = navOnOrBefore(asOf);
    openingUnits = navAsOf > 0 ? (asOfMarketValue / navAsOf) : 0;
    openingInvested = (asOfInvested !== null && !isNaN(asOfInvested)) ? asOfInvested : asOfMarketValue;
    simulateFrom = advanceDate_(new Date(asOf), frequency); // first NEW installment is the one after the snapshot date
  } else {
    simulateFrom = new Date(startDate);
  }

  let totalUnits = openingUnits, installments = 0, d = new Date(simulateFrom), guard = 0;
  while (d <= today && guard < 3000) {
    const nav = navOnOrBefore(d);
    if (nav > 0) { totalUnits += amount / nav; installments++; }
    advanceDate_(d, frequency);
    guard++;
  }
  const latestNav = series[series.length - 1].nav;
  const result = {
    ok: true, totalUnits, installments, totalInvested: openingInvested + installments * amount,
    openingUnits, openingInvested, usedAsOf: useAsOf,
    latestNav, latestDate: Utilities.formatDate(series[series.length - 1].date, Session.getScriptTimeZone() || 'Etc/UTC', 'yyyy-MM-dd')
  };
  cache.put(cacheKey, JSON.stringify(result), 6 * 3600);
  return result;
}

function parseMfapiDate_(str) {
  // mfapi.in dates come as "dd-mm-yyyy"
  const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(str || '');
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

// Full historical NAV series for one scheme, for the Dashboard's per-scheme "NAV Performance"
// chart. Proxied server-side (same reason as everything else that talks to mfapi.in/AMFI in
// this file) so it works from any investor's browser without a CORS round-trip depending on
// mfapi.in's own headers. No login required — NAV history isn't personal data — matching
// getMarketIndices below.
function getNavHistory(body) {
  const schemeCode = String(body.schemeCode || '').trim();
  if (!schemeCode) return { ok: false, error: 'schemeCode is required' };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'navhist_' + schemeCode;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let raw;
  try {
    const res = UrlFetchApp.fetch('https://api.mfapi.in/mf/' + encodeURIComponent(schemeCode), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, error: 'Historical NAV service unavailable (HTTP ' + res.getResponseCode() + ')' };
    raw = JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, error: 'Could not reach historical NAV service: ' + e.message };
  }

  const parsed = (raw.data || [])
    .map(d => ({ date: parseMfapiDate_(d.date), nav: Number(d.nav) }))
    .filter(d => d.date && !isNaN(d.nav))
    .sort((a, b) => a.date - b.date); // mfapi.in returns newest-first; the chart wants ascending
  if (parsed.length === 0) return { ok: false, error: 'No historical NAV data found for this scheme' };

  // Cap to the most recent ~1100 trading days (a little over 4 years) so the cached JSON
  // comfortably fits CacheService's 100KB-per-value limit while still covering every period
  // the Dashboard's chart offers (1M/6M/1Y/3Y/All-since-cap).
  const capped = parsed.length > 1100 ? parsed.slice(parsed.length - 1100) : parsed;
  const tz = Session.getScriptTimeZone() || 'Etc/UTC';
  const out = {
    ok: true,
    schemeCode: schemeCode,
    schemeName: (raw.meta && raw.meta.scheme_name) || '',
    data: capped.map(d => ({ date: Utilities.formatDate(d.date, tz, 'yyyy-MM-dd'), nav: d.nav }))
  };
  try { cache.put(cacheKey, JSON.stringify(out), 900); } catch (e) { /* payload too large to cache — fine, just skip caching */ }
  return out;
}

// IMPORTANT (perf): this used to also return true once the list was >24h old, which meant
// an ordinary investor opening the Dashboard could randomly be the one who pays for a
// synchronous 10-20 second AMFI fetch-and-rewrite of ~10,000+ scheme rows. That had nothing
// to do with login/password security — it was just a slow network call sitting in the
// request path. Now a live request only forces a synchronous refresh if the Funds tab has
// never been populated at all; a merely-stale list is still served instantly (a search/NAV
// lookup a day old is a fine tradeoff for a page that opens instantly). Use
// setupDailyFundRefreshTrigger() (Admin Console > "Enable automatic daily refresh") so the
// list refreshes itself in the background at a fixed time and effectively never goes stale
// in the first place.
function fundsNeedRefresh() {
  const sheet = getFundsSheet();
  return sheet.getLastRow() < 2;
}

function fundsAreStale_() {
  const sheet = getFundsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return true;
  const lastRefreshed = sheet.getRange(2, 4).getValue();
  if (!lastRefreshed) return true;
  const ageHours = (Date.now() - new Date(lastRefreshed).getTime()) / 36e5;
  return ageHours > 24;
}

/**
 * Run once (Admin Console > AMFI Fund List > "Enable automatic daily refresh", or from the
 * Apps Script editor function dropdown) to refresh the Funds tab automatically every night —
 * so no live investor request ever has to pay for the AMFI fetch itself. Safe to re-run.
 */
function setupDailyFundRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshFundListIfStale_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshFundListIfStale_').timeBased().everyDays(1).atHour(3).create();
  return { ok: true, message: 'Fund list will now refresh automatically in the background (~3 AM daily).' };
}

function refreshFundListIfStale_() {
  if (fundsAreStale_()) refreshFundList();
}

function searchFunds(query) {
  if (!query || query.trim().length < 2) return { ok: true, funds: [] };
  if (fundsNeedRefresh()) {
    const r = refreshFundList();
    if (!r.ok) return r;
  }
  const sheet = getFundsSheet();
  const data = sheet.getDataRange().getValues();
  const q = query.trim().toLowerCase();
  const matches = [];
  for (let i = 1; i < data.length && matches.length < 40; i++) {
    const name = (data[i][1] || '').toString();
    if (name.toLowerCase().indexOf(q) !== -1) {
      matches.push({ name: name, nav: data[i][2] || '', code: data[i][0] || '' });
    }
  }
  return { ok: true, funds: matches };
}

// =====================================================================================
// Password hashing — PBKDF2-HMAC-SHA256, salted, ~10,000 iterations.
//
// The old scheme (see legacyHashPw_ below) was one unsalted SHA-256 round: fast enough
// that a stolen PasswordHash column could be brute-forced at billions of guesses/second
// on a GPU, and identical passwords produced identical hashes (no salt). PBKDF2 with a
// random salt per user and many iterations makes brute-forcing thousands of times slower
// and defeats precomputed rainbow tables.
//
// Apps Script has no native bcrypt/scrypt/Argon2, so this hand-rolls PBKDF2 on top of
// Utilities.computeHmacSha256Signature. That function's Byte[] values are SIGNED
// (-128..127, Java-style) — toSigned_/toUnsigned_ below convert to/from the normal 0..255
// range used everywhere else here. After pasting this file in, run runPbkdf2SelfTest_()
// once from the Apps Script editor (function dropdown > select it > Run) — it checks the
// output against a published PBKDF2 test vector and throws if anything is wrong, so you
// find out before any real password depends on it.
// =====================================================================================

const PBKDF2_ITERATIONS = 10000;
const PBKDF2_KEY_BYTES = 32;
const PBKDF2_SALT_BYTES = 16;

function toSigned_(b) { return b > 127 ? b - 256 : b; }
function toUnsigned_(b) { return b & 0xff; }
function bytesToSigned_(bytes) { return bytes.map(toSigned_); }
function bytesToUnsigned_(bytes) { return bytes.map(toUnsigned_); }

function hmacSha256Bytes_(valueBytesUnsigned, keyBytesUnsigned) {
  const sig = Utilities.computeHmacSha256Signature(bytesToSigned_(valueBytesUnsigned), bytesToSigned_(keyBytesUnsigned));
  return bytesToUnsigned_(sig);
}
function intToBytes4_(i) { return [(i >>> 24) & 0xff, (i >>> 16) & 0xff, (i >>> 8) & 0xff, i & 0xff]; }
function xorBytes_(a, b) { return a.map((v, idx) => v ^ b[idx]); }

function pbkdf2Sha256_(passwordBytes, saltBytes, iterations, keyLenBytes) {
  const hLen = 32;
  const blocks = Math.ceil(keyLenBytes / hLen);
  let dk = [];
  for (let i = 1; i <= blocks; i++) {
    let u = hmacSha256Bytes_(saltBytes.concat(intToBytes4_(i)), passwordBytes);
    let t = u.slice();
    for (let j = 1; j < iterations; j++) {
      u = hmacSha256Bytes_(u, passwordBytes);
      t = xorBytes_(t, u);
    }
    dk = dk.concat(t);
  }
  return dk.slice(0, keyLenBytes);
}

// Utilities has no direct CSPRNG byte generator, but getUuid() is backed by a secure
// random source; stitching enough UUIDs together and truncating gives cryptographically
// random bytes.
function randomBytes_(n) {
  let bytes = [];
  while (bytes.length < n) {
    const hex = Utilities.getUuid().replace(/-/g, '');
    for (let i = 0; i < hex.length; i += 2) bytes.push(parseInt(hex.substr(i, 2), 16));
  }
  return bytes.slice(0, n);
}

function bytesToBase64_(bytes) { return Utilities.base64Encode(bytesToSigned_(bytes)); }
function base64ToBytes_(b64) { return bytesToUnsigned_(Utilities.base64Decode(b64)); }
function strToUtf8Bytes_(s) { return bytesToUnsigned_(Utilities.newBlob(s).getBytes()); }

// New-format stored hash: "pbkdf2$<iterations>$<saltBase64>$<hashBase64>". The iteration
// count travels WITH the hash (rather than being assumed from a global constant) so
// PBKDF2_ITERATIONS can be raised later without invalidating hashes created under the old value.
function makePasswordHash_(password) {
  const salt = randomBytes_(PBKDF2_SALT_BYTES);
  const dk = pbkdf2Sha256_(strToUtf8Bytes_(password), salt, PBKDF2_ITERATIONS, PBKDF2_KEY_BYTES);
  return 'pbkdf2$' + PBKDF2_ITERATIONS + '$' + bytesToBase64_(salt) + '$' + bytesToBase64_(dk);
}

// The original hashing scheme (unsalted SHA-256 hex digest) — kept only so existing
// accounts' stored hashes can still be verified once, to transparently upgrade them.
function legacyHashPw_(pw) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  return digest.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

function timingSafeEqual_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const len = Math.max(a.length, b.length);
  let diff = a.length === b.length ? 0 : 1;
  for (let i = 0; i < len; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

// Verifies a password against a stored hash of either format. needsUpgrade is true only
// for a successfully-verified LEGACY hash, so callers can transparently re-hash with the
// new scheme on that login — no forced password reset for anyone.
function verifyPassword_(password, storedHash) {
  if (!storedHash) return { ok: false, needsUpgrade: false };
  if (storedHash.indexOf('pbkdf2$') === 0) {
    const parts = storedHash.split('$');
    const iterations = Number(parts[1]);
    const salt = base64ToBytes_(parts[2]);
    const dk = pbkdf2Sha256_(strToUtf8Bytes_(password), salt, iterations, PBKDF2_KEY_BYTES);
    return { ok: timingSafeEqual_(bytesToBase64_(dk), parts[3]), needsUpgrade: false };
  }
  const ok = timingSafeEqual_(legacyHashPw_(password), storedHash);
  return { ok: ok, needsUpgrade: ok };
}

// Run this once from the Apps Script editor after pasting this file in (function dropdown
// > runPbkdf2SelfTest_ > Run). Confirms PBKDF2 produces correct, known-good output on this
// account's runtime before any real password depends on it. Check View > Logs afterward —
// it throws (visible as a red error) if anything doesn't match.
function runPbkdf2SelfTest_() {
  const hex1 = pbkdf2Sha256_(strToUtf8Bytes_('password'), strToUtf8Bytes_('salt'), 1, 32).map(b => b.toString(16).padStart(2, '0')).join('');
  if (hex1 !== '120fb6cffcf8b32c43e7225256c4f837a86548c92ccc35480805987cb70be17b') throw new Error('PBKDF2 self-test FAILED (1 iter): ' + hex1);
  const hex2 = pbkdf2Sha256_(strToUtf8Bytes_('password'), strToUtf8Bytes_('salt'), 2000, 32).map(b => b.toString(16).padStart(2, '0')).join('');
  if (hex2 !== '9209a0c90243e88b89488f99cd7ea010c244cc7a9d4bf65c157f2d8f642eb952') throw new Error('PBKDF2 self-test FAILED (2000 iter): ' + hex2);
  const hash = makePasswordHash_('correct horse battery staple');
  if (!verifyPassword_('correct horse battery staple', hash).ok) throw new Error('PBKDF2 self-test FAILED: round-trip hash/verify mismatch');
  if (verifyPassword_('wrong password', hash).ok) throw new Error('PBKDF2 self-test FAILED: an incorrect password verified as correct');
  Logger.log('PBKDF2 self-test passed.');
}

// =====================================================================================
// Sessions — issued on login/signup, checked on every other action. Replaces the old
// design where the browser held the plaintext password in memory/sessionStorage and
// resent it with every single request. A token is a random opaque bearer credential
// (~244 bits of randomness from two stitched-together secure UUIDs) that only maps to a
// session inside this script's properties — it reveals nothing about the password.
//
// Sessions do NOT expire on their own: a token stays valid until the user logs out, or
// until an admin blocks that account (resolveSession_ checks Blocked on every use and
// invalidates immediately if so). Stored in PropertiesService rather than CacheService
// specifically because CacheService entries are capped at 6 hours no matter what.
// =====================================================================================

function createSession_(email, isAdvisor) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('sess_' + token, JSON.stringify({ email: email, isAdvisor: !!isAdvisor, createdAt: Date.now() }));
  return token;
}

// Resolves a token to its session (or null if missing, or the account is blocked),
// keeping isAdvisor fresh in case admin rights changed since login.
function resolveSession_(token) {
  if (!token) return null;
  const props = PropertiesService.getScriptProperties();
  const key = 'sess_' + token;
  const raw = props.getProperty(key);
  if (!raw) return null;
  const sess = JSON.parse(raw);
  const u = findRow(getUsersSheet(), sess.email);
  if (!u || isBlocked(u.data)) { props.deleteProperty(key); return null; }
  const freshIsAdvisor = (u.data[2] === true || u.data[2] === 'TRUE');
  if (freshIsAdvisor !== sess.isAdvisor) {
    sess.isAdvisor = freshIsAdvisor;
    props.setProperty(key, JSON.stringify(sess));
  }
  return sess;
}

function destroySession_(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
}

function logout(token) {
  destroySession_(token);
  return { ok: true };
}

// Lets the client silently check "am I still signed in?" on page load using only the
// token it already has — it no longer has the password to re-login with, by design.
function whoami(token) {
  const sess = resolveSession_(token);
  if (!sess) return { ok: false, error: 'Not signed in' };
  return { ok: true, email: sess.email, isAdvisor: sess.isAdvisor };
}

// =====================================================================================
// Login rate-limiting — 5 failed attempts on one account locks it out for 15 minutes.
// Scoped by email (not IP, which Apps Script doesn't expose) since the threat this
// defends against is a specific account's password being guessed/stuffed, which is what
// matters most here — a single admin account guards every investor's data.
// =====================================================================================

const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_LOCKOUT_SECONDS = 15 * 60;

function loginAttemptKey_(email) { return 'loginfail_' + (email || '').toLowerCase().trim(); }

function checkLoginLockout_(email) {
  const raw = CacheService.getScriptCache().get(loginAttemptKey_(email));
  if (!raw) return { locked: false };
  const info = JSON.parse(raw);
  const remainingSec = Math.max(0, Math.round((info.lockedUntil - Date.now()) / 1000));
  return (info.count >= LOGIN_MAX_ATTEMPTS && remainingSec > 0) ? { locked: true, remainingSec: remainingSec } : { locked: false };
}

function recordLoginFailure_(email) {
  const cache = CacheService.getScriptCache();
  const key = loginAttemptKey_(email);
  const raw = cache.get(key);
  const info = raw ? JSON.parse(raw) : { count: 0 };
  info.count++;
  info.lockedUntil = Date.now() + LOGIN_LOCKOUT_SECONDS * 1000;
  cache.put(key, JSON.stringify(info), LOGIN_LOCKOUT_SECONDS);
}

function clearLoginFailures_(email) { CacheService.getScriptCache().remove(loginAttemptKey_(email)); }

function findRow(sheet, email) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase() === (email || '').toLowerCase()) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

function isBlocked(userRowData) {
  const v = userRowData[3];
  return v === true || v === 'TRUE';
}

function signup(email, password) {
  if (!email || !password) return { ok: false, error: 'Email and password required' };
  if (password.length < 8) return { ok: false, error: 'Password must be at least 8 characters' };
  const users = getUsersSheet();
  if (findRow(users, email)) return { ok: false, error: 'An account with that email already exists' };
  users.appendRow([email, makePasswordHash_(password), false, false]);
  getClientsSheet().appendRow([email, '', '', '', '', '', '', '[]', new Date(), '[]']);
  const token = createSession_(email, false);
  return { ok: true, isAdvisor: false, token: token };
}

function login(email, password) {
  const lockout = checkLoginLockout_(email);
  if (lockout.locked) {
    const mins = Math.ceil(lockout.remainingSec / 60);
    return { ok: false, error: 'Too many failed attempts on this account. Try again in about ' + mins + ' minute' + (mins === 1 ? '' : 's') + '.' };
  }
  const u = findRow(getUsersSheet(), email);
  if (!u) return { ok: false, error: 'No account found for that email' };
  const check = verifyPassword_(password, u.data[1]);
  if (!check.ok) {
    recordLoginFailure_(email);
    return { ok: false, error: 'Incorrect password' };
  }
  if (isBlocked(u.data)) return { ok: false, error: 'This account has been blocked. Please contact your advisor.' };
  clearLoginFailures_(email);
  if (check.needsUpgrade) {
    // Transparent migration from the old unsalted-SHA-256 scheme — no action needed from
    // the user, their next login already uses the stronger hash.
    getUsersSheet().getRange(u.row, 2).setValue(makePasswordHash_(password));
  }
  const isAdvisor = (u.data[2] === true || u.data[2] === 'TRUE');
  const token = createSession_(email, isAdvisor);
  return { ok: true, isAdvisor: isAdvisor, token: token };
}

// Replaces the old password-per-request check. email is the caller's claimed identity;
// token must resolve to a live session for that SAME email — this ties every action back
// to a real login rather than trusting a client-supplied email string on its own.
function authOk(email, token) {
  const sess = resolveSession_(token);
  return !!sess && sess.email.toLowerCase() === (email || '').toLowerCase();
}

// Verifies the caller holds a live, non-blocked session with IsAdvisor = TRUE.
function requireAdmin(email, token) {
  const sess = resolveSession_(token);
  if (!sess) return { ok: false, error: 'Not signed in — please sign in again.' };
  if (sess.email.toLowerCase() !== (email || '').toLowerCase()) return { ok: false, error: 'Not authenticated' };
  if (!sess.isAdvisor) return { ok: false, error: 'Super-user rights required for this action' };
  return { ok: true };
}

// Wraps a zero-argument admin-only action (e.g. refreshFundList, setupDailyReminderTrigger)
// behind a requireAdmin() check, so it can't be invoked by an unauthenticated caller who
// simply POSTs the action name directly to the Apps Script URL.
function requireAdminAction(body, fn) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  return fn();
}

function clientRowFor(sheet, email) {
  // Ensures 10 columns even for rows created before HoldingsJSON existed.
  const found = findRow(sheet, email);
  if (!found) return null;
  const d = found.data.slice();
  while (d.length < 10) d.push(d.length === 7 ? '[]' : (d.length === 9 ? '[]' : ''));
  return { row: found.row, data: d };
}

function saveData(body) {
  if (!authOk(body.email, body.token)) return { ok: false, error: 'Not authenticated' };
  const sheet = getClientsSheet();
  const inv = body.investor || {};
  const risk = body.riskAssessment || {};
  const existing = clientRowFor(sheet, body.email);
  const holdingsJson = existing ? (existing.data[9] || '[]') : '[]';
  const row = [
    body.email, inv.name || '', inv.phone || '', inv.pan || '', inv.dob || '',
    risk.score || '', risk.date || '', JSON.stringify(body.sips || []), new Date(), holdingsJson
  ];
  if (existing) sheet.getRange(existing.row, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { ok: true };
}

function loadData(email, token) {
  if (!authOk(email, token)) return { ok: false, error: 'Not authenticated' };
  const c = clientRowFor(getClientsSheet(), email);
  if (!c) return { ok: true, data: null };
  const d = c.data;
  return {
    ok: true,
    data: {
      investor: { name: d[1], email: email, phone: d[2], pan: d[3], dob: d[4] },
      riskAssessment: d[5] ? { score: Number(d[5]), date: d[6] } : null,
      sips: d[7] ? JSON.parse(d[7]) : [],
      holdings: d[9] ? JSON.parse(d[9]) : []
    }
  };
}

/**
 * Returns the full client roster for an admin: profile, SIPs, CAS holdings, and account status.
 * Used to power the Advisor/Admin Dashboard, the "select any user" dropdowns, the AUM tab and drill-downs.
 */
function listClients(email, token) {
  const admin = requireAdmin(email, token);
  if (!admin.ok) return admin;
  const users = getUsersSheet().getDataRange().getValues();
  const blockedByEmail = {};
  const advisorByEmail = {};
  for (let i = 1; i < users.length; i++) {
    const key = (users[i][0] || '').toString().toLowerCase();
    blockedByEmail[key] = isBlocked(users[i]);
    advisorByEmail[key] = (users[i][2] === true || users[i][2] === 'TRUE');
  }
  const data = getClientsSheet().getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    const key = d[0].toString().toLowerCase();
    clients.push({
      email: d[0],
      name: d[1],
      phone: d[2],
      pan: d[3],
      dob: d[4] || '',
      riskScore: d[5] || null,
      sips: d[7] ? JSON.parse(d[7]) : [],
      holdings: d[9] ? JSON.parse(d[9]) : [],
      blocked: !!blockedByEmail[key],
      isAdvisor: !!advisorByEmail[key]
    });
  }
  return { ok: true, clients };
}

/** Super-user: add a SIP to any investor's register. */
function adminAddSip(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.targetEmail);
  if (!existing) return { ok: false, error: 'No such investor account' };
  const sips = existing.data[7] ? JSON.parse(existing.data[7]) : [];
  sips.push(body.sip);
  sheet.getRange(existing.row, 8).setValue(JSON.stringify(sips));
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true };
}

/** Super-user: update the full details of an investor's existing SIP (matched by body.sip.id). */
function adminUpdateSip(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.targetEmail);
  if (!existing) return { ok: false, error: 'No such investor account' };
  const sips = existing.data[7] ? JSON.parse(existing.data[7]) : [];
  const idx = sips.findIndex(s => s.id === (body.sip || {}).id);
  if (idx === -1) return { ok: false, error: 'SIP not found — it may have already been changed elsewhere. Refresh and try again.' };
  sips[idx] = body.sip;
  sheet.getRange(existing.row, 8).setValue(JSON.stringify(sips));
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true };
}

/** Super-user: delete one of an investor's SIPs (matched by body.sipId). */
function adminDeleteSip(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.targetEmail);
  if (!existing) return { ok: false, error: 'No such investor account' };
  const sips = existing.data[7] ? JSON.parse(existing.data[7]) : [];
  const next = sips.filter(s => s.id !== body.sipId);
  if (next.length === sips.length) return { ok: false, error: 'SIP not found — it may have already been removed. Refresh and try again.' };
  sheet.getRange(existing.row, 8).setValue(JSON.stringify(next));
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true };
}

/**
 * Self-service: an investor uploads/appends their OWN CAS holdings (no admin rights needed).
 * Same shape as adminUploadCas but scoped to the caller's own account.
 */
function uploadOwnCas(body) {
  if (!authOk(body.email, body.token)) return { ok: false, error: 'Not authenticated' };
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.email);
  if (!existing) return { ok: false, error: 'Account not found' };
  let holdings = (body.mode === 'append') ? (existing.data[9] ? JSON.parse(existing.data[9]) : []) : [];
  const incoming = (body.holdings || []).map(h => ({
    id: 'h_' + Utilities.getUuid().slice(0, 8),
    fundName: h.fundName || '',
    folio: h.folio || '',
    schemeType: h.schemeType || 'Equity',
    units: Number(h.units || 0),
    nav: Number(h.nav || 0),
    investedValue: Number(h.investedValue || 0)
  }));
  holdings = holdings.concat(incoming);
  sheet.getRange(existing.row, 10).setValue(JSON.stringify(holdings));
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true, count: incoming.length };
}

/** Super-user: upload/append CAS (Consolidated Account Statement) holdings on behalf of any investor.
 * Expects body.holdings = [{ fundName, folio, schemeType, units, nav, investedValue }]
 * mode: "replace" (default) overwrites the investor's holdings list, "append" adds to it.
 */
function adminUploadCas(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.targetEmail);
  if (!existing) return { ok: false, error: 'No such investor account' };
  let holdings = (body.mode === 'append') ? (existing.data[9] ? JSON.parse(existing.data[9]) : []) : [];
  const incoming = (body.holdings || []).map(h => ({
    id: 'h_' + Utilities.getUuid().slice(0, 8),
    fundName: h.fundName || '',
    folio: h.folio || '',
    schemeType: h.schemeType || 'Equity',
    units: Number(h.units || 0),
    nav: Number(h.nav || 0),
    investedValue: Number(h.investedValue || 0)
  }));
  holdings = holdings.concat(incoming);
  sheet.getRange(existing.row, 10).setValue(JSON.stringify(holdings));
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true, count: incoming.length };
}

/**
 * Self-service: an investor wipes their OWN CAS holdings clean (units/folios/values from
 * old statements), so a fresh CAS PDF can be uploaded without old and new rows mixing
 * together. Does NOT touch SIPs, risk profile, or personal details — holdings only.
 */
function flushHoldings(body) {
  if (!authOk(body.email, body.token)) return { ok: false, error: 'Not authenticated' };
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.email);
  if (!existing) return { ok: false, error: 'Account not found' };
  sheet.getRange(existing.row, 10).setValue('[]');
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true };
}

/** Super-user: wipe any investor's CAS holdings clean (e.g. before uploading a corrected
 * CAS on their behalf). Does NOT touch SIPs, risk profile, or personal details. */
function adminFlushHoldings(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getClientsSheet();
  const existing = clientRowFor(sheet, body.targetEmail);
  if (!existing) return { ok: false, error: 'No such investor account' };
  sheet.getRange(existing.row, 10).setValue('[]');
  sheet.getRange(existing.row, 9).setValue(new Date());
  return { ok: true };
}

/** Super-user: change any investor's password. */
function adminChangePassword(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  if (!body.newPassword || body.newPassword.length < 8) return { ok: false, error: 'New password must be at least 8 characters' };
  const users = getUsersSheet();
  const u = findRow(users, body.targetEmail);
  if (!u) return { ok: false, error: 'No such user account' };
  users.getRange(u.row, 2).setValue(makePasswordHash_(body.newPassword));
  return { ok: true };
}

/** Super-user: block or unblock any investor's ability to log in. */
function adminSetBlocked(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const users = getUsersSheet();
  const u = findRow(users, body.targetEmail);
  if (!u) return { ok: false, error: 'No such user account' };
  if ((u.data[0] || '').toString().toLowerCase() === (body.adminEmail || '').toLowerCase()) {
    return { ok: false, error: "You can't block your own account" };
  }
  users.getRange(u.row, 4).setValue(!!body.blocked);
  return { ok: true };
}

/** Super-user: create a brand-new user account (investor or advisor) directly — they can
 * sign in immediately with the password set here, no separate self-signup needed. */
function adminCreateUser(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const email = (body.targetEmail || '').trim();
  const password = body.newPassword || '';
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'A valid email is required' };
  if (!password || password.length < 8) return { ok: false, error: 'Password must be at least 8 characters' };
  const users = getUsersSheet();
  if (findRow(users, email)) return { ok: false, error: 'An account with that email already exists' };
  users.appendRow([email, makePasswordHash_(password), !!body.isAdvisor, false]);
  getClientsSheet().appendRow([email, body.name || '', body.phone || '', body.pan || '', body.dob || '', '', '', '[]', new Date(), '[]']);
  return { ok: true };
}

/** Super-user: edit an existing user's profile fields (name/phone/PAN/DOB) and/or their
 * super-user (IsAdvisor) rights. Deliberately does NOT change the email address itself —
 * it's the key every other sheet, session, message and request references, so renaming it
 * safely would mean rewriting all of those; delete and recreate the account instead if an
 * email was entered wrong. */
function adminUpdateUser(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const targetEmail = (body.targetEmail || '').trim();
  const users = getUsersSheet();
  const u = findRow(users, targetEmail);
  if (!u) return { ok: false, error: 'No such user account' };
  if (body.isAdvisor !== undefined) {
    const isSelf = targetEmail.toLowerCase() === (body.adminEmail || '').toLowerCase();
    if (isSelf && !body.isAdvisor) return { ok: false, error: "You can't remove your own super-user rights" };
    users.getRange(u.row, 3).setValue(!!body.isAdvisor);
  }
  const clientsSheet = getClientsSheet();
  const c = clientRowFor(clientsSheet, targetEmail);
  if (c) {
    const row = c.data.slice();
    if (body.name !== undefined) row[1] = body.name;
    if (body.phone !== undefined) row[2] = body.phone;
    if (body.pan !== undefined) row[3] = body.pan;
    if (body.dob !== undefined) row[4] = body.dob;
    clientsSheet.getRange(c.row, 1, 1, row.length).setValues([row]);
  }
  return { ok: true };
}

/** Super-user: permanently delete a user account — removes their Users row and their entire
 * Clients row (profile, SIPs, holdings, risk data), plus any banner messages addressed
 * specifically to them and any purchase requests they submitted. Cannot delete your own
 * account. IRREVERSIBLE — the client confirms with the investor's email before calling this.
 */
function adminDeleteUser(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const targetEmail = (body.targetEmail || '').trim();
  if (targetEmail.toLowerCase() === (body.adminEmail || '').toLowerCase()) {
    return { ok: false, error: "You can't delete your own account" };
  }
  const users = getUsersSheet();
  const u = findRow(users, targetEmail);
  if (!u) return { ok: false, error: 'No such user account' };
  users.deleteRow(u.row);
  const clients = getClientsSheet();
  const c = findRow(clients, targetEmail);
  if (c) clients.deleteRow(c.row);
  const messagesRemoved = deleteMessagesForUser_(targetEmail);
  const requestsRemoved = deleteRequestsForUser_(targetEmail);
  // Any live session tokens for this email stop working on their very next use — resolveSession_
  // looks the email up in the Users sheet on every call and treats "not found" the same as blocked.
  return { ok: true, messagesRemoved, requestsRemoved };
}

/** Super-user: wipe everything ASSOCIATED with a user — portfolio (SIPs + CAS holdings),
 * risk profile, banner messages addressed to them, and their purchase requests — while
 * leaving their login account itself intact so they can still sign in to a clean slate.
 * Use adminDeleteUser instead if the account itself should be removed too. */
function adminFlushUserData(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const targetEmail = (body.targetEmail || '').trim();
  const clients = getClientsSheet();
  const c = clientRowFor(clients, targetEmail);
  if (!c) return { ok: false, error: 'No such investor account' };
  const row = c.data.slice();
  row[5] = '';       // RiskScore
  row[6] = '';       // RiskDate
  row[7] = '[]';     // SIPsJSON
  row[8] = new Date();
  row[9] = '[]';     // HoldingsJSON
  clients.getRange(c.row, 1, 1, row.length).setValues([row]);
  const messagesRemoved = deleteMessagesForUser_(targetEmail);
  const requestsRemoved = deleteRequestsForUser_(targetEmail);
  return { ok: true, messagesRemoved, requestsRemoved };
}

// Removes every banner message addressed specifically to this email (leaves firm-wide "ALL"
// messages and any addressed to other investors untouched). Returns how many were removed.
function deleteMessagesForUser_(email) {
  const sheet = getMessagesSheet();
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][1] || '').toString().toLowerCase() === (email || '').toLowerCase()) { sheet.deleteRow(i + 1); count++; }
  }
  return count;
}

// Removes every New Purchase Request submitted by this email. Returns how many were removed.
function deleteRequestsForUser_(email) {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][1] || '').toString().toLowerCase() === (email || '').toLowerCase()) { sheet.deleteRow(i + 1); count++; }
  }
  return count;
}

/** Super-user: post a banner message to one investor's email, or 'ALL' for everyone. */
function adminBroadcast(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  if (!body.message || !body.message.trim()) return { ok: false, error: 'Message text is required' };
  const sheet = getMessagesSheet();
  const id = 'm_' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, body.target || 'ALL', body.message.trim(), new Date(), true]);
  return { ok: true, id };
}

function adminListMessages(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const data = getMessagesSheet().getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    msgs.push({ id: d[0], target: d[1], message: d[2], createdAt: d[3], active: d[4] === true || d[4] === 'TRUE' });
  }
  msgs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { ok: true, messages: msgs };
}

function adminDeleteMessage(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getMessagesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) { sheet.getRange(i + 1, 5).setValue(false); return { ok: true }; }
  }
  return { ok: false, error: 'Message not found' };
}

/** Public, no-auth: returns the active flash message(s) for the login screen, split by
 * placement ('LOGIN_TOP' or 'LOGIN_BOTTOM'). Unlike getMessages, this is callable before
 * anyone signs in — same trust level as getMarketIndices (read-only, no investor data). */
function getLoginMessages() {
  const data = getMessagesSheet().getDataRange().getValues();
  const top = [], bottom = [], note = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    const active = d[4] === true || d[4] === 'TRUE';
    if (!active) continue;
    const target = (d[1] || '').toString();
    if (target === 'LOGIN_TOP') top.push({ id: d[0], message: d[2], createdAt: d[3] });
    else if (target === 'LOGIN_BOTTOM') bottom.push({ id: d[0], message: d[2], createdAt: d[3] });
    else if (target === 'LOGIN_MARKET_NOTE') note.push({ id: d[0], message: d[2], createdAt: d[3] });
  }
  top.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  bottom.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  note.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { ok: true, top: top, bottom: bottom, note: note.length ? note[0].message : '' };
}

/** Returns the active banner message(s) visible to a given investor: firm-wide "ALL" plus any addressed to them. */
function getMessages(email, token) {
  if (!authOk(email, token)) return { ok: false, error: 'Not authenticated' };
  const data = getMessagesSheet().getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    const active = d[4] === true || d[4] === 'TRUE';
    if (!active) continue;
    const target = (d[1] || 'ALL').toString();
    if (target === 'ALL' || target.toLowerCase() === (email || '').toLowerCase()) {
      msgs.push({ id: d[0], message: d[2], createdAt: d[3] });
    }
  }
  msgs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { ok: true, messages: msgs };
}

/** Investor submits a New Investment / Additional Purchase request, which lands with the advisor. */
function submitPurchaseRequest(body) {
  if (!authOk(body.email, body.token)) return { ok: false, error: 'Not authenticated' };
  const sheet = getRequestsSheet();
  const id = 'r_' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, body.email, body.name || '', body.type || 'New Investment', body.fundName || '', Number(body.amount || 0), body.notes || '', 'Pending', new Date()]);
  return { ok: true, id };
}

function adminListRequests(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const data = getRequestsSheet().getDataRange().getValues();
  const requests = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    requests.push({ id: d[0], email: d[1], name: d[2], type: d[3], fundName: d[4], amount: d[5], notes: d[6], status: d[7], createdAt: d[8] });
  }
  requests.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { ok: true, requests };
}

function adminUpdateRequestStatus(body) {
  const admin = requireAdmin(body.adminEmail, body.adminToken);
  if (!admin.ok) return admin;
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) { sheet.getRange(i + 1, 8).setValue(body.status || 'Pending'); return { ok: true }; }
  }
  return { ok: false, error: 'Request not found' };
}

/**
 * Looks up the latest AMFI NAV for a list of fund names (used by the investor Dashboard and AUM roll-up).
 */
function lookupNavs(names) {
  if (!names || names.length === 0) return { ok: true, navs: {} };
  if (fundsNeedRefresh()) {
    const r = refreshFundList();
    if (!r.ok) return r;
  }
  const sheet = getFundsSheet();
  const data = sheet.getDataRange().getValues();
  const byName = {};
  for (let i = 1; i < data.length; i++) {
    const n = (data[i][1] || '').toString().toLowerCase().trim();
    if (n) byName[n] = data[i][2];
  }
  const navs = {};
  names.forEach(name => {
    const key = (name || '').toLowerCase().trim();
    if (byName.hasOwnProperty(key)) { navs[name] = byName[key]; return; }
    const found = Object.keys(byName).find(k => k.indexOf(key) !== -1 || key.indexOf(k) !== -1);
    navs[name] = found ? byName[found] : null;
  });
  return { ok: true, navs };
}

const DOMESTIC_FEEDS = [
  { source: 'Economic Times', url: 'https://economictimes.indiatimes.com/markets/rssfeeds/1977021501.cms' },
  { source: 'Economic Times MF', url: 'https://economictimes.indiatimes.com/mf/rssfeeds/359241701.cms' },
  { source: 'Zee Business', url: 'https://zeenews.india.com/rss/business.xml' }
];
const INTL_FEEDS = [
  { source: 'Economic Times World', url: 'https://economictimes.indiatimes.com/news/international/rssfeeds/858478126.cms' }
];

// =====================================================================================
// Market index ticker — NIFTY 50, Bank NIFTY, Dow Jones, and a China index for the top-right
// of the home/login screen and the Dashboard. Uses Yahoo Finance's public, keyless chart
// endpoint (the same kind of free, no-signup source already used elsewhere here for AMFI/
// mfapi.in data). Cached 5 minutes so the ticker never adds real latency to a page open —
// worst case it's a few minutes behind, which is fine for a glance indicator.
// =====================================================================================
const MARKET_INDEX_SYMBOLS = [
  { key: 'NIFTY50', label: 'NIFTY 50', symbol: '%5ENSEI' },
  { key: 'BANKNIFTY', label: 'BANK NIFTY', symbol: '%5ENSEBANK' },
  { key: 'DOW', label: 'DOW', symbol: '%5EDJI' },
  { key: 'CHINA', label: 'CHINA (SSE)', symbol: '000001.SS' }
];

function getMarketIndices() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('fc_indices');
  if (cached) return { ok: true, indices: JSON.parse(cached), cached: true };

  const indices = MARKET_INDEX_SYMBOLS.map(def => {
    try {
      const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + def.symbol + '?range=1d&interval=1d';
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return { key: def.key, label: def.label, ok: false };
      const json = JSON.parse(res.getContentText());
      const meta = json.chart && json.chart.result && json.chart.result[0] && json.chart.result[0].meta;
      if (!meta || meta.regularMarketPrice === undefined) return { key: def.key, label: def.label, ok: false };
      const price = meta.regularMarketPrice;
      const prevClose = meta.previousClose || meta.chartPreviousClose || price;
      const change = price - prevClose;
      const changePct = prevClose ? (change / prevClose * 100) : 0;
      return { key: def.key, label: def.label, ok: true, price: price, change: change, changePct: changePct, currency: meta.currency || '' };
    } catch (e) {
      return { key: def.key, label: def.label, ok: false };
    }
  });
  cache.put('fc_indices', JSON.stringify(indices), 300);
  return { ok: true, indices: indices, cached: false };
}

function getNews() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('fc_news');
  if (cached) return { ok: true, ...JSON.parse(cached), cached: true };

  const domestic = fetchFeedItems(DOMESTIC_FEEDS).slice(0, 5);
  const international = fetchFeedItems(INTL_FEEDS).slice(0, 5);
  const result = { domestic, international };
  cache.put('fc_news', JSON.stringify(result), 900);
  return { ok: true, ...result, cached: false };
}

function fetchFeedItems(feeds) {
  const items = [];
  feeds.forEach(feed => {
    try {
      const res = UrlFetchApp.fetch(feed.url, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() !== 200) return;
      const doc = XmlService.parse(res.getContentText());
      const root = doc.getRootElement();
      const channel = root.getChild('channel');
      if (!channel) return;
      const rssItems = channel.getChildren('item');
      rssItems.slice(0, 8).forEach(it => {
        const title = (it.getChildText('title') || '').trim();
        const link = (it.getChildText('link') || '').trim();
        const pubDateRaw = it.getChildText('pubDate') || '';
        const pubDate = pubDateRaw ? new Date(pubDateRaw) : new Date(0);
        if (title) items.push({ title, link, source: feed.source, pubDate: pubDate.getTime() });
      });
    } catch (e) { /* skip broken feed */ }
  });
  items.sort((a, b) => b.pubDate - a.pubDate);
  return items;
}

/**
 * Sends an email to every investor whose next SIP due date is exactly 2 days away.
 */
function sendSipReminders() {
  const sheet = getClientsSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let sent = 0;
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    if (!email) continue;
    const sips = data[i][7] ? JSON.parse(data[i][7]) : [];
    sips.forEach(sip => {
      const due = nextDueDateServer(sip);
      if (!due) return;
      const diffDays = Math.round((due.getTime() - today.getTime()) / 86400000);
      if (diffDays === 2) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: 'SIP due in 2 days — ' + sip.fundName,
            body:
              'Hello,\n\nThis is a reminder that your ' + sip.frequency + ' SIP of ₹' + sip.amount +
              ' for ' + sip.fundName + ' is due on ' + due.toDateString() + '.\n\n' +
              'Please ensure sufficient balance for the auto-debit.\n\n' +
              '— EYM Financial · ARN-364244 · Contact us: 7292078936'
          });
          sent++;
        } catch (e) { /* keep going even if one email fails */ }
      }
    });
  }
  return { ok: true, sent };
}

function nextDueDateServer(sip) {
  const start = new Date(sip.startDate);
  if (isNaN(start)) return null;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let d = new Date(start);
  let guard = 0;
  while (d < today && guard < 3000) {
    advanceDate_(d, sip.frequency);
    guard++;
  }
  return d;
}

/**
 * Run this ONCE (select it in the function dropdown above and click Run) to schedule
 * sendSipReminders() to run automatically every day. Safe to re-run — it won't create duplicates.
 */
function setupDailyReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendSipReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendSipReminders').timeBased().everyDays(1).atHour(8).create();
  return { ok: true, message: 'Daily SIP reminder emails scheduled for ~8 AM every day.' };
}
