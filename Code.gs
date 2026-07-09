/**
 * Folio & Compass — Google Sheets backend.
 * Paste this into Extensions > Apps Script in your Google Sheet, then deploy as a Web App.
 * See SHEETS_SETUP.md for step-by-step instructions.
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
 * post banner messages, view firm-wide AUM, and action New Purchase Requests.
 */

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'signup') result = signup(body.email, body.password);
    else if (action === 'login') result = login(body.email, body.password);
    else if (action === 'saveData') result = saveData(body);
    else if (action === 'loadData') result = loadData(body.email, body.password);
    else if (action === 'listClients') result = listClients(body.email, body.password);
    else if (action === 'searchFunds') result = searchFunds(body.query);
    else if (action === 'refreshFunds') result = refreshFundList();
    else if (action === 'lookupNavs') result = lookupNavs(body.names || []);
    else if (action === 'getNews') result = getNews();
    else if (action === 'setupReminderTrigger') result = setupDailyReminderTrigger();
    else if (action === 'getMessages') result = getMessages(body.email);
    else if (action === 'submitPurchaseRequest') result = submitPurchaseRequest(body);
    // ---- super-user / admin actions ----
    else if (action === 'adminAddSip') result = adminAddSip(body);
    else if (action === 'adminUploadCas') result = adminUploadCas(body);
    else if (action === 'adminChangePassword') result = adminChangePassword(body);
    else if (action === 'adminSetBlocked') result = adminSetBlocked(body);
    else if (action === 'adminBroadcast') result = adminBroadcast(body);
    else if (action === 'adminDeleteMessage') result = adminDeleteMessage(body);
    else if (action === 'adminListMessages') result = adminListMessages(body);
    else if (action === 'adminListRequests') result = adminListRequests(body);
    else if (action === 'adminUpdateRequestStatus') result = adminUpdateRequestStatus(body);
    else result = { ok: false, error: 'Unknown action' };
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

function fundsNeedRefresh() {
  const sheet = getFundsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return true;
  const lastRefreshed = sheet.getRange(2, 4).getValue();
  if (!lastRefreshed) return true;
  const ageHours = (Date.now() - new Date(lastRefreshed).getTime()) / 36e5;
  return ageHours > 24;
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
      matches.push({ name: name, nav: data[i][2] || '' });
    }
  }
  return { ok: true, funds: matches };
}

function hashPw(pw) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  return digest.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

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
  if (password.length < 6) return { ok: false, error: 'Password must be at least 6 characters' };
  const users = getUsersSheet();
  if (findRow(users, email)) return { ok: false, error: 'An account with that email already exists' };
  users.appendRow([email, hashPw(password), false, false]);
  getClientsSheet().appendRow([email, '', '', '', '', '', '', '[]', new Date(), '[]']);
  return { ok: true, isAdvisor: false };
}

function login(email, password) {
  const u = findRow(getUsersSheet(), email);
  if (!u) return { ok: false, error: 'No account found for that email' };
  if (u.data[1] !== hashPw(password)) return { ok: false, error: 'Incorrect password' };
  if (isBlocked(u.data)) return { ok: false, error: 'This account has been blocked. Please contact your advisor.' };
  const isAdvisor = (u.data[2] === true || u.data[2] === 'TRUE');
  return { ok: true, isAdvisor };
}

function authOk(email, password) {
  const u = findRow(getUsersSheet(), email);
  return !!u && u.data[1] === hashPw(password) && !isBlocked(u.data);
}

// Verifies the caller is a non-blocked account with IsAdvisor = TRUE.
function requireAdmin(email, password) {
  const u = findRow(getUsersSheet(), email);
  if (!u || u.data[1] !== hashPw(password)) return { ok: false, error: 'Not authenticated' };
  if (isBlocked(u.data)) return { ok: false, error: 'This account has been blocked.' };
  if (!(u.data[2] === true || u.data[2] === 'TRUE')) return { ok: false, error: 'Super-user rights required for this action' };
  return { ok: true };
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
  if (!authOk(body.email, body.password)) return { ok: false, error: 'Not authenticated' };
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

function loadData(email, password) {
  if (!authOk(email, password)) return { ok: false, error: 'Not authenticated' };
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
function listClients(email, password) {
  const admin = requireAdmin(email, password);
  if (!admin.ok) return admin;
  const users = getUsersSheet().getDataRange().getValues();
  const blockedByEmail = {};
  for (let i = 1; i < users.length; i++) {
    blockedByEmail[(users[i][0] || '').toString().toLowerCase()] = isBlocked(users[i]);
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
      riskScore: d[5] || null,
      sips: d[7] ? JSON.parse(d[7]) : [],
      holdings: d[9] ? JSON.parse(d[9]) : [],
      blocked: !!blockedByEmail[key]
    });
  }
  return { ok: true, clients };
}

/** Super-user: add a SIP to any investor's register. */
function adminAddSip(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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

/**
 * Super-user: upload/append CAS (Consolidated Account Statement) holdings on behalf of any investor.
 * Expects body.holdings = [{ fundName, folio, schemeType, units, nav, investedValue }]
 * mode: "replace" (default) overwrites the investor's holdings list, "append" adds to it.
 */
function adminUploadCas(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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

/** Super-user: change any investor's password. */
function adminChangePassword(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
  if (!admin.ok) return admin;
  if (!body.newPassword || body.newPassword.length < 6) return { ok: false, error: 'New password must be at least 6 characters' };
  const users = getUsersSheet();
  const u = findRow(users, body.targetEmail);
  if (!u) return { ok: false, error: 'No such user account' };
  users.getRange(u.row, 2).setValue(hashPw(body.newPassword));
  return { ok: true };
}

/** Super-user: block or unblock any investor's ability to log in. */
function adminSetBlocked(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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

/** Super-user: post a banner message to one investor's email, or 'ALL' for everyone. */
function adminBroadcast(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
  if (!admin.ok) return admin;
  if (!body.message || !body.message.trim()) return { ok: false, error: 'Message text is required' };
  const sheet = getMessagesSheet();
  const id = 'm_' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, body.target || 'ALL', body.message.trim(), new Date(), true]);
  return { ok: true, id };
}

function adminListMessages(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
  if (!admin.ok) return admin;
  const sheet = getMessagesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) { sheet.getRange(i + 1, 5).setValue(false); return { ok: true }; }
  }
  return { ok: false, error: 'Message not found' };
}

/** Returns the active banner message(s) visible to a given investor: firm-wide "ALL" plus any addressed to them. */
function getMessages(email) {
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
  if (!authOk(body.email, body.password)) return { ok: false, error: 'Not authenticated' };
  const sheet = getRequestsSheet();
  const id = 'r_' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([id, body.email, body.name || '', body.type || 'New Investment', body.fundName || '', Number(body.amount || 0), body.notes || '', 'Pending', new Date()]);
  return { ok: true, id };
}

function adminListRequests(body) {
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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
  const admin = requireAdmin(body.adminEmail, body.adminPassword);
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
    if (sip.frequency === 'Daily') d.setDate(d.getDate() + 1);
    else if (sip.frequency === 'Weekly') d.setDate(d.getDate() + 7);
    else if (sip.frequency === 'Monthly') d.setMonth(d.getMonth() + 1);
    else if (sip.frequency === 'Quarterly') d.setMonth(d.getMonth() + 3);
    else d.setMonth(d.getMonth() + 1);
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
