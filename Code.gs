/**
 * Folio & Compass — Google Sheets backend.
 * Paste this into Extensions > Apps Script in your Google Sheet, then deploy as a Web App.
 * See SHEETS_SETUP.md for step-by-step instructions.
 *
 * Expects two tabs in the spreadsheet:
 *   "Users"   columns: Email | PasswordHash | IsAdvisor
 *   "Clients" columns: Email | Name | Phone | PAN | DOB | RiskScore | RiskDate | SIPsJSON | UpdatedAt
 * A third tab, "Funds", is created automatically to cache the AMFI scheme list.
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

/**
 * Fetches AMFI's full daily NAV list (https://www.amfiindia.com/spages/NAVAll.txt),
 * parses every scheme row, and rewrites the "Funds" tab with the latest data.
 * Browsers can't fetch this file directly (AMFI doesn't send CORS headers), so this
 * server-side fetch is what lets the app offer a searchable fund list.
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
    if (!line || line.indexOf(';') === -1) continue; // skip blank lines & category headers
    const parts = line.split(';');
    if (parts.length < 5) continue;
    const code = parts[0].trim();
    const name = parts[3].trim();
    const navRaw = parts[4].trim();
    if (!code || !name || isNaN(Number(code))) continue; // real scheme rows start with a numeric code
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

function signup(email, password) {
  if (!email || !password) return { ok: false, error: 'Email and password required' };
  if (password.length < 6) return { ok: false, error: 'Password must be at least 6 characters' };
  const users = getUsersSheet();
  if (findRow(users, email)) return { ok: false, error: 'An account with that email already exists' };
  users.appendRow([email, hashPw(password), false]);
  getClientsSheet().appendRow([email, '', '', '', '', '', '', '[]', new Date()]);
  return { ok: true, isAdvisor: false };
}

function login(email, password) {
  const u = findRow(getUsersSheet(), email);
  if (!u) return { ok: false, error: 'No account found for that email' };
  if (u.data[1] !== hashPw(password)) return { ok: false, error: 'Incorrect password' };
  const isAdvisor = (u.data[2] === true || u.data[2] === 'TRUE');
  return { ok: true, isAdvisor };
}

function authOk(email, password) {
  const u = findRow(getUsersSheet(), email);
  return !!u && u.data[1] === hashPw(password);
}

function saveData(body) {
  if (!authOk(body.email, body.password)) return { ok: false, error: 'Not authenticated' };
  const sheet = getClientsSheet();
  const inv = body.investor || {};
  const risk = body.riskAssessment || {};
  const row = [
    body.email, inv.name || '', inv.phone || '', inv.pan || '', inv.dob || '',
    risk.score || '', risk.date || '', JSON.stringify(body.sips || []), new Date()
  ];
  const existing = findRow(sheet, body.email);
  if (existing) sheet.getRange(existing.row, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { ok: true };
}

function loadData(email, password) {
  if (!authOk(email, password)) return { ok: false, error: 'Not authenticated' };
  const c = findRow(getClientsSheet(), email);
  if (!c) return { ok: true, data: null };
  const d = c.data;
  return {
    ok: true,
    data: {
      investor: { name: d[1], email: email, phone: d[2], pan: d[3], dob: d[4] },
      riskAssessment: d[5] ? { score: Number(d[5]), date: d[6] } : null,
      sips: d[7] ? JSON.parse(d[7]) : []
    }
  };
}

function listClients(email, password) {
  const u = findRow(getUsersSheet(), email);
  if (!u || u.data[1] !== hashPw(password)) return { ok: false, error: 'Not authenticated' };
  if (!(u.data[2] === true || u.data[2] === 'TRUE')) return { ok: false, error: 'This account is not marked as an advisor' };
  const data = getClientsSheet().getDataRange().getValues();
  const clients = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    clients.push({
      email: d[0],
      name: d[1],
      riskScore: d[5] || null,
      sipsCount: d[7] ? JSON.parse(d[7]).length : 0,
      sips: d[7] ? JSON.parse(d[7]) : []
    });
  }
  return { ok: true, clients };
}
