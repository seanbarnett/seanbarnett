// Tiller Financial Dashboard - Google Apps Script API
//
// DEPLOY:
//   Extensions -> Apps Script -> paste this code -> Deploy -> New deployment
//   Type: Web App | Execute as: Me | Who has access: Anyone
//   Copy the /exec URL into dashboard.html -> CONFIG.GAS_URL
//
// SHEET ID: 1g0kUMLJWsS5pGwDhMOFHfyoeXldx_2htl2ROvAFpzFc

var SHEET_ID = '1g0kUMLJWsS5pGwDhMOFHfyoeXldx_2htl2ROvAFpzFc';

function doGet() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var payload = {
      balances: getBalances(ss),
      history: getBalanceHistory(ss),
      assumptions: getAssumptions(ss),
      budget: getBudget(ss),
      targets: getTargets(ss),
      lastUpdated: new Date().toISOString(),
    };
    return json(payload);
  } catch (e) {
    return json({ error: e.message });
  }
}

// ─── Balances ────────────────────────────────────────────────────────────────
// The "Balances" sheet is a formatted dashboard (gauges, colored banners),
// not a plain data grid, so its visual columns shift around and can't be
// parsed positionally. Instead we read its "Sorted Assets" / "Sorted
// Liabilities" reference tables, which list every account in a fixed
// Row | Id | Group | Account | Last Updated | Balance layout. We locate
// each table by its header cells rather than hardcoding column numbers,
// since the dashboard's charts/formatting can shift columns over time.
function getBalances(ss) {
  var sheet = ss.getSheetByName('Balances');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var accounts = { checking: [], savings: [], investment: [], credit: [] };
  var tableStarts = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var c = 0; c + 5 < row.length; c++) {
      if (String(row[c]).trim() === 'Row' &&
          String(row[c + 1]).trim() === 'Id' &&
          String(row[c + 2]).trim() === 'Group' &&
          String(row[c + 3]).trim() === 'Account') {
        tableStarts.push({ row: i, col: c });
      }
    }
  }

  tableStarts.forEach(function(start) {
    for (var r = start.row + 1; r < data.length; r++) {
      var row = data[r];
      var group   = String(row[start.col + 2] || '').trim().toLowerCase();
      var name    = String(row[start.col + 3] || '').trim();
      var balance = row[start.col + 5];

      if (!group && !name) break; // blank row = end of this table

      if (!name || typeof balance !== 'number') continue;

      var entry = { name: name, updated: formatUpdated(row[start.col + 4]), balance: balance };
      if (group === 'checking') accounts.checking.push(entry);
      else if (group === 'investment') accounts.investment.push(entry);
      else if (group === 'savings') accounts.savings.push(entry);
      else if (group === 'credit card') accounts.credit.push(entry);
    }
  });

  return accounts;
}

function formatUpdated(daysAgo) {
  if (typeof daysAgo !== 'number') return String(daysAgo || '');
  if (daysAgo === 0) return 'today';
  if (daysAgo === 1) return '1 day ago';
  return daysAgo + ' days ago';
}

// ─── Balance History ─────────────────────────────────────────────────────────
// Built from Tiller's "Balance History" sheet — one row per account per
// snapshot. For each of the last 12 months we take each account's most
// recent snapshot in that month, then sum by the account's Group.
//
// Grouping comes from the Accounts sheet, not Balance History's own "Type"
// column: Type holds raw institution types (ROTH_IRA, CREDIT,
// HEALTH_SAVINGS_ACCOUNT_HSA), whereas Group holds the Checking / Savings /
// Investment / Credit Card buckets this dashboard reports on. An HSA, for
// example, may sit under Investment or Savings depending on how it's been
// categorized in Accounts.
var MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function getBalanceHistory(ss) {
  var sheet = ss.getSheetByName('Balance History');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var header = data[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  var iDate    = header.indexOf('date');
  var iAcctId  = header.indexOf('account id');
  var iAccount = header.indexOf('account');
  var iBalance = header.indexOf('balance');
  var iClass   = header.indexOf('class');
  if (iDate === -1 || iBalance === -1) return null;

  var groupById = getAccountGroups(ss);

  // The 12 months ending with the current one.
  var now = new Date();
  var monthKeys = {};
  var labels = [];
  for (var m = 11; m >= 0; m--) {
    var d = new Date(now.getFullYear(), now.getMonth() - m, 1);
    monthKeys[d.getFullYear() + '-' + d.getMonth()] = 11 - m;
    labels.push(MONTH_ABBR[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2));
  }

  // Most recent snapshot per account per month.
  var latest = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dt = row[iDate];
    if (!(dt instanceof Date)) continue;

    var mi = monthKeys[dt.getFullYear() + '-' + dt.getMonth()];
    if (mi === undefined) continue;

    var bal = row[iBalance];
    if (typeof bal !== 'number') continue;

    var acctId = String((iAcctId !== -1 ? row[iAcctId] : row[iAccount]) || '').trim();
    if (!acctId) continue;

    var key = acctId + '|' + mi;
    var existing = latest[key];
    if (!existing || dt.getTime() > existing.time) {
      latest[key] = {
        time: dt.getTime(),
        balance: bal,
        acctId: acctId,
        monthIdx: mi,
        cls: String((iClass !== -1 ? row[iClass] : '') || '').trim().toLowerCase(),
      };
    }
  }

  var zeros = function() { return labels.map(function() { return 0; }); };
  var result = {
    labels: labels,
    checking: zeros(),
    investment: zeros(),
    savings: zeros(),
    liabilities: zeros(),
  };

  for (var k in latest) {
    var e = latest[k];
    var group = (groupById[e.acctId] || '').toLowerCase();

    var bucket = null;
    if (group === 'checking') bucket = 'checking';
    else if (group === 'investment') bucket = 'investment';
    else if (group === 'savings') bucket = 'savings';
    else if (group === 'credit card' || e.cls === 'liability') bucket = 'liabilities';
    if (!bucket) continue;

    // Liabilities are reported as positive magnitudes; the dashboard
    // subtracts them when computing net worth.
    result[bucket][e.monthIdx] += bucket === 'liabilities' ? Math.abs(e.balance) : e.balance;
  }

  ['checking', 'investment', 'savings', 'liabilities'].forEach(function(b) {
    result[b] = result[b].map(round2);
  });

  return result;
}

// Maps Account Id -> Group using the Accounts sheet, locating both columns
// by header name so it survives column reordering.
function getAccountGroups(ss) {
  var sheet = ss.getSheetByName('Accounts');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();

  var iId = -1, iGroup = -1, headerRow = -1;
  for (var r = 0; r < data.length && headerRow === -1; r++) {
    var row = data[r].map(function(v) { return String(v || '').trim().toLowerCase(); });
    var idIdx = row.indexOf('account id');
    var groupIdx = row.indexOf('group');
    if (idIdx !== -1 && groupIdx !== -1) {
      iId = idIdx;
      iGroup = groupIdx;
      headerRow = r;
    }
  }
  if (headerRow === -1) return {};

  var map = {};
  for (var i = headerRow + 1; i < data.length; i++) {
    var id = String(data[i][iId] || '').trim();
    var group = String(data[i][iGroup] || '').trim();
    if (id && group) map[id] = group;
  }
  return map;
}

// ─── Assumptions ─────────────────────────────────────────────────────────────
function getAssumptions(ss) {
  var sheet = ss.getSheetByName('ASSUMPTIONS');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var map = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var label = String(row[0] || '').trim();
    var val = row[1];
    if (label && val !== '' && val !== null) {
      map[label.toLowerCase()] = val;
    }
  }

  function pick() {
    for (var i = 0; i < arguments.length; i++) {
      var k = arguments[i].toLowerCase();
      if (map[k] !== undefined) return map[k];
    }
    return null;
  }

  function toPct(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return v <= 1 ? v : v / 100;
    var n = parseFloat(String(v).replace('%', ''));
    return isNaN(n) ? null : (n > 1 ? n / 100 : n);
  }

  return {
    nominalReturn: toPct(pick('expected nominal return')),
    inflation: toPct(pick('expected annual inflation', 'expected inflation')),
    realReturn: toPct(pick('real return (after inflation)')),
    swr: toPct(pick('safe withdrawal rate (real)', 'safe withdrawal rate')),
    annualSpending: pick('annual spending (3-year avg)', 'annual spending', 'normalized baseline spending'),
    retirementAdj: toPct(pick('retirement spending factor', 'retirement adjustment factor', 'retirement adjustment')),
    geoAdj: toPct(pick('selected geographic scenario')),
    inflationAnchor: pick('inflation anchor year'),
    totalMonthly: pick('total monthly investment'),
    bonusAmt: pick('bonus amount invested'),
    annualTotal: pick('total annual investment'),
  };
}

// ─── Budget ───────────────────────────────────────────────────────────────────
function getBudget(ss) {
  var sheet = ss.getSheetByName('Monthly Budget');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var TYPES = { 'income': 1, 'expense': 1 };
  var GROUPS = { 'income': 1, 'fixed bills': 1, 'guilt free spend': 1, 'save & invest': 1 };

  var actuals = {};
  var budgets = {};

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    for (var c = 0; c + 3 < row.length; c++) {
      var typeVal = String(row[c] || '').trim().toLowerCase();
      var groupVal = String(row[c + 1] || '').trim().toLowerCase();
      var catVal = String(row[c + 2] || '').trim();
      var amtVal = row[c + 3];

      if (TYPES[typeVal] && GROUPS[groupVal] && catVal && typeof amtVal === 'number') {
        var key = typeVal + '|' + groupVal + '|' + catVal;
        actuals[key] = (actuals[key] || 0) + Math.abs(amtVal);
      }
    }
  }

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    for (var c = 0; c + 3 < row.length; c++) {
      var typeVal = String(row[c] || '').trim().toLowerCase();
      var groupVal = String(row[c + 1] || '').trim().toLowerCase();
      var catVal = String(row[c + 2] || '').trim();
      var amtVal = row[c + 3];

      if (TYPES[typeVal] && GROUPS[groupVal] && catVal && typeof amtVal === 'number' && amtVal > 0) {
        var key = typeVal + '|' + groupVal + '|' + catVal;
        if (!budgets[key]) budgets[key] = amtVal;
      }
    }
  }

  var grouped = {};
  for (var key in actuals) {
    var parts = key.split('|');
    var type = parts[0];
    var group = parts[1];
    var cat = parts[2];
    if (!grouped[group]) {
      grouped[group] = { name: titleCase(group), type: type, categories: [] };
    }
    grouped[group].categories.push({
      name: cat,
      actual: round2(actuals[key]),
      budget: round2(budgets[key] || 0),
    });
  }

  for (var g in grouped) {
    grouped[g].categories.sort(function(a, b) { return b.actual - a.actual; });
    grouped[g].actual = round2(grouped[g].categories.reduce(function(s, c) { return s + c.actual; }, 0));
    grouped[g].budget = round2(grouped[g].categories.reduce(function(s, c) { return s + c.budget; }, 0));
  }

  var groups = [];
  for (var g in grouped) {
    groups.push(grouped[g]);
  }

  return {
    period: findCellValue(data, 'Budget Setup/Budget Period') || '—',
    groups: groups,
  };
}

// ─── Targets ─────────────────────────────────────────────────────────────────
function getTargets(ss) {
  var sheet = ss.getSheetByName('TARGETS');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var label = String(row[0] || '').trim();
    if (label) map[label] = row[1];
  }
  return map;
}

// ─── Utilities ───────────────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function titleCase(s) {
  return s.replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function findCellValue(data, label) {
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var c = 0; c < row.length - 1; c++) {
      if (String(row[c] || '').trim() === label) {
        return row[c + 1];
      }
    }
  }
  return null;
}
