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
function getBalances(ss) {
  var sheet = ss.getSheetByName('Balances');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var accounts = { checking: [], savings: [], investment: [], credit: [] };
  var currentType = null;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var label = String(row[0] || '').trim().toUpperCase();

    if (label === 'CHECKING') {
      currentType = 'checking';
      continue;
    } else if (label === 'INVESTMENT') {
      currentType = 'investment';
      continue;
    } else if (label === 'SAVINGS') {
      currentType = 'savings';
      continue;
    } else if (label === 'CREDIT CARD') {
      currentType = 'credit';
      continue;
    }

    if (!currentType) continue;

    var name = String(row[0] || '').trim();
    if (!name || name === '' || label.indexOf('TILLER') > -1 || label.indexOf('YEAR') === 0) continue;

    var raw = row[2];
    var bal = typeof raw === 'number' ? raw : (parseFloat(String(raw).replace(/[$, ]/g, '')) || null);
    if (bal === null || isNaN(bal)) continue;

    var updated = String(row[1] || '').trim();
    accounts[currentType].push({ name: name, updated: updated, balance: bal });
  }

  return accounts;
}

// ─── Balance History ─────────────────────────────────────────────────────────
function getBalanceHistory(ss) {
  var sheet = ss.getSheetByName('Balances');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var result = { checking: null, investment: null, savings: null, liabilities: null };
  var histRows = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var count = 0;
    for (var j = 0; j < 12 && j < row.length; j++) {
      if (typeof row[j] === 'number' && row[j] >= 0) count++;
    }
    if (count >= 10 && String(row[0] || '').trim() === '') {
      histRows.push({ idx: i, row: row });
    }
  }

  if (histRows.length >= 4) {
    result.checking = histRows[0].row.slice(0, 12).map(function(v) { return typeof v === 'number' ? v : 0; });
    var invIdx = -1;
    for (var i = 1; i < histRows.length; i++) {
      if (histRows[i].idx > histRows[0].idx + 8) {
        invIdx = i;
        break;
      }
    }
    if (invIdx > 0) result.investment = histRows[invIdx].row.slice(0, 12).map(function(v) { return typeof v === 'number' ? v : 0; });
    var savIdx = -1;
    for (var i = invIdx + 1; i < histRows.length; i++) {
      savIdx = i;
      break;
    }
    if (savIdx > invIdx && savIdx >= 0) result.savings = histRows[savIdx].row.slice(0, 12).map(function(v) { return typeof v === 'number' ? v : 0; });
    var lastHist = histRows[histRows.length - 1];
    result.liabilities = lastHist.row.slice(0, 12).map(function(v) { return typeof v === 'number' ? v : 0; });
  }

  return result;
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
    retirementAdj: toPct(pick('retirement adjustment factor', 'retirement adjustment')),
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
