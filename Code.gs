// ===================================================
// pH Finance Dashboard — Google Apps Script Backend
// Paste this into a new Google Apps Script project,
// deploy as Web App (Anyone can access), and copy
// the deployment URL into the dashboard toolbar.
// ===================================================

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── Entry Points ──────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'load';
  if (action === 'load') {
    return handleLoad();
  }
  return jsonResponse({ error: 'Unknown GET action' });
}

function doPost(e) {
  var payload = {};
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON' });
  }
  if (payload.action === 'save') {
    return handleSave(payload.data);
  }
  return jsonResponse({ error: 'Unknown POST action' });
}

// ── Load ─────────────────────────────────────────────

function handleLoad() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ensureSheet(ss, 'RawData', '#37474F');
  var raw = sheet.getRange('B1').getValue();
  if (!raw) return jsonResponse({});
  try {
    return jsonResponse(JSON.parse(raw));
  } catch (e) {
    return jsonResponse({ error: 'Corrupt data in RawData!B1' });
  }
}

// ── Save ─────────────────────────────────────────────

function handleSave(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var rawSheet = ensureSheet(ss, 'RawData', '#37474F');
  rawSheet.getRange('B1').setValue(JSON.stringify(data));
  writeAllTables(ss, data);
  return jsonResponse({ success: true, timestamp: new Date().toISOString() });
}

// ── Write All Tables ──────────────────────────────────

function writeAllTables(ss, data) {
  writeSavingsSheet(ss, data.savings || []);
  writeInvestmentsSheet(ss, data.investments || []);
  writeBudgetSheet(ss, data.budget || {});
}

// ── Savings Sheet ─────────────────────────────────────

function writeSavingsSheet(ss, savings) {
  var sheet = ensureSheet(ss, 'Savings', '#1B5E20');
  sheet.clear();
  sheet.setFrozenRows(0);

  // Dark background
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 50), 6)
    .setBackground('#0F1117').setFontColor('#E8EAF0');

  // Title
  var titleCell = sheet.getRange('A1:F1');
  titleCell.merge().setValue('💰 Savings Jars')
    .setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
    .setFontColor('#C6FF00').setBackground('#141720').setHorizontalAlignment('center');

  // Timestamp
  sheet.getRange('A2:F2').merge()
    .setValue('Last synced: ' + new Date().toLocaleString('en-IN'))
    .setFontSize(9).setFontColor('#555A6A').setBackground('#0F1117');

  // Header row (row 3)
  var headers = ['#', 'Jar Name', 'Saved (₹)', 'Goal (₹)', 'Progress %', 'Remaining (₹)'];
  var headerRow = sheet.getRange(3, 1, 1, 6);
  headerRow.setValues([headers])
    .setBackground('#1A1E28').setFontColor('#8B90A0')
    .setFontWeight('bold').setFontSize(10);

  sheet.setFrozenRows(3);

  if (!savings.length) {
    sheet.getRange(4, 1).setValue('No savings jars');
    return;
  }

  // Data rows
  var rows = savings.map(function(j, i) {
    var pct = j.goal ? (j.amount / j.goal) : 0;
    var remaining = Math.max(0, j.goal - j.amount);
    return [i + 1, j.name, j.amount, j.goal, pct, remaining];
  });

  var dataRange = sheet.getRange(4, 1, rows.length, 6);
  dataRange.setValues(rows);
  dataRange.setFontFamily('Courier New').setFontSize(10);

  // Number formats
  sheet.getRange(4, 3, rows.length, 1).setNumberFormat('#,##0');
  sheet.getRange(4, 4, rows.length, 1).setNumberFormat('#,##0');
  sheet.getRange(4, 5, rows.length, 1).setNumberFormat('0%');
  sheet.getRange(4, 6, rows.length, 1).setNumberFormat('#,##0');

  // Zebra stripe + progress colour-coding
  rows.forEach(function(r, i) {
    var row = 4 + i;
    var bg = i % 2 === 0 ? '#141720' : '#0F1117';
    sheet.getRange(row, 1, 1, 6).setBackground(bg).setFontColor('#E8EAF0');
    // Progress cell colour
    var pct = r[4];
    var pctColor = pct < 0.33 ? '#FF3B30' : pct < 0.66 ? '#FFC94A' : '#C6FF00';
    sheet.getRange(row, 5).setFontColor(pctColor).setFontWeight('bold');
    sheet.getRange(row, 1).setFontColor('#555A6A'); // # column dim
  });

  // Totals row
  var totalRow = 4 + rows.length;
  var totalSaved = savings.reduce(function(a, j) { return a + j.amount; }, 0);
  var totalGoal = savings.reduce(function(a, j) { return a + j.goal; }, 0);
  sheet.getRange(totalRow, 1, 1, 6)
    .setValues([['', 'TOTAL', totalSaved, totalGoal, totalGoal ? totalSaved / totalGoal : 0, Math.max(0, totalGoal - totalSaved)]])
    .setBackground('#1A1E28').setFontColor('#C6FF00').setFontWeight('bold');
  sheet.getRange(totalRow, 3).setNumberFormat('#,##0');
  sheet.getRange(totalRow, 4).setNumberFormat('#,##0');
  sheet.getRange(totalRow, 5).setNumberFormat('0%');
  sheet.getRange(totalRow, 6).setNumberFormat('#,##0');

  // Column widths
  sheet.setColumnWidths(1, 6, 120);
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(5, 100);
}

// ── Investments Sheet ─────────────────────────────────

function writeInvestmentsSheet(ss, investments) {
  var sheet = ensureSheet(ss, 'Investments', '#004D40');
  sheet.clear();
  sheet.setFrozenRows(0);

  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 50), 7)
    .setBackground('#0F1117').setFontColor('#E8EAF0');

  sheet.getRange('A1:G1').merge().setValue('📈 Investment Jars')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor('#00E5C8').setBackground('#141720').setHorizontalAlignment('center');

  sheet.getRange('A2:G2').merge()
    .setValue('Last synced: ' + new Date().toLocaleString('en-IN'))
    .setFontSize(9).setFontColor('#555A6A').setBackground('#0F1117');

  var headers = ['#', 'Jar Name', 'Invested (₹)', 'Returns (₹)', 'Total Value (₹)', 'Target (₹)', 'Progress %'];
  sheet.getRange(3, 1, 1, 7)
    .setValues([headers])
    .setBackground('#1A1E28').setFontColor('#8B90A0')
    .setFontWeight('bold').setFontSize(10);

  sheet.setFrozenRows(3);

  if (!investments.length) {
    sheet.getRange(4, 1).setValue('No investment jars');
    return;
  }

  var rows = investments.map(function(j, i) {
    var total = (j.amount || 0) + (j.returns || 0);
    var pct = j.target ? total / j.target : 0;
    return [i + 1, j.name, j.amount || 0, j.returns || 0, total, j.target || 0, pct];
  });

  var dataRange = sheet.getRange(4, 1, rows.length, 7);
  dataRange.setValues(rows).setFontFamily('Courier New').setFontSize(10);

  sheet.getRange(4, 3, rows.length, 4).setNumberFormat('#,##0');
  sheet.getRange(4, 7, rows.length, 1).setNumberFormat('0%');

  rows.forEach(function(r, i) {
    var row = 4 + i;
    var bg = i % 2 === 0 ? '#141720' : '#0F1117';
    sheet.getRange(row, 1, 1, 7).setBackground(bg).setFontColor('#E8EAF0');
    var pct = r[6];
    var pctColor = pct < 0.33 ? '#FF3B30' : pct < 0.66 ? '#FFC94A' : '#00E5C8';
    sheet.getRange(row, 7).setFontColor(pctColor).setFontWeight('bold');
    sheet.getRange(row, 1).setFontColor('#555A6A');
  });

  var totalRow = 4 + rows.length;
  var ti = investments.reduce(function(a, j) { return a + (j.amount || 0); }, 0);
  var tr = investments.reduce(function(a, j) { return a + (j.returns || 0); }, 0);
  var tt = investments.reduce(function(a, j) { return a + (j.target || 0); }, 0);
  sheet.getRange(totalRow, 1, 1, 7)
    .setValues([['', 'TOTAL', ti, tr, ti + tr, tt, tt ? (ti + tr) / tt : 0]])
    .setBackground('#1A1E28').setFontColor('#00E5C8').setFontWeight('bold');
  sheet.getRange(totalRow, 3, 1, 4).setNumberFormat('#,##0');
  sheet.getRange(totalRow, 7).setNumberFormat('0%');

  sheet.setColumnWidths(1, 7, 120);
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 180);
}

// ── Budget Sheet ──────────────────────────────────────

function writeBudgetSheet(ss, budget) {
  var sheet = ensureSheet(ss, 'Budget', '#B71C1C');
  sheet.clear();
  sheet.setFrozenRows(0);

  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 100), 6)
    .setBackground('#0F1117').setFontColor('#E8EAF0');

  sheet.getRange('A1:F1').merge().setValue('📊 Monthly Budget & Expenses')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor('#FF6B5B').setBackground('#141720').setHorizontalAlignment('center');

  sheet.getRange('A2:F2').merge()
    .setValue('Last synced: ' + new Date().toLocaleString('en-IN'))
    .setFontSize(9).setFontColor('#555A6A').setBackground('#0F1117');

  var colHeaders = ['Year', 'Month', 'Budget (₹)', 'Expenses (₹)', 'Surplus/Deficit (₹)', 'Status'];
  sheet.getRange(3, 1, 1, 6)
    .setValues([colHeaders])
    .setBackground('#1A1E28').setFontColor('#8B90A0')
    .setFontWeight('bold').setFontSize(10);

  sheet.setFrozenRows(3);

  var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  // Sort entries chronologically
  var entries = Object.keys(budget).map(function(k) {
    var parts = k.split('-');
    return { year: parseInt(parts[0]), month: parseInt(parts[1]), data: budget[k] };
  }).sort(function(a, b) {
    return a.year !== b.year ? a.year - b.year : a.month - b.month;
  });

  if (!entries.length) {
    sheet.getRange(4, 1).setValue('No budget data');
    return;
  }

  // Group by year
  var byYear = {};
  entries.forEach(function(e) {
    if (!byYear[e.year]) byYear[e.year] = [];
    byYear[e.year].push(e);
  });

  var currentRow = 4;
  var years = Object.keys(byYear).map(Number).sort();

  years.forEach(function(year) {
    var yearEntries = byYear[year];

    // Year header row
    sheet.getRange(currentRow, 1, 1, 6).merge()
      .setValue('── ' + year + ' ──')
      .setBackground('#1A2530').setFontColor('#FFC94A')
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    currentRow++;

    // Month rows
    var yearBudget = 0, yearExpense = 0;
    yearEntries.forEach(function(e, i) {
      var diff = e.data.budget - e.data.expense;
      var status = diff >= 0 ? '✅ On track' : '🔴 Over budget';
      yearBudget += e.data.budget;
      yearExpense += e.data.expense;

      var row = [year, MONTHS[e.month - 1], e.data.budget, e.data.expense, diff, status];
      var dataRange = sheet.getRange(currentRow, 1, 1, 6);
      dataRange.setValues([row]);
      var bg = i % 2 === 0 ? '#141720' : '#0F1117';
      dataRange.setBackground(bg).setFontColor('#E8EAF0').setFontSize(10).setFontFamily('Courier New');
      // Surplus/deficit colour
      var diffColor = diff >= 0 ? '#C6FF00' : '#FF6B5B';
      sheet.getRange(currentRow, 5).setFontColor(diffColor).setFontWeight('bold');
      currentRow++;
    });

    // Year subtotal
    var yearDiff = yearBudget - yearExpense;
    sheet.getRange(currentRow, 1, 1, 6)
      .setValues([['', 'Year Total', yearBudget, yearExpense, yearDiff, yearDiff >= 0 ? '✅ Net Savings' : '🔴 Overspent']])
      .setBackground('#1A1E28').setFontColor('#FFC94A').setFontWeight('bold').setFontSize(10);
    sheet.getRange(currentRow, 5).setFontColor(yearDiff >= 0 ? '#C6FF00' : '#FF6B5B');
    currentRow++;
    currentRow++; // spacer
  });

  // Number formats — apply to data area
  var dataRows = currentRow - 4;
  if (dataRows > 0) {
    sheet.getRange(4, 3, dataRows, 3).setNumberFormat('#,##0');
  }

  sheet.setColumnWidths(1, 6, 130);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(6, 140);
}

// ── Helper: Ensure Sheet ──────────────────────────────

function ensureSheet(ss, name, tabColor) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  try { sheet.setTabColor(tabColor); } catch(e) {}
  return sheet;
}

// ── Helper: JSON Response ─────────────────────────────

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
