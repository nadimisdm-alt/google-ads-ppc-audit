/**
 * ============================================================
 * GOOGLE ADS PPC AUDIT DASHBOARD SCRIPT
 * ============================================================
 * Pulls key performance insights straight into Google Sheets.
 * Spreadsheet is auto-renamed to: "[Account Name] — Google Ads PPC Audit — YYYY-MM-DD"
 *
 * TABS GENERATED:
 *  0. 🏠 Overview               — Account summary + health scorecard
 *  1. 💀 Zombie Campaigns       — Spending but 0 conversions (14 days)
 *  2. ❌ Non-Converting Terms    — Search terms burning budget (14 days)
 *  3. 💰 High CPC Terms         — CPC above threshold + no conversions
 *  4. 🛑 Negative KW Candidates — Terms to exclude
 *  5. 📈 Week-over-Week         — Last 7d vs prior 7d performance
 *
 * HOW TO USE:
 *  1. Open Google Ads → Tools & Settings → Scripts
 *  2. Create a new script and paste this entire file
 *  3. Edit CONFIG below (especially SPREADSHEET_URL)
 *  4. Authorize and Run (or schedule weekly)
 *
 * FIXES (v2):
 *  - CTR was being multiplied by 100 incorrectly in multiple places.
 *    The Google Ads API returns Ctr as a decimal fraction (e.g. 0.0017
 *    for 0.17%). Multiplying by 100 once gives the correct percentage.
 *    The original script multiplied in some paths more than once,
 *    causing values like 169 instead of 0.17%.
 * ============================================================
 */

// ============================================================
// CONFIG — Edit these values before running
// ============================================================
var CONFIG = {
  SPREADSHEET_URL: 'SPREADSHEET_URL',
  HIGH_CPC_THRESHOLD: 3.00,
  MIN_IMPRESSIONS_FOR_AUDIT: 10,
  MIN_ZOMBIE_SPEND: 1.00,
};
// ============================================================


function main() {
  var ss;
  try {
    ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  } catch (e) {
    throw new Error(
      'Cannot open spreadsheet. Check SPREADSHEET_URL and share it with the Ads account email.\nError: ' + e.message
    );
  }

  Logger.log('=== PPC AUDIT SCRIPT START ===');

  var accountName = AdsApp.currentAccount().getName();
  var tz = AdsApp.currentAccount().getTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // ── Auto-rename spreadsheet ──
  ss.rename(accountName + ' — Google Ads PPC Audit — ' + today);
  Logger.log('Spreadsheet renamed: ' + accountName + ' — Google Ads PPC Audit — ' + today);

  // ── Run audit tabs first so Overview can count their rows ──
  runZombieCampaigns(ss);
  runNonConvertingSearchTerms(ss);
  runHighCpcSearchTerms(ss);
  runNegativeKeywordCandidates(ss);
  runWeekOverWeek(ss);

  // ── Overview last (reads row counts from other tabs) ──
  runOverview(ss, accountName, tz);

  // ── Move Overview to first position ──
  var overviewSheet = ss.getSheetByName('🏠 Overview');
  if (overviewSheet) {
    ss.setActiveSheet(overviewSheet);
    ss.moveActiveSheet(1);
  }

  Logger.log('=== PPC AUDIT SCRIPT COMPLETE ===');
}


// ============================================================
// HELPERS
// ============================================================

function daysAgo(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, AdsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function writeHeader(sheet, headers, bgColor) {
  bgColor = bgColor || '#1a1a2e';
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setFontColor('#FFFFFF');
  range.setBackground(bgColor);
  sheet.setFrozenRows(1);
}

function autoResize(sheet, numCols) {
  for (var i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }
}

function r2(n) { return Math.round(n * 100) / 100; }

/**
 * Fetch account-level totals for a given date range.
 *
 * FIX: The Google Ads API returns `Ctr` as a decimal fraction
 * (e.g., 0.0017 = 0.17%). We multiply by 100 ONCE here to convert
 * to a human-readable percentage. The original code multiplied by 100
 * in multiple places, causing inflated CTR values like 169 instead of 0.17%.
 */
function getAccountTotals(dateFrom, dateTo) {
  var t = { cost:0, impressions:0, clicks:0, conversions:0, convValue:0, ctr:0, avgCpc:0, cpa:0 };
  var report = AdsApp.report(
    'SELECT Cost, Impressions, Clicks, Conversions, ConversionValue, Ctr, AverageCpc, CostPerConversion ' +
    'FROM ACCOUNT_PERFORMANCE_REPORT ' +
    'DURING ' + dateFrom + ',' + dateTo
  );
  var iter = report.rows();
  if (iter.hasNext()) {
    var row = iter.next();
    t.cost        = r2(parseFloat(row['Cost']) || 0);
    t.impressions = parseInt(row['Impressions']) || 0;
    t.clicks      = parseInt(row['Clicks']) || 0;
    t.conversions = r2(parseFloat(row['Conversions']) || 0);
    t.convValue   = r2(parseFloat(row['ConversionValue']) || 0);
    // ✅ FIX: Multiply by 100 only once to convert decimal → percentage
    t.ctr         = r2((parseFloat(row['Ctr']) || 0) * 100);
    t.avgCpc      = r2(parseFloat(row['AverageCpc']) || 0);
    t.cpa         = r2(parseFloat(row['CostPerConversion']) || 0);
  }
  return t;
}

/** Count data rows in an audit tab (excludes header + ✅ no-data message) */
function countSheetDataRows(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  var firstCell = sheet.getRange(2, 1).getValue().toString();
  if (firstCell.indexOf('✅') === 0) return 0;
  return lastRow - 1;
}

/** Write a two-row KPI card (label row + value row) */
function writeKpiCard(sheet, labelRow, col, label, value, bgColor, valueColor) {
  sheet.getRange(labelRow, col)
    .setValue(label)
    .setFontSize(9).setFontWeight('bold').setFontColor('#777777')
    .setBackground(bgColor).setHorizontalAlignment('center');
  sheet.getRange(labelRow + 1, col)
    .setValue(value)
    .setFontSize(15).setFontWeight('bold').setFontColor(valueColor || '#1a1a2e')
    .setBackground(bgColor).setHorizontalAlignment('center');
}


// ============================================================
// TAB 0 — OVERVIEW  (runs AFTER all other tabs)
// ============================================================
function runOverview(ss, accountName, tz) {
  Logger.log('Running: Overview...');
  var sheet = getOrCreateSheet(ss, '🏠 Overview');
  var currency = AdsApp.currentAccount().getCurrencyCode();
  var todayLabel = Utilities.formatDate(new Date(), tz, 'MMMM dd, yyyy');

  // ── TITLE ──────────────────────────────────────────────────
  sheet.getRange('A1:H1').merge().setValue('📊  Google Ads PPC Audit Report')
    .setFontSize(22).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#1a1a2e').setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 44);

  sheet.getRange('A2:H2').merge()
    .setValue(accountName + '   |   ' + todayLabel + '   |   Currency: ' + currency)
    .setFontSize(11).setFontColor('#555555').setBackground('#f4f6f8');
  sheet.setRowHeight(2, 28);

  // ── 30-DAY KPI CARDS ───────────────────────────────────────
  var acct30 = getAccountTotals(daysAgo(30), daysAgo(1));

  sheet.getRange('A4:H4').merge().setValue('ACCOUNT SNAPSHOT  —  Last 30 Days')
    .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#34495e').setHorizontalAlignment('left');
  sheet.setRowHeight(4, 24);

  var kpis = [
    { label: 'Total Spend',    value: currency + ' ' + r2(acct30.cost),    bg: '#fef9f0', color: '#c0392b' },
    { label: 'Impressions',    value: acct30.impressions.toLocaleString(), bg: '#eaf4fb', color: '#1a5276' },
    { label: 'Clicks',         value: acct30.clicks.toLocaleString(),      bg: '#eafaf1', color: '#1e8449' },
    // ✅ FIX: CTR value already correctly converted in getAccountTotals()
    { label: 'CTR',            value: r2(acct30.ctr) + '%',                bg: '#f5eef8', color: '#6c3483' },
    { label: 'Avg CPC',        value: currency + ' ' + r2(acct30.avgCpc), bg: '#fdf2f8', color: '#922b21' },
    { label: 'Conversions',    value: r2(acct30.conversions),              bg: '#edfbf7', color: '#1a7a4a' },
    { label: 'CPA',            value: currency + ' ' + r2(acct30.cpa),    bg: '#fefbeb', color: '#9a6b00' },
    { label: 'Conv. Value',    value: currency + ' ' + r2(acct30.convValue), bg: '#eaf0fb', color: '#1a3a8e' },
  ];
  for (var i = 0; i < kpis.length; i++) {
    writeKpiCard(sheet, 5, i + 1, kpis[i].label, kpis[i].value, kpis[i].bg, kpis[i].color);
    sheet.setColumnWidth(i + 1, 112);
    sheet.setRowHeight(5, 22);
    sheet.setRowHeight(6, 32);
  }

  // ── WEEK-OVER-WEEK SNAPSHOT ────────────────────────────────
  var acct7  = getAccountTotals(daysAgo(7),  daysAgo(1));
  var acct7p = getAccountTotals(daysAgo(14), daysAgo(8));

  sheet.getRange('A8:H8').merge().setValue('WEEK-OVER-WEEK SNAPSHOT  —  Last 7 Days vs Prior 7 Days')
    .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#0d3349').setHorizontalAlignment('left');
  sheet.setRowHeight(8, 24);

  var wowHdrs = [['Metric', 'This Week', 'Last Week', '∆ Diff', '∆ %', '', '', '']];
  sheet.getRange('A9:E9').setValues([wowHdrs[0].slice(0,5)])
    .setFontWeight('bold').setBackground('#d6eaf8').setFontColor('#1a1a2e');

  function fmtDiff(nw, ow) { var d = r2(nw - ow); return (d >= 0 ? '+' : '') + d; }
  function fmtPct(nw, ow)  {
    if (ow === 0) return nw > 0 ? '+100%' : '0%';
    var p = r2(((nw - ow) / ow) * 100);
    return (p >= 0 ? '+' : '') + p + '%';
  }

  // ✅ FIX: acct7.ctr and acct7p.ctr are already correct percentages
  // from getAccountTotals() — no extra multiplication needed here
  var wowData = [
    ['Spend (' + currency + ')', r2(acct7.cost),        r2(acct7p.cost),        fmtDiff(acct7.cost, acct7p.cost),        fmtPct(acct7.cost, acct7p.cost),        true  ],
    ['Impressions',              acct7.impressions,      acct7p.impressions,      fmtDiff(acct7.impressions, acct7p.impressions), fmtPct(acct7.impressions, acct7p.impressions), true],
    ['Clicks',                   acct7.clicks,           acct7p.clicks,           fmtDiff(acct7.clicks, acct7p.clicks),           fmtPct(acct7.clicks, acct7p.clicks),           true ],
    ['CTR (%)',                  r2(acct7.ctr),           r2(acct7p.ctr),          fmtDiff(acct7.ctr, acct7p.ctr),                fmtPct(acct7.ctr, acct7p.ctr),                true ],
    ['Avg CPC (' + currency + ')', r2(acct7.avgCpc),    r2(acct7p.avgCpc),       fmtDiff(acct7.avgCpc, acct7p.avgCpc),           fmtPct(acct7.avgCpc, acct7p.avgCpc),           false],
    ['Conversions',              r2(acct7.conversions),  r2(acct7p.conversions),  fmtDiff(acct7.conversions, acct7p.conversions),  fmtPct(acct7.conversions, acct7p.conversions),  true ],
    ['CPA (' + currency + ')',   r2(acct7.cpa),          r2(acct7p.cpa),          fmtDiff(acct7.cpa, acct7p.cpa),                 fmtPct(acct7.cpa, acct7p.cpa),                 false],
    ['Conv. Value',              r2(acct7.convValue),    r2(acct7p.convValue),    fmtDiff(acct7.convValue, acct7p.convValue),      fmtPct(acct7.convValue, acct7p.convValue),      true ],
  ];

  for (var wi = 0; wi < wowData.length; wi++) {
    var wr = 10 + wi;
    var higherIsBetter = wowData[wi][5];
    sheet.getRange(wr, 1).setValue(wowData[wi][0]).setFontWeight('bold');
    sheet.getRange(wr, 2).setValue(wowData[wi][1]);
    sheet.getRange(wr, 3).setValue(wowData[wi][2]);
    sheet.getRange(wr, 4).setValue(wowData[wi][3]);
    sheet.getRange(wr, 5).setValue(wowData[wi][4]);

    var rawPct = parseFloat(wowData[wi][4]);
    var isGood = higherIsBetter ? rawPct >= 0 : rawPct <= 0;
    var deltaBg = rawPct === 0 ? '#ffffff' : (isGood ? '#e8f5e9' : '#fde8e8');
    sheet.getRange(wr, 4, 1, 2).setBackground(deltaBg)
      .setFontColor(rawPct === 0 ? '#333333' : (isGood ? '#1e8449' : '#c0392b'))
      .setFontWeight('bold');
    sheet.getRange(wr, 1, 1, 3).setBackground(wi % 2 === 0 ? '#f9f9f9' : '#ffffff');
  }

  // ── HEALTH SCORECARD ───────────────────────────────────────
  var zombieCount  = countSheetDataRows(ss, '💀 Zombie Campaigns');
  var nonConvCount = countSheetDataRows(ss, '❌ Non-Converting Terms');
  var highCpcCount = countSheetDataRows(ss, '💰 High CPC Terms');
  var negKwCount   = countSheetDataRows(ss, '🛑 Negative KW Candidates');

  var scStart = 10 + wowData.length + 2;
  sheet.getRange(scStart, 1, 1, 5).merge()
    .setValue('🏥  AUDIT HEALTH SCORECARD')
    .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#4a0e8f');
  sheet.setRowHeight(scStart, 24);

  sheet.getRange(scStart + 1, 1, 1, 3)
    .setValues([['Check', 'Status', 'Detail']])
    .setFontWeight('bold').setBackground('#ede7f6');

  var checks = [
    ['💀 Zombie Campaigns',          zombieCount  === 0, zombieCount,  'campaign(s) spending with zero conversions in 14 days',    'No campaigns spending without conversions'],
    ['❌ Non-Converting Terms',       nonConvCount === 0, nonConvCount, 'search term(s) spending without converting in 14 days',     'No wasted spend on search terms'],
    ['💰 High CPC Terms',            highCpcCount === 0, highCpcCount, 'term(s) above ' + CONFIG.HIGH_CPC_THRESHOLD + ' CPC with no conversions', 'No terms above CPC threshold'],
    ['🛑 Negative KW Candidates',    negKwCount   === 0, negKwCount,   'term(s) suggested as negatives to save budget',             'No negative keyword candidates found'],
  ];

  var passed = 0;
  for (var ci = 0; ci < checks.length; ci++) {
    var cr = scStart + 2 + ci;
    var clean = checks[ci][1];
    var count = checks[ci][2];
    if (clean) passed++;
    var status = clean ? '✅ Clean' : '⚠️ ' + count + ' Issue(s)';
    var detail = clean ? checks[ci][4] : count + ' ' + checks[ci][3];
    sheet.getRange(cr, 1).setValue(checks[ci][0]).setFontWeight('bold');
    sheet.getRange(cr, 2).setValue(status).setFontWeight('bold')
      .setFontColor(clean ? '#1e8449' : '#c0392b');
    sheet.getRange(cr, 3).setValue(detail).setFontColor('#444444');
    sheet.getRange(cr, 1, 1, 3).setBackground(ci % 2 === 0 ? '#f9f9f9' : '#ffffff');
  }

  // Overall score banner
  var totalChecks = checks.length;
  var scoreRow = scStart + 2 + totalChecks + 1;
  var scoreBg    = passed === totalChecks ? '#e8f5e9' : (passed >= 2 ? '#fff8e1' : '#fde8e8');
  var scoreColor = passed === totalChecks ? '#1e8449' : (passed >= 2 ? '#9a6b00' : '#c0392b');
  var scoreEmoji = passed === totalChecks ? '🎉' : (passed >= 2 ? '⚠️' : '🚨');
  sheet.getRange(scoreRow, 1, 1, 3).merge()
    .setValue(scoreEmoji + '  Overall Score: ' + passed + ' / ' + totalChecks + ' checks passed')
    .setFontSize(12).setFontWeight('bold')
    .setFontColor(scoreColor).setBackground(scoreBg)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(scoreRow, 30);

  // Column widths
  sheet.setColumnWidth(1, 235);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 390);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 80);

  Logger.log('Overview complete. Score: ' + passed + '/' + totalChecks);
}


// ============================================================
// TAB 1 — ZOMBIE CAMPAIGNS
// ============================================================
function runZombieCampaigns(ss) {
  Logger.log('Running: Zombie Campaigns...');
  var sheet = getOrCreateSheet(ss, '💀 Zombie Campaigns');

  var headers = [
    'Campaign Name', 'Status', 'Cost (14d)', 'Clicks (14d)',
    'Impressions (14d)', 'Conversions (14d)', 'CPC (14d)'
  ];
  writeHeader(sheet, headers, '#7b0d1e');

  var rows = [];
  var dateFrom = daysAgo(14);
  var dateTo   = daysAgo(1);

  var report = AdsApp.report(
    'SELECT CampaignName, CampaignStatus, Cost, Clicks, Impressions, Conversions ' +
    'FROM CAMPAIGN_PERFORMANCE_REPORT ' +
    'WHERE CampaignStatus IN ["ENABLED", "PAUSED"] ' +
    'DURING ' + dateFrom + ',' + dateTo
  );

  var iter = report.rows();
  while (iter.hasNext()) {
    var row = iter.next();
    var cost = parseFloat(row['Cost']) || 0;
    var conversions = parseFloat(row['Conversions']) || 0;
    var clicks = parseInt(row['Clicks']) || 0;

    if (cost >= CONFIG.MIN_ZOMBIE_SPEND && conversions === 0) {
      var cpc = clicks > 0 ? r2(cost / clicks) : 0;
      rows.push([
        row['CampaignName'], row['CampaignStatus'],
        r2(cost), clicks, parseInt(row['Impressions']) || 0, conversions, cpc
      ]);
    }
  }

  rows.sort(function(a, b) { return b[2] - a[2]; });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 3, rows.length, 1).setBackground('#ffe0e0');
  } else {
    sheet.getRange(2, 1).setValue('✅ No zombie campaigns found in the last 14 days.');
  }

  autoResize(sheet, headers.length);
  Logger.log('Zombie campaigns found: ' + rows.length);
}


// ============================================================
// TAB 2 — NON-CONVERTING SEARCH TERMS
// ============================================================
function runNonConvertingSearchTerms(ss) {
  Logger.log('Running: Non-Converting Search Terms...');
  var sheet = getOrCreateSheet(ss, '❌ Non-Converting Terms');

  var headers = [
    'Search Term', 'Campaign', 'Ad Group', 'Match Type',
    'Impressions', 'Clicks', 'Cost', 'CTR%', 'Avg CPC', 'Conversions'
  ];
  writeHeader(sheet, headers, '#b5451b');

  var rows = [];
  var dateFrom = daysAgo(14);
  var dateTo   = daysAgo(1);

  var report = AdsApp.report(
    'SELECT Query, CampaignName, AdGroupName, QueryMatchTypeWithVariant, ' +
    'Impressions, Clicks, Cost, Ctr, AverageCpc, Conversions ' +
    'FROM SEARCH_QUERY_PERFORMANCE_REPORT ' +
    'WHERE Conversions = 0 ' +
    'AND Impressions >= ' + CONFIG.MIN_IMPRESSIONS_FOR_AUDIT + ' ' +
    'DURING ' + dateFrom + ',' + dateTo
  );

  var iter = report.rows();
  while (iter.hasNext()) {
    var row = iter.next();
    var cost = parseFloat(row['Cost']) || 0;
    if (cost > 0) {
      rows.push([
        row['Query'], row['CampaignName'], row['AdGroupName'],
        row['QueryMatchTypeWithVariant'],
        parseInt(row['Impressions']) || 0,
        parseInt(row['Clicks']) || 0,
        r2(cost),
        // ✅ FIX: Multiply by 100 only once — API returns decimal fraction
        r2((parseFloat(row['Ctr']) || 0) * 100),
        r2(parseFloat(row['AverageCpc'])) || 0,
        0
      ]);
    }
  }

  rows.sort(function(a, b) { return b[6] - a[6]; });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 7, rows.length, 1).setBackground('#fff3cd');
  } else {
    sheet.getRange(2, 1).setValue('✅ No non-converting search terms found.');
  }

  autoResize(sheet, headers.length);
  Logger.log('Non-converting search terms found: ' + rows.length);
}


// ============================================================
// TAB 3 — HIGH CPC SEARCH TERMS
// ============================================================
function runHighCpcSearchTerms(ss) {
  Logger.log('Running: High CPC Search Terms...');
  var sheet = getOrCreateSheet(ss, '💰 High CPC Terms');

  var headers = [
    'Search Term', 'Campaign', 'Ad Group', 'Clicks',
    'Cost', 'Avg CPC', 'Conversions', 'CPC Threshold'
  ];
  writeHeader(sheet, headers, '#2d5016');

  var rows = [];
  var dateFrom  = daysAgo(14);
  var dateTo    = daysAgo(1);
  var threshold = CONFIG.HIGH_CPC_THRESHOLD;

  var report = AdsApp.report(
    'SELECT Query, CampaignName, AdGroupName, Clicks, Cost, AverageCpc, Conversions ' +
    'FROM SEARCH_QUERY_PERFORMANCE_REPORT ' +
    'WHERE AverageCpc > ' + threshold + ' ' +
    'AND Conversions = 0 ' +
    'AND Impressions >= ' + CONFIG.MIN_IMPRESSIONS_FOR_AUDIT + ' ' +
    'DURING ' + dateFrom + ',' + dateTo
  );

  var iter = report.rows();
  while (iter.hasNext()) {
    var row = iter.next();
    rows.push([
      row['Query'], row['CampaignName'], row['AdGroupName'],
      parseInt(row['Clicks']) || 0,
      r2(parseFloat(row['Cost'])) || 0,
      r2(parseFloat(row['AverageCpc'])) || 0,
      parseFloat(row['Conversions']) || 0,
      threshold
    ]);
  }

  rows.sort(function(a, b) { return b[5] - a[5]; });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 6, rows.length, 1).setBackground('#fde8e8');
  } else {
    sheet.getRange(2, 1).setValue('✅ No high-CPC non-converting terms found.');
  }

  autoResize(sheet, headers.length);
  Logger.log('High CPC terms found: ' + rows.length);
}


// ============================================================
// TAB 4 — NEGATIVE KEYWORD CANDIDATES
// ============================================================
function runNegativeKeywordCandidates(ss) {
  Logger.log('Running: Negative Keyword Candidates...');
  var sheet = getOrCreateSheet(ss, '🛑 Negative KW Candidates');

  var headers = [
    'Search Term', 'Campaign', 'Ad Group', 'Impressions',
    'Clicks', 'Cost', 'Conversions', 'Suggested Match Type', 'Reason'
  ];
  writeHeader(sheet, headers, '#4a0e8f');

  var rows = [];
  var dateFrom = daysAgo(30);
  var dateTo   = daysAgo(1);

  var report = AdsApp.report(
    'SELECT Query, CampaignName, AdGroupName, Impressions, Clicks, Cost, Conversions ' +
    'FROM SEARCH_QUERY_PERFORMANCE_REPORT ' +
    'WHERE Conversions = 0 ' +
    'AND Impressions >= ' + CONFIG.MIN_IMPRESSIONS_FOR_AUDIT + ' ' +
    'DURING ' + dateFrom + ',' + dateTo
  );

  var iter = report.rows();
  while (iter.hasNext()) {
    var row = iter.next();
    var cost        = parseFloat(row['Cost']) || 0;
    var clicks      = parseInt(row['Clicks']) || 0;
    var impressions = parseInt(row['Impressions']) || 0;
    var ctr         = impressions > 0 ? clicks / impressions : 0;

    var reason = '';
    if (cost > CONFIG.HIGH_CPC_THRESHOLD * 3) {
      reason = 'High spend, zero conversions';
    } else if (impressions > 100 && ctr < 0.005) {
      reason = 'High impressions, very low CTR';
    } else if (clicks > 20 && cost > 0) {
      reason = 'Many clicks, zero conversions';
    }

    if (reason) {
      rows.push([
        row['Query'], row['CampaignName'], row['AdGroupName'],
        impressions, clicks, r2(cost), 0, 'Broad', reason
      ]);
    }
  }

  rows.sort(function(a, b) { return b[5] - a[5]; });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 9, rows.length, 1).setFontStyle('italic');
  } else {
    sheet.getRange(2, 1).setValue('✅ No negative keyword candidates identified.');
  }

  autoResize(sheet, headers.length);
  Logger.log('Negative KW candidates found: ' + rows.length);
}


// ============================================================
// TAB 5 — WEEK-OVER-WEEK PERFORMANCE
// ============================================================
function runWeekOverWeek(ss) {
  Logger.log('Running: Week-over-Week Performance...');
  var sheet = getOrCreateSheet(ss, '📈 Week-over-Week');

  var headers = [
    'Campaign',
    'Impr (This Wk)', 'Clicks (This Wk)', 'Cost (This Wk)',
    'Conv (This Wk)', 'CTR% (This Wk)', 'CPA (This Wk)', 'Conv Value (This Wk)',
    'Impr (Last Wk)', 'Clicks (Last Wk)', 'Cost (Last Wk)',
    'Conv (Last Wk)', 'CTR% (Last Wk)', 'CPA (Last Wk)', 'Conv Value (Last Wk)',
    '∆ Impr%', '∆ Clicks%', '∆ Cost%', '∆ Conv%', '∆ CPA%'
  ];
  writeHeader(sheet, headers, '#0d3349');

  var thisWeekFrom = daysAgo(7);
  var thisWeekTo   = daysAgo(1);
  var lastWeekFrom = daysAgo(14);
  var lastWeekTo   = daysAgo(8);

  /**
   * FIX: CTR is returned by the API as a decimal fraction.
   * Multiply by 100 exactly once here to get the correct percentage.
   * The original code multiplied by 100 here AND again elsewhere.
   */
  function fetchCampaignData(dateFrom, dateTo) {
    var data = {};
    var report = AdsApp.report(
      'SELECT CampaignName, Impressions, Clicks, Cost, Conversions, Ctr, ' +
      'CostPerConversion, ConversionValue ' +
      'FROM CAMPAIGN_PERFORMANCE_REPORT ' +
      'WHERE CampaignStatus IN ["ENABLED", "PAUSED"] ' +
      'DURING ' + dateFrom + ',' + dateTo
    );
    var iter = report.rows();
    while (iter.hasNext()) {
      var row = iter.next();
      data[row['CampaignName']] = {
        impressions: parseInt(row['Impressions']) || 0,
        clicks:      parseInt(row['Clicks']) || 0,
        cost:        r2(parseFloat(row['Cost'])) || 0,
        conversions: r2(parseFloat(row['Conversions'])) || 0,
        // ✅ FIX: Multiply by 100 once only — API returns decimal fraction
        ctr:         r2((parseFloat(row['Ctr']) || 0) * 100),
        cpa:         r2(parseFloat(row['CostPerConversion'])) || 0,
        convValue:   r2(parseFloat(row['ConversionValue'])) || 0
      };
    }
    return data;
  }

  function pctChange(nw, ow) {
    if (ow === 0) return nw > 0 ? 100 : 0;
    return r2(((nw - ow) / ow) * 100);
  }

  var thisWeek = fetchCampaignData(thisWeekFrom, thisWeekTo);
  var lastWeek = fetchCampaignData(lastWeekFrom, lastWeekTo);

  var campaigns = {};
  Object.keys(thisWeek).forEach(function(k) { campaigns[k] = true; });
  Object.keys(lastWeek).forEach(function(k) { campaigns[k] = true; });

  var rows = [];
  var empty = { impressions:0, clicks:0, cost:0, conversions:0, ctr:0, cpa:0, convValue:0 };

  Object.keys(campaigns).sort().forEach(function(name) {
    var tw = thisWeek[name] || empty;
    var lw = lastWeek[name] || empty;
    rows.push([
      name,
      tw.impressions, tw.clicks, tw.cost, tw.conversions, tw.ctr, tw.cpa, tw.convValue,
      lw.impressions, lw.clicks, lw.cost, lw.conversions, lw.ctr, lw.cpa, lw.convValue,
      pctChange(tw.impressions, lw.impressions),
      pctChange(tw.clicks, lw.clicks),
      pctChange(tw.cost, lw.cost),
      pctChange(tw.conversions, lw.conversions),
      pctChange(tw.cpa, lw.cpa)
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    var deltaStartCol = 16;
    var cpaDeltaCol   = 20;
    for (var r = 0; r < rows.length; r++) {
      for (var c = deltaStartCol; c <= 20; c++) {
        var val = rows[r][c - 1];
        var isCpa = (c === cpaDeltaCol);
        var good = isCpa ? val <= 0 : val >= 0;
        var bg = val === 0 ? '#ffffff' : (good ? '#e8f5e9' : '#fde8e8');
        sheet.getRange(r + 2, c).setBackground(bg);
      }
    }
  } else {
    sheet.getRange(2, 1).setValue('No campaign data found for the selected date ranges.');
  }

  autoResize(sheet, headers.length);
  Logger.log('Week-over-Week campaigns processed: ' + rows.length);
}
