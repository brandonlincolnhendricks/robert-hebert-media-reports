/**
 * Robert Hebert Media - Google Ads MCC Data Export Script
 *
 * Install this script at the MCC (Manager Account) level
 * It will automatically pull data from all 3 client accounts
 * Schedule to run every Sunday at 11pm
 */

// ===========================================
// CONFIGURATION
// ===========================================

// MUST match the spreadsheet the Apps Script report generator is bound to
// (Extensions > Apps Script on that sheet). They were previously two different
// sheets, so exported data never reached the generator. Do not change without
// updating the bound report sheet to match.
var SPREADSHEET_ID = '1NRlxowBBoDkaOhVBCgkFqVMMhMtIFYNNfAgoJ5vcpps';

// Map Google Ads Customer IDs to spreadsheet rows
// Format: 'XXX-XXX-XXXX': row number
var CLIENT_CONFIG = {
  '917-597-4799': { name: 'JFTx2025', row: 2 },
  '343-027-6201': { name: 'PFBHNC', row: 3 },
  '326-336-6442': { name: 'ReOptica', row: 4 }
};

// ===========================================
// MAIN FUNCTION - Schedule this weekly
// ===========================================

function main() {
  Logger.log('Starting RHM Weekly Data Export...');
  Logger.log('Processing ' + Object.keys(CLIENT_CONFIG).length + ' accounts');

  // Calculate date ranges
  var dates = calculateDateRanges();
  Logger.log('This week: ' + dates.thisWeekStart + ' to ' + dates.thisWeekEnd);
  Logger.log('Previous week: ' + dates.prevWeekStart + ' to ' + dates.prevWeekEnd);

  // Open spreadsheet
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var thisWeekSheet = spreadsheet.getSheetByName('This Week');
  var prevWeekSheet = spreadsheet.getSheetByName('Previous Week');
  var dateRangeSheet = spreadsheet.getSheetByName('Date Range');

  // Update date ranges first
  dateRangeSheet.getRange('B2:B5').setValues([
    [dates.thisWeekStartISO],
    [dates.thisWeekEndISO],
    [dates.prevWeekStartISO],
    [dates.prevWeekEndISO]
  ]);
  Logger.log('Date ranges updated');

  // Process ONLY the 3 configured accounts. Without withIds(), the selector
  // walks the entire MCC tree, which blows the 30-minute limit on a large
  // agency manager account. withIds() fetches just these accounts directly.
  var accountSelector = AdsManagerApp.accounts().withIds(Object.keys(CLIENT_CONFIG));
  var accountIterator = accountSelector.get();

  var processedCount = 0;

  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var customerId = account.getCustomerId();

    // Check if this account is in our config
    var config = CLIENT_CONFIG[customerId];
    if (!config) {
      Logger.log('Skipping account ' + customerId + ' (not in config)');
      continue;
    }

    Logger.log('Processing: ' + config.name + ' (' + customerId + ')');

    // Switch to this account
    AdsManagerApp.select(account);

    // Get performance data
    var thisWeekData = getPerformanceData(dates.thisWeekStart, dates.thisWeekEnd);
    var prevWeekData = getPerformanceData(dates.prevWeekStart, dates.prevWeekEnd);

    Logger.log('  This week - Spend: $' + thisWeekData.spend + ', Impressions: ' + thisWeekData.impressions + ', Clicks: ' + thisWeekData.clicks);
    Logger.log('  Prev week - Spend: $' + prevWeekData.spend + ', Impressions: ' + prevWeekData.impressions + ', Clicks: ' + prevWeekData.clicks);

    // Update spreadsheet
    thisWeekSheet.getRange(config.row, 1, 1, 6).setValues([
      [config.name, thisWeekData.spend, thisWeekData.impressions, thisWeekData.clicks, thisWeekData.conversions, thisWeekData.views]
    ]);

    prevWeekSheet.getRange(config.row, 1, 1, 6).setValues([
      [config.name, prevWeekData.spend, prevWeekData.impressions, prevWeekData.clicks, prevWeekData.conversions, prevWeekData.views]
    ]);

    processedCount++;
    Logger.log('  ✓ Updated row ' + config.row);
  }

  Logger.log('');
  Logger.log('=== COMPLETE ===');
  Logger.log('Processed ' + processedCount + ' of ' + Object.keys(CLIENT_CONFIG).length + ' accounts');
  Logger.log('Spreadsheet updated: https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID);
}

// ===========================================
// DATA RETRIEVAL
// ===========================================

function getPerformanceData(startDate, endDate) {
  var report = AdsApp.report(
    'SELECT Cost, Impressions, Clicks, Conversions, VideoViews ' +
    'FROM ACCOUNT_PERFORMANCE_REPORT ' +
    'DURING ' + startDate + ',' + endDate
  );

  var rows = report.rows();
  var data = {
    spend: 0,
    impressions: 0,
    clicks: 0,
    conversions: 0,
    views: 0
  };

  while (rows.hasNext()) {
    var row = rows.next();
    var cost = parseFloat(row['Cost'].replace(/,/g, '')) || 0;
    data.spend += cost;
    data.impressions += parseInt(row['Impressions'].replace(/,/g, ''), 10) || 0;
    data.clicks += parseInt(row['Clicks'].replace(/,/g, ''), 10) || 0;
    data.conversions += parseFloat(row['Conversions'].replace(/,/g, '')) || 0;
    data.views += parseInt(row['VideoViews'].replace(/,/g, ''), 10) || 0;
  }

  data.spend = Math.round(data.spend * 100) / 100;
  data.conversions = Math.round(data.conversions * 100) / 100;
  return data;
}

// ===========================================
// DATE CALCULATIONS
// ===========================================

function calculateDateRanges() {
  var today = new Date();

  // This week ends yesterday
  var thisWeekEnd = new Date(today);
  thisWeekEnd.setDate(today.getDate() - 1);

  // This week starts 6 days before end (7 days total)
  var thisWeekStart = new Date(thisWeekEnd);
  thisWeekStart.setDate(thisWeekEnd.getDate() - 6);

  // Previous week ends day before this week starts
  var prevWeekEnd = new Date(thisWeekStart);
  prevWeekEnd.setDate(thisWeekStart.getDate() - 1);

  // Previous week starts 6 days before its end
  var prevWeekStart = new Date(prevWeekEnd);
  prevWeekStart.setDate(prevWeekEnd.getDate() - 6);

  return {
    thisWeekStart: formatDateForQuery(thisWeekStart),
    thisWeekEnd: formatDateForQuery(thisWeekEnd),
    prevWeekStart: formatDateForQuery(prevWeekStart),
    prevWeekEnd: formatDateForQuery(prevWeekEnd),
    thisWeekStartISO: formatDateISO(thisWeekStart),
    thisWeekEndISO: formatDateISO(thisWeekEnd),
    prevWeekStartISO: formatDateISO(prevWeekStart),
    prevWeekEndISO: formatDateISO(prevWeekEnd)
  };
}

function formatDateForQuery(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + month + day;
}

function formatDateISO(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

// ===========================================
// PREVIEW - Run this to test without updating
// ===========================================

function preview() {
  Logger.log('=== PREVIEW MODE ===');
  Logger.log('This will show what accounts are accessible and their IDs');
  Logger.log('');

  var dates = calculateDateRanges();
  Logger.log('Date ranges that would be used:');
  Logger.log('  This week: ' + dates.thisWeekStartISO + ' to ' + dates.thisWeekEndISO);
  Logger.log('  Prev week: ' + dates.prevWeekStartISO + ' to ' + dates.prevWeekEndISO);
  Logger.log('');

  Logger.log('Accounts in this MCC:');
  var accountIterator = AdsManagerApp.accounts().get();

  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var customerId = account.getCustomerId();
    var name = account.getName();
    var config = CLIENT_CONFIG[customerId];

    if (config) {
      Logger.log('  ✓ ' + name + ' (' + customerId + ') -> Row ' + config.row + ' as "' + config.name + '"');
    } else {
      Logger.log('  ✗ ' + name + ' (' + customerId + ') -> NOT IN CONFIG');
    }
  }

  Logger.log('');
  Logger.log('Accounts configured but not found in MCC:');
  for (var cid in CLIENT_CONFIG) {
    Logger.log('  - ' + CLIENT_CONFIG[cid].name + ' (' + cid + ')');
  }
}
