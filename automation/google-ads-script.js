/**
 * Robert Hebert Media - Google Ads Data Export Script
 *
 * Install this script in each Google Ads account (JFTx2025, PFBHNC, ReOptica)
 * Schedule to run every Sunday at 11pm
 *
 * This script exports weekly performance data to the RHM Google Sheet
 */

// ===========================================
// CONFIGURATION - Update these values
// ===========================================

// MUST match the spreadsheet the Apps Script report generator is bound to
// (Extensions > Apps Script on that sheet). They were previously two different
// sheets, so exported data never reached the generator.
var SPREADSHEET_ID = '1NRlxowBBoDkaOhVBCgkFqVMMhMtIFYNNfAgoJ5vcpps';

// Map Google Ads account names to spreadsheet row numbers
// Row 2 = first data row (row 1 is headers)
var CLIENT_ROW_MAP = {
  'JFTx2025': 2,
  'PFBHNC': 3,
  'ReOptica': 4
};

// ===========================================
// MAIN FUNCTION - Schedule this to run weekly
// ===========================================

function main() {
  // Get account name to determine which row to update
  var accountName = AdsApp.currentAccount().getName();
  var row = getRowForAccount(accountName);

  if (!row) {
    Logger.log('ERROR: Could not find row mapping for account: ' + accountName);
    Logger.log('Please update CLIENT_ROW_MAP with this account name.');
    return;
  }

  Logger.log('Processing account: ' + accountName + ' (Row ' + row + ')');

  // Calculate date ranges
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.

  // This week: Monday to Sunday (or up to today if mid-week)
  // For Sunday run: get Mon-Sat of current week
  var thisWeekEnd = new Date(today);
  thisWeekEnd.setDate(today.getDate() - 1); // Yesterday (Saturday if running Sunday)

  var thisWeekStart = new Date(thisWeekEnd);
  thisWeekStart.setDate(thisWeekEnd.getDate() - 6); // 7 days total

  // Previous week: 7 days before this week
  var prevWeekEnd = new Date(thisWeekStart);
  prevWeekEnd.setDate(thisWeekStart.getDate() - 1);

  var prevWeekStart = new Date(prevWeekEnd);
  prevWeekStart.setDate(prevWeekEnd.getDate() - 6);

  Logger.log('This week: ' + formatDate(thisWeekStart) + ' to ' + formatDate(thisWeekEnd));
  Logger.log('Previous week: ' + formatDate(prevWeekStart) + ' to ' + formatDate(prevWeekEnd));

  // Get performance data
  var thisWeekData = getPerformanceData(thisWeekStart, thisWeekEnd);
  var prevWeekData = getPerformanceData(prevWeekStart, prevWeekEnd);

  Logger.log('This week - Spend: $' + thisWeekData.spend + ', Impressions: ' + thisWeekData.impressions + ', Clicks: ' + thisWeekData.clicks);
  Logger.log('Previous week - Spend: $' + prevWeekData.spend + ', Impressions: ' + prevWeekData.impressions + ', Clicks: ' + prevWeekData.clicks);

  // Update Google Sheet
  updateSpreadsheet(row, accountName, thisWeekData, prevWeekData, thisWeekStart, thisWeekEnd, prevWeekStart, prevWeekEnd);

  Logger.log('Successfully updated spreadsheet for ' + accountName);
}

// ===========================================
// DATA RETRIEVAL
// ===========================================

function getPerformanceData(startDate, endDate) {
  var report = AdsApp.report(
    'SELECT Cost, Impressions, Clicks, Conversions, VideoViews ' +
    'FROM ACCOUNT_PERFORMANCE_REPORT ' +
    'DURING ' + formatDateForQuery(startDate) + ',' + formatDateForQuery(endDate)
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

  // Round spend to 2 decimal places
  data.spend = Math.round(data.spend * 100) / 100;
  data.conversions = Math.round(data.conversions * 100) / 100;

  return data;
}

// ===========================================
// SPREADSHEET UPDATE
// ===========================================

function updateSpreadsheet(row, accountName, thisWeekData, prevWeekData, thisWeekStart, thisWeekEnd, prevWeekStart, prevWeekEnd) {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Update "This Week" sheet
  var thisWeekSheet = spreadsheet.getSheetByName('This Week');
  thisWeekSheet.getRange(row, 1, 1, 6).setValues([
    [accountName, thisWeekData.spend, thisWeekData.impressions, thisWeekData.clicks, thisWeekData.conversions, thisWeekData.views]
  ]);

  // Update "Previous Week" sheet
  var prevWeekSheet = spreadsheet.getSheetByName('Previous Week');
  prevWeekSheet.getRange(row, 1, 1, 6).setValues([
    [accountName, prevWeekData.spend, prevWeekData.impressions, prevWeekData.clicks, prevWeekData.conversions, prevWeekData.views]
  ]);

  // Update "Date Range" sheet
  var dateRangeSheet = spreadsheet.getSheetByName('Date Range');
  dateRangeSheet.getRange('B2:B5').setValues([
    [formatDateISO(thisWeekStart)],
    [formatDateISO(thisWeekEnd)],
    [formatDateISO(prevWeekStart)],
    [formatDateISO(prevWeekEnd)]
  ]);

  Logger.log('Spreadsheet updated successfully');
}

// ===========================================
// HELPER FUNCTIONS
// ===========================================

function getRowForAccount(accountName) {
  // Try exact match first
  if (CLIENT_ROW_MAP[accountName]) {
    return CLIENT_ROW_MAP[accountName];
  }

  // Try partial match (account name contains the key)
  for (var key in CLIENT_ROW_MAP) {
    if (accountName.toLowerCase().indexOf(key.toLowerCase()) !== -1) {
      return CLIENT_ROW_MAP[key];
    }
  }

  return null;
}

function formatDate(date) {
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
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
// PREVIEW FUNCTION - Run this to test without updating
// ===========================================

function preview() {
  var accountName = AdsApp.currentAccount().getName();
  Logger.log('Account Name: ' + accountName);
  Logger.log('Customer ID: ' + AdsApp.currentAccount().getCustomerId());

  var row = getRowForAccount(accountName);
  Logger.log('Would update row: ' + (row || 'NOT FOUND - update CLIENT_ROW_MAP'));

  // Show what dates would be used
  var today = new Date();
  var thisWeekEnd = new Date(today);
  thisWeekEnd.setDate(today.getDate() - 1);
  var thisWeekStart = new Date(thisWeekEnd);
  thisWeekStart.setDate(thisWeekEnd.getDate() - 6);

  Logger.log('This week range: ' + formatDate(thisWeekStart) + ' to ' + formatDate(thisWeekEnd));

  // Get sample data
  var data = getPerformanceData(thisWeekStart, thisWeekEnd);
  Logger.log('Sample data - Spend: $' + data.spend + ', Impressions: ' + data.impressions + ', Clicks: ' + data.clicks);
}
