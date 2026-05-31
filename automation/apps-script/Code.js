/**
 * Robert Hebert Media - Weekly Report Generator
 * Google Apps Script for automated report generation and deployment
 *
 * Setup:
 * 1. Copy this code to Extensions > Apps Script
 * 2. Configure the sheets as described in google-sheets-setup.md
 * 3. Set up a weekly trigger for generateWeeklyReports()
 */

// ============================================
// MAIN FUNCTION - Run this weekly
// ============================================

function generateWeeklyReports() {
  try {
    return _generateWeeklyReportsImpl();
  } catch (err) {
    notifyFailure_('generateWeeklyReports', err);
    throw err;
  }
}

// Run from the editor anytime auth feels stale. Touches every OAuth scope
// the trigger needs so Google re-prompts for any missing grant.
function forceReauth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getName();
  UrlFetchApp.fetch('https://api.github.com/zen', { muteHttpExceptions: true });
  const quota = MailApp.getRemainingDailyQuota();
  ScriptApp.getProjectTriggers();
  Logger.log('Reauth check passed for ' + Session.getEffectiveUser().getEmail() + ' | mail quota remaining: ' + quota);
}

// ============================================
// AUTOMATION (time-based trigger)
// ============================================
// Run "Turn On Weekly Automation" once from the RHM Reports menu and approve
// the OAuth prompt. This installs a Monday 08:00 (script timezone, America/Chicago)
// trigger on generateWeeklyReports so reports build and deploy with no manual run.
// Schedule the Google Ads data-export script for Sunday 11pm so the sheet is fresh
// before this fires.

var WEEKLY_TRIGGER_HANDLER = 'generateWeeklyReports';

function enableWeeklyAutomation() {
  removeWeeklyAutomation_();
  ScriptApp.newTrigger(WEEKLY_TRIGGER_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  _notify_(
    'Weekly automation is ON.\n\n' +
    'generateWeeklyReports will run every Monday around 8:00 AM America/Chicago.\n' +
    'Make sure the Google Ads export script runs Sunday 11pm so the sheet is fresh,\n' +
    'and that the Config tab holds a valid GitHub PAT (repo scope).'
  );
}

function disableWeeklyAutomation() {
  var removed = removeWeeklyAutomation_();
  _notify_('Weekly automation is OFF. Removed ' + removed + ' trigger(s).');
}

function removeWeeklyAutomation_() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === WEEKLY_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

function _generateWeeklyReportsImpl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get configuration
  const config = getConfig(ss);
  const clients = getClients(ss);
  const dateRange = getDateRange(ss);
  const thisWeekData = getWeekData(ss, 'This Week');
  const prevWeekData = getWeekData(ss, 'Previous Week');

  Logger.log('Starting report generation as: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Date range: ' + dateRange.thisWeek);

  const reportUrls = [];
  const folderSuffix = getFolderSuffix(dateRange);

  // Generate report for each active client
  clients.forEach(client => {
    if (!client.active) return;

    const currentData = thisWeekData[client.name];
    const previousData = prevWeekData[client.name];

    if (!currentData) {
      Logger.log('No data for client: ' + client.name);
      return;
    }

    Logger.log('Generating report for: ' + client.name);

    // Generate HTML report
    const html = generateReportHtml(client, currentData, previousData || {spend: 0, impressions: 0, clicks: 0}, dateRange);

    // Deploy to GitHub
    const folderName = client.slug + '-' + folderSuffix;
    deployToGitHub(config, folderName, html);

    reportUrls.push({
      name: client.name,
      url: 'https://reports.roberthebertmedia.com/' + folderName + '/'
    });
  });

  // Update index.html
  updateIndexHtml(config, clients, dateRange, folderSuffix);

  // Send email notification
  sendEmailNotification(config, reportUrls, dateRange.thisWeek);

  Logger.log('Report generation complete!');
}

function notifyFailure_(fnName, err) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig(ss);
    const to = config.email_bcc || config.email_to || Session.getEffectiveUser().getEmail();
    const subject = 'RHM Report Generator FAILED — ' + fnName;
    const body = [
      'The RHM Weekly Report Generator hit an exception.',
      '',
      'Function: ' + fnName,
      'User: ' + Session.getEffectiveUser().getEmail(),
      'Time: ' + new Date().toString(),
      '',
      'Error: ' + (err && err.message ? err.message : String(err)),
      '',
      'Stack:',
      (err && err.stack) ? err.stack : '(no stack)'
    ].join('\n');
    GmailApp.sendEmail(to, subject, body, { name: 'RHM Report System' });
  } catch (e) {
    // If even the failure email fails (likely the same auth issue),
    // surface it in the Stackdriver log so the daily digest still fires.
    console.error('notifyFailure_ could not send email: ' + e);
  }
}

// ============================================
// DATA READING FUNCTIONS
// ============================================

function getConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};

  data.forEach(row => {
    if (row[0] && row[1]) {
      config[row[0]] = row[1];
    }
  });

  return config;
}

function getClients(ss) {
  const sheet = ss.getSheetByName('Clients');
  const data = sheet.getDataRange().getValues();
  const clients = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      clients.push({
        name: data[i][0],
        slug: data[i][1],
        customerId: data[i][2],
        active: data[i][3] === true || data[i][3] === 'TRUE'
      });
    }
  }

  return clients;
}

function getDateRange(ss) {
  const sheet = ss.getSheetByName('Date Range');
  const data = sheet.getDataRange().getValues();
  const dates = {};

  data.forEach(row => {
    if (row[0] && row[1]) {
      dates[row[0]] = row[1];
    }
  });

  // Format date ranges
  const formatDate = (date) => {
    if (date instanceof Date) {
      return Utilities.formatDate(date, 'America/Chicago', 'MMMM d, yyyy');
    }
    return date;
  };

  return {
    thisWeek: formatDate(dates.this_week_start) + ' &ndash; ' + formatDate(dates.this_week_end),
    prevWeek: formatDate(dates.prev_week_start) + ' &ndash; ' + formatDate(dates.prev_week_end),
    thisWeekStart: dates.this_week_start,
    thisWeekEnd: dates.this_week_end
  };
}

function getWeekData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const weekData = {};

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      weekData[data[i][0]] = {
        spend: parseFloat(data[i][1]) || 0,
        impressions: parseInt(data[i][2]) || 0,
        clicks: parseInt(data[i][3]) || 0,
        conversions: parseFloat(data[i][4]) || 0,
        views: parseInt(data[i][5]) || 0
      };
    }
  }

  return weekData;
}

function getFolderSuffix(dateRange) {
  const start = dateRange.thisWeekStart;
  const end = dateRange.thisWeekEnd;

  const startDate = start instanceof Date ? start : new Date(start);
  const endDate = end instanceof Date ? end : new Date(end);

  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

  const startMonth = months[startDate.getMonth()];
  const endMonth = months[endDate.getMonth()];
  const startDay = startDate.getDate();
  const endDay = endDate.getDate();

  if (startMonth === endMonth) {
    return startMonth + startDay + '-' + endDay;
  } else {
    return startMonth + startDay + '-' + endMonth + endDay;
  }
}

// ============================================
// REPORT GENERATION
// ============================================

function generateReportHtml(client, current, previous, dateRange) {
  // Derived metrics: CTR, CPC, cost-per-lead, conversion rate, view rate, CPV.
  computeDerived_(current);
  computeDerived_(previous);

  const calcChange = (curr, prev) => prev === 0 ? (curr === 0 ? 0 : 100) : ((curr - prev) / prev * 100);
  const changes = {
    spend: calcChange(current.spend, previous.spend),
    impressions: calcChange(current.impressions, previous.impressions),
    clicks: calcChange(current.clicks, previous.clicks),
    ctr: calcChange(current.ctr, previous.ctr),
    cpc: calcChange(current.cpc, previous.cpc),
    conversions: calcChange(current.conversions, previous.conversions),
    cpl: calcChange(current.cpl, previous.cpl),
    convRate: calcChange(current.convRate, previous.convRate),
    views: calcChange(current.views, previous.views),
    viewRate: calcChange(current.viewRate, previous.viewRate),
    cpv: calcChange(current.cpv, previous.cpv)
  };

  // JFTx2025 = video/awareness (impressions, views, clicks); others = lead-gen.
  const profile = getKpiProfile_(client);
  const summary = generateSummary(client.name, current, changes, profile);
  const insights = generateInsights(current, previous, changes, profile);

  const endDate = dateRange.thisWeekEnd instanceof Date ? dateRange.thisWeekEnd : new Date(dateRange.thisWeekEnd);
  const weekNum = getWeekNumber(endDate);
  const reportId = 'RHM-' + client.slug.toUpperCase().substring(0, 3) + '-' + endDate.getFullYear() + '-W' + String(weekNum).padStart(2, '0');

  // Function replacements so '$' in currency/HTML is never read as a regex backreference.
  return getReportTemplate()
    .replace(/{client_name}/g, function () { return client.name; })
    .replace(/{date_range}/g, function () { return dateRange.thisWeek; })
    .replace(/{generated_date}/g, function () { return Utilities.formatDate(new Date(), 'America/Chicago', 'MMMM d, yyyy'); })
    .replace(/{executive_summary}/g, function () { return summary; })
    .replace(/{kpi_cards_html}/g, function () { return buildKpiCards_(profile, current, previous, changes); })
    .replace(/{metric_rows_html}/g, function () { return buildMetricRows_(profile, current, previous, changes); })
    .replace(/{insights_html}/g, function () { return insights; })
    .replace(/{report_id}/g, function () { return reportId; });
}

// ============================================
// KPI PROFILES (client-aware metric sets)
// ============================================
// JFTx2025 is a video / awareness account: lead with impressions, views, clicks.
// PFBHNC and ReOptica are lead-gen: lead with conversions and cost-per-lead.
var KPI_PROFILES = {
  jftx2025: 'video',
  pfbhnc: 'leadgen',
  reoptica: 'leadgen'
};

function getKpiProfile_(client) {
  var key = String(client.slug || client.name || '').toLowerCase();
  return KPI_PROFILES[key] || 'leadgen';
}

function computeDerived_(d) {
  d.spend = d.spend || 0;
  d.impressions = d.impressions || 0;
  d.clicks = d.clicks || 0;
  d.conversions = d.conversions || 0;
  d.views = d.views || 0;
  d.ctr = d.impressions > 0 ? (d.clicks / d.impressions * 100) : 0;
  d.cpc = d.clicks > 0 ? (d.spend / d.clicks) : 0;
  d.cpl = d.conversions > 0 ? (d.spend / d.conversions) : 0;
  d.convRate = d.clicks > 0 ? (d.conversions / d.clicks * 100) : 0;
  d.viewRate = d.impressions > 0 ? (d.views / d.impressions * 100) : 0;
  d.cpv = d.views > 0 ? (d.spend / d.views) : 0;
}

function formatConversions_(value) {
  // Conversions can be fractional (modeled). Show a decimal only when meaningful.
  if (Math.abs(value - Math.round(value)) < 0.05) return Math.round(value).toLocaleString();
  return (Math.round(value * 10) / 10).toLocaleString();
}

function changeLabel_(change, prevVal, invert, suffix) {
  if (!prevVal) return 'New this week';
  return formatChangeText(change, suffix || ' vs last week', invert);
}

function changeColor_(change, prevVal, invert) {
  if (!prevVal) return 'neutral';
  return getChangeClass(change, invert);
}

function kpiCard_(label, value, changeClass, changeText, highlight, badge) {
  return '' +
    '<div class="kpi-card' + (highlight ? ' highlight' : '') + '">' +
    (badge || '') +
    '<div class="kpi-label">' + label + '</div>' +
    '<div class="kpi-value">' + value + '</div>' +
    '<div class="kpi-change ' + changeClass + '">' + changeText + '</div>' +
    '</div>';
}

function buildKpiCards_(profile, current, previous, changes) {
  var cards = [];
  var spendSuffix = current.spend < previous.spend ? ' cost savings' : ' vs last week';
  if (profile === 'video') {
    cards.push(kpiCard_('Impressions', formatNumber(current.impressions),
      changeColor_(changes.impressions, previous.impressions), changeLabel_(changes.impressions, previous.impressions), true));
    cards.push(kpiCard_('Video Views', formatNumber(current.views),
      changeColor_(changes.views, previous.views), changeLabel_(changes.views, previous.views), true));
    cards.push(kpiCard_('Clicks', formatNumber(current.clicks),
      changeColor_(changes.clicks, previous.clicks), changeLabel_(changes.clicks, previous.clicks), true, getClicksBadge(changes.clicks)));
    cards.push(kpiCard_('View Rate', current.viewRate.toFixed(1) + '%',
      changeColor_(changes.viewRate, previous.viewRate), changeLabel_(changes.viewRate, previous.viewRate), false));
    cards.push(kpiCard_('Click-Through Rate', current.ctr.toFixed(2) + '%',
      changeColor_(changes.ctr, previous.ctr), changeLabel_(changes.ctr, previous.ctr), false, getCtrBadge(current.ctr)));
    cards.push(kpiCard_('Total Spend', formatCurrency(current.spend),
      changeColor_(changes.spend, previous.spend, true), changeLabel_(changes.spend, previous.spend, true, spendSuffix), false));
  } else {
    var cplValue = current.conversions > 0 ? formatCurrency(current.cpl) : '&mdash;';
    cards.push(kpiCard_('Conversions (Leads)', formatConversions_(current.conversions),
      changeColor_(changes.conversions, previous.conversions), changeLabel_(changes.conversions, previous.conversions), true));
    cards.push(kpiCard_('Cost Per Lead', cplValue,
      changeColor_(changes.cpl, previous.cpl, true), previous.cpl ? changeLabel_(changes.cpl, previous.cpl, true) : 'New this week', true));
    cards.push(kpiCard_('Conversion Rate', current.convRate.toFixed(2) + '%',
      changeColor_(changes.convRate, previous.convRate), changeLabel_(changes.convRate, previous.convRate), false));
    cards.push(kpiCard_('Total Spend', formatCurrency(current.spend),
      changeColor_(changes.spend, previous.spend, true), changeLabel_(changes.spend, previous.spend, true, spendSuffix), false));
    cards.push(kpiCard_('Clicks', formatNumber(current.clicks),
      changeColor_(changes.clicks, previous.clicks), changeLabel_(changes.clicks, previous.clicks), false, getClicksBadge(changes.clicks)));
    cards.push(kpiCard_('Click-Through Rate', current.ctr.toFixed(2) + '%',
      changeColor_(changes.ctr, previous.ctr), changeLabel_(changes.ctr, previous.ctr), false, getCtrBadge(current.ctr)));
  }
  return cards.join('\n                ');
}

function metricRow_(name, thisVal, prevVal, tableClass, changePct) {
  return '<tr><td class="metric-name">' + name + '</td><td class="metric-value">' + thisVal + '</td><td>' + prevVal + '</td><td class="' + tableClass + '">' + changePct + '</td></tr>';
}

function buildMetricRows_(profile, current, previous, changes) {
  var rows = [];
  rows.push(metricRow_('Total Ad Spend', '$' + current.spend.toFixed(2), '$' + previous.spend.toFixed(2), getTableClass(changes.spend, true), formatChangePct(changes.spend)));
  if (profile === 'video') {
    rows.push(metricRow_('Impressions', formatNumber(current.impressions), formatNumber(previous.impressions), getTableClass(changes.impressions), formatChangePct(changes.impressions)));
    rows.push(metricRow_('Video Views', formatNumber(current.views), formatNumber(previous.views), getTableClass(changes.views), formatChangePct(changes.views)));
    rows.push(metricRow_('View Rate', current.viewRate.toFixed(2) + '%', previous.viewRate.toFixed(2) + '%', getTableClass(changes.viewRate), formatChangePct(changes.viewRate)));
    rows.push(metricRow_('Clicks', formatNumber(current.clicks), formatNumber(previous.clicks), getTableClass(changes.clicks), formatChangePct(changes.clicks)));
    rows.push(metricRow_('Click-Through Rate (CTR)', current.ctr.toFixed(2) + '%', previous.ctr.toFixed(2) + '%', getTableClass(changes.ctr), formatChangePct(changes.ctr)));
    rows.push(metricRow_('Avg. Cost Per View', current.views > 0 ? '$' + current.cpv.toFixed(3) : '&mdash;', previous.views > 0 ? '$' + previous.cpv.toFixed(3) : '&mdash;', getTableClass(changes.cpv, true), formatChangePct(changes.cpv)));
  } else {
    rows.push(metricRow_('Conversions (Leads)', formatConversions_(current.conversions), formatConversions_(previous.conversions), getTableClass(changes.conversions), formatChangePct(changes.conversions)));
    rows.push(metricRow_('Cost Per Lead', current.conversions > 0 ? '$' + current.cpl.toFixed(2) : '&mdash;', previous.conversions > 0 ? '$' + previous.cpl.toFixed(2) : '&mdash;', getTableClass(changes.cpl, true), formatChangePct(changes.cpl)));
    rows.push(metricRow_('Conversion Rate', current.convRate.toFixed(2) + '%', previous.convRate.toFixed(2) + '%', getTableClass(changes.convRate), formatChangePct(changes.convRate)));
    rows.push(metricRow_('Impressions', formatNumber(current.impressions), formatNumber(previous.impressions), getTableClass(changes.impressions), formatChangePct(changes.impressions)));
    rows.push(metricRow_('Clicks', formatNumber(current.clicks), formatNumber(previous.clicks), getTableClass(changes.clicks), formatChangePct(changes.clicks)));
    rows.push(metricRow_('Click-Through Rate (CTR)', current.ctr.toFixed(2) + '%', previous.ctr.toFixed(2) + '%', getTableClass(changes.ctr), formatChangePct(changes.ctr)));
    rows.push(metricRow_('Average CPC', '$' + current.cpc.toFixed(2), '$' + previous.cpc.toFixed(2), getTableClass(changes.cpc, true), formatChangePct(changes.cpc)));
  }
  return rows.join('\n                    ');
}

function buildPrimaryInsight_(profile, current, previous, changes) {
  if (profile === 'leadgen') {
    if (current.conversions === 0) {
      return createInsightCard('warning', 'No Conversions Recorded',
        'No conversions were tracked this week. Recommend verifying conversion tracking is firing and reviewing lead-form and landing-page performance.');
    }
    if (changes.conversions > 20) {
      return createInsightCard('success', 'Lead Volume Growth',
        'Conversions increased ' + changes.conversions.toFixed(0) + '% week-over-week to ' + formatConversions_(current.conversions) + ' leads' +
        (current.cpl > 0 ? ' at a $' + current.cpl.toFixed(2) + ' cost per lead.' : '.'));
    }
    if (changes.cpl < -10) {
      return createInsightCard('success', 'More Efficient Lead Acquisition',
        'Cost per lead dropped ' + Math.abs(changes.cpl).toFixed(0) + '% to $' + current.cpl.toFixed(2) + ', improving campaign efficiency.');
    }
    if (changes.conversions < -20) {
      return createInsightCard('warning', 'Lead Volume Decline',
        'Conversions decreased ' + Math.abs(changes.conversions).toFixed(0) + '% week-over-week. Recommend reviewing budget pacing, search impression share, and landing-page conversion rate.');
    }
    return '';
  }
  // video / awareness
  if (changes.views > 25 || changes.impressions > 25) {
    return createInsightCard('success', 'Expanding Reach',
      'Impressions reached ' + formatNumber(current.impressions) + ' and video views ' + (changes.views >= 0 ? 'rose ' + changes.views.toFixed(0) + '%' : 'held steady') + ' this week, growing top-of-funnel awareness.');
  }
  if (changes.views < -20) {
    return createInsightCard('warning', 'Reach Pullback',
      'Video views declined ' + Math.abs(changes.views).toFixed(0) + '% week-over-week. Recommend reviewing budget pacing and creative rotation to sustain reach.');
  }
  return createInsightCard('success', 'Steady Reach',
    'The campaign delivered ' + formatNumber(current.impressions) + ' impressions and ' + formatNumber(current.views) + ' video views at a ' + current.viewRate.toFixed(1) + '% view rate, holding awareness steady week-over-week.');
}

function generateVideoRecommendations_(current, changes) {
  var recs = [];
  if (changes.views > 30) {
    recs.push('Reach is scaling well, maintain budget and watch view rate as volume grows');
  } else if (changes.views < -15) {
    recs.push('Review budget pacing and refresh creative to recover reach');
  }
  if (current.viewRate < 15) {
    recs.push('Test stronger hooks in the first 5 seconds to lift view rate');
  } else {
    recs.push('Strong view rate, consider expanding to similar audiences to scale reach');
  }
  recs.push('Build remarketing audiences from video engagement for lower-funnel follow-up');
  recs.push('Rotate in fresh creative to limit frequency fatigue');
  return recs.slice(0, 4).map(function (r, i) { return (i + 1) + ') ' + r; }).join('<br>');
}

function generateSummary(clientName, current, changes, profile) {
  profile = profile || 'leadgen';
  let tone = '';
  let details = [];

  if (profile === 'video') {
    if (changes.views > 30 || changes.impressions > 30) {
      tone = 'Strong reach growth this week. ';
    } else if (changes.views < -20 || changes.impressions < -20) {
      tone = 'Reach softened this week. ';
    } else {
      tone = 'Steady awareness performance this week. ';
    }
    details.push('Impressions of <strong>' + formatNumber(current.impressions) + '</strong>');
    details.push('video views of <strong>' + formatNumber(current.views) + '</strong>');
    details.push('clicks of <strong>' + formatNumber(current.clicks) + '</strong>');
    if (Math.abs(changes.ctr) > 10) {
      details.push('CTR ' + (changes.ctr > 0 ? 'improved' : 'declined') + ' to <strong>' + current.ctr.toFixed(2) + '%</strong>');
    }
    details.push('on total investment of <strong>$' + current.spend.toFixed(2) + '</strong>');
  } else {
    if (changes.conversions > 30 || (current.conversions > 0 && changes.cpl < -10)) {
      tone = 'Strong lead generation this week. ';
    } else if (current.conversions === 0) {
      tone = 'No conversions recorded this week. ';
    } else if (changes.conversions < -20) {
      tone = 'Lead volume softened this week. ';
    } else {
      tone = 'Solid lead performance this week. ';
    }
    details.push('Generated <strong>' + formatConversions_(current.conversions) + '</strong> leads');
    if (current.conversions > 0) {
      details.push('at a cost per lead of <strong>$' + current.cpl.toFixed(2) + '</strong>');
    }
    if (Math.abs(changes.clicks) > 10) {
      details.push('clicks ' + (changes.clicks > 0 ? 'rose' : 'fell') + ' <strong>' + Math.abs(changes.clicks).toFixed(0) + '%</strong> to <strong>' + formatNumber(current.clicks) + '</strong>');
    }
    details.push('on total spend of <strong>$' + current.spend.toFixed(2) + '</strong>');
  }

  return tone + details.join(', ') + '.';
}

function generateInsights(current, previous, changes, profile) {
  profile = profile || 'leadgen';
  let html = '';

  // Profile-specific headline insight (leads for lead-gen, reach for video)
  html += buildPrimaryInsight_(profile, current, previous, changes);

  // Video/awareness uses reach-oriented insights only. The search-oriented
  // CTR and CPC warnings below do not apply to a video campaign, where a low
  // CTR is normal and would read as a false problem to the client.
  if (profile === 'video') {
    html += createInsightCard('info', 'Recommendations', generateVideoRecommendations_(current, changes));
    return html;
  }

  // CTR insight
  if (current.ctr >= 10) {
    html += createInsightCard('success', 'Outstanding CTR Performance',
      'A ' + current.ctr.toFixed(2) + '% CTR is exceptional&mdash;approximately 5-7x the industry average. ' +
      (changes.ctr > 5 ? 'The ' + changes.ctr.toFixed(0) + '% improvement week-over-week indicates strong ad relevance.' : 'This indicates highly effective ad copy and targeting.'));
  } else if (current.ctr >= 5) {
    html += createInsightCard('success', 'Strong Click-Through Rate',
      'A ' + current.ctr.toFixed(2) + '% CTR is approximately 2x the industry average, indicating strong ad relevance and effective messaging.');
  } else if (current.ctr < 2) {
    html += createInsightCard('warning', 'CTR Optimization Opportunity',
      'A ' + current.ctr.toFixed(2) + '% CTR is below industry average (2-3%). Recommend testing new ad copy and reviewing keyword relevance.');
  }

  // Efficiency insight
  if (changes.clicks > 50 && changes.cpc < 0) {
    html += createInsightCard('success', 'Exceptional Scale Achievement',
      'Clicks increased ' + changes.clicks.toFixed(0) + '% while CPC decreased ' + Math.abs(changes.cpc).toFixed(0) + '%. The campaign successfully scaled with improved efficiency.');
  } else if (changes.cpc < -15) {
    html += createInsightCard('success', 'Improved Cost Efficiency',
      'CPC dropped ' + Math.abs(changes.cpc).toFixed(0) + '% from $' + previous.cpc.toFixed(2) + ' to $' + current.cpc.toFixed(2) + '. This demonstrates excellent optimization results.');
  } else if (changes.clicks < -15) {
    html += createInsightCard('warning', 'Traffic Volume Decline',
      'Clicks decreased ' + Math.abs(changes.clicks).toFixed(0) + '% week-over-week. This may be due to seasonal factors, competitive pressure, or budget pacing. Recommend reviewing search impression share.');
  }

  // Recommendations
  html += createInsightCard('info', 'Recommendations', generateRecommendations(current, changes));

  return html;
}

function generateRecommendations(current, changes) {
  let recs = [];

  if (changes.clicks > 30) {
    recs.push('Continue current strategy&mdash;the scaling approach is working well');
    recs.push('Monitor CTR trends as volume increases to ensure quality');
  } else if (changes.clicks < -15) {
    recs.push('Review search impression share to identify if budget or rank is limiting visibility');
    recs.push('Analyze search terms report for new keyword opportunities');
  }

  if (changes.ctr < -10) {
    recs.push('Test new ad copy variations to improve click-through rate');
  }

  if (changes.cpc > 15) {
    recs.push('Review bid strategy and quality scores to improve efficiency');
  } else if (changes.cpc < -10) {
    recs.push('Consider reinvesting cost savings to expand reach');
  }

  if (recs.length < 2) {
    recs.push('Monitor competitive landscape for opportunities');
    recs.push('Test similar audiences to scale while maintaining efficiency');
  }

  return recs.slice(0, 4).map((r, i) => (i + 1) + ') ' + r).join('<br>');
}

function createInsightCard(type, title, text) {
  const iconChar = type === 'success' ? '&#10003;' : (type === 'warning' ? '!' : '&rarr;');
  return `
    <div class="insight-card">
      <div class="insight-header">
        <div class="insight-icon ${type}">${iconChar}</div>
        <div class="insight-title">${title}</div>
      </div>
      <p class="insight-text">${text}</p>
    </div>
  `;
}

// ============================================
// FORMATTING HELPERS
// ============================================

function formatCurrency(value) {
  if (value >= 1000) {
    return '$' + Math.round(value).toLocaleString();
  }
  return '$' + value.toFixed(2);
}

function formatNumber(value) {
  return Math.round(value).toLocaleString();
}

function formatChangeText(change, suffix, invert) {
  suffix = suffix || ' vs last week';
  if (Math.abs(change) < 0.5) return 'No change';
  const sign = change > 0 ? '+' : '';
  return sign + change.toFixed(1) + '%' + suffix;
}

function formatChangePct(change) {
  const sign = change > 0 ? '+' : '';
  return sign + change.toFixed(2) + '%';
}

function getChangeClass(change, invert) {
  if (Math.abs(change) < 0.5) return 'neutral';
  const isPositive = invert ? change < 0 : change > 0;
  return isPositive ? 'positive' : 'negative';
}

function getTableClass(change, invert) {
  if (Math.abs(change) < 0.5) return '';
  const isPositive = invert ? change < 0 : change > 0;
  return isPositive ? 'change-positive' : 'change-negative';
}

function getCtrBadge(ctr) {
  if (ctr >= 10) return '<span class="performance-badge excellent">Excellent</span>';
  if (ctr >= 5) return '<span class="performance-badge good">Strong</span>';
  return '';
}

function getCpcBadge(cpc) {
  if (cpc < 0.20) return '<span class="performance-badge excellent">Efficient</span>';
  if (cpc < 0.50) return '<span class="performance-badge good">Good</span>';
  return '';
}

function getClicksBadge(change) {
  if (change > 50) return '<span class="performance-badge excellent">+' + Math.round(change) + '%</span>';
  if (Math.abs(change) < 5) return '<span class="performance-badge good">Stable</span>';
  return '';
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ============================================
// GITHUB DEPLOYMENT
// ============================================

function deployToGitHub(config, folderName, htmlContent) {
  const token = config.github_token;
  const repo = config.github_repo;
  const path = folderName + '/index.html';

  const url = 'https://api.github.com/repos/' + repo + '/contents/' + path;

  // Check if file exists (to get SHA for update)
  let sha = null;
  try {
    const existingResponse = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'token ' + token,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });

    if (existingResponse.getResponseCode() === 200) {
      sha = JSON.parse(existingResponse.getContentText()).sha;
    }
  } catch (e) {
    // File doesn't exist, that's fine
  }

  // Create or update file
  const payload = {
    message: 'Update ' + folderName + ' report',
    content: Utilities.base64Encode(htmlContent),
    branch: 'main'
  };

  if (sha) {
    payload.sha = sha;
  }

  const response = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });

  Logger.log('Deployed ' + folderName + ': ' + response.getResponseCode());
}

function updateIndexHtml(config, clients, dateRange, folderSuffix) {
  const token = config.github_token;
  const repo = config.github_repo;

  // Generate new index content
  const indexHtml = generateIndexHtml(clients, dateRange.thisWeek, folderSuffix);

  // Deploy
  const url = 'https://api.github.com/repos/' + repo + '/contents/index.html';

  // Get existing SHA
  const existingResponse = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json'
    }
  });

  const sha = JSON.parse(existingResponse.getContentText()).sha;

  const response = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      message: 'Update index with new reports',
      content: Utilities.base64Encode(indexHtml),
      sha: sha,
      branch: 'main'
    })
  });

  Logger.log('Updated index.html: ' + response.getResponseCode());
}

function generateIndexHtml(clients, dateRange, folderSuffix) {
  let sections = '';

  clients.forEach(client => {
    if (!client.active) return;

    sections += `
        <!-- ${client.name} -->
        <section class="client-section">
            <div class="section-header">
                <span>${client.name}</span>
                <span class="client-badge">Google Ads</span>
            </div>
            <div class="report-card">
                <a href="/${client.slug}-${folderSuffix}/">
                    <div class="report-info">
                        <div class="report-title">
                            Weekly Performance Report
                            <span class="badge-new">New</span>
                        </div>
                        <div class="report-date">${dateRange}</div>
                    </div>
                    <span class="arrow-icon">&rarr;</span>
                </a>
            </div>
        </section>`;
  });

  return getIndexTemplate().replace('{client_sections}', sections);
}

// ============================================
// EMAIL NOTIFICATION
// ============================================

function sendEmailNotification(config, reportUrls, dateRange) {
  const to = config.email_to;
  const bcc = config.email_bcc;

  let body = 'The weekly performance reports for ' + dateRange + ' are ready:\n\n';
  body += 'Reports Portal: https://reports.roberthebertmedia.com/\n\n';
  body += 'Direct Links:\n';

  reportUrls.forEach(report => {
    body += '• ' + report.name + ': ' + report.url + '\n';
  });

  body += '\nLet me know if you need any changes.\n\nBrandon';

  GmailApp.sendEmail(to, 'Weekly Google Ads Reports Ready - ' + dateRange, body, {
    bcc: bcc,
    name: 'Brandon Hendricks'
  });

  Logger.log('Email sent to: ' + to);
}

// ============================================
// TEMPLATES
// ============================================

function getReportTemplate() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Weekly Performance Report | {client_name} | {date_range}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #0066CC;
            --success: #059669;
            --warning: #D97706;
            --danger: #DC2626;
            --gray-50: #F9FAFB;
            --gray-100: #F3F4F6;
            --gray-200: #E5E7EB;
            --gray-400: #9CA3AF;
            --gray-500: #6B7280;
            --gray-600: #4B5563;
            --gray-700: #374151;
            --gray-800: #1F2937;
            --gray-900: #111827;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #fff;
            color: var(--gray-800);
            line-height: 1.5;
            font-size: 14px;
        }
        .report-container { max-width: 1100px; margin: 0 auto; padding: 40px; }
        .report-header { border-bottom: 3px solid var(--primary); padding-bottom: 24px; margin-bottom: 32px; }
        .header-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; }
        .brand { font-size: 12px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 1.5px; }
        .report-date { font-size: 12px; color: var(--gray-500); text-align: right; }
        .client-name { font-size: 32px; font-weight: 700; color: var(--gray-900); margin-bottom: 4px; }
        .report-title { font-size: 18px; font-weight: 400; color: var(--gray-600); }
        .report-period { display: inline-block; background: var(--primary); color: white; padding: 6px 16px; border-radius: 4px; font-size: 13px; font-weight: 500; margin-top: 12px; }
        .executive-summary { background: var(--gray-50); border-left: 4px solid var(--primary); padding: 24px 28px; margin-bottom: 40px; }
        .summary-title { font-size: 11px; font-weight: 600; color: var(--primary); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px; }
        .summary-text { font-size: 16px; color: var(--gray-700); line-height: 1.7; }
        .summary-text strong { color: var(--gray-900); font-weight: 600; }
        .kpi-section { margin-bottom: 48px; }
        .section-header { font-size: 11px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 20px; padding-bottom: 8px; border-bottom: 1px solid var(--gray-200); }
        .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; }
        .kpi-card { background: white; border: 1px solid var(--gray-200); border-radius: 8px; padding: 24px; position: relative; }
        .kpi-card.highlight { border-color: var(--primary); border-width: 2px; }
        .kpi-label { font-size: 12px; font-weight: 500; color: var(--gray-500); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .kpi-value { font-size: 36px; font-weight: 700; color: var(--gray-900); line-height: 1; margin-bottom: 8px; }
        .kpi-card.highlight .kpi-value { color: var(--primary); }
        .kpi-change { font-size: 13px; font-weight: 500; display: flex; align-items: center; gap: 4px; }
        .kpi-change.positive { color: var(--success); }
        .kpi-change.negative { color: var(--danger); }
        .kpi-change.neutral { color: var(--gray-400); }
        .kpi-subtitle { font-size: 12px; color: var(--gray-500); margin-top: 4px; }
        .performance-badge { position: absolute; top: 16px; right: 16px; padding: 4px 10px; border-radius: 4px; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .performance-badge.excellent { background: #D1FAE5; color: #065F46; }
        .performance-badge.good { background: #DBEAFE; color: #1E40AF; }
        .table-section { margin-bottom: 48px; }
        .data-table { width: 100%; border-collapse: collapse; background: white; border: 1px solid var(--gray-200); border-radius: 8px; overflow: hidden; }
        .data-table thead { background: var(--gray-50); }
        .data-table th { padding: 14px 16px; text-align: left; font-size: 11px; font-weight: 600; color: var(--gray-600); text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid var(--gray-200); }
        .data-table th:not(:first-child) { text-align: right; }
        .data-table td { padding: 16px; border-bottom: 1px solid var(--gray-100); font-size: 14px; }
        .data-table td:not(:first-child) { text-align: right; font-variant-numeric: tabular-nums; }
        .data-table tbody tr:hover { background: var(--gray-50); }
        .data-table tbody tr:last-child td { border-bottom: none; }
        .metric-name { font-weight: 500; color: var(--gray-800); }
        .metric-value { font-weight: 600; color: var(--gray-900); }
        .change-positive { color: var(--success); font-weight: 500; }
        .change-negative { color: var(--danger); font-weight: 500; }
        .insights-section { margin-bottom: 48px; }
        .insight-card { background: white; border: 1px solid var(--gray-200); border-radius: 8px; padding: 24px; margin-bottom: 16px; }
        .insight-card:last-child { margin-bottom: 0; }
        .insight-header { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
        .insight-icon { width: 32px; height: 32px; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 16px; }
        .insight-icon.success { background: #D1FAE5; }
        .insight-icon.info { background: #DBEAFE; }
        .insight-icon.warning { background: #FEF3C7; }
        .insight-title { font-size: 14px; font-weight: 600; color: var(--gray-800); }
        .insight-text { font-size: 14px; color: var(--gray-600); line-height: 1.6; }
        .report-footer { border-top: 1px solid var(--gray-200); padding-top: 24px; margin-top: 48px; display: flex; justify-content: space-between; align-items: center; }
        .footer-brand { font-size: 12px; color: var(--gray-500); }
        .footer-brand a { color: var(--primary); text-decoration: none; font-weight: 500; }
        .footer-meta { font-size: 11px; color: var(--gray-400); }
        @media (max-width: 768px) { .report-container { padding: 20px; } .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 16px; } .kpi-value { font-size: 28px; } .header-top { flex-direction: column; gap: 8px; } }
    </style>
</head>
<body>
    <div class="report-container">
        <header class="report-header">
            <div class="header-top">
                <div class="brand">Robert Hebert Media</div>
                <div class="report-date">Report Generated: {generated_date}<br>Confidential</div>
            </div>
            <h1 class="client-name">{client_name}</h1>
            <p class="report-title">Weekly Google Ads Performance Report</p>
            <span class="report-period">{date_range}</span>
        </header>

        <div class="executive-summary">
            <div class="summary-title">Executive Summary</div>
            <p class="summary-text">{executive_summary}</p>
        </div>

        <section class="kpi-section">
            <div class="section-header">Key Performance Indicators</div>
            <div class="kpi-grid">
                {kpi_cards_html}
            </div>
        </section>

        <section class="table-section">
            <div class="section-header">Detailed Metrics</div>
            <table class="data-table">
                <thead>
                    <tr><th>Metric</th><th>This Week</th><th>Previous Week</th><th>Change</th></tr>
                </thead>
                <tbody>
                    {metric_rows_html}
                </tbody>
            </table>
        </section>

        <section class="insights-section">
            <div class="section-header">Key Insights & Recommendations</div>
            {insights_html}
        </section>

        <footer class="report-footer">
            <div class="footer-brand">Prepared by <a href="https://roberthebertmedia.com">Robert Hebert Media</a></div>
            <div class="footer-meta">Report ID: {report_id} | Page 1 of 1</div>
        </footer>
    </div>
</body>
</html>`;
}

function getIndexTemplate() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Client Reports | Robert Hebert Media</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #0066CC;
            --success: #059669;
            --gray-50: #F9FAFB;
            --gray-100: #F3F4F6;
            --gray-200: #E5E7EB;
            --gray-400: #9CA3AF;
            --gray-500: #6B7280;
            --gray-600: #4B5563;
            --gray-700: #374151;
            --gray-800: #1F2937;
            --gray-900: #111827;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #fff;
            color: var(--gray-800);
            line-height: 1.5;
            font-size: 14px;
        }
        .page-container { max-width: 900px; margin: 0 auto; padding: 40px; }
        .page-header { border-bottom: 3px solid var(--primary); padding-bottom: 24px; margin-bottom: 40px; }
        .header-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; }
        .brand { font-size: 12px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 1.5px; }
        .header-meta { font-size: 12px; color: var(--gray-500); text-align: right; }
        .page-title { font-size: 32px; font-weight: 700; color: var(--gray-900); margin-bottom: 4px; }
        .page-subtitle { font-size: 18px; font-weight: 400; color: var(--gray-600); }
        .client-section { margin-bottom: 32px; }
        .section-header { font-size: 11px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 1px solid var(--gray-200); display: flex; align-items: center; gap: 12px; }
        .client-badge { background: var(--primary); color: white; padding: 3px 10px; border-radius: 4px; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .report-card { background: white; border: 1px solid var(--gray-200); border-radius: 8px; padding: 20px 24px; margin-bottom: 12px; transition: all 0.2s ease; position: relative; }
        .report-card:hover { border-color: var(--primary); box-shadow: 0 4px 12px rgba(0, 102, 204, 0.1); transform: translateY(-2px); }
        .report-card:last-child { margin-bottom: 0; }
        .report-card a { text-decoration: none; color: inherit; display: flex; justify-content: space-between; align-items: center; }
        .report-info { flex: 1; }
        .report-title { font-size: 15px; font-weight: 600; color: var(--gray-900); margin-bottom: 4px; display: flex; align-items: center; gap: 10px; }
        .report-date { font-size: 13px; color: var(--gray-500); }
        .badge-new { display: inline-block; background: #D1FAE5; color: #065F46; padding: 2px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .arrow-icon { color: var(--gray-400); font-size: 18px; transition: transform 0.2s ease, color 0.2s ease; }
        .report-card:hover .arrow-icon { color: var(--primary); transform: translateX(4px); }
        .page-footer { border-top: 1px solid var(--gray-200); padding-top: 24px; margin-top: 48px; display: flex; justify-content: space-between; align-items: center; }
        .footer-brand { font-size: 12px; color: var(--gray-500); }
        .footer-brand a { color: var(--primary); text-decoration: none; font-weight: 500; }
        .footer-brand a:hover { text-decoration: underline; }
        .footer-meta { font-size: 11px; color: var(--gray-400); }
        @media (max-width: 768px) { .page-container { padding: 20px; } .header-top { flex-direction: column; gap: 8px; } .page-title { font-size: 24px; } .report-card a { flex-direction: column; align-items: flex-start; gap: 8px; } .arrow-icon { display: none; } }
    </style>
</head>
<body>
    <div class="page-container">
        <header class="page-header">
            <div class="header-top">
                <div class="brand">Robert Hebert Media</div>
                <div class="header-meta">Client Reports Portal</div>
            </div>
            <h1 class="page-title">Performance Reports</h1>
            <p class="page-subtitle">Weekly Google Ads Analytics</p>
        </header>

        {client_sections}

        <footer class="page-footer">
            <div class="footer-brand">Powered by <a href="https://roberthebertmedia.com">Robert Hebert Media</a></div>
            <div class="footer-meta">Confidential Client Portal</div>
        </footer>
    </div>
</body>
</html>`;
}

// ============================================
// MENU FUNCTIONS
// ============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RHM Reports')
    .addItem('Generate Reports Now', 'generateWeeklyReports')
    .addItem('Turn On Weekly Automation', 'enableWeeklyAutomation')
    .addItem('Turn Off Weekly Automation', 'disableWeeklyAutomation')
    .addItem('Setup / Reset Tabs', 'setupSpreadsheet')
    .addItem('Test Email', 'testEmail')
    .addItem('Force Reauth', 'forceReauth')
    .addToUi();
}

// ============================================
// SPREADSHEET SETUP
// ============================================
// Run once after creating a new spreadsheet. Creates every tab the report
// generator reads from, with headers and reasonable defaults pre-filled.

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheet_(ss, 'Config', _configTabData_());
  _ensureSheet_(ss, 'Clients', _clientsTabData_());
  _ensureSheet_(ss, 'Date Range', _dateRangeTabData_());
  _ensureSheet_(ss, 'This Week', _weekTabData_());
  _ensureSheet_(ss, 'Previous Week', _weekTabData_());

  // Remove the default empty Sheet1 if it's still hanging around.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  _notify_(
    'RHM Report tabs ready.\n\n' +
    'Next steps:\n' +
    '1. Fill the Config tab with your GitHub PAT (repo scope) and recipient emails.\n' +
    '2. Add real Google Ads customer IDs to the Clients tab.\n' +
    '3. Install the Google Ads export script (MCC level) and schedule it Sunday 11pm. ' +
    'It writes This Week / Previous Week / Date Range into THIS sheet automatically. ' +
    '(Manual paste still works as a fallback.)\n' +
    '4. Run "Force Reauth" once to grant OAuth scopes.\n' +
    '5. Run "Turn On Weekly Automation" once to install the Monday 8am trigger.'
  );
}

// Show a UI alert when running from the spreadsheet menu; fall back to
// Logger when running from the editor (where getUi() is unavailable).
function _notify_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

function _ensureSheet_(ss, name, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
  for (let c = 1; c <= rows[0].length; c++) sheet.autoResizeColumn(c);
}

function _configTabData_() {
  return [
    ['Key', 'Value'],
    ['github_token', '<<paste GitHub PAT with repo scope>>'],
    ['github_repo', 'BLincoln711/robert-hebert-media-reports'],
    // Reports email Brandon during review. Flip email_to to robert@roberthebertmedia.com once approved.
    ['email_to', 'brandon@hendricks.ai'],
    ['email_bcc', '']
  ];
}

function _clientsTabData_() {
  return [
    ['Name', 'Slug', 'Customer ID', 'Active'],
    ['JFTx2025', 'jftx2025', '', true],
    ['PFBHNC', 'pfbhnc', '', true],
    ['ReOptica', 'reoptica', '', true]
  ];
}

function _dateRangeTabData_() {
  // Default to the most recently completed Sunday-Saturday week.
  const today = new Date();
  const dow = today.getDay(); // 0 Sun ... 6 Sat
  const lastSaturday = new Date(today);
  lastSaturday.setDate(today.getDate() - (dow === 6 ? 7 : dow + 1));
  const lastSunday = new Date(lastSaturday);
  lastSunday.setDate(lastSaturday.getDate() - 6);
  const prevSaturday = new Date(lastSunday);
  prevSaturday.setDate(lastSunday.getDate() - 1);
  const prevSunday = new Date(prevSaturday);
  prevSunday.setDate(prevSaturday.getDate() - 6);

  return [
    ['Key', 'Value'],
    ['this_week_start', lastSunday],
    ['this_week_end', lastSaturday],
    ['prev_week_start', prevSunday],
    ['prev_week_end', prevSaturday]
  ];
}

function _weekTabData_() {
  return [
    ['Client', 'Spend', 'Impressions', 'Clicks', 'Conversions', 'Views'],
    ['JFTx2025', 0, 0, 0, 0, 0],
    ['PFBHNC', 0, 0, 0, 0, 0],
    ['ReOptica', 0, 0, 0, 0, 0]
  ];
}

function testEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss);

  const recipient = config.email_bcc || config.email_to || Session.getEffectiveUser().getEmail();
  GmailApp.sendEmail(recipient, 'Test - RHM Report System', 'This is a test email from the RHM Report Generator.', {
    name: 'RHM Report System'
  });

  _notify_('Test email sent to ' + recipient);
}
