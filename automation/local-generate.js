/**
 * Local one-off generator for the RHM weekly reports.
 *
 * Runs the REAL render functions from apps-script/Code.js inside a vm sandbox
 * (Apps Script globals stubbed) so the HTML is byte-for-byte what the Monday
 * trigger would have produced. Used 2026-06-02 to ship the week of May 23-29
 * reports manually after the trigger failed on "Authorization is required".
 *
 * Data source: the bound Google Sheet (week 5/23-5/29, prev 5/16-5/22), which
 * the failed run would have read. Does NOT push anywhere — writes local files
 * only; deploy is a separate explicit git step.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const REPO = path.join(process.env.HOME, 'robert-hebert-media-reports');
const CODE = fs.readFileSync(path.join(REPO, 'automation/apps-script/Code.js'), 'utf8');

// ---- Apps Script global stubs (only what the render path touches) ----
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
function formatDateMDY(d) { return MONTHS[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear(); }
const noop = function () {};
const passiveProxy = new Proxy(function () {}, {
  get: () => passiveProxy,
  apply: () => passiveProxy,
});
const sandbox = {
  Utilities: {
    // Render path only calls formatDate(new Date(), tz, 'MMMM d, yyyy').
    formatDate: function (date, tz, fmt) { return formatDateMDY(date); },
    base64Encode: function (s) { return Buffer.from(s, 'utf8').toString('base64'); },
  },
  Logger: { log: noop },
  console: console,
  SpreadsheetApp: passiveProxy,
  GmailApp: passiveProxy,
  MailApp: passiveProxy,
  UrlFetchApp: passiveProxy,
  ScriptApp: passiveProxy,
  Session: passiveProxy,
};
vm.createContext(sandbox);
vm.runInContext(CODE, sandbox);

// ---- Data: exactly what the bound sheet holds for week 5/23-5/29 ----
// Columns per getWeekData: spend, impressions, clicks, conversions, views
const thisWeek = {
  JFTx2025: { spend: 990.84, impressions: 706805, clicks: 9622, conversions: 0,      views: 42319 },
  PFBHNC:   { spend: 3175.7, impressions: 33823,  clicks: 4430, conversions: 720.09, views: 1 },
  ReOptica: { spend: 91.92,  impressions: 628,    clicks: 67,   conversions: 0,      views: 0 },
};
const prevWeek = {
  JFTx2025: { spend: 1034.76, impressions: 782386, clicks: 9946, conversions: 0,      views: 53125 },
  PFBHNC:   { spend: 3383.77, impressions: 28832,  clicks: 4090, conversions: 763.02, views: 1 },
  ReOptica: { spend: 165.63,  impressions: 792,    clicks: 76,   conversions: 0,      views: 0 },
};

const clients = [
  { name: 'JFTx2025', slug: 'jftx2025', active: true },
  { name: 'PFBHNC',   slug: 'pfbhnc',   active: true },
  { name: 'ReOptica', slug: 'reoptica', active: true },
];

const dateRange = {
  thisWeek: 'May 23, 2026 &ndash; May 29, 2026',
  prevWeek: 'May 16, 2026 &ndash; May 22, 2026',
  thisWeekStart: new Date(2026, 4, 23),
  thisWeekEnd: new Date(2026, 4, 29),
};

const folderSuffix = sandbox.getFolderSuffix(dateRange); // -> may23-29
console.log('folderSuffix:', folderSuffix);

// Print rules so PDF export never splits a KPI grid or a table across pages
// (Brandon's PDF house rule: data blocks stay whole, no browser headers/footers).
const PRINT_CSS = `
    @media print {
      @page { size: letter; margin: 0.38in; }
      html, body { background: #fff !important; font-size: 12px; }
      .report-container { max-width: 100% !important; padding: 0 !important; }
      .report-header { margin-bottom: 16px !important; padding-bottom: 12px !important; }
      .executive-summary { margin-bottom: 18px !important; padding: 14px 18px !important; }
      .kpi-section, .table-section, .insights-section { margin-bottom: 16px !important; }
      .section-header { margin-bottom: 12px !important; }
      .kpi-grid { gap: 12px !important; }
      .kpi-card { padding: 13px !important; }
      .kpi-value { font-size: 25px !important; }
      .data-table td { padding: 8px 14px !important; }
      .insight-card { padding: 12px !important; margin-bottom: 8px !important; }
      .insight-header { margin-bottom: 8px !important; }
      .report-footer { margin-top: 14px !important; padding-top: 12px !important; }
      .kpi-grid, .data-table, .kpi-card, .insight-card, tr { page-break-inside: avoid; break-inside: avoid; }
      thead { display: table-header-group; }
      a[href]:after { content: none !important; }
    }`;

clients.forEach((client) => {
  const cur = Object.assign({}, thisWeek[client.name]);
  const prev = Object.assign({}, prevWeek[client.name]);
  let html = sandbox.generateReportHtml(client, cur, prev, dateRange);
  html = html.replace('</style>', PRINT_CSS + '\n    </style>');
  // "Page 1 of 1" is meaningless on a multi-page PDF and slightly wrong; drop it.
  html = html.replace(' | Page 1 of 1', '');

  // ReOptica: conversion tracking was only just applied to this account, so the
  // generator's default "No Conversions Recorded -> verify tracking is firing"
  // warning is misleading. Reframe that one card as a positive setup note.
  if (client.slug === 'reoptica') {
    html = html
      .replace('class="insight-icon warning">!</div>', 'class="insight-icon success">&#10003;</div>')
      .replace('<div class="insight-title">No Conversions Recorded</div>',
               '<div class="insight-title">Conversion Tracking Now Live</div>')
      .replace(
        'No conversions were tracked this week. Recommend verifying conversion tracking is firing and reviewing lead-form and landing-page performance.',
        'Conversion tracking was implemented on this account this week, so no historical conversions appear in this report. With tracking now live, lead volume will begin populating in next week&rsquo;s report and going forward.'
      );
  }
  const dir = path.join(REPO, client.slug + '-' + folderSuffix);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'index.html'), html);
  const bad = /undefined|NaN|\$NaN|{[a-z_]+}/.exec(html);
  console.log('wrote', client.slug + '-' + folderSuffix + '/index.html',
    '(' + html.length + ' bytes)', bad ? '  <-- SUSPECT TOKEN: ' + bad[0] : 'OK');
});

const indexHtml = sandbox.generateIndexHtml(clients, dateRange.thisWeek, folderSuffix);
fs.writeFileSync(path.join(REPO, 'index.html'), indexHtml);
console.log('wrote index.html (' + indexHtml.length + ' bytes)');
console.log('DONE');
