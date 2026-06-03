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

clients.forEach((client) => {
  const cur = Object.assign({}, thisWeek[client.name]);
  const prev = Object.assign({}, prevWeek[client.name]);
  const html = sandbox.generateReportHtml(client, cur, prev, dateRange);
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
