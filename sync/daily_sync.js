// LPH Performance Tracker — canonical daily ServiceTitan -> GitHub sync script.
//
// This is the SINGLE SOURCE OF TRUTH for parsing the ServiceTitan "Daily Job
// Detail" report and merging it into data/dd.json, data/td.json and
// data/install.json in mattsbaker1980-dev/LPH-Performance-Tracker.
//
// WHY THIS EXISTS: the scheduled sync tasks used to be instructed to re-fetch
// index.html every single run and re-derive (re-type, from scratch, in a
// fresh LLM context) the parsing/merge logic (NAME_MAP, classifyBU,
// resolveTech, the DD/TD/INST schemas, the merge routines) before doing
// anything else. That step was slow, token-heavy, and is the prime suspect
// for why both the primary and watchdog runs stalled and never completed on
// 2026-09-09/2026-09-10. This script removes that step entirely: the logic
// lives here, version-controlled, and a scheduled run just executes it.
//
// USAGE:
//   node daily_sync.js <path-to-report-dump.txt> [run_context]
//
//   <path-to-report-dump.txt> — plain text produced by reading the
//     ServiceTitan .xlsx attachment via the Microsoft 365 connector's
//     read_resource tool. IMPORTANT: when that attachment is large, the tool
//     does NOT return its content inline — it saves the content to a local
//     .txt file and reports that file's path in its response text (this is
//     normal/expected, not an error). Pass THAT saved file's path directly
//     as this argument. Do NOT read the file into your own context first —
//     this script reads it from disk itself. There is nothing to transcribe.
//
//   [run_context] — optional label written into data/_sync_log.json, e.g.
//     "scheduled" (primary) or "watchdog". Defaults to "scheduled".
//
// GITHUB TOKEN: reads from the GH_TOKEN environment variable if set,
// otherwise falls back to reading /root/work/.ghtoken (a file containing
// just the token, no trailing content expected beyond whitespace).
//
// This script fetches+applies its own merge logic; it does NOT need
// index.html or any other page re-read/re-derived at runtime.

const fs = require('fs');
const path = require('path');

const OWNER = 'mattsbaker1980-dev';
const REPO = 'LPH-Performance-Tracker';
const BRANCH = 'main';
const API = `https://api.github.com/repos/${OWNER}/${REPO}/contents/`;

function getToken() {
  if (process.env.GH_TOKEN && process.env.GH_TOKEN.trim()) return process.env.GH_TOKEN.trim();
  const fallbackPath = '/root/work/.ghtoken';
  if (fs.existsSync(fallbackPath)) return fs.readFileSync(fallbackPath, 'utf8').trim();
  throw new Error('No GitHub token found. Set GH_TOKEN env var or create /root/work/.ghtoken.');
}
const TOKEN = getToken();

function headers(extra) {
  return Object.assign({ Authorization: 'Bearer ' + TOKEN, 'User-Agent': 'lph-sync', Accept: 'application/vnd.github+json' }, extra || {});
}
async function ghGet(p) {
  const res = await fetch(API + encodeURIComponent(p).replace(/%2F/g, '/') + '?ref=' + BRANCH + '&_cb=' + Date.now(), { headers: headers() });
  if (res.status === 404) return { json: null, sha: null };
  if (!res.ok) throw new Error(`GET ${p} failed: ${res.status} ${await res.text()}`);
  const body = await res.json();
  const content = Buffer.from(body.content, 'base64').toString('utf8');
  return { json: JSON.parse(content), sha: body.sha };
}
async function ghPut(p, obj, sha, message) {
  const content = Buffer.from(JSON.stringify(obj, null, 2)).toString('base64');
  let lastErr;
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = await fetch(API + encodeURIComponent(p).replace(/%2F/g, '/'), {
      method: 'PUT',
      headers: headers({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ message, content, sha: sha || undefined, branch: BRANCH }),
    });
    if (res.ok) return res.json();
    const bodyText = await res.text();
    if (res.status === 409 && attempt < 2) {
      // sha conflict - re-GET and retry
      const fresh = await ghGet(p);
      sha = fresh.sha;
      lastErr = new Error(`PUT ${p} 409, retrying with fresh sha`);
      continue;
    }
    throw new Error(`PUT ${p} failed: ${res.status} ${bodyText}`);
  }
  throw lastErr;
}
async function writeLog(stage, detail, runContext, extra) {
  try {
    const { sha } = await ghGet('data/_sync_log.json');
    const body = Object.assign(
      { time: new Date().toISOString(), stage, detail: detail || '', run_context: runContext || 'scheduled' },
      extra || {}
    );
    await ghPut('data/_sync_log.json', body, sha, `Daily sync (${runContext || 'scheduled'}): ${stage}`);
  } catch (e) {
    console.error('WARNING: failed to write _sync_log.json:', e.message);
  }
}

// ---- Parsing logic (mirrors index.html's Upload Data tool / sync_today.py) ----

const NAME_MAP = {
  'Electrical|tyson|k': 'Tyson Kline', 'Electrical|ryan|k': 'Ryan Kurtz', 'Electrical|brian|m': 'Brian McNamara',
  'Electrical|cole|p': 'Cole Portner', 'Electrical|ryan|s': 'Ryan Santiago', 'Electrical|ryan|w': 'Ryan Walten',
  'HVAC|garrick|b': 'Garrick Budrow', 'HVAC|robert|f': 'Robert Fenton', 'HVAC|andrew|g': 'Andrew Gambler',
  'HVAC|shane|h': 'Shane Hetrich', 'HVAC|cesar|r': 'Cesar Rodriguez', 'HVAC|cayden|w': 'Cayden Warner',
  'HVAC|uriah|w': 'Uriah Warner', 'HVAC|brian|w': 'Brian Wasche', 'HVAC|jeff|w': 'Jeff Watson',
  'HVAC|eric|v': 'Eric Vazquez', 'HVAC|traevon|h': 'Traevon Hinton', 'HVAC|scott|h': 'Scott Huber',
  'HVAC|brian|m': 'Brian McGlynn', 'Plumbing|layne|b': 'Layne Beard', 'Plumbing|derrick|b': 'Derrick Bender',
  'Plumbing|chris|d': 'Chris DeCicco', 'Plumbing|john|d': 'John Dwyer', 'Plumbing|jason|g': 'Jason Gantt',
  'Plumbing|nick|e': 'Nick Elam', 'Plumbing|tylor|h': 'Tylor Hughes', 'Plumbing|austin|j': 'Austin Johnson',
  'Plumbing|mason|k': 'Mason Kephart', 'Plumbing|robert|r': 'Robert Reynolds', 'Plumbing|mike|d': 'Mike DeCicco',
  'Plumbing|ethan|j': 'Ethan Jones', 'Plumbing|kevin|j': 'Kevin Jones', 'Plumbing|reese|g': 'Reese Gassert',
  'Sewer|ethan|j': 'Ethan Jones', 'Sewer|mike|d': 'Mike DeCicco', 'Sewer|layne|b': 'Layne Beard',
  'Sewer|derrick|b': 'Derrick Bender', 'Sewer|chris|d': 'Chris DeCicco', 'Sewer|john|d': 'John Dwyer',
  'Sewer|jason|g': 'Jason Gantt', 'Sewer|nick|e': 'Nick Elam', 'Sewer|tylor|h': 'Tylor Hughes',
  'Sewer|austin|j': 'Austin Johnson', 'Sewer|mason|k': 'Mason Kephart', 'Sewer|kevin|j': 'Kevin Jones',
  'Sewer|robert|r': 'Robert Reynolds', 'Sewer|reese|g': 'Reese Gassert', 'Plumbing|bill|p': 'William Price',
  'Sewer|bill|p': 'William Price', 'Plumbing|josh|r': 'Joshua Rosario', 'Sewer|josh|r': 'Joshua Rosario',
};

const METRIC_FIELDS = ['rev','calls','money_calls','tasks','hours_paid','on_job','sold',
  'demand_rev','demand_calls','demand_money','demand_tasks',
  'spp_rev','spp_calls','spp_money','spp_tasks',
  'cod_rev','cod_calls','cod_money','cod_tasks',
  'warranty','callbacks','service_agreements','spp_missed'];

function emptyMetrics() { const m = {}; METRIC_FIELDS.forEach(f => m[f] = 0); return m; }
function round2(n) { return Math.round((n || 0) * 100) / 100; }
function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

function classifyBU(bu) {
  if (!bu) return [null, null];
  const b = String(bu).toLowerCase();
  let dept = null;
  if (b.includes('sewer')) dept = 'Sewer';
  else if (b.includes('plumbing')) dept = 'Plumbing';
  else if (b.includes('hvac')) dept = 'HVAC';
  else if (b.includes('elec')) dept = 'Electrical';
  let cat = null;
  if (b.includes('cod')) cat = 'COD';
  else if (b.includes('system check') || b.includes('sys ck')) cat = 'SystemCheck';
  else if (b.includes('demand')) cat = 'Demand';
  else if (b.includes('install sales')) cat = 'InstallSales';
  else if (b.includes('installs')) cat = 'Installs';
  else if (b.includes('service')) cat = 'Service';
  if (dept === 'Sewer' && cat === 'Service') cat = 'Demand';
  return [dept, cat];
}
function resolveTech(dept, rawName) {
  if (!rawName) return rawName;
  const parts = String(rawName).replace(/\./g, '').trim().split(/\s+/);
  if (parts.length < 2) return rawName;
  const key = dept + '|' + parts[0].toLowerCase() + '|' + parts[1][0].toLowerCase();
  return NAME_MAP[key] || rawName;
}
function toIsoDateSerial(serial) {
  const n = parseFloat(serial);
  if (!n || n <= 0 || isNaN(n)) return null;
  const base = Date.UTC(1899, 11, 30);
  const d = new Date(base + n * 86400000);
  return d.toISOString().slice(0, 10);
}

function parseDump(dumpPath) {
  const text = fs.readFileSync(dumpPath, 'utf8');
  const lines = text.split('\n');
  let start = null;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].startsWith('Job ID')) { start = i; break; }
  }
  if (start === null) throw new Error("Couldn't find header row (line starting 'Job ID') in dump file");

  const rows = lines.slice(start).filter(l => l.trim() !== '').map(l => l.split('\t'));
  const header = rows[0];

  const techDay = {};
  const deptDay = {};
  const installDay = {};
  let totalReportRevenue = 0, capturedRevenue = 0, excludedRevenue = 0, nRows = 0;

  for (const r of rows.slice(1)) {
    if (!r || !r.length) continue;
    const col = i => (i < r.length ? r[i] : undefined);
    const bu = col(5), status = col(6), opp = col(7), compDateRaw = col(10);
    const warrantyFlag = col(14), recallFlag = col(15), convertedFlag = col(16);
    const primaryTech = col(18);
    const soldHrs = num(col(19)), paidTime = num(col(20)), hoursWorked = num(col(21));
    const revenue = num(col(28));

    nRows++;
    totalReportRevenue += revenue;

    if (!bu || !primaryTech) { excludedRevenue += revenue; continue; }
    const compDate = toIsoDateSerial(compDateRaw);
    if (!compDate) { excludedRevenue += revenue; continue; }
    const [dept, cat] = classifyBU(bu);
    if (!dept || !cat) { excludedRevenue += revenue; continue; }

    capturedRevenue += revenue;
    const tech = resolveTech(dept, primaryTech);
    const isCompleted = status === 'Completed';
    const isMoney = isCompleted && revenue > 0;
    const isWarranty = warrantyFlag === 'TRUE' || warrantyFlag === true;
    const isRecall = recallFlag === 'TRUE' || recallFlag === true;
    const isConverted = convertedFlag === 'TRUE' || convertedFlag === true;
    const isOpportunity = opp === 'TRUE' || opp === true;

    if (cat === 'Installs' || cat === 'InstallSales' || cat === 'Service') {
      const ikey = dept + '|' + tech + '|' + compDate;
      installDay[ikey] = (installDay[ikey] || 0) + revenue;
      continue;
    }

    const tdKey = dept + '|' + tech + '|' + compDate;
    if (!techDay[tdKey]) techDay[tdKey] = { dept, tech, date: compDate, src: 'jd', ...emptyMetrics() };
    const t = techDay[tdKey];
    t.hours_paid = Math.max(t.hours_paid, paidTime);
    t.on_job = Math.max(t.on_job, hoursWorked);
    if (isWarranty) t.warranty += 1;
    if (isRecall) t.callbacks += 1;
    if (cat === 'SystemCheck' && isOpportunity && !isConverted) t.spp_missed += 1;

    const ddKey = dept + '|' + compDate;
    if (!deptDay[ddKey]) deptDay[ddKey] = { dept, date: compDate, ...emptyMetrics() };
    const d = deptDay[ddKey];

    t.rev += revenue; t.sold += soldHrs;
    if (isCompleted) { t.calls += 1; d.calls += 1; }
    if (isMoney) { t.money_calls += 1; d.money_calls += 1; }
    if (cat === 'Demand') {
      t.demand_rev += revenue; d.demand_rev += revenue;
      if (isCompleted) { t.demand_calls += 1; d.demand_calls += 1; }
      if (isMoney) { t.demand_money += 1; d.demand_money += 1; }
    }
    if (cat === 'SystemCheck') {
      t.spp_rev += revenue; d.spp_rev += revenue;
      if (isCompleted) { t.spp_calls += 1; d.spp_calls += 1; }
      if (isMoney) { t.spp_money += 1; d.spp_money += 1; }
    }
    if (cat === 'COD') {
      t.cod_rev += revenue; d.cod_rev += revenue;
      if (isCompleted) { t.cod_calls += 1; d.cod_calls += 1; }
      if (isMoney) { t.cod_money += 1; d.cod_money += 1; }
    }
    d.rev += revenue; d.sold += soldHrs;
    d.warranty += isWarranty ? 1 : 0;
    d.callbacks += isRecall ? 1 : 0;
    if (cat === 'SystemCheck' && isOpportunity && !isConverted) d.spp_missed += 1;
  }

  // roll up hours_paid/on_job into dept_day from tech_day
  for (const t of Object.values(techDay)) {
    const ddKey = t.dept + '|' + t.date;
    if (!deptDay[ddKey]) deptDay[ddKey] = { dept: t.dept, date: t.date, ...emptyMetrics() };
    deptDay[ddKey].hours_paid += t.hours_paid;
    deptDay[ddKey].on_job += t.on_job;
  }

  const installRows = Object.entries(installDay).map(([k, rev]) => {
    const [dept, tech, dt] = k.split('|');
    return [dept, tech, dt, round2(rev)];
  });

  const reconcileOk = Math.abs(capturedRevenue + excludedRevenue - totalReportRevenue) < 0.01;
  const datesSeen = [...new Set(Object.values(deptDay).map(v => v.date))].sort();

  return {
    techDay: Object.values(techDay), deptDay: Object.values(deptDay), installRows,
    nRows, totalReportRevenue: round2(totalReportRevenue), capturedRevenue: round2(capturedRevenue),
    excludedRevenue: round2(excludedRevenue), reconcileOk, datesSeen,
  };
}

// ---- Merge logic (same key scheme as merge_and_push.js) ----

const DD_DEPT = 0, DD_DATE = 1;
const TD_DEPT = 0, TD_TECH = 1, TD_DATE = 2;
const INST_DEPT = 0, INST_TECH = 1, INST_DATE = 2;

function rowFromMetrics(dept, tech, date, src, m) {
  const row = [dept, tech, date, src];
  METRIC_FIELDS.forEach(f => row.push(round2(m[f] || 0)));
  return row;
}
function ddRowFromMetrics(dept, date, m) {
  const row = [dept, date];
  METRIC_FIELDS.forEach(f => row.push(round2(m[f] || 0)));
  return row;
}

async function mergeAndPush(parsed, runContext) {
  const summary = {};

  {
    const { json, sha } = await ghGet('data/dd.json');
    const dd = json && json.dd ? json.dd : [];
    const before = dd.length;
    const idx = {}; dd.forEach((r, i) => { idx[r[DD_DEPT] + '|' + r[DD_DATE]] = i; });
    let updated = 0, added = 0;
    parsed.deptDay.forEach(nr => {
      const key = nr.dept + '|' + nr.date;
      const row = ddRowFromMetrics(nr.dept, nr.date, nr);
      if (idx[key] !== undefined) { dd[idx[key]] = row; updated++; } else { dd.push(row); idx[key] = dd.length - 1; added++; }
    });
    const merged = Object.assign({}, json || {}, { dd });
    await ghPut('data/dd.json', merged, sha, `Daily sync (${runContext}): update dd.json`);
    summary.dd = { before, after: dd.length, updated, added };
  }
  await writeLog('wrote_dd', JSON.stringify(summary.dd), runContext);

  {
    const { json, sha } = await ghGet('data/td.json');
    const td = json && json.td ? json.td : [];
    const before = td.length;
    const idx = {}; td.forEach((r, i) => { idx[r[TD_DEPT] + '|' + r[TD_TECH] + '|' + r[TD_DATE]] = i; });
    let updated = 0, added = 0;
    parsed.techDay.forEach(nr => {
      const key = nr.dept + '|' + nr.tech + '|' + nr.date;
      const row = rowFromMetrics(nr.dept, nr.tech, nr.date, nr.src, nr);
      if (idx[key] !== undefined) { td[idx[key]] = row; updated++; } else { td.push(row); idx[key] = td.length - 1; added++; }
    });
    const merged = Object.assign({}, json || {}, { td });
    await ghPut('data/td.json', merged, sha, `Daily sync (${runContext}): update td.json`);
    summary.td = { before, after: td.length, updated, added };
  }
  await writeLog('wrote_td', JSON.stringify(summary.td), runContext);

  {
    const { json, sha } = await ghGet('data/install.json');
    const install = json && json.install ? json.install : [];
    const before = install.length;
    const idx = {}; install.forEach((r, i) => { idx[r[INST_DEPT] + '|' + r[INST_TECH] + '|' + r[INST_DATE]] = i; });
    let updated = 0, added = 0;
    parsed.installRows.forEach(nr => {
      const key = nr[INST_DEPT] + '|' + nr[INST_TECH] + '|' + nr[INST_DATE];
      if (idx[key] !== undefined) { install[idx[key]] = nr; updated++; } else { install.push(nr); idx[key] = install.length - 1; added++; }
    });
    const merged = Object.assign({}, json || {}, { install });
    await ghPut('data/install.json', merged, sha, `Daily sync (${runContext}): update install.json`);
    summary.install = { before, after: install.length, updated, added };
  }
  await writeLog('wrote_install', JSON.stringify(summary.install), runContext);

  return summary;
}

async function main() {
  const dumpPath = process.argv[2];
  const runContext = process.argv[3] || 'scheduled';
  if (!dumpPath) {
    console.error('Usage: node daily_sync.js <path-to-report-dump.txt> [run_context]');
    process.exit(1);
  }
  if (!fs.existsSync(dumpPath)) {
    console.error(`Dump file not found: ${dumpPath}`);
    process.exit(1);
  }

  await writeLog('started', `Parsing ${dumpPath}`, runContext);

  let parsed;
  try {
    parsed = parseDump(dumpPath);
  } catch (e) {
    await writeLog('failed', `Parse error: ${e.message}`, runContext);
    console.error('PARSE FAILED:', e);
    process.exit(1);
  }

  console.log(`Parsed ${parsed.nRows} data rows.`);
  console.log(`Total report revenue: ${parsed.totalReportRevenue}`);
  console.log(`Captured revenue: ${parsed.capturedRevenue}, Excluded: ${parsed.excludedRevenue}`);
  console.log(`Reconcile OK: ${parsed.reconcileOk}`);
  console.log(`Dates covered: ${parsed.datesSeen.join(', ')}`);
  console.log(`techDay rows: ${parsed.techDay.length}, deptDay rows: ${parsed.deptDay.length}, installRows: ${parsed.installRows.length}`);

  if (!parsed.reconcileOk) {
    await writeLog('reconciled_mismatch', `captured=${parsed.capturedRevenue} excluded=${parsed.excludedRevenue} total=${parsed.totalReportRevenue}`, runContext);
    console.error('RECONCILIATION MISMATCH — stopping before writing any data. Report this, do not guess.');
    process.exit(1);
  }
  await writeLog('reconciled_ok', `captured=${parsed.capturedRevenue} excluded=${parsed.excludedRevenue} total=${parsed.totalReportRevenue}`, runContext);

  let summary;
  try {
    summary = await mergeAndPush(parsed, runContext);
  } catch (e) {
    await writeLog('failed', `Merge/push error: ${e.message}`, runContext);
    console.error('MERGE/PUSH FAILED:', e);
    process.exit(1);
  }

  // Verify
  const { json: ddFinal } = await ghGet('data/dd.json');
  const maxDate = ddFinal.dd.reduce((m, r) => (r[DD_DATE] > m ? r[DD_DATE] : m), '');

  await writeLog('done', 'Sync completed successfully', runContext, {
    confirmed_max_date: maxDate,
    confirmed_row_counts: { dd: ddFinal.dd.length, td: summary.td.after, install: summary.install.after },
    report_revenue_reconciled: parsed.reconcileOk,
  });

  console.log('DONE.');
  console.log(`Confirmed max date in dd.json: ${maxDate}`);
  console.log(`Row counts — dd: ${ddFinal.dd.length}, td: ${summary.td.after}, install: ${summary.install.after}`);
}

main().catch(e => {
  console.error('FATAL:', e);
  process.exit(1);
});
