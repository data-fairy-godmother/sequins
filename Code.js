/**
 * Sequins ✨ — Code.gs
 * Assembly line sequencing agent for Farmer's Fridge
 * OPSICLE vNext
 *
 * Deploy as Web App:
 *   Execute as: Me
 *   Who has access: Anyone in Farmer's Fridge
 */

// ─── SHEET IDs ────────────────────────────────────────────────────────────────
const DEMANDS_SHEET_ID  = '1yOQ_xp3kGZ3hwqbNuZA_l5v2KsziDwr2YjRbdMOL_00';
const FORECAST_SHEET_ID = '1wyHr4QhvRGfyHgYURY7k5vLJFrpV3AX_wo5hkkk151A';
const FORECAST_TAB      = 'Summary';

// SKU attribute sources (replaces the old Seq Input aggregator)
const MASTER_DOC_SHEET_ID   = '1FRJ77-304M51SLwxrmljjZlrqv3YGMO_DRVrGu0pDBQ'; // Assembly Sequencing 2.0
const MASTER_DOC_TAB        = 'Master Document';   // cols G(name) - J(UPM), H(Optimal HC)
const MENU_LIBRARY_SHEET_ID = '1Exdh-emJxD7TohJ3IzjIQZDP3siPgjuVXhp3J7Gw2Ik'; // Menu Library
const MENU_LIBRARY_TAB      = 'Full Menu Summary';  // B=Category, C=SKU Name, L=Package, M=Allergens
const PROCESSING_SHEET_ID   = '1v_C2ZUR9_PjTqCO4XU16x2oRTvmvpdT43cs0d3tyh54'; // FPLModel Engine
const PROCESSING_TAB        = 'Processing Complexity'; // B=SKU Name, E=90 Day Duration/Unit

// ─── STORAGE ──────────────────────────────────────────────────────────────────
const STATE_KEY  = 'sequins_state';
const ADMINS_KEY = 'sequins_admins';
const AUDIT_SHEET_ID = '10yoKW7U76VW-GTuPTfNRxIQSTiegZqpOrDRZnZPI1Es';

// Stop reading SKU rows when we hit this sentinel in Demands 2025
const DEMANDS_STOP_SKU = 'VITAL_FARMS_EGGS';

// Default admins — always has full access
const DEFAULT_ADMINS = [
  'cori.blackburn@farmersfridge.com',
];

// Default rules editors — can edit line config + sequencing rules
const DEFAULT_RULES_EDITORS = [
  'cori.blackburn@farmersfridge.com',
  'smunshi@farmersfridge.com',
];

// ─── SERVE UI ─────────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Sequins ✨')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
function getCurrentUser() {
  const email      = Session.getActiveUser().getEmail();
  const admins     = getAdminList_();
  const isAdmin    = admins.map(a => a.toLowerCase()).includes(email.toLowerCase());
  const state      = getState() || {};
  const planners   = state.planners || [];
  const isPlanner  = planners.map(p => p.toLowerCase()).includes(email.toLowerCase());
  const rulesEditors = getRulesEditorList_();
  const canEditRules = isAdmin || rulesEditors.map(r => r.toLowerCase()).includes(email.toLowerCase());
  return { email, isAdmin, isPlanner, canEditRules, name: email.split('@')[0] };
}

function getAdminList_() {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty(ADMINS_KEY);
  try { return stored ? JSON.parse(stored) : DEFAULT_ADMINS; }
  catch(e) { return DEFAULT_ADMINS; }
}

// Public wrapper — client-side google.script.run cannot call functions
// with a trailing underscore (private-by-convention), so this exposes
// the same data through a normal function name.
function getAdminListForClient() {
  return getAdminList_();
}

function getRulesEditorList_() {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty('sequins_rules_editors');
  try { return stored ? JSON.parse(stored) : DEFAULT_RULES_EDITORS; }
  catch(e) { return DEFAULT_RULES_EDITORS; }
}

function saveRulesEditors(list) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  PropertiesService.getScriptProperties().setProperty('sequins_rules_editors', JSON.stringify(list));
  return { ok: true };
}

function addAdmin(email) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  const list = getAdminList_();
  if (!list.map(e => e.toLowerCase()).includes(email.toLowerCase())) list.push(email.toLowerCase());
  PropertiesService.getScriptProperties().setProperty(ADMINS_KEY, JSON.stringify(list));
  return { ok: true };
}

function removeAdmin(email) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  if (email.toLowerCase() === user.email.toLowerCase()) throw new Error("Can't remove yourself");
  const list = getAdminList_().filter(e => e.toLowerCase() !== email.toLowerCase());
  PropertiesService.getScriptProperties().setProperty(ADMINS_KEY, JSON.stringify(list));
  return { ok: true };
}

// ─── STATE ────────────────────────────────────────────────────────────────────
function getState() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty(STATE_KEY);
  try { return raw ? JSON.parse(raw) : null; }
  catch(e) { return null; }
}

function saveState(state) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  state.lastModified = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return { ok: true };
}

function saveStateAsEditor(state) {
  // Allows rules editors (Samad) to save line config + rules
  const user = getCurrentUser();
  if (!user.isAdmin && !user.canEditRules) throw new Error('Not authorized');
  state.lastModified = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return { ok: true };
}

function getLastModified() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty(STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw).lastModified || null; }
  catch(e) { return null; }
}

// ─── DEMAND FETCH: COMPILED FORECAST ─────────────────────────────────────────
function fetchForecastWeeks() {
  const ss    = SpreadsheetApp.openById(FORECAST_SHEET_ID);
  const sheet = ss.getSheetByName(FORECAST_TAB);
  if (!sheet) throw new Error('Tab "' + FORECAST_TAB + '" not found in Compiled Forecast');

  const lastCol = sheet.getLastColumn();
  const row1    = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row2    = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const row3    = sheet.getRange(3, 1, 1, lastCol).getValues()[0];

  const weeks = {};
  const DAYS  = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];

  row1.forEach((cell, ci) => {
    const wkMatch = String(cell).match(/Week\s+(\d+)/i);
    if (!wkMatch) return;
    const wkNum   = parseInt(wkMatch[1]);
    const dayName = String(row2[ci]).trim();
    if (!DAYS.includes(dayName)) return;
    const dateVal = row3[ci];
    const dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    const yr    = new Date().getFullYear();
    const label = 'Wk ' + wkNum + ' · ' + yr;
    if (!weeks[label]) weeks[label] = { label, wkNum, days: [] };
    weeks[label].days.push(dayName);
  });

  return Object.values(weeks).sort((a, b) => a.wkNum - b.wkNum);
}

function fetchForecastWeekData(weekLabel) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');

  const ss    = SpreadsheetApp.openById(FORECAST_SHEET_ID);
  const sheet = ss.getSheetByName(FORECAST_TAB);
  if (!sheet) throw new Error('Tab "' + FORECAST_TAB + '" not found');

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const row1    = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row2    = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const row3    = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  const DAYS    = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];

  const tz = Session.getScriptTimeZone();
  // Find the column range for this week using the row1 header (handles merged cells)
  // then collect all day columns within that range using row3 dates
  const allWkHeaders = [];
  row1.forEach((cell, ci) => {
    const m = String(cell).match(/Week\s+(\d+)/i);
    if (m) allWkHeaders.push({ ci, label: 'Wk ' + parseInt(m[1]) + ' · ' + new Date().getFullYear() });
  });
  const thisWk = allWkHeaders.find(w => w.label === weekLabel);
  if (!thisWk) throw new Error('Week ' + weekLabel + ' not found in Summary tab');
  const nextWk = allWkHeaders.find(w => w.ci > thisWk.ci);
  const colEnd = nextWk ? nextWk.ci : row1.length;
  const weekCols = [];
  for (let ci = thisWk.ci; ci < colEnd; ci++) {
    const dayName = String(row2[ci]).trim();
    if (!DAYS.includes(dayName)) continue;
    const dateVal = row3[ci];
    const dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd') : '';
    weekCols.push({ col: ci, day: dayName, date: dateStr });
  }
  if (!weekCols.length) throw new Error('Week ' + weekLabel + ' not found in Summary tab');

  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const skuData = {};
  const dates   = {};
  weekCols.forEach(wc => { dates[wc.day] = wc.date; });

  const skuLib = (getState() || {}).skuLibrary || {};
  const hasLib = Object.keys(skuLib).length > 0;
  for (let r = 15; r < allData.length; r++) {
    const skuVal = allData[r][2];
    if (!skuVal) continue;
    const skuName = String(skuVal).trim();
    if (!skuName || skuName === 'SKU') continue;
    if (hasLib) {
      const libEntry = skuLib[skuName.toUpperCase()];
      if (!libEntry || libEntry.active === false) continue;
    }
    weekCols.forEach(wc => {
      const qty = Math.round(parseFloat(allData[r][wc.col]) || 0);
      if (qty <= 0) return;
      if (!skuData[wc.day]) skuData[wc.day] = {};
      skuData[wc.day][skuName] = qty;
    });
  }

  return { weekLabel, skuData, dates, mode: 'forecast' };
}

// ─── DEMAND FETCH: ACTUALS (DEMANDS 2025) ─────────────────────────────────────
function fetchActualDemand(startDate, endDate) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');

  const start = new Date(startDate + 'T12:00:00');
  const end   = new Date(endDate   + 'T12:00:00');
  const tz    = Session.getScriptTimeZone();
  const ss    = SpreadsheetApp.openById(DEMANDS_SHEET_ID);

  const allSheets  = ss.getSheets();
  const weekSheets = allSheets.filter(s => /\d{4}\s+Week\s+\d+/i.test(s.getName()));
  if (!weekSheets.length) throw new Error('No weekly tabs found in Demands 2025. Expected names like "2026 Week 27".');

  const skuData = {};
  const byDate  = {};

  weekSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol < 3 || lastRow < 4) return;

    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const dateRow = allData[2]; // Row 3 = dates

    const dateCols = [];
    for (let ci = 2; ci <= 8 && ci < dateRow.length; ci++) {
      const cellVal  = dateRow[ci];
      if (!cellVal) continue;
      const cellDate = cellVal instanceof Date ? new Date(cellVal) : new Date(cellVal);
      if (isNaN(cellDate.getTime())) continue;
      cellDate.setHours(12, 0, 0, 0);
      if (cellDate >= start && cellDate <= end) {
        const dateStr = Utilities.formatDate(cellDate, tz, 'yyyy-MM-dd');
        const dayStr  = Utilities.formatDate(cellDate, tz, 'EEEE');
        dateCols.push({ col: ci, date: dateStr, day: dayStr });
        byDate[dateStr] = { day: dayStr, col: ci };
      }
    }
    if (!dateCols.length) return;

    // SKUs start row 4 (index 3), col A (index 0). Stop at VITAL_FARMS_EGGS.
    for (let r = 3; r < allData.length; r++) {
      const skuVal = allData[r][0];
      if (!skuVal) continue;
      const skuName = String(skuVal).trim();
      if (!skuName) continue;
      if (skuName.toUpperCase() === DEMANDS_STOP_SKU) break;
      // No library gate
      dateCols.forEach(dc => {
        const qty = Math.round(parseFloat(allData[r][dc.col]) || 0);
        if (qty <= 0) return;
        const key = dc.day + '|' + dc.date;
        if (!skuData[key]) skuData[key] = {};
        skuData[key][skuName] = (skuData[key][skuName] || 0) + qty;
      });
    }
  });

  if (!Object.keys(byDate).length)
    throw new Error('No dates found between ' + startDate + ' and ' + endDate + ' in Demands 2025.');

  const datesList = Object.entries(byDate)
    .map(([date, info]) => ({ date, day: info.day, col: info.col }))
    .sort((a, b) => a.date.localeCompare(b.date));

  return { skuData, byDate, dates: datesList, mode: 'actual' };
}

// ─── PUBLISH DEMAND ───────────────────────────────────────────────────────────
function publishForecastWeek(weekLabel, skuData, dates) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');

  let state = getState() || {};
  if (!state.demand) state.demand = {};
  if (!state.demand[weekLabel]) state.demand[weekLabel] = {};
  let daysLoaded = 0;

  Object.entries(skuData).forEach(([day, skus]) => {
    const existing = state.demand[weekLabel][day];
    if (existing && existing.mode === 'actual') return;
    const history = existing ? [
      { skus: existing.skus, mode: existing.mode, date: existing.date, savedAt: new Date().toISOString() },
      ...(existing.history || [])
    ].slice(0, 5) : [];
    state.demand[weekLabel][day] = {
      skus, mode: 'forecast', date: dates[day] || '',
      publishedBy: user.email, publishedAt: new Date().toISOString(), history
    };
    daysLoaded++;
  });

  writeAuditLog_(user.email, 'publish_forecast', weekLabel, '', daysLoaded + ' days');
  saveState(state);
  return { ok: true, weekLabel, daysLoaded };
}

function publishActualDays(entries) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');

  let state = getState() || {};
  if (!state.demand) state.demand = {};

  entries.forEach(entry => {
    const { weekLabel, day, date, skus } = entry;
    if (!state.demand[weekLabel]) state.demand[weekLabel] = {};
    const existing = state.demand[weekLabel][day];
    const history  = existing ? [
      { skus: existing.skus, mode: existing.mode, date: existing.date, savedAt: new Date().toISOString() },
      ...(existing.history || [])
    ].slice(0, 5) : [];
    state.demand[weekLabel][day] = {
      skus, mode: 'actual', date,
      publishedBy: user.email, publishedAt: new Date().toISOString(), history
    };
    writeAuditLog_(user.email, 'publish_actual', weekLabel, day, Object.keys(skus).length + ' SKUs');
  });

  saveState(state);
  return { ok: true, daysLoaded: entries.length };
}

// ─── SKU ATTRIBUTES (real sources — no guessing) ──────────────────────────────
/**
 * Pulls SKU attributes for a given list of SKU names directly from the
 * three source sheets that the old Seq Input formulas referenced:
 *   - Master Document (UPM, Optimal HC)
 *   - Full Menu Summary (Package, Allergens, Menu Category)
 *   - Processing Complexity (90 Day Duration/Unit)
 * Returns a map keyed by SKU name (uppercased) -> attributes.
 * SKUs not found in these sources are flagged, not guessed.
 */
function fetchSkuAttributesFor(skuNames) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.canEditRules) throw new Error('Not authorized');

  const wantedKeys = new Set(skuNames.map(s => normalizeSku_(s)));
  const result = {};
  const notFound = new Set(skuNames.map(s => s.toUpperCase()));

  // 1. Master Document — UPM (col J) + Optimal HC (col H), keyed by col G
  try {
    const mdSheet = SpreadsheetApp.openById(MASTER_DOC_SHEET_ID).getSheetByName(MASTER_DOC_TAB);
    if (mdSheet) {
      const lastRow = mdSheet.getLastRow();
      const data = mdSheet.getRange(1, 7, lastRow, 4).getValues(); // cols G:J
      data.forEach(row => {
        const name = String(row[0] || '').trim();
        if (!name) return;
        const key = normalizeSku_(name);
        if (!wantedKeys.has(key)) return;
        const hc  = parseFloat(row[1]); // col H
        const upm = parseFloat(row[3]); // col J
        if (!result[key]) result[key] = {};
        if (isFinite(hc))  result[key].optimalHC = hc;
        if (isFinite(upm)) result[key].upm = upm;
      });
    }
  } catch(e) { Logger.log('Master Document fetch failed: ' + e.message); }

  // 2. Full Menu Summary — Category (col B), Package (col L), Allergens (col M), keyed by col C
  try {
    const mlSheet = SpreadsheetApp.openById(MENU_LIBRARY_SHEET_ID).getSheetByName(MENU_LIBRARY_TAB);
    if (mlSheet) {
      const lastRow = mlSheet.getLastRow();
      const data = mlSheet.getRange(2, 2, lastRow - 1, 16).getValues(); // cols B:Q starting row 2
      data.forEach(row => {
        const name = String(row[1] || '').trim(); // col C = index 1 (0=B,1=C)
        if (!name) return;
        const key = normalizeSku_(name);
        if (!wantedKeys.has(key)) return;
        const category  = String(row[0] || '').trim();  // col B = index 0
        const packageTy = String(row[10] || '').trim();  // col L = index 10
        const allergens = String(row[11] || '').trim();  // col M = index 11
        if (!result[key]) result[key] = {};
        if (category)  result[key].category    = category;
        if (packageTy)  result[key].packageType = packageTy;
        if (allergens)  result[key].allergens   = allergens;
      });
    }
  } catch(e) { Logger.log('Menu Library fetch failed: ' + e.message); }

  // 3. Processing Complexity — 90 Day Duration/Unit (col E), keyed by col B
  try {
    const pcSheet = SpreadsheetApp.openById(PROCESSING_SHEET_ID).getSheetByName(PROCESSING_TAB);
    if (pcSheet) {
      const lastRow = pcSheet.getLastRow();
      const data = pcSheet.getRange(2, 2, lastRow - 1, 4).getValues(); // cols B:E starting row 2
      data.forEach(row => {
        const name = String(row[0] || '').trim(); // col B = index 0
        if (!name) return;
        const key = normalizeSku_(name);
        if (!wantedKeys.has(key)) return;
        const dur90 = parseFloat(row[3]); // col E = index 3
        if (!result[key]) result[key] = {};
        if (isFinite(dur90)) result[key].duration90Day = dur90;
      });
    }
  } catch(e) { Logger.log('Processing Complexity fetch failed: ' + e.message); }

  // Mark which originally-requested SKUs found nothing at all
  Object.keys(result).forEach(key => {
    // crude reverse-match: if we found anything for this key, clear it from notFound
    skuNames.forEach(orig => {
      if (normalizeSku_(orig) === key) notFound.delete(orig.toUpperCase());
    });
  });

  return { attributes: result, notFound: Array.from(notFound) };
}

function normalizeSku_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}


function saveSkuLibrary(library) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  let state = getState() || {};
  state.skuLibrary = library;
  saveState(state);
  return { ok: true };
}

// ─── LINE CONFIG + RULES ──────────────────────────────────────────────────────
function saveLineConfig(lineConfig) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.canEditRules) throw new Error('Not authorized');
  let state = getState() || {};
  state.lineConfig = lineConfig;
  writeAuditLog_(user.email, 'save_line_config', '', '', lineConfig.length + ' lines');
  saveStateAsEditor(state);
  return { ok: true };
}

function saveSequencingRules(rules) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.canEditRules) throw new Error('Not authorized');
  let state = getState() || {};
  state.sequencingRules = rules;
  writeAuditLog_(user.email, 'save_rules', '', '', JSON.stringify(rules));
  saveStateAsEditor(state);
  return { ok: true };
}

// ─── PLANNER MANAGEMENT ───────────────────────────────────────────────────────
function savePlanners(planners) {
  const user = getCurrentUser();
  if (!user.isAdmin) throw new Error('Not authorized');
  let state = getState() || {};
  state.planners = planners;
  saveState(state);
  return { ok: true };
}

// ─── AUDIT LOG ────────────────────────────────────────────────────────────────
function writeAuditLog_(email, action, week, day, detail) {
  try {
    const ss    = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet   = ss.getSheetByName('Sequins Audit');
    if (!sheet) {
      sheet = ss.insertSheet('Sequins Audit');
      sheet.appendRow(['Timestamp','Email','Action','Week','Day','Detail']);
      sheet.getRange(1,1,1,6).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), email, action, week||'', day||'', detail||'']);
  } catch(e) {
    Logger.log('Audit log write failed: ' + e.message);
  }
}

/**
 * Logs a SKU move in the Workbench to its own sheet tab — richer record
 * than the general audit log since moves need to show what rules (if any)
 * were broken, and carry an optional note.
 */
function writeSkuMoveLog_(entry) {
  try {
    const ss    = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet   = ss.getSheetByName('SKU Moves');
    if (!sheet) {
      sheet = ss.insertSheet('SKU Moves');
      sheet.appendRow(['Timestamp','Approved By','Week','Day','SKU','From Line','To Line','Violations','Note']);
      sheet.getRange(1,1,1,9).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      entry.email || '',
      entry.weekLabel || '',
      entry.day || '',
      entry.sku || '',
      entry.fromLine || '',
      entry.toLine || '',
      (entry.violations || []).join('; '),
      entry.note || ''
    ]);
  } catch(e) {
    Logger.log('SKU move log write failed: ' + e.message);
  }
}

/**
 * Saves a manual SKU placement override for a given week/day, and logs it.
 * Overrides are stored separately from the base demand so the sequencer's
 * automatic placement can still run underneath, with overrides applied
 * on top (move SKU X to line Y, regardless of where auto-placement put it).
 */
function saveSkuMove(weekLabel, day, sku, fromLine, toLine, violations, note) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.isPlanner && !user.canEditRules) throw new Error('Not authorized');

  let state = getState() || {};
  if (!state.overrides) state.overrides = {};
  if (!state.overrides[weekLabel]) state.overrides[weekLabel] = {};
  if (!state.overrides[weekLabel][day]) state.overrides[weekLabel][day] = {};

  state.overrides[weekLabel][day][sku] = {
    line: toLine,
    movedBy: user.email,
    movedAt: new Date().toISOString(),
    violations: violations || [],
    note: note || ''
  };

  writeSkuMoveLog_({
    email: user.email, weekLabel, day, sku, fromLine, toLine, violations, note
  });

  saveStateAsEditor(state);
  return { ok: true };
}


// ─── PUBLISHED PLAN ───────────────────────────────────────────────────────────
function savePublishedPlan(weekLabel, day, snapshot) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.isPlanner && !user.canEditRules) throw new Error('Not authorized');
  let state = getState() || {};
  if (!state.publishedPlans) state.publishedPlans = {};
  if (!state.publishedPlans[weekLabel]) state.publishedPlans[weekLabel] = {};
  state.publishedPlans[weekLabel][day] = snapshot;
  writeAuditLog_(user.email, 'publish_plan', weekLabel, day, Object.keys(snapshot.lineState || {}).length + ' lines');
  state.lastModified = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return { ok: true };
}


// ─── PUBLISHED PLAN ───────────────────────────────────────────────────────────
function savePublishedPlan(weekLabel, day, snapshot) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.isPlanner && !user.canEditRules) throw new Error('Not authorized');
  let state = getState() || {};
  if (!state.publishedPlans) state.publishedPlans = {};
  if (!state.publishedPlans[weekLabel]) state.publishedPlans[weekLabel] = {};
  state.publishedPlans[weekLabel][day] = snapshot;
  writeAuditLog_(user.email, 'publish_plan', weekLabel, day, Object.keys(snapshot.lineState || {}).length + ' lines');
  state.lastModified = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  return { ok: true };
}

function clearSkuMove(weekLabel, day, sku) {
  const user = getCurrentUser();
  if (!user.isAdmin && !user.isPlanner && !user.canEditRules) throw new Error('Not authorized');

  let state = getState() || {};
  if (state.overrides?.[weekLabel]?.[day]?.[sku]) {
    delete state.overrides[weekLabel][day][sku];
    writeAuditLog_(user.email, 'clear_sku_override', weekLabel, day, sku);
    saveStateAsEditor(state);
  }
  return { ok: true };
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getRecentSkuMoves(limit) {
  try {
    const ss    = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName('SKU Moves');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const lastRow = sheet.getLastRow();
    const n = Math.min(limit || 100, lastRow - 1);
    const rows = sheet.getRange(lastRow - n + 1, 1, n, 9).getValues();
    return rows.reverse().map(r => ({
      timestamp:  r[0] ? new Date(r[0]).toLocaleString('en-US', {month:'short',day:'numeric',hour:'numeric',minute:'2-digit',hour12:true}) : '',
      approvedBy: r[1], weekLabel: r[2], day: r[3], sku: r[4],
      fromLine: r[5], toLine: r[6], violations: r[7], note: r[8]
    }));
  } catch(e) {
    Logger.log('getRecentSkuMoves failed: ' + e.message);
    return [];
  }
}