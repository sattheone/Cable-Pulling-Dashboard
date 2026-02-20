/**
 * ══════════════════════════════════════════════════════════════
 *  E&I Cable Pulling Dashboard — Google Apps Script Backend
 * ══════════════════════════════════════════════════════════════
 *
 *  SETUP INSTRUCTIONS:
 *  1. Create a new Google Sheet
 *  2. Go to Extensions → Apps Script
 *  3. Delete any code in Code.gs and paste this entire file
 *  4. Click the 💾 save button
 *  5. Click "Deploy" → "New deployment"
 *  6. Choose type: "Web app"
 *     - Description: "Cable Dashboard API"
 *     - Execute as: "Me"
 *     - Who has access: "Anyone" (or "Anyone with Google Account" for auth)
 *  7. Click "Deploy" and copy the Web App URL
 *  8. Paste that URL into your HTML dashboard's SCRIPT_URL variable
 *
 *  The script will auto-create 5 sheets on first use:
 *    • Config    — project name, dates, cable types
 *    • BOQ       — bill of quantities per type
 *    • SRN       — delivery log
 *    • Manual    — manual entry overrides
 *    • Snapshots — weekly pulling snapshots
 */

// ─── Sheet names (auto-created) ───────────────────────────────
const SHEET_CONFIG    = 'Config';
const SHEET_BOQ       = 'BOQ';
const SHEET_SRN       = 'SRN';
const SHEET_MANUAL    = 'Manual';
const SHEET_SNAPSHOTS = 'Snapshots';

// ─── Helpers ──────────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function sheetToJSON(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
}

function clearAndWrite(sheet, headers, rows) {
  sheet.clearContents();
  const data = [headers, ...rows];
  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

// ─── Initialize sheets ───────────────────────────────────────
function initSheets() {
  getOrCreateSheet(SHEET_CONFIG, ['key', 'value']);
  getOrCreateSheet(SHEET_BOQ, ['type', 'total', 'color']);
  getOrCreateSheet(SHEET_SRN, ['type', 'date', 'length', 'ref']);
  getOrCreateSheet(SHEET_MANUAL, ['type', 'delivered', 'pulled', 'lastWeek', 'thisWeek']);
  getOrCreateSheet(SHEET_SNAPSHOTS, ['id', 'weekLabel', 'date', 'pulledJSON', 'total']);
}

// ─── READ all data ───────────────────────────────────────────
function readAllData() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config → key-value pairs → object
  const configSheet = ss.getSheetByName(SHEET_CONFIG);
  const configRows = sheetToJSON(configSheet);
  const config = {};
  configRows.forEach(r => {
    try { config[r.key] = JSON.parse(r.value); } 
    catch(e) { config[r.key] = r.value; }
  });

  // BOQ
  const boqSheet = ss.getSheetByName(SHEET_BOQ);
  const boqRows = sheetToJSON(boqSheet);
  const boq = {};
  boqRows.forEach(r => { boq[r.type] = { total: Number(r.total) || 0, color: r.color || '#6366f1' }; });

  // SRN
  const srnSheet = ss.getSheetByName(SHEET_SRN);
  const srn = sheetToJSON(srnSheet).map(r => ({
    type: r.type,
    date: String(r.date),
    length: Number(r.length) || 0,
    ref: r.ref || ''
  }));

  // Manual
  const manualSheet = ss.getSheetByName(SHEET_MANUAL);
  const manualRows = sheetToJSON(manualSheet);
  const manual = {};
  manualRows.forEach(r => {
    manual[r.type] = {
      delivered: (r.delivered === '' || r.delivered === null || r.delivered === 'null') ? null : Number(r.delivered),
      pulled: Number(r.pulled) || 0,
      lastWeek: Number(r.lastWeek) || 0,
      thisWeek: Number(r.thisWeek) || 0
    };
  });

  // Snapshots
  const snapSheet = ss.getSheetByName(SHEET_SNAPSHOTS);
  const snapRows = sheetToJSON(snapSheet);
  const snapshots = snapRows.map(r => {
    let pulled = {};
    try { pulled = JSON.parse(r.pulledJSON); } catch(e) {}
    return {
      id: Number(r.id) || Date.now(),
      weekLabel: r.weekLabel,
      date: String(r.date),
      pulled: pulled,
      total: Number(r.total) || 0
    };
  });

  return {
    project: {
      name: config.name || '',
      startDate: config.startDate || '',
      targetDate: config.targetDate || '',
      asOf: config.asOf || '',
      types: config.types || []
    },
    boq,
    srn,
    manual,
    snapshots
  };
}

// ─── WRITE all data ──────────────────────────────────────────
function writeAllData(payload) {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config
  if (payload.project) {
    const configSheet = ss.getSheetByName(SHEET_CONFIG);
    const configRows = [
      ['name', JSON.stringify(payload.project.name || '')],
      ['startDate', JSON.stringify(payload.project.startDate || '')],
      ['targetDate', JSON.stringify(payload.project.targetDate || '')],
      ['asOf', JSON.stringify(payload.project.asOf || '')],
      ['types', JSON.stringify(payload.project.types || [])]
    ];
    clearAndWrite(configSheet, ['key', 'value'], configRows);
  }

  // BOQ
  if (payload.boq) {
    const boqSheet = ss.getSheetByName(SHEET_BOQ);
    const boqRows = Object.entries(payload.boq).map(([type, data]) => {
      if (typeof data === 'object') return [type, data.total || 0, data.color || ''];
      return [type, Number(data) || 0, ''];
    });
    clearAndWrite(boqSheet, ['type', 'total', 'color'], boqRows);
  }

  // SRN
  if (payload.srn) {
    const srnSheet = ss.getSheetByName(SHEET_SRN);
    const srnRows = payload.srn.map(r => [r.type, r.date, r.length, r.ref || '']);
    clearAndWrite(srnSheet, ['type', 'date', 'length', 'ref'], srnRows);
  }

  // Manual
  if (payload.manual) {
    const manualSheet = ss.getSheetByName(SHEET_MANUAL);
    const manualRows = Object.entries(payload.manual).map(([type, data]) => [
      type,
      data.delivered === null ? 'null' : (data.delivered || 0),
      data.pulled || 0,
      data.lastWeek || 0,
      data.thisWeek || 0
    ]);
    clearAndWrite(manualSheet, ['type', 'delivered', 'pulled', 'lastWeek', 'thisWeek'], manualRows);
  }

  // Snapshots
  if (payload.snapshots) {
    const snapSheet = ss.getSheetByName(SHEET_SNAPSHOTS);
    const snapRows = payload.snapshots.map(s => [
      s.id,
      s.weekLabel,
      s.date,
      JSON.stringify(s.pulled || {}),
      s.total || 0
    ]);
    clearAndWrite(snapSheet, ['id', 'weekLabel', 'date', 'pulledJSON', 'total'], snapRows);
  }

  return { success: true };
}

// ─── Partial update: just one sheet ──────────────────────────
function writePartial(sheetKey, payload) {
  initSheets();
  const wrapper = {};
  wrapper[sheetKey] = payload;
  return writeAllData(wrapper);
}

// ═══ Web App Entry Points ════════════════════════════════════

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'read';
  let result;

  try {
    if (action === 'read') {
      result = readAllData();
    } else {
      result = { error: 'Unknown GET action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;

  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'writeAll';

    if (action === 'writeAll') {
      result = writeAllData(body.data || {});
    } else if (action === 'writePartial') {
      result = writePartial(body.sheetKey, body.data);
    } else {
      result = { error: 'Unknown POST action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Test from Apps Script editor ────────────────────────────
function testRead() {
  const data = readAllData();
  Logger.log(JSON.stringify(data, null, 2));
}

function testSeedData() {
  // Run this once to populate the sheet with sample data
  writeAllData({
    project: {
      name: 'E&I Cable Pulling',
      startDate: '2024-09-01',
      targetDate: '2026-09-30',
      asOf: '2026-01-30',
      types: [
        { name: 'HV',  color: '#3b82f6' },
        { name: 'LV',  color: '#8b5cf6' },
        { name: 'Ins', color: '#ec4899' },
        { name: 'Tel', color: '#f59e0b' },
        { name: 'HT',  color: '#10b981' }
      ]
    },
    boq: { HV: { total: 35275, color: '#3b82f6' }, LV: { total: 367298, color: '#8b5cf6' }, Ins: { total: 462247, color: '#ec4899' }, Tel: { total: 196468, color: '#f59e0b' }, HT: { total: 276075, color: '#10b981' } },
    srn: [
      { type:'HV', date:'05/11/2024', length:9754,  ref:'SRN-00286' },
      { type:'HV', date:'19/12/2024', length:3023,  ref:'SRN-00356' },
      { type:'HV', date:'19/09/2025', length:16327, ref:'SRN-00286' },
      { type:'HV', date:'09/01/2026', length:4321,  ref:'SRN-01170' },
      { type:'HV', date:'19/04/2026', length:1850,  ref:'No SRN' },
      { type:'LV', date:'19/11/2025', length:1752,  ref:'SRN-01033' },
      { type:'LV', date:'25/11/2025', length:28713, ref:'SRN-01090' },
      { type:'LV', date:'22/03/2025', length:75090, ref:'SRN-00553' },
      { type:'LV', date:'12/01/2026', length:95640, ref:'SRN-01152' },
      { type:'LV', date:'22/01/2026', length:36163, ref:'SRN-01180' },
      { type:'LV', date:'26/12/2025', length:112635, ref:'SRN-01252' },
      { type:'LV', date:'19/04/2026', length:17305, ref:'No SRN' },
      { type:'Ins', date:'07/07/2025', length:44328, ref:'SRN-00840' },
      { type:'Ins', date:'31/01/2026', length:217648, ref:'SRN-01239' },
      { type:'Ins', date:'11/03/2026', length:196475, ref:'SRN-01426' },
      { type:'Ins', date:'04/03/2026', length:3796, ref:'SRN-01435' },
      { type:'Tel', date:'07/11/2024', length:102302, ref:'SRN-00274' },
      { type:'Tel', date:'04/07/2025', length:84266, ref:'SRN-00757' },
      { type:'Tel', date:'15/12/2025', length:3500, ref:'SRN-01303' },
      { type:'Tel', date:'19/04/2026', length:4900, ref:'No SRN' },
      { type:'Tel', date:'27/06/2026', length:1500, ref:'No SRN' },
      { type:'HT', date:'22/09/2025', length:580, ref:'SRN-00334' },
      { type:'HT', date:'22/09/2025', length:9103, ref:'SRN-00855' },
      { type:'HT', date:'26/09/2025', length:30665, ref:'SRN-00933' },
      { type:'HT', date:'22/10/2025', length:21771, ref:'SRN-00939' },
      { type:'HT', date:'25/02/2026', length:70117, ref:'SRN-01295' },
      { type:'HT', date:'19/04/2026', length:60420, ref:'SRN-01446' },
      { type:'HT', date:'19/04/2026', length:83419, ref:'No SRN' }
    ],
    manual: {
      HV:  { delivered: null, pulled: 12750, lastWeek: 4346, thisWeek: 0 },
      LV:  { delivered: null, pulled: 14098, lastWeek: 0, thisWeek: 8480 },
      Ins: { delivered: null, pulled: 0, lastWeek: 0, thisWeek: 0 },
      Tel: { delivered: null, pulled: 0, lastWeek: 0, thisWeek: 0 },
      HT:  { delivered: null, pulled: 3300, lastWeek: 0, thisWeek: 0 }
    },
    snapshots: []
  });
  Logger.log('✅ Seed data written!');
}
