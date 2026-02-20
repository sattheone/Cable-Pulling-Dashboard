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
 *  5. Run ▶ "testSeedData" once to populate sample data
 *  6. Click "Deploy" → "New deployment"
 *  7. Choose type: "Web app"
 *     - Description: "Cable Dashboard API"
 *     - Execute as: "Me"
 *     - Who has access: "Anyone"
 *  8. Click "Deploy" and copy the Web App URL
 *  9. Paste that URL into your Dashboard Settings → Import/Export
 *
 *  Auto-created sheets:
 *    📊 Dashboard — live summary (auto-updated, DO NOT EDIT)
 *    ⚙️ Config    — project name, dates, cable types
 *    📦 BOQ       — bill of quantities per cable type
 *    🚛 SRN       — delivery log (Store Receipt Notes)
 *    ✏️ Manual    — manual entry overrides for pulling data
 *    📸 Snapshots — weekly pulling snapshots
 */

// ─── Sheet names ──────────────────────────────────────────────
const SHEET_DASHBOARD = 'Dashboard';
const SHEET_CONFIG    = 'Config';
const SHEET_BOQ       = 'BOQ';
const SHEET_SRN       = 'SRN';
const SHEET_MANUAL    = 'Manual';
const SHEET_SNAPSHOTS = 'Snapshots';

// ─── Colors ───────────────────────────────────────────────────
const CLR_HEADER   = '#1e3a5f';
const CLR_HDR_TEXT = '#ffffff';
const CLR_HELP     = '#e8f0fe';
const CLR_HELP_TXT = '#5f6368';
const CLR_ALT_ROW  = '#f8f9fa';

// ─── Helpers ──────────────────────────────────────────────────
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function sheetToJSON(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  // Skip help row (row 2) if it starts with an info emoji
  const startRow = (data.length > 1 && String(data[1][0]).match(/^[ℹ💡⬇🔒📝⚙📅📏🔖🎨]/u)) ? 2 : 1;
  return data.slice(startRow).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function clearAndWrite(sheet, headers, helpRow, rows) {
  sheet.clearContents();
  sheet.clearFormats();
  const numCols = headers.length;

  // Header row
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setValues([headers]);
  headerRange.setBackground(CLR_HEADER).setFontColor(CLR_HDR_TEXT)
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);

  let dataStart = 2;

  // Help/instruction row
  if (helpRow && helpRow.length > 0) {
    const helpRange = sheet.getRange(2, 1, 1, numCols);
    helpRange.setValues([helpRow]);
    helpRange.setBackground(CLR_HELP).setFontColor(CLR_HELP_TXT)
      .setFontSize(9).setFontStyle('italic').setWrap(true);
    sheet.setFrozenRows(2);
    dataStart = 3;
  }

  // Data rows
  if (rows.length > 0) {
    sheet.getRange(dataStart, 1, rows.length, numCols).setValues(rows);
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(dataStart + i, 1, 1, numCols).setBackground(CLR_ALT_ROW);
      }
    }
  }

  // Auto-resize columns with min width
  for (let c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
    if (sheet.getColumnWidth(c) < 100) sheet.setColumnWidth(c, 100);
  }
}

// ─── Format SRN date to DD/MM/YYYY string ─────────────────────
function formatSrnDate(dateVal) {
  if (!dateVal) return '';
  const s = String(dateVal);
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return s;
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

// ─── Read transposed sheet (types as columns) ────────────────
function readTransposedSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { types: [], metrics: {} };

  const types = data[0].slice(1).map(String).filter(t => t.trim());

  // Skip help row if present
  let startRow = 1;
  if (data.length > 1 && String(data[1][0]).match(/^[📝✏️ℹ💡]/u)) startRow = 2;

  const metrics = {};
  for (let r = startRow; r < data.length; r++) {
    const label = String(data[r][0]).trim();
    if (!label) continue;
    metrics[label] = {};
    types.forEach((t, i) => {
      metrics[label][t] = data[r][i + 1];
    });
  }
  return { types, metrics };
}

// ─── Write transposed sheet (types as columns, metrics as rows) ─
function writeTransposedSheet(sheet, types, helpTexts, metricRows) {
  sheet.clearContents();
  sheet.clearFormats();
  const numCols = types.length + 1;

  // Header row: Cable Type | HV | LV | ...
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setValues([['Cable Type', ...types]]);
  headerRange.setBackground(CLR_HEADER).setFontColor(CLR_HDR_TEXT)
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  let dataStart = 2;

  // Help row
  if (helpTexts && helpTexts.length > 0) {
    const helpRange = sheet.getRange(2, 1, 1, numCols);
    helpRange.setValues([helpTexts]);
    helpRange.setBackground(CLR_HELP).setFontColor(CLR_HELP_TXT)
      .setFontSize(9).setFontStyle('italic').setWrap(true);
    sheet.setFrozenRows(2);
    dataStart = 3;
  }

  // Metric rows
  if (metricRows.length > 0) {
    sheet.getRange(dataStart, 1, metricRows.length, numCols).setValues(metricRows);
    for (let i = 0; i < metricRows.length; i++) {
      const r = dataStart + i;
      sheet.getRange(r, 1).setFontWeight('bold');
      if (i % 2 === 1) sheet.getRange(r, 1, 1, numCols).setBackground(CLR_ALT_ROW);
    }
  }

  // Auto-resize with min width
  for (let c = 1; c <= numCols; c++) {
    sheet.autoResizeColumn(c);
    if (sheet.getColumnWidth(c) < 120) sheet.setColumnWidth(c, 120);
  }
  sheet.setColumnWidth(1, 180);
}

// ─── Initialize sheets ───────────────────────────────────────
function initSheets() {
  getOrCreateSheet(SHEET_CONFIG);
  getOrCreateSheet(SHEET_BOQ);
  getOrCreateSheet(SHEET_SRN);
  getOrCreateSheet(SHEET_MANUAL);
  getOrCreateSheet(SHEET_SNAPSHOTS);
  getOrCreateSheet(SHEET_DASHBOARD);
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
    if (!r.key) return;
    try { config[r.key] = JSON.parse(r.value); }
    catch(e) { config[r.key] = r.value; }
  });

  // BOQ — transposed: types as columns
  const boqSheet = ss.getSheetByName(SHEET_BOQ);
  const { types: boqTypes, metrics: boqMetrics } = readTransposedSheet(boqSheet);
  const boq = {};
  boqTypes.forEach(t => {
    boq[t] = {
      total: Number((boqMetrics['BOQ Total (m)'] || {})[t]) || 0,
      color: String((boqMetrics['Color'] || {})[t]) || '#6366f1'
    };
  });

  // SRN — normalize dates
  const srnSheet = ss.getSheetByName(SHEET_SRN);
  const srn = sheetToJSON(srnSheet).filter(r => r.type).map(r => ({
    type: r.type,
    date: formatSrnDate(r.date),
    length: Number(r.length) || 0,
    ref: r.ref || ''
  }));

  // Manual — transposed: types as columns
  const manualSheet = ss.getSheetByName(SHEET_MANUAL);
  const { types: manTypes, metrics: manMetrics } = readTransposedSheet(manualSheet);
  const manual = {};
  manTypes.forEach(t => {
    const delVal = (manMetrics['Delivered (m)'] || {})[t];
    const isAuto = (delVal === '' || delVal === null || delVal === undefined ||
      String(delVal).toLowerCase() === 'auto' || String(delVal) === 'null');
    manual[t] = {
      delivered: isAuto ? null : Number(delVal),
      pulled: Number((manMetrics['Pulled (m)'] || {})[t]) || 0,
      lastWeek: Number((manMetrics['Last Week (m)'] || {})[t]) || 0,
      thisWeek: Number((manMetrics['This Week (m)'] || {})[t]) || 0
    };
  });

  // Snapshots
  const snapSheet = ss.getSheetByName(SHEET_SNAPSHOTS);
  const snapRows = sheetToJSON(snapSheet);
  const snapshots = snapRows.filter(r => r.id).map(r => {
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
    boq, srn, manual, snapshots
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
      ['name',       JSON.stringify(payload.project.name || '')],
      ['startDate',  JSON.stringify(payload.project.startDate || '')],
      ['targetDate', JSON.stringify(payload.project.targetDate || '')],
      ['asOf',       JSON.stringify(payload.project.asOf || '')],
      ['types',      JSON.stringify(payload.project.types || [])]
    ];
    clearAndWrite(configSheet,
      ['key', 'value'],
      ['⚙️ Setting name', '📝 JSON value (auto-managed by dashboard)'],
      configRows
    );
  }

  // BOQ — transposed layout (types as columns)
  if (payload.boq) {
    const boqSheet = ss.getSheetByName(SHEET_BOQ);
    const types = Object.keys(payload.boq);
    const totals = types.map(t => {
      const d = payload.boq[t];
      return typeof d === 'object' ? (d.total || 0) : (Number(d) || 0);
    });
    const colors = types.map(t => {
      const d = payload.boq[t];
      return typeof d === 'object' ? (d.color || '#6366f1') : '#6366f1';
    });
    writeTransposedSheet(boqSheet, types,
      ['📝 Metric', ...types.map(() => '✏️ Edit value')],
      [
        ['BOQ Total (m)', ...totals],
        ['Color',         ...colors]
      ]
    );
    // Number format for totals row
    if (types.length > 0) {
      boqSheet.getRange(3, 2, 1, types.length).setNumberFormat('#,##0');
    }
  }

  // SRN — force date column as plain text to prevent auto-formatting
  if (payload.srn) {
    const srnSheet = ss.getSheetByName(SHEET_SRN);
    const srnRows = payload.srn.map(r => [r.type, formatSrnDate(r.date), r.length, r.ref || '']);
    clearAndWrite(srnSheet,
      ['type', 'date', 'length', 'ref'],
      ['📝 Cable type', '📅 DD/MM/YYYY', '📏 Length in meters', '🔖 SRN reference'],
      srnRows
    );
    // Set date column as plain text to prevent Sheets auto-parsing
    srnSheet.getRange(1, 2, srnSheet.getMaxRows(), 1).setNumberFormat('@');
    if (srnRows.length > 0) {
      srnSheet.getRange(3, 2, srnRows.length, 1)
        .setValues(srnRows.map(r => [r[1]]));
      srnSheet.getRange(3, 3, srnRows.length, 1).setNumberFormat('#,##0');
    }
  }

  // Manual — transposed layout (types as columns)
  if (payload.manual) {
    const manualSheet = ss.getSheetByName(SHEET_MANUAL);
    const types = Object.keys(payload.manual);
    const delivered = types.map(t => payload.manual[t].delivered === null ? 'auto' : (payload.manual[t].delivered || 0));
    const pulled    = types.map(t => payload.manual[t].pulled || 0);
    const lastWk    = types.map(t => payload.manual[t].lastWeek || 0);
    const thisWk    = types.map(t => payload.manual[t].thisWeek || 0);
    writeTransposedSheet(manualSheet, types,
      ['📝 Metric', ...types.map(() => '✏️ Edit values in meters. Type "auto" for auto-calc delivered')],
      [
        ['Delivered (m)',  ...delivered],
        ['Pulled (m)',     ...pulled],
        ['Last Week (m)',  ...lastWk],
        ['This Week (m)',  ...thisWk]
      ]
    );
    // Number format for pulled/lastWeek/thisWeek rows (rows 4-6), skip delivered which may be "auto"
    if (types.length > 0) {
      manualSheet.getRange(4, 2, 3, types.length).setNumberFormat('#,##0');
    }
  }

  // Snapshots
  if (payload.snapshots) {
    const snapSheet = ss.getSheetByName(SHEET_SNAPSHOTS);
    const snapRows = payload.snapshots.map(s => [
      s.id, s.weekLabel, s.date,
      JSON.stringify(s.pulled || {}),
      s.total || 0
    ]);
    clearAndWrite(snapSheet,
      ['id', 'weekLabel', 'date', 'pulledJSON', 'total'],
      ['🔒 Auto ID', '📅 Week label', '📅 Date', '🔒 JSON data', '📏 Total pulled (m)'],
      snapRows
    );
  }

  // Refresh Dashboard summary sheet
  try { refreshDashboard(payload); } catch(e) { Logger.log('Dashboard refresh error: ' + e); }

  return { success: true };
}

// ─── Refresh the Dashboard summary sheet ─────────────────────
function refreshDashboard(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getOrCreateSheet(SHEET_DASHBOARD);
  dash.clearContents();
  dash.clearFormats();

  const project = payload.project || {};
  const boq = payload.boq || {};
  const srn = payload.srn || [];
  const manual = payload.manual || {};
  const types = (project.types || []).map(t => t.name || t);

  if (types.length === 0) {
    dash.getRange('A1').setValue('⚠️ No cable types configured yet. Run testSeedData or configure from the dashboard app.');
    return;
  }

  const numCols = types.length + 1;
  let row = 1;

  function writeRow(label, values, opts) {
    opts = opts || {};
    const rowData = [label, ...values];
    const range = dash.getRange(row, 1, 1, numCols);
    range.setValues([rowData]);
    if (opts.bold) range.setFontWeight('bold');
    if (opts.bg) range.setBackground(opts.bg);
    if (opts.fg) range.setFontColor(opts.fg);
    if (opts.fmt) dash.getRange(row, 2, 1, types.length).setNumberFormat(opts.fmt);
    if (opts.size) range.setFontSize(opts.size);
    row++;
  }

  function spacer() { row++; }

  // Title
  dash.getRange(row, 1, 1, numCols).merge();
  dash.getRange(row, 1).setValue('📊 E&I Cable Pulling Dashboard — ' + (project.name || 'Summary'))
    .setFontSize(14).setFontWeight('bold').setFontColor(CLR_HEADER);
  row++;

  dash.getRange(row, 1, 1, numCols).merge();
  dash.getRange(row, 1).setValue('📅 As of: ' + (project.asOf || 'N/A') + '   |   🎯 Target: ' + (project.targetDate || 'N/A') + '   |   🔄 Auto-updated — DO NOT EDIT this sheet')
    .setFontSize(9).setFontColor(CLR_HELP_TXT).setFontStyle('italic');
  row++;
  spacer();

  writeRow('Cable', types, { bold: true, bg: CLR_HEADER, fg: CLR_HDR_TEXT, size: 11 });

  // BOQ Total
  const boqTotals = types.map(t => {
    const b = boq[t];
    return (b ? (typeof b === 'object' ? b.total : Number(b)) : 0) / 1000;
  });
  writeRow('BOQ Total (km)', boqTotals, { fmt: '#,##0.000', bold: true });
  spacer();

  // Delivered
  const delivered = types.map(t => {
    const m = manual[t];
    if (m && m.delivered !== null && m.delivered !== 'null' && m.delivered !== undefined) {
      return Number(m.delivered) / 1000;
    }
    return srn.filter(r => r.type === t).reduce((s, r) => s + (Number(r.length) || 0), 0) / 1000;
  });
  writeRow('Delivered (km)', delivered, { fmt: '#,##0.000', bg: '#e6f4ea', bold: true });

  const pending = types.map((t, i) => boqTotals[i] - delivered[i]);
  writeRow('Pending to receive (km)', pending, { fmt: '#,##0.000', bg: '#fce8e6' });
  spacer();

  // Pulled
  const pulled = types.map(t => {
    const m = manual[t];
    return (m ? (m.pulled || 0) : 0) / 1000;
  });
  writeRow('Pulled (km)', pulled, { fmt: '#,##0.000', bg: '#e8f0fe', bold: true });

  const notPulled = types.map((t, i) => delivered[i] - pulled[i]);
  writeRow('Delivered − not pulled (km)', notPulled, { fmt: '#,##0.000' });

  const remaining = types.map((t, i) => boqTotals[i] - pulled[i]);
  writeRow('Remaining vs BOQ (km)', remaining, { fmt: '#,##0.000' });
  spacer();

  // Weekly
  const lastWeek = types.map(t => ((manual[t] || {}).lastWeek || 0) / 1000);
  const thisWeek = types.map(t => ((manual[t] || {}).thisWeek || 0) / 1000);
  writeRow('Last Week (km)', lastWeek, { fmt: '#,##0.00', bg: '#fff3e0' });
  writeRow('This Week (km)', thisWeek, { fmt: '#,##0.00', bg: '#fff3e0' });
  spacer();

  // Percentages
  const delPct = types.map((t, i) => boqTotals[i] > 0 ? delivered[i] / boqTotals[i] : 0);
  const pullPct = types.map((t, i) => boqTotals[i] > 0 ? pulled[i] / boqTotals[i] : 0);
  writeRow('Delivery %', delPct, { fmt: '0.0%', bold: true, bg: '#e6f4ea' });
  writeRow('Pulling %', pullPct, { fmt: '0.0%', bold: true, bg: '#e8f0fe' });
  spacer();

  const srnCount = types.map(t => srn.filter(r => r.type === t).length);
  writeRow('SRN Deliveries (count)', srnCount, { fmt: '0' });

  // Column widths
  dash.setColumnWidth(1, 230);
  for (let c = 2; c <= numCols; c++) dash.setColumnWidth(c, 120);

  // Protect dashboard
  try {
    const protection = dash.protect().setDescription('Auto-generated dashboard — do not edit');
    protection.setWarningOnly(true);
  } catch(e) {}

  // Move Dashboard to first position
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
}

// ─── Partial update: just one sheet ──────────────────────────
function writePartial(sheetKey, payload) {
  initSheets();
  try {
    const fullData = readAllData();
    if (sheetKey === 'project') fullData.project = payload;
    else if (sheetKey === 'boq') fullData.boq = payload;
    else if (sheetKey === 'srn') fullData.srn = payload;
    else if (sheetKey === 'manual') fullData.manual = payload;
    else if (sheetKey === 'snapshots') fullData.snapshots = payload;
    return writeAllData(fullData);
  } catch(e) {
    const wrapper = {};
    wrapper[sheetKey] = payload;
    return writeAllData(wrapper);
  }
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

// ─── Test functions ──────────────────────────────────────────
function testRead() {
  const data = readAllData();
  Logger.log(JSON.stringify(data, null, 2));
}

function testSeedData() {
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
    boq: {
      HV:  { total: 35275,  color: '#3b82f6' },
      LV:  { total: 367298, color: '#8b5cf6' },
      Ins: { total: 462247, color: '#ec4899' },
      Tel: { total: 196468, color: '#f59e0b' },
      HT:  { total: 276075, color: '#10b981' }
    },
    srn: [
      { type:'HV',  date:'05/11/2024', length:9754,   ref:'SRN-00286' },
      { type:'HV',  date:'19/12/2024', length:3023,   ref:'SRN-00356' },
      { type:'HV',  date:'19/09/2025', length:16327,  ref:'SRN-00286' },
      { type:'HV',  date:'09/01/2026', length:4321,   ref:'SRN-01170' },
      { type:'HV',  date:'19/04/2026', length:1850,   ref:'No SRN' },
      { type:'LV',  date:'19/11/2025', length:1752,   ref:'SRN-01033' },
      { type:'LV',  date:'25/11/2025', length:28713,  ref:'SRN-01090' },
      { type:'LV',  date:'22/03/2025', length:75090,  ref:'SRN-00553' },
      { type:'LV',  date:'12/01/2026', length:95640,  ref:'SRN-01152' },
      { type:'LV',  date:'22/01/2026', length:36163,  ref:'SRN-01180' },
      { type:'LV',  date:'26/12/2025', length:112635, ref:'SRN-01252' },
      { type:'LV',  date:'19/04/2026', length:17305,  ref:'No SRN' },
      { type:'Ins', date:'07/07/2025', length:44328,  ref:'SRN-00840' },
      { type:'Ins', date:'31/01/2026', length:217648, ref:'SRN-01239' },
      { type:'Ins', date:'11/03/2026', length:196475, ref:'SRN-01426' },
      { type:'Ins', date:'04/03/2026', length:3796,   ref:'SRN-01435' },
      { type:'Tel', date:'07/11/2024', length:102302, ref:'SRN-00274' },
      { type:'Tel', date:'04/07/2025', length:84266,  ref:'SRN-00757' },
      { type:'Tel', date:'15/12/2025', length:3500,   ref:'SRN-01303' },
      { type:'Tel', date:'19/04/2026', length:4900,   ref:'No SRN' },
      { type:'Tel', date:'27/06/2026', length:1500,   ref:'No SRN' },
      { type:'HT',  date:'22/09/2025', length:580,    ref:'SRN-00334' },
      { type:'HT',  date:'22/09/2025', length:9103,   ref:'SRN-00855' },
      { type:'HT',  date:'26/09/2025', length:30665,  ref:'SRN-00933' },
      { type:'HT',  date:'22/10/2025', length:21771,  ref:'SRN-00939' },
      { type:'HT',  date:'25/02/2026', length:70117,  ref:'SRN-01295' },
      { type:'HT',  date:'19/04/2026', length:60420,  ref:'SRN-01446' },
      { type:'HT',  date:'19/04/2026', length:83419,  ref:'No SRN' }
    ],
    manual: {
      HV:  { delivered: null, pulled: 12750,  lastWeek: 4346, thisWeek: 0 },
      LV:  { delivered: null, pulled: 14098,  lastWeek: 0,    thisWeek: 8480 },
      Ins: { delivered: null, pulled: 0,      lastWeek: 0,    thisWeek: 0 },
      Tel: { delivered: null, pulled: 0,      lastWeek: 0,    thisWeek: 0 },
      HT:  { delivered: null, pulled: 3300,   lastWeek: 0,    thisWeek: 0 }
    },
    snapshots: []
  });

  // Delete default Sheet1/Sheet2 if they exist
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const s1 = ss.getSheetByName('Sheet1');
    if (s1) ss.deleteSheet(s1);
    const s2 = ss.getSheetByName('Sheet2');
    if (s2) ss.deleteSheet(s2);
  } catch(e) {}

  Logger.log('✅ Seed data written with formatting! Check the Dashboard sheet.');
}
