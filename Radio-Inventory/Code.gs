var TEMPLATE_ROWS = 100, THEME_COLOR = '#356854', HEADER_TEXT_COLOR = '#ffffff';
var LABEL_BG = '#dbe8e1', READONLY_BG = '#f2f2f2';
var FORM_LABEL_COL_WIDTH = 220, FORM_INPUT_COL_WIDTH = 200, FORM_ROW_HEIGHT = 28;
var FORM_INPUT_FONT_SIZE = 10;
var CONDITION_LIST = ['Good', 'Repair', 'Awaiting Parts', 'Renewal', 'Missing', 'Retired'];
var RADIO_TYPE_LIST = ['Portable', 'Mobile'];
var CONDITION_COLORS = [
  { v: 'Good',           bg: '#b6d7a8', fg: '#000000' },
  { v: 'Awaiting Parts', bg: '#ffe599', fg: '#000000' },
  { v: 'Repair',         bg: '#cccccc', fg: '#000000' },
  { v: 'Missing',        bg: '#ea9999', fg: '#000000' },
  { v: 'Renewal',        bg: '#f9cb9c', fg: '#000000' },
  { v: 'Retired',        bg: '#666666', fg: '#ffffff' }
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Inventory & Maintenance')
    .addItem('Setup Radio Workbook', 'runSetupSafe_')
    .addSeparator()
    .addItem('Apply Condition Colors', 'applyConditionColorsOnly')
    .addItem('Refresh Validations Only', 'runRefreshValidationsSafe_')
    .addToUi();
}

function runSetupSafe_() { safeRun_(setupRadioWorkbook); }
function runRefreshValidationsSafe_() { safeRun_(refreshFlowAndValidations_); }
function isTypedColumnError_(m) { return /typed column|number format of cells in a typed/i.test(String(m || '')); }

function safeRun_(fn) {
  try {
    fn();
    try { SpreadsheetApp.flush(); } catch (eF) { if (!isTypedColumnError_(eF && eF.message ? eF.message : eF)) throw eF; }
  } catch (err) {
    var m = String(err && err.message ? err.message : err);
    if (isTypedColumnError_(m)) return;
    SpreadsheetApp.getUi().alert('Error: ' + m);
  }
}

function safeSetNumberFormat_(r, f) { try { r.setNumberFormat(f); } catch (e) {} }
function safeSetDataValidation_(r, v) { try { r.setDataValidation(v); } catch (e) {} }
function safeSetValue_(r, v) { try { r.setValue(v); } catch (e) {} }
function safeSetValues_(r, v) { try { r.setValues(v); } catch (e) {} }
function safeSetFormula_(r, f) { try { r.setFormula(f); } catch (e) {} }
function safeSetFormulas_(r, f) { try { r.setFormulas(f); } catch (e) {} }
function safeClearContent_(r) { try { r.clearContent(); } catch (e) {} }

function clearSheetSafe_(s) {
  try { s.clearContents(); } catch (e) {}
  try { s.clearFormats(); } catch (e) {}
  try { s.clearConditionalFormatRules(); } catch (e) {}
  try { if (s.getFilter()) s.getFilter().remove(); } catch (e) {}
  try { s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).clearDataValidations(); } catch (e) {}
}

function applyConditionColorsOnly() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ASSET_MASTER');
  if (!s) { SpreadsheetApp.getUi().alert('ASSET_MASTER not found'); return; }
  applyAssetMasterConditionColors_(s);
  SpreadsheetApp.getUi().alert('Condition colors applied.');
}

function applyAssetMasterConditionColors_(s) {
  var maxRows = s.getMaxRows();
  if (maxRows < 2) return;
  var rowRange  = s.getRange(2, 1, maxRows - 1, 14);
  var condRange = s.getRange(2, 12, maxRows - 1, 1);

  var existing = s.getConditionalFormatRules() || [];
  var kept = [];
  for (var i = 0; i < existing.length; i++) {
    var bc = existing[i].getBooleanCondition && existing[i].getBooleanCondition();
    var keep = true;
    if (bc) {
      var vals = bc.getCriteriaValues() || [];
      for (var k = 0; k < CONDITION_COLORS.length; k++) {
        var cv = CONDITION_COLORS[k].v.toUpperCase();
        for (var vi = 0; vi < vals.length; vi++) {
          var sv = String(vals[vi]).toUpperCase();
          if (sv === cv || sv.indexOf('"' + cv + '"') >= 0 || sv.indexOf(cv) >= 0) { keep = false; break; }
        }
        if (!keep) break;
      }
    }
    if (keep) kept.push(existing[i]);
  }

  var rules = kept.slice();
  for (var j = 0; j < CONDITION_COLORS.length; j++) {
    var c = CONDITION_COLORS[j];
    var upper = c.v.toUpperCase();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=UPPER(TRIM($L2))="' + upper + '"')
      .setBackground(c.bg)
      .setRanges([rowRange])
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=UPPER(TRIM($L2))="' + upper + '"')
      .setBackground(c.bg)
      .setFontColor(c.fg)
      .setBold(true)
      .setRanges([condRange])
      .build());
  }
  try { s.setConditionalFormatRules(rules); } catch (e) {}
}

function setupRadioWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var names = ['ASSET_MASTER', 'ISSUANCE_LOG', 'MAINTENANCE_LOG', 'STATUS_DASHBOARD', 'INVENTORY_AUDIT', 'PRINTABLE_RADIO_LOG'];
  var map = {};
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]); if (!sh) sh = ss.insertSheet(names[i]); map[names[i]] = sh;
  }
  var errors = [];
  try { setupAssetMasterSafe_(map['ASSET_MASTER']); } catch (e) { if (!isTypedColumnError_(e.message)) errors.push('ASSET_MASTER: ' + e.message); }
  try { setupIssuanceLogSafe_(map['ISSUANCE_LOG']); } catch (e) { if (!isTypedColumnError_(e.message)) errors.push('ISSUANCE_LOG: ' + e.message); }
  try { setupMaintenanceLogSafe_(map['MAINTENANCE_LOG']); } catch (e) { if (!isTypedColumnError_(e.message)) errors.push('MAINTENANCE_LOG: ' + e.message); }
  try { setupInventoryAuditSafe_(map['INVENTORY_AUDIT']); } catch (e) { if (!isTypedColumnError_(e.message)) errors.push('INVENTORY_AUDIT: ' + e.message); }
  try { setupPrintableRadioLog_(map['PRINTABLE_RADIO_LOG']); } catch (e) { if (!isTypedColumnError_(e.message)) errors.push('PRINTABLE: ' + e.message); }
  try { setupFormsAndFlow_(); } catch (e2) { if (!isTypedColumnError_(e2.message)) errors.push('FORMS: ' + e2.message); }
  try { setupStatusDashboard_(map['STATUS_DASHBOARD']); } catch (eD) { if (!isTypedColumnError_(eD.message)) errors.push('DASH: ' + eD.message); }
  var oldFlow = ss.getSheetByName('_FLOW');
  if (oldFlow) try { ss.deleteSheet(oldFlow); } catch (e) {}
  reorderSheets_(['AssetMasterForm', 'RadioIssuanceForm', 'MaintenanceEntryForm', 'InventoryAuditForm', 'RadioReturnForm',
    'ASSET_MASTER', 'ISSUANCE_LOG', 'MAINTENANCE_LOG', 'STATUS_DASHBOARD', 'INVENTORY_AUDIT', 'PRINTABLE_RADIO_LOG']);
  try { SpreadsheetApp.flush(); } catch (e) {}
  var f1 = ss.getSheetByName('AssetMasterForm'); if (f1) ss.setActiveSheet(f1);
  if (errors.length) ui.alert('Setup completed with issues:\n\n' + errors.join('\n'));
  else ui.alert('Setup completed. Tables and existing data preserved.');
}

function setHeadersIfBlank_(s, headers) {
  var first = s.getRange(1, 1).getValue();
  if (first === '' || first === null) {
    safeSetValues_(s.getRange(1, 1, 1, headers.length), [headers]);
    try { s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(THEME_COLOR).setFontColor(HEADER_TEXT_COLOR).setHorizontalAlignment('center'); } catch (e) {}
    try { s.setFrozenRows(1); } catch (e) {}
  }
}

function setupAssetMasterSafe_(s) {
  ensureRowsAtLeast_(s, TEMPLATE_ROWS); ensureColumns_(s, 14);
  var h = ['No.', 'Asset Tag', 'Serial Number', 'Radio Model', 'Radio Type', 'Date Acquired',
    'Assigned To', 'Unit', 'Current Status', 'Last Repair Date', 'Last Inventory Check',
    'Condition', 'Storage Location', 'Notes'];
  setHeadersIfBlank_(s, h);
  var w = [65, 120, 150, 180, 120, 130, 170, 130, 150, 150, 160, 140, 170, 250];
  for (var i = 0; i < w.length; i++) try { s.setColumnWidth(i + 1, w[i]); } catch (e) {}
  var er = s.getMaxRows();
  var typeDV = SpreadsheetApp.newDataValidation().requireValueInList(RADIO_TYPE_LIST, true).setAllowInvalid(false).build();
  safeSetDataValidation_(s.getRange(2, 5, er - 1, 1), typeDV);
  var condDV = SpreadsheetApp.newDataValidation().requireValueInList(CONDITION_LIST, true).setAllowInvalid(false).build();
  safeSetDataValidation_(s.getRange(2, 12, er - 1, 1), condDV);
  safeSetNumberFormat_(s.getRange(2, 2, er - 1, 1), '@');
  safeSetNumberFormat_(s.getRange(2, 6, er - 1, 1), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange(2, 10, er - 1, 1), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange(2, 11, er - 1, 1), 'yyyy-mm-dd');
  applyAssetMasterConditionColors_(s);
}

function setupIssuanceLogSafe_(s) {
  ensureRowsAtLeast_(s, TEMPLATE_ROWS); ensureColumns_(s, 12);
  var h = ['No.', 'Transaction ID', 'Asset Tag', 'Serial Number', 'User Name', 'Rank', 'Unit',
    'Date Issued', 'Issued By', 'Return Date', 'Status', 'Remarks'];
  setHeadersIfBlank_(s, h);
  var w = [65, 130, 120, 150, 170, 100, 120, 130, 140, 130, 110, 220];
  for (var i = 0; i < w.length; i++) try { s.setColumnWidth(i + 1, w[i]); } catch (e) {}
  var er = s.getMaxRows();
  applyAssetTagValidation_(s, 3);
  safeSetNumberFormat_(s.getRange(2, 2, er - 1, 1), '@');
  safeSetNumberFormat_(s.getRange(2, 8, er - 1, 1), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange(2, 10, er - 1, 1), 'yyyy-mm-dd');
}

function setupMaintenanceLogSafe_(s) {
  ensureRowsAtLeast_(s, TEMPLATE_ROWS); ensureColumns_(s, 13);
  var h = ['No.', 'Service ID', 'Asset Tag', 'Date Received', 'Problem Reported', 'Diagnosis',
    'Repair Action', 'Parts Replaced', 'Technician', 'Repair Start Date',
    'Repair Completion Date', 'Status', 'Notes'];
  setHeadersIfBlank_(s, h);
  var w = [65, 120, 120, 130, 200, 180, 200, 180, 150, 140, 170, 120, 220];
  for (var i = 0; i < w.length; i++) try { s.setColumnWidth(i + 1, w[i]); } catch (e) {}
  var er = s.getMaxRows();
  applyAssetTagValidation_(s, 3);
  safeSetNumberFormat_(s.getRange(2, 2, er - 1, 1), '@');
  safeSetNumberFormat_(s.getRange(2, 4, er - 1, 1), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange(2, 10, er - 1, 1), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange(2, 11, er - 1, 1), 'yyyy-mm-dd');
}

function setupInventoryAuditSafe_(s) {
  ensureRowsAtLeast_(s, TEMPLATE_ROWS); ensureColumns_(s, 9);
  var h = ['No.', 'Audit Date', 'Asset Tag', 'Serial Number', 'Assigned User',
    'Physical Verification', 'Condition', 'Verified By', 'Notes'];
  setHeadersIfBlank_(s, h);
  var w = [65, 130, 120, 150, 170, 170, 120, 140, 240];
  for (var i = 0; i < w.length; i++) try { s.setColumnWidth(i + 1, w[i]); } catch (e) {}
  var er = s.getMaxRows();
  applyAssetTagValidation_(s, 3);
  safeSetDataValidation_(s.getRange(2, 6, er - 1, 1),
    SpreadsheetApp.newDataValidation().requireValueInList(['Verified', 'Not Found', 'Sent to Repair'], true).setAllowInvalid(false).build());
  safeSetDataValidation_(s.getRange(2, 7, er - 1, 1),
    SpreadsheetApp.newDataValidation().requireValueInList(CONDITION_LIST, true).setAllowInvalid(false).build());
  safeSetNumberFormat_(s.getRange(2, 2, er - 1, 1), 'yyyy-mm-dd');
}

function setupPrintableRadioLog_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, TEMPLATE_ROWS); ensureColumns_(s, 8);
  var h = ['No.', 'Asset Tag', 'Serial', 'Model', 'Assigned User', 'Unit', 'Status', 'Last Repair'];
  safeSetValues_(s.getRange(1, 1, 1, h.length), [h]);
  try { s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground(THEME_COLOR).setFontColor(HEADER_TEXT_COLOR).setHorizontalAlignment('center'); } catch (e) {}
  try { s.setFrozenRows(1); } catch (e) {}
  var w = [65, 120, 140, 180, 170, 130, 130, 140];
  for (var i = 0; i < w.length; i++) try { s.setColumnWidth(i + 1, w[i]); } catch (e) {}
  safeSetFormula_(s.getRange('A2'), '=ARRAYFORMULA(IF(B2:B="","",ROW(B2:B)-1))');
  safeSetFormula_(s.getRange('B2'),
    '=IFERROR(FILTER({ASSET_MASTER!$B$2:$B,ASSET_MASTER!$C$2:$C,ASSET_MASTER!$D$2:$D,ASSET_MASTER!$G$2:$G,ASSET_MASTER!$H$2:$H,ASSET_MASTER!$I$2:$I,ASSET_MASTER!$J$2:$J},ASSET_MASTER!$B$2:$B<>""),"")');
  safeSetNumberFormat_(s.getRange('H2:H'), 'yyyy-mm-dd');
}

function setupStatusDashboard_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 80); ensureColumns_(s, 8);
  safeSetValues_(s.getRange('A1:B1'), [['Metric', 'Value']]);
  safeSetValues_(s.getRange('A2:A8'), [['Total Radios'], ['Assigned Radios'], ['Available Radios'],
    ['Radios Under Repair'], ['Awaiting Parts'], ['Renewal'], ['Missing Radios']]);
  safeSetFormulas_(s.getRange('B2:B8'), [
    ['=COUNTA(ASSET_MASTER!B:B)-1'], ['=COUNTIF(ASSET_MASTER!I:I,"ASSIGNED")'],
    ['=COUNTIF(ASSET_MASTER!I:I,"AVAILABLE")'], ['=COUNTIF(ASSET_MASTER!I:I,"UNDER REPAIR")'],
    ['=COUNTIF(ASSET_MASTER!I:I,"AWAITING PARTS")'], ['=COUNTIF(ASSET_MASTER!I:I,"RENEWAL")'],
    ['=COUNTIF(ASSET_MASTER!I:I,"MISSING")']
  ]);
  safeSetValues_(s.getRange('D1:E1'), [['Unit', 'Count']]);
  safeSetFormula_(s.getRange('D2'),
    '=IFERROR(QUERY(ASSET_MASTER!H2:H,"select H, count(H) where H is not null group by H order by H asc label H \'\' , count(H) \'\' ",0),"")');
  safeSetValues_(s.getRange('G1:H1'), [['Month', 'Repairs']]);
  safeSetFormula_(s.getRange('G2'),
    '=IFERROR(QUERY({TEXT(MAINTENANCE_LOG!K2:K,"yyyy-mm")},"select Col1, count(Col1) where Col1 is not null and Col1 <> \'1899-12\' group by Col1 order by Col1 asc label Col1 \'\' , count(Col1) \'\' ",0),"")');
  var hr = ['A1:B1', 'D1:E1', 'G1:H1'];
  for (var i = 0; i < hr.length; i++) {
    try { s.getRange(hr[i]).setFontWeight('bold').setBackground(THEME_COLOR).setFontColor(HEADER_TEXT_COLOR).setHorizontalAlignment('center'); } catch (e) {}
  }
  try { s.getRange('A2:B8').setBorder(true, true, true, true, true, true); } catch (e) {}
  try { s.getRange('B2:B8').setHorizontalAlignment('center'); } catch (e) {}
  try { s.setFrozenRows(1); } catch (e) {}
  var cw = [220, 110, null, 160, 110, null, 140, 110];
  for (var c = 0; c < cw.length; c++) if (cw[c]) try { s.setColumnWidth(c + 1, cw[c]); } catch (e) {}
  buildDashboardCharts_(s);
}

function applyFormStyle_(s, lastRow) {
  try { s.setColumnWidth(1, FORM_LABEL_COL_WIDTH); } catch (e) {}
  for (var c = 2; c <= 6; c++) try { s.setColumnWidth(c, FORM_INPUT_COL_WIDTH); } catch (e) {}
  for (var r = 6; r <= lastRow; r++) try { s.setRowHeight(r, FORM_ROW_HEIGHT); } catch (e) {}
  try { s.getRange('A6:F' + lastRow).setBorder(true, true, true, true, true, true).setFontSize(FORM_INPUT_FONT_SIZE).setVerticalAlignment('middle'); } catch (e) {}
  try { s.getRange('A6:A' + lastRow).setBackground(LABEL_BG).setFontWeight('bold').setHorizontalAlignment('left'); } catch (e) {}
  try { s.getRange('B6:F' + lastRow).setHorizontalAlignment('left'); } catch (e) {}
}

function setupFormsAndFlow_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupForm1Sheet_(getOrCreateSheet_(ss, 'AssetMasterForm'));
  setupForm2Sheet_(getOrCreateSheet_(ss, 'RadioIssuanceForm'));
  setupForm3Sheet_(getOrCreateSheet_(ss, 'MaintenanceEntryForm'));
  setupForm4Sheet_(getOrCreateSheet_(ss, 'InventoryAuditForm'));
  setupForm5Sheet_(getOrCreateSheet_(ss, 'RadioReturnForm'));
  refreshFlowAndValidations_();
}

function setupForm1Sheet_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 50); ensureColumns_(s, 6);
  styleFormTitle_(s, 'ASSET REGISTRATION');

  setFormLabel_(s, 6, 'Asset Tag');
  setFormLabel_(s, 7, 'Serial Number *');
  setFormLabel_(s, 8, 'Radio Model *');
  setFormLabel_(s, 9, 'Radio Type *');
  setFormLabel_(s, 10, 'Date Acquired *');
  setFormLabel_(s, 11, 'Assigned To');
  setFormLabel_(s, 12, 'Unit *');
  setFormLabel_(s, 13, 'Condition *');
  setFormLabel_(s, 14, 'Storage Location');
  setFormLabel_(s, 15, 'Notes');

  mergeA1Ranges_(s, ['B6:F6', 'B7:F7', 'B8:F8', 'B9:F9', 'B10:F10', 'B11:F11',
    'B12:F12', 'B13:F13', 'B14:F14', 'B15:F17']);
  applyFormStyle_(s, 17);

  safeSetValue_(s.getRange('B6'), 'Auto-generated on submit');
  safeSetNumberFormat_(s.getRange('B10'), 'yyyy-mm-dd');
  try { s.getRange('B15').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
  safeSetDataValidation_(s.getRange('B9'),
    SpreadsheetApp.newDataValidation().requireValueInList(RADIO_TYPE_LIST, true).setAllowInvalid(false).build());
  safeSetDataValidation_(s.getRange('B13'),
    SpreadsheetApp.newDataValidation().requireValueInList(CONDITION_LIST, true).setAllowInvalid(false).build());
  safeSetDataValidation_(s.getRange('B10'),
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
}

function setupForm2Sheet_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 50); ensureColumns_(s, 6);
  styleFormTitle_(s, 'RADIO ISSUANCE');

  setFormLabel_(s, 4, 'Transaction ID');
  setFormLabel_(s, 5, 'Asset Tag *');
  setFormLabel_(s, 6, 'Serial Number');
  setFormLabel_(s, 7, 'User Name *');
  setFormLabel_(s, 8, 'Rank');
  setFormLabel_(s, 9, 'Unit *');
  setFormLabel_(s, 10, 'Date Issued *');
  setFormLabel_(s, 11, 'Issued By *');
  setFormLabel_(s, 12, 'Return Date');
  setFormLabel_(s, 13, 'Remarks');

  mergeA1Ranges_(s, ['B4:F4', 'B5:F5', 'B6:F6', 'B7:F7', 'B8:F8', 'B9:F9',
    'B10:F10', 'B11:F11', 'B12:F12', 'B13:F15']);
  try { s.setColumnWidth(1, FORM_LABEL_COL_WIDTH); } catch (e) {}
  for (var c = 2; c <= 6; c++) try { s.setColumnWidth(c, FORM_INPUT_COL_WIDTH); } catch (e) {}
  for (var r = 4; r <= 15; r++) try { s.setRowHeight(r, FORM_ROW_HEIGHT); } catch (e) {}
  try { s.getRange('A4:F15').setBorder(true, true, true, true, true, true).setFontSize(FORM_INPUT_FONT_SIZE).setVerticalAlignment('middle'); } catch (e) {}
  try { s.getRange('A4:A15').setBackground(LABEL_BG).setFontWeight('bold').setHorizontalAlignment('left'); } catch (e) {}
  try { s.getRange('B4:F15').setHorizontalAlignment('left'); } catch (e) {}

  safeSetValue_(s.getRange('B4'), 'Auto-generated on submit');
  safeSetFormula_(s.getRange('B6'), '=IFERROR(VLOOKUP(B5,ASSET_MASTER!B:C,2,FALSE),"")');
  try { s.getRange('B6').setBackground(READONLY_BG); } catch (e) {}
  safeSetNumberFormat_(s.getRange('B10'), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange('B12'), 'yyyy-mm-dd');
  try { s.getRange('B13').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
  var dv = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  safeSetDataValidation_(s.getRange('B10'), dv); safeSetDataValidation_(s.getRange('B12'), dv);
}

function setupForm3Sheet_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 60); ensureColumns_(s, 6);
  styleFormTitle_(s, 'MAINTENANCE ENTRY');

  try { s.getRange('A4:A5').merge(); } catch (e) {}
  safeSetValue_(s.getRange('A4'), 'Search SRV/Asset:');
  try { s.getRange('A4').setBackground('#E8F0FE').setFontWeight('bold').setVerticalAlignment('middle'); } catch (e) {}
  try { s.getRange('B4:F4').merge(); } catch (e) {}
  safeSetValue_(s.getRange('B4'), '');
  try { s.getRange('B4').setBackground('#FFFFFF').setBorder(true, true, true, true, null, null); } catch (e) {}

  setFormLabel_(s, 6, 'Service ID');
  setFormLabel_(s, 7, 'Asset Tag *');
  setFormLabel_(s, 8, 'Date Received *');
  setFormLabel_(s, 9, 'Problem Reported *');
  setFormLabel_(s, 11, 'Diagnosis');
  setFormLabel_(s, 13, 'Repair Action');
  setFormLabel_(s, 15, 'Parts Replaced');
  setFormLabel_(s, 17, 'Technician *');
  setFormLabel_(s, 18, 'Repair Start Date *');
  setFormLabel_(s, 19, 'Repair Completion Date');
  setFormLabel_(s, 20, 'Notes');

  mergeA1Ranges_(s, ['B6:F6', 'B7:F7', 'B8:F8', 'B9:F10', 'B11:F12', 'B13:F14',
    'B15:F16', 'B17:F17', 'B18:F18', 'B19:F19', 'B20:F22']);
  applyFormStyle_(s, 22);

  safeSetValue_(s.getRange('B6'), 'Auto-generated on submit');
  safeSetNumberFormat_(s.getRange('B8'), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange('B18'), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange('B19'), 'yyyy-mm-dd');
  var taller = ['B9', 'B11', 'B13', 'B15', 'B20'];
  for (var i = 0; i < taller.length; i++) try { s.getRange(taller[i]).setWrap(true).setVerticalAlignment('top'); } catch (e) {}
  var dv = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  safeSetDataValidation_(s.getRange('B8'), dv);
  safeSetDataValidation_(s.getRange('B18'), dv);
  safeSetDataValidation_(s.getRange('B19'), dv);
}

function setupForm4Sheet_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 55); ensureColumns_(s, 6);
  styleFormTitle_(s, 'INVENTORY AUDIT');

  try { s.getRange('A4:A5').merge(); } catch (e) {}
  safeSetValue_(s.getRange('A4'), 'Search Asset:');
  try { s.getRange('A4').setBackground('#E8F0FE').setFontWeight('bold').setVerticalAlignment('middle'); } catch (e) {}
  try { s.getRange('B4:F4').merge(); } catch (e) {}
  safeSetValue_(s.getRange('B4'), '');
  try { s.getRange('B4').setBackground('#FFFFFF').setBorder(true, true, true, true, null, null); } catch (e) {}

  setFormLabel_(s, 6, 'Audit Row #');
  setFormLabel_(s, 7, 'Audit Date *');
  setFormLabel_(s, 8, 'Asset Tag *');
  setFormLabel_(s, 9, 'Serial Number');
  setFormLabel_(s, 10, 'Assigned User');
  setFormLabel_(s, 11, 'Audit Type *');
  setFormLabel_(s, 12, 'Physical Verification *');
  setFormLabel_(s, 13, 'Condition');
  setFormLabel_(s, 14, 'Verified By *');
  setFormLabel_(s, 15, 'Findings');
  setFormLabel_(s, 17, 'Notes');

  mergeA1Ranges_(s, ['B6:F6', 'B7:F7', 'B8:F8', 'B9:F9', 'B10:F10', 'B11:F11',
    'B12:F12', 'B13:F13', 'B14:F14', 'B15:F16', 'B17:F19']);
  applyFormStyle_(s, 19);

  safeSetValue_(s.getRange('B6'), 'Auto-set on submit');
  safeSetNumberFormat_(s.getRange('B7'), 'yyyy-mm-dd');
  safeSetFormula_(s.getRange('B9'), '=IFERROR(VLOOKUP(B8,ASSET_MASTER!B:C,2,FALSE),"")');
  safeSetFormula_(s.getRange('B10'), '=IFERROR(VLOOKUP(B8,ASSET_MASTER!B:G,6,FALSE),"")');
  try { s.getRange('B9:B10').setBackground(READONLY_BG); } catch (e) {}

  safeSetDataValidation_(s.getRange('B7'),
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  safeSetDataValidation_(s.getRange('B11'),
    SpreadsheetApp.newDataValidation().requireValueInList(['Monthly', 'Quarterly', 'Annual'], true).setAllowInvalid(false).build());
  safeSetDataValidation_(s.getRange('B12'),
    SpreadsheetApp.newDataValidation().requireValueInList(['Verified', 'Not Found', 'Sent to Repair'], true).setAllowInvalid(false).build());
  safeSetDataValidation_(s.getRange('B13'),
    SpreadsheetApp.newDataValidation().requireValueInList(CONDITION_LIST, true).setAllowInvalid(false).build());
  try { s.getRange('B15').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
  try { s.getRange('B17').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
}

function setupForm5Sheet_(s) {
  clearSheetSafe_(s); ensureRowsAtLeast_(s, 55); ensureColumns_(s, 6);
  styleFormTitle_(s, 'RADIO RETURN');

  try { s.getRange('A4:A5').merge(); } catch (e) {}
  safeSetValue_(s.getRange('A4'), 'Search TRX:');
  try { s.getRange('A4').setBackground('#E8F0FE').setFontWeight('bold').setVerticalAlignment('middle'); } catch (e) {}
  try { s.getRange('B4:F4').merge(); } catch (e) {}
  safeSetValue_(s.getRange('B4'), '');
  try { s.getRange('B4').setBackground('#FFFFFF').setBorder(true, true, true, true, null, null); } catch (e) {}

  setFormLabel_(s, 6, 'Transaction ID *');
  setFormLabel_(s, 7, 'Asset Tag');
  setFormLabel_(s, 8, 'Serial Number');
  setFormLabel_(s, 9, 'User Name');
  setFormLabel_(s, 10, 'Unit');
  setFormLabel_(s, 11, 'Date Issued');
  setFormLabel_(s, 12, 'Return Date *');
  setFormLabel_(s, 13, 'Returned By *');
  setFormLabel_(s, 14, 'Return Condition *');
  setFormLabel_(s, 15, 'Inspection Notes');
  setFormLabel_(s, 17, 'Remarks');

  mergeA1Ranges_(s, ['B6:F6', 'B7:F7', 'B8:F8', 'B9:F9', 'B10:F10', 'B11:F11',
    'B12:F12', 'B13:F13', 'B14:F14', 'B15:F16', 'B17:F19']);
  applyFormStyle_(s, 19);

  safeSetFormula_(s.getRange('B7'), '=IFERROR(INDEX(ISSUANCE_LOG!C:C,MATCH(B6,ISSUANCE_LOG!B:B,0)),"")');
  safeSetFormula_(s.getRange('B8'), '=IFERROR(INDEX(ISSUANCE_LOG!D:D,MATCH(B6,ISSUANCE_LOG!B:B,0)),"")');
  safeSetFormula_(s.getRange('B9'), '=IFERROR(INDEX(ISSUANCE_LOG!E:E,MATCH(B6,ISSUANCE_LOG!B:B,0)),"")');
  safeSetFormula_(s.getRange('B10'), '=IFERROR(INDEX(ISSUANCE_LOG!G:G,MATCH(B6,ISSUANCE_LOG!B:B,0)),"")');
  safeSetFormula_(s.getRange('B11'), '=IFERROR(INDEX(ISSUANCE_LOG!H:H,MATCH(B6,ISSUANCE_LOG!B:B,0)),"")');
  try { s.getRange('B7:B11').setBackground(READONLY_BG); } catch (e) {}
  safeSetNumberFormat_(s.getRange('B11'), 'yyyy-mm-dd');
  safeSetNumberFormat_(s.getRange('B12'), 'yyyy-mm-dd');
  safeSetDataValidation_(s.getRange('B12'),
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  safeSetDataValidation_(s.getRange('B14'),
    SpreadsheetApp.newDataValidation().requireValueInList(CONDITION_LIST, true).setAllowInvalid(false).build());
  try { s.getRange('B15').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
  try { s.getRange('B17').setWrap(true).setVerticalAlignment('top'); } catch (e) {}
}

function refreshFlowAndValidations_() { applyFlowValidations_(); }

function applyFlowValidations_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var m = ss.getSheetByName('ASSET_MASTER'); if (!m) return;
  var lr = Math.max(m.getMaxRows(), 2);
  var ar = m.getRange(2, 2, lr - 1, 1);
  var av = SpreadsheetApp.newDataValidation().requireValueInRange(ar, true).setAllowInvalid(true).build();
  var f2 = ss.getSheetByName('RadioIssuanceForm');
  if (f2) { safeSetDataValidation_(f2.getRange('B5:F5'), av); safeSetFormula_(f2.getRange('B6'), '=IFERROR(VLOOKUP(B5,ASSET_MASTER!B:C,2,FALSE),"")'); }
  var f3 = ss.getSheetByName('MaintenanceEntryForm'); if (f3) safeSetDataValidation_(f3.getRange('B7:F7'), av);
  var f4 = ss.getSheetByName('InventoryAuditForm'); if (f4) safeSetDataValidation_(f4.getRange('B8:F8'), av);
  var f5 = ss.getSheetByName('RadioReturnForm');
  var log = ss.getSheetByName('ISSUANCE_LOG');
  if (f5 && log) {
    var tr = log.getRange(2, 2, Math.max(log.getMaxRows() - 1, 1), 1);
    var tv = SpreadsheetApp.newDataValidation().requireValueInRange(tr, true).setAllowInvalid(true).build();
    safeSetDataValidation_(f5.getRange('B6:F6'), tv);
  }
  applyAssetTagValidation_(ss.getSheetByName('ISSUANCE_LOG'), 3);
  applyAssetTagValidation_(ss.getSheetByName('MAINTENANCE_LOG'), 3);
  applyAssetTagValidation_(ss.getSheetByName('INVENTORY_AUDIT'), 3);
  var am = ss.getSheetByName('ASSET_MASTER');
  if (am) applyAssetMasterConditionColors_(am);
}

function F1_SEARCH() { if (!ensureActiveForm_('AssetMasterForm')) return; safeRun_(btnSearchForm1_); }
function F1_SUBMIT() { if (!ensureActiveForm_('AssetMasterForm')) return; safeRun_(submitForm1ToAssetMaster); }
function F1_CLEAR() { if (!ensureActiveForm_('AssetMasterForm')) return; safeRun_(clearForm1_); }
function F1_UPDATE() { if (!ensureActiveForm_('AssetMasterForm')) return; safeRun_(btnUpdateForm1_); }
function F2_SEARCH() { if (!ensureActiveForm_('RadioIssuanceForm')) return; safeRun_(btnSearchForm2_); }
function F2_SUBMIT() { if (!ensureActiveForm_('RadioIssuanceForm')) return; safeRun_(submitForm2ToIssuanceLog); }
function F2_CLEAR() { if (!ensureActiveForm_('RadioIssuanceForm')) return; safeRun_(clearForm2_); }
function F2_UPDATE() { if (!ensureActiveForm_('RadioIssuanceForm')) return; safeRun_(btnUpdateForm2_); }
function F3_SEARCH() { if (!ensureActiveForm_('MaintenanceEntryForm')) return; safeRun_(btnSearchForm3_); }
function F3_SUBMIT() { if (!ensureActiveForm_('MaintenanceEntryForm')) return; safeRun_(submitForm3ToMaintenanceLog); }
function F3_CLEAR() { if (!ensureActiveForm_('MaintenanceEntryForm')) return; safeRun_(clearForm3_); }
function F3_UPDATE() { if (!ensureActiveForm_('MaintenanceEntryForm')) return; safeRun_(btnUpdateForm3_); }
function F4_SEARCH() { if (!ensureActiveForm_('InventoryAuditForm')) return; safeRun_(btnSearchForm4_); }
function F4_SUBMIT() { if (!ensureActiveForm_('InventoryAuditForm')) return; safeRun_(submitForm4ToInventoryAudit); }
function F4_CLEAR() { if (!ensureActiveForm_('InventoryAuditForm')) return; safeRun_(clearForm4_); }
function F4_UPDATE() { if (!ensureActiveForm_('InventoryAuditForm')) return; safeRun_(btnUpdateForm4_); }
function F5_SEARCH() { if (!ensureActiveForm_('RadioReturnForm')) return; safeRun_(btnSearchForm5_); }
function F5_SUBMIT() { if (!ensureActiveForm_('RadioReturnForm')) return; safeRun_(submitForm5RadioReturn); }
function F5_CLEAR() { if (!ensureActiveForm_('RadioReturnForm')) return; safeRun_(clearForm5_); }
function F5_UPDATE() { if (!ensureActiveForm_('RadioReturnForm')) return; safeRun_(btnUpdateForm5_); }

function submitForm1ToAssetMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('AssetMasterForm'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !m) { ui.alert('Missing sheets'); return; }
  var assetTag = trimText_(f.getRange('B6').getValue());
  var serial = trimText_(f.getRange('B7').getValue());
  var model = trimText_(f.getRange('B8').getValue());
  var type = trimText_(f.getRange('B9').getValue());
  var acquired = f.getRange('B10').getValue();
  var assignedTo = trimText_(f.getRange('B11').getValue());
  var unit = trimText_(f.getRange('B12').getValue());
  var cond = trimText_(f.getRange('B13').getValue()) || 'Good';
  var storage = trimText_(f.getRange('B14').getValue());
  var notes = trimText_(f.getRange('B15').getValue());
  if (!serial || !model || !type || isBlankValue_(acquired) || !unit) { ui.alert('Required: Serial, Model, Type, Date Acquired, Unit'); return; }
  if (!assetTag || /^Auto/i.test(assetTag)) assetTag = nextAssetTag_(m);
  else if (findAssetRowByTag_(m, assetTag) > 0) { ui.alert('Asset Tag exists.'); return; }
  var row = firstBlankRow_(m, 2, 2);
  if (row > m.getMaxRows()) try { m.insertRowsAfter(m.getMaxRows(), row - m.getMaxRows()); } catch (e) {}
  safeSetValue_(m.getRange(row, 2), assetTag);
  safeSetValue_(m.getRange(row, 3), serial);
  safeSetValue_(m.getRange(row, 4), model);
  safeSetValue_(m.getRange(row, 5), type);
  safeSetValue_(m.getRange(row, 6), acquired);
  safeSetNumberFormat_(m.getRange(row, 6), 'yyyy-mm-dd');
  safeSetValue_(m.getRange(row, 7), assignedTo);
  safeSetValue_(m.getRange(row, 8), unit);
  safeSetValue_(m.getRange(row, 12), cond);
  safeSetValue_(m.getRange(row, 13), storage);
  safeSetValue_(m.getRange(row, 14), notes);
  setMasterDerivedFormulasAtRow_(m, row);
  clearForm1_(); refreshFlowAndValidations_();
  ui.alert('Saved: ' + assetTag);
}

function submitForm2ToIssuanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('RadioIssuanceForm'), log = ss.getSheetByName('ISSUANCE_LOG'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !log || !m) { ui.alert('Missing sheets'); return; }
  var assetTag = trimText_(f.getRange('B5').getValue());
  var userName = trimText_(f.getRange('B7').getValue());
  var rank = trimText_(f.getRange('B8').getValue());
  var unit = trimText_(f.getRange('B9').getValue());
  var dateIssued = f.getRange('B10').getValue();
  var issuedBy = trimText_(f.getRange('B11').getValue());
  var rd = f.getRange('B12').getValue();
  var rem = trimText_(f.getRange('B13').getValue());
  if (!assetTag || !userName || !unit || isBlankValue_(dateIssued) || !issuedBy) { ui.alert('Required: Asset Tag, User, Unit, Date Issued, Issued By'); return; }
  var mRow = findAssetRowByTag_(m, assetTag);
  if (!mRow) { ui.alert('Asset Tag not found: ' + assetTag); return; }
  var cs = trimText_(m.getRange(mRow, 9).getValue()).toUpperCase();
  if (cs === 'ASSIGNED') {
    var resp = ui.alert('Already ASSIGNED. Issue anyway?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  } else if (['UNDER REPAIR', 'AWAITING PARTS', 'RENEWAL', 'MISSING', 'RETIRED'].indexOf(cs) >= 0) {
    ui.alert('Cannot issue: ' + cs); return;
  }
  var trx = nextId_(log, 2, 'TRX', 5);
  var row = firstBlankRow_(log, 2, 2);
  if (row > log.getMaxRows()) try { log.insertRowsAfter(log.getMaxRows(), row - log.getMaxRows()); } catch (e) {}
  safeSetValue_(log.getRange(row, 2), trx);
  safeSetValue_(log.getRange(row, 3), assetTag);
  safeSetValue_(log.getRange(row, 5), userName);
  safeSetValue_(log.getRange(row, 6), rank);
  safeSetValue_(log.getRange(row, 7), unit);
  safeSetValue_(log.getRange(row, 8), dateIssued);
  safeSetNumberFormat_(log.getRange(row, 8), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 9), issuedBy);
  safeSetValue_(log.getRange(row, 10), isBlankValue_(rd) ? '' : rd);
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 12), rem);
  setIssuanceDerivedFormulasAtRow_(log, row);
  safeSetValue_(m.getRange(mRow, 7), userName);
  safeSetValue_(m.getRange(mRow, 8), unit);
  clearForm2_(); refreshFlowAndValidations_();
  ui.alert('Saved: ' + trx);
}

function submitForm3ToMaintenanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('MaintenanceEntryForm'), log = ss.getSheetByName('MAINTENANCE_LOG'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !log || !m) { ui.alert('Missing sheets'); return; }
  var assetTag = trimText_(f.getRange('B7').getValue());
  var dr = f.getRange('B8').getValue();
  var prob = trimText_(f.getRange('B9').getValue());
  var diag = trimText_(f.getRange('B11').getValue());
  var act = trimText_(f.getRange('B13').getValue());
  var parts = trimText_(f.getRange('B15').getValue());
  var tech = trimText_(f.getRange('B17').getValue());
  var sd = f.getRange('B18').getValue();
  var dd = f.getRange('B19').getValue();
  var notes = trimText_(f.getRange('B20').getValue());
  if (!assetTag || isBlankValue_(dr) || !prob || !tech || isBlankValue_(sd)) { ui.alert('Required: Asset Tag, Date Received, Problem, Technician, Repair Start'); return; }
  if (!findAssetRowByTag_(m, assetTag)) { ui.alert('Asset Tag not found: ' + assetTag); return; }
  var srv = nextId_(log, 2, 'SRV', 5);
  var row = firstBlankRow_(log, 2, 2);
  if (row > log.getMaxRows()) try { log.insertRowsAfter(log.getMaxRows(), row - log.getMaxRows()); } catch (e) {}
  safeSetValue_(log.getRange(row, 2), srv);
  safeSetValue_(log.getRange(row, 3), assetTag);
  safeSetValue_(log.getRange(row, 4), dr);
  safeSetNumberFormat_(log.getRange(row, 4), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 5), prob);
  safeSetValue_(log.getRange(row, 6), diag);
  safeSetValue_(log.getRange(row, 7), act);
  safeSetValue_(log.getRange(row, 8), parts);
  safeSetValue_(log.getRange(row, 9), tech);
  safeSetValue_(log.getRange(row, 10), sd);
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 11), isBlankValue_(dd) ? '' : dd);
  safeSetNumberFormat_(log.getRange(row, 11), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 13), notes);
  setMaintenanceDerivedFormulasAtRow_(log, row);
  updateAssetConditionFromMaintenance_(assetTag, isBlankValue_(dd) ? 'IN REPAIR' : 'COMPLETED');
  clearForm3_(); refreshFlowAndValidations_();
  ui.alert('Saved: ' + srv);
}

function submitForm4ToInventoryAudit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('InventoryAuditForm'), audit = ss.getSheetByName('INVENTORY_AUDIT'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !audit || !m) { ui.alert('Missing sheets'); return; }
  var ad = f.getRange('B7').getValue();
  var assetTag = trimText_(f.getRange('B8').getValue());
  var atype = trimText_(f.getRange('B11').getValue());
  var ver = trimText_(f.getRange('B12').getValue());
  var cond = trimText_(f.getRange('B13').getValue());
  var verBy = trimText_(f.getRange('B14').getValue());
  var find = trimText_(f.getRange('B15').getValue());
  var notes = trimText_(f.getRange('B17').getValue());
  if (isBlankValue_(ad) || !assetTag || !atype || !ver || !verBy) { ui.alert('Required: Audit Date, Asset Tag, Audit Type, Verification, Verified By'); return; }
  var mRow = findAssetRowByTag_(m, assetTag);
  if (!mRow) { ui.alert('Asset Tag not found: ' + assetTag); return; }
  var row = firstBlankRow_(audit, 3, 2);
  if (row > audit.getMaxRows()) try { audit.insertRowsAfter(audit.getMaxRows(), row - audit.getMaxRows()); } catch (e) {}
  safeSetValue_(audit.getRange(row, 2), ad);
  safeSetNumberFormat_(audit.getRange(row, 2), 'yyyy-mm-dd');
  safeSetValue_(audit.getRange(row, 3), assetTag);
  safeSetValue_(audit.getRange(row, 6), ver);
  if (cond) safeSetValue_(audit.getRange(row, 7), cond);
  safeSetValue_(audit.getRange(row, 8), verBy);
  var combo = '[' + atype + ']';
  if (find) combo += ' Findings: ' + find;
  if (notes) combo += ' | Notes: ' + notes;
  safeSetValue_(audit.getRange(row, 9), combo);
  if (ver === 'Not Found') safeSetValue_(m.getRange(mRow, 12), 'Missing');
  else if (ver === 'Sent to Repair') safeSetValue_(m.getRange(mRow, 12), 'Repair');
  else if (ver === 'Verified' && cond) safeSetValue_(m.getRange(mRow, 12), cond);
  safeSetFormula_(audit.getRange(row, 1), '=IF($C' + row + '="","",ROW()-1)');
  safeSetNumberFormat_(audit.getRange(row, 1), '0');
  safeSetFormula_(audit.getRange(row, 4), '=IF($C' + row + '="","",IFERROR(VLOOKUP($C' + row + ',ASSET_MASTER!$B:$C,2,FALSE),""))');
  safeSetFormula_(audit.getRange(row, 5), '=IF($C' + row + '="","",IFERROR(VLOOKUP($C' + row + ',ASSET_MASTER!$B:$G,6,FALSE),""))');
  clearForm4_(); refreshFlowAndValidations_();
  ui.alert('Saved row ' + row);
}

function submitForm5RadioReturn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('RadioReturnForm'), log = ss.getSheetByName('ISSUANCE_LOG'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !log || !m) { ui.alert('Missing sheets'); return; }
  var trx = trimText_(f.getRange('B6').getValue());
  var rd = f.getRange('B12').getValue();
  var rby = trimText_(f.getRange('B13').getValue());
  var rcond = trimText_(f.getRange('B14').getValue());
  var insp = trimText_(f.getRange('B15').getValue());
  var rem = trimText_(f.getRange('B17').getValue());
  if (!trx || isBlankValue_(rd) || !rby || !rcond) { ui.alert('Required: Transaction ID, Return Date, Returned By, Condition'); return; }
  var row = btnFindRowByValue_(log, 2, trx);
  if (!row) { ui.alert('Transaction not found: ' + trx); return; }
  var ex = log.getRange(row, 10).getValue();
  if (!isBlankValue_(ex)) {
    var resp = ui.alert('Already has Return Date. Overwrite?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }
  var assetTag = trimText_(log.getRange(row, 3).getValue());
  safeSetValue_(log.getRange(row, 10), rd);
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  var er = trimText_(log.getRange(row, 12).getValue());
  var nr = 'RETURNED by ' + rby + ' (' + rcond + ')';
  if (insp) nr += ' | Inspection: ' + insp;
  if (rem) nr += ' | Remarks: ' + rem;
  safeSetValue_(log.getRange(row, 12), er ? er + ' | ' + nr : nr);
  setIssuanceDerivedFormulasAtRow_(log, row);
  var mRow = findAssetRowByTag_(m, assetTag);
  if (mRow) {
    safeClearContent_(m.getRange(mRow, 7));
    safeSetValue_(m.getRange(mRow, 12), rcond);
  }
  clearForm5_(); refreshFlowAndValidations_();
  ui.alert('Returned: ' + trx + ' (' + assetTag + ')');
}

function btnSearchForm1_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('AssetMasterForm'), m = ss.getSheetByName('ASSET_MASTER');
  var t = trimText_(f.getRange('B6').getValue());
  if (!t || /^Auto/i.test(t)) { ui.alert('Enter Asset Tag in field B6 to search.'); return; }
  var row = btnFindRowByValue_(m, 2, t);
  if (!row) { ui.alert('Not found.'); return; }
  safeSetValue_(f.getRange('B6'), m.getRange(row, 2).getValue());
  safeSetValue_(f.getRange('B7'), m.getRange(row, 3).getValue());
  safeSetValue_(f.getRange('B8'), m.getRange(row, 4).getValue());
  safeSetValue_(f.getRange('B9'), m.getRange(row, 5).getValue());
  btnSetDateOrBlank_(f.getRange('B10'), m.getRange(row, 6).getValue());
  safeSetValue_(f.getRange('B11'), m.getRange(row, 7).getValue());
  safeSetValue_(f.getRange('B12'), m.getRange(row, 8).getValue());
  safeSetValue_(f.getRange('B13'), m.getRange(row, 12).getValue());
  safeSetValue_(f.getRange('B14'), m.getRange(row, 13).getValue());
  safeSetValue_(f.getRange('B15'), m.getRange(row, 14).getValue());
  ui.alert('Loaded: ' + t);
}

function btnSearchForm2_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('RadioIssuanceForm'), log = ss.getSheetByName('ISSUANCE_LOG');
  var t = trimText_(f.getRange('B2').getValue());
  var row = 0;
  if (t) row = btnFindRowByValue_(log, 2, t);
  else { ui.alert('Enter Transaction ID or Asset Tag in search bar.'); return; }
  if (!row) row = btnFindLastRowByValue_(log, 3, t);
  if (!row) { ui.alert('Not found.'); return; }
  safeSetValue_(f.getRange('B4'), log.getRange(row, 2).getValue());
  safeSetValue_(f.getRange('B5'), log.getRange(row, 3).getValue());
  safeSetFormula_(f.getRange('B6'), '=IFERROR(VLOOKUP(B5,ASSET_MASTER!B:C,2,FALSE),"")');
  safeSetValue_(f.getRange('B7'), log.getRange(row, 5).getValue());
  safeSetValue_(f.getRange('B8'), log.getRange(row, 6).getValue());
  safeSetValue_(f.getRange('B9'), log.getRange(row, 7).getValue());
  btnSetDateOrBlank_(f.getRange('B10'), log.getRange(row, 8).getValue());
  safeSetValue_(f.getRange('B11'), log.getRange(row, 9).getValue());
  btnSetDateOrBlank_(f.getRange('B12'), log.getRange(row, 10).getValue());
  safeSetValue_(f.getRange('B13'), log.getRange(row, 12).getValue());
  ui.alert('Loaded transaction');
}

function btnSearchForm3_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('MaintenanceEntryForm'), log = ss.getSheetByName('MAINTENANCE_LOG');
  var sv = trimText_(f.getRange('B4').getValue());
  var row = 0;
  if (sv) row = btnFindRowByValue_(log, 2, sv);
  else { ui.alert('Enter Service ID or Asset Tag in search bar.'); return; }
  if (!row) row = btnFindLastRowByValue_(log, 3, sv);
  if (!row) { ui.alert('Not found.'); return; }
  safeSetValue_(f.getRange('B6'), log.getRange(row, 2).getValue());
  safeSetValue_(f.getRange('B7'), log.getRange(row, 3).getValue());
  btnSetDateOrBlank_(f.getRange('B8'), log.getRange(row, 4).getValue());
  safeSetValue_(f.getRange('B9'), log.getRange(row, 5).getValue());
  safeSetValue_(f.getRange('B11'), log.getRange(row, 6).getValue());
  safeSetValue_(f.getRange('B13'), log.getRange(row, 7).getValue());
  safeSetValue_(f.getRange('B15'), log.getRange(row, 8).getValue());
  safeSetValue_(f.getRange('B17'), log.getRange(row, 9).getValue());
  btnSetDateOrBlank_(f.getRange('B18'), log.getRange(row, 10).getValue());
  btnSetDateOrBlank_(f.getRange('B19'), log.getRange(row, 11).getValue());
  safeSetValue_(f.getRange('B20'), log.getRange(row, 13).getValue());
  ui.alert('Loaded service record');
}

function btnSearchForm4_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('InventoryAuditForm'), audit = ss.getSheetByName('INVENTORY_AUDIT');
  var rt = trimText_(f.getRange('B4').getValue());
  var row = 0;
  if (rt && /^\d+$/.test(rt)) row = parseInt(rt, 10);
  else if (rt) row = btnFindLastRowByValue_(audit, 3, rt);
  else { ui.alert('Enter Audit Row # or Asset Tag in search bar.'); return; }
  if (!row || row < 2) { ui.alert('Not found.'); return; }
  safeSetValue_(f.getRange('B6'), row);
  btnSetDateOrBlank_(f.getRange('B7'), audit.getRange(row, 2).getValue());
  safeSetValue_(f.getRange('B8'), audit.getRange(row, 3).getValue());
  safeSetValue_(f.getRange('B12'), audit.getRange(row, 6).getValue());
  safeSetValue_(f.getRange('B13'), audit.getRange(row, 7).getValue());
  safeSetValue_(f.getRange('B14'), audit.getRange(row, 8).getValue());
  safeSetValue_(f.getRange('B17'), audit.getRange(row, 9).getValue());
  ui.alert('Loaded audit record');
}

function btnSearchForm5_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var log = ss.getSheetByName('ISSUANCE_LOG'), f = ss.getSheetByName('RadioReturnForm');
  var trx = trimText_(f.getRange('B4').getValue());
  if (!trx) { ui.alert('Enter Transaction ID in search bar.'); return; }
  var row = btnFindRowByValue_(log, 2, trx);
  if (!row) { ui.alert('Not found: ' + trx); return; }
  safeSetValue_(f.getRange('B6'), trx);
  btnSetDateOrBlank_(f.getRange('B12'), log.getRange(row, 10).getValue());
  ui.alert('Loaded ' + trx + '\nAsset: ' + log.getRange(row, 3).getValue() + '\nUser: ' + log.getRange(row, 5).getValue());
}

function btnUpdateForm1_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('AssetMasterForm'), m = ss.getSheetByName('ASSET_MASTER');
  var t = trimText_(f.getRange('B6').getValue());
  if (!t) { ui.alert('Enter existing Asset Tag in field B6.'); return; }
  var row = btnFindRowByValue_(m, 2, t);
  if (!row) { ui.alert('Not found.'); return; }
  safeSetValue_(m.getRange(row, 3), f.getRange('B7').getValue());
  safeSetValue_(m.getRange(row, 4), f.getRange('B8').getValue());
  safeSetValue_(m.getRange(row, 5), f.getRange('B9').getValue());
  safeSetValue_(m.getRange(row, 6), f.getRange('B10').getValue());
  safeSetNumberFormat_(m.getRange(row, 6), 'yyyy-mm-dd');
  safeSetValue_(m.getRange(row, 7), f.getRange('B11').getValue());
  safeSetValue_(m.getRange(row, 8), f.getRange('B12').getValue());
  safeSetValue_(m.getRange(row, 12), f.getRange('B13').getValue());
  safeSetValue_(m.getRange(row, 13), f.getRange('B14').getValue());
  safeSetValue_(m.getRange(row, 14), f.getRange('B15').getValue());
  setMasterDerivedFormulasAtRow_(m, row);
  refreshFlowAndValidations_();
  ui.alert('Updated: ' + t);
}

function btnUpdateForm2_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('RadioIssuanceForm'), log = ss.getSheetByName('ISSUANCE_LOG');
  var trx = trimText_(f.getRange('B4').getValue());
  if (!trx) { ui.alert('Enter existing Transaction ID.'); return; }
  var row = btnFindRowByValue_(log, 2, trx);
  if (!row) { ui.alert('Not found.'); return; }
  var rd = f.getRange('B12').getValue();
  safeSetValue_(log.getRange(row, 3), f.getRange('B5').getValue());
  safeSetValue_(log.getRange(row, 5), f.getRange('B7').getValue());
  safeSetValue_(log.getRange(row, 6), f.getRange('B8').getValue());
  safeSetValue_(log.getRange(row, 7), f.getRange('B9').getValue());
  safeSetValue_(log.getRange(row, 8), f.getRange('B10').getValue());
  safeSetNumberFormat_(log.getRange(row, 8), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 9), f.getRange('B11').getValue());
  safeSetValue_(log.getRange(row, 10), isBlankValue_(rd) ? '' : rd);
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 12), f.getRange('B13').getValue());
  setIssuanceDerivedFormulasAtRow_(log, row);
  var m = ss.getSheetByName('ASSET_MASTER');
  if (m) {
    var mRow = findAssetRowByTag_(m, trimText_(f.getRange('B5').getValue()));
    if (mRow) {
      safeSetValue_(m.getRange(mRow, 7), f.getRange('B7').getValue());
      safeSetValue_(m.getRange(mRow, 8), f.getRange('B9').getValue());
    }
  }
  refreshFlowAndValidations_();
  ui.alert('Updated: ' + trx);
}

function btnUpdateForm3_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('MaintenanceEntryForm'), log = ss.getSheetByName('MAINTENANCE_LOG');
  var sv = trimText_(f.getRange('B6').getValue());
  if (!sv) { ui.alert('Enter existing Service ID.'); return; }
  var row = btnFindRowByValue_(log, 2, sv);
  if (!row) { ui.alert('Not found.'); return; }
  var assetTag = trimText_(f.getRange('B7').getValue()), dd = f.getRange('B19').getValue();
  safeSetValue_(log.getRange(row, 3), assetTag);
  safeSetValue_(log.getRange(row, 4), f.getRange('B8').getValue());
  safeSetNumberFormat_(log.getRange(row, 4), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 5), f.getRange('B9').getValue());
  safeSetValue_(log.getRange(row, 6), f.getRange('B11').getValue());
  safeSetValue_(log.getRange(row, 7), f.getRange('B13').getValue());
  safeSetValue_(log.getRange(row, 8), f.getRange('B15').getValue());
  safeSetValue_(log.getRange(row, 9), f.getRange('B17').getValue());
  safeSetValue_(log.getRange(row, 10), f.getRange('B18').getValue());
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 11), isBlankValue_(dd) ? '' : dd);
  safeSetNumberFormat_(log.getRange(row, 11), 'yyyy-mm-dd');
  safeSetValue_(log.getRange(row, 13), f.getRange('B20').getValue());
  setMaintenanceDerivedFormulasAtRow_(log, row);
  updateAssetConditionFromMaintenance_(assetTag, isBlankValue_(dd) ? 'IN REPAIR' : 'COMPLETED');
  refreshFlowAndValidations_();
  ui.alert('Updated: ' + sv);
}

function btnUpdateForm4_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('InventoryAuditForm'), audit = ss.getSheetByName('INVENTORY_AUDIT');
  var rt = trimText_(f.getRange('B6').getValue());
  if (!rt || !/^\d+$/.test(rt)) { ui.alert('Enter Audit Row #.'); return; }
  var row = parseInt(rt, 10);
  if (row < 2) { ui.alert('Invalid row.'); return; }
  safeSetValue_(audit.getRange(row, 2), f.getRange('B7').getValue());
  safeSetNumberFormat_(audit.getRange(row, 2), 'yyyy-mm-dd');
  safeSetValue_(audit.getRange(row, 3), f.getRange('B8').getValue());
  safeSetValue_(audit.getRange(row, 6), f.getRange('B12').getValue());
  safeSetValue_(audit.getRange(row, 7), f.getRange('B13').getValue());
  safeSetValue_(audit.getRange(row, 8), f.getRange('B14').getValue());
  var atype = trimText_(f.getRange('B11').getValue());
  var find = trimText_(f.getRange('B15').getValue());
  var notes = trimText_(f.getRange('B17').getValue());
  var combo = '[' + atype + ']';
  if (find) combo += ' Findings: ' + find;
  if (notes) combo += ' | Notes: ' + notes;
  safeSetValue_(audit.getRange(row, 9), combo);
  refreshFlowAndValidations_();
  ui.alert('Updated row ' + row);
}

function btnUpdateForm5_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var f = ss.getSheetByName('RadioReturnForm'), log = ss.getSheetByName('ISSUANCE_LOG'), m = ss.getSheetByName('ASSET_MASTER');
  if (!f || !log || !m) { ui.alert('Missing sheets'); return; }
  var trx = trimText_(f.getRange('B6').getValue());
  if (!trx) { ui.alert('Enter existing Transaction ID.'); return; }
  var row = btnFindRowByValue_(log, 2, trx);
  if (!row) { ui.alert('Not found: ' + trx); return; }
  var rd = f.getRange('B12').getValue();
  var rby = trimText_(f.getRange('B13').getValue());
  var rcond = trimText_(f.getRange('B14').getValue());
  var insp = trimText_(f.getRange('B15').getValue());
  var rem = trimText_(f.getRange('B17').getValue());
  if (isBlankValue_(rd) || !rby || !rcond) { ui.alert('Required: Return Date, Returned By, Condition'); return; }
  var assetTag = trimText_(log.getRange(row, 3).getValue());
  safeSetValue_(log.getRange(row, 10), rd);
  safeSetNumberFormat_(log.getRange(row, 10), 'yyyy-mm-dd');
  var nr = 'RETURNED by ' + rby + ' (' + rcond + ')';
  if (insp) nr += ' | Inspection: ' + insp;
  if (rem) nr += ' | Remarks: ' + rem;
  safeSetValue_(log.getRange(row, 12), nr);
  setIssuanceDerivedFormulasAtRow_(log, row);
  var mRow = findAssetRowByTag_(m, assetTag);
  if (mRow) safeSetValue_(m.getRange(mRow, 12), rcond);
  refreshFlowAndValidations_();
  ui.alert('Updated return: ' + trx + ' (' + assetTag + ')');
}

function clearForm1_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AssetMasterForm'); if (!f) return;
  safeSetValue_(f.getRange('B6'), 'Auto-generated on submit');
  safeClearContent_(f.getRange('B7:B9'));
  safeClearContent_(f.getRange('B10'));
  safeSetNumberFormat_(f.getRange('B10'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B11:B12'));
  safeSetValue_(f.getRange('B13'), 'Good');
  safeClearContent_(f.getRange('B14:B15'));
}

function clearForm2_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RadioIssuanceForm'); if (!f) return;
  safeSetValue_(f.getRange('B4'), 'Auto-generated on submit');
  safeClearContent_(f.getRange('B5'));
  safeSetFormula_(f.getRange('B6'), '=IFERROR(VLOOKUP(B5,ASSET_MASTER!B:C,2,FALSE),"")');
  safeClearContent_(f.getRange('B7:B9'));
  safeClearContent_(f.getRange('B10'));
  safeSetNumberFormat_(f.getRange('B10'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B11:B12'));
  safeSetNumberFormat_(f.getRange('B12'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B13'));
}

function clearForm3_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MaintenanceEntryForm'); if (!f) return;
  safeClearContent_(f.getRange('B4'));
  safeSetValue_(f.getRange('B6'), 'Auto-generated on submit');
  safeClearContent_(f.getRange('B7'));
  safeClearContent_(f.getRange('B8'));
  safeSetNumberFormat_(f.getRange('B8'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B9'));
  safeClearContent_(f.getRange('B11'));
  safeClearContent_(f.getRange('B13'));
  safeClearContent_(f.getRange('B15'));
  safeClearContent_(f.getRange('B17'));
  safeClearContent_(f.getRange('B18'));
  safeSetNumberFormat_(f.getRange('B18'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B19'));
  safeSetNumberFormat_(f.getRange('B19'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B20'));
}

function clearForm4_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('InventoryAuditForm'); if (!f) return;
  safeClearContent_(f.getRange('B4'));
  safeSetValue_(f.getRange('B6'), 'Auto-set on submit');
  safeClearContent_(f.getRange('B7'));
  safeSetNumberFormat_(f.getRange('B7'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B8'));
  safeClearContent_(f.getRange('B11:B15'));
  safeClearContent_(f.getRange('B17'));
}

function clearForm5_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RadioReturnForm'); if (!f) return;
  safeClearContent_(f.getRange('B4'));
  safeClearContent_(f.getRange('B6'));
  safeClearContent_(f.getRange('B12'));
  safeSetNumberFormat_(f.getRange('B12'), 'yyyy-mm-dd');
  safeClearContent_(f.getRange('B13:B15'));
  safeClearContent_(f.getRange('B17'));
}

function updateAssetConditionFromMaintenance_(at, st) {
  var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ASSET_MASTER');
  if (!m) return;
  var row = findAssetRowByTag_(m, at);
  if (!row) return;
  var c = trimText_(m.getRange(row, 12).getValue()).toUpperCase();
  if (st === 'IN REPAIR') {
    if (['MISSING', 'RETIRED', 'AWAITING PARTS', 'RENEWAL'].indexOf(c) < 0) safeSetValue_(m.getRange(row, 12), 'Repair');
  } else { if (c === '' || c === 'REPAIR') safeSetValue_(m.getRange(row, 12), 'Good'); }
}

function findAssetRowByTag_(m, at) {
  var t = trimText_(at).toUpperCase(); if (!t) return 0;
  var lr = m.getLastRow(); if (lr < 2) return 0;
  var v = m.getRange(2, 2, lr - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (trimText_(v[i][0]).toUpperCase() === t) return i + 2;
  return 0;
}

function reorderSheets_(o) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < o.length; i++) {
    var sh = ss.getSheetByName(o[i]); if (!sh) sh = ss.insertSheet(o[i]);
    try { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); } catch (e) {}
  }
}

function getOrCreateSheet_(ss, n) { var s = ss.getSheetByName(n); if (!s) s = ss.insertSheet(n); return s; }

function styleFormTitle_(s, t) {
  try { s.getRange('A1:F1').merge(); } catch (e) {}
  try { s.getRange('A1:F1').setValue(t).setBackground(THEME_COLOR).setFontColor(HEADER_TEXT_COLOR).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center'); } catch (e) {}
  try { s.setRowHeight(1, 32); } catch (e) {}
}

function setFormLabel_(s, r, t) { try { s.getRange(r, 1).setValue(t); } catch (e) {} }

function mergeA1Ranges_(s, ranges) {
  for (var i = 0; i < ranges.length; i++) try { s.getRange(ranges[i]).breakApart(); } catch (e) {}
  for (var j = 0; j < ranges.length; j++) try { s.getRange(ranges[j]).merge(); } catch (e) {}
}

function applyAssetTagValidation_(t, c) {
  if (!t) return;
  var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ASSET_MASTER');
  if (!m) return;
  var ar = m.getRange(2, 2, Math.max(m.getMaxRows() - 1, 1), 1);
  var v = SpreadsheetApp.newDataValidation().requireValueInRange(ar, true).setAllowInvalid(true).build();
  safeSetDataValidation_(t.getRange(2, c, Math.max(t.getMaxRows() - 1, 1), 1), v);
}

function buildDashboardCharts_(s) {
  var ch = s.getCharts();
  for (var i = 0; i < ch.length; i++) try { s.removeChart(ch[i]); } catch (e) {}
  try {
    s.insertChart(s.newChart().setChartType(Charts.ChartType.PIE).addRange(s.getRange('A3:B8')).setNumHeaders(0)
      .setOption('title', 'Radios by Status').setOption('legend', { position: 'right' }).setPosition(12, 1, 0, 0).build());
  } catch (e) {}
  try {
    s.insertChart(s.newChart().setChartType(Charts.ChartType.COLUMN).addRange(s.getRange('D1:E200')).setNumHeaders(1)
      .setOption('title', 'Radios by Unit').setOption('legend', { position: 'none' }).setPosition(12, 4, 0, 0).build());
  } catch (e) {}
  try {
    s.insertChart(s.newChart().setChartType(Charts.ChartType.LINE).addRange(s.getRange('G1:H200')).setNumHeaders(1)
      .setOption('title', 'Repairs per Month').setOption('legend', { position: 'none' }).setPosition(12, 7, 0, 0).build());
  } catch (e) {}
}

function setMasterDerivedFormulasAtRow_(s, r) {
  if (r < 2) return;
  safeSetFormula_(s.getRange(r, 1), '=IF($B' + r + '="","",ROW()-1)');
  safeSetFormula_(s.getRange(r, 9),
    '=IF($B' + r + '="","",IF($L' + r + '="Missing","MISSING",IF($L' + r + '="Retired","RETIRED",IF($L' + r + '="Awaiting Parts","AWAITING PARTS",IF($L' + r + '="Renewal","RENEWAL",IF($L' + r + '="Repair","UNDER REPAIR",IF($G' + r + '="","AVAILABLE","ASSIGNED"))))))))');
  safeSetFormula_(s.getRange(r, 10),
    '=IF($B' + r + '="","",IFERROR(IF(MAXIFS(MAINTENANCE_LOG!$K:$K,MAINTENANCE_LOG!$C:$C,$B' + r + ')=0,"",MAXIFS(MAINTENANCE_LOG!$K:$K,MAINTENANCE_LOG!$C:$C,$B' + r + ')),""))');
  safeSetFormula_(s.getRange(r, 11),
    '=IF($B' + r + '="","",IFERROR(IF(MAXIFS(INVENTORY_AUDIT!$B:$B,INVENTORY_AUDIT!$C:$C,$B' + r + ')=0,"",MAXIFS(INVENTORY_AUDIT!$B:$B,INVENTORY_AUDIT!$C:$C,$B' + r + ')),""))');
}

function setIssuanceDerivedFormulasAtRow_(s, r) {
  if (r < 2) return;
  safeSetFormula_(s.getRange(r, 1), '=IF($B' + r + '="","",ROW()-1)');
  safeSetFormula_(s.getRange(r, 4), '=IF($C' + r + '="","",IFERROR(VLOOKUP($C' + r + ',ASSET_MASTER!$B:$C,2,FALSE),""))');
  safeSetFormula_(s.getRange(r, 11), '=IF($C' + r + '="","",IF(ISBLANK($J' + r + '),"ISSUED","RETURNED"))');
}

function setMaintenanceDerivedFormulasAtRow_(s, r) {
  if (r < 2) return;
  safeSetFormula_(s.getRange(r, 1), '=IF($B' + r + '="","",ROW()-1)');
  safeSetFormula_(s.getRange(r, 12), '=IF($C' + r + '="","",IF(ISBLANK($K' + r + '),"IN REPAIR","COMPLETED"))');
}

function nextId_(s, c, p, d) {
  var mr = s.getMaxRows(), max = 0;
  var re = new RegExp('^' + escapeRegExp_(p) + '-(\\d+)$');
  if (mr >= 2) {
    var v = s.getRange(2, c, mr - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      var t = trimText_(v[i][0]); if (!t) continue;
      var m = t.match(re); if (m) { var n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
    }
  }
  return p + '-' + padNumber_(max + 1, d || 5);
}

function nextAssetTag_(m) {
  var v = m.getRange(2, 2, Math.max(m.getMaxRows() - 1, 1), 1).getValues(), max = 0;
  var re = /^(?:COM-)?RAD-(\d+)$/i;
  for (var i = 0; i < v.length; i++) {
    var t = trimText_(v[i][0]); if (!t) continue;
    var mt = t.match(re); if (mt) { var n = parseInt(mt[1], 10); if (!isNaN(n) && n > max) max = n; }
  }
  return 'RAD-' + padNumber_(max + 1, 3);
}

function ensureRowsAtLeast_(s, n) {
  var c = s.getMaxRows();
  if (c < n) try { s.insertRowsAfter(c, n - c); } catch (e) {}
}
function ensureColumns_(s, n) { var c = s.getMaxColumns(); if (c < n) try { s.insertColumnsAfter(c, n - c); } catch (e) {} }

function firstBlankRow_(s, c, sr) {
  var mr = s.getMaxRows();
  var v = s.getRange(sr, c, mr - sr + 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === '' || v[i][0] === null) return sr + i;
  return mr + 1;
}

function padNumber_(n, d) { var s = String(n); while (s.length < d) s = '0' + s; return s; }
function trimText_(v) { if (v === null || v === undefined) return ''; return String(v).trim(); }
function isBlankValue_(v) { return v === '' || v === null || v === undefined; }
function escapeRegExp_(t) { return String(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function ensureActiveForm_(n) {
  var a = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  if (a !== n) { SpreadsheetApp.getUi().alert('This button is for ' + n + ' only.'); return false; }
  return true;
}

function btnFindRowByValue_(s, c, n) {
  if (!s) return 0; var lr = s.getLastRow(); if (lr < 2) return 0;
  var k = trimText_(n).toUpperCase(); if (!k) return 0;
  var v = s.getRange(2, c, lr - 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (trimText_(v[i][0]).toUpperCase() === k) return i + 2;
  return 0;
}

function btnFindLastRowByValue_(s, c, n) {
  if (!s) return 0; var lr = s.getLastRow(); if (lr < 2) return 0;
  var k = trimText_(n).toUpperCase(); if (!k) return 0;
  var v = s.getRange(2, c, lr - 1, 1).getDisplayValues();
  for (var i = v.length - 1; i >= 0; i--) if (trimText_(v[i][0]).toUpperCase() === k) return i + 2;
  return 0;
}

function btnSetDateOrBlank_(r, v) {
  if (isBlankValue_(v) || v === 0) { safeClearContent_(r); safeSetNumberFormat_(r, 'yyyy-mm-dd'); return; }
  safeSetValue_(r, v); safeSetNumberFormat_(r, 'yyyy-mm-dd');
}
