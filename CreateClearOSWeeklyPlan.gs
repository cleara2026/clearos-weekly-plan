/**
 * CLEAR OS Weekly Plan — Google Apps Script
 *
 * Run createClearOSWeeklyPlan() from the Apps Script editor to build
 * a brand-new CLEAR OS Weekly Plan spreadsheet in your Google Drive.
 *
 * Script is standalone (script.google.com). No bound spreadsheet needed.
 */

// ─── Brand colours ───────────────────────────────────────────────────────────
var COLOUR = {
  teal:          '#009a98',
  purple:        '#652d90',
  yellow:        '#edc530',
  orange:        '#f6931e',
  red:           '#eb1c24',
  lightTeal:     '#e6f7f7',
  lightPurple:   '#f0e8f7',
  lightOrange:   '#fff3e0',
  lightRed:      '#fce4ec',
  white:         '#ffffff',
  offWhite:      '#f9f9f9',
  lightGrey:     '#f2f2f2',
  midGrey:       '#cccccc',
  darkGrey:      '#444444',
  black:         '#000000',
};

// ─── Main entry point ─────────────────────────────────────────────────────────
function createClearOSWeeklyPlan() {
  var ss = SpreadsheetApp.create('CLEAR OS Weekly Plan');

  // Build tabs in order (the default Sheet1 will be renamed to Instructions)
  var instructionsSheet = ss.getSheets()[0];
  instructionsSheet.setName('Instructions');

  var listSheet      = ss.insertSheet('1. List');
  var estimateSheet  = ss.insertSheet('2. Estimate');
  var urgentSheet    = ss.insertSheet('3. Urgent');
  var highSheet      = ss.insertSheet('4. High');
  var capacitySheet  = ss.insertSheet('5. Capacity');
  var allocateSheet  = ss.insertSheet('6. Allocate');
  var roadmapSheet   = ss.insertSheet('7. Roadmap');
  var dataSheet      = ss.insertSheet('Data');

  // Build each tab
  buildDataTab(dataSheet);
  buildInstructionsTab(instructionsSheet, ss.getUrl());
  buildListTab(listSheet);
  buildEstimateTab(estimateSheet);
  buildUrgentTab(urgentSheet);
  buildHighTab(highSheet);
  buildCapacityTab(capacitySheet);
  buildAllocateTab(allocateSheet);
  buildRoadmapTab(roadmapSheet);

  // Hide the Data tab
  dataSheet.hideSheet();

  // Reorder so Instructions is first
  ss.setActiveSheet(instructionsSheet);
  ss.moveActiveSheet(1);

  // Open the file for the user
  var url = ss.getUrl();

  Logger.log('Your CLEAR OS Weekly Plan is ready! URL: ' + url);
}

// ─── DATA TAB ────────────────────────────────────────────────────────────────
function buildDataTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();

  // Columns: A=ID | B=Task | C=Minutes | D=Priority | E=Day | F=Status | G=Notes
  var headers = ['ID', 'Task', 'Minutes', 'Priority', 'Day', 'Status', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  styleHeaderRow(sheet, 1, headers.length, COLOUR.darkGrey, COLOUR.lightGrey);

  // Column widths
  sheet.setColumnWidth(1, 60);   // ID
  sheet.setColumnWidth(2, 300);  // Task
  sheet.setColumnWidth(3, 80);   // Minutes
  sheet.setColumnWidth(4, 100);  // Priority
  sheet.setColumnWidth(5, 110);  // Day
  sheet.setColumnWidth(6, 110);  // Status
  sheet.setColumnWidth(7, 240);  // Notes

  // Pre-fill 50 rows with IDs and empty task data
  for (var i = 2; i <= 51; i++) {
    sheet.getRange(i, 1).setValue(i - 1); // ID 1–50
  }

  // Data validation — Priority
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Urgent', 'High', 'Normal'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 50, 1).setDataValidation(priorityRule);

  // Data validation — Day
  var dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Unassigned'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 50, 1).setDataValidation(dayRule);

  // Data validation — Status
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['To Do', 'Done', 'Delegated', 'Dropped'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, 50, 1).setDataValidation(statusRule);

  // Default values for new rows
  sheet.getRange(2, 4, 50, 1).setValue('Normal');
  sheet.getRange(2, 5, 50, 1).setValue('Unassigned');
  sheet.getRange(2, 6, 50, 1).setValue('To Do');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Font
  sheet.getRange(1, 1, 51, headers.length).setFontFamily('Arial').setFontSize(10);
}

// ─── INSTRUCTIONS TAB ────────────────────────────────────────────────────────
function buildInstructionsTab(sheet, ssUrl) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.teal);

  // Set column widths
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 500);
  sheet.setColumnWidth(3, 300);

  // ── Brand header block ──
  sheet.getRange('B1:C3').merge();
  sheet.getRange('B1').setValue('CLEAR OS');
  sheet.getRange('B1').setFontFamily('Arial')
    .setFontSize(28)
    .setFontWeight('bold')
    .setFontColor(COLOUR.white)
    .setBackground(COLOUR.teal)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  sheet.getRange('B4:C4').merge();
  sheet.getRange('B4').setValue('Weekly Plan')
    .setFontFamily('Arial')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor(COLOUR.teal)
    .setBackground(COLOUR.lightTeal)
    .setHorizontalAlignment('left');

  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 40);
  sheet.setRowHeight(3, 40);
  sheet.setRowHeight(4, 30);

  // ── Tagline ──
  sheet.getRange('B5').setValue(
    'Clarity over chaos. One week at a time.'
  ).setFontFamily('Arial').setFontSize(11).setFontColor(COLOUR.darkGrey).setFontStyle('italic');
  sheet.setRowHeight(5, 28);

  var gap = sheet.getRange('B6');
  gap.setValue('').setBackground(COLOUR.white);
  sheet.setRowHeight(6, 16);

  // ── How to use ──
  var instructions = [
    ['', ''],
    ['HOW TO USE THIS TOOL', ''],
    ['', ''],
    ['Each tab walks you through the CLEAR weekly planning process.', ''],
    ['Work through the steps in order, starting with Step 1 (List).', ''],
    ['', ''],
    ['Step 1 — List',    'Capture every task on your mind. Don\'t filter yet.'],
    ['Step 2 — Estimate', 'Add honest time estimates (in minutes) to each task.'],
    ['Step 3 — Urgent',  'Identify what MUST happen by this time next week.'],
    ['Step 4 — High',    'Mark tasks you\'d like to get done if capacity allows.'],
    ['Step 5 — Capacity', 'Enter how many hours you have available each day.'],
    ['Step 6 — Allocate', 'Assign your Urgent and High tasks to specific days.'],
    ['Step 7 — Roadmap', 'Review your committed week at a glance.'],
  ];

  var startRow = 7;
  for (var i = 0; i < instructions.length; i++) {
    var row = startRow + i;
    var label = instructions[i][0];
    var desc  = instructions[i][1];

    if (label === 'HOW TO USE THIS TOOL') {
      sheet.getRange(row, 2).setValue(label)
        .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
        .setFontColor(COLOUR.teal);
      sheet.setRowHeight(row, 24);
    } else if (label !== '' && desc !== '') {
      sheet.getRange(row, 2).setValue(label)
        .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
        .setFontColor(COLOUR.darkGrey);
      sheet.getRange(row, 3).setValue(desc)
        .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey);
      sheet.setRowHeight(row, 22);
    } else if (label !== '') {
      sheet.getRange(row, 2).setValue(label)
        .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey);
      sheet.setRowHeight(row, 20);
    } else {
      sheet.setRowHeight(row, 10);
    }
  }

  var tipRow = startRow + instructions.length + 1;
  sheet.getRange(tipRow, 2, 1, 2).merge();
  sheet.getRange(tipRow, 2).setValue(
    '💡  Make a fresh copy every Monday morning. File → Make a copy'
  ).setFontFamily('Arial').setFontSize(10)
    .setFontColor(COLOUR.white)
    .setBackground(COLOUR.teal)
    .setFontWeight('bold');
  sheet.setRowHeight(tipRow, 28);

  sheet.getRange(tipRow + 1, 2, 1, 2).merge();
  sheet.getRange(tipRow + 1, 2).setValue(
    '📌  Bookmark the original and always copy from it, never use it directly.'
  ).setFontFamily('Arial').setFontSize(10)
    .setFontColor(COLOUR.teal)
    .setBackground(COLOUR.lightTeal);
  sheet.setRowHeight(tipRow + 1, 28);
}

// ─── 1. LIST TAB ─────────────────────────────────────────────────────────────
function buildListTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.teal);

  // Column widths
  sheet.setColumnWidth(1, 40);   // row number spacer
  sheet.setColumnWidth(2, 40);   // #
  sheet.setColumnWidth(3, 320);  // Task
  sheet.setColumnWidth(4, 90);   // Minutes
  sheet.setColumnWidth(5, 110);  // Priority
  sheet.setColumnWidth(6, 240);  // Notes

  // ── Tab header ──
  buildTabHeader(sheet, '1. List', COLOUR.teal, 'Capture everything on your mind for the week. Don\'t filter yet — just get it all out.');

  // ── Column headers (row 4) ──
  var colHeaders = ['#', 'Task', 'Minutes', 'Priority', 'Notes'];
  sheet.getRange(4, 2, 1, colHeaders.length).setValues([colHeaders]);
  styleHeaderRow(sheet, 4, colHeaders.length, COLOUR.white, COLOUR.teal, 2);

  // ── 50 input rows ──
  for (var i = 1; i <= 50; i++) {
    var row = 4 + i;
    // Row number
    sheet.getRange(row, 2).setValue(i)
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.lightGrey);

    // Task (editable)
    styleInputCell(sheet.getRange(row, 3));

    // Minutes (read hint)
    sheet.getRange(row, 4)
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Priority (read hint)
    sheet.getRange(row, 5)
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Notes (editable)
    styleInputCell(sheet.getRange(row, 6));

    sheet.setRowHeight(row, 24);
  }

  // ── Instruction cell below ──
  var bottomRow = 55;
  sheet.getRange(bottomRow, 3, 1, 3).merge();
  sheet.getRange(bottomRow, 3).setValue('Fill in Task and Notes here. You\'ll add Minutes and Priority in the next steps.')
    .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey).setFontStyle('italic');

  // Freeze rows 1–4
  sheet.setFrozenRows(4);

  // Note about Data tab sync
  sheet.getRange(3, 3).setValue(
    '⬇  Enter tasks below. This list feeds all subsequent steps.'
  ).setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.teal);

  // Set a note on task column header
  sheet.getRange(4, 3).setNote(
    'Type your tasks here. Each task is stored and referenced throughout the planner.'
  );

  addStepFooter(sheet, 56, 3, 'Next: go to "2. Estimate" to add time estimates to each task.');
}

// ─── 2. ESTIMATE TAB ─────────────────────────────────────────────────────────
function buildEstimateTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.teal);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);

  buildTabHeader(sheet, '2. Estimate', COLOUR.teal, 'How long will each task actually take? Be honest. Underestimating is the #1 planning mistake.');

  var colHeaders = ['#', 'Task', 'Minutes', 'Hours'];
  sheet.getRange(4, 2, 1, colHeaders.length).setValues([colHeaders]);
  styleHeaderRow(sheet, 4, colHeaders.length, COLOUR.white, COLOUR.teal, 2);

  // 50 rows referencing the List tab
  for (var i = 1; i <= 50; i++) {
    var row = 4 + i;

    // Row number
    sheet.getRange(row, 2).setValue(i)
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center').setBackground(COLOUR.lightGrey);

    // Task — pulled from '1. List' tab column C
    sheet.getRange(row, 3)
      .setFormula("='1. List'!C" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Minutes — editable input
    styleInputCell(sheet.getRange(row, 4));
    sheet.getRange(row, 4).setHorizontalAlignment('center');

    // Hours — formula
    sheet.getRange(row, 5)
      .setFormula('=IF(D' + row + '="","",ROUND(D' + row + '/60,1))')
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.lightTeal)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(row, 24);
  }

  // Totals row
  var totRow = 55;
  sheet.getRange(totRow, 3).setValue('TOTAL')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOUR.teal).setHorizontalAlignment('right');
  sheet.getRange(totRow, 4)
    .setFormula('=SUM(D5:D54)')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOUR.teal).setHorizontalAlignment('center')
    .setBackground(COLOUR.lightTeal)
    .setBorder(true, true, true, true, false, false, COLOUR.teal, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totRow, 5)
    .setFormula('=ROUND(D55/60,1)')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOUR.teal).setHorizontalAlignment('center')
    .setBackground(COLOUR.lightTeal)
    .setBorder(true, true, true, true, false, false, COLOUR.teal, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(totRow + 1, 3).setValue('hours')
    .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
    .setHorizontalAlignment('right');

  sheet.setFrozenRows(4);

  // Instruction note
  sheet.getRange(3, 3).setValue('⬇  Enter minutes for each task. The Hours column calculates automatically.')
    .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.teal);

  addStepFooter(sheet, 57, 3, 'Next: go to "3. Urgent" to identify what must be done this week.');
}

// ─── 3. URGENT TAB ───────────────────────────────────────────────────────────
function buildUrgentTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.red);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 90);

  buildTabHeader(sheet, '3. Urgent', COLOUR.red,
    'What MUST be done by this time next week — or you\'re in trouble? Be strict. Most things are not truly urgent.');

  // Warning formula row
  sheet.getRange(3, 2, 1, 4).merge();
  sheet.getRange(3, 2)
    .setFormula(
      '=IF(COUNTIF(F5:F54,"Urgent")>3,' +
      '"⚠  You have "&COUNTIF(F5:F54,"Urgent")&" urgent tasks. Urgent means it HAS to be done by this time next week. Or you\'re in trouble. Cut the list.",' +
      'IF(COUNTIF(F5:F54,"Urgent")=0,"Mark tasks as Urgent using the dropdown in the Priority column →",' +
      '"✓  "&COUNTIF(F5:F54,"Urgent")&" urgent task(s) marked. Keep it tight."))'
    )
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // We'll conditionally format this cell: red bg if >3 urgent, teal bg otherwise
  // Use conditional formatting rules
  var warningRange = sheet.getRange('B3');
  var rules = sheet.getConditionalFormatRules();

  // Rule: red background when count > 3
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COUNTIF($F$5:$F$54,"Urgent")>3')
    .setBackground(COLOUR.lightRed)
    .setFontColor(COLOUR.red)
    .setRanges([warningRange])
    .build();

  // Rule: teal background when count <= 3 and > 0
  var okRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(COUNTIF($F$5:$F$54,"Urgent")<=3,COUNTIF($F$5:$F$54,"Urgent")>0)')
    .setBackground(COLOUR.lightTeal)
    .setFontColor(COLOUR.teal)
    .setRanges([warningRange])
    .build();

  rules.push(redRule);
  rules.push(okRule);
  sheet.setConditionalFormatRules(rules);
  sheet.setRowHeight(3, 32);

  // Column headers row 4
  var colHeaders = ['#', 'Task', 'Minutes', 'Priority', 'Mark Urgent?'];
  sheet.getRange(4, 2, 1, colHeaders.length).setValues([colHeaders]);
  styleHeaderRow(sheet, 4, colHeaders.length, COLOUR.white, COLOUR.red, 2);

  // Priority data validation
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Urgent', 'High', 'Normal'], true)
    .setAllowInvalid(false)
    .build();

  for (var i = 1; i <= 50; i++) {
    var row = 4 + i;

    sheet.getRange(row, 2).setValue(i)
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center').setBackground(COLOUR.lightGrey);

    // Task — from List tab
    sheet.getRange(row, 3)
      .setFormula("='1. List'!C" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Minutes — from Estimate tab
    sheet.getRange(row, 4)
      .setFormula("='2. Estimate'!D" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Priority — editable dropdown
    sheet.getRange(row, 5).setDataValidation(priorityRule)
      .setFontFamily('Arial').setFontSize(10)
      .setHorizontalAlignment('center');
    styleInputCell(sheet.getRange(row, 5));

    // "Mark Urgent?" helper — instruction text in header, checkbox concept
    // Show a visual cue: formula shows "← set to Urgent" if not already urgent
    sheet.getRange(row, 6)
      .setFormula('=IF(C' + row + '="","",IF(E' + row + '="Urgent","✓ Urgent","← set Urgent?"))')
      .setFontFamily('Arial').setFontSize(9).setFontStyle('italic')
      .setHorizontalAlignment('center');

    sheet.setRowHeight(row, 24);
  }

  // Conditional formatting: highlight entire row if Priority = Urgent
  var dataRange = sheet.getRange('B5:F54');
  var urgentRowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E5="Urgent"')
    .setBackground(COLOUR.lightRed)
    .setFontColor(COLOUR.red)
    .setRanges([dataRange])
    .build();

  var currentRules = sheet.getConditionalFormatRules();
  currentRules.push(urgentRowRule);
  sheet.setConditionalFormatRules(currentRules);

  sheet.setFrozenRows(4);
  addStepFooter(sheet, 56, 3, 'Next: go to "4. High" to mark your high-priority tasks.');
}

// ─── 4. HIGH TAB ─────────────────────────────────────────────────────────────
function buildHighTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.purple);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 90);

  buildTabHeader(sheet, '4. High', COLOUR.purple,
    'Which tasks would you love to get done this week, if capacity allows? These are important but not crisis-level.');

  sheet.getRange(3, 2, 1, 4).merge();
  sheet.getRange(3, 2)
    .setFormula(
      '=IF(COUNTIF(E5:E54,"High")=0,"Mark tasks as High priority using the dropdown →",' +
      '"✓  "&COUNTIF(E5:E54,"High")&" high-priority task(s) identified.")'
    )
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOUR.purple).setBackground(COLOUR.lightPurple)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(3, 28);

  var colHeaders = ['#', 'Task', 'Minutes', 'Priority', 'Mark High?'];
  sheet.getRange(4, 2, 1, colHeaders.length).setValues([colHeaders]);
  styleHeaderRow(sheet, 4, colHeaders.length, COLOUR.white, COLOUR.purple, 2);

  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Urgent', 'High', 'Normal'], true)
    .setAllowInvalid(false)
    .build();

  for (var i = 1; i <= 50; i++) {
    var row = 4 + i;

    sheet.getRange(row, 2).setValue(i)
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center').setBackground(COLOUR.lightGrey);

    sheet.getRange(row, 3)
      .setFormula("='1. List'!C" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(row, 4)
      .setFormula("='2. Estimate'!D" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Priority editable
    sheet.getRange(row, 5).setDataValidation(priorityRule);
    styleInputCell(sheet.getRange(row, 5));
    sheet.getRange(row, 5).setHorizontalAlignment('center');

    // Helper hint
    sheet.getRange(row, 6)
      .setFormula('=IF(C' + row + '="","",IF(E' + row + '="High","✓ High",IF(E' + row + '="Urgent","(Urgent)","← set High?")))')
      .setFontFamily('Arial').setFontSize(9).setFontStyle('italic')
      .setHorizontalAlignment('center');

    sheet.setRowHeight(row, 24);
  }

  // Highlight high rows purple
  var dataRange = sheet.getRange('B5:F54');
  var highRowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E5="High"')
    .setBackground(COLOUR.lightPurple)
    .setFontColor(COLOUR.purple)
    .setRanges([dataRange])
    .build();

  var currentRules = sheet.getConditionalFormatRules();
  currentRules.push(highRowRule);
  sheet.setConditionalFormatRules(currentRules);

  sheet.setFrozenRows(4);
  addStepFooter(sheet, 56, 3, 'Next: go to "5. Capacity" to enter your available hours for the week.');
}

// ─── 5. CAPACITY TAB ─────────────────────────────────────────────────────────
function buildCapacityTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.orange);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 200);

  buildTabHeader(sheet, '5. Capacity', COLOUR.orange,
    'How many hours do you actually have available each day? Enter them below. The sheet will calculate your real capacity.');

  // ── Day grid ──
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  var gridStartRow = 5;

  // Sub-header
  var gridHeaders = ['Day', 'Hours Available', 'Realistic Capacity (hrs)', 'Minutes'];
  sheet.getRange(gridStartRow, 2, 1, gridHeaders.length).setValues([gridHeaders]);
  styleHeaderRow(sheet, gridStartRow, gridHeaders.length, COLOUR.white, COLOUR.orange, 2);

  for (var d = 0; d < days.length; d++) {
    var row = gridStartRow + 1 + d;

    sheet.getRange(row, 2).setValue(days[d])
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
      .setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.lightOrange)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Hours available — editable
    styleInputCell(sheet.getRange(row, 3));
    sheet.getRange(row, 3).setHorizontalAlignment('center').setValue(d < 5 ? 8 : 0);

    // Realistic capacity = hours × 0.5
    sheet.getRange(row, 4)
      .setFormula('=IF(C' + row + '="","",C' + row + '*0.5)')
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.lightTeal)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Minutes
    sheet.getRange(row, 5)
      .setFormula('=IF(D' + row + '="","",D' + row + '*60)')
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.lightTeal)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(row, 26);
  }

  // ── Totals ──
  var totRow = gridStartRow + days.length + 1;
  sheet.getRange(totRow, 2).setValue('TOTAL WEEK')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.orange)
    .setBackground(COLOUR.lightOrange)
    .setBorder(true, true, true, true, false, false, COLOUR.orange, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totRow, 3)
    .setFormula('=SUM(C' + (gridStartRow+1) + ':C' + (gridStartRow+days.length) + ')')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.orange)
    .setHorizontalAlignment('center')
    .setBackground(COLOUR.lightOrange)
    .setBorder(true, true, true, true, false, false, COLOUR.orange, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totRow, 4)
    .setFormula('=SUM(D' + (gridStartRow+1) + ':D' + (gridStartRow+days.length) + ')')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.orange)
    .setHorizontalAlignment('center')
    .setBackground(COLOUR.lightTeal)
    .setBorder(true, true, true, true, false, false, COLOUR.orange, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totRow, 5)
    .setFormula('=SUM(E' + (gridStartRow+1) + ':E' + (gridStartRow+days.length) + ')')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.orange)
    .setHorizontalAlignment('center')
    .setBackground(COLOUR.lightTeal)
    .setBorder(true, true, true, true, false, false, COLOUR.orange, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(totRow, 28);

  // ── Reality Delta ──
  var deltaStartRow = totRow + 2;

  sheet.getRange(deltaStartRow, 2, 1, 3).merge();
  sheet.getRange(deltaStartRow, 2).setValue('REALITY DELTA')
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor(COLOUR.orange);
  sheet.setRowHeight(deltaStartRow, 28);

  // Total task minutes
  var estimateTotRow = deltaStartRow + 1;
  sheet.getRange(estimateTotRow, 2).setValue('Total task minutes (from Estimate):')
    .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey);
  sheet.getRange(estimateTotRow, 3)
    .setFormula("='2. Estimate'!D55")
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(COLOUR.offWhite)
    .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

  // Realistic capacity minutes
  var capRow = estimateTotRow + 1;
  sheet.getRange(capRow, 2).setValue('Realistic capacity (minutes):')
    .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey);
  sheet.getRange(capRow, 3)
    .setFormula('=E' + totRow)
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(COLOUR.lightTeal)
    .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

  // Delta
  var deltaRow = capRow + 1;
  sheet.getRange(deltaRow, 2).setValue('Delta (capacity minus tasks):')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.darkGrey);
  sheet.getRange(deltaRow, 3)
    .setFormula('=C' + capRow + '-C' + estimateTotRow)
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Conditional formatting on delta cell
  var deltaCell = sheet.getRange(deltaRow, 3);
  var deltaRules = sheet.getConditionalFormatRules();

  var greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(30)
    .setBackground('#d9ead3')
    .setFontColor('#274e13')
    .setRanges([deltaCell])
    .build();
  var amberRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(-30, 30)
    .setBackground('#fff2cc')
    .setFontColor('#7f6000')
    .setRanges([deltaCell])
    .build();
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-30)
    .setBackground(COLOUR.lightRed)
    .setFontColor(COLOUR.red)
    .setRanges([deltaCell])
    .build();

  deltaRules.push(greenRule);
  deltaRules.push(amberRule);
  deltaRules.push(redRule);
  sheet.setConditionalFormatRules(deltaRules);

  // Status label
  var statusRow = deltaRow + 1;
  sheet.getRange(statusRow, 2, 1, 2).merge();
  sheet.getRange(statusRow, 2)
    .setFormula(
      '=IF(C' + deltaRow + '>30,"✅  You have capacity. Assign more tasks if needed.",' +
      'IF(C' + deltaRow + '>=(-30),"⚠  You\'re close to capacity. Tread carefully.",' +
      '"🔴  Overloaded. Cut tasks until the delta is positive."))'
    )
    .setFontFamily('Arial').setFontSize(10).setFontStyle('italic').setFontColor(COLOUR.darkGrey);
  sheet.setRowHeight(statusRow, 24);

  // Note about realistic capacity
  var noteRow = statusRow + 2;
  sheet.getRange(noteRow, 2, 1, 3).merge();
  sheet.getRange(noteRow, 2)
    .setValue('Note: "Realistic capacity" = your available hours × 0.5. This accounts for context switching, emails, and the fact that focused work rarely fills the whole day.')
    .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey).setFontStyle('italic')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(noteRow, 36);

  sheet.setFrozenRows(4);
  addStepFooter(sheet, noteRow + 2, 2, 'Next: go to "6. Allocate" to assign your tasks to specific days.');
}

// ─── 6. ALLOCATE TAB ─────────────────────────────────────────────────────────
function buildAllocateTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.teal);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);

  buildTabHeader(sheet, '6. Allocate', COLOUR.teal,
    'Assign your Urgent and High priority tasks to specific days. Be realistic — use the capacity you calculated in Step 5.');

  // Instructions row 3
  sheet.getRange(3, 2, 1, 5).merge();
  sheet.getRange(3, 2).setValue(
    'This view shows all tasks. Set the Day column for each task you want to schedule. Tasks marked Urgent or High are highlighted.'
  ).setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.teal).setFontStyle('italic')
   .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(3, 30);

  // Column headers
  var colHeaders = ['#', 'Task', 'Minutes', 'Priority', 'Assign Day', 'Status', 'Day Minutes Used'];
  sheet.getRange(4, 2, 1, colHeaders.length).setValues([colHeaders]);
  styleHeaderRow(sheet, 4, colHeaders.length, COLOUR.white, COLOUR.teal, 2);

  var dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Unassigned'], true)
    .setAllowInvalid(false)
    .build();

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['To Do', 'Done', 'Delegated', 'Dropped'], true)
    .setAllowInvalid(false)
    .build();

  for (var i = 1; i <= 50; i++) {
    var row = 4 + i;

    sheet.getRange(row, 2).setValue(i)
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey)
      .setHorizontalAlignment('center').setBackground(COLOUR.lightGrey);

    // Task
    sheet.getRange(row, 3)
      .setFormula("='1. List'!C" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Minutes
    sheet.getRange(row, 4)
      .setFormula("='2. Estimate'!D" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Priority (read from Urgent/High tabs — use Urgent tab as master since it aggregates both)
    sheet.getRange(row, 5)
      .setFormula("='3. Urgent'!E" + (4 + i))
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    // Day — editable
    sheet.getRange(row, 6).setDataValidation(dayRule).setValue('Unassigned');
    styleInputCell(sheet.getRange(row, 6));
    sheet.getRange(row, 6).setHorizontalAlignment('center');

    // Status — editable
    sheet.getRange(row, 7).setDataValidation(statusRule).setValue('To Do');
    styleInputCell(sheet.getRange(row, 7));
    sheet.getRange(row, 7).setHorizontalAlignment('center');

    // Day minutes used (cumulative for that day)
    // This shows total minutes allocated to the same day as this task
    sheet.getRange(row, 8)
      .setFormula('=IF(C' + row + '="","",IF(F' + row + '="Unassigned","",SUMIF($F$5:$F$54,F' + row + ',$D$5:$D$54)))')
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center')
      .setBackground(COLOUR.lightTeal)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(row, 24);
  }

  // Conditional formatting — urgent = light red, high = light purple
  var dataRange = sheet.getRange('B5:H54');
  var rules = sheet.getConditionalFormatRules();

  var urgentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E5="Urgent"')
    .setBackground(COLOUR.lightRed)
    .setRanges([dataRange])
    .build();

  var highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E5="High"')
    .setBackground(COLOUR.lightPurple)
    .setRanges([dataRange])
    .build();

  var doneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G5="Done"')
    .setFontColor(COLOUR.midGrey)
    .setRanges([dataRange])
    .build();

  rules.push(urgentRule);
  rules.push(highRule);
  rules.push(doneRule);
  sheet.setConditionalFormatRules(rules);

  // ── Per-day summary below the task rows ──
  var sumStartRow = 56;
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  sheet.getRange(sumStartRow, 2, 1, 5).merge();
  sheet.getRange(sumStartRow, 2).setValue('DAY SUMMARY')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.white)
    .setBackground(COLOUR.teal);
  sheet.setRowHeight(sumStartRow, 24);

  var sumHeaders = ['Day', 'Tasks', 'Minutes Allocated', 'Realistic Capacity (min)', 'Remaining'];
  sheet.getRange(sumStartRow + 1, 2, 1, sumHeaders.length).setValues([sumHeaders]);
  styleHeaderRow(sheet, sumStartRow + 1, sumHeaders.length, COLOUR.white, COLOUR.darkGrey, 2);

  var capSheet = '5. Capacity';
  // Capacity rows in the 5. Capacity sheet are rows 6–12 (Mon=6, Tue=7, ... Sun=12)
  var capMinCols = ['E6', 'E7', 'E8', 'E9', 'E10', 'E11', 'E12'];

  for (var d = 0; d < days.length; d++) {
    var sRow = sumStartRow + 2 + d;
    sheet.getRange(sRow, 2).setValue(days[d])
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setBackground(COLOUR.lightGrey)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(sRow, 3)
      .setFormula('=COUNTIF($F$5:$F$54,"' + days[d] + '")')
      .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(sRow, 4)
      .setFormula('=SUMIF($F$5:$F$54,"' + days[d] + '",$D$5:$D$54)')
      .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center')
      .setBackground(COLOUR.offWhite)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(sRow, 5)
      .setFormula("='5. Capacity'!" + capMinCols[d])
      .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center')
      .setBackground(COLOUR.lightTeal)
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(sRow, 6)
      .setFormula('=E' + sRow + '-D' + sRow)
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Conditional: green if positive, red if negative
    var remCell = sheet.getRange(sRow, 6);
    var sumRules = sheet.getConditionalFormatRules();
    sumRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0)
        .setBackground('#d9ead3').setFontColor('#274e13')
        .setRanges([remCell]).build()
    );
    sumRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground(COLOUR.lightRed).setFontColor(COLOUR.red)
        .setRanges([remCell]).build()
    );
    sheet.setConditionalFormatRules(sumRules);

    sheet.setRowHeight(sRow, 24);
  }

  sheet.setFrozenRows(4);
  addStepFooter(sheet, sumStartRow + 2 + days.length + 1, 2, 'Next: go to "7. Roadmap" for your final committed week view.');
}

// ─── 7. ROADMAP TAB ──────────────────────────────────────────────────────────
function buildRoadmapTab(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setTabColor(COLOUR.teal);

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);

  buildTabHeader(sheet, '7. Roadmap', COLOUR.teal, 'Your committed week at a glance. Review this every morning.');

  // ── Summary stats ──
  var statsRow = 4;
  sheet.getRange(statsRow, 2, 1, 4).merge();
  sheet.getRange(statsRow, 2)
    .setFormula(
      '="This week you\'ve committed to "&ROUND(SUMIF(\'6. Allocate\'!$F$5:$F$54,"<>Unassigned",\'6. Allocate\'!$D$5:$D$54)/60,1)' +
      '&" hours of focused work across "&COUNTIF(\'6. Allocate\'!$F$5:$F$54,"<>Unassigned")&" tasks."'
    )
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor(COLOUR.teal).setBackground(COLOUR.lightTeal)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, COLOUR.teal, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(statsRow, 36);

  // Quick stats row
  var q = statsRow + 1;
  var statLabels = [
    ['Urgent tasks:', '=COUNTIF(\'3. Urgent\'!E5:E54,"Urgent")'],
    ['High tasks:', '=COUNTIF(\'4. High\'!E5:E54,"High")'],
    ['Total tasks:', '=COUNTA(\'1. List\'!C5:C54)'],
    ['Total minutes:', '=\'2. Estimate\'!D55'],
  ];

  for (var s = 0; s < statLabels.length; s++) {
    var statCol = 2 + s * 2;
    sheet.getRange(q, statCol).setValue(statLabels[s][0])
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey);
    sheet.getRange(q, statCol + 1)
      .setFormula(statLabels[s][1])
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold').setFontColor(COLOUR.darkGrey)
      .setHorizontalAlignment('center');
  }
  sheet.setRowHeight(q, 20);

  // ── Per-day sections ──
  var days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  var currentRow = q + 2;

  for (var d = 0; d < days.length; d++) {
    var day = days[d];

    // Day header
    sheet.getRange(currentRow, 2, 1, 4).merge();
    sheet.getRange(currentRow, 2).setValue(day.toUpperCase())
      .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setFontColor(COLOUR.white).setBackground(COLOUR.teal)
      .setHorizontalAlignment('left');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;

    // Column mini-headers
    var miniHeaders = ['Task', 'Minutes', 'Priority', 'Status'];
    sheet.getRange(currentRow, 3, 1, miniHeaders.length).setValues([miniHeaders]);
    sheet.getRange(currentRow, 3, 1, miniHeaders.length)
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setFontColor(COLOUR.white).setBackground(COLOUR.darkGrey);
    sheet.setRowHeight(currentRow, 20);
    currentRow++;

    // QUERY formula to pull tasks for this day from Allocate tab
    // We use a QUERY on the Allocate tab columns: C=Task, D=Minutes, E=Priority, F=Day, G=Status
    // Allocate tab data rows 5:54, columns B–H
    // Col indices: B=1, C=2, D=3, E=4(Priority), F=5(Day), G=6(Status)
    sheet.getRange(currentRow, 3)
      .setFormula(
        '=IFERROR(QUERY({\'6. Allocate\'!C5:C54,\'6. Allocate\'!D5:D54,\'6. Allocate\'!E5:E54,\'6. Allocate\'!F5:F54,\'6. Allocate\'!G5:G54},' +
        '"SELECT Col1,Col2,Col3,Col5 WHERE Col4=\'' + day + '\' AND Col1 <> \'\' ORDER BY Col3 DESC",0),' +
        '"(no tasks scheduled)")'
      )
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Reserve 8 rows per day for query results
    for (var r = 0; r < 8; r++) {
      sheet.setRowHeight(currentRow + r, 24);
      sheet.getRange(currentRow + r, 3, 1, 4)
        .setBorder(false, false, true, false, false, false, COLOUR.lightGrey, SpreadsheetApp.BorderStyle.SOLID);
    }

    // Day totals
    var dayTotRow = currentRow + 8;
    sheet.getRange(dayTotRow, 2).setValue('Total for ' + day + ':')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold').setFontColor(COLOUR.teal)
      .setHorizontalAlignment('right');
    sheet.getRange(dayTotRow, 3)
      .setFormula('=SUMIF(\'6. Allocate\'!$F$5:$F$54,"' + day + '",\'6. Allocate\'!$D$5:$D$54)&" min / "&ROUND(SUMIF(\'6. Allocate\'!$F$5:$F$54,"' + day + '",\'6. Allocate\'!$D$5:$D$54)/60,1)&" hrs"')
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.teal);
    sheet.setRowHeight(dayTotRow, 20);

    currentRow = dayTotRow + 2;
  }

  // ── Completion summary at bottom ──
  sheet.getRange(currentRow, 2, 1, 4).merge();
  sheet.getRange(currentRow, 2).setValue('COMPLETION TRACKER')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOUR.white).setBackground(COLOUR.purple)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(currentRow, 28);
  currentRow++;

  var completionData = [
    ['Done:', '=COUNTIF(\'6. Allocate\'!G5:G54,"Done")', 'tasks completed'],
    ['To Do:', '=COUNTIF(\'6. Allocate\'!G5:G54,"To Do")', 'remaining'],
    ['Delegated:', '=COUNTIF(\'6. Allocate\'!G5:G54,"Delegated")', 'delegated'],
    ['Dropped:', '=COUNTIF(\'6. Allocate\'!G5:G54,"Dropped")', 'dropped'],
  ];

  for (var c = 0; c < completionData.length; c++) {
    var cRow = currentRow + c;
    sheet.getRange(cRow, 2).setValue(completionData[c][0])
      .setFontFamily('Arial').setFontSize(10).setFontColor(COLOUR.darkGrey);
    sheet.getRange(cRow, 3).setFormula(completionData[c][1])
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
      .setFontColor(COLOUR.teal).setHorizontalAlignment('center');
    sheet.getRange(cRow, 4).setValue(completionData[c][2])
      .setFontFamily('Arial').setFontSize(9).setFontColor(COLOUR.midGrey);
    sheet.setRowHeight(cRow, 22);
  }

  // Completion % bar formula
  var pctRow = currentRow + completionData.length + 1;
  sheet.getRange(pctRow, 2).setValue('Completion rate:')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor(COLOUR.darkGrey);
  sheet.getRange(pctRow, 3)
    .setFormula('=IFERROR(ROUND(COUNTIF(\'6. Allocate\'!G5:G54,"Done")/COUNTA(\'1. List\'!C5:C54)*100,0)&"%","0%")')
    .setFontFamily('Arial').setFontSize(12).setFontWeight('bold')
    .setFontColor(COLOUR.teal).setHorizontalAlignment('center');
  sheet.setRowHeight(pctRow, 28);

  sheet.setFrozenRows(3);
}

// ─── HELPER: Build tab header (rows 1–2) ─────────────────────────────────────
function buildTabHeader(sheet, title, colour, subtitle) {
  // Row 1: big title
  sheet.getRange(1, 1, 1, 7).merge();
  sheet.getRange(1, 1).setValue(title)
    .setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
    .setFontColor(COLOUR.white).setBackground(colour)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sheet.setRowHeight(1, 40);

  // Row 2: subtitle
  sheet.getRange(2, 1, 1, 7).merge();
  sheet.getRange(2, 1).setValue(subtitle)
    .setFontFamily('Arial').setFontSize(9).setFontStyle('italic')
    .setFontColor(colour).setBackground(COLOUR.lightTeal)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(2, 28);
}

// ─── HELPER: Style a header row ──────────────────────────────────────────────
function styleHeaderRow(sheet, row, numCols, fontColour, bgColour, startCol) {
  startCol = startCol || 1;
  var range = sheet.getRange(row, startCol, 1, numCols);
  range
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(fontColour)
    .setBackground(bgColour)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, bgColour, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(row, 28);
}

// ─── HELPER: Style an input cell ─────────────────────────────────────────────
function styleInputCell(range) {
  range
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontColor(COLOUR.darkGrey)
    .setBackground(COLOUR.white)
    .setBorder(true, true, true, true, false, false, COLOUR.midGrey, SpreadsheetApp.BorderStyle.SOLID);
}

// ─── HELPER: Add a footer hint ───────────────────────────────────────────────
function addStepFooter(sheet, row, col, text) {
  sheet.getRange(row, col, 1, 4).merge();
  sheet.getRange(row, col).setValue('→  ' + text)
    .setFontFamily('Arial').setFontSize(9).setFontStyle('italic')
    .setFontColor(COLOUR.white).setBackground(COLOUR.teal);
  sheet.setRowHeight(row, 24);
}

