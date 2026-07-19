/**
 * ULTIMATE HABIT TRACKER — Core Generator (batched for performance)
 * Builds an Overview dashboard + 12 month tabs: colored week blocks,
 * weekend tint, native checkboxes, progress/complete/incomplete rows
 * with color-coded progress, today's-date highlight, monthly focus box,
 * daily reflections row.
 *
 * All reads/writes are batched (setValues/setBackgrounds over whole
 * ranges) instead of looping cell-by-cell — looping per-cell across
 * 12 months x ~31 days is thousands of network calls and blows past
 * Apps Script's 6-minute execution limit.
 */

const CONFIG = {
  DEFAULT_HABITS: [
    "Wake up at 6am",
    "Hydrate 2.5L water",
    "30 min exercise",
    "Meditate 10 min",
    "Read 20 pages",
    "GATE study 2 hrs",
    "Journal",
    "Sleep 8 hours"
  ],
  MONTH_NAMES: ["January","February","March","April","May","June",
                "July","August","September","October","November","December"],
  WEEK_COLORS: ["#F9C9D2","#FBE0B4","#C8E6C9","#BBDEFB","#D7BDE2"],
  WEEKEND_COLOR: "#FFCDD2",
  HEADER_COLOR: "#4A4A6A",
  TITLE_COLOR: "#2F3E9E"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🌟 Habit Tracker")
    .addItem("Build / Rebuild Full Tracker", "buildHabitTracker")
    .addItem("Change Year & Rebuild...", "changeYear")
    .addSeparator()
    .addItem("Add New Habit...", "addNewHabit")
    .addSeparator()
    .addItem("Enable Daily Reminder (8 PM)", "enableDailyReminder")
    .addItem("Disable Daily Reminder", "disableDailyReminder")
    .addToUi();
}

// Habit list lives in Script Properties (as JSON) so "Add New Habit" can grow
// it over time without editing code. DEFAULT_HABITS is only the initial seed.
function getHabits() {
  const stored = PropertiesService.getScriptProperties().getProperty("HABITS");
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    } catch (e) { /* fall through to default below */ }
  }
  return CONFIG.DEFAULT_HABITS.slice();
}

function saveHabits(list) {
  PropertiesService.getScriptProperties().setProperty("HABITS", JSON.stringify(list));
}

// Reads the active year from Script Properties (defaults to 2026 the first time)
function getYear() {
  const stored = PropertiesService.getScriptProperties().getProperty("YEAR");
  return stored ? parseInt(stored, 10) : 2026;
}

// Menu action: ask the user for a new year, save it, then rebuild everything
function changeYear() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Change Year",
    "Enter the year you want to build the tracker for (e.g. 2027):",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const year = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(year) || year < 2000 || year > 2100) {
    ui.alert("Please enter a valid 4-digit year.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("YEAR", String(year));
  buildHabitTracker();
}

function buildHabitTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildOverviewSheet(ss);
  CONFIG.MONTH_NAMES.forEach((name, idx) => buildMonthSheet(ss, name, idx));
  buildHeatmapSheet(ss);
  removeDefaultSheet(ss);
  ss.setActiveSheet(ss.getSheetByName("Overview"));
  SpreadsheetApp.getUi().alert("Habit Tracker built! Check the month tabs at the bottom.");
}

function removeDefaultSheet(ss) {
  const sheet = ss.getSheetByName("Sheet1");
  if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
}

function buildOverviewSheet(ss) {
  let sheet = ss.getSheetByName("Overview");
  if (!sheet) sheet = ss.insertSheet("Overview", 0);
  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setRowHeight(1, 35);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 30);

  // ===== TITLE =====
  sheet.getRange("A1:H2").merge()
    .setValue("🌟 ULTIMATE HABIT TRACKER " + getYear())
    .setFontSize(22).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#2F3E9E").setFontColor("white");

  // ===== OVERALL STATS =====
  sheet.getRange("A4").setValue("Overall Statistics").setFontSize(14).setFontWeight("bold");
  sheet.getRange("A6").setValue("Current Year");
  sheet.getRange("B6").setValue(getYear()).setNumberFormat("0");
  sheet.getRange("A7").setValue("Total Habits");
  sheet.getRange("B7").setValue(getHabits().length).setNumberFormat("0");
  sheet.getRange("A8").setValue("Average Completion");
  // Month table lives in rows 12-23 (12 months) — this range must match that.
  sheet.getRange("B8").setFormula("=AVERAGE(B12:B23)").setNumberFormat("0%");
  sheet.getRange("B6:B7").setHorizontalAlignment("left");

  // ===== MONTH TABLE =====
  sheet.getRange("A11:C11").setValues([["Month", "Progress", "Visual"]])
    .setBackground(CONFIG.HEADER_COLOR).setFontColor("white").setFontWeight("bold");

  const monthNameRows = CONFIG.MONTH_NAMES.map(m => [m]);
  sheet.getRange(12, 1, monthNameRows.length, 1).setValues(monthNameRows);

  const progressFormulas = CONFIG.MONTH_NAMES.map(m => [`=IFERROR('${m}'!J1,0)`]);
  sheet.getRange(12, 2, progressFormulas.length, 1).setValues(progressFormulas).setNumberFormat("0%");

  const visualFormulas = CONFIG.MONTH_NAMES.map((m, i) => {
    const row = 12 + i;
    return [`=REPT("🟩",ROUND(B${row}*10))&REPT("⬜",10-ROUND(B${row}*10))`];
  });
  sheet.getRange(12, 3, visualFormulas.length, 1).setValues(visualFormulas);

  sheet.autoResizeColumns(1, 2);

  // ===== DASHBOARD CARDS =====
  sheet.getRange("D4:E5").merge().setValue("📅 YEAR\n" + getYear())
    .setBackground("#E3F2FD").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

  sheet.getRange("D7:E8").merge().setValue("✅ HABITS\n" + getHabits().length)
    .setBackground("#E8F5E9").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

  sheet.getRange("G4:H5").merge()
    .setFormula('="📊 PROGRESS"&CHAR(10)&TEXT(B8,"0%")')
    .setBackground("#FFF8E1").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

  sheet.getRange("D10:E11").merge().setValue("🎯 MONTHLY GOAL\n85%")
    .setBackground("#E8F5E9").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

  sheet.getRange("G10:H11").merge()
    .setFormula('=IF(B8>=0.85,"🏆 Excellent",IF(B8>=0.70,"😊 Good","💪 Keep Going"))')
    .setBackground("#E3F2FD").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

  sheet.getRange("G7:H8").merge().setValue("🔥 STREAK\nComing Soon")
    .setBackground("#FCE4EC").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);
}

// Row where each month's "Progress %" line lives — depends on how many
// habits there are (firstHabitRow=5, so lastHabitRow=4+count, +2 for the gap).
// Both buildMonthSheet and buildHeatmapSheet must agree on this number.
function getProgressRow() {
  return getHabits().length + 6;
}

function buildMonthSheet(ss, monthName, monthIndex) {
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) sheet = ss.insertSheet(monthName);

  const daysInMonth = new Date(getYear(), monthIndex + 1, 0).getDate();
  const firstDataCol = 2; // Column B
  const lastDataCol = firstDataCol + daysInMonth - 1;
  const habitCount = getHabits().length;
  const firstHabitRow = 5;
  const lastHabitRow = firstHabitRow + habitCount - 1;

  // Preserve any already-checked boxes before wiping the sheet — sheet.clear()
  // erases everything, so without this a rebuild would silently lose all
  // tracked days. Only restorable when habit count/day count line up with
  // what's already there; if not (brand-new sheet, size mismatch), skip safely.
  let savedGrid = null;
  try {
    savedGrid = sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth).getValues();
  } catch (e) {
    savedGrid = null;
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();

  // Expand columns if needed (new sheets default to only 26 columns)
  const neededCols = lastDataCol + 3;
  const currentCols = sheet.getMaxColumns();
  if (currentCols < neededCols) {
    sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  // ---- Title + top stats (one batched row) ----
  sheet.getRange(1, 1, 1, 10).setValues([[
    monthName + " " + getYear(), "", "", "",
    "Habits", `=COUNTA($A$5:$A$${lastHabitRow})`,
    "Completed", `=COUNTIF($B$5:${columnLetter(lastDataCol)}$${lastHabitRow},TRUE)`,
    "Progress %", `=IFERROR(H1/(F1*${daysInMonth}),0)`
  ]]);
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setFontColor(CONFIG.TITLE_COLOR);
  sheet.getRange("E1").setFontWeight("bold");
  sheet.getRange("G1").setFontWeight("bold");
  sheet.getRange("I1").setFontWeight("bold");
  sheet.getRange("J1").setNumberFormat("0%");

  // ---- Day headers: weekday row(3) + date row(4), weekend gets a red tint ----
  const weekdayRow = [];
  const dateRow = [];
  const dayColors = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(getYear(), monthIndex, d);
    weekdayRow.push(Utilities.formatDate(date, Session.getScriptTimeZone(), "EEE"));
    dateRow.push(d);
    const isWeekend = (date.getDay() === 0 || date.getDay() === 6);
    const weekIndex = Math.floor((d - 1) / 7) % CONFIG.WEEK_COLORS.length;
    dayColors.push(isWeekend ? CONFIG.WEEKEND_COLOR : CONFIG.WEEK_COLORS[weekIndex]);
  }
  const headerRange3 = sheet.getRange(3, firstDataCol, 1, daysInMonth);
  const headerRange4 = sheet.getRange(4, firstDataCol, 1, daysInMonth);
  headerRange3.setValues([weekdayRow]);
  headerRange4.setValues([dateRow]);
  headerRange3.setBackgrounds([dayColors]);
  headerRange4.setBackgrounds([dayColors]);
  sheet.getRange(3, firstDataCol, 2, daysInMonth)
    .setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange("A3").setValue("Daily Habits").setFontWeight("bold")
    .setFontColor("white").setBackground(CONFIG.HEADER_COLOR);
  sheet.getRange("A4").setBackground(CONFIG.HEADER_COLOR);

  // ---- Habit rows + checkboxes (batched) ----
  sheet.getRange(firstHabitRow, 1, habitCount, 1)
    .setValues(getHabits().map(h => [h]))
    .setFontWeight("bold");
  sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth).insertCheckboxes();

  // Restore whatever checkbox data we saved, now that the fresh checkboxes exist
  if (savedGrid) {
    try {
      sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth).setValues(savedGrid);
    } catch (e) { /* dimensions changed unexpectedly — skip restore, not fatal */ }
  }

  // ---- Conditional formatting (collected into one array, applied once) ----
  const rules = [];
  const gridRange = sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth);
  const topLeftA1 = `${columnLetter(firstDataCol)}${firstHabitRow}`;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${topLeftA1}=TRUE`)
    .setBackground("#A9D18E")
    .setRanges([gridRange])
    .build());

  // ---- Progress / Complete / Incomplete rows (batched formulas) ----
  const progressRow = getProgressRow();
  const completeRow = progressRow + 1;
  const incompleteRow = progressRow + 2;
  sheet.getRange(progressRow, 1, 3, 1)
    .setValues([["Progress %"], ["Complete"], ["Incomplete"]])
    .setFontWeight("bold");

  const completeFormulas = [];
  const incompleteFormulas = [];
  const progressFormulas = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const colL = columnLetter(firstDataCol + d - 1);
    completeFormulas.push(`=COUNTIF(${colL}${firstHabitRow}:${colL}${lastHabitRow},TRUE)`);
    incompleteFormulas.push(`=COUNTIF(${colL}${firstHabitRow}:${colL}${lastHabitRow},FALSE)`);
    progressFormulas.push(`=IFERROR(${colL}${completeRow}/(${colL}${completeRow}+${colL}${incompleteRow}),0)`);
  }
  const progressRange = sheet.getRange(progressRow, firstDataCol, 1, daysInMonth);
  progressRange.setValues([progressFormulas]).setNumberFormat("0%");
  sheet.getRange(completeRow, firstDataCol, 1, daysInMonth).setValues([completeFormulas]);
  sheet.getRange(incompleteRow, firstDataCol, 1, daysInMonth).setValues([incompleteFormulas]);

  // Color-code the whole progress row in one shot: green >=80%, yellow 50-79%, red <50%
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.8).setBackground("#A5D6A7")
    .setRanges([progressRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.79).setBackground("#FFE082")
    .setRanges([progressRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.5).setBackground("#EF9A9A")
    .setRanges([progressRange]).build());

  sheet.setConditionalFormatRules(rules);

  // ---- Monthly Focus box ----
  // Header spans rows 1-4 (fully inside frozen zone), body spans rows 5-8
  // (fully outside it) — merges must never straddle a freeze boundary.
  const focusCol = lastDataCol + 2;
  sheet.getRange(1, focusCol, 4, 2).merge().setValue("This Month's Focus")
    .setFontWeight("bold").setBackground(CONFIG.HEADER_COLOR).setFontColor("white")
    .setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true);
  const focusBox = sheet.getRange(5, focusCol, 4, 2);
  focusBox.merge().setValue("Click here and write your focus for the month...")
    .setFontStyle("italic").setFontColor("#888888")
    .setVerticalAlignment("top").setWrap(true);
  focusBox.setBorder(true, true, true, true, false, false);

  // ---- Daily Reflections row ----
  const reflectRow = incompleteRow + 2;
  sheet.getRange(reflectRow, 1).setValue("Daily Reflections").setFontWeight("bold")
    .setBackground(CONFIG.HEADER_COLOR).setFontColor("white");
  sheet.getRange(reflectRow + 1, firstDataCol, 1, daysInMonth)
    .setBorder(true, true, true, true, true, true);

  // ---- Column widths / freeze ----
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidths(firstDataCol, daysInMonth, 28);
  sheet.setColumnWidth(focusCol, 100);
  sheet.setColumnWidth(focusCol + 1, 100);
  sheet.setFrozenColumns(2); // keep habit names + Day 1 visible while scrolling right
  sheet.setFrozenRows(4);

  // ---- Highlight today's column with a blue border, only on the current month tab ----
  const today = new Date();
  if (today.getFullYear() === getYear() && today.getMonth() === monthIndex) {
    const todayCol = firstDataCol + today.getDate() - 1;
    const highlightHeight = (reflectRow + 1) - 3 + 1;
    sheet.getRange(3, todayCol, highlightHeight, 1)
      .setBorder(true, true, true, true, false, false, "blue", SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
}

/**
 * Builds a GitHub-style contribution heatmap: one row per weekday (Sun-Sat),
 * one column per week of the year, each cell colored by that day's
 * completion % (pulled from the matching month sheet's Progress row).
 */
function buildHeatmapSheet(ss) {
  let sheet = ss.getSheetByName("Heatmap");
  if (!sheet) sheet = ss.insertSheet("Heatmap");
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);

  const year = getYear();
  const jan1 = new Date(year, 0, 1);
  const dec31 = new Date(year, 11, 31);
  const startWeekday = jan1.getDay(); // 0=Sun ... 6=Sat
  const totalDays = Math.round((dec31 - jan1) / 86400000) + 1;
  const totalCells = startWeekday + totalDays;
  const totalWeeks = Math.ceil(totalCells / 7);
  const progressRow = getProgressRow();

  sheet.getRange("A1").setValue("🔥 " + year + " HABIT HEATMAP")
    .setFontSize(18).setFontWeight("bold").setFontColor(CONFIG.TITLE_COLOR);
  sheet.getRange("A2").setValue("Each cell = one day's completion % (darker = better)")
    .setFontStyle("italic").setFontColor("#888888");

  // Make sure there's room: 1 label column + totalWeeks columns
  const neededCols = totalWeeks + 2;
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  const HEATMAP_FIRST_ROW = 4; // rows 4-10 = Sun..Sat
  const dayLabels = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  sheet.getRange(HEATMAP_FIRST_ROW, 1, 7, 1)
    .setValues(dayLabels.map(d => [d]))
    .setFontWeight("bold");

  // Build the whole formula grid in memory first, write it in one batch call
  const grid = [];
  for (let r = 0; r < 7; r++) grid.push(new Array(totalWeeks).fill(""));

  for (let d = 0; d < totalDays; d++) {
    const date = new Date(year, 0, 1 + d);
    const cellIndex = startWeekday + d;
    const col = Math.floor(cellIndex / 7);
    const row = cellIndex % 7;
    const monthName = CONFIG.MONTH_NAMES[date.getMonth()];
    const dayColLetter = columnLetter(2 + date.getDate() - 1); // firstDataCol=2 in month sheets
    grid[row][col] = `=IFERROR('${monthName}'!${dayColLetter}${progressRow},0)`;
  }

  const heatRange = sheet.getRange(HEATMAP_FIRST_ROW, 2, 7, totalWeeks);
  heatRange.setValues(grid);
  heatRange.setNumberFormat("0%").setHorizontalAlignment("center");

  // Color scale: light grey (0%) -> light green (50%) -> dark green (100%)
  const colorRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#EBEDF0", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMidpointWithValue("#66BB6A", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMaxpointWithValue("#1B5E20", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setRanges([heatRange])
    .build();
  sheet.setConditionalFormatRules([colorRule]);

  // Legend row
  const legendRow = HEATMAP_FIRST_ROW + 8;
  sheet.getRange(legendRow, 1).setValue("Less");
  sheet.getRange(legendRow, 2, 1, 4).setBackgrounds([["#EBEDF0", "#C8E6C9", "#66BB6A", "#1B5E20"]]);
  sheet.getRange(legendRow, 6).setValue("More");

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidths(2, totalWeeks, 20);
  for (let r = 0; r < 7; r++) sheet.setRowHeight(HEATMAP_FIRST_ROW + r, 20);
}

/**
 * Adds a new habit to every month sheet by inserting one row directly
 * (not a full rebuild) — this means existing checked days for every
 * other habit are completely untouched.
 */
function addNewHabit() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Add New Habit",
    "Enter the new habit's name:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const name = response.getResponseText().trim();
  if (!name) { ui.alert("Please enter a habit name."); return; }

  const habits = getHabits();
  if (habits.indexOf(name) !== -1) {
    ui.alert("That habit already exists.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldCount = habits.length;
  habits.push(name);
  saveHabits(habits);

  const firstHabitRow = 5;
  const firstDataCol = 2;

  CONFIG.MONTH_NAMES.forEach((monthName, monthIndex) => {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return; // not built yet — the next full build will include it

    const daysInMonth = new Date(getYear(), monthIndex + 1, 0).getDate();
    const newHabitRow = firstHabitRow + oldCount; // one row after the last existing habit

    sheet.insertRowBefore(newHabitRow);
    sheet.getRange(newHabitRow, 1).setValue(name).setFontWeight("bold");
    sheet.getRange(newHabitRow, firstDataCol, 1, daysInMonth).insertCheckboxes();

    // Row totals/formulas and conditional formatting need their ranges
    // widened by one row — cheap to just recompute them.
    refreshMonthTotals(sheet, monthIndex, oldCount + 1);
  });

  ui.alert(`"${name}" added to all 12 months! Existing checkmarks were not touched.`);
}

/**
 * Recomputes the row-1 stats, Progress/Complete/Incomplete formulas, and
 * conditional formatting for a month sheet to match a new habit count —
 * without touching the checkbox grid itself.
 */
function refreshMonthTotals(sheet, monthIndex, habitCount) {
  const daysInMonth = new Date(getYear(), monthIndex + 1, 0).getDate();
  const firstDataCol = 2;
  const lastDataCol = firstDataCol + daysInMonth - 1;
  const firstHabitRow = 5;
  const lastHabitRow = firstHabitRow + habitCount - 1;
  const progressRow = getProgressRow();
  const completeRow = progressRow + 1;
  const incompleteRow = progressRow + 2;

  sheet.getRange("F1").setFormula(`=COUNTA($A$5:$A$${lastHabitRow})`);
  sheet.getRange("H1").setFormula(`=COUNTIF($B$5:${columnLetter(lastDataCol)}$${lastHabitRow},TRUE)`);
  sheet.getRange("J1").setFormula(`=IFERROR(H1/(F1*${daysInMonth}),0)`);

  const completeFormulas = [];
  const incompleteFormulas = [];
  const progressFormulas = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const colL = columnLetter(firstDataCol + d - 1);
    completeFormulas.push(`=COUNTIF(${colL}${firstHabitRow}:${colL}${lastHabitRow},TRUE)`);
    incompleteFormulas.push(`=COUNTIF(${colL}${firstHabitRow}:${colL}${lastHabitRow},FALSE)`);
    progressFormulas.push(`=IFERROR(${colL}${completeRow}/(${colL}${completeRow}+${colL}${incompleteRow}),0)`);
  }
  const progressRange = sheet.getRange(progressRow, firstDataCol, 1, daysInMonth);
  progressRange.setValues([progressFormulas]).setNumberFormat("0%");
  sheet.getRange(completeRow, firstDataCol, 1, daysInMonth).setValues([completeFormulas]);
  sheet.getRange(incompleteRow, firstDataCol, 1, daysInMonth).setValues([incompleteFormulas]);

  const gridRange = sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth);
  const topLeftA1 = `${columnLetter(firstDataCol)}${firstHabitRow}`;
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${topLeftA1}=TRUE`)
    .setBackground("#A9D18E").setRanges([gridRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.8).setBackground("#A5D6A7")
    .setRanges([progressRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.79).setBackground("#FFE082")
    .setRanges([progressRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.5).setBackground("#EF9A9A")
    .setRanges([progressRange]).build());
  sheet.setConditionalFormatRules(rules);
}

/**
 * Checks today's row in the current month sheet; if any habit is still
 * unchecked, emails a reminder to the sheet owner. Meant to run on a
 * daily time-based trigger (see enableDailyReminder).
 */
function checkAndSendReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  if (today.getFullYear() !== getYear()) return; // tracker isn't built for this year

  const monthName = CONFIG.MONTH_NAMES[today.getMonth()];
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) return;

  const firstDataCol = 2;
  const firstHabitRow = 5;
  const habits = getHabits();
  const todayCol = firstDataCol + today.getDate() - 1;

  const values = sheet.getRange(firstHabitRow, todayCol, habits.length, 1).getValues();
  const checkedCount = values.filter(row => row[0] === true).length;
  const remaining = habits.length - checkedCount;

  if (remaining > 0) {
    const email = Session.getActiveUser().getEmail();
    if (!email) return;
    MailApp.sendEmail({
      to: email,
      subject: `⏰ Habit Tracker: ${remaining} habit(s) left today`,
      body: `You've completed ${checkedCount} of ${habits.length} habits today ` +
            `(${monthName} ${today.getDate()}, ${getYear()}).\n\n` +
            `${remaining} habit(s) still unchecked. Open your Habit Tracker and mark them off!\n\n` +
            `— Ultimate Habit Tracker`
    });
  }
}

function enableDailyReminder() {
  disableDailyReminder(); // avoid stacking duplicate triggers
  ScriptApp.newTrigger("checkAndSendReminder")
    .timeBased()
    .everyDays(1)
    .atHour(20) // Apps Script runs this within an hour window around 8 PM
    .create();
  SpreadsheetApp.getUi().alert(
    "Daily reminder enabled. You'll get an email around 8 PM if any habits are still unchecked."
  );
}

function disableDailyReminder() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "checkAndSendReminder") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function columnLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}