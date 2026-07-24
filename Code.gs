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
  WEEKDAY_COLOR: "#DCE6F1",  // light blue — Mon-Fri
  WEEKEND_COLOR: "#F5B7B1",  // light coral — Sat-Sun, clearly different hue
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
    .addItem("Rename Habit...", "renameHabit")
    .addItem("Delete Habit...", "deleteHabit")
    .addItem("Reset Habits to Default...", "resetHabitsToDefault")
    .addSeparator()
    .addItem("Enable Daily Reminder (8 PM)", "enableDailyReminder")
    .addItem("Disable Daily Reminder", "disableDailyReminder")
    .addSeparator()
    .addItem("Refresh Streak", "refreshStreakCard")
    .addItem("Move Today's Highlight", "refreshTodayHighlight")
    .addToUi();
}

// Habit list lives in Script Properties (as JSON) so "Add New Habit" can grow
// it over time without editing code. DEFAULT_HABITS is only the initial seed.
function getHabits() {
  const stored = PropertiesService.getScriptProperties().getProperty("HABITS");
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      // Sanity check: must be a reasonable array of plain strings. Without this,
      // a corrupted property (e.g. from a bad script run) could silently make
      // habitCount huge, which then produces nonsense COUNTIF ranges everywhere.
      if (Array.isArray(parsed) && parsed.length > 0 && parsed.length <= 50 &&
          parsed.every(h => typeof h === "string" && h.length > 0)) {
        return parsed;
      }
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
  buildHabitTracker(false); // false = don't carry over old checkmarks; a new year starts blank
}

function buildHabitTracker(preserveData) {
  if (preserveData === undefined) preserveData = true; // default: normal rebuilds keep existing data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildOverviewSheet(ss);
  CONFIG.MONTH_NAMES.forEach((name, idx) => buildMonthSheet(ss, name, idx, preserveData));
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
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);

  // Muted, professional palette — one accent color, greys for structure.
  const ACCENT = "#1E3A5F";     // deep slate blue
  const ACCENT_LIGHT = "#EAF1F8";
  const TEXT_DARK = "#1F2937";
  const TEXT_GREY = "#6B7280";
  const BORDER_GREY = "#E2E5E9";

  sheet.setColumnWidth(1, 26);   // left margin
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 26);   // gap
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 26);   // gap
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 150);

  // ===== HEADER BAR =====
  sheet.setRowHeight(1, 46);
  sheet.getRange(1, 1, 1, 8).merge()
    .setValue("  Ultimate Habit Tracker   ·   " + getYear())
    .setFontSize(16).setFontWeight("bold").setFontColor("white")
    .setFontFamily("Arial")
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBackground(ACCENT);

  // ===== SECTION LABEL: OVERVIEW =====
  sheet.getRange("B3").setValue("OVERVIEW")
    .setFontSize(10).setFontWeight("bold").setFontColor(TEXT_GREY)
    .setFontFamily("Arial");
  sheet.getRange(3, 2, 1, 6).setBorder(false, false, true, false, false, false, BORDER_GREY, SpreadsheetApp.BorderStyle.SOLID);

  // ===== STAT CARDS (2x2 grid, muted, thin accent top-border) =====
  const cardRows = [4, 4, 4, 7, 7];
  const cardCols = [2, 5, 8, 2, 5];
  const cardLabels = ["YEAR", "TOTAL HABITS", "CURRENT STREAK", "AVG COMPLETION", "BEST MONTH"];
  const streakDays = calculateStreak();
  const cardValueFormulas = [
    String(getYear()),
    String(getHabits().length),
    streakDays + (streakDays === 1 ? " day" : " days"),
    `=TEXT(AVERAGE(C13:C24),"0%")`,
    `=IFERROR(INDEX(B13:B24, MATCH(MAX(C13:C24), C13:C24, 0)), "—")`
  ];

  for (let i = 0; i < 4; i++) {
    const r = cardRows[i], c = cardCols[i];

    // Border the whole 2x2 card area BEFORE merging its two rows separately —
    // merging the entire block into one cell was the bug: writing to what
    // used to be the "value" row afterward had no real cell to land in.
    sheet.getRange(r, c, 2, 2)
      .setBackground("white")
      .setBorder(true, true, true, true, false, false, BORDER_GREY, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, c, 1, 2)
      .setBorder(true, null, null, null, false, false, ACCENT, SpreadsheetApp.BorderStyle.SOLID_THICK);

    const labelRange = sheet.getRange(r, c, 1, 2);
    labelRange.merge().setValue(cardLabels[i])
      .setFontSize(9).setFontWeight("bold").setFontColor(TEXT_GREY).setFontFamily("Arial")
      .setVerticalAlignment("bottom");

    const valueRange = sheet.getRange(r + 1, c, 1, 2);
    valueRange.merge();
    if (String(cardValueFormulas[i]).startsWith("=")) {
      valueRange.setFormula(cardValueFormulas[i]);
    } else {
      valueRange.setValue(cardValueFormulas[i]);
    }
    valueRange.setFontSize(20).setFontWeight("bold").setFontColor(TEXT_DARK).setFontFamily("Arial")
      .setVerticalAlignment("top");
  }
  sheet.setRowHeight(4, 22);
  sheet.setRowHeight(5, 32);
  sheet.setRowHeight(7, 22);
  sheet.setRowHeight(8, 32);

  // ===== SECTION LABEL: MONTHLY PROGRESS =====
  sheet.getRange("B11").setValue("MONTHLY PROGRESS")
    .setFontSize(10).setFontWeight("bold").setFontColor(TEXT_GREY)
    .setFontFamily("Arial");
  sheet.getRange(11, 2, 1, 6).setBorder(false, false, true, false, false, false, BORDER_GREY, SpreadsheetApp.BorderStyle.SOLID);

  // ===== MONTH TABLE (clean: month, %, single color-scaled progress bar) =====
  const tableHeaderRow = 12;
  sheet.getRange(tableHeaderRow, 2, 1, 2).setValues([["Month", "Completion"]])
    .setBackground(ACCENT).setFontColor("white").setFontWeight("bold").setFontFamily("Arial");

  const monthNameRows = CONFIG.MONTH_NAMES.map(m => [m]);
  sheet.getRange(13, 2, monthNameRows.length, 1).setValues(monthNameRows).setFontFamily("Arial");

  const progressFormulas = CONFIG.MONTH_NAMES.map(m => [`=IFERROR('${m}'!J1,0)`]);
  const progressRange = sheet.getRange(13, 3, progressFormulas.length, 1);
  progressRange.setValues(progressFormulas).setNumberFormat("0%").setFontFamily("Arial");

  // A single, subtle horizontal bar per month instead of emoji squares —
  // built as a colored rectangle whose width mirrors the % via a formula,
  // achieved by overlaying a background color-scale on the % cell itself.
  const barColorRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#F3F4F6", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMidpointWithValue("#93C5FD", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMaxpointWithValue(ACCENT, SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setRanges([progressRange])
    .build();
  sheet.setConditionalFormatRules([barColorRule]);

  const tableRange = sheet.getRange(tableHeaderRow, 2, 13, 2);
  tableRange.setBorder(true, true, true, true, true, true, BORDER_GREY, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(13, 2, 12, 2).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 100);
}

// Row where each month's "Progress %" line lives — depends on how many
// habits there are (firstHabitRow=5, so lastHabitRow=4+count, +2 for the gap).
// Both buildMonthSheet and buildHeatmapSheet must agree on this number.
function getProgressRow() {
  return getHabits().length + 6;
}

function buildMonthSheet(ss, monthName, monthIndex, preserveData) {
  if (preserveData === undefined) preserveData = true;
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
  // tracked days. Skipped entirely on a year change (preserveData=false),
  // since a new year's checkboxes should start blank, not inherit the old
  // year's ticks at the same day-of-month position.
  let savedGrid = null;
  if (preserveData) {
    try {
      savedGrid = sheet.getRange(firstHabitRow, firstDataCol, habitCount, daysInMonth).getValues();
    } catch (e) {
      savedGrid = null;
    }
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();
  // sheet.clear() clears content/formatting but NOT data validations — checkboxes
  // are implemented as a data validation rule, so without this, any leftover
  // checkbox columns from a previous run (e.g. before a column-count fix) stick
  // around forever as "ghost" checkboxes with no header above them.
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  // Expand columns if needed (new sheets default to only 26 columns)
  const neededCols = lastDataCol + 3;
  const currentCols = sheet.getMaxColumns();
  if (currentCols < neededCols) {
    sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  // ---- Title + hidden stat cells (F1/H1/J1 kept as real numbers — Overview reads J1) ----
  sheet.getRange("A1").setValue(monthName + " " + getYear())
    .setFontSize(16).setFontWeight("bold").setFontColor(CONFIG.TITLE_COLOR);
  sheet.getRange("F1").setFormula(`=COUNTA($A$5:$A$${lastHabitRow})`);
  sheet.getRange("H1").setFormula(`=COUNTIF($B$5:${columnLetter(lastDataCol)}$${lastHabitRow},TRUE)`);
  sheet.getRange("J1").setFormula(`=IFERROR(H1/(F1*${daysInMonth}),0)`).setNumberFormat("0%");

  // Readable summary line in row 2 — merged across columns C-J so it gets enough
  // width WITHOUT widening any individual day-checkbox column (which would make
  // some days visually wider than others). Starts at column C (not B) so the
  // merge stays fully outside the frozen-columns zone (A-B are frozen). Explicit
  // white background stops it from visually blending into the colored day-header
  // row right below it.
  sheet.getRange(2, 3, 1, 8).merge()
    .setFormula('="Habits: "&F1&"   |   Done: "&H1&"   |   Progress: "&TEXT(J1,"0%")')
    .setFontSize(10).setFontColor("#6B7280").setFontFamily("Arial").setBackground("white")
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  // ---- Day headers: weekday row(3) + date row(4), weekend gets a red tint ----
  const weekdayRow = [];
  const dateRow = [];
  const dayColors = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(getYear(), monthIndex, d);
    weekdayRow.push(Utilities.formatDate(date, Session.getScriptTimeZone(), "EEE"));
    dateRow.push(d);
    const isWeekend = (date.getDay() === 0 || date.getDay() === 6);
    dayColors.push(isWeekend ? CONFIG.WEEKEND_COLOR : CONFIG.WEEKDAY_COLOR);
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
    const highlightHeight = incompleteRow - 3 + 1;
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

  const HEATMAP_FIRST_ROW = 4; // rows 4-10 = Sun..Sat (row 3 is reserved for month labels)
  const dayLabels = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  sheet.getRange(HEATMAP_FIRST_ROW, 1, 7, 1)
    .setValues(dayLabels.map(d => [d]))
    .setFontWeight("bold");

  // Month labels: place an abbreviation ("Jan", "Feb"...) at the first week-column
  // where each month begins, GitHub-calendar style.
  const monthLabels = new Array(totalWeeks).fill("");
  let lastMonthPlaced = -1;
  for (let d = 0; d < totalDays; d++) {
    const date = new Date(year, 0, 1 + d);
    const cellIndex = startWeekday + d;
    const col = Math.floor(cellIndex / 7);
    if (date.getMonth() !== lastMonthPlaced) {
      monthLabels[col] = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM");
      lastMonthPlaced = date.getMonth();
    }
  }
  sheet.getRange(HEATMAP_FIRST_ROW - 1, 2, 1, totalWeeks)
    .setValues([monthLabels])
    .setFontWeight("bold")
    .setFontColor(CONFIG.TITLE_COLOR);

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
  // ";;;" is a custom number format with every section empty — hides the
  // displayed value entirely (GitHub's graph shows colored squares, no text).
  // The underlying number is untouched, so conditional-format coloring still works.
  heatRange.setNumberFormat(";;;").setHorizontalAlignment("center");
  // White borders between cells create the visible "gap" GitHub's graph has.
  heatRange.setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

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
  sheet.setColumnWidths(2, totalWeeks, 24);
  for (let r = 0; r < 7; r++) sheet.setRowHeight(HEATMAP_FIRST_ROW + r, 24);
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
 * Deletes a habit from every month sheet by removing its row directly —
 * every other habit's checked days are untouched.
 */
function deleteHabit() {
  const ui = SpreadsheetApp.getUi();
  const habits = getHabits();

  if (habits.length <= 1) {
    ui.alert("Can't delete the last remaining habit.");
    return;
  }

  const response = ui.prompt(
    "Delete Habit",
    "Enter the exact name of the habit to delete:\n\n" + habits.join(", "),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const name = response.getResponseText().trim();
  const idx = habits.indexOf(name);
  if (idx === -1) {
    ui.alert("Habit not found — check the spelling exactly as shown in the list.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const firstHabitRow = 5;
  const rowToDelete = firstHabitRow + idx;

  habits.splice(idx, 1);
  saveHabits(habits);

  CONFIG.MONTH_NAMES.forEach((monthName, monthIndex) => {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    sheet.deleteRow(rowToDelete);
    refreshMonthTotals(sheet, monthIndex, habits.length);
  });

  ui.alert(`"${name}" deleted from all 12 months.`);
}

/**
 * Renames a habit in every month sheet (just updates the label — the
 * checkbox grid position and all checked data stay exactly where they are).
 */
function renameHabit() {
  const ui = SpreadsheetApp.getUi();
  const habits = getHabits();

  const response1 = ui.prompt(
    "Rename Habit",
    "Enter the exact current name of the habit:\n\n" + habits.join(", "),
    ui.ButtonSet.OK_CANCEL
  );
  if (response1.getSelectedButton() !== ui.Button.OK) return;

  const oldName = response1.getResponseText().trim();
  const idx = habits.indexOf(oldName);
  if (idx === -1) {
    ui.alert("Habit not found — check the spelling exactly as shown in the list.");
    return;
  }

  const response2 = ui.prompt(
    "Rename Habit",
    `Enter the new name for "${oldName}":`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response2.getSelectedButton() !== ui.Button.OK) return;

  const newName = response2.getResponseText().trim();
  if (!newName) { ui.alert("Please enter a name."); return; }

  habits[idx] = newName;
  saveHabits(habits);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const row = 5 + idx; // firstHabitRow + idx
  CONFIG.MONTH_NAMES.forEach(monthName => {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    sheet.getRange(row, 1).setValue(newName);
  });

  ui.alert(`Renamed "${oldName}" to "${newName}" in all 12 months.`);
}

/**
 * Recovery action: resets the habit list back to the original 8 defaults
 * and does a full rebuild. Use this if the habit list ever gets corrupted
 * (e.g. showing huge nonsense numbers in the Complete/Incomplete rows).
 */
function resetHabitsToDefault() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Reset Habits",
    "This resets your habit list to the original 8 defaults and rebuilds everything. Continue?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  saveHabits(CONFIG.DEFAULT_HABITS.slice());
  buildHabitTracker(false);
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

  updateStreakCard();

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
  ScriptApp.newTrigger("refreshTodayHighlight")
    .timeBased()
    .everyDays(1)
    .atHour(1) // just after midnight, so the highlight moves early each day
    .create();
  SpreadsheetApp.getUi().alert(
    "Daily reminder enabled. You'll get an email around 8 PM if any habits are still unchecked, " +
    "and the today-highlight will move to the new date automatically each night."
  );
}

function disableDailyReminder() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "checkAndSendReminder" || fn === "refreshTodayHighlight") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * Counts consecutive days, walking backward from today, where every habit
 * was checked. Stops at the first day that isn't 100% complete (or has no
 * data yet). Crosses month-sheet boundaries automatically.
 */
/**
 * Moves the blue "today" border to the current date, without touching
 * anything else — clears yesterday's border (tracked via a Script Property)
 * and draws today's. Meant to run once a day via a trigger, or on demand.
 */
function refreshTodayHighlight() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = getYear();
  const today = new Date();
  if (today.getFullYear() !== year) return; // tracker isn't built for this year

  const firstDataCol = 2;
  const highlightHeight = (getProgressRow() + 2) - 3 + 1; // rows 3 through Incomplete row

  // Remove yesterday's highlight (if we have a record of where it was)
  const lastStr = PropertiesService.getScriptProperties().getProperty("LAST_HIGHLIGHT_DATE");
  if (lastStr) {
    const lastDate = new Date(lastStr);
    if (lastDate.getFullYear() === year) {
      const lastSheet = ss.getSheetByName(CONFIG.MONTH_NAMES[lastDate.getMonth()]);
      if (lastSheet) {
        const lastCol = firstDataCol + lastDate.getDate() - 1;
        lastSheet.getRange(3, lastCol, highlightHeight, 1)
          .setBorder(false, false, false, false, false, false);
      }
    }
  }

  // Draw today's highlight
  const sheet = ss.getSheetByName(CONFIG.MONTH_NAMES[today.getMonth()]);
  if (sheet) {
    const todayCol = firstDataCol + today.getDate() - 1;
    sheet.getRange(3, todayCol, highlightHeight, 1)
      .setBorder(true, true, true, true, false, false, "blue", SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  PropertiesService.getScriptProperties().setProperty("LAST_HIGHLIGHT_DATE", today.toDateString());
}

function calculateStreak() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = getYear();
  const today = new Date();
  if (today.getFullYear() !== year) return 0;

  const habitCount = getHabits().length;
  const firstHabitRow = 5;
  const firstDataCol = 2;
  let streak = 0;
  const cursor = new Date(today);

  while (streak < 1000) { // safety cap, avoids any risk of an infinite loop
    const monthName = CONFIG.MONTH_NAMES[cursor.getMonth()];
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) break;

    const col = firstDataCol + cursor.getDate() - 1;
    let values;
    try {
      values = sheet.getRange(firstHabitRow, col, habitCount, 1).getValues();
    } catch (e) {
      break; // sheet not built yet for this date — streak stops here
    }
    const allDone = values.length > 0 && values.every(row => row[0] === true);
    if (!allDone) break;

    streak++;
    cursor.setDate(cursor.getDate() - 1);
  }
  return streak;
}

// Core streak update — no UI calls, so it's safe to run from an automated
// trigger (SpreadsheetApp.getUi() throws if called outside a user action).
function updateStreakCard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Overview");
  if (!sheet) return null;
  const streakDays = calculateStreak();
  sheet.getRange(5, 8).setValue(streakDays + (streakDays === 1 ? " day" : " days"));
  return streakDays;
}

// Menu action version — same update, plus a confirmation popup for the user
function refreshStreakCard() {
  const streakDays = updateStreakCard();
  if (streakDays === null) {
    SpreadsheetApp.getUi().alert("Build the tracker first.");
  } else {
    SpreadsheetApp.getUi().alert(`Current streak: ${streakDays} day(s).`);
  }
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