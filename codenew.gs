// =============================================================
// VERTICAL LAYOUT — each month is a self-contained block of rows
//
// Block structure (N = block start row):
//   N+0 : Company name header  (merged across all cols, green/red)
//   N+1 : Stat headers (PRESENT…NAME) + Day headers (Mar 1, Mar 2…)
//   N+2 : (blank under stats)  + STATUS / TIME IN / TIME OUT per day + TOTAL DAYS
//   N+3…: Employee data rows (one row per employee)
//   (3 blank gap rows, then next month block starts)
//
// Column layout (starting at col B = col 2):
//   RV  cols 2-12 : PRESENT ABSENT LATE TOTAL_LATES UNDERTIME OVERTIME AWOL SICK_LEAVE WFH SCHED_TIME NAME
//   COMS cols 2-10: PRESENT ABSENT LATE UNDERTIME AWOL SICK_LEAVE WFH SCHED_TIME NAME
//   Daily cols start at col 2+dashColCount, 3 cols per day + 1 TOTAL DAYS col
//
// Month blocks are tracked via a hidden marker in col A:
//   "MONTH_BLOCK:YYYY-M"  at the block's start row
// =============================================================

// ── constants ─────────────────────────────────────────────────

function getDashHeaders(isRV) {
  return isRV
    ? ['PRESENT','ABSENT','LATE','TOTAL LATES (MINS)','UNDERTIME','OVERTIME','AWOL','SICK LEAVE / VACATION LEAVE','WORK FROM HOME','SCHEDULE TIME','NAME']
    : ['PRESENT','ABSENT','LATE','UNDERTIME','AWOL','SICK LEAVE / VACATION LEAVE','WORK FROM HOME','SCHEDULE TIME','NAME'];
}

// RV late minutes use an 11-minute grace period.
// Stored lateMinutes from the app should already be net of grace,
// but this helper keeps server aggregation safe for missing/legacy data.
function getLateMinutesForDept(record, dept) {
  const mins = Number(record && record.lateMinutes);
  if (!isNaN(mins)) return Math.max(0, mins);
  // Legacy fallback only when value is missing
  return dept === 'rv' ? 0 : 15;
}

function fmtT(t) {
  if (!t) return '';
  const [h, m] = t.split(':');
  const hr = parseInt(h);
  return (hr % 12 || 12) + ':' + m + ' ' + (hr >= 12 ? 'PM' : 'AM');
}

// ── block discovery ───────────────────────────────────────────

// Returns array of { key, startRow, year, monthNum }
function findMonthBlocks(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const vals = sheet.getRange(1, 1, lastRow, 1).getValues();
  const blocks = [];
  for (let i = 0; i < vals.length; i++) {
    const v = (vals[i][0] || '').toString();
    if (v.startsWith('MONTH_BLOCK:')) {
      const key = v.replace('MONTH_BLOCK:', '');
      const [y, mo] = key.split('-').map(Number);
      blocks.push({ key, startRow: i + 1, year: y, monthNum: mo });
    }
  }
  return blocks;
}

// Returns [{name, row}] for employee data rows inside a block
function getEmpRowsInBlock(sheet, blockStartRow, nameCol) {
  const dataStart = blockStartRow + 3;
  const lastRow   = sheet.getLastRow();
  const result    = [];
  for (let r = dataStart; r <= lastRow; r++) {
    const marker = (sheet.getRange(r, 1).getValue() || '').toString();
    if (marker.startsWith('MONTH_BLOCK:')) break; // next block started
    const name = (sheet.getRange(r, nameCol).getValue() || '').toString().trim();
    if (name) result.push({ name, row: r });
    else break;
  }
  return result;
}

// ── write one month block's header rows ───────────────────────

function writeBlockHeaders(sheet, startRow, year, month, isRV) {
  const props = PropertiesService.getScriptProperties();
  const savedName = props.getProperty(isRV ? 'companyName_rv' : 'companyName_coms');
  const companyName = savedName || (isRV ? 'RED VICTORY CONSUMERS GOODS TRADING' : 'C. OPERATIONS MANAGEMENT SERVICES');
  const headerColor  = isRV ? '#4CAF50' : '#ef4444';
  const dashHeaders  = getDashHeaders(isRV);
  const dashColCount = dashHeaders.length;          // 11 (RV) or 9 (COMS)
  const daysInMonth  = new Date(year, month + 1, 0).getDate();
  const totalCols    = dashColCount + daysInMonth * 3 + 1; // stats + days*3 + TOTAL DAYS
  const startCol     = 2;                           // col B

  // ── Row N+0: company header (split at freeze boundary so freeze works) ──
  // Left part: cols 2 to NAME col (dashboard columns)
  sheet.getRange(startRow, startCol, 1, dashColCount).merge()
    .setValue(companyName)
    .setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(14);
  // Right part: daily columns — "WEEKLY REPORT — Month Year"
  const dailyCols = daysInMonth * 3 + 1;
  const monthName = new Date(year, month, 1).toLocaleDateString('en-US', { month: 'long' });
  const savedWeekly = props.getProperty((isRV ? 'weeklyLabel_rv' : 'weeklyLabel_coms') + '_' + year + '_' + month);
  const weeklyLabel = savedWeekly || ('WEEKLY REPORT — ' + monthName + ' ' + year);
  sheet.getRange(startRow, startCol + dashColCount, 1, dailyCols).merge()
    .setValue(weeklyLabel)
    .setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(14);

  // ── Row N+1: stat headers + day headers ──────────────────────
  // Stat headers
  sheet.getRange(startRow + 1, startCol, 1, dashColCount)
    .setValues([dashHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);

  // Day headers (merged 3 cols each)
  let col = startCol + dashColCount;
  for (let day = 1; day <= daysInMonth; day++) {
    sheet.getRange(startRow + 1, col, 1, 3).merge()
      .setValue(monthName + ' ' + day)
      .setFontWeight('bold')
      .setBackground('#BFDBFE')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    col += 3;
  }
  // TOTAL DAYS spans rows N+1 and N+2
  sheet.getRange(startRow + 1, col, 2, 1).merge()
    .setValue('TOTAL DAYS')
    .setFontWeight('bold')
    .setBackground('#D1FAE5')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);

  // ── Row N+2: blank under stats + STATUS/TIME IN/TIME OUT ─────
  sheet.getRange(startRow + 2, startCol, 1, dashColCount)
    .setBackground('#f3f3f3')
    .setBorder(true, true, true, true, true, true);

  col = startCol + dashColCount;
  for (let day = 1; day <= daysInMonth; day++) {
    ['STATUS', 'TIME IN', 'TIME OUT'].forEach((lbl, i) => {
      sheet.getRange(startRow + 2, col + i)
        .setValue(lbl)
        .setFontWeight('bold')
        .setBackground('#E3F2FD')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true);
    });
    col += 3;
  }

  // Freeze rows (header rows) and columns up to NAME
  sheet.setFrozenColumns(isRV ? 12 : 10);

  Logger.log('Block headers written: ' + monthName + ' ' + year + ' at row ' + startRow);
}

// ── get or create a month block ───────────────────────────────

function getOrCreateBlock(sheet, year, month, isRV) {
  const key    = year + '-' + month;
  const blocks = findMonthBlocks(sheet);
  const found  = blocks.find(b => b.key === key);
  if (found) return found;

  const dept = isRV ? 'rv' : 'coms';
  const monthValue = (y, m) => y * 12 + m;
  const targetVal = monthValue(year, month);

  // Estimate how many rows this new block needs.
  let empCount = 0;
  try {
    const props = PropertiesService.getScriptProperties();
    const stateJson = props.getProperty('appdata_' + dept);
    if (stateJson) {
      const state = JSON.parse(stateJson);
      if (state.employees && state.employees.length) {
        empCount = state.employees.filter(e => (e.department || dept) === dept && employeeAppearsInSheetBlock(e, year, month)).length;
      }
    }
  } catch (e) {
    Logger.log('Error estimating employee count for new block: ' + e);
  }
  const needed = 3 + empCount + 3; // 3 header rows + employees + 3-row gap

  // Keep blocks ordered by month DESC in sheet: newest on top (May > Apr > Mar).
  let appendRow = 2; // first block starts at row 2 (row 1 = sentinel)
  if (blocks.length > 0) {
    let insertionRow = null;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      const bVal = monthValue(b.year, b.monthNum);
      if (targetVal > bVal) {
        insertionRow = b.startRow; // insert before first older block
        break;
      }
    }
    if (insertionRow === null) {
      const last = blocks[blocks.length - 1];
      const lastNameCol = isRV ? 12 : 10;
      const lastEmpCount = getEmpRowsInBlock(sheet, last.startRow, lastNameCol).length;
      insertionRow = last.startRow + 3 + lastEmpCount + 3;
    }
    if (needed > 0) {
      if (insertionRow <= sheet.getMaxRows()) sheet.insertRows(insertionRow, needed);
      else sheet.insertRowsAfter(sheet.getMaxRows(), needed);
    }
    appendRow = insertionRow;
    Logger.log('Inserted new month block ' + key + ' at row ' + appendRow + ' (needed rows: ' + needed + ')');
  }

  // Write marker in col A (hidden)
  sheet.getRange(appendRow, 1).setValue('MONTH_BLOCK:' + key);
  writeBlockHeaders(sheet, appendRow, year, month, isRV);

  // Seed employee rows into the new block from saved app state (so new months start populated)
  try {
    const dept = isRV ? 'rv' : 'coms';
    seedEmployeesIntoBlock(sheet, appendRow, isRV, dept, year, month);
  } catch (e) {
    Logger.log('seedEmployeesIntoBlock error: ' + e);
  }

  // Reapply headers after seeding employees to guard against any inadvertent
  // format/merge clears that may have occurred during seeding. Harmless
  // if headers are already correct.
  try { writeBlockHeaders(sheet, appendRow, year, month, isRV); } catch (e) { Logger.log('re-writeBlockHeaders error: ' + e); }
  try { applyCompactLayoutToBlock(sheet, appendRow, isRV, year, month); } catch (e) { Logger.log('applyCompactLayoutToBlock error: ' + e); }

  return { key, startRow: appendRow, year, monthNum: month };
}

// ── sheet initialisation ──────────────────────────────────────

function ensureSheetReady(sheet) {
  const sentinel = (sheet.getRange(1, 1).getValue() || '').toString();
  if (sentinel !== 'V2') {
    sheet.clear();
    sheet.clearFormats();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.getRange(1, 1).setValue('V2');
    // Hide the marker column (col A) — make it 2px wide and white text
    sheet.setColumnWidth(1, 2);
    sheet.getRange(1, 1).setFontColor('#ffffff').setBackground('#ffffff');
  }
}

// Rebuild one department sheet from saved/full state to normalize:
// - month order (latest at top)
// - compact/consistent block formatting
// - existing synced data (no logical data changes)
function normalizeSheetLayoutFromState(sheet, dept, state) {
  const isRV = sheet.getName() === 'RV';
  const attendance = (state && state.attendanceData) ? state.attendanceData : [];
  const employees = (state && state.employees) ? state.employees : [];

  // Hard reset layout canvas, then recreate blocks from state.
  sheet.clear();
  sheet.clearFormats();
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.getRange(1, 1).setValue('V2');
  sheet.setColumnWidth(1, 2);
  sheet.getRange(1, 1).setFontColor('#ffffff').setBackground('#ffffff');

  // Determine all months from attendance data and sort DESC (latest first).
  const months = {};
  attendance.forEach(r => {
    if (!r || !r.date || r.department !== dept) return;
    const d = new Date(r.date + 'T00:00:00');
    months[d.getFullYear() + '-' + d.getMonth()] = { y: d.getFullYear(), mo: d.getMonth() };
  });
  const sortedMonths = Object.values(months).sort((a, b) => (b.y * 12 + b.mo) - (a.y * 12 + a.mo));

  // If no attendance yet, create current month block so sheet still has a usable layout.
  if (sortedMonths.length === 0) {
    const now = new Date();
    sortedMonths.push({ y: now.getFullYear(), mo: now.getMonth() });
  }

  // Create ordered month blocks first.
  sortedMonths.forEach(({ y, mo }) => {
    getOrCreateBlock(sheet, y, mo, isRV);
  });

  // Write dashboard stats for each month block.
  sortedMonths.forEach(({ y, mo }) => {
    const monthStats = buildDashStats(employees, attendance, dept, y, mo);
    updateDashboard(sheet, monthStats, y, mo);
  });

  // Write daily cells and compact formatting.
  updateWeeklyReportFull(sheet, attendance);
}

// Apply a compact, consistent visual style for one month block.
// This keeps all blocks aligned with the clean "April" look.
function applyCompactLayoutToBlock(sheet, blockStartRow, isRV, year, month) {
  const dashCount = getDashHeaders(isRV).length;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const totalCols = dashCount + daysInMonth * 3 + 1;
  const blocks = findMonthBlocks(sheet);
  let nextStart = sheet.getLastRow() + 1;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].startRow === blockStartRow) {
      if (i + 1 < blocks.length) nextStart = blocks[i + 1].startRow;
      break;
    }
  }
  const rowsInBlock = Math.max(3, nextStart - blockStartRow);
  sheet.getRange(blockStartRow, 2, rowsInBlock, totalCols)
    .setVerticalAlignment('middle')
    .setFontSize(8);

  // Compact row heights to avoid "bloated" older blocks.
  sheet.setRowHeightsForced(blockStartRow, 1, 20);
  sheet.setRowHeightsForced(blockStartRow + 1, 2, 18);
  const dataRows = Math.max(0, nextStart - (blockStartRow + 3));
  if (dataRows > 0) sheet.setRowHeightsForced(blockStartRow + 3, dataRows, 18);
}

// Employee exists on sheet for block Y-M only if added on or before that month (addedYear/addedMonth from web app).
// Missing addedYear/addedMonth = legacy data → show in every month block.
function employeeAppearsInSheetBlock(emp, blockYear, blockMonth) {
  if (typeof emp.addedYear !== 'number' || typeof emp.addedMonth !== 'number') return true;
  if (blockYear > emp.addedYear) return true;
  if (blockYear === emp.addedYear && blockMonth >= emp.addedMonth) return true;
  return false;
}

// Seed employee rows into a specific month block using saved `appdata_{dept}` state.
function seedEmployeesIntoBlock(sheet, blockStartRow, isRV, dept, blockYear, blockMonth) {
  if (!sheet || !blockStartRow) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const stateJson = props.getProperty('appdata_' + (dept || (isRV ? 'rv' : 'coms')));
    if (!stateJson) return;
    const state = JSON.parse(stateJson);
    if (!state.employees || state.employees.length === 0) return;

    const dashCount = getDashHeaders(isRV).length;
    const nameCol = 1 + dashCount;
    const dataStart = blockStartRow + 3;

    const by = (typeof blockYear === 'number') ? blockYear : new Date().getFullYear();
    const bm = (typeof blockMonth === 'number') ? blockMonth : new Date().getMonth();

    const deptEmps = state.employees.filter(e => e.department === dept && employeeAppearsInSheetBlock(e, by, bm));
    deptEmps.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

    deptEmps.forEach((e, i) => {
      const row = dataStart + i;
      const empObj = {
        name: e.name,
        present: 0, absent: 0, late: 0, totalLates: 0,
        undertime: 0, overtime: 0, awol: 0, sickLeave: 0, wfh: 0,
        scheduleDisplay: e.scheduleDisplay || ''
      };
      writeDashRow(sheet, row, empObj, isRV);
      applyDashFormatting(sheet, row, empObj, isRV);
    });
    // Ensure trailing gap rows after seeding employees
    try { ensureBlockTrailingGap(sheet, blockStartRow, isRV); } catch (e) { Logger.log('ensureBlockTrailingGap error (seed): ' + e); }
  } catch (err) {
    Logger.log('seedEmployeesIntoBlock error: ' + err);
  }
}

// ── doPost ────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    const rvSheet   = ss.getSheetByName('RV')   || ss.insertSheet('RV');
    const comsSheet = ss.getSheetByName('COMS') || ss.insertSheet('COMS');
    const sheet     = data.department === 'rv' ? rvSheet : (data.department === 'coms' ? comsSheet : null);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown department' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    ensureSheetReady(sheet);

    if (data.type === 'dashboard')    updateDashboard(sheet, data.employees);
    if (data.type === 'weeklyReport') {
      // Ensure dashboard (employee rows) exist in the block before writing daily data
      const props = PropertiesService.getScriptProperties();
      const stateJson = props.getProperty('appdata_' + (data.department || 'rv'));
      if (stateJson) {
        const state = JSON.parse(stateJson);
        if (state.employees && state.employees.length > 0) {
          const isRV = sheet.getName() === 'RV';
          const now = new Date();
          // For each month in records, ensure employee rows exist in that block
          const months = {};
          (data.records || []).forEach(r => {
            if (!r.date) return;
            const d = new Date(r.date + 'T00:00:00');
            months[d.getFullYear() + '-' + d.getMonth()] = { y: d.getFullYear(), mo: d.getMonth() };
          });
          Object.values(months).forEach(({ y, mo }) => {
            const block = getOrCreateBlock(sheet, y, mo, isRV);
            const nameCol = 1 + getDashHeaders(isRV).length;
            const empRows = getEmpRowsInBlock(sheet, block.startRow, nameCol);
            if (empRows.length === 0) {
              // No employees yet in this block — seed them into THIS month block
              try {
                seedEmployeesIntoBlock(sheet, block.startRow, isRV, data.department, y, mo);
              } catch (e) {
                Logger.log('seedEmployeesIntoBlock (doPost) error: ' + e);
              }
            }
          });
        }
      }
      updateWeeklyReport(sheet, data.records);
    }
    if (data.type === 'deleteEmployee' && data.employeeName) {
      // blockYear / blockMonth (0-based month) from web app = only that month block + its daily columns
      // Avoid data.month from clients that send a string month name — use blockMonth only.
      if (typeof data.blockYear === 'number' && !isNaN(data.blockYear) &&
          typeof data.blockMonth === 'number' && !isNaN(data.blockMonth)) {
        // One month block only: row removal removes dashboard + daily cells for that month
        deleteEmployeeFromBlock(sheet, data.employeeName, data.blockYear, data.blockMonth);
      } else {
        deleteEmployeeFromAllBlocks(sheet, data.employeeName);
      }
    }
    if (data.type === 'forceSync' && (data.action === 'deleteEmployee' || data.action === 'deleteRecord')) {
      clearEmployeeWeeklyData(sheet, data.employeeName);
    }
    if (data.type === 'fullState' && data.state) {
      const dept = data.department || 'rv';
      const props = PropertiesService.getScriptProperties();
      const revKey = 'apprev_' + dept;
      const currentRevision = Number(props.getProperty(revKey) || '0');
      const hasBaseRevision = typeof data.baseRevision !== 'undefined' && data.baseRevision !== null;
      const baseRevision = Number(data.baseRevision);

      // Reject stale full-state writes (device edited older snapshot).
      if (hasBaseRevision && !isNaN(baseRevision) && baseRevision !== currentRevision) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'conflict',
          message: 'Stale state. Reload before syncing again.',
          department: dept,
          currentRevision: currentRevision
        })).setMimeType(ContentService.MimeType.JSON);
      }

      props.setProperty('appdata_' + dept, JSON.stringify(data.state));
      props.setProperty(revKey, String(currentRevision + 1));
      // Normalize whole sheet layout from current full state (same data, cleaned order/format).
      normalizeSheetLayoutFromState(sheet, dept, data.state);
    }

    // Import manually-edited weekly report cells back into PropertiesService state
    if (data.type === 'importFromSheet') {
      try {
        importFromSheetToState(sheet, data.department);
      } catch (ie) {
        Logger.log('importFromSheet error: ' + ie);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Synced' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── updateDashboard ───────────────────────────────────────────
// Writes PRESENT/ABSENT/LATE/… stats into the CURRENT month block.
// Each month block has its own fresh stats — previous months are untouched.

function updateDashboard(sheet, employees, targetYear, targetMonth) {
  // If targetYear/targetMonth are provided, write dashboard into that month block.
  if (!employees || employees.length === 0) return;
  ensureSheetReady(sheet);

  const isRV       = sheet.getName() === 'RV';
  const dashHdrs   = getDashHeaders(isRV);
  const dashCount  = dashHdrs.length;
  const nameCol    = 1 + dashCount; // 1-based: col B=2, so nameCol = dashCount+1

  const now = new Date();
  const y = (typeof targetYear !== 'undefined' && typeof targetMonth !== 'undefined') ? targetYear : now.getFullYear();
  const mo = (typeof targetYear !== 'undefined' && typeof targetMonth !== 'undefined') ? targetMonth : now.getMonth();

  const block = getOrCreateBlock(sheet, y, mo, isRV);
  const dataStart = block.startRow + 3;

  employees.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

  // Clear only the dashboard stat columns inside THIS block's data rows
  // (previous behavior cleared from this block to sheet end which wiped
  // later month blocks when updating the current month)
  const blocks = findMonthBlocks(sheet);
  let nextStart = sheet.getLastRow() + 1;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].startRow === block.startRow) {
      if (i + 1 < blocks.length) nextStart = blocks[i + 1].startRow;
      break;
    }
  }
  const clearCount = Math.max(0, nextStart - dataStart);
  const daysInMonth = new Date(y, mo + 1, 0).getDate();
  const numCols = dashCount + daysInMonth * 3 + 1; // stats + NAME + daily (×3) + TOTAL DAYS
  if (clearCount > 0) {
    sheet.getRange(dataStart, 2, clearCount, numCols).clearContent().clearFormat();
  }

  employees.forEach((emp, i) => {
    const row = dataStart + i;
    writeDashRow(sheet, row, emp, isRV);
    applyDashFormatting(sheet, row, emp, isRV);
  });

  // After adding rows, restore the 3 blank gap rows before the next month block (push next section down)
  try {
    ensureBlockTrailingGap(sheet, block.startRow, isRV);
  } catch (e) {
    Logger.log('ensureBlockTrailingGap after updateDashboard: ' + e);
  }
  try {
    applyCompactLayoutToBlock(sheet, block.startRow, isRV, y, mo);
  } catch (e) {
    Logger.log('applyCompactLayoutToBlock after updateDashboard: ' + e);
  }

  Logger.log('Dashboard updated: ' + employees.length + ' employees, block row ' + block.startRow + ' (' + y + '-' + mo + ')');
}

function writeDashRow(sheet, row, emp, isRV) {
  if (isRV) {
    sheet.getRange(row, 2).setValue(emp.present    || 0);
    sheet.getRange(row, 3).setValue(emp.absent     || 0);
    sheet.getRange(row, 4).setValue(emp.late       || 0);
    sheet.getRange(row, 5).setValue(emp.totalLates || 0);
    sheet.getRange(row, 6).setValue(emp.undertime  || 0);
    sheet.getRange(row, 7).setValue(emp.overtime   || 0);
    sheet.getRange(row, 8).setValue(emp.awol       || 0);
    sheet.getRange(row, 9).setValue(emp.sickLeave  || 0);
    sheet.getRange(row, 10).setValue(emp.wfh       || 0);
    sheet.getRange(row, 11).setValue(emp.scheduleDisplay || '');
    sheet.getRange(row, 12).setValue(emp.name      || '').setFontWeight('bold');
  } else {
    sheet.getRange(row, 2).setValue(emp.present    || 0);
    sheet.getRange(row, 3).setValue(emp.absent     || 0);
    sheet.getRange(row, 4).setValue(emp.late       || 0);
    sheet.getRange(row, 5).setValue(emp.undertime  || 0);
    sheet.getRange(row, 6).setValue(emp.awol       || 0);
    sheet.getRange(row, 7).setValue(emp.sickLeave  || 0);
    sheet.getRange(row, 8).setValue(emp.wfh        || 0);
    sheet.getRange(row, 9).setValue(emp.scheduleDisplay || '');
    sheet.getRange(row, 10).setValue(emp.name      || '').setFontWeight('bold');
  }
}

function applyDashFormatting(sheet, row, emp, isRV) {
  const numCols = isRV ? 11 : 9;
  sheet.getRange(row, 2, 1, numCols)
    .clearFormat()
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Name col — left align
  sheet.getRange(row, isRV ? 12 : 10).setHorizontalAlignment('left');

  const habitual = (emp.absent || 0) + (emp.late || 0) + (emp.undertime || 0);

  if ((emp.present || 0) > 0)
    sheet.getRange(row, 2).setBackground('#d1fae5').setFontColor('#065f46').setFontWeight('bold');

  const hColor = getHabitualColor(habitual);
  if (hColor) {
    if ((emp.absent    || 0) > 0) sheet.getRange(row, 3).setBackground(hColor).setFontColor('#000').setFontWeight('bold');
    if ((emp.late      || 0) > 0) sheet.getRange(row, 4).setBackground(hColor).setFontColor('#000').setFontWeight('bold');
    if ((emp.undertime || 0) > 0) sheet.getRange(row, isRV ? 6 : 5).setBackground(hColor).setFontColor('#000').setFontWeight('bold');
  }

  if ((emp.wfh || 0) > 0)
    sheet.getRange(row, isRV ? 10 : 8).setBackground('#D1FAE5').setFontColor('#065f46').setFontWeight('bold');

  const aColor = getAWOLColor(emp.awol || 0);
  if (aColor) {
    const ac = sheet.getRange(row, isRV ? 8 : 6);
    ac.setBackground(aColor).setFontColor('#000').setFontWeight('bold');
    if ((emp.awol || 0) >= 4) ac.setFontColor('#fff');
  }
}

// ── buildDashStats ───────────────────────────────────────────
// Computes dashboard stats from raw employees + attendanceData arrays
// (mirrors what script.js updateDashboard does client-side)
function buildDashStats(employees, attendanceData, dept, targetYear, targetMonth) {
  const now = new Date();
  const curMonth = (typeof targetMonth !== 'undefined') ? Number(targetMonth) : now.getMonth();
  const curYear  = (typeof targetYear !== 'undefined') ? Number(targetYear) : now.getFullYear();

  const stats = {};
  employees.filter(e => e.department === dept && employeeAppearsInSheetBlock(e, curYear, curMonth))
    .sort((a, b) => (a.name || '').localeCompare(b.name || ''))
    .forEach(e => {
      stats[e.name] = {
        name: e.name, present: 0, absent: 0, late: 0, totalLates: 0,
        undertime: 0, overtime: 0, awol: 0, sickLeave: 0, wfh: 0,
        scheduleDisplay: e.scheduleDisplay || ''
      };
    });

  attendanceData.forEach(r => {
    if (r.department !== dept || !stats[r.name]) return;
    if (!r.date) return;
    const d = new Date(r.date + 'T00:00:00');
    if (d.getMonth() !== curMonth || d.getFullYear() !== curYear) return;
    const s = stats[r.name];
    if (r.status === 'Present')        s.present++;
    else if (r.status === 'Absent')    s.absent++;
    else if (r.status === 'Late')      { s.late++; s.totalLates += getLateMinutesForDept(r, dept); }
    else if (r.status === 'Undertime') s.undertime++;
    else if (r.status === 'Overtime')  {
      // For COMS department, treat Overtime as a counted "present" day on the dashboard
      if (dept === 'coms') s.present++; else s.overtime++;
    }
    else if (r.status === 'AWOL')      s.awol++;
    else if (r.status === 'Sick Leave') s.sickLeave++;
    else if (r.status === 'Work From Home') s.wfh++;
  });

  return Object.values(stats);
}

// ── updateWeeklyReportFull ────────────────────────────────────
// Like updateWeeklyReport but ALSO clears every employee's daily cells
// for months that have NO records (handles the delete-all-records case).
function updateWeeklyReportFull(sheet, records) {
  ensureSheetReady(sheet);
  const isRV      = sheet.getName() === 'RV';
  const dashCount = getDashHeaders(isRV).length;
  const nameCol   = 1 + dashCount;

  // Build lookup: byMonth[mk][normalizedEmpName][dateStr] = record
  // Normalize names by trimming, collapsing whitespace and uppercasing to avoid mismatch
  const byMonth = {};
  (records || []).forEach(r => {
    if (!r.date || !r.name) return;
    const d  = new Date(r.date + 'T00:00:00');
    const mk = d.getFullYear() + '-' + d.getMonth();
    if (!byMonth[mk]) byMonth[mk] = {};
    const normName = r.name.toString().replace(/\s+/g, ' ').trim().toUpperCase();
    if (!byMonth[mk][normName]) byMonth[mk][normName] = {};
    byMonth[mk][normName][r.date] = r;
  });

  // Process every existing block in the sheet — clear+rewrite all
  findMonthBlocks(sheet).forEach(block => {
    const mk  = block.key; // "YYYY-M"
    const [y, mo] = mk.split('-').map(Number);
    try { writeBlockHeaders(sheet, block.startRow, y, mo, isRV); } catch (e) { Logger.log('writeBlockHeaders (weeklyFull) error: ' + e); }
    const daysInMonth   = new Date(y, mo + 1, 0).getDate();
    const dailyStartCol = 2 + dashCount;
    const monthData     = byMonth[mk] || {};

    getEmpRowsInBlock(sheet, block.startRow, nameCol).forEach(({ name, row }) => {
      const normSheetName = (name || '').toString().replace(/\s+/g, ' ').trim().toUpperCase();
      const empRecs = monthData[normSheetName] || {};
      let totalDays = 0;
      let col = dailyStartCol;

      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = y + '-' + String(mo + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        const rec = empRecs[dateStr];

        sheet.getRange(row, col, 1, 3)
          .clearContent().clearFormat()
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');

        if (rec) {
          const si   = getStatusInfo(rec.status);
          const code = si ? si.code : (rec.status || '');
          sheet.getRange(row, col).setValue(code);
          sheet.getRange(row, col + 1).setValue(fmtT(rec.timeIn));
          sheet.getRange(row, col + 2).setValue(fmtT(rec.timeOut));
          if (si) sheet.getRange(row, col, 1, 3).setBackground(si.bg).setFontColor(si.text).setFontWeight('bold');
          if (/(present|late|undertime|overtime)/i.test(rec.status || '')) totalDays++;
        }
        col += 3;
      }

      // TOTAL DAYS — clear if zero
      const tdCell = sheet.getRange(row, col);
      if (totalDays > 0) {
        tdCell.setValue(totalDays + (totalDays === 1 ? ' day' : ' days'))
          .setBackground('#D1FAE5').setFontWeight('bold')
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true, true, true, true, true, true);
      } else {
        tdCell.clearContent().clearFormat()
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      }
    });

    Logger.log('WeeklyFull written: ' + mk + ' — ' + getEmpRowsInBlock(sheet, block.startRow, nameCol).length + ' employees');
    try { applyCompactLayoutToBlock(sheet, block.startRow, isRV, y, mo); } catch (e) { Logger.log('applyCompactLayoutToBlock (weeklyFull) error: ' + e); }
  });
}

// ── updateWeeklyReport ────────────────────────────────────────
// Writes daily STATUS/TIME IN/TIME OUT into the correct month block.
// Each month in the records gets its own vertical block.

function updateWeeklyReport(sheet, records) {
  if (!sheet || !records || records.length === 0) return;
  ensureSheetReady(sheet);

  const isRV      = sheet.getName() === 'RV';
  const dashCount = getDashHeaders(isRV).length;
  const nameCol   = 1 + dashCount;

  // Group records: byMonth[YYYY-M][normalizedEmpName][dateStr] = record
  // Normalize names by trimming, collapsing whitespace and uppercasing to avoid mismatch
  const byMonth = {};
  records.forEach(r => {
    if (!r.date || !r.name) return;
    const d  = new Date(r.date + 'T00:00:00');
    const mk = d.getFullYear() + '-' + d.getMonth();
    if (!byMonth[mk]) byMonth[mk] = {};
    const normName = r.name.toString().replace(/\s+/g, ' ').trim().toUpperCase();
    if (!byMonth[mk][normName]) byMonth[mk][normName] = {};
    byMonth[mk][normName][r.date] = r;
  });

  Object.keys(byMonth).forEach(mk => {
    const [y, mo]  = mk.split('-').map(Number);
    const block    = getOrCreateBlock(sheet, y, mo, isRV);
    try { writeBlockHeaders(sheet, block.startRow, y, mo, isRV); } catch (e) { Logger.log('writeBlockHeaders (weekly) error: ' + e); }
    const dataStart = block.startRow + 3;
    const daysInMonth  = new Date(y, mo + 1, 0).getDate();
    const dailyStartCol = 2 + dashCount;

    const empRows = getEmpRowsInBlock(sheet, block.startRow, nameCol);

    empRows.forEach(({ name, row }) => {
      const normSheetName = (name || '').toString().replace(/\s+/g, ' ').trim().toUpperCase();
      const empRecs = byMonth[mk][normSheetName] || {};
      let totalDays = 0;
      let col = dailyStartCol;

      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = y + '-' + String(mo + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        const rec = empRecs[dateStr];

        sheet.getRange(row, col, 1, 3)
          .clearContent().clearFormat()
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');

        if (rec) {
          const si   = getStatusInfo(rec.status);
          const code = si ? si.code : (rec.status || '');
          sheet.getRange(row, col).setValue(code);
          sheet.getRange(row, col + 1).setValue(fmtT(rec.timeIn));
          sheet.getRange(row, col + 2).setValue(fmtT(rec.timeOut));
          if (si) {
            sheet.getRange(row, col, 1, 3)
              .setBackground(si.bg).setFontColor(si.text).setFontWeight('bold');
          }
          if (/(present|late|undertime|overtime)/i.test(rec.status || '')) totalDays++;
        }
        col += 3;
      }

      // TOTAL DAYS
      sheet.getRange(row, col)
        .setValue(totalDays > 0 ? totalDays + (totalDays === 1 ? ' day' : ' days') : '')
        .setBackground('#D1FAE5').setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true);
    });

    Logger.log('Weekly written: ' + mk + ' — ' + empRows.length + ' employees');
    try { applyCompactLayoutToBlock(sheet, block.startRow, isRV, y, mo); } catch (e) { Logger.log('applyCompactLayoutToBlock (weekly) error: ' + e); }
  });
}

// ── deleteEmployeeFromAllBlocks ──────────────────────────────
// Deletes the employee row in EVERY month block (removes row entirely)
// so both the stat columns and daily columns are gone.

function deleteEmployeeFromAllBlocks(sheet, employeeName) {
  if (!employeeName) return;
  const isRV    = sheet.getName() === 'RV';
  const nameCol = 1 + getDashHeaders(isRV).length; // 12 for RV, 10 for COMS

  // Collect rows to delete (descending order so row indices stay valid)
  const rowsToDelete = [];
  findMonthBlocks(sheet).forEach(block => {
    getEmpRowsInBlock(sheet, block.startRow, nameCol).forEach(({ name, row }) => {
      if (name.trim() === employeeName.trim()) rowsToDelete.push(row);
    });
  });

  // Delete from bottom to top
  rowsToDelete.sort((a, b) => b - a).forEach(row => {
    sheet.deleteRow(row);
    Logger.log('Deleted employee row ' + row + ': ' + employeeName);
  });

  // Ensure each month block has a 3-row trailing gap after deletions
  try {
    findMonthBlocks(sheet).forEach(b => ensureBlockTrailingGap(sheet, b.startRow, isRV));
  } catch (e) { Logger.log('ensureBlockTrailingGap error (all): ' + e); }
}

// Delete an employee row only within a specific month block (year, month are numeric, month is 0-based)
function deleteEmployeeFromBlock(sheet, employeeName, year, month) {
  if (!employeeName || typeof year === 'undefined' || typeof month === 'undefined') return;
  const isRV = sheet.getName() === 'RV';
  const nameCol = 1 + getDashHeaders(isRV).length;
  const key = year + '-' + month;
  const blocks = findMonthBlocks(sheet);
  const block = blocks.find(b => b.key === key);
  if (!block) return;

  const rowsToDelete = [];
  getEmpRowsInBlock(sheet, block.startRow, nameCol).forEach(({ name, row }) => {
    if (name.trim() === employeeName.trim()) rowsToDelete.push(row);
  });

  rowsToDelete.sort((a, b) => b - a).forEach(row => {
    sheet.deleteRow(row);
    Logger.log('Deleted employee row ' + row + ' from block ' + key + ': ' + employeeName);
  });

  // Ensure this block has a 3-row trailing gap after deletion
  try { ensureBlockTrailingGap(sheet, block.startRow, isRV); } catch (e) { Logger.log('ensureBlockTrailingGap error (block): ' + e); }
}

// ── deleteEmployeeData ────────────────────────────────────────

function deleteEmployeeData(sheet, employeeName) {
  if (!employeeName) return;
  const isRV    = sheet.getName() === 'RV';
  const nameCol = isRV ? 12 : 10;
  const lastRow = sheet.getLastRow();
  for (let i = 1; i <= lastRow; i++) {
    const v = (sheet.getRange(i, nameCol).getValue() || '').toString().trim();
    if (v === employeeName.trim()) {
      sheet.deleteRow(i);
      Logger.log('Deleted employee row: ' + employeeName);
      // After deletion, ensure block gaps are intact
      try { findMonthBlocks(sheet).forEach(b => ensureBlockTrailingGap(sheet, b.startRow, isRV)); } catch (e) { Logger.log('ensureBlockTrailingGap error (data): ' + e); }
      return;
    }
  }
}

// Ensure a given month block has at least a 3-row trailing gap after the last employee row
function ensureBlockTrailingGap(sheet, blockStartRow, isRV) {
  const dashCount = getDashHeaders(isRV).length;
  const nameCol = 1 + dashCount;
  const dataStart = blockStartRow + 3;
  const blocks = findMonthBlocks(sheet);
  // Find next block startRow or end of sheet
  let nextStart = sheet.getLastRow() + 1;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].startRow === blockStartRow) {
      if (i + 1 < blocks.length) nextStart = blocks[i + 1].startRow;
      break;
    }
  }

  // Find last non-empty name row in this block
  let lastNonEmpty = dataStart - 1;
  for (let r = dataStart; r < nextStart; r++) {
    const name = (sheet.getRange(r, nameCol).getValue() || '').toString().trim();
    if (name) lastNonEmpty = r;
  }

  const requiredGap = 3;
  const currentGap = nextStart - (lastNonEmpty + 1);
  if (currentGap < requiredGap) {
    const toInsert = requiredGap - currentGap;
    sheet.insertRows(nextStart, toInsert);
    Logger.log('Inserted ' + toInsert + ' gap rows at ' + nextStart + ' for block starting ' + blockStartRow);
    // Reapply block headers after inserting rows — inserting can split merged
    // header cells or clear formatting. Find the block metadata and rewrite
    // headers to ensure STATUS / TIME IN / TIME OUT and company header remain.
    try {
      const blocks = findMonthBlocks(sheet);
      const block = blocks.find(b => b.startRow === blockStartRow);
      if (block) {
        writeBlockHeaders(sheet, blockStartRow, block.year, block.monthNum, isRV);
      }
    } catch (e) {
      Logger.log('reapply headers after gap insert error: ' + e);
    }
  }
}

// ── clearEmployeeWeeklyData ───────────────────────────────────

function clearEmployeeWeeklyData(sheet, employeeName) {
  if (!employeeName) return;
  const isRV      = sheet.getName() === 'RV';
  const dashCount = getDashHeaders(isRV).length;
  const nameCol   = 1 + dashCount;

  findMonthBlocks(sheet).forEach(block => {
    getEmpRowsInBlock(sheet, block.startRow, nameCol).forEach(({ name, row }) => {
      if (name !== employeeName.trim()) return;
      const days  = new Date(block.year, block.monthNum + 1, 0).getDate();
      const start = 2 + dashCount;
      sheet.getRange(row, start, 1, days * 3 + 1).clearContent().clearFormat();
    });
  });
}

// ── policy color helpers ──────────────────────────────────────

function getAWOLColor(n) {
  if (n <= 0) return null;
  if (n === 1) return '#F4E4A6';
  if (n === 2) return '#E8C89A';
  if (n === 3) return '#D9956B';
  return '#C1503A';
}

function getHabitualColor(n) {
  if (n <= 0) return null;
  if (n === 1) return '#F4E4A6';
  if (n === 2) return '#D9C4A8';
  if (n === 3) return '#D9956B';
  if (n === 4) return '#D4B5A8';
  if (n === 5) return '#C98B7A';
  if (n === 6) return '#B85C52';
  if (n === 7) return '#A63D3D';
  if (n === 8) return '#8B2E2E';
  return '#3D3D3D';
}

function getStatusInfo(status) {
  const map = {
    'Present':            { bg: '#BFDBFE', text: '#1e3a8a', code: 'P'    },
    'Absent':             { bg: '#FCA5A5', text: '#7f1d1d', code: 'A'    },
    'Late':               { bg: '#FED7AA', text: '#7c2d12', code: 'L'    },
    'Undertime':          { bg: '#E9D5FF', text: '#581c87', code: 'UT'   },
    'Overtime':         { bg: '#E9D5FF', text: '#581c87', code: 'OT'   },
    'AWOL':               { bg: '#FCA5A5', text: '#7f1d1d', code: 'AWOL' },
    'Sick Leave':         { bg: '#BFDBFE', text: '#1e3a8a', code: 'LEAVE'   },
    'No Schedule':        { bg: '#BFDBFE', text: '#1e3a8a', code: 'NS'   },
    'Late and Undertime': { bg: '#FED7AA', text: '#7c2d12', code: 'L&UT' },
    'Work From Home':     { bg: '#D1FAE5', text: '#065f46', code: 'WFH'  }
  };
  return map[status] || null;
}

// ── doGet ─────────────────────────────────────────────────────

function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  const dept  = e && e.parameter && e.parameter.department ? e.parameter.department : null;
  if (dept) {
    const state = JSON.parse(props.getProperty('appdata_' + dept) || '{"employees":[],"attendanceData":[]}');
    state.revision = Number(props.getProperty('apprev_' + dept) || '0');
    return ContentService.createTextOutput(
      JSON.stringify(state)
    ).setMimeType(ContentService.MimeType.JSON);
  }
  const rv   = JSON.parse(props.getProperty('appdata_rv')   || '{"employees":[],"attendanceData":[]}');
  const coms = JSON.parse(props.getProperty('appdata_coms') || '{"employees":[],"attendanceData":[]}');
  rv.revision = Number(props.getProperty('apprev_rv') || '0');
  coms.revision = Number(props.getProperty('apprev_coms') || '0');
  return ContentService.createTextOutput(
    JSON.stringify({ rv, coms })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── menu ──────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance Monitor')
    .addItem('Force Reset Sheets', 'forceRecreateHeaders')
    .addItem('Set RV Company Name', 'setRVCompanyName')
    .addItem('Set COMS Company Name', 'setCOMSCompanyName')
    .addItem('Set RV Weekly Report Label', 'setRVWeeklyLabel')
    .addItem('Set COMS Weekly Report Label', 'setCOMSWeeklyLabel')
      .addItem('Add Karin (RV)', 'addKarinToRV')
    .addToUi();
}

  // Add a named employee into saved app state for a department if missing.
  function addEmployeeToState(dept, emp) {
    if (!dept || !emp || !emp.name) return;
    const props = PropertiesService.getScriptProperties();
    const key = 'appdata_' + dept;
    const stateJson = props.getProperty(key) || '{"employees":[],"attendanceData":[]}';
    let state = {};
    try { state = JSON.parse(stateJson); } catch (e) { state = { employees: [], attendanceData: [] }; }
    state.employees = state.employees || [];
    const exists = state.employees.some(e => (e.name || '').trim().toUpperCase() === (emp.name || '').trim().toUpperCase());
    if (exists) return;
    // Ensure basic fields
    const nowId = Date.now();
    const n = new Date();
    const toAdd = Object.assign({
      id: nowId,
      addedYear: n.getFullYear(),
      addedMonth: n.getMonth(),
      name: emp.name,
      scheduleStart: emp.scheduleStart || '',
      scheduleEnd: emp.scheduleEnd || '',
      scheduleDisplay: emp.scheduleDisplay || '',
      scheduleNotes: emp.scheduleNotes || '',
      weeklyDays: emp.weeklyDays || '',
      specialDays: emp.specialDays || [],
      department: dept
    }, emp);
    state.employees.push(toAdd);
    props.setProperty(key, JSON.stringify(state));
    // Update sheets immediately
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(dept.toUpperCase()) || ss.getSheetByName('RV');
    if (sheet) {
      const y = n.getFullYear();
      const mo = n.getMonth();
      updateDashboard(sheet, buildDashStats(state.employees, state.attendanceData || [], dept, y, mo), y, mo);
      updateWeeklyReportFull(sheet, state.attendanceData || []);
    }
  }

  function addKarinToRV() {
    const emp = {
      name: 'Karin',
      scheduleStart: '09:00',
      scheduleEnd: '18:00',
      scheduleDisplay: '9:00 - 6:00',
      scheduleNotes: 'MANUAL',
      weeklyDays: 'Mon,Tue,Wed,Thu,Fri',
      specialDays: [],
      department: 'rv'
    };
    addEmployeeToState('rv', emp);
    SpreadsheetApp.getUi().alert('Karin added to saved state (RV). Dashboard & Weekly Report updated.');
  }

function setRVCompanyName() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('companyName_rv') || 'RED VICTORY CONSUMERS GOODS TRADING';
  const result = ui.prompt('Set RV Company Name', 'Current: ' + current + '\n\nEnter new name:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    const name = result.getResponseText().trim();
    if (name) {
      props.setProperty('companyName_rv', name);
      ui.alert('Saved! RV company name set to: ' + name + '\n\nIt will appear on the next sync or reset.');
    }
  }
}

function setCOMSCompanyName() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('companyName_coms') || 'C. OPERATIONS MANAGEMENT SERVICES';
  const result = ui.prompt('Set COMS Company Name', 'Current: ' + current + '\n\nEnter new name:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    const name = result.getResponseText().trim();
    if (name) {
      props.setProperty('companyName_coms', name);
      ui.alert('Saved! COMS company name set to: ' + name + '\n\nIt will appear on the next sync or reset.');
    }
  }
}

function forceRecreateHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const depts = [
    { name: 'RV',   isRV: true  },
    { name: 'COMS', isRV: false }
  ];
  const now = new Date();

  depts.forEach(({ name, isRV }) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.clearFormats();
    sh.setFrozenRows(0);
    sh.setFrozenColumns(0);

    // Write sentinel in col A row 1
    sh.getRange(1, 1).setValue('V2');
    sh.setColumnWidth(1, 2);
    sh.getRange(1, 1).setFontColor('#ffffff').setBackground('#ffffff');

    // Write current month block starting at row 2
    sh.getRange(2, 1).setValue('MONTH_BLOCK:' + now.getFullYear() + '-' + now.getMonth());
    writeBlockHeaders(sh, 2, now.getFullYear(), now.getMonth(), isRV);
  });

  SpreadsheetApp.getUi().alert('Sheets reset with current month layout. Sync from the app to populate employee data.');
}

function setRVWeeklyLabel() { setWeeklyLabel_('rv'); }
function setCOMSWeeklyLabel() { setWeeklyLabel_('coms'); }

function setWeeklyLabel_(dept) {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const y = now.getFullYear(), mo = now.getMonth();
  const monthName = now.toLocaleDateString('en-US', { month: 'long' });
  const key = 'weeklyLabel_' + dept + '_' + y + '_' + mo;
  const current = props.getProperty(key) || ('WEEKLY REPORT — ' + monthName + ' ' + y);
  const result = ui.prompt('Set Weekly Report Label (' + dept.toUpperCase() + ')',
    'Current: ' + current + '\n\nEnter new label (leave blank to reset to default):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    const val = result.getResponseText().trim();
    if (val) {
      props.setProperty(key, val);
      ui.alert('Saved! Label set to: ' + val + '\n\nRun Force Reset Sheets to apply.');
    } else {
      props.deleteProperty(key);
      ui.alert('Reset to default: WEEKLY REPORT — ' + monthName + ' ' + y);
    }
  }
}
