function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let rvSheet = ss.getSheetByName('RV');
    let comsSheet = ss.getSheetByName('COMS');
    
    // Create sheets if they don't exist
    if (!rvSheet) {
      rvSheet = ss.insertSheet('RV');
    }
    
    if (!comsSheet) {
      comsSheet = ss.insertSheet('COMS');
    }
    
    const sheet = data.department === 'rv' ? rvSheet : (data.department === 'coms' ? comsSheet : null);
    
    if (!sheet) {
      Logger.log('Unknown department: ' + data.department);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Unknown department: ' + data.department
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Always ensure headers are present before any operation
    ensureHeadersExist(sheet);
    
    if (data.type === 'dashboard') {
      updateDashboard(sheet, data.employees);
    }
    
    if (data.type === 'weeklyReport') {
      updateWeeklyReport(sheet, data.records);
    }
    
    if (data.type === 'deleteEmployee') {
      deleteEmployeeData(sheet, data.employeeName);
    }
    
    // Handle force sync - complete refresh of both dashboard and weekly report
    if (data.type === 'forceSync') {
      Logger.log('Force sync requested for: ' + data.employeeName);
      if (data.action === 'deleteEmployee' || data.action === 'deleteRecord') {
        // Clear specific employee data and force refresh
        clearEmployeeWeeklyData(sheet, data.employeeName);
      }
    }

    // Save full app state for cross-device sharing
    if (data.employees !== undefined || data.records !== undefined || data.type === 'fullState') {
      const props = PropertiesService.getScriptProperties();
      const dept = data.department || 'rv';
      if (data.type === 'fullState' && data.state) {
        props.setProperty('appdata_' + dept, JSON.stringify(data.state));
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data synced successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function ensureHeadersExist(sheet) {
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const companyName = isRV ? 'RED VICTORY CONSUMERS GOODS TRADING' : 'C. OPERATIONS MANAGEMENT SERVICES';
  
  // Check if headers exist by looking for the company name in row 2
  const companyHeader = sheet.getRange(2, 2).getValue();
  const dashboardHeader = sheet.getRange(3, 2).getValue();
  const wfhHeader = sheet.getRange(3, isRV ? 10 : 8).getValue();
  
  // If headers don't exist or are corrupted, recreate them
  if (!companyHeader || companyHeader.toString() !== companyName || dashboardHeader !== 'PRESENT' || wfhHeader !== 'WORK FROM HOME') {
    Logger.log('Headers missing or corrupted, recreating for ' + sheetName);
    setupCompleteHeaders(sheet, companyName);
  } else {
    // Headers exist, just update the weekly report section for current month
    updateWeeklyReportHeaders(sheet);
  }
}

function setupCompleteHeaders(sheet, companyName) {
  // Clear everything first
  sheet.clear();
  sheet.clearFormats();
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  
  const isRV = companyName.includes('RED VICTORY');
  const headerColor = isRV ? '#4CAF50' : '#ef4444';
  
  // Company header (Row 2)
  sheet.getRange(2, 2, 1, 11).merge()
    .setValue(companyName)
    .setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(14);
  
  // Dashboard headers (Row 3)
  let dashboardHeaders;
  if (isRV) {
    dashboardHeaders = ['PRESENT', 'ABSENT', 'LATE', 'TOTAL LATES (MINS)', 'UNDERTIME', 'OVERTIME', 'AWOL', 'SICK LEAVE / VACATION LEAVE', 'WORK FROM HOME', 'SCHEDULE TIME', 'NAME'];
  } else {
    dashboardHeaders = ['PRESENT', 'ABSENT', 'LATE', 'UNDERTIME', 'AWOL', 'SICK LEAVE / VACATION LEAVE', 'WORK FROM HOME', 'SCHEDULE TIME', 'NAME'];
  }
  
  const headerRange = sheet.getRange(3, 2, 1, dashboardHeaders.length);
  headerRange.setValues([dashboardHeaders])
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  
  // Empty row for spacing (Row 4)
  sheet.getRange(4, 2, 1, dashboardHeaders.length).setBackground('#ffffff');
  
  // Setup weekly report headers
  setupWeeklyReportHeaders(sheet, dashboardHeaders.length);
  
  // Set frozen rows and columns
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(isRV ? 13 : 11);
  
  Logger.log('Complete headers setup completed for ' + sheet.getName());
}

function setupWeeklyReportHeaders(sheet, dashboardColCount) {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  
  // Find existing months and add current month if not exists
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  const currentMonthKey = year + '-' + month;
  
  if (!existingMonths.some(m => m.key === currentMonthKey)) {
    // Add new month section
    addNewMonthSection(sheet, dashboardColCount, year, month, existingMonths);
  }
}

function findExistingMonths(sheet, dashboardColCount) {
  const existingMonths = [];
  const weeklyStartCol = 2 + dashboardColCount;
  let col = weeklyStartCol;
  
  // Scan row 2 for existing month headers
  while (col <= sheet.getLastColumn()) {
    const headerValue = sheet.getRange(2, col).getValue();
    if (headerValue && headerValue.toString().includes('WEEKLY REPORT')) {
      const headerText = headerValue.toString();
      const monthMatch = headerText.match(/(\w+)\s+(\d{4})\s+WEEKLY REPORT/);
      if (monthMatch) {
        const monthName = monthMatch[1];
        const year = parseInt(monthMatch[2]);
        const monthNum = new Date(Date.parse(monthName + ' 1, 2000')).getMonth();
        const daysInMonth = new Date(year, monthNum + 1, 0).getDate();
        const totalCols = daysInMonth * 3 + 1;
        
        existingMonths.push({
          key: year + '-' + monthNum,
          startCol: col,
          endCol: col + totalCols - 1,
          monthName: monthName,
          year: year,
          monthNum: monthNum
        });
        
        col += totalCols;
      } else {
        col++;
      }
    } else {
      col++;
    }
  }
  
  return existingMonths;
}

function addNewMonthSection(sheet, dashboardColCount, year, month, existingMonths) {
  const monthName = new Date(year, month, 1).toLocaleDateString('en-US', { month: 'long' }).toUpperCase();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const totalWeeklyCols = daysInMonth * 3 + 1;
  
  // Calculate starting column for new month
  let startCol = 2 + dashboardColCount;
  if (existingMonths.length > 0) {
    const lastMonth = existingMonths[existingMonths.length - 1];
    startCol = lastMonth.endCol + 1;
  }
  
  // Monthly header (Row 2)
  sheet.getRange(2, startCol, 1, totalWeeklyCols).merge()
    .setValue(monthName + ' ' + year + ' WEEKLY REPORT')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(14)
    .setBackground('#BFDBFE');
  
  // Daily headers (Row 3)
  let col = startCol;
  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = monthName + ' ' + day;
    
    // Merge 3 columns for each day
    sheet.getRange(3, col, 1, 3).merge()
      .setValue(dateStr)
      .setFontWeight('bold')
      .setBackground('#BFDBFE')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    
    // Sub-headers for STATUS, TIME IN, TIME OUT (Row 4)
    sheet.getRange(4, col).setValue('STATUS')
      .setFontWeight('bold')
      .setBackground('#E3F2FD')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    
    sheet.getRange(4, col + 1).setValue('TIME IN')
      .setFontWeight('bold')
      .setBackground('#E3F2FD')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(4, col + 2).setValue('TIME OUT')
      .setFontWeight('bold')
      .setBackground('#E3F2FD')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    
    col += 3;
  }
  
  // TOTAL DAYS column (Rows 3-4 merged)
  sheet.getRange(3, col, 2, 1).merge()
    .setValue('TOTAL DAYS')
    .setFontWeight('bold')
    .setBackground('#D1FAE5')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  
  Logger.log('Added new month section: ' + monthName + ' ' + year + ' starting at column ' + startCol);
}

function updateWeeklyReportHeaders(sheet) {
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const dashboardColCount = isRV ? 11 : 9;
  
  // Check if current month section exists, if not add it
  const now = new Date();
  const currentMonth = now.toLocaleDateString('en-US', { month: 'long' }).toUpperCase();
  const currentYear = now.getFullYear();
  const currentMonthKey = currentYear + '-' + now.getMonth();
  
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  const hasCurrentMonth = existingMonths.some(m => m.key === currentMonthKey);
  
  if (!hasCurrentMonth) {
    Logger.log('Adding new month section: ' + currentMonth + ' ' + currentYear);
    addNewMonthSection(sheet, dashboardColCount, currentYear, now.getMonth(), existingMonths);
  }
}

function updateDashboard(sheet, employees) {
  if (!employees || employees.length === 0) {
    Logger.log('No employees data to update');
    return;
  }
  
  // Ensure headers exist before updating data
  ensureHeadersExist(sheet);
  
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const numCols = isRV ? 11 : 9;
  const dataStartRow = 5; // Data starts from row 5
  
  // Clear existing employee data AND formatting (preserve headers)
  const lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    const numColsToName = isRV ? 12 : 10; // Include name column
    const clearRange = sheet.getRange(dataStartRow, 2, lastRow - dataStartRow + 1, numColsToName);
    clearRange.clearContent();
    clearRange.clearFormat();
    
    // DON'T clear monthly data when updating dashboard - only clear dashboard columns
    // Monthly data should only be cleared by updateWeeklyReport function
  }
  
  // Add employee data
  employees.forEach((emp, index) => {
    const row = dataStartRow + index;
    
    if (isRV) {
      // RV format: PRESENT, ABSENT, LATE, TOTAL LATES, UNDERTIME, OVERTIME, AWOL, SICK LEAVE, WFH, SCHEDULE TIME, NAME
      sheet.getRange(row, 2).setValue(emp.present || 0);
      sheet.getRange(row, 3).setValue(emp.absent || 0);
      sheet.getRange(row, 4).setValue(emp.late || 0);
      sheet.getRange(row, 5).setValue(emp.totalLates || 0);
      sheet.getRange(row, 6).setValue(emp.undertime || 0);
      sheet.getRange(row, 7).setValue(emp.overtime || 0);
      sheet.getRange(row, 8).setValue(emp.awol || 0);
      sheet.getRange(row, 9).setValue(emp.sickLeave || 0);
      sheet.getRange(row, 10).setValue(emp.wfh || 0);
      sheet.getRange(row, 11).setValue(emp.scheduleDisplay || '');
      sheet.getRange(row, 12).setValue(emp.name || '');
    } else {
      // COMS format: PRESENT, ABSENT, LATE, UNDERTIME, AWOL, SICK LEAVE, WFH, SCHEDULE TIME, NAME
      sheet.getRange(row, 2).setValue(emp.present || 0);
      sheet.getRange(row, 3).setValue(emp.absent || 0);
      sheet.getRange(row, 4).setValue(emp.late || 0);
      sheet.getRange(row, 5).setValue(emp.undertime || 0);
      sheet.getRange(row, 6).setValue(emp.awol || 0);
      sheet.getRange(row, 7).setValue(emp.sickLeave || 0);
      sheet.getRange(row, 8).setValue(emp.wfh || 0);
      sheet.getRange(row, 9).setValue(emp.scheduleDisplay || '');
      sheet.getRange(row, 10).setValue(emp.name || '');
    }
    
    // Apply formatting and colors
    applyEmployeeRowFormatting(sheet, row, emp, isRV);
  });
  
  Logger.log('Dashboard updated with ' + employees.length + ' employees - dashboard data refreshed, monthly data preserved');
}

function applyEmployeeRowFormatting(sheet, row, emp, isRV) {
  const numCols = isRV ? 12 : 10;
  
  // Clear all existing formatting first
  sheet.getRange(row, 2, 1, numCols).clearFormat();
  
  // Apply borders and alignment to all cells
  sheet.getRange(row, 2, 1, numCols)
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Apply policy colors ONLY if values are greater than 0
  const habitualCount = (emp.absent || 0) + (emp.late || 0) + (emp.undertime || 0);
  
  // Present - green background ONLY if actually present
  if ((emp.present || 0) > 0) {
    sheet.getRange(row, 2).setBackground('#d1fae5').setFontColor('#065f46').setFontWeight('bold');
  }
  
  // Habitual policy colors (Absent, Late, Undertime) ONLY if values > 0
  const habitualColor = getHabitualPolicyColor(habitualCount);
  if (habitualColor && habitualCount > 0) {
    if ((emp.absent || 0) > 0) {
      sheet.getRange(row, 3).setBackground(habitualColor).setFontColor('#000000').setFontWeight('bold');
    }
    if ((emp.late || 0) > 0) {
      sheet.getRange(row, 4).setBackground(habitualColor).setFontColor('#000000').setFontWeight('bold');
    }
    if ((emp.undertime || 0) > 0) {
      const undertimeCol = isRV ? 6 : 5;
      sheet.getRange(row, undertimeCol).setBackground(habitualColor).setFontColor('#000000').setFontWeight('bold');
    }
  }

  // WFH color (green) ONLY if WFH > 0
  if ((emp.wfh || 0) > 0) {
    const wfhCol = isRV ? 10 : 8;
    sheet.getRange(row, wfhCol).setBackground('#D1FAE5').setFontColor('#065f46').setFontWeight('bold');
  }
  
  // AWOL policy colors ONLY if AWOL > 0
  const awolColor = getAWOLPolicyColor(emp.awol || 0);
  if (awolColor && (emp.awol || 0) > 0) {
    const awolCol = isRV ? 8 : 6;
    sheet.getRange(row, awolCol).setBackground(awolColor).setFontColor('#000000').setFontWeight('bold');
    if ((emp.awol || 0) >= 4) {
      sheet.getRange(row, awolCol).setFontColor('#ffffff');
    }
  }
}

function updateWeeklyReport(sheet, records) {
  if (!sheet) {
    Logger.log('updateWeeklyReport: sheet is undefined, skipping');
    return;
  }
  if (!records || records.length === 0) {
    Logger.log('No records to update in weekly report - clearing all monthly data');
    // If no records, clear all monthly report data but keep headers
    clearAllMonthlyData(sheet);
    return;
  }
  
  // Ensure headers exist before updating data
  ensureHeadersExist(sheet);
  
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const dashboardColCount = isRV ? 11 : 9;
  const nameCol = isRV ? 12 : 10;
  const dataStartRow = 5;
  
  // Group records by employee and month
  const employeeRecords = {};
  records.forEach(record => {
    if (!employeeRecords[record.name]) {
      employeeRecords[record.name] = {};
    }
    employeeRecords[record.name][record.date] = record;
  });
  
  // Find all employees in the dashboard
  const lastRow = sheet.getLastRow();
  const employees = [];
  
  for (let i = dataStartRow; i <= lastRow; i++) {
    const name = sheet.getRange(i, nameCol).getValue();
    if (name && name.toString().trim() !== '') {
      employees.push({ name: name.toString().trim(), row: i });
    }
  }
  
  if (employees.length === 0) {
    Logger.log('No employees found in dashboard to update weekly report');
    return;
  }
  
  // Get existing months
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  
  // Clear only current month data first, then update
  clearCurrentMonthWeeklyData(sheet);
  
  // Update weekly report for each employee and each month
  employees.forEach(emp => {
    const row = emp.row;
    
    existingMonths.forEach(monthInfo => {
      const year = monthInfo.year;
      const month = monthInfo.monthNum;
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      
      let col = monthInfo.startCol;
      let totalDays = 0;
      
      // Fill in daily data for this month
      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = year + '-' + String(month + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        const record = employeeRecords[emp.name] ? employeeRecords[emp.name][dateStr] : null;
        
        if (record) {
          const statusInfo = getStatusDisplayInfo(record.status);
          const statusCode = statusInfo ? statusInfo.code : (record.status || '');
          
          // Set status, time in, and time out
          const fmtT = t => {
            if (!t) return '';
            const [h, m] = t.split(':');
            const hr = parseInt(h);
            return (hr % 12 || 12) + ':' + m + ' ' + (hr >= 12 ? 'PM' : 'AM');
          };
          sheet.getRange(row, col).setValue(statusCode);
          sheet.getRange(row, col + 1).setValue(fmtT(record.timeIn));
          sheet.getRange(row, col + 2).setValue(fmtT(record.timeOut));
          
          // Apply colors
          if (statusInfo) {
            sheet.getRange(row, col).setBackground(statusInfo.bg).setFontColor(statusInfo.text).setFontWeight('bold');
            sheet.getRange(row, col + 1).setBackground(statusInfo.bg).setFontColor(statusInfo.text).setFontWeight('bold');
            sheet.getRange(row, col + 2).setBackground(statusInfo.bg).setFontColor(statusInfo.text).setFontWeight('bold');
          }
          
          if (record.status !== 'Work From Home') totalDays++;
        }
        
        // Apply borders and alignment
        sheet.getRange(row, col, 1, 3)
          .setBorder(true, true, true, true, true, true)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        
        col += 3;
      }
      
      // Set TOTAL DAYS for this month
      sheet.getRange(row, col).setValue(totalDays > 0 ? totalDays + (totalDays === 1 ? ' day' : ' days') : '')
        .setBackground('#D1FAE5')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true);
    });
  });
  
  Logger.log('Weekly report updated for ' + employees.length + ' employees across ' + existingMonths.length + ' months');
}

function deleteEmployeeData(sheet, employeeName) {
  if (!employeeName) {
    Logger.log('No employee name provided for deletion');
    return;
  }
  
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const nameCol = isRV ? 12 : 10;
  const dataStartRow = 5;
  const lastRow = sheet.getLastRow();
  
  // Find and delete the employee row
  for (let i = dataStartRow; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, nameCol).getValue();
    if (cellValue && cellValue.toString().trim() === employeeName.trim()) {
      sheet.deleteRow(i);
      Logger.log('Deleted employee: ' + employeeName + ' from row ' + i);
      return;
    }
  }
  
  Logger.log('Employee not found for deletion: ' + employeeName);
}

// Clear weekly data for a specific employee across all months
function clearEmployeeWeeklyData(sheet, employeeName) {
  if (!employeeName) {
    Logger.log('No employee name provided for weekly data clearing');
    return;
  }
  
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const dashboardColCount = isRV ? 11 : 9;
  const nameCol = isRV ? 12 : 10;
  const dataStartRow = 5;
  const lastRow = sheet.getLastRow();
  
  // Get existing months
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  
  // Find the employee and clear their weekly data across all months
  for (let i = dataStartRow; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, nameCol).getValue();
    if (cellValue && cellValue.toString().trim() === employeeName.trim()) {
      // Clear weekly data for this employee across all months
      existingMonths.forEach(monthInfo => {
        const year = monthInfo.year;
        const month = monthInfo.monthNum;
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const totalWeeklyCols = (daysInMonth * 3) + 1;
        
        sheet.getRange(i, monthInfo.startCol, 1, totalWeeklyCols).clearContent().clearFormat();
        
        // Reapply borders and basic formatting
        for (let day = 1; day <= daysInMonth; day++) {
          const col = monthInfo.startCol + ((day - 1) * 3);
          sheet.getRange(i, col, 1, 3)
            .setBorder(true, true, true, true, true, true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        }
        
        // Format TOTAL DAYS column
        sheet.getRange(i, monthInfo.startCol + (daysInMonth * 3))
          .setBackground('#D1FAE5')
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBorder(true, true, true, true, true, true);
      });
      
      Logger.log('Cleared weekly data for employee: ' + employeeName + ' at row ' + i + ' across all months');
      return;
    }
  }
  
  Logger.log('Employee not found for weekly data clearing: ' + employeeName);
}

// Clear current month weekly data while preserving headers and other months
function clearCurrentMonthWeeklyData(sheet) {
  if (!sheet) {
    Logger.log('clearCurrentMonthWeeklyData: sheet is undefined, skipping');
    return;
  }
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const dashboardColCount = isRV ? 11 : 9;
  const dataStartRow = 5;
  const lastRow = sheet.getLastRow();
  
  if (lastRow < dataStartRow) {
    Logger.log('No data rows to clear');
    return;
  }
  
  // Get current month info
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  const currentMonthKey = currentYear + '-' + currentMonth;
  
  // Find current month section
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  const currentMonthInfo = existingMonths.find(m => m.key === currentMonthKey);
  
  if (currentMonthInfo) {
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
    const totalWeeklyCols = (daysInMonth * 3) + 1;
    
    // Clear current month weekly data for all employees
    sheet.getRange(dataStartRow, currentMonthInfo.startCol, lastRow - dataStartRow + 1, totalWeeklyCols)
      .clearContent()
      .clearFormat();
    
    Logger.log('Cleared current month weekly data from rows ' + dataStartRow + ' to ' + lastRow);
  } else {
    Logger.log('Current month section not found');
  }
}

// Clear all monthly sections data while preserving headers
function clearAllMonthlyData(sheet) {
  if (!sheet) {
    Logger.log('clearAllMonthlyData: sheet is undefined, skipping');
    return;
  }
  const sheetName = sheet.getName();
  const isRV = sheetName === 'RV';
  const dashboardColCount = isRV ? 11 : 9;
  const dataStartRow = 5;
  const lastRow = sheet.getLastRow();
  
  if (lastRow < dataStartRow) {
    Logger.log('No data rows to clear');
    return;
  }
  
  // Get all existing months
  const existingMonths = findExistingMonths(sheet, dashboardColCount);
  
  if (existingMonths.length > 0) {
    // Clear data for all monthly sections
    existingMonths.forEach(monthInfo => {
      const year = monthInfo.year;
      const month = monthInfo.monthNum;
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      const totalWeeklyCols = (daysInMonth * 3) + 1;
      
      // Clear all monthly data for all employees
      sheet.getRange(dataStartRow, monthInfo.startCol, lastRow - dataStartRow + 1, totalWeeklyCols)
        .clearContent()
        .clearFormat();
    });
    
    Logger.log('Cleared all monthly data from rows ' + dataStartRow + ' to ' + lastRow + ' across ' + existingMonths.length + ' months');
  } else {
    Logger.log('No monthly sections found to clear');
  }
}

// Policy color functions
function getAWOLPolicyColor(awolCount) {
  if (awolCount === 0) return null;
  if (awolCount === 1) return '#F4E4A6'; // Light yellow
  if (awolCount === 2) return '#E8C89A'; // Light orange
  if (awolCount === 3) return '#D9956B'; // Orange
  return '#C1503A'; // Dark red
}

function getHabitualPolicyColor(count) {
  if (count === 0) return null;
  if (count === 1) return '#F4E4A6'; // Light yellow
  if (count === 2) return '#D9C4A8'; // Light brown
  if (count === 3) return '#D9956B'; // Orange
  if (count === 4) return '#D4B5A8'; // Light red-brown
  if (count === 5) return '#C98B7A'; // Red-brown
  if (count === 6) return '#B85C52'; // Dark red-brown
  if (count === 7) return '#A63D3D'; // Dark red
  if (count === 8) return '#8B2E2E'; // Very dark red
  return '#3D3D3D'; // Almost black
}

function getStatusDisplayInfo(status) {
  const statusMap = {
    'Present': { bg: '#BFDBFE', text: '#1e3a8a', code: 'P' },
    'Absent': { bg: '#FCA5A5', text: '#7f1d1d', code: 'A' },
    'Late': { bg: '#FED7AA', text: '#7c2d12', code: 'L' },
    'Undertime': { bg: '#E9D5FF', text: '#581c87', code: 'UT' },
    'AWOL': { bg: '#FCA5A5', text: '#7f1d1d', code: 'AWOL' },
    'Sick Leave': { bg: '#BFDBFE', text: '#1e3a8a', code: 'SL' },
    'No Schedule': { bg: '#BFDBFE', text: '#1e3a8a', code: 'NS' },
    'Late and Undertime': { bg: '#FED7AA', text: '#7c2d12', code: 'L&UT' },
    'Work From Home': { bg: '#D1FAE5', text: '#065f46', code: 'WF' }
  };
  return statusMap[status] || null;
}

function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  const dept = (e && e.parameter && e.parameter.department) ? e.parameter.department : null;

  if (dept) {
    const data = props.getProperty('appdata_' + dept) || '{"employees":[],"attendanceData":[]}';
    return ContentService.createTextOutput(data).setMimeType(ContentService.MimeType.JSON);
  }

  // Return both departments
  const rv   = props.getProperty('appdata_rv')   || '{"employees":[],"attendanceData":[]}';
  const coms = props.getProperty('appdata_coms') || '{"employees":[],"attendanceData":[]}';
  return ContentService.createTextOutput(JSON.stringify({ rv: JSON.parse(rv), coms: JSON.parse(coms) }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Add menu to Google Sheets UI
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance Monitor')
    .addItem('Update Headers (Add WFH Column)', 'updateHeadersFromMenu')
    .addItem('Force Recreate All Headers', 'forceRecreateHeaders')
    .addItem('Force Complete Refresh', 'forceCompleteRefresh')
    .addToUi();
}

// Update headers from menu — adds WFH column if missing, preserves data
function updateHeadersFromMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { sheet: ss.getSheetByName('RV'),   company: 'RED VICTORY CONSUMERS GOODS TRADING',  isRV: true  },
    { sheet: ss.getSheetByName('COMS'), company: 'C. OPERATIONS MANAGEMENT SERVICES',     isRV: false }
  ];

  sheets.forEach(({ sheet, company, isRV }) => {
    if (!sheet) return;

    // Expected dashboard headers
    const expectedHeaders = isRV
      ? ['PRESENT', 'ABSENT', 'LATE', 'TOTAL LATES (MINS)', 'UNDERTIME', 'OVERTIME', 'AWOL', 'SICK LEAVE / VACATION LEAVE', 'WORK FROM HOME', 'SCHEDULE TIME', 'NAME']
      : ['PRESENT', 'ABSENT', 'LATE', 'UNDERTIME', 'AWOL', 'SICK LEAVE / VACATION LEAVE', 'WORK FROM HOME', 'SCHEDULE TIME', 'NAME'];

    // Read current row 3 headers starting at col 2
    const headerCount = expectedHeaders.length;
    const currentHeaders = sheet.getRange(3, 2, 1, headerCount).getValues()[0];

    // Check if headers already match
    const headersMatch = expectedHeaders.every((h, i) => currentHeaders[i] === h);
    if (!headersMatch) {
      // Overwrite row 3 dashboard headers
      sheet.getRange(3, 2, 1, headerCount)
        .setValues([expectedHeaders])
        .setFontWeight('bold')
        .setBackground('#f3f3f3')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true);
      Logger.log('Updated dashboard headers for ' + sheet.getName());
    }

    // Also ensure company header row 2 is correct
    const headerColor = isRV ? '#4CAF50' : '#ef4444';
    sheet.getRange(2, 2, 1, headerCount).merge()
      .setValue(company)
      .setBackground(headerColor)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(14);

    // Ensure weekly report month section exists
    updateWeeklyReportHeaders(sheet);
  });

  SpreadsheetApp.getUi().alert('Headers updated successfully for RV and COMS sheets!');
}

// Manual function to force recreate all headers
function forceRecreateHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let rvSheet = ss.getSheetByName('RV');
  let comsSheet = ss.getSheetByName('COMS');
  
  if (!rvSheet) {
    rvSheet = ss.insertSheet('RV');
  }
  
  if (!comsSheet) {
    comsSheet = ss.insertSheet('COMS');
  }
  
  Logger.log('Force recreating RV headers...');
  setupCompleteHeaders(rvSheet, 'RED VICTORY CONSUMERS GOODS TRADING');
  
  Logger.log('Force recreating COMS headers...');
  setupCompleteHeaders(comsSheet, 'C. OPERATIONS MANAGEMENT SERVICES');
  
  Logger.log('All headers recreated successfully!');
  SpreadsheetApp.getUi().alert('Headers recreated successfully for both RV and COMS sheets!');
  
  return 'Headers recreated successfully!';
}

// Force complete data refresh - clears everything and rebuilds
function forceCompleteRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let rvSheet = ss.getSheetByName('RV');
  let comsSheet = ss.getSheetByName('COMS');
  
  if (rvSheet) {
    Logger.log('Completely clearing and refreshing RV sheet...');
    rvSheet.clear();
    rvSheet.clearFormats();
    setupCompleteHeaders(rvSheet, 'RED VICTORY CONSUMERS GOODS TRADING');
  }
  
  if (comsSheet) {
    Logger.log('Completely clearing and refreshing COMS sheet...');
    comsSheet.clear();
    comsSheet.clearFormats();
    setupCompleteHeaders(comsSheet, 'C. OPERATIONS MANAGEMENT SERVICES');
  }
  
  Logger.log('Complete refresh finished!');
  SpreadsheetApp.getUi().alert('Complete refresh finished! All data and formatting cleared.');
  
  return 'Complete refresh finished!';
}