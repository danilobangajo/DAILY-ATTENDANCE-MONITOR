// Storage wrapper to handle tracking prevention
const storage = (() => {
    const mem = {};
    let _ls = null;
    // Test localStorage once silently
    try {
        localStorage.setItem('__test__', '1');
        localStorage.removeItem('__test__');
        _ls = localStorage;
    } catch (e) { /* blocked by tracking prevention, fall back to in-memory */ }
    return {
        getItem(key) {
            return _ls ? _ls.getItem(key) : (mem[key] ?? null);
        },
        setItem(key, value) {
            if (_ls) { try { _ls.setItem(key, value); return; } catch(e) { _ls = null; } }
            mem[key] = value;
        }
    };
})();

// Note: syncToGoogleSheets function is defined in config.js

async function syncDashboardToSheets() {
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const attendanceData = loadAttendanceData();
    
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    const deptEmployees = employees.filter(e => e.department === currentDepartment)
        .sort((a, b) => a.name.localeCompare(b.name));
    
    const employeeStats = deptEmployees.map(emp => {
        const stats = {
            name: emp.name,
            present: 0,
            absent: 0,
            late: 0,
            totalLates: 0,
            undertime: 0,
            overtime: 0,
            awol: 0,
            sickLeave: 0,
            wfh: 0,
            scheduleDisplay: emp.scheduleDisplay
        };
        
        attendanceData.forEach(record => {
            if (record.name !== emp.name || record.department !== currentDepartment) return;
            const d = new Date(record.date + 'T00:00:00');
            if (d.getMonth() !== currentMonth || d.getFullYear() !== currentYear) return;
            switch(record.status) {
                case 'Present': stats.present++; break;
                case 'Absent': stats.absent++; break;
                case 'Late': stats.late++; stats.totalLates += Number(record.lateMinutes) || 0; break;
                case 'Undertime': stats.undertime++; break;
                case 'Overtime': if (currentDepartment === 'rv') stats.overtime++; else stats.present++; break;
                case 'AWOL': stats.awol++; break;
                case 'Sick Leave': stats.sickLeave++; break;
                case 'Work From Home': stats.wfh++; break;
            }
        });
        
        return stats;
    });
    
    return syncToGoogleSheets('dashboard', { employees: employeeStats });
}

async function syncWeeklyReportToSheets() {
    const attendanceData = loadAttendanceData();
    const deptRecords = attendanceData.filter(r => r.department === currentDepartment);
    return syncToGoogleSheets('weeklyReport', { records: deptRecords });
}

// Get policy color based on AWOL count
function getAWOLPolicyColor(awolCount) {
    if (awolCount === 0) return '';
    if (awolCount === 1) return 'policy-color-1';
    if (awolCount === 2) return 'policy-color-2';
    if (awolCount === 3) return 'policy-color-3';
    return 'policy-color-4'; // 4 or more
}

// Get policy color based on Absent/Late count (Habitual)
function getHabitualPolicyColor(count) {
    if (count === 0) return '';
    if (count === 1) return 'policy-color-5';
    if (count === 2) return 'policy-color-6';
    if (count === 3) return 'policy-color-7';
    if (count === 4) return 'policy-color-8';
    if (count === 5) return 'policy-color-9';
    if (count === 6) return 'policy-color-10';
    if (count === 7) return 'policy-color-11';
    if (count === 8) return 'policy-color-12';
    return 'policy-color-13'; // 9 or more
}

// Helper: get local date in ISO (YYYY-MM-DD) accounting for timezone offset
function getLocalISODate() {
    const d = new Date();
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().split('T')[0];
}

/** YYYY-MM for input type="month" */
function getLocalMonthInputValue(d = new Date()) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
}

// Current filter date (use local date)
let currentFilterDate = getLocalISODate();
let currentFilterMode = 'day'; // 'day' or 'week'
let currentWeek = null;

// Filter by week
function filterByWeek(weekNumber) {
    currentFilterMode = 'week';
    currentWeek = weekNumber;
    
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();
    
    // Calculate week range (assuming 4 weeks per month, ~7 days each)
    const startDay = (weekNumber - 1) * 7 + 1;
    const endDay = Math.min(weekNumber * 7, new Date(year, month + 1, 0).getDate());
    
    // Update display
    const dateElement = document.getElementById('selectedDate');
    const monthName = now.toLocaleDateString('en-US', { month: 'long' });
    dateElement.textContent = `Week ${weekNumber} (${monthName} ${startDay}-${endDay}, ${year})`;
    
    // Clear date input
    document.getElementById('reportDate').value = '';
    
    updateDailyReportByWeek(weekNumber, year, month);
}

// Filter by week from select dropdown
function filterByWeekSelect() {
    const weekSelect = document.getElementById('weekFilter');
    const weekSelectModal = document.getElementById('weekFilterModal');
    const weekNumber = parseInt(weekSelect?.value || weekSelectModal?.value);
    
    if (weekNumber) {
        // Get selected month and year from calendar input
        const monthYearSelect = document.getElementById('monthYearSelect');
        const selectedValue = monthYearSelect.value; // Format: "2026-03"
        
        if (!selectedValue) return;
        
        const [selectedYear, selectedMonthStr] = selectedValue.split('-');
        const selectedMonth = parseInt(selectedMonthStr) - 1; // Convert to 0-based month
        const year = parseInt(selectedYear);
        
        // Calculate week range
        const startDay = (weekNumber - 1) * 7 + 1;
        const endDay = Math.min(weekNumber * 7, new Date(year, selectedMonth + 1, 0).getDate());
        
        // Update display
        const dateElement = document.getElementById('selectedDateModal');
        if (dateElement) {
            const date = new Date(year, selectedMonth);
            const monthName = date.toLocaleDateString('en-US', { month: 'long' });
            dateElement.textContent = `Week ${weekNumber} (${monthName} ${startDay}-${endDay}, ${year})`;
        }
        
        updateWeeklyReportWithFilter(selectedMonth, year, weekNumber);
    } else {
        // Show all for selected month/year if no week selected
        updateWeeklyReportByMonthYear();
    }
}

// Toggle Daily Report visibility
function toggleDailyReport() {
    const content = document.getElementById('dailyReportContent');
    const icon = document.getElementById('toggleReportIcon');
    const btn = document.getElementById('toggleReportBtn');
    const dateDisplay = document.getElementById('selectedDate').parentElement;
    
    if (content.style.display === 'none') {
        content.style.display = 'block';
        dateDisplay.style.display = 'block';
        icon.className = 'bi bi-dash-lg';
        btn.title = 'Minimize';
    } else {
        content.style.display = 'none';
        dateDisplay.style.display = 'none';
        icon.className = 'bi bi-plus-lg';
        btn.title = 'Maximize';
    }
}

// Open Daily Report in modal (zoom view)
function openReportModal() {
    const modalTable = document.getElementById('modalDailyReportTable');
    const mainTable = document.getElementById('dailyReportTable');
    const selectedDate = document.getElementById('selectedDate').textContent;
    const modalSelectedDate = document.getElementById('modalSelectedDate');
    
    // Copy table content
    modalTable.innerHTML = mainTable.innerHTML;
    modalSelectedDate.textContent = selectedDate;
    
    // Open modal
    const modal = new bootstrap.Modal(document.getElementById('reportZoomModal'));
    modal.show();
}

// Open Attendance Monitor in modal (zoom view)
function openMonitorModal() {
    const modalTableHead = document.getElementById('modalMonitorTableHead');
    const modalTableBody = document.getElementById('modalMonitorTableBody');
    const mainTableHead = document.querySelector('.dashboard-table thead');
    const mainTableBody = document.getElementById('dashboardTable');
    const companyName = document.getElementById('companyName').textContent;
    const modalCompanyName = document.getElementById('modalCompanyName');
    
    // Copy table content
    modalTableHead.innerHTML = mainTableHead.innerHTML;
    modalTableBody.innerHTML = mainTableBody.innerHTML;
    modalCompanyName.textContent = companyName;
    
    // Open modal
    const modal = new bootstrap.Modal(document.getElementById('monitorZoomModal'));
    modal.show();
}

// Update employee name datalist for autocomplete
function updateEmployeeDatalist() {
    const datalist = document.getElementById('employeeList');
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const attendanceData = loadAttendanceData();
    const today = getLocalISODate();
    
    // Filter employees by current department
    const departmentEmployees = employees.filter(emp => emp.department === currentDepartment);
    
    // Get today's time in records (without timeout)
    // Also include previous day for overnight shifts (e.g. 7AM-1AM)
    const prevDay = (() => {
        const d = new Date(today + 'T00:00:00');
        d.setDate(d.getDate() - 1);
        const y = d.getFullYear();
        const mo = String(d.getMonth() + 1).padStart(2, '0');
        const dy = String(d.getDate()).padStart(2, '0');
        return `${y}-${mo}-${dy}`;
    })();
    const todayTimeInRecords = attendanceData.filter(record => 
        (record.date === today || record.date === prevDay) && 
        record.department === currentDepartment && 
        record.timeIn && 
        !record.timeOut
    );
    
    // Clear and populate datalist
    datalist.innerHTML = '';
    
    // Add employees with existing time in records first
    todayTimeInRecords.forEach(record => {
        const option = document.createElement('option');
        option.value = record.name;
        option.setAttribute('data-has-timein', 'true');
        option.setAttribute('data-timein', record.timeIn);
        option.setAttribute('data-status', record.status);
        option.setAttribute('data-record-id', record.id);
        datalist.appendChild(option);
    });
    
    // Sort alphabetically
    departmentEmployees.sort((a, b) => a.name.localeCompare(b.name));

    // Add other employees
    departmentEmployees.forEach(emp => {
        // Skip if already has time in record today
        const hasTimeIn = todayTimeInRecords.some(record => record.name === emp.name);
        if (!hasTimeIn) {
            const option = document.createElement('option');
            option.value = emp.name;
            datalist.appendChild(option);
        }
    });
}

// Filter by date
function filterByDate() {
    currentFilterMode = 'day';
    currentWeek = null;
    const dateInput = document.getElementById('reportDate');
    currentFilterDate = dateInput.value;
    
    // Clear week select
    document.getElementById('weekFilter').value = '';
    
    updateSelectedDateDisplay();
    updateDailyReport();
}

// Update daily report by week
function updateDailyReportByWeek(weekNumber, year, month) {
    const tableBody = document.getElementById('dailyReportTable');
    const attendanceData = loadAttendanceData();
    
    // Calculate week range
    const startDay = (weekNumber - 1) * 7 + 1;
    const endDay = Math.min(weekNumber * 7, new Date(year, month + 1, 0).getDate());
    
    // Filter records within the week range
    const filteredRecords = attendanceData.filter(record => {
        if (record.department !== currentDepartment) return false;
        
        const recordDate = new Date(record.date + 'T00:00:00');
        const recordDay = recordDate.getDate();
        const recordMonth = recordDate.getMonth();
        const recordYear = recordDate.getFullYear();
        
        return recordYear === year && recordMonth === month && recordDay >= startDay && recordDay <= endDay;
    });
    
    if (filteredRecords.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="6" class="text-center text-muted py-4">
                    <i class="bi bi-inbox fs-1 d-block mb-2"></i>
                    No attendance records for selected week
                </td>
            </tr>
        `;
        return;
    }
    
    filteredRecords.sort((a, b) => a.name.localeCompare(b.name));
    tableBody.innerHTML = '';
    filteredRecords.forEach(record => {
        const safeId = parseInt(record.id, 10);
        const isLate = record.status === 'Late';
        const tr = document.createElement('tr');

        const tdName = document.createElement('td');
        tdName.textContent = record.name;

        const tdStatus = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `badge-status ${getStatusBadgeClass(record.status)}`;
        badge.textContent = getDisplayStatus(record.status);
        tdStatus.appendChild(badge);

        const tdTimeIn = document.createElement('td');
        if (isLate) { tdTimeIn.className = 'text-warning fw-bold'; }
        tdTimeIn.textContent = formatTime(record.timeIn);

        const tdSched = document.createElement('td');
        tdSched.textContent = record.scheduleTime || '-';

        const tdDate = document.createElement('td');
        tdDate.textContent = record.date;

        const tdReason = document.createElement('td');
        if (record.reason) {
            const badge = document.createElement('span');
            badge.className = 'badge bg-warning text-dark';
            badge.textContent = record.reason;
            tdReason.appendChild(badge);
        } else {
            tdReason.textContent = '-';
        }

        const tdActions = document.createElement('td');
        tdActions.className = 'text-center';
        const editBtn = document.createElement('button');
        editBtn.className = 'btn btn-sm btn-outline-primary';
        editBtn.title = 'Edit';
        editBtn.innerHTML = '<i class="bi bi-pencil"></i>';
        editBtn.addEventListener('click', () => editAttendanceRecord(safeId));
        const delBtn = document.createElement('button');
        delBtn.className = 'btn btn-sm btn-outline-danger';
        delBtn.title = 'Delete';
        delBtn.innerHTML = '<i class="bi bi-trash"></i>';
        delBtn.addEventListener('click', () => deleteAttendanceRecord(safeId));
        tdActions.appendChild(editBtn);
        tdActions.appendChild(delBtn);

        tr.append(tdName, tdStatus, tdTimeIn, tdSched, tdDate, tdReason, tdActions);
        tableBody.appendChild(tr);
    });
}

// Filter by date
function filterByDate() {
    const dateInput = document.getElementById('reportDate');
    currentFilterDate = dateInput.value;
    updateSelectedDateDisplay();
    updateDailyReport();
}

// Update selected date display
function updateSelectedDateDisplay() {
    const dateElement = document.getElementById('selectedDate');
    const modalDateElement = document.getElementById('selectedDateModal');
    
    if (currentFilterDate) {
        const date = new Date(currentFilterDate + 'T00:00:00');
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        const formattedDate = date.toLocaleDateString('en-US', options);
        
        if (dateElement) dateElement.textContent = formattedDate;
        if (modalDateElement) modalDateElement.textContent = formattedDate;
    }
}

function selectSchedNote(btn, groupId) {
    document.querySelectorAll('#' + groupId + ' .sn-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
}

function applyFloatSchedule(startId, endId, specialListId) {
    document.getElementById(startId).value = '';
    document.getElementById(endId).value = '';
    document.getElementById(specialListId).innerHTML = '';
    // Uncheck all weekly days
    const prefix = startId === 'scheduleStart' ? 'day' : 'editDay';
    ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].forEach(d => {
        const cb = document.getElementById(prefix + d);
        if (cb) cb.checked = false;
    });
}

const DAYS_OF_WEEK = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

function getDayCheckboxId(containerId, day) {
    const prefix = containerId === 'specialDaysList' ? 'day' : 'editDay';
    return prefix + day;
}

function addSpecialDayRow(containerId, data = {}) {
    const container = document.getElementById(containerId);
    const row = document.createElement('div');
    row.className = 'd-flex align-items-center gap-2 mb-2 special-day-row';
    row.innerHTML = `
        <select class="form-select form-select-sm" style="width:90px" aria-label="Day">
            ${DAYS_OF_WEEK.map(d => `<option value="${d}" ${data.day === d ? 'selected' : ''}>${d}</option>`).join('')}
        </select>
        <input type="time" class="form-control form-control-sm" style="width:110px" placeholder="Start" value="${data.start || ''}" aria-label="Start">
        <input type="time" class="form-control form-control-sm" style="width:110px" placeholder="End" value="${data.end || ''}" aria-label="End">
        <button type="button" class="btn btn-outline-danger btn-sm" title="Remove"><i class="bi bi-x"></i></button>
    `;
    container.appendChild(row);

    const select = row.querySelector('select');
    const removeBtn = row.querySelector('button');
    const startInput = row.querySelectorAll('input[type="time"]')[0];
    const endInput = row.querySelectorAll('input[type="time"]')[1];

    startInput.addEventListener('change', function() {
        if (this.value) {
            const [h, m] = this.value.split(':');
            const endHour = (parseInt(h) + 9) % 24;
            endInput.value = `${endHour.toString().padStart(2, '0')}:${m}`;
        }
    });

    // Auto-check the day in Weekly Days when selected
    const syncCheck = (day) => {
        const cb = document.getElementById(getDayCheckboxId(containerId, day));
        if (cb) cb.checked = true;
    };

    // Check on initial load
    if (data.day) syncCheck(data.day);
    else syncCheck(select.value);

    select.addEventListener('change', () => syncCheck(select.value));

    removeBtn.addEventListener('click', () => row.remove());
}

function getSpecialDays(containerId) {
    const rows = document.querySelectorAll('#' + containerId + ' .special-day-row');
    const result = [];
    rows.forEach(row => {
        const [dayEl, startEl, endEl] = row.querySelectorAll('select, input[type="time"]');
        if (dayEl.value && startEl.value && endEl.value) {
            result.push({ day: dayEl.value, start: startEl.value, end: endEl.value });
        }
    });
    return result;
}

function loadScheduleOverrides() {
    const raw = storage.getItem('scheduleOverrides');
    return raw ? JSON.parse(raw) : [];
}

function saveScheduleOverrides(overrides) {
    storage.setItem('scheduleOverrides', JSON.stringify(overrides || []));
}

function getScheduleOverrideForDate(employeeName, dateStr, department) {
    if (!employeeName || !dateStr || !department) return null;
    const keyName = employeeName.toString().trim().toUpperCase();
    return loadScheduleOverrides().find(o =>
        (o.name || '').toString().trim().toUpperCase() === keyName &&
        o.date === dateStr &&
        o.department === department
    ) || null;
}

function getScheduleForDate(employee, dateStr, includeOverride = true) {
    if (!dateStr || !employee) return { start: employee.scheduleStart, end: employee.scheduleEnd, display: employee.scheduleDisplay || '' };

    const dept = employee.department || currentDepartment;
    if (includeOverride) {
        const o = getScheduleOverrideForDate(employee.name, dateStr, dept);
        if (o && o.type === 'shift') {
            return { start: o.start, end: o.end, display: `${formatTime(o.start)} - ${formatTime(o.end)}`, source: 'override-shift' };
        }
        if (o && o.type === 'broken' && Array.isArray(o.segments) && o.segments.length > 0) {
            const s1 = o.segments[0];
            const s2 = o.segments[1] || o.segments[0];
            return {
                start: s1.start,
                end: s2.end,
                segments: o.segments,
                isBroken: true,
                display: `${formatTime(s1.start)} - ${formatTime(s1.end)} / ${formatTime(s2.start)} - ${formatTime(s2.end)}`,
                source: 'override-broken'
            };
        }
    }

    const dayName = new Date(dateStr + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'short' });
    // dayName is like "Mon", "Tue", etc.
    const special = (employee.specialDays || []).find(s => s.day === dayName);
    if (special) return { start: special.start, end: special.end, display: `${formatTime(special.start)} - ${formatTime(special.end)}`, source: 'special' };
    return { start: employee.scheduleStart, end: employee.scheduleEnd, display: employee.scheduleDisplay || `${formatTime(employee.scheduleStart)} - ${formatTime(employee.scheduleEnd)}`, source: 'regular' };
}

function initScheduleOverrideControls() {
    const modeShift = document.getElementById('overrideModeShift');
    const modeBroken = document.getElementById('overrideModeBroken');
    const shiftFields = document.getElementById('overrideShiftFields');
    const brokenFields = document.getElementById('overrideBrokenFields');
    const dateInput = document.getElementById('attendanceDate');
    const overrideDate = document.getElementById('overrideDate');

    if (!modeShift || !modeBroken || !shiftFields || !brokenFields) return;

    const refreshMode = () => {
        const broken = modeBroken.checked;
        shiftFields.style.display = broken ? 'none' : '';
        brokenFields.style.display = broken ? '' : 'none';
    };
    modeShift.addEventListener('change', refreshMode);
    modeBroken.addEventListener('change', refreshMode);
    refreshMode();

    if (dateInput && overrideDate && !overrideDate.value) {
        overrideDate.value = dateInput.value || getLocalISODate();
    }
}

function openScheduleOverrideModal() {
    const name = (document.getElementById('employeeName')?.value || '').trim();
    const date = document.getElementById('attendanceDate')?.value || getLocalISODate();
    if (!name) {
        showNotification('Select employee name first.', 'warning');
        return;
    }

    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const employee = employees.find(e => e.name === name && e.department === currentDepartment);
    if (!employee) {
        showNotification('Employee not found in current department.', 'warning');
        return;
    }

    const existing = getScheduleOverrideForDate(name, date, currentDepartment);
    const info = document.getElementById('scheduleOverrideEmployee');
    const overrideDate = document.getElementById('overrideDate');
    const shiftMode = document.getElementById('overrideModeShift');
    const brokenMode = document.getElementById('overrideModeBroken');

    if (info) info.textContent = `${name} — ${date}`;
    if (overrideDate) overrideDate.value = date;

    if (existing && existing.type === 'broken') {
        brokenMode.checked = true;
        const seg1 = existing.segments?.[0] || {};
        const seg2 = existing.segments?.[1] || {};
        document.getElementById('overrideB1Start').value = seg1.start || '';
        document.getElementById('overrideB1End').value = seg1.end || '';
        document.getElementById('overrideB2Start').value = seg2.start || '';
        document.getElementById('overrideB2End').value = seg2.end || '';
        document.getElementById('overrideStart').value = '';
        document.getElementById('overrideEnd').value = '';
    } else {
        shiftMode.checked = true;
        const base = existing && existing.type === 'shift'
            ? { start: existing.start, end: existing.end }
            : getScheduleForDate(employee, date, false);
        document.getElementById('overrideStart').value = base.start || '';
        document.getElementById('overrideEnd').value = base.end || '';
        document.getElementById('overrideB1Start').value = '';
        document.getElementById('overrideB1End').value = '';
        document.getElementById('overrideB2Start').value = '';
        document.getElementById('overrideB2End').value = '';
    }

    initScheduleOverrideControls();
    new bootstrap.Modal(document.getElementById('scheduleOverrideModal')).show();
}

function saveScheduleOverride() {
    const name = (document.getElementById('employeeName')?.value || '').trim();
    const date = document.getElementById('overrideDate')?.value;
    const isBroken = document.getElementById('overrideModeBroken')?.checked;
    if (!name || !date) {
        showNotification('Employee and date are required.', 'warning');
        return;
    }

    const overrides = loadScheduleOverrides().filter(o =>
        !((o.name || '').toString().trim().toUpperCase() === name.toUpperCase() && o.date === date && o.department === currentDepartment)
    );

    if (isBroken) {
        const b1s = document.getElementById('overrideB1Start').value;
        const b1e = document.getElementById('overrideB1End').value;
        const b2s = document.getElementById('overrideB2Start').value;
        const b2e = document.getElementById('overrideB2End').value;
        if (!b1s || !b1e || !b2s || !b2e) {
            showNotification('Complete all broken schedule fields.', 'warning');
            return;
        }
        overrides.push({
            id: Date.now(),
            name,
            department: currentDepartment,
            date,
            type: 'broken',
            segments: [{ start: b1s, end: b1e }, { start: b2s, end: b2e }],
            createdAt: new Date().toISOString()
        });
    } else {
        const start = document.getElementById('overrideStart').value;
        const end = document.getElementById('overrideEnd').value;
        if (!start || !end) {
            showNotification('Set shift start and end time.', 'warning');
            return;
        }
        overrides.push({
            id: Date.now(),
            name,
            department: currentDepartment,
            date,
            type: 'shift',
            start,
            end,
            createdAt: new Date().toISOString()
        });
    }

    saveScheduleOverrides(overrides);
    bootstrap.Modal.getInstance(document.getElementById('scheduleOverrideModal'))?.hide();
    showNotification('One-day schedule override saved.', 'success');
}

function clearScheduleOverride() {
    const name = (document.getElementById('employeeName')?.value || '').trim();
    const date = document.getElementById('overrideDate')?.value;
    if (!name || !date) return;
    const next = loadScheduleOverrides().filter(o =>
        !((o.name || '').toString().trim().toUpperCase() === name.toUpperCase() && o.date === date && o.department === currentDepartment)
    );
    saveScheduleOverrides(next);
    bootstrap.Modal.getInstance(document.getElementById('scheduleOverrideModal'))?.hide();
    showNotification('One-day schedule override cleared.', 'info');
}

// Save new employee
function saveEmployee() {
    const name = document.getElementById('newEmployeeName').value;
    const scheduleStart = document.getElementById('scheduleStart').value;
    const scheduleEnd = document.getElementById('scheduleEnd').value;
    const scheduleNotes = document.querySelector('#snBtnGroup .sn-btn.active')?.dataset.value || '';
    
    // Get selected days
    const days = [];
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].forEach(day => {
        const checkbox = document.getElementById('day' + day);
        if (checkbox.checked) {
            days.push(day);
        }
    });
    
    if (!name || !scheduleStart || !scheduleEnd || days.length === 0) {
        const isFloat = (document.querySelector('#snBtnGroup .sn-btn.active')?.dataset.value || '').toUpperCase() === 'FLOAT';
        if (!name || (!isFloat && (!scheduleStart || !scheduleEnd || days.length === 0))) {
            showNotification('Please fill all fields and select at least one day', 'warning');
            return;
        }
    }
    
    // Format time
    const formatTime = (time) => {
        const [hours, minutes] = time.split(':');
        const hour = parseInt(hours);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        const displayHour = hour % 12 || 12;
        return `${displayHour}:${minutes} ${ampm}`;
    };
    
    const scheduleDisplay = `${formatTime(scheduleStart)} - ${formatTime(scheduleEnd)}`;
    const weeklyDays = days.join(', ');
    const specialDays = getSpecialDays('specialDaysList');
    
    // Save to storage (addedYear/addedMonth → Google Sheet only lists this emp from that month forward)
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const addedNow = new Date();
    employees.push({
        id: Date.now(),
        name: name,
        scheduleStart: scheduleStart,
        scheduleEnd: scheduleEnd,
        scheduleDisplay: scheduleDisplay,
        scheduleNotes: scheduleNotes,
        weeklyDays: weeklyDays,
        specialDays: specialDays,
        department: currentDepartment,
        addedYear: addedNow.getFullYear(),
        addedMonth: addedNow.getMonth()
    });
    storage.setItem('employees', JSON.stringify(employees));
    
    // Close modal and reset form
    document.activeElement?.blur();
    const modal = bootstrap.Modal.getInstance(document.getElementById('addEmployeeModal'));
    modal.hide();
    document.getElementById('addEmployeeForm').reset();
    document.getElementById('specialDaysList').innerHTML = '';
    document.querySelectorAll('#snBtnGroup .sn-btn').forEach((b,i) => b.classList.toggle('active', i===0));
    
    // Update dashboard
    updateDashboard();
    showNotification('Employee added successfully!', 'success');
    
    // Sync to Google Sheets
    syncDashboardToSheets();
    syncFullState();
}

// Current department
let currentDepartment = 'rv';

// Switch department
function switchDepartment(dept) {
    currentDepartment = dept;
    sessionStorage.setItem('activeDepartment', dept);
    
    // Update button states
    const rvTab = document.getElementById('rvTab');
    const comsTab = document.getElementById('comsTab');
    const body = document.body;
    
    const dot = document.getElementById('deptDot');
    if (dept === 'rv') {
        rvTab.classList.add('active');
        rvTab.classList.remove('btn-outline-primary');
        rvTab.classList.add('btn-primary');
        comsTab.classList.remove('active');
        comsTab.classList.add('btn-outline-primary');
        comsTab.classList.remove('btn-primary');
        document.getElementById('companyName').textContent = 'Red Victory Consumers Goods Trading';
        body.classList.remove('coms-active');
        body.classList.add('rv-active');
        // department background image
        body.classList.remove('coms-bg');
        body.classList.add('rv-bg');
        dot.classList.remove('text-danger');
        dot.classList.add('text-success');
    } else {
        comsTab.classList.add('active');
        comsTab.classList.remove('btn-outline-primary');
        comsTab.classList.add('btn-primary');
        rvTab.classList.remove('active');
        rvTab.classList.add('btn-outline-primary');
        rvTab.classList.remove('btn-primary');
        document.getElementById('companyName').textContent = 'C. Operations Management Services';
        body.classList.remove('rv-active');
        body.classList.add('coms-active');
        // department background image
        body.classList.remove('rv-bg');
        body.classList.add('coms-bg');
        dot.classList.remove('text-success');
        dot.classList.add('text-danger');
    }
    
    // Clear attendance form on department switch
    document.getElementById('attendanceForm').reset();
    setDefaultDate();
    document.getElementById('employeeName').removeAttribute('data-existing-record-id');
    document.getElementById('employeeName').removeAttribute('data-early-out-record-id');
    const timeInEl = document.getElementById('timeIn');
    const timeOutEl = document.getElementById('timeOut');
    timeInEl.value = '';
    timeInEl.disabled = false;
    timeOutEl.value = '';
    timeOutEl.disabled = true;
    timeOutEl.style.backgroundColor = '#e9ecef';
    timeOutEl.style.cursor = 'not-allowed';
    const statusEl = document.getElementById('attendanceStatus');
    statusEl.disabled = false;
    statusEl.style.backgroundColor = '';
    statusEl.style.cursor = '';
    const returnSection = document.getElementById('returnToWorkSection');
    if (returnSection) returnSection.style.display = 'none';
    document.getElementById('earlyOutBtn').style.display = 'inline-flex';
    document.getElementById('returnWorkBtn').style.display = 'none';

    // Update tables
    updateDailyReport();
    updateDashboard();

    // Animate cards on department switch
    document.querySelectorAll('.card').forEach(c => {
        c.classList.remove('card-switch');
        void c.offsetWidth;
        c.classList.add('card-switch');
    });

    const weeklyModal = document.getElementById('weeklyReportModal');
    if (weeklyModal && weeklyModal.classList.contains('show')) {
        updateWeeklyReportByDate();
    }
}

// Initialize date and time
function updateDateTime() {
    const now = new Date();
    
    // Update time
    const timeElement = document.getElementById('currentTime');
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const seconds = now.getSeconds();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const displayHours = hours % 12 || 12;
    timeElement.textContent = `${displayHours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')} ${ampm}`;
    
    // Update date
    const dateElement = document.getElementById('currentDate');
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    dateElement.textContent = now.toLocaleDateString('en-US', options);
    
    // Update current month for Daily Report
    const monthElement = document.getElementById('currentMonth');
    if (monthElement) {
        const monthName = now.toLocaleDateString('en-US', { month: 'long' });
        monthElement.textContent = monthName;
    }
}

// Timezone helper: populate select and handle display
const TZ_LIST = [
    'Asia/Manila', 'UTC', 'America/New_York', 'Europe/London', 'Asia/Tokyo', 'Australia/Sydney'
];

function initTimezoneControls() {
    const btn = document.getElementById('timezoneBtn');
    const container = document.getElementById('timezoneContainer');
    const select = document.getElementById('timezoneSelect');
    const display = document.getElementById('timezoneDisplay');

    // Populate options
    select.innerHTML = '';
    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = 'Use default (Philippine Time)';
    select.appendChild(emptyOpt);
    TZ_LIST.forEach(tz => {
        const opt = document.createElement('option');
        opt.value = tz;
        opt.textContent = tz;
        select.appendChild(opt);
    });

    btn.addEventListener('click', () => {
        container.style.display = container.style.display === 'none' ? 'block' : 'none';
    });

    let tzInterval = null;
    select.addEventListener('change', () => {
        const tz = select.value;
        if (!tz) {
            display.textContent = 'Using Philippines (Asia/Manila)';
            if (tzInterval) { clearInterval(tzInterval); tzInterval = null; }
            // Persist selection
            try { localStorage.removeItem('selectedTimezone'); } catch(e){}
            // clear any stored interval id on the select
            try { if (select._tzIntervalId) { clearInterval(select._tzIntervalId); delete select._tzIntervalId; } } catch(e){}
            return;
        }

        const update = () => {
            try {
                const now = new Date();
                const options = { timeZone: tz, hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' };
                const parts = new Intl.DateTimeFormat('en-US', options).formatToParts(now);
                const h = parts.find(p => p.type === 'hour').value;
                const m = parts.find(p => p.type === 'minute').value;
                const s = parts.find(p => p.type === 'second').value;
                // Also show date
                const dateOptions = { timeZone: tz, year: 'numeric', month: 'short', day: '2-digit' };
                const dateStr = new Intl.DateTimeFormat('en-US', dateOptions).format(now);
                display.textContent = `${dateStr} ${h}:${m}:${s} (${tz})`;
            } catch (e) {
                display.textContent = `Time unavailable for ${tz}`;
            }
        };

        update();
        if (tzInterval) clearInterval(tzInterval);
        tzInterval = setInterval(update, 1000);
        // expose interval id so other code (submit handler) can clear it
        try { select._tzIntervalId = tzInterval; } catch(e){}
        try { window._tzIntervalId = tzInterval; } catch(e){}

        // Also set attendance date and time inputs to the selected timezone's current date/time
        try {
            const now = new Date();
            const dateOptions = { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' };
            const dparts = new Intl.DateTimeFormat('en-CA', dateOptions).format(now).split('-');
            // en-CA gives YYYY-MM-DD
            const dateVal = dparts.join('-');
            const dateInput = document.getElementById('attendanceDate');
            if (dateInput) dateInput.value = dateVal;
            // Persist selection
            try { localStorage.setItem('selectedTimezone', tz); } catch(e){}
        } catch (e) {
            // ignore
        }
    });

    // Initialize display and restore selection
    const saved = (function(){ try { return localStorage.getItem('selectedTimezone'); } catch(e) { return null; } })();
    if (saved) {
        select.value = saved;
        select.dispatchEvent(new Event('change'));
    } else {
        display.textContent = 'Using Philippines (Asia/Manila)';
    }

    // When user clicks/focuses Time In or Time Out, fill with current time for selected timezone (if empty)
    const timeInInput = document.getElementById('timeIn');
    const timeOutInput = document.getElementById('timeOut');
    const fillWithTZTime = (inputEl) => {
        if (!inputEl) return;
        if (inputEl.value) return; // don't override existing value
        const selectedTZ = select.value || Intl.DateTimeFormat().resolvedOptions().timeZone;
        try {
            const now = new Date();
            const timeOptions = { timeZone: selectedTZ, hour12: false, hour: '2-digit', minute: '2-digit' };
            const timeStr = new Intl.DateTimeFormat('en-GB', timeOptions).format(now); // HH:MM
            inputEl.value = timeStr;
        } catch (e) { /* ignore */ }
    };

    if (timeInInput) {
        timeInInput.addEventListener('focus', () => fillWithTZTime(timeInInput));
        timeInInput.addEventListener('click', () => fillWithTZTime(timeInInput));
    }
    if (timeOutInput) {
        timeOutInput.addEventListener('focus', () => fillWithTZTime(timeOutInput));
        timeOutInput.addEventListener('click', () => fillWithTZTime(timeOutInput));
    }
}

// Convert a local time (HH:MM or HH:MM:SS) and date (YYYY-MM-DD) to time in target timezone
function convertToTZ(dateStr, timeStr, targetTZ) {
    if (!targetTZ) return { date: dateStr, time: timeStr };
    // Build a Date from the parts in UTC by interpreting the given date+time as if in targetTZ
    // Approach: get the equivalent instant for the targetTZ by using Date.toLocaleString with timeZone
    try {
        const [y, mo, d] = dateStr.split('-').map(Number);
        const timeParts = (timeStr || '00:00:00').split(':').map(Number);
        const hour = timeParts[0] || 0;
        const minute = timeParts[1] || 0;
        const second = timeParts[2] || 0;

        // Create an ISO-like string in targetTZ by building UTC components from formatted parts
        // Get the milliseconds value for the same local wall-clock in targetTZ by searching for an instant
        // We'll create a Date in UTC from the targetTZ wall time by using Date.UTC and then adjusting by the timezone offset
        // Simpler method: use Intl to get the offset between targetTZ and UTC at the desired instant
        const asUTCString = new Date(Date.UTC(y, mo - 1, d, hour, minute, second));
        // Find the equivalent wall time in targetTZ for that UTC instant
        const parts = new Intl.DateTimeFormat('en-US', { timeZone: targetTZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(asUTCString);
        const ty = parts.find(p => p.type === 'year').value;
        const tmo = parts.find(p => p.type === 'month').value;
        const td = parts.find(p => p.type === 'day').value;
        const th = parts.find(p => p.type === 'hour').value;
        const tm = parts.find(p => p.type === 'minute').value;
        const ts = parts.find(p => p.type === 'second').value;

        return { date: `${ty}-${tmo}-${td}`, time: `${th}:${tm}:${ts}` };
    } catch (e) {
        return { date: dateStr, time: timeStr };
    }
}

// Set default date to today
function setDefaultDate() {
    const dateInput = document.getElementById('attendanceDate');
    if (!dateInput) return;
    const today = getLocalISODate();
    dateInput.value = today;
}

// Run initializers on DOM ready
document.addEventListener('DOMContentLoaded', () => {
    setDefaultDate();
    updateDateTime();
    initTimezoneControls();
    setInterval(updateDateTime, 1000);
});

// Get status badge class
function getStatusBadgeClass(status) {
    const statusMap = {
        'Present': 'status-present',
        'Late': 'status-late',
        'Absent': 'status-absent',
        'Undertime': 'status-undertime',
        'Overtime': 'status-overtime',
        'AWOL': 'status-awol',
        'Sick Leave': 'status-sick',
        'Work From Home': 'status-wfh',
        'No Schedule': 'status-ns'
    };
    return statusMap[status] || 'status-present';
}

// Display-friendly status text (UI only). Map various leave types to a
// unified "Leave" label while preserving the underlying `record.status` value.
function getDisplayStatus(status) {
    if (!status) return '';
    const s = status.toString();
    if (/sick|vacation|vac|leave/i.test(s)) return 'Leave';
    return status;
}

// Return true if the status should be counted as a worked day in totals
function isCountedStatus(status) {
    if (!status) return false;
    const s = status.toString().toLowerCase();
    // Count only actual work statuses — Present, Late, Undertime, Overtime
    return /present|late|undertime|overtime/.test(s);
}

// Format time to AM/PM
function formatTime(time) {
    if (!time) return '-';
    const parts = time.split(':');
    const hours = parts[0];
    const minutes = parts[1] || '00';
    const seconds = parts[2];
    const hour = parseInt(hours);
    const ampm = hour >= 12 ? 'PM' : 'AM';
    const displayHour = hour % 12 || 12;
    if (typeof seconds !== 'undefined') {
        return `${displayHour.toString().padStart(2,'0')}:${minutes}:${seconds} ${ampm}`;
    }
    return `${displayHour.toString().padStart(2,'0')}:${minutes} ${ampm}`;
}

// Load attendance data from storage
function loadAttendanceData() {
    const data = storage.getItem('attendanceData');
    return data ? JSON.parse(data) : [];
}

// Get data by department
function getDataByDepartment(data) {
    return data.filter(record => record.department === currentDepartment);
}

// Save attendance data to storage
function saveAttendanceData(data) {
    // Deduplicate: per employee+date+department, keep only the record with the highest id
    const seen = {};
    const deduped = [];
    // Sort by id ascending so the last one (highest id) wins
    const sorted = [...data].sort((a, b) => a.id - b.id);
    sorted.forEach(r => {
        const key = `${r.name}|${r.date}|${r.department}`;
        seen[key] = r;
    });
    data = Object.values(seen);
    storage.setItem('attendanceData', JSON.stringify(data));
}

// RV late minutes use an 11-minute grace period:
// e.g. 9:13 vs 9:00 schedule => 2 late minutes.
function computeLateMinutes(diffMinutes, department) {
    const diff = Number(diffMinutes) || 0;
    if (department === 'rv') {
        return Math.max(0, diff - 11);
    }
    return Math.max(0, diff);
}

function toMinutesFromTime(t) {
    if (!t) return 0;
    const [h, m] = t.split(':').map(Number);
    return (h || 0) * 60 + (m || 0);
}

function overlapMinutes(aStart, aEnd, bStart, bEnd) {
    return Math.max(0, Math.min(aEnd, bEnd) - Math.max(aStart, bStart));
}

// For broken schedules: count only overlaps with part1 + part2 windows.
// For normal schedules: regular timeIn-timeOut duration.
function calculateWorkedHoursBySchedule(timeIn, timeOut, scheduleMeta) {
    if (!timeIn || !timeOut) return 0;
    let inM = toMinutesFromTime(timeIn);
    let outM = toMinutesFromTime(timeOut);
    if (outM < inM) outM += 24 * 60;

    if (scheduleMeta && scheduleMeta.isBroken && Array.isArray(scheduleMeta.segments) && scheduleMeta.segments.length > 0) {
        let worked = 0;
        scheduleMeta.segments.forEach(seg => {
            let s = toMinutesFromTime(seg.start);
            let e = toMinutesFromTime(seg.end);
            if (e < s) e += 24 * 60;
            worked += overlapMinutes(inM, outM, s, e);
        });
        return parseFloat((worked / 60).toFixed(2));
    }

    return parseFloat(((outM - inM) / 60).toFixed(2));
}

// Update daily report table
function updateDailyReport() {
    const tableBody = document.getElementById('dailyReportTable');
    const modalTableBody = document.getElementById('dailyReportTableModal');
    const attendanceData = loadAttendanceData();
    
    // Filter by selected date
    const filteredRecords = attendanceData.filter(record => 
        record.date === currentFilterDate && record.department === currentDepartment
    );
    filteredRecords.sort((a, b) => a.name.localeCompare(b.name));
    
    const tableHTML = filteredRecords.length === 0 ? `
        <tr>
            <td colspan="6" class="text-center text-muted py-4">
                <i class="bi bi-inbox fs-1 d-block mb-2"></i>
                No attendance records for selected date
            </td>
        </tr>
    ` : filteredRecords.map(record => {
        const isLate = record.status === 'Late';
        const safeId = parseInt(record.id, 10);
        return `
        <tr>
            <td>${record.name}</td>
            <td><span class="badge-status ${getStatusBadgeClass(record.status)}">${getDisplayStatus(record.status)}</span></td>
            <td class="${isLate ? 'text-warning fw-bold' : ''}">${formatTime(record.timeIn)}</td>
            <td>${record.scheduleTime || '-'}</td>
            <td>${record.date}</td>
            <td>${record.reason ? `<span class="badge bg-warning text-dark">${record.reason}</span>` : '-'}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-outline-primary" onclick="editAttendanceRecord(${safeId})" title="Edit"><i class="bi bi-pencil"></i></button>
                <button class="btn btn-sm btn-outline-danger" onclick="deleteAttendanceRecord(${safeId})" title="Delete"><i class="bi bi-trash"></i></button>
            </td>
        </tr>
    `}).join('');
    
    if (tableBody) tableBody.innerHTML = tableHTML;
    if (modalTableBody) modalTableBody.innerHTML = tableHTML;
}

// Update weekly report with month/year filter
function updateWeeklyReportWithFilter(selectedMonth = null, selectedYear = null, weekNumber = null) {
    const modalTableBody = document.getElementById('weeklyReportTableModal');
    if (!modalTableBody) return;
    
    const attendanceData = loadAttendanceData();
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    
    // Use current month/year if not specified
    if (selectedMonth === null || selectedYear === null) {
        const now = new Date();
        selectedMonth = selectedMonth !== null ? selectedMonth : now.getMonth();
        selectedYear = selectedYear !== null ? selectedYear : now.getFullYear();
    }
    
    // Filter by month, year, and optionally week
    let filteredData = attendanceData.filter(record => {
        if (record.department !== currentDepartment) return false;
        
        const recordDate = new Date(record.date + 'T00:00:00');
        const recordMonth = recordDate.getMonth();
        const recordYear = recordDate.getFullYear();
        
        // Check month and year
        if (recordMonth !== selectedMonth || recordYear !== selectedYear) return false;
        
        // Check week if specified
        if (weekNumber) {
            const startDay = (weekNumber - 1) * 7 + 1;
            const endDay = Math.min(weekNumber * 7, new Date(selectedYear, selectedMonth + 1, 0).getDate());
            const recordDay = recordDate.getDate();
            return recordDay >= startDay && recordDay <= endDay;
        }
        
        return true;
    });
    
    // Group by employee name
    const employeeRecords = {};
    
    filteredData.forEach(record => {
        if (!employeeRecords[record.name]) {
            const employee = employees.find(e => e.name === record.name && e.department === currentDepartment);
            employeeRecords[record.name] = {
                name: record.name,
                dates: [],
                scheduleDisplay: employee ? employee.scheduleDisplay : 'Not Set',
                totalHours: 0,
                absentCount: 0,
                leaveCount: 0,
                records: []
            };
        }
        // Add date if not already in the list (exclude WFH from day count)
        if (!employeeRecords[record.name].dates.includes(record.date) && isCountedStatus(record.status)) {
            employeeRecords[record.name].dates.push(record.date);
        }
        // Count absences and leaves separately
        if (/absent/i.test(record.status || '')) employeeRecords[record.name].absentCount++;
        if (/sick|vacation|vac|leave/i.test(record.status || '')) employeeRecords[record.name].leaveCount++;
        employeeRecords[record.name].totalHours += parseFloat(record.totalHours || 0);
        employeeRecords[record.name].records.push(record);
    });
    
    if (Object.keys(employeeRecords).length === 0) {
        modalTableBody.innerHTML = `
            <tr>
                <td colspan="5" class="text-center text-muted py-4">
                    <i class="bi bi-inbox fs-1 d-block mb-2"></i>
                    No attendance data available for selected period
                </td>
            </tr>
        `;
        return;
    }
    
    modalTableBody.innerHTML = Object.values(employeeRecords).sort((a, b) => a.name.localeCompare(b.name)).map(data => {
        return `
        <tr>
            <td><strong>${data.name}</strong></td>
            <td><span class="badge bg-info text-dark">${data.scheduleDisplay}</span></td>
            <td class="text-center fw-bold text-primary">${data.totalHours.toFixed(2)} hrs</td>
            <td class="text-center fw-bold text-danger">${data.absentCount || 0}</td>
            <td class="text-center fw-bold text-warning">${data.leaveCount || 0}</td>
            <td class="text-center fw-bold text-success">${data.dates.length} day${data.dates.length !== 1 ? 's' : ''}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-outline-info" onclick="viewEmployeeMonthDetails('${data.name}', ${selectedMonth}, ${selectedYear})" title="View Details"><i class="bi bi-eye"></i></button>
            </td>
        </tr>
    `}).join('');
}

// Update employee dashboard
function updateDashboard() {
    const tableBody = document.getElementById('dashboardTable');
    const tableHead = document.querySelector('.dashboard-table thead tr');
    const attendanceData = loadAttendanceData();
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    
    // Update employee name datalist
    updateEmployeeDatalist();
    
    // Update table headers based on department
    if (currentDepartment === 'rv') {
        tableHead.innerHTML = `
            <th>Name</th>
            <th class="text-center">Present</th>
            <th class="text-center">Absent</th>
            <th class="text-center">Late</th>
            <th class="text-center">Total Lates (mins)</th>
            <th class="text-center">Undertime</th>
            <th class="text-center">Overtime</th>
            <th class="text-center">AWOL</th>
            <th class="text-center">Leave</th>
            <th class="text-center">WFH</th>
            <th class="text-center">Sched Time</th>
            <th class="text-center">Notes</th>
            <th class="text-center">Actions</th>
        `;
    } else {
        tableHead.innerHTML = `
            <th>Name</th>
            <th class="text-center">Present</th>
            <th class="text-center">Absent</th>
            <th class="text-center">Late</th>
            <th class="text-center">Undertime</th>
            <th class="text-center">AWOL</th>
            <th class="text-center">Leave</th>
            <th class="text-center">WFH</th>
            <th class="text-center">Sched Time</th>
            <th class="text-center">Notes</th>
            <th class="text-center">Actions</th>
        `;
    }
    
    // Group by employee - ONLY for employees that exist in the employees list
    const employeeStats = {};
    
    // Sort employees alphabetically
    employees.sort((a, b) => a.name.localeCompare(b.name));

    // Initialize all saved employees with zero stats
    employees.forEach(emp => {
        if (emp.department === currentDepartment) {
            employeeStats[emp.name] = {
                present: 0,
                absent: 0,
                late: 0,
                totalLates: 0,
                undertime: 0,
                overtime: 0,
                awol: 0,
                sickLeave: 0,
                wfh: 0,
                noSchedule: 0,
                scheduleDisplay: emp.scheduleDisplay,
                weeklyDays: emp.weeklyDays
            };
        }
    });
    
    // Update stats from attendance records - ONLY for employees in the list AND current month only
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    attendanceData.forEach(record => {
        if (record.department !== currentDepartment) return;
        
        // Filter by current month and year
        const recordDate = new Date(record.date + 'T00:00:00');
        const recordMonth = recordDate.getMonth();
        const recordYear = recordDate.getFullYear();
        
        // Only process current month data
        if (recordMonth !== currentMonth || recordYear !== currentYear) return;
        
        // Only process if employee exists in employeeStats (i.e., in employees list)
        if (employeeStats[record.name]) {
            const stats = employeeStats[record.name];
            
            switch(record.status) {
                case 'Present':
                    stats.present++;
                    break;
                case 'Absent':
                    stats.absent++;
                    break;
                case 'Late':
                    stats.late++;
                    stats.totalLates += Number(record.lateMinutes) || 0;
                    break;
                case 'Undertime':
                    stats.undertime++;
                    break;
                case 'Overtime':
                    if (currentDepartment === 'rv') stats.overtime++; else stats.present++;
                    break;
                case 'AWOL':
                    stats.awol++;
                    break;
                case 'Sick Leave':
                    stats.sickLeave++;
                    break;
                case 'Work From Home':
                    stats.wfh++;
                    break;
                case 'No Schedule':
                    stats.noSchedule++;
                    break;
            }
        }
    });
    
    if (Object.keys(employeeStats).length === 0) {
        const colspan = currentDepartment === 'rv' ? '13' : '11';
        tableBody.innerHTML = `
            <tr>
                <td colspan="${colspan}" class="text-center text-muted py-4">
                    <i class="bi bi-people fs-1 d-block mb-2"></i>
                    No employee data available
                </td>
            </tr>
        `;
        return;
    }
    
    tableBody.innerHTML = Object.entries(employeeStats).map(([name, stats]) => {
        const employee = employees.find(e => e.name === name && e.department === currentDepartment);
        const empId = employee ? employee.id : name;
        
        // Calculate habitual count (Absent + Late + Undertime)
        const habitualCount = stats.absent + stats.late + stats.undertime;
        
        // Get policy colors
        const awolColorClass = getAWOLPolicyColor(stats.awol);
        const absentColorClass = getHabitualPolicyColor(habitualCount);
        const lateColorClass = getHabitualPolicyColor(habitualCount);
        const undertimeColorClass = getHabitualPolicyColor(habitualCount);
        const empNotes = employee && employee.scheduleNotes ? employee.scheduleNotes.toUpperCase() : '';
        const isFloat = empNotes.includes('FLOAT');
        const schedStart = employee && employee.scheduleStart ? (() => { const [h,m] = employee.scheduleStart.split(':'); const hr = parseInt(h); return `${hr%12||12}:${m} ${hr>=12?'PM':'AM'}`; })() : (stats.scheduleDisplay || 'Not Set');
        const schedDisplay = isFloat ? '-' : `<span class="badge bg-info text-dark">${schedStart}</span>`;
        
        if (currentDepartment === 'rv') {
            return `
            <tr>
                <td><strong>${name}</strong></td>
                <td class="text-center text-present">${stats.present}</td>
                <td class="text-center text-absent ${absentColorClass}">${stats.absent}</td>
                <td class="text-center text-late ${lateColorClass}">${stats.late}</td>
                <td class="text-center">${stats.totalLates}</td>
                <td class="text-center text-undertime ${undertimeColorClass}">${stats.undertime}</td>
                <td class="text-center text-overtime">${stats.overtime}</td>
                <td class="text-center text-awol ${awolColorClass}">${stats.awol}</td>
                <td class="text-center text-sick">${stats.sickLeave}</td>
                <td class="text-center text-wfh">${stats.wfh}</td>
                <td class="text-center">${schedDisplay}</td>
                <td class="text-center">${employee && employee.scheduleNotes ? `<span class="badge bg-secondary">${employee.scheduleNotes}</span>` : '-'}</td>
                <td class="text-center">
                    <button class="btn btn-sm btn-outline-primary" onclick="openEditModal('${empId}')" title="Edit"><i class="bi bi-pencil"></i></button>
                    <button class="btn btn-sm btn-outline-info" onclick="openScheduleModal('${empId}')" title="View Schedule"><i class="bi bi-calendar3"></i></button>
                    <button class="btn btn-sm btn-outline-danger" onclick="confirmDelete('${empId}')" title="Delete"><i class="bi bi-trash"></i></button>
                </td>
            </tr>
        `;
        } else {
            return `
            <tr>
                <td><strong>${name}</strong></td>
                <td class="text-center text-present">${stats.present}</td>
                <td class="text-center text-absent ${absentColorClass}">${stats.absent}</td>
                <td class="text-center text-late ${lateColorClass}">${stats.late}</td>
                <td class="text-center text-undertime ${undertimeColorClass}">${stats.undertime}</td>
                <td class="text-center text-awol ${awolColorClass}">${stats.awol}</td>
                <td class="text-center text-sick">${stats.sickLeave}</td>
                <td class="text-center text-wfh">${stats.wfh}</td>
                <td class="text-center">${schedDisplay}</td>
                <td class="text-center">${employee && employee.scheduleNotes ? `<span class="badge bg-secondary">${employee.scheduleNotes}</span>` : '-'}</td>
                <td class="text-center">
                    <button class="btn btn-sm btn-outline-primary" onclick="openEditModal('${empId}')" title="Edit"><i class="bi bi-pencil"></i></button>
                    <button class="btn btn-sm btn-outline-info" onclick="openScheduleModal('${empId}')" title="View Schedule"><i class="bi bi-calendar3"></i></button>
                    <button class="btn btn-sm btn-outline-danger" onclick="confirmDelete('${empId}')" title="Delete"><i class="bi bi-trash"></i></button>
                </td>
            </tr>
        `;
        }
    }).join('');
}

// Handle form submission
document.getElementById('attendanceForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    // Capture current TZ at submit and immediately clear/hide the TZ UI so
    // the panel closes and live updating stops even before submit completes.
    const tzSelectEl = document.getElementById('timezoneSelect');
    const tzDisplayEl = document.getElementById('timezoneDisplay');
    const tzContainerEl = document.getElementById('timezoneContainer');
    const tzAtSubmit = tzSelectEl ? tzSelectEl.value : '';
    try { localStorage.removeItem('selectedTimezone'); } catch(e){}
    try { if (tzSelectEl && tzSelectEl._tzIntervalId) { clearInterval(tzSelectEl._tzIntervalId); delete tzSelectEl._tzIntervalId; } } catch(e){}
    try { if (window._tzIntervalId) { clearInterval(window._tzIntervalId); delete window._tzIntervalId; } } catch(e){}
    if (tzContainerEl) tzContainerEl.style.display = 'none';
    if (tzDisplayEl) tzDisplayEl.textContent = '';

    const name = document.getElementById('employeeName').value;
    const date = document.getElementById('attendanceDate').value;
    let status = document.getElementById('attendanceStatus').value;
    let timeIn = document.getElementById('timeIn').value;
    let timeOut = document.getElementById('timeOut').value;
    const reason = document.getElementById('attendanceReason').value;
    const existingRecordId = document.getElementById('employeeName').getAttribute('data-existing-record-id');
    let lateMinutes = 0;
    let scheduleTime = '';
    let totalHours = 0;
    
    // Initial total hours (may be adjusted below for broken schedules)
    if (timeIn && timeOut) {
        const [inHours, inMinutes] = timeIn.split(':').map(Number);
        const [outHours, outMinutes] = timeOut.split(':').map(Number);
        
        const inTotalMinutes = inHours * 60 + inMinutes;
        let outTotalMinutes = outHours * 60 + outMinutes;
        
        // Handle overnight shift (e.g. time in 22:00, time out 06:00)
        if (outTotalMinutes < inTotalMinutes) {
            outTotalMinutes += 24 * 60;
        }
        
        const diffMinutes = outTotalMinutes - inTotalMinutes;
        totalHours = parseFloat((diffMinutes / 60).toFixed(2));
    }
    
    // Get employee data
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const employee = employees.find(e => e.name === name && e.department === currentDepartment);

    // Use the TZ captured at submit time (so we don't lose it when we cleared the UI)
    const tz = typeof tzAtSubmit !== 'undefined' ? tzAtSubmit : (document.getElementById('timezoneSelect') ? document.getElementById('timezoneSelect').value : '');
    const hadTZ = !!tz; // remember whether user had selected a timezone at submit time
    
    if (employee) {
        const daySched = getScheduleForDate(employee, date);
        scheduleTime = daySched?.display || employee.scheduleDisplay || '';
    }

    let attendanceData = loadAttendanceData();

    // --- Overnight early-arrival correction ---
    // If schedule starts at midnight or early morning (00:00–05:59) and the employee
    // logs in late at night (20:00–23:59) on the form date, they are actually arriving
    // early for the NEXT day's shift. Shift the record date forward by 1 day.
    let recordDate = date;
    if (employee && status === 'Present' && timeIn && !existingRecordId) {
        const checkSched = getScheduleForDate(employee, date);
        const schedStartForCheck = checkSched?.start || employee.scheduleStart;
        if (schedStartForCheck) {
            const [schedH] = schedStartForCheck.split(':').map(Number);
            const [timeH] = timeIn.split(':').map(Number);
            // Schedule starts between midnight and 6 AM, and time-in is 8 PM or later
            if (schedH >= 0 && schedH < 6 && timeH >= 20) {
                const nextDay = new Date(date + 'T00:00:00');
                nextDay.setDate(nextDay.getDate() + 1);
                const y = nextDay.getFullYear();
                const mo = String(nextDay.getMonth() + 1).padStart(2, '0');
                const dy = String(nextDay.getDate()).padStart(2, '0');
                recordDate = `${y}-${mo}-${dy}`;
                // Also update the form date input so subsequent logic uses the correct date
                document.getElementById('attendanceDate').value = recordDate;
            }
        }
    }

    // Auto-detect late if status is Present
    // Skip late detection if employee already has a record today (returning from break)
    const existingTodayRecord = attendanceData.find(r => 
        r.name === name && r.date === recordDate && r.department === currentDepartment
    );

    const notes = employee && employee.scheduleNotes ? employee.scheduleNotes.toUpperCase() : '';
    const isFlex = notes.includes('FLEX') || notes.includes('FLOAT');

    // Block FLOAT employees from submitting attendance
    if (notes.includes('FLOAT') && !existingRecordId) {
        showNotification(`${name} is on FLOAT schedule and is not required to log attendance.`, 'warning');
        return;
    }

    // FLEX policy: never count/select "Late" status for flex shifts.
    // Keep actual clock-in time but normalize status to Present.
    if (notes.includes('FLEX') && status === 'Late') {
        status = 'Present';
        lateMinutes = 0;
    }

    if (status === 'Present' && timeIn && employee && !existingTodayRecord && !isFlex) {
        const daySched = getScheduleForDate(employee, recordDate);
        const schedStart = daySched?.start || employee.scheduleStart;
        if (schedStart) {
            const [schedHours, schedMins] = schedStart.split(':').map(Number);
            const [timeHours, timeMins] = timeIn.split(':').map(Number);
            
            let schedMinutes = schedHours * 60 + schedMins;
            let timeMinutes = timeHours * 60 + timeMins;
            // For midnight-start schedules (00:00-05:59), a late-night time-in (20:00+)
            // means the employee arrived early for the next day — treat as negative late (early)
            if (schedHours < 6 && timeHours >= 20) {
                // time-in is on the previous calendar day relative to schedule
                // e.g. 23:51 vs 00:00 → employee is 9 mins early, not 1431 mins late
                timeMinutes = timeMinutes - 24 * 60; // shift back to make it negative
            }
            const delayMinutes = timeMinutes - schedMinutes;
            lateMinutes = computeLateMinutes(delayMinutes, currentDepartment);
            
            if (delayMinutes >= 11) {
                status = 'Late';
            } else {
                lateMinutes = 0;
            }
        }
    }
    
    // Auto-detect Overtime: if total hours >= 10 and status is Present or Late
    if (timeOut && totalHours >= 10 && (status === 'Present' || status === 'Late')) {
        status = 'Overtime';
    }

    // Calculate late minutes for manually selected Late status
    if (status === 'Late' && timeIn && lateMinutes === 0 && employee && !isFlex) {
        const daySched = getScheduleForDate(employee, recordDate);
        const schedStart = daySched?.start || employee.scheduleStart;
        if (schedStart) {
            const [schedHours, schedMins] = schedStart.split(':').map(Number);
            const [timeHours, timeMins] = timeIn.split(':').map(Number);
            
            const schedMinutes = schedHours * 60 + schedMins;
            const timeMinutes = timeHours * 60 + timeMins;
            lateMinutes = computeLateMinutes(timeMinutes - schedMinutes, currentDepartment);
        }
    }

    // Broken schedule hour logic:
    // automatically excludes break gap between part 1 and part 2.
    if (timeIn && timeOut && employee) {
        const scheduleMeta = getScheduleForDate(employee, recordDate);
        totalHours = calculateWorkedHoursBySchedule(timeIn, timeOut, scheduleMeta);
    }
    
    // Also check for any existing record on the same date (catches early-out re-submit)
    if (!existingRecordId) {
        const sameDay = attendanceData.find(r =>
            r.name === name && r.date === recordDate && r.department === currentDepartment
        );
        if (sameDay && status === 'Undertime') {
            // Update the existing record instead of creating a duplicate
            const idx = attendanceData.indexOf(sameDay);
            attendanceData[idx].timeOut    = timeOut;
            attendanceData[idx].totalHours = totalHours;
            attendanceData[idx].status     = 'Undertime';
            if (reason) attendanceData[idx].reason = reason;
            saveAttendanceData(attendanceData);
            this.reset();
            setDefaultDate();
            document.getElementById('employeeName').removeAttribute('data-existing-record-id');
            updateDailyReport();
            updateWeeklyReportWithFilter();
            updateDashboard();
            updateEmployeeDatalist();
            await syncFullState();
            await syncDashboardToSheets();
            await syncWeeklyReportToSheets();
            showNotification('Attendance recorded successfully!', 'success');
            return;
        }
    }

    if (existingRecordId) {
        const recordIndex = attendanceData.findIndex(r => r.id == existingRecordId);
        if (recordIndex !== -1) {
            const existingRecord = attendanceData[recordIndex];

            // If employee returned from early out, calculate session 1 + session 2 hours
            if (existingRecord.returnTimeIn && timeOut) {
                const [s2InH, s2InM]   = existingRecord.returnTimeIn.split(':').map(Number);
                const [s2OutH, s2OutM] = timeOut.split(':').map(Number);
                let s2In  = s2InH * 60 + s2InM;
                let s2Out = s2OutH * 60 + s2OutM;
                if (s2Out < s2In) s2Out += 24 * 60;
                const session2Hours = parseFloat(((s2Out - s2In) / 60).toFixed(2));
                const session1Hours = parseFloat(existingRecord.session1Hours || 0);
                totalHours = parseFloat((session1Hours + session2Hours).toFixed(2));

                // Restore original status (not Undertime anymore).
                // Prefer explicit previousStatus if available; otherwise keep
                // the existing record.status. Do NOT default to 'Present' as
                // this can unintentionally clear 'Late' or other statuses.
                if (existingRecord.previousStatus) {
                    status = existingRecord.previousStatus;
                } else if (existingRecord.status) {
                    status = existingRecord.status;
                }

                attendanceData[recordIndex] = {
                    ...existingRecord,
                    timeOut:      timeOut,
                    totalHours:   totalHours,
                    status:       status
                };
            } else {
                attendanceData[recordIndex].timeOut    = timeOut;
                attendanceData[recordIndex].totalHours = totalHours;
                if (reason) attendanceData[recordIndex].reason = reason;
                // If early out was used, override status to Undertime
                if (status === 'Undertime') attendanceData[recordIndex].status = 'Undertime';
                // If total hours >= 10, override status to Overtime
                else if (totalHours >= 10 && (attendanceData[recordIndex].status === 'Present' || attendanceData[recordIndex].status === 'Late')) {
                    attendanceData[recordIndex].status = 'Overtime';
                }
            }
        }
        showNotification('Attendance recorded successfully!', 'success');
    } else {
        const newRecord = {
            id: Date.now(),
            name: name,
            date: recordDate,
            status: status,
            timeIn: timeIn,
            timeOut: timeOut,
            totalHours: totalHours,
            scheduleTime: scheduleTime,
            lateMinutes: lateMinutes,
            reason: reason,
            tz: tz || '',
            department: currentDepartment
        };
        attendanceData.push(newRecord);
        showNotification('Attendance recorded successfully!', 'success');
    }
    
    saveAttendanceData(attendanceData);
    
    // Reset form
    this.reset();
    setDefaultDate();
    document.getElementById('employeeName').removeAttribute('data-existing-record-id');
    document.getElementById('employeeName').removeAttribute('data-early-out-record-id');
    document.getElementById('earlyOutBtn').style.display = 'inline-flex';
    document.getElementById('returnWorkBtn').style.display = 'none';
    // Re-lock timeout, unlock status
    const toField = document.getElementById('timeOut');
    toField.disabled = true;
    toField.style.backgroundColor = '#e9ecef';
    toField.style.cursor = 'not-allowed';
    const statusField = document.getElementById('attendanceStatus');
    statusField.disabled = false;
    statusField.style.backgroundColor = '';
    statusField.style.cursor = '';
    const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');
    submitBtn.disabled = false;
    submitBtn.style.opacity = '1';
    
    // Update tables
    updateDailyReport();
    updateWeeklyReportWithFilter();
    updateDashboard();
    
    // Update weekly report modal if open: day view if DATE set, else month summary (selected month)
    const weeklyModal = document.getElementById('weeklyReportModal');
    if (weeklyModal && weeklyModal.classList.contains('show')) {
        const reportDatePicker = document.getElementById('reportDatePicker');
        if (reportDatePicker && reportDatePicker.value) {
            updateWeeklyReportByDate();
        } else {
            const mp = document.getElementById('reportMonthSummaryPicker');
            if (mp && !mp.value) mp.value = getLocalMonthInputValue();
            updateWeeklyReportMonthSummary();
        }
    }
    
    // Update employee datalist for autocomplete
    updateEmployeeDatalist();
    
    // Sync to Google Sheets — ensure full state is sent first so server can seed blocks
    await syncFullState();
    await syncDashboardToSheets();
    await syncWeeklyReportToSheets();

    // After successful submit, clear any stored timezone and update the UI.
    try { localStorage.removeItem('selectedTimezone'); } catch (e) { /* ignore */ }
    const tzSelect = document.getElementById('timezoneSelect');
    const tzDisplay = document.getElementById('timezoneDisplay');
    const tzContainer = document.getElementById('timezoneContainer');
    if (tzSelect) {
        // Clear the select value for next use
        tzSelect.value = '';
        // Clear any running tz interval to stop live updates (both select-bound and global)
        try { if (tzSelect._tzIntervalId) { clearInterval(tzSelect._tzIntervalId); delete tzSelect._tzIntervalId; } } catch(e) {}
        try { if (window._tzIntervalId) { clearInterval(window._tzIntervalId); delete window._tzIntervalId; } } catch(e) {}

        // Always hide the timezone panel after submit and remove the live text
        if (tzContainer) tzContainer.style.display = 'none';
        if (tzDisplay) tzDisplay.textContent = '';
    } else if (tzDisplay) {
        tzDisplay.textContent = 'Using Philippines (Asia/Manila)';
    }
});

// Show notification
function showNotification(message, type) {
    const notification = document.createElement('div');
    notification.className = `alert alert-${type} position-fixed top-0 start-50 translate-middle-x mt-3`;
    notification.style.zIndex = '9999';
    notification.style.minWidth = '300px';
    notification.innerHTML = `
        <i class="bi bi-check-circle-fill me-2"></i>${message}
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transition = 'opacity 0.5s ease';
        setTimeout(() => notification.remove(), 500);
    }, 3000);
}

// Export data to CSV
function exportData() {
    const attendanceData = loadAttendanceData();
    
    if (attendanceData.length === 0) {
        showNotification('No data to export', 'warning');
        return;
    }
    
    let csv = 'Name,Date,Status,Time In\n';
    
    attendanceData.forEach(record => {
        csv += `${record.name},${record.date},${record.status},${formatTime(record.timeIn)}\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `attendance_report_${getLocalISODate()}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
    
    showNotification('Data exported successfully!', 'success');
}

// Open edit employee modal
function openEditModal(empId) {
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const employee = employees.find(e => e.id == empId);
    
    if (!employee) return;
    
    document.getElementById('editEmployeeId').value = employee.id;
    document.getElementById('editEmployeeName').value = employee.name;
    document.getElementById('editScheduleStart').value = employee.scheduleStart;
    document.getElementById('editScheduleEnd').value = employee.scheduleEnd;
    const esnVal = employee.scheduleNotes || '';
    document.querySelectorAll('#esnBtnGroup .sn-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.value === esnVal);
    });
    
    // Set checkboxes
    const days = employee.weeklyDays.split(', ');
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].forEach(day => {
        const checkbox = document.getElementById('editDay' + day);
        checkbox.checked = days.includes(day);
    });

    // Load special days
    const editSpecialList = document.getElementById('editSpecialDaysList');
    editSpecialList.innerHTML = '';
    (employee.specialDays || []).forEach(sd => addSpecialDayRow('editSpecialDaysList', sd));
    
    const modal = new bootstrap.Modal(document.getElementById('editEmployeeModal'));
    modal.show();
}

// Save edited employee
function saveEditEmployee() {
    const empId = document.getElementById('editEmployeeId').value;
    const name = document.getElementById('editEmployeeName').value;
    const scheduleStart = document.getElementById('editScheduleStart').value;
    const scheduleEnd = document.getElementById('editScheduleEnd').value;
    const scheduleNotes = document.querySelector('#esnBtnGroup .sn-btn.active')?.dataset.value || '';
    const isFloat = scheduleNotes.toUpperCase() === 'FLOAT';

    const days = [];
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].forEach(day => {
        const checkbox = document.getElementById('editDay' + day);
        if (checkbox.checked) days.push(day);
    });

    if (!name || (!isFloat && (!scheduleStart || !scheduleEnd || days.length === 0))) {
        showNotification('Please fill all fields', 'warning');
        return;
    }

    const fmtT = (time) => {
        if (!time) return '';
        const [hours, minutes] = time.split(':');
        const hour = parseInt(hours);
        const ampm = hour >= 12 ? 'PM' : 'AM';
        return `${hour % 12 || 12}:${minutes} ${ampm}`;
    };

    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const index = employees.findIndex(e => e.id == empId);

    if (index !== -1) {
        employees[index] = {
            ...employees[index],
            name: name,
            scheduleStart: scheduleStart,
            scheduleEnd: scheduleEnd,
            scheduleDisplay: isFloat ? 'Float' : `${fmtT(scheduleStart)} - ${fmtT(scheduleEnd)}`,
            scheduleNotes: scheduleNotes,
            weeklyDays: days.join(', ') || '',
            specialDays: getSpecialDays('editSpecialDaysList')
        };
        storage.setItem('employees', JSON.stringify(employees));
    }

    document.activeElement?.blur();
    const modal = bootstrap.Modal.getInstance(document.getElementById('editEmployeeModal'));
    modal.hide();
    updateDashboard();
    showNotification('Employee updated successfully!', 'success');

    syncDashboardToSheets();
    syncFullState();
}

// Open schedule view modal
function openScheduleModal(empId) {
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const employee = employees.find(e => e.id == empId);
    
    if (!employee) return;
    
    document.getElementById('scheduleEmployeeName').textContent = employee.name;
    document.getElementById('scheduleTime').textContent = employee.scheduleDisplay;
    document.getElementById('scheduleDays').textContent = employee.weeklyDays;
    
    // Show special days
    let specialSection = document.getElementById('scheduleSpecialDays');
    if (!specialSection) {
        specialSection = document.createElement('div');
        specialSection.id = 'scheduleSpecialDays';
        specialSection.className = 'mb-3';
        document.getElementById('scheduleDays').parentElement.after(specialSection);
    }
    if (employee.specialDays && employee.specialDays.length > 0) {
        const fmt = t => formatTime(t);
        specialSection.innerHTML = `<label class="form-label fw-bold">Special Day Schedules</label><ul class="mb-0">${employee.specialDays.map(s => `<li><strong>${s.day}:</strong> ${fmt(s.start)} - ${fmt(s.end)}</li>`).join('')}</ul>`;
        specialSection.style.display = '';
    } else {
        specialSection.style.display = 'none';
    }
    
    // Show/hide schedule notes section
    const notesSection = document.getElementById('scheduleNotesSection');
    const notesDisplay = document.getElementById('scheduleNotesDisplay');
    if (employee.scheduleNotes) {
        notesDisplay.textContent = employee.scheduleNotes;
        notesSection.classList.add('show');
    } else {
        notesSection.classList.remove('show');
    }
    
    const modal = new bootstrap.Modal(document.getElementById('scheduleModal'));
    modal.show();
}

// Confirm delete employee
function confirmDelete(empId) {
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    const employee = employees.find(e => e.id == empId);
    
    if (!employee) return;
    
    document.getElementById('deleteEmployeeId').value = empId;
    document.getElementById('deleteEmployeeName').textContent = employee.name;
    
    const modal = new bootstrap.Modal(document.getElementById('deleteModal'));
    modal.show();
}

// Delete employee
async function deleteEmployee() {
    const empId = document.getElementById('deleteEmployeeId').value;
    let employees = JSON.parse(storage.getItem('employees') || '[]');
    
    const employee = employees.find(e => e.id == empId);
    const employeeName = employee ? employee.name : null;
    const employeeDept = employee ? employee.department : null;
    
    employees = employees.filter(e => e.id != empId);
    storage.setItem('employees', JSON.stringify(employees));
    
    if (employeeName && employeeDept) {
        let attendanceData = loadAttendanceData();
        attendanceData = attendanceData.filter(record => !(record.name === employeeName && record.department === employeeDept));
        saveAttendanceData(attendanceData);
    }

    // 1. Google Sheet: remove this employee row only in the current calendar month block
    if (employeeName && employeeDept) {
        const delNow = new Date();
        await syncToGoogleSheets('deleteEmployee', {
            employeeName: employeeName,
            department: employeeDept,
            blockYear: delNow.getFullYear(),
            blockMonth: delNow.getMonth()
        });
    }

    // 2. Sync full state so all devices get the updated list
    await syncFullState();
    
    document.activeElement?.blur();
    const modal = bootstrap.Modal.getInstance(document.getElementById('deleteModal'));
    modal.hide();
    updateDashboard();
    updateDailyReport();
    
    syncDashboardToSheets();
    syncWeeklyReportToSheets();
    
    showNotification('Employee and all attendance records deleted successfully!', 'success');
}

// Edit attendance record
function editAttendanceRecord(recordId) {
    const attendanceData = loadAttendanceData();
    const record = attendanceData.find(r => r.id === recordId);
    
    if (!record) return;
    
    // Populate form with record data
    document.getElementById('employeeName').value = record.name;
    document.getElementById('attendanceDate').value = record.date;
    document.getElementById('attendanceStatus').value = record.status;
    document.getElementById('timeIn').value = record.timeIn || '';
    document.getElementById('timeOut').value = record.timeOut || '';
    
    // Delete the old record
    deleteAttendanceRecord(recordId, true);
    
    // Close ALL open modals cleanly
    document.querySelectorAll('.modal.show').forEach(m => {
        const instance = bootstrap.Modal.getInstance(m);
        if (instance) instance.hide();
    });
    // Remove any lingering backdrops and body class
    setTimeout(() => {
        document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
        document.body.classList.remove('modal-open');
        document.body.style.removeProperty('overflow');
        document.body.style.removeProperty('padding-right');
        document.getElementById('attendanceForm').scrollIntoView({ behavior: 'smooth' });
        showNotification('Edit the record and submit to update', 'info');
    }, 300);
}

// Delete attendance record
async function deleteAttendanceRecord(recordId, silent = false) {
    let attendanceData = loadAttendanceData();
    const deletedRecord = attendanceData.find(record => record.id === recordId);
    
    attendanceData = attendanceData.filter(record => record.id !== recordId);
    saveAttendanceData(attendanceData);
    
    updateDailyReport();
    updateDashboard();
    
    // Auto-refresh weekly report if modal is open
    const weeklyModal = document.getElementById('weeklyReportModal');
    if (weeklyModal && weeklyModal.classList.contains('show')) {
        const monthYearSelect = document.getElementById('monthYearSelect');
        const weekFilter = document.getElementById('weekFilterModal');
        
        if (monthYearSelect && monthYearSelect.value) {
            const [selectedYear, selectedMonthStr] = monthYearSelect.value.split('-');
            const selectedMonth = parseInt(selectedMonthStr) - 1;
            const year = parseInt(selectedYear);
            const weekNumber = weekFilter ? parseInt(weekFilter.value) : null;
            
            if (weekNumber) {
                updateWeeklyReportWithFilter(selectedMonth, year, weekNumber);
            } else {
                updateWeeklyReportWithFilter(selectedMonth, year);
            }
        } else {
            const reportDatePicker = document.getElementById('reportDatePicker');
            if (reportDatePicker && reportDatePicker.value) {
                updateWeeklyReportByDate();
            } else {
                const mp = document.getElementById('reportMonthSummaryPicker');
                if (mp && !mp.value) mp.value = getLocalMonthInputValue();
                updateWeeklyReportMonthSummary();
            }
        }
    }
    
    // Auto-refresh employee details modal if open
    const employeeModal = document.getElementById('employeeDetailsModal');
    if (employeeModal && employeeModal.classList.contains('show') && deletedRecord) {
        const modal = bootstrap.Modal.getInstance(employeeModal);
        if (modal) {
            modal.hide();
            setTimeout(() => {
                viewEmployeeDetails(deletedRecord.name);
            }, 300);
        }
    }
    
    // Sync deleted state to Google Sheets FIRST, then notify
    await syncFullState();
    syncDashboardToSheets();
    syncWeeklyReportToSheets();
    
    if (!silent) {
        showNotification('Attendance record deleted successfully!', 'success');
    }
}

// View employee attendance details
function viewEmployeeDetails(employeeName) {
    const attendanceData = loadAttendanceData();
    
    // Get selected date or week from the new filters
    const reportDatePicker = document.getElementById('reportDatePicker');
    
    let filteredRecords = [];
    let periodText = '';
    
    if (reportDatePicker && reportDatePicker.value) {
        // Filter by selected date
        const selectedDate = reportDatePicker.value;
        filteredRecords = attendanceData.filter(r => {
            return r.name === employeeName && r.department === currentDepartment && r.date === selectedDate;
        });
        
        const date = new Date(selectedDate + 'T00:00:00');
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        periodText = date.toLocaleDateString('en-US', options);
        
    } else {
        // No filter selected, show current month
        const now = new Date();
        const year = now.getFullYear();
        const month = now.getMonth();
        
        filteredRecords = attendanceData.filter(r => {
            if (r.name !== employeeName || r.department !== currentDepartment) return false;
            
            const recordDate = new Date(r.date + 'T00:00:00');
            return recordDate.getMonth() === month && recordDate.getFullYear() === year;
        });
        
        const monthName = now.toLocaleDateString('en-US', { month: 'long' });
        periodText = `${monthName} ${year}`;
    }
    
    if (filteredRecords.length === 0) {
        showNotification(`No attendance records found for ${employeeName} in ${periodText}`, 'info');
        return;
    }
    
    // Sort by date descending, then by id descending (latest record per date wins)
    filteredRecords.sort((a, b) => new Date(b.date) - new Date(a.date) || b.id - a.id);
    // After dedup, sort alphabetically by name then by date descending
    // (dedup happens below, final sort applied after)

    // Deduplicate: keep only the first (latest id) record per date
    const seenDates = new Set();
    filteredRecords = filteredRecords.filter(r => {
        if (seenDates.has(r.date)) return false;
        seenDates.add(r.date);
        return true;
    });

    // Remove existing modal if any
    const existingModal = document.getElementById('employeeDetailsModal');
    if (existingModal) existingModal.remove();

    // Build modal using DOM to avoid injection
    const modalDiv = document.createElement('div');
    modalDiv.className = 'modal fade';
    modalDiv.id = 'employeeDetailsModal';
    modalDiv.tabIndex = -1;
    modalDiv.innerHTML = `
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <div>
                        <h5 class="modal-title"><i class="bi bi-person-circle me-2"></i><span id="detailsModalName"></span></h5>
                        <p class="text-muted mb-0" style="font-size: 0.9rem;" id="detailsModalPeriod"></p>
                    </div>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <table class="table table-hover table-bordered">
                        <thead>
                            <tr>
                                <th>Date</th><th>Status</th><th>Time In</th>
                                <th>Time Out</th><th>Total Hours</th>
                                <th class="text-center">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="detailsModalBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
    document.body.appendChild(modalDiv);

    // Blur focused element on hide to prevent aria-hidden focus conflict
    modalDiv.addEventListener('hide.bs.modal', function() {
        if (document.activeElement && modalDiv.contains(document.activeElement)) {
            document.activeElement.blur();
        }
    });

    // Set text safely via textContent
    document.getElementById('detailsModalName').textContent = `${employeeName} - Attendance Details`;
    document.getElementById('detailsModalPeriod').textContent = periodText;

    // Build rows via DOM
    const tbody = document.getElementById('detailsModalBody');
    filteredRecords.forEach(r => {
        const safeRId = parseInt(r.id, 10);
        const recordDate = new Date(r.date + 'T00:00:00');
        const formattedDate = recordDate.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });

        const tr = document.createElement('tr');

        const tdDate = document.createElement('td');
        tdDate.textContent = formattedDate;

        const tdStatus = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `badge-status ${getStatusBadgeClass(r.status)}`;
        badge.textContent = getDisplayStatus(r.status);
        tdStatus.appendChild(badge);

        const tdIn = document.createElement('td');
        tdIn.textContent = formatTime(r.timeIn);

        const tdOut = document.createElement('td');
        tdOut.textContent = formatTime(r.timeOut);

        const tdHours = document.createElement('td');
        tdHours.className = 'fw-bold';
        tdHours.textContent = `${r.totalHours || 0} hrs`;

        const tdActions = document.createElement('td');
        tdActions.className = 'text-center';
        const editBtn = document.createElement('button');
        editBtn.className = 'btn btn-sm btn-outline-primary';
        editBtn.title = 'Edit';
        editBtn.innerHTML = '<i class="bi bi-pencil"></i>';
        editBtn.addEventListener('click', () => editAttendanceRecord(safeRId));

        const delBtn = document.createElement('button');
        delBtn.className = 'btn btn-sm btn-outline-danger';
        delBtn.title = 'Delete';
        delBtn.innerHTML = '<i class="bi bi-trash"></i>';
        delBtn.addEventListener('click', () => deleteAttendanceRecord(safeRId));

        tdActions.appendChild(editBtn);

        // Show reason button if record has any reason (Early Out, etc.)
        const reasonText = r.reason || r.attendanceReason || '';
        if (reasonText) {
            const reasonBtn = document.createElement('button');
            reasonBtn.className = 'btn btn-sm btn-outline-warning ms-1';
            reasonBtn.title = 'View Reason';
            reasonBtn.innerHTML = '<i class="bi bi-chat-left-text"></i>';
            reasonBtn.addEventListener('click', () => showReasonModal(reasonText, formattedDate));
            tdActions.appendChild(reasonBtn);
        }

        tdActions.appendChild(delBtn);

        tr.append(tdDate, tdStatus, tdIn, tdOut, tdHours, tdActions);
        tbody.appendChild(tr);
    });

    const modal = new bootstrap.Modal(modalDiv);
    modal.show();
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', async function() {
    updateDateTime();
    setInterval(updateDateTime, 1000); // Update every second
    setDefaultDate();

    // Always sync from Google Sheets on load so all devices stay up to date
    showNotification('Syncing data from server...', 'info');
    const loaded = await loadFromSheets();
    if (loaded) showNotification('Data synced from server!', 'success');
    
    updateDailyReport();
    updateWeeklyReportWithFilter();
    updateDashboard();
    initScheduleOverrideControls();
    
    // Set initial RV theme
    document.body.classList.add('rv-active');

    // Blur focus on any modal hide to prevent aria-hidden focus conflict
    document.querySelectorAll('.modal').forEach(modalEl => {
        modalEl.addEventListener('hide.bs.modal', function() {
            if (document.activeElement && this.contains(document.activeElement)) {
                document.activeElement.blur();
            }
        });
    });

    // Restore department state after refresh
    const savedDept = sessionStorage.getItem('activeDepartment');
    if (savedDept && savedDept !== 'rv') {
        switchDepartment(savedDept);
    }

    // Restore admin panel state after refresh
    if (sessionStorage.getItem('adminPanelOpen') === '1') {
        openAdminPanel();
    }
    
    // Auto-calculate end time when start time changes
    const scheduleStartInput = document.getElementById('scheduleStart');
    const scheduleEndInput = document.getElementById('scheduleEnd');
    
    if (scheduleStartInput && scheduleEndInput) {
        scheduleStartInput.addEventListener('change', function() {
            if (this.value) {
                const [hours, minutes] = this.value.split(':');
                const startHour = parseInt(hours);
                const notesVal = (document.querySelector('#snBtnGroup .sn-btn.active')?.dataset.value || '').toUpperCase();
                const hoursToAdd = notesVal.includes('OT') ? 10 : 9;
                const endHour = (startHour + hoursToAdd) % 24;
                scheduleEndInput.value = `${endHour.toString().padStart(2, '0')}:${minutes}`;
            }
        });
    }
    
    // Auto-capture current time for Time In field
    const timeInInput = document.getElementById('timeIn');
    if (timeInInput) {
        // Make time input readonly and remove picker
        timeInInput.type = 'text';
        timeInInput.readOnly = true;
        timeInInput.style.backgroundColor = '#f8f9fa';
        timeInInput.style.cursor = 'not-allowed';
        timeInInput.placeholder = '';

        timeInInput.addEventListener('focus', function() {
            // Only auto-fill if field is empty
            if (!this.value) {
                const now = new Date();
                const hours = now.getHours().toString().padStart(2, '0');
                const minutes = now.getMinutes().toString().padStart(2, '0');
                this.value = `${hours}:${minutes}`;

                showNotification('Time automatically captured and locked for security', 'info');
            }
        });
    }
    
    // Setup Time Out field — disabled by default until employee has existing Time In record
    const timeOutInitial = document.getElementById('timeOut');
    timeOutInitial.disabled = true;
    timeOutInitial.style.backgroundColor = '#e9ecef';
    timeOutInitial.style.cursor = 'not-allowed';

    // Setup Time Out field
    addManualTimeoutButton();

    // Add employee name change listener for emergency timeout
    addEmployeeNameListener();

    // Disable Time In when status is Absent, but never enable Time Out from here
    const statusSelect = document.getElementById('attendanceStatus');
    const timeInField = document.getElementById('timeIn');
    const timeOutField = document.getElementById('timeOut');
    statusSelect.addEventListener('change', function() {
        const isAbsent = this.value === 'Absent' || this.value === 'Sick Leave';
        if (isAbsent) {
            timeInField.value = '';
            timeInField.disabled = true;
            timeInField.readOnly = true;
            timeInField.style.backgroundColor = '#e9ecef';
            timeInField.style.cursor = 'not-allowed';
        } else {
            timeInField.disabled = false;
            timeInField.readOnly = true;
            timeInField.style.backgroundColor = '#f8f9fa';
            timeInField.style.cursor = 'not-allowed';
        }
        // Time Out is NEVER enabled from status change — only enabled when existing record found
    });
    
    // Refresh weekly report when modal opens — full month summary, default = this month
    const weeklyReportModal = document.getElementById('weeklyReportModal');
    if (weeklyReportModal) {
        weeklyReportModal.addEventListener('show.bs.modal', function() {
            document.getElementById('reportDatePicker').value = '';
            const mp = document.getElementById('reportMonthSummaryPicker');
            if (mp) mp.value = getLocalMonthInputValue();
            updateWeeklyReportMonthSummary();
        });
    }
});



// Update weekly report for current week (Mon-Fri)
function updateWeeklyReportCurrentWeek(weekStart, weekEnd) {
    const modalTableBody = document.getElementById('weeklyReportTableModal');
    if (!modalTableBody) return;
    
    const attendanceData = loadAttendanceData();
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    
    const filteredData = attendanceData.filter(record => {
        if (record.department !== currentDepartment) return false;
        return record.date >= weekStart && record.date <= weekEnd;
    });
    
    if (filteredData.length === 0) {
        modalTableBody.innerHTML = `
            <tr>
                <td colspan="5" class="text-center text-muted py-5">
                    <i class="bi bi-calendar-x fs-1 d-block mb-3 text-muted opacity-50"></i>
                    <h6 class="text-muted">No Data This Week</h6>
                    <p class="small mb-0">No attendance records for the current week</p>
                </td>
            </tr>
        `;
        return;
    }
    
    const employeeRecords = {};
    filteredData.forEach(record => {
        if (!employeeRecords[record.name]) {
            const employee = employees.find(e => e.name === record.name && e.department === currentDepartment);
            employeeRecords[record.name] = {
                name: record.name,
                scheduleDisplay: employee ? employee.scheduleDisplay : 'Not Set',
                totalHours: 0,
                totalDays: 0
            };
        }
        employeeRecords[record.name].totalHours += parseFloat(record.totalHours || 0);
        if (isCountedStatus(record.status)) employeeRecords[record.name].totalDays++;
    });
    
    modalTableBody.innerHTML = Object.values(employeeRecords).sort((a, b) => a.name.localeCompare(b.name)).map(data => `
        <tr>
            <td><strong>${data.name}</strong></td>
            <td class="text-center"><span class="badge bg-info text-dark">${data.scheduleDisplay}</span></td>
            <td class="text-center fw-bold text-primary">${data.totalHours.toFixed(2)} hrs</td>
            <td class="text-center fw-bold text-success">${data.totalDays} day${data.totalDays !== 1 ? 's' : ''}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-outline-info" onclick="viewEmployeeDetails('${data.name}')" title="View Details"><i class="bi bi-eye"></i></button>
            </td>
        </tr>
    `).join('');
}

// Full month summary for Weekly Report modal (month = 0–11)
function showMonthSummary(year, month) {
    const monthName = new Date(year, month, 1).toLocaleDateString('en-US', { month: 'long' });

    const dateElement = document.getElementById('selectedDateModal');
    if (dateElement) dateElement.textContent = `${monthName} ${year} — Monthly Summary`;

    const modalTableBody = document.getElementById('weeklyReportTableModal');
    if (!modalTableBody) return;

    const attendanceData = loadAttendanceData();
    const employees = JSON.parse(storage.getItem('employees') || '[]');

    const filteredData = attendanceData.filter(r => {
        if (r.department !== currentDepartment) return false;
        const d = new Date(r.date + 'T00:00:00');
        return d.getMonth() === month && d.getFullYear() === year;
    });

    if (filteredData.length === 0) {
        modalTableBody.innerHTML = `
            <tr><td colspan="5" class="text-center text-muted py-4">
                <i class="bi bi-inbox fs-1 d-block mb-2"></i>No attendance records for ${monthName} ${year}
            </td></tr>`;
        return;
    }

    const empMap = {};
    filteredData.forEach(r => {
        if (!empMap[r.name]) {
            const emp = employees.find(e => e.name === r.name && e.department === currentDepartment);
            empMap[r.name] = { name: r.name, scheduleDisplay: emp ? emp.scheduleDisplay : 'Not Set', totalHours: 0, dates: new Set() };
        }
        empMap[r.name].totalHours += parseFloat(r.totalHours || 0);
        if (isCountedStatus(r.status)) empMap[r.name].dates.add(r.date);
    });

    modalTableBody.innerHTML = Object.values(empMap)
        .sort((a, b) => a.name.localeCompare(b.name))
        .map(d => `
        <tr>
            <td><strong>${d.name}</strong></td>
            <td><span class="badge bg-info text-dark">${d.scheduleDisplay}</span></td>
            <td class="text-center fw-bold text-primary">${d.totalHours.toFixed(2)} hrs</td>
            <td class="text-center fw-bold text-success">${d.dates.size} day${d.dates.size !== 1 ? 's' : ''}</td>
            <td class="text-center"><button class="btn btn-sm btn-outline-info" onclick="viewEmployeeMonthDetails('${d.name}', ${month}, ${year})" title="View Details"><i class="bi bi-eye"></i></button></td>
        </tr>`).join('');
}

/** Uses #reportMonthSummaryPicker; clears single-day filter. Default month = current. */
function updateWeeklyReportMonthSummary() {
    const monthInput = document.getElementById('reportMonthSummaryPicker');
    const datePicker = document.getElementById('reportDatePicker');
    if (datePicker) datePicker.value = '';

    let year;
    let month;
    if (monthInput && monthInput.value) {
        const parts = monthInput.value.split('-');
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
    } else {
        const now = new Date();
        year = now.getFullYear();
        month = now.getMonth();
        if (monthInput) monthInput.value = getLocalMonthInputValue(now);
    }
    showMonthSummary(year, month);
}

// View employee full month attendance details
function viewEmployeeMonthDetails(employeeName, month, year) {
    const attendanceData = loadAttendanceData();
    const monthName = new Date(year, month, 1).toLocaleDateString('en-US', { month: 'long' });

    let records = attendanceData.filter(r => {
        if (r.name !== employeeName || r.department !== currentDepartment) return false;
        const d = new Date(r.date + 'T00:00:00');
        return d.getMonth() === month && d.getFullYear() === year;
    });

    if (records.length === 0) {
        showNotification(`No records for ${employeeName} in ${monthName} ${year}`, 'info');
        return;
    }

    // Deduplicate by date (keep latest id)
    records.sort((a, b) => new Date(a.date) - new Date(b.date) || b.id - a.id);
    const seen = new Set();
    records = records.filter(r => { if (seen.has(r.date)) return false; seen.add(r.date); return true; });

    const existing = document.getElementById('employeeDetailsModal');
    if (existing) existing.remove();

    const modalDiv = document.createElement('div');
    modalDiv.className = 'modal fade';
    modalDiv.id = 'employeeDetailsModal';
    modalDiv.tabIndex = -1;
    // Calculate week ranges for this month
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const weeks = [];
    for (let w = 1; w <= 4; w++) {
        const startDay = (w - 1) * 7 + 1;
        const endDay = Math.min(w * 7, daysInMonth);
        if (startDay <= daysInMonth) weeks.push({ w, startDay, endDay });
    }

    modalDiv.innerHTML = `
        <div class="modal-dialog modal-lg">
            <div class="modal-content">
                <div class="modal-header">
                    <div>
                        <h5 class="modal-title"><i class="bi bi-person-circle me-2"></i>${employeeName} — Attendance Details</h5>
                        <p class="text-muted mb-0" style="font-size:0.9rem" id="detailsModalPeriodLabel">${monthName} ${year}</p>
                    </div>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <div class="d-flex align-items-center gap-2 mb-3">
                        <label class="mb-0 small fw-semibold">Week</label>
                        <input type="number" class="form-control form-control-sm" id="detailWeekInput" placeholder="1-${weeks.length}" min="0" max="${weeks.length}" style="width:80px" oninput="filterDetailWeek(this, '${employeeName}', ${month}, ${year})">
                        <span class="text-muted small" id="detailWeekLabel"></span>
                    </div>
                    <table class="table table-hover table-bordered">
                        <thead><tr><th>Date</th><th>Status</th><th>Time In</th><th>Time Out</th><th>Total Hours</th><th class="text-center">Actions</th></tr></thead>
                        <tbody id="detailsModalBody"></tbody>
                    </table>
                </div>
            </div>
        </div>`;
    document.body.appendChild(modalDiv);

    modalDiv.addEventListener('hide.bs.modal', () => { if (document.activeElement && modalDiv.contains(document.activeElement)) document.activeElement.blur(); });

    const tbody = document.getElementById('detailsModalBody');
    records.forEach(r => {
        const safeRId = parseInt(r.id, 10);
        const tr = document.createElement('tr');
        const fmtDate = new Date(r.date + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });

        const tdDate = document.createElement('td'); tdDate.textContent = fmtDate;

        const tdStatus = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `badge-status ${getStatusBadgeClass(r.status)}`;
        badge.textContent = getDisplayStatus(r.status);
        tdStatus.appendChild(badge);

        const tdIn = document.createElement('td'); tdIn.textContent = formatTime(r.timeIn);
        const tdOut = document.createElement('td'); tdOut.textContent = formatTime(r.timeOut);
        const tdHrs = document.createElement('td'); tdHrs.className = 'fw-bold'; tdHrs.textContent = `${r.totalHours || 0} hrs`;

        const tdActions = document.createElement('td');
        tdActions.className = 'text-center';

        const editBtn = document.createElement('button');
        editBtn.className = 'btn btn-sm btn-outline-primary';
        editBtn.title = 'Edit';
        editBtn.innerHTML = '<i class="bi bi-pencil"></i>';
        editBtn.addEventListener('click', () => editAttendanceRecord(safeRId));
        tdActions.appendChild(editBtn);

        const reasonText = r.reason || r.attendanceReason || '';
        if (reasonText) {
            const reasonBtn = document.createElement('button');
            reasonBtn.className = 'btn btn-sm btn-outline-warning ms-1';
            reasonBtn.title = 'View Reason';
            reasonBtn.innerHTML = '<i class="bi bi-chat-left-text"></i>';
            reasonBtn.addEventListener('click', () => showReasonModal(reasonText, fmtDate));
            tdActions.appendChild(reasonBtn);
        }

        const delBtn = document.createElement('button');
        delBtn.className = 'btn btn-sm btn-outline-danger ms-1';
        delBtn.title = 'Delete';
        delBtn.innerHTML = '<i class="bi bi-trash"></i>';
        delBtn.addEventListener('click', () => deleteAttendanceRecord(safeRId));
        tdActions.appendChild(delBtn);

        tr.append(tdDate, tdStatus, tdIn, tdOut, tdHrs, tdActions);
        tbody.appendChild(tr);
    });

    new bootstrap.Modal(modalDiv).show();
}

// Filter attendance detail modal by week
function filterDetailWeek(input, employeeName, month, year) {
    const weekNumber = parseInt(input.value) || 0;
    const attendanceData = loadAttendanceData();
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    let records = attendanceData.filter(r => {
        if (r.name !== employeeName || r.department !== currentDepartment) return false;
        const d = new Date(r.date + 'T00:00:00');
        if (d.getMonth() !== month || d.getFullYear() !== year) return false;
        if (!weekNumber) return true;
        const day = d.getDate();
        const startDay = (weekNumber - 1) * 7 + 1;
        const endDay = Math.min(weekNumber * 7, daysInMonth);
        return day >= startDay && day <= endDay;
    });

    records.sort((a, b) => new Date(a.date) - new Date(b.date) || b.id - a.id);
    const seen = new Set();
    records = records.filter(r => { if (seen.has(r.date)) return false; seen.add(r.date); return true; });

    const monthName = new Date(year, month, 1).toLocaleDateString('en-US', { month: 'long' });
    const periodLabel = document.getElementById('detailsModalPeriodLabel');
    const weekLabel = document.getElementById('detailWeekLabel');
    if (periodLabel) {
        if (!weekNumber) {
            periodLabel.textContent = `${monthName} ${year}`;
            if (weekLabel) weekLabel.textContent = 'All';
        } else {
            const startDay = (weekNumber - 1) * 7 + 1;
            const endDay = Math.min(weekNumber * 7, daysInMonth);
            periodLabel.textContent = `Week ${weekNumber} — ${monthName} ${startDay}-${endDay}, ${year}`;
            if (weekLabel) weekLabel.textContent = `${monthName} ${startDay}–${endDay}`;
        }
    }

    const tbody = document.getElementById('detailsModalBody');
    tbody.innerHTML = '';
    if (records.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3"><i class="bi bi-inbox fs-4 d-block mb-1"></i>No records</td></tr>`;
        return;
    }
    records.forEach(r => {
        const safeRId = parseInt(r.id, 10);
        const tr = document.createElement('tr');
        const fmtDate = new Date(r.date + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
        const tdDate = document.createElement('td'); tdDate.textContent = fmtDate;
        const tdStatus = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `badge-status ${getStatusBadgeClass(r.status)}`;
        badge.textContent = getDisplayStatus(r.status);
        tdStatus.appendChild(badge);
        const tdIn = document.createElement('td'); tdIn.textContent = formatTime(r.timeIn);
        const tdOut = document.createElement('td'); tdOut.textContent = formatTime(r.timeOut);
        const tdHrs = document.createElement('td'); tdHrs.className = 'fw-bold'; tdHrs.textContent = `${r.totalHours || 0} hrs`;
        const tdActions = document.createElement('td'); tdActions.className = 'text-center';
        const editBtn = document.createElement('button');
        editBtn.className = 'btn btn-sm btn-outline-primary'; editBtn.title = 'Edit';
        editBtn.innerHTML = '<i class="bi bi-pencil"></i>';
        editBtn.addEventListener('click', () => editAttendanceRecord(safeRId));
        tdActions.appendChild(editBtn);
        const reasonText = r.reason || r.attendanceReason || '';
        if (reasonText) {
            const reasonBtn = document.createElement('button');
            reasonBtn.className = 'btn btn-sm btn-outline-warning ms-1'; reasonBtn.title = 'View Reason';
            reasonBtn.innerHTML = '<i class="bi bi-chat-left-text"></i>';
            reasonBtn.addEventListener('click', () => showReasonModal(reasonText, fmtDate));
            tdActions.appendChild(reasonBtn);
        }
        const delBtn = document.createElement('button');
        delBtn.className = 'btn btn-sm btn-outline-danger ms-1'; delBtn.title = 'Delete';
        delBtn.innerHTML = '<i class="bi bi-trash"></i>';
        delBtn.addEventListener('click', () => deleteAttendanceRecord(safeRId));
        tdActions.appendChild(delBtn);
        tr.append(tdDate, tdStatus, tdIn, tdOut, tdHrs, tdActions);
        tbody.appendChild(tr);
    });
}

// Update weekly report by selected date from date picker
function updateWeeklyReportByDate() {
    const datePicker = document.getElementById('reportDatePicker');
    const selectedDate = datePicker.value;

    if (!selectedDate) {
        updateWeeklyReportMonthSummary();
        return;
    }

    const monthInput = document.getElementById('reportMonthSummaryPicker');
    if (monthInput && selectedDate.length >= 7) {
        monthInput.value = selectedDate.slice(0, 7);
    }
    
    // Clear week number input when date is selected
    const weekNumEl = document.getElementById('weekNumberInput');
    if (weekNumEl) weekNumEl.value = '';
    
    const date = new Date(selectedDate + 'T00:00:00');
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    const formattedDate = date.toLocaleDateString('en-US', options);
    
    // Update display
    const dateElement = document.getElementById('selectedDateModal');
    if (dateElement) {
        dateElement.textContent = formattedDate;
    }
    
    // Filter and display data for the selected date
    updateWeeklyReportWithDateFilter(selectedDate);
}

// Update weekly report by week number input
function updateWeeklyReportByWeek() {
    const weekInput = document.getElementById('weekNumberInput');
    const datePicker = document.getElementById('reportDatePicker');
    const weekNumber = parseInt(weekInput.value);
    
    if (!weekNumber || weekNumber < 1 || weekNumber > 4) {
        clearWeeklyReportDisplay();
        return;
    }
    
    // Get selected date or use current date
    let selectedDate = datePicker.value;
    let targetDate;
    
    if (selectedDate) {
        targetDate = new Date(selectedDate + 'T00:00:00');
    } else {
        targetDate = new Date();
    }
    
    const year = targetDate.getFullYear();
    const month = targetDate.getMonth();
    
    // Calculate week range
    const startDay = (weekNumber - 1) * 7 + 1;
    const endDay = Math.min(weekNumber * 7, new Date(year, month + 1, 0).getDate());
    
    // Update display
    const dateElement = document.getElementById('selectedDateModal');
    const monthName = targetDate.toLocaleDateString('en-US', { month: 'long' });
    if (dateElement) {
        dateElement.textContent = `Week ${weekNumber} (${monthName} ${startDay}-${endDay}, ${year})`;
    }
    
    // Filter and display data for the selected week
    updateWeeklyReportWithWeekFilter(month, year, weekNumber);
}

// Clear weekly report filters
function clearWeeklyFilters() {
    document.getElementById('reportDatePicker').value = '';
    document.getElementById('weekNumberInput').value = '';
    clearWeeklyReportDisplay();
}

// Clear weekly report display
function clearWeeklyReportDisplay() {
    const dateElement = document.getElementById('selectedDateModal');
    const tableBody = document.getElementById('weeklyReportTableModal');
    
    if (dateElement) {
        dateElement.textContent = 'Select date or week to view report';
    }
    
    if (tableBody) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="5" class="text-center text-muted py-5">
                    <i class="bi bi-calendar-x fs-1 d-block mb-3 text-muted opacity-50"></i>
                    <h6 class="text-muted">No Data Available</h6>
                    <p class="small mb-0">Please select a date or week number to view attendance data</p>
                </td>
            </tr>
        `;
    }
}

// Update weekly report with date filter
function updateWeeklyReportWithDateFilter(selectedDate) {
    const modalTableBody = document.getElementById('weeklyReportTableModal');
    if (!modalTableBody) return;
    
    const attendanceData = loadAttendanceData();
    const employees = JSON.parse(storage.getItem('employees') || '[]');
    
    // Filter by selected date and current department
    const filteredData = attendanceData.filter(record => {
        return record.date === selectedDate && record.department === currentDepartment;
    });
    
    // Group by employee name
    const employeeRecords = {};
    
    filteredData.forEach(record => {
        if (!employeeRecords[record.name]) {
            const employee = employees.find(e => e.name === record.name && e.department === currentDepartment);
            employeeRecords[record.name] = {
                name: record.name,
                scheduleDisplay: employee ? employee.scheduleDisplay : 'Not Set',
                totalHours: 0,
                totalDays: 0,
                records: []
            };
        }
        employeeRecords[record.name].totalHours += parseFloat(record.totalHours || 0);
        if (isCountedStatus(record.status)) employeeRecords[record.name].totalDays++;
        employeeRecords[record.name].records.push(record);
    });
    
    if (Object.keys(employeeRecords).length === 0) {
        modalTableBody.innerHTML = `
            <tr>
                <td colspan="5" class="text-center text-muted py-4">
                    <i class="bi bi-inbox fs-1 d-block mb-2"></i>
                    No attendance data available for selected date
                </td>
            </tr>
        `;
        return;
    }
    
    modalTableBody.innerHTML = Object.values(employeeRecords).sort((a, b) => a.name.localeCompare(b.name)).map(data => {
        return `
        <tr>
            <td><strong>${data.name}</strong></td>
            <td class="text-center"><span class="badge bg-info text-dark">${data.scheduleDisplay}</span></td>
            <td class="text-center fw-bold text-primary">${data.totalHours.toFixed(2)} hrs</td>
            <td class="text-center fw-bold text-success">${data.totalDays} day${data.totalDays !== 1 ? 's' : ''}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-outline-info" onclick="viewEmployeeDetails('${data.name}')" title="View Details">
                    <i class="bi bi-eye"></i>
                </button>
            </td>
        </tr>
    `}).join('');
}

function addManualTimeoutButton() {
    // Time Out is a plain manual input, nothing to setup
}

// Show reason modal
function showReasonModal(reason, dateLabel) {
    let modal = document.getElementById('reasonViewModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.className = 'modal fade';
        modal.id = 'reasonViewModal';
        modal.tabIndex = -1;
        modal.innerHTML = `
            <div class="modal-dialog modal-sm">
                <div class="modal-content">
                    <div class="modal-header bg-warning text-dark py-2">
                        <h6 class="modal-title"><i class="bi bi-chat-left-text me-2"></i>Reason</h6>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <p class="text-muted small mb-1" id="reasonModalDate"></p>
                        <p class="mb-0" id="reasonModalText"></p>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
        modal.addEventListener('hide.bs.modal', function() {
            if (document.activeElement && modal.contains(document.activeElement)) {
                document.activeElement.blur();
            }
        });
    }
    document.getElementById('reasonModalDate').textContent = dateLabel || '';
    document.getElementById('reasonModalText').textContent = reason;
    const existing = bootstrap.Modal.getInstance(modal);
    if (existing) existing.dispose();
    new bootstrap.Modal(modal).show();
}

// Open Early Out modal
function openEarlyOutModal() {
    const now = new Date();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    document.getElementById('earlyOutTime').value = `${hours}:${minutes}`;
    document.getElementById('earlyOutReason').value = '';
    const modal = new bootstrap.Modal(document.getElementById('earlyOutModal'));
    modal.show();
}

// Submit Early Out
function submitEarlyOut() {
    const time = document.getElementById('earlyOutTime').value;
    const reason = document.getElementById('earlyOutReason').value.trim();

    if (!time) {
        showNotification('Please set a time out', 'warning');
        return;
    }

    // Create or update a provisional attendance record immediately so
    // Emergency Out pauses the working session and Return Work can resume it later.
    const employeeNameInput = document.getElementById('employeeName');
    const name = employeeNameInput.value?.trim();
    const date = document.getElementById('attendanceDate').value || getLocalISODate();
    if (!name) {
        showNotification('Select an employee first', 'warning');
        return;
    }

    const formattedTime = formatTime(time);
    const reasonText = reason ? `${reason} (Emergency Out: ${formattedTime})` : `Emergency Out: ${formattedTime}`;

    let attendanceData = loadAttendanceData();

    // Find open record for this employee (has timeIn but no timeOut)
    let record = attendanceData.find(r =>
        r.name === name && r.date === date && r.department === currentDepartment && r.timeIn && !r.timeOut
    );

    // Helper to compute total hours between two HH:MM values
    const computeHours = (inTime, outTime) => {
        if (!inTime || !outTime) return 0;
        const [inH, inM] = inTime.split(':').map(Number);
        const [outH, outM] = outTime.split(':').map(Number);
        let inTotal = inH * 60 + inM;
        let outTotal = outH * 60 + outM;
        if (outTotal < inTotal) outTotal += 24 * 60;
        return parseFloat(((outTotal - inTotal) / 60).toFixed(2));
    };

    if (record) {
        // update open record with provisional early out
        record.timeOut = time;
        record.totalHours = computeHours(record.timeIn, time);
        // Save previous status so Return can restore it
        if (!record.previousStatus) record.previousStatus = record.status || '';
        record.status = 'Undertime';
        record.reason = reasonText;
        record.earlyOut = true;
    } else {
        // No open record — create provisional early-out record
        const tzEarly = document.getElementById('timezoneSelect') ? document.getElementById('timezoneSelect').value : '';
        const newRecord = {
            id: Date.now(),
            name: name,
            date: date,
            status: 'Undertime',
            timeIn: '',
            timeOut: time,
            totalHours: 0,
            scheduleTime: '',
            lateMinutes: 0,
            reason: reasonText,
            tz: tzEarly || '',
            department: currentDepartment,
            earlyOut: true
        };
        attendanceData.push(newRecord);
        record = newRecord;
    }

    saveAttendanceData(attendanceData);

    // Set marker so Return to Work knows which record to update
    employeeNameInput.setAttribute('data-early-out-record-id', record.id);

    // Update UI: show return button, disable timeout editing
    document.getElementById('earlyOutBtn').style.display = 'none';
    document.getElementById('returnWorkBtn').style.display = 'inline-flex';
    const toField = document.getElementById('timeOut');
    if (toField) {
        toField.value = time;
        toField.disabled = true;
        toField.style.backgroundColor = '#e9ecef';
        toField.style.cursor = 'not-allowed';
    }

    document.activeElement?.blur();
    const modal = bootstrap.Modal.getInstance(document.getElementById('earlyOutModal'));
    if (modal) modal.hide();

    // Refresh UI and sync
    updateDailyReport();
    updateDashboard();
    updateEmployeeDatalist();
    syncDashboardToSheets();
    syncWeeklyReportToSheets();

    showNotification('Emergency out recorded. Use Return when back to continue your shift.', 'info');
}

// Open Return to Work modal
function openReturnToWorkModal() {
    const now = new Date();
    const hh = now.getHours().toString().padStart(2, '0');
    const mm = now.getMinutes().toString().padStart(2, '0');
    document.getElementById('returnTimeInInput').value = `${hh}:${mm}`;
    new bootstrap.Modal(document.getElementById('returnToWorkModal')).show();
}

// Submit Return to Work
function submitReturnToWork() {
    const returnTime = document.getElementById('returnTimeInInput').value;
    if (!returnTime) {
        showNotification('Please set a return time', 'warning');
        return;
    }

    const employeeNameInput = document.getElementById('employeeName');
    const earlyOutRecordId = employeeNameInput.getAttribute('data-early-out-record-id');
    if (!earlyOutRecordId) {
        showNotification('No early-out record found', 'warning');
        return;
    }

    let attendanceData = loadAttendanceData();
    const idx = attendanceData.findIndex(r => r.id == earlyOutRecordId);
    if (idx === -1) {
        showNotification('Record not found', 'warning');
        return;
    }

    // Reopen the record: clear timeOut so it shows as open (Time In only)
    // Store session1 hours for later total calculation, set returnTimeIn for tracking
    attendanceData[idx].session1Hours = parseFloat(attendanceData[idx].totalHours || 0);
    attendanceData[idx].returnTimeIn = returnTime;
    attendanceData[idx].timeOut = '';        // clear timeout — record is "open" again
    attendanceData[idx].totalHours = 0;     // reset until final timeout
    attendanceData[idx].earlyOut = false;
    // Restore original status
    if (attendanceData[idx].previousStatus) {
        attendanceData[idx].status = attendanceData[idx].previousStatus;
    }
    saveAttendanceData(attendanceData);

    document.activeElement?.blur();
    bootstrap.Modal.getInstance(document.getElementById('returnToWorkModal')).hide();

    // Fully reset the form — employee will come back and select their name to do Time Out
    document.getElementById('attendanceForm').reset();
    setDefaultDate();
    employeeNameInput.removeAttribute('data-existing-record-id');
    employeeNameInput.removeAttribute('data-early-out-record-id');
    document.getElementById('earlyOutBtn').style.display = 'inline-flex';
    document.getElementById('returnWorkBtn').style.display = 'none';
    const toField = document.getElementById('timeOut');
    toField.disabled = true;
    toField.style.backgroundColor = '#e9ecef';
    toField.style.cursor = 'not-allowed';
    const statusField = document.getElementById('attendanceStatus');
    statusField.disabled = false;
    statusField.style.backgroundColor = '';
    statusField.style.cursor = '';
    const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');
    submitBtn.disabled = false;
    submitBtn.style.opacity = '1';

    updateDailyReport();
    updateDashboard();
    updateEmployeeDatalist();

    showNotification(`Return time recorded at ${formatTime(returnTime)}. Select your name again to log Time Out.`, 'success');
}

// Add employee name change listener
function addEmployeeNameListener() {
    const employeeNameInput = document.getElementById('employeeName');
    const dateInput = document.getElementById('attendanceDate');
    const statusSelect = document.getElementById('attendanceStatus');
    const timeInInput = document.getElementById('timeIn');
    
    if (!employeeNameInput) return;
    
    // Lightweight handler while typing: don't populate records until the user
    // commits the value (change event). This prevents partial input from
    // triggering time-in/time-out modes prematurely.
    employeeNameInput.addEventListener('input', handleEmployeeNameTyping);
    employeeNameInput.addEventListener('change', handleEmployeeNameChange);
    dateInput.addEventListener('change', handleEmployeeNameChange);

    // If user is resuming after Return, enable submit only when final Time Out is entered
    const timeOutInputGlobal = document.getElementById('timeOut');
    if (timeOutInputGlobal) {
        timeOutInputGlobal.addEventListener('input', function () {
            const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');
            const employeeNameInput = document.getElementById('employeeName');
            const existingId = employeeNameInput.getAttribute('data-existing-record-id');
            if (existingId && this.value && submitBtn) {
                submitBtn.disabled = false;
                submitBtn.style.opacity = '1';
            }
        });
    }

    function handleEmployeeNameChange() {
        const selectedName = employeeNameInput.value;
        const employees = JSON.parse(storage.getItem('employees') || '[]');
        const selectedEmployee = employees.find(e => e.name === selectedName && e.department === currentDepartment);
        const attendanceData = loadAttendanceData();
        const today = getLocalISODate();
        const selectedDate = dateInput.value || today;
        // Check for resumed session (Return confirmed) awaiting final Time Out
        const resumedRecordId = employeeNameInput.getAttribute('data-existing-record-id');
        if (resumedRecordId) {
            const resumed = attendanceData.find(r => r.id == resumedRecordId);
            if (resumed && resumed.returnTimeIn) {
                // RESUMED MODE — user should enter final Time Out; keep submit disabled until they do
                const timeOutInput = document.getElementById('timeOut');
                const statusSelect = document.getElementById('attendanceStatus');
                const timeInInput = document.getElementById('timeIn');
                const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');

                statusSelect.value = resumed.status || '';
                statusSelect.disabled = true;
                statusSelect.style.backgroundColor = '#e9ecef';
                statusSelect.style.cursor = 'not-allowed';

                timeInInput.value = resumed.timeIn || '';
                timeInInput.readOnly = true;
                timeInInput.style.backgroundColor = '#f8f9fa';
                timeInInput.style.cursor = 'not-allowed';

                // Allow entering final Time Out but keep submit disabled until filled
                timeOutInput.disabled = false;
                timeOutInput.style.backgroundColor = '';
                timeOutInput.style.cursor = '';
                timeOutInput.value = '';

                if (submitBtn) {
                    submitBtn.disabled = true;
                    submitBtn.style.opacity = '0.6';
                }

                const earlyOutBtn = document.getElementById('earlyOutBtn');
                const returnWorkBtn = document.getElementById('returnWorkBtn');
                const returnSection = document.getElementById('returnToWorkSection');
                if (earlyOutBtn) earlyOutBtn.style.display = 'inline-flex';
                if (returnWorkBtn) returnWorkBtn.style.display = 'none';
                if (returnSection) returnSection.style.display = 'none';

                return; // handled
            }
        }
        
        // Check for existing record with no timeout (normal flow)
        // Also check previous day for overnight shifts (e.g. 7AM-1AM spanning midnight)
        const prevDate = (() => {
            const d = new Date(selectedDate + 'T00:00:00');
            d.setDate(d.getDate() - 1);
            const y = d.getFullYear();
            const m = String(d.getMonth() + 1).padStart(2, '0');
            const day = String(d.getDate()).padStart(2, '0');
            return `${y}-${m}-${day}`;
        })();
        const existingRecord = attendanceData.find(record => 
            record.name === selectedName && 
            (record.date === selectedDate || record.date === prevDate) && 
            record.department === currentDepartment && 
            record.timeIn && 
            !record.timeOut
        );

        // Check for early-out record (has timeOut but marked as Undertime/Early Out)
        // Also check previous day for overnight shifts
        const earlyOutRecord = attendanceData.find(record =>
            record.name === selectedName &&
            (record.date === selectedDate || record.date === prevDate) &&
            record.department === currentDepartment &&
            record.timeIn &&
            record.timeOut &&
            ((record.reason || '').toLowerCase().includes('early out') || (record.reason || '').toLowerCase().includes('emergency out')) &&
            !record.returnTimeIn
        );
        
        const returnSection = document.getElementById('returnToWorkSection');

        const earlyOutBtn = document.getElementById('earlyOutBtn');
        const returnWorkBtn = document.getElementById('returnWorkBtn');
        const timeOutInput = document.getElementById('timeOut');
        const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');

        if (existingRecord) {
            // TIME OUT MODE — status locked, timeout enabled, submit enabled
            statusSelect.value = existingRecord.status;
            statusSelect.disabled = true;
            statusSelect.style.backgroundColor = '#e9ecef';
            statusSelect.style.cursor = 'not-allowed';
            timeInInput.value = existingRecord.timeIn;
            timeInInput.readOnly = true;
            timeInInput.style.backgroundColor = '#f8f9fa';
            timeInInput.style.cursor = 'not-allowed';
            timeOutInput.disabled = false;
            timeOutInput.style.backgroundColor = '';
            timeOutInput.style.cursor = '';
            submitBtn.disabled = false;
            submitBtn.style.opacity = '1';
            employeeNameInput.setAttribute('data-existing-record-id', existingRecord.id);
            employeeNameInput.removeAttribute('data-early-out-record-id');
            // For overnight shifts: sync the date input to the record's actual date
            // so the form submit updates the correct record
            if (existingRecord.date !== selectedDate) {
                dateInput.value = existingRecord.date;
            }
            if (earlyOutBtn) earlyOutBtn.style.display = 'inline-flex';
            if (returnWorkBtn) returnWorkBtn.style.display = 'none';
            if (returnSection) returnSection.style.display = 'none';
            return; // handled
        } else if (earlyOutRecord) {
            // RETURN TO WORK MODE — status locked, show Return button
            statusSelect.value = earlyOutRecord.status;
            statusSelect.disabled = true;
            statusSelect.style.backgroundColor = '#e9ecef';
            statusSelect.style.cursor = 'not-allowed';
            timeInInput.value = earlyOutRecord.timeIn;
            timeInInput.readOnly = true;
            timeInInput.style.backgroundColor = '#f8f9fa';
            timeInInput.style.cursor = 'not-allowed';
            timeOutInput.disabled = true;
            timeOutInput.style.backgroundColor = '#e9ecef';
            timeOutInput.style.cursor = 'not-allowed';
            submitBtn.disabled = true;
            submitBtn.style.opacity = '0.6';
            employeeNameInput.setAttribute('data-early-out-record-id', earlyOutRecord.id);
            employeeNameInput.removeAttribute('data-existing-record-id');
            if (earlyOutBtn) earlyOutBtn.style.display = 'none';
            if (returnWorkBtn) returnWorkBtn.style.display = 'inline-flex';
            if (returnSection) returnSection.style.display = 'none';
            return; // handled
        } else {
            // TIME IN MODE — status editable, timeout disabled, submit disabled
            statusSelect.disabled = false;
            statusSelect.style.backgroundColor = '';
            statusSelect.style.cursor = '';
            // Time In should always reflect actual clock-in time (regular/flex alike).
            timeInInput.value = '';
            employeeNameInput.removeAttribute('data-flex');
            timeInInput.readOnly = false;
            timeInInput.disabled = false;
            timeInInput.style.backgroundColor = '';
            timeInInput.style.cursor = '';
            timeOutInput.disabled = true;
            timeOutInput.value = '';
            timeOutInput.style.backgroundColor = '#e9ecef';
            timeOutInput.style.cursor = 'not-allowed';
            submitBtn.disabled = false;
            submitBtn.style.opacity = '1';
            employeeNameInput.removeAttribute('data-existing-record-id');
            employeeNameInput.removeAttribute('data-early-out-record-id');
            statusSelect.value = '';
            if (earlyOutBtn) earlyOutBtn.style.display = 'inline-flex';
            if (returnWorkBtn) returnWorkBtn.style.display = 'none';
            if (returnSection) returnSection.style.display = 'none';
        }
    }

    // Called on each keystroke — keep this minimal: clear any existing markers
    // and disable submit until the user selects/commits a full name.
    function handleEmployeeNameTyping() {
        const employeeNameInput = document.getElementById('employeeName');
        const timeInInput = document.getElementById('timeIn');
        const timeOutInput = document.getElementById('timeOut');
        const statusSelect = document.getElementById('attendanceStatus');
        const submitBtn = document.querySelector('#attendanceForm button[type="submit"]');

        // Remove any markers from prior selection
        employeeNameInput.removeAttribute('data-existing-record-id');
        employeeNameInput.removeAttribute('data-early-out-record-id');
        employeeNameInput.removeAttribute('data-flex');

        // Reset fields to neutral typing state
        if (timeInInput) {
            timeInInput.value = '';
            timeInInput.readOnly = false;
            timeInInput.disabled = false;
            timeInInput.style.backgroundColor = '';
            timeInInput.style.cursor = '';
        }
        if (timeOutInput) {
            timeOutInput.value = '';
            timeOutInput.disabled = true;
            timeOutInput.style.backgroundColor = '#e9ecef';
            timeOutInput.style.cursor = 'not-allowed';
        }
        if (statusSelect) {
            statusSelect.value = '';
            statusSelect.disabled = false;
            statusSelect.style.backgroundColor = '';
            statusSelect.style.cursor = '';
        }
        if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.style.opacity = '0.6';
        }
    }
}

function openAdminPanel() {
    document.getElementById('adminPanel').style.display = 'block';
    document.getElementById('adminControls').style.display = 'flex';
    document.getElementById('adminControls').style.visibility = 'visible';
    document.getElementById('formCol').className = 'col-lg-3';
    document.getElementById('mainRow').classList.add('admin-open');
    const btn = document.getElementById('adminToggleBtn');
    btn.innerHTML = '<i class="bi bi-x-lg"></i>';
    btn.classList.remove('btn-outline-secondary');
    btn.classList.add('btn-secondary');
    sessionStorage.setItem('adminPanelOpen', '1');
    // Hide User Manual button when admin panel is open
    const userManualBtn = document.querySelector('[data-bs-target="#userManualModal"]');
    if (userManualBtn) userManualBtn.style.display = 'none';
}

function closeAdminPanel() {
    document.getElementById('adminPanel').style.display = 'none';
    document.getElementById('adminControls').style.display = 'none';
    document.getElementById('adminControls').style.visibility = 'hidden';
    document.getElementById('formCol').className = 'col-lg-3';
    document.getElementById('mainRow').classList.remove('admin-open');
    const btn = document.getElementById('adminToggleBtn');
    btn.innerHTML = '<i class="bi bi-shield-lock me-1"></i>Admin';
    btn.classList.remove('btn-secondary');
    btn.classList.add('btn-outline-secondary');
    sessionStorage.removeItem('adminPanelOpen');
    document.getElementById('attendanceForm').scrollIntoView({ behavior: 'smooth' });
    // Show User Manual button when admin panel is closed
    const userManualBtn = document.querySelector('[data-bs-target="#userManualModal"]');
    if (userManualBtn) userManualBtn.style.display = '';
}

function toggleAdminView() {
    const panel = document.getElementById('adminPanel');
    const isOpen = panel.style.display !== 'none';

    if (isOpen) {
        closeAdminPanel();
    } else {
        document.getElementById('adminPasswordInput').value = '';
        document.getElementById('adminPasswordError').style.display = 'none';
        new bootstrap.Modal(document.getElementById('adminPasswordModal')).show();
    }
}

function verifyAdminPassword() {
    const input = document.getElementById('adminPasswordInput');
    if (input.value === 'COMSADMIN2026!') {
        input.value = '';
        document.activeElement?.blur();
        const modalEl = document.getElementById('adminPasswordModal');
        modalEl.addEventListener('hidden.bs.modal', function handler() {
            modalEl.removeEventListener('hidden.bs.modal', handler);
            openAdminPanel();
        });
        bootstrap.Modal.getInstance(modalEl).hide();
    } else {
        document.getElementById('adminPasswordError').style.display = 'block';
        input.value = '';
        input.focus();
    }
}

// Reset Monthly Statistics
function resetMonthlyStats() {
    let confirmModal = document.getElementById('resetStatsConfirmModal');
    if (!confirmModal) {
        confirmModal = document.createElement('div');
        confirmModal.className = 'modal fade';
        confirmModal.id = 'resetStatsConfirmModal';
        confirmModal.tabIndex = -1;
        confirmModal.innerHTML = `
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">Reset Attendance Records</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">Are you sure you want to reset all attendance records? All counts will go back to zero. <strong>Google Sheets data will not be affected.</strong></div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-danger" id="resetStatsConfirmBtn">Reset</button>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(confirmModal);
    }
    const bsModal = new bootstrap.Modal(confirmModal);
    document.getElementById('resetStatsConfirmBtn').onclick = function() {
        bsModal.hide();
        // Clear only local attendance data — Google Sheets is NOT touched
        storage.setItem('attendanceData', '[]');
        updateDashboard();
        updateDailyReport();
        showNotification('Attendance records reset to zero. Google Sheets data preserved.', 'success');
    };
    bsModal.show();
}

// --- Test helper: seed sample employees + attendance records ---
function seedTestData() {
    const employees = [
        { id: Date.now() + 101, name: 'Alice', scheduleStart: '08:00', scheduleEnd: '17:00', scheduleDisplay: '08:00 - 17:00', scheduleNotes: '', weeklyDays: 'Mon, Tue, Wed, Thu, Fri', specialDays: [], department: currentDepartment },
        { id: Date.now() + 102, name: 'Bob', scheduleStart: '09:00', scheduleEnd: '18:00', scheduleDisplay: '09:00 - 18:00', scheduleNotes: '', weeklyDays: 'Mon, Tue, Wed, Thu, Fri', specialDays: [], department: currentDepartment }
    ];

    const today = getLocalISODate();
    const yesterday = (() => { const d = new Date(); d.setDate(d.getDate() - 1); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().split('T')[0]; })();

    const attendanceData = [
        // Alice: Present today -> should count
        { id: Date.now() + 201, name: 'Alice', date: today, timeIn: '08:05', timeOut: '17:10', status: 'Present', totalHours: 9, department: currentDepartment },
        // Alice: Sick Leave yesterday -> should NOT count
        { id: Date.now() + 202, name: 'Alice', date: yesterday, timeIn: '', timeOut: '', status: 'Sick Leave', totalHours: 0, department: currentDepartment },
        // Bob: Absent today -> should NOT count
        { id: Date.now() + 203, name: 'Bob', date: today, timeIn: '', timeOut: '', status: 'Absent', totalHours: 0, department: currentDepartment },
        // Bob: WFH yesterday -> should NOT count
        { id: Date.now() + 204, name: 'Bob', date: yesterday, timeIn: '09:00', timeOut: '18:00', status: 'Work From Home', totalHours: 9, department: currentDepartment }
    ];

    storage.setItem('employees', JSON.stringify(employees));
    storage.setItem('attendanceData', JSON.stringify(attendanceData));

    // Refresh UI
    updateDailyReport();
    updateWeeklyReportWithFilter();
    updateDashboard();

    showNotification('Seeded test data: Alice (Present + Sick), Bob (Absent + WFH).', 'info');
}

window.seedTestData = seedTestData;