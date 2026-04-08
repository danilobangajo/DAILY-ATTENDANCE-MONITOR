// Configuration - Paste your Google Apps Script Web App URL here
const GOOGLE_SHEETS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwqnEHkM943ag95i_ZzWmopgMLgdKcYp2BIUpib8s42xvAcdgqDlJwvrkXBmaw6Kryh/exec';

function getDeptRevision(dept) {
    try {
        const store = typeof storage !== 'undefined' ? storage : localStorage;
        const v = parseInt(store.getItem(`syncRevision_${dept}`) || '0', 10);
        return Number.isFinite(v) ? v : 0;
    } catch (e) {
        return 0;
    }
}

function setDeptRevision(dept, rev) {
    try {
        const store = typeof storage !== 'undefined' ? storage : localStorage;
        const n = parseInt(rev, 10);
        if (Number.isFinite(n)) store.setItem(`syncRevision_${dept}`, String(n));
    } catch (e) {
        // ignore
    }
}

// Sync to Google Sheets
async function syncToGoogleSheets(type, data) {
    if (!GOOGLE_SHEETS_WEB_APP_URL || GOOGLE_SHEETS_WEB_APP_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') {
        console.log('Google Sheets not configured');
        return;
    }
    
    try {
        const now = new Date();
        const currentMonth = now.toLocaleDateString('en-US', { month: 'long' });
        const currentYear = now.getFullYear();
        
        const response = await fetch(GOOGLE_SHEETS_WEB_APP_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                type: type,
                department: typeof currentDepartment !== 'undefined' ? currentDepartment : 'rv',
                month: currentMonth,
                year: currentYear,
                monthYear: `${currentMonth} ${currentYear}`,
                ...data
            })
        });
        console.log(`✅ Synced to Google Sheets (${currentMonth} ${currentYear})`);
    } catch (error) {
        console.error('❌ Sync error:', error);
    }
}

// Save full app state (employees + attendance) to Google Sheets for cross-device sharing
// Syncs BOTH departments so all devices always have complete data
async function syncFullState() {
    if (!GOOGLE_SHEETS_WEB_APP_URL || GOOGLE_SHEETS_WEB_APP_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') return;
    const store = typeof storage !== 'undefined' ? storage : localStorage;
    const allEmployees = JSON.parse(store.getItem('employees') || '[]');
    const allAttendance = JSON.parse(store.getItem('attendanceData') || '[]');

    // Split by department and sync each separately so doGet returns correct structure
    for (const dept of ['rv', 'coms']) {
        const deptEmployees = allEmployees.filter(e => e.department === dept);
        const deptAttendance = allAttendance.filter(r => r.department === dept);
        // Retry up to 3 times to ensure delete/edit propagates to remote
        let success = false;
        for (let attempt = 0; attempt < 3 && !success; attempt++) {
                try {
                await fetch(GOOGLE_SHEETS_WEB_APP_URL, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify({
                        type: 'fullState',
                        department: dept,
                        baseRevision: getDeptRevision(dept),
                        state: { employees: deptEmployees, attendanceData: deptAttendance }
                    })
                });
                success = true;
                console.log('✅ Full state synced for', dept);
            } catch (e) {
                console.error(`❌ Full state sync error for ${dept} (attempt ${attempt + 1}):`, e);
                if (attempt < 2) await new Promise(r => setTimeout(r, 1000));
            }
        }
    }
}

// Load data from Google Sheets — remote is the single source of truth.
// Fully replaces local data so deletes/edits on one device reflect on all devices.
async function loadFromSheets() {
    if (!GOOGLE_SHEETS_WEB_APP_URL || GOOGLE_SHEETS_WEB_APP_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') return false;
    try {
        const res = await fetch(GOOGLE_SHEETS_WEB_APP_URL + '?t=' + Date.now());
        const data = await res.json();
        if (!data.rv && !data.coms) return false;

        const remoteEmployees = [...(data.rv?.employees || []), ...(data.coms?.employees || [])];
        const remoteAttendance = [...(data.rv?.attendanceData || []), ...(data.coms?.attendanceData || [])];

        if (remoteEmployees.length === 0 && remoteAttendance.length === 0) return false;

        // Remote fully replaces local — no merge, so deletes propagate to all devices
        const store = typeof storage !== 'undefined' ? storage : localStorage;
        store.setItem('employees', JSON.stringify(remoteEmployees));
        store.setItem('attendanceData', JSON.stringify(remoteAttendance));
        setDeptRevision('rv', data.rv?.revision);
        setDeptRevision('coms', data.coms?.revision);
        return true;
    } catch (e) {
        console.error('❌ Load from Sheets error:', e);
    }
    return false;
}
