// --- Global Variables ---
let studentsData = {};
let transactions = [];
let activeLoans = [];

// --- Google Sheets API Configuration ---
// IMPORTANT: Replace these placeholder values with your actual credentials.
const CLIENT_ID = '310556038461-e3b6dn1vtph2vhq4g3vu39a1fi9r64s8.apps.googleusercontent.com'; 
const API_KEY = 'AIzaSyDeEKbvpaboKtojryxDcLx7iGuPCeFdv50'; 
const SPREADSHEET_ID = '1FwXvfXQtgovokYTPVQOo0Aee4J6SRtpEecVkrVhQKq8'; // <<< SET YOUR FIXED ID HERE
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

let tokenClient;
let gapiInited = false;
let gisInited = false;


// --- Utility Functions ---
function showStatus(elementId, message, type) {
    const element = document.getElementById(elementId);
    element.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
    setTimeout(() => { element.innerHTML = ''; }, 5000);
}
function formatDateTime(date) { return date.toLocaleString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
function formatDate(date) { return date.toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' }); }
function formatTime(date) { return date.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }); }


// --- Data Persistence Functions ---
function saveData() {
    localStorage.setItem('activeLoans', JSON.stringify(activeLoans));
    localStorage.setItem('transactions', JSON.stringify(transactions));
}

function loadData() {
    const savedLoans = localStorage.getItem('activeLoans');
    const savedTransactions = localStorage.getItem('transactions');

    if (savedLoans) {
        const parsedLoans = JSON.parse(savedLoans);
        activeLoans = parsedLoans.map(loan => ({ ...loan, outTime: new Date(loan.outTime) }));
    }

    if (savedTransactions) {
        const parsedTransactions = JSON.parse(savedTransactions);
        let allTransactions = parsedTransactions.map(t => ({ ...t, outTime: new Date(t.outTime), inTime: t.inTime ? new Date(t.inTime) : null }));
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        transactions = allTransactions.filter(t => t.outTime >= today);
    }
    
    updateActiveLoansTable();
    updateTransactionTable();
}


// --- Table Update Functions ---
function updateActiveLoansTable() {
    const tbody = document.getElementById('activeLoansBody');
    tbody.innerHTML = (activeLoans.length === 0) 
        ? '<tr><td colspan="7" style="text-align: center; padding: 30px;">No active loans</td></tr>'
        : activeLoans.map(loan => `
            <tr>
                <td>${loan.rollNo}</td><td>${loan.name}</td><td>${loan.roomNo}</td>
                <td>${loan.phoneNo}</td><td>${loan.equipment}</td><td>${formatDateTime(loan.outTime)}</td>
                <td><button class="return-btn" onclick="returnEquipment(${loan.id})">Return</button></td>
            </tr>`).join('');
}
function updateTransactionTable() {
    const tbody = document.getElementById('transactionBody');
    tbody.innerHTML = (transactions.length === 0)
        ? '<tr><td colspan="8" style="text-align: center; padding: 30px;">No transactions yet</td></tr>'
        : transactions.map(t => `
            <tr>
                <td>${formatDate(t.outTime)}</td><td>${t.rollNo}</td><td>${t.name}</td>
                <td>${t.phoneNo}</td><td>${t.equipment}</td><td>${formatTime(t.outTime)}</td>
                <td>${t.inTime ? formatTime(t.inTime) : '-'}</td>
                <td><span class="status-${t.status}">${t.status.toUpperCase()}</span></td>
            </tr>`).join('');
}


// --- Core Application Functions ---
async function loadStudentDatabase() {
    try {
        const response = await fetch('boarders.xlsx');
        if (!response.ok) throw new Error(`Could not find boarders.xlsx`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        if (jsonData.length === 0) throw new Error('boarders.xlsx is empty');

        studentsData = {};
        jsonData.forEach(row => {
            let rollNo = '';
            for (const key of Object.keys(row)) { if (key.toLowerCase().includes('roll')) { rollNo = String(row[key] || '').trim(); break; } }
            if (rollNo) {
                let name = '', roomNo = '', phoneNo = '';
                for (const key of Object.keys(row)) { if (key.toLowerCase().includes('name')) { name = String(row[key] || '').trim(); break; } }
                for (const key of Object.keys(row)) { if (key.toLowerCase().includes('room')) { roomNo = String(row[key] || '').trim(); break; } }
                for (const key of Object.keys(row)) { if (key.toLowerCase().includes('phone')) { phoneNo = String(row[key] || '').trim(); break; } }
                studentsData[rollNo] = { name: name || 'N/A', rollNo, roomNo: roomNo || 'N/A', phoneNo: phoneNo || 'N/A' };
            }
        });
        showStatus('issueStatus', `✅ Student database loaded: ${Object.keys(studentsData).length} records found.`, 'success');
    } catch (error) {
        showStatus('issueStatus', `❌ Error loading database: ${error.message}`, 'error');
    }
}

function lookupStudent() {
    const rollNo = document.getElementById('rollNo').value.trim();
    if (!rollNo) { showStatus('issueStatus', 'Please enter a roll number', 'error'); return; }
    if (Object.keys(studentsData).length === 0) {
        showStatus('issueStatus', 'Student database not loaded. Please check if boarders.xlsx is available.', 'error');
        return;
    }
    if (!studentsData[rollNo]) {
        showStatus('issueStatus', 'Student not found.', 'error');
        document.getElementById('studentInfo').style.display = 'none';
        return;
    }
    const student = studentsData[rollNo];
    document.getElementById('studentName').textContent = student.name;
    document.getElementById('studentRoll').textContent = student.rollNo;
    document.getElementById('studentRoom').textContent = student.roomNo;
    document.getElementById('studentPhone').textContent = student.phoneNo;
    document.getElementById('studentInfo').style.display = 'block';
    showStatus('issueStatus', '', 'success');
}

function issueEquipment() {
    const rollNo = document.getElementById('rollNo').value.trim();
    if (!studentsData[rollNo]) { showStatus('issueStatus', 'Please look up student first', 'error'); return; }

    const facilityType = document.getElementById('facilityType').value;
    if (!facilityType) { showStatus('issueStatus', 'Please select a facility type', 'error'); return; }

    let equipment = facilityType;
    if (facilityType === 'Sports') {
        const sportsEq = document.getElementById('sportsEquipment').value;
        if (!sportsEq) { showStatus('issueStatus', 'Please select a sports equipment', 'error'); return; }
        equipment = sportsEq;
    }

    if (activeLoans.find(loan => loan.rollNo === rollNo)) {
        showStatus('issueStatus', 'Student already has an active loan.', 'error'); return;
    }

    const student = studentsData[rollNo];
    const currentTime = new Date();
    const transaction = {
        id: Date.now(), rollNo, name: student.name, roomNo: student.roomNo,
        phoneNo: student.phoneNo, equipment, facility: facilityType,
        outTime: currentTime, inTime: null, status: 'active'
    };

    transactions.unshift(transaction);
    activeLoans.unshift(transaction);
    updateActiveLoansTable();
    updateTransactionTable();
    showStatus('issueStatus', `✅ Equipment issued to ${student.name}.`, 'success');
    
    saveData();
    addNewRowToSheet(transaction);

    document.getElementById('rollNo').value = '';
    document.getElementById('facilityType').value = '';
    document.getElementById('studentInfo').style.display = 'none';
    document.getElementById('sportsEquipmentGroup').style.display = 'none';
}

function returnEquipment(transactionId) {
    const loanIndex = activeLoans.findIndex(loan => loan.id === transactionId);
    const transactionIndex = transactions.findIndex(t => t.id === transactionId);

    if (loanIndex !== -1 && transactionIndex !== -1) {
        const returnTime = new Date();
        transactions[transactionIndex].inTime = returnTime;
        transactions[transactionIndex].status = 'returned';
        activeLoans.splice(loanIndex, 1);
        updateActiveLoansTable();
        updateTransactionTable();

        saveData();
        updateSheetRow(transactions[transactionIndex]);
    }
}


// --- Excel Export Function ---
function downloadLogs(facilityType) {
    let dataToExport = transactions;
    let fileName = 'all-logs.xlsx';

    if (facilityType) {
        dataToExport = transactions.filter(t => t.facility === facilityType);
        fileName = `${facilityType.toLowerCase().replace(' ', '-')}-logs.xlsx`;
    }

    if (dataToExport.length === 0) {
        alert(`No logs found for ${facilityType || 'any facility'}.`);
        return;
    }

    const formattedData = dataToExport.map(t => ({
        "Date": formatDate(t.outTime), "Roll No": t.rollNo, "Name": t.name,
        "Room": t.roomNo, "Phone": t.phoneNo, "Facility": t.facility,
        "Equipment": t.equipment, "Out Time": formatTime(t.outTime),
        "In Time": t.inTime ? formatTime(t.inTime) : 'N/A', "Status": t.status
    }));

    const worksheet = XLSX.utils.json_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Logs');
    XLSX.writeFile(workbook, fileName);
}


// --- Google Sheets API Functions ---
function gapiLoaded() { gapi.load('client', initializeGapiClient); }
async function initializeGapiClient() {
    await gapi.client.init({ apiKey: API_KEY, discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"], });
    gapiInited = true;
    maybeEnableButtons();
}
function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({ client_id: CLIENT_ID, scope: SCOPES, callback: '' });
    gisInited = true;
    maybeEnableButtons();
}
function maybeEnableButtons() {
    if (gapiInited && gisInited) {
        document.getElementById('authorize_button').style.display = 'block';
        document.getElementById('authorize_button').onclick = handleAuthClick;
        document.getElementById('signout_button').onclick = handleSignoutClick;
    }
}
function handleAuthClick() {
    tokenClient.callback = async (resp) => {
        if (resp.error !== undefined) { throw (resp); }
        document.getElementById('signout_button').style.display = 'block';
        document.getElementById('authorize_button').innerText = 'Refresh Connection';
        showStatus('g-sheet-status', '✅ Successfully connected to Google.', 'success');
        await checkSheetHeaders();
    };
    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
        tokenClient.requestAccessToken({ prompt: '' });
    }
}
function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token);
        gapi.client.setToken('');
        document.getElementById('authorize_button').innerText = 'Connect Google Account';
        document.getElementById('signout_button').style.display = 'none';
        showStatus('g-sheet-status', 'Disconnected from Google.', 'success');
    }
}
async function addNewRowToSheet(transaction) {
    if (gapi.client.getToken() === null || !SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') return;

    const facilitySheetName = transaction.facility || 'All Logs';
    const values = [[
        transaction.id, formatDate(transaction.outTime), transaction.rollNo, transaction.name,
        transaction.roomNo, transaction.phoneNo, transaction.equipment,
        formatTime(transaction.outTime), '', 'active'
    ]];
    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: `${facilitySheetName}!A:J`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: values },
        });
    } catch (err) {
        showStatus('g-sheet-status', `Error writing to sheet: ${err.result.error.message}`, 'error');
    }
}
async function updateSheetRow(transaction) {
    if (gapi.client.getToken() === null || !SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') return;

    const sheetName = transaction.facility || 'All Logs';
    try {
        const idColumnRange = `${sheetName}!A:A`;
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: idColumnRange,
        });

        const allIds = response.result.values;
        if (!allIds || allIds.length === 0) { return; }

        const rowIndex = allIds.findIndex(row => row[0] && String(row[0]) === String(transaction.id));

        if (rowIndex === -1) {
            console.error(`Transaction ID ${transaction.id} not found in sheet. Cannot update.`);
            return;
        }
        
        const rowToUpdate = rowIndex + 1;

        const updatedValues = [
            transaction.id, formatDate(transaction.outTime), transaction.rollNo, transaction.name,
            transaction.roomNo, transaction.phoneNo, transaction.equipment,
            formatTime(transaction.outTime), transaction.inTime ? formatTime(transaction.inTime) : '',
            transaction.status
        ];

        const updateRange = `${sheetName}!A${rowToUpdate}:J${rowToUpdate}`;
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [updatedValues] }
        });

        showStatus('g-sheet-status', `✅ Log updated for Roll No ${transaction.rollNo}.`, 'success');

    } catch (err) {
        showStatus('g-sheet-status', `Error updating sheet: ${err.result.error.message}`, 'error');
        console.error("Google Sheets update error:", err);
    }
}
async function checkSheetHeaders() {
    if (gapi.client.getToken() === null || !SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
        return;
    }

    const headers = ['Transaction ID', 'Date', 'Roll No', 'Name', 'Room', 'Phone', 'Equipment', 'Out Time', 'In Time', 'Status'];
    try {
        const sheetToCheck = 'Sports';
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${sheetToCheck}!A1:J1`,
        });
        if (!response.result.values || response.result.values.length === 0) {
            await gapi.client.sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID,
                range: `${sheetToCheck}!A1`,
                valueInputOption: 'RAW',
                resource: { values: [headers] },
            });
            showStatus('g-sheet-status', `Added headers to '${sheetToCheck}' sheet.`, 'success');
        }
    } catch (err) {
        showStatus('g-sheet-status', `Ensure sheets ('Sports', 'Gym', etc.) exist.`, 'error');
    }
}

// --- Event Listeners ---
window.addEventListener('load', () => {
    loadStudentDatabase();
    loadData();
    document.getElementById('facilityType').addEventListener('change', function() {
        document.getElementById('sportsEquipmentGroup').style.display = (this.value === 'Sports') ? 'block' : 'none';
    });
});
document.addEventListener('keydown', function(e) {
    if (e.key === 'F5' || (e.ctrlKey && e.key === 'r') || (e.ctrlKey && e.shiftKey && e.key === 'R')) {
        e.preventDefault();
    }
});