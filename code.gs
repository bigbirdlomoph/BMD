// ==========================================
// 1. CONFIGURATION & SETUP
// ==========================================
var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ';
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = 'Version : 690318';

// ==========================================
// 📌 1.1 DYNAMIC COLUMN MAPPING (แผนที่หัวคอลัมน์)
// ==========================================
var COL_NAME = {
    ID: 'รหัสโครงการ',
    ORDER: 'ลำดับโครงการ',
    DEPT: 'กลุ่มงาน/งาน',
    PLAN: 'แผนงาน',
    PROJ: 'โครงการ',
    ACT: 'กิจกรรมหลัก',
    SUB: 'กิจกรรมย่อย',
    TYPE: 'ประเภทงบ',
    SOURCE: 'แหล่งงบประมาณ',
    APPROVE: 'อนุมัติตามแผน',
    ALLOC: 'จัดสรร',
    SPENT: 'เบิกจ่าย',
    LOAN: 'เงินยืม',
    BAL: 'คงเหลือ (ไม่รวมเงินยืม)', // 👈 ต้องตรงกับหัวชีตเป๊ะๆ
    STATUS: 'สถานะ',
    REMARK: 'หมายเหตุการปรับ',
    M_OCT: 'ต.ค.', M_NOV: 'พ.ย.', M_DEC: 'ธ.ค.', M_JAN: 'ม.ค.', M_FEB: 'ก.พ.', M_MAR: 'มี.ค.',
    M_APR: 'เม.ย.', M_MAY: 'พ.ค.', M_JUN: 'มิ.ย.', M_JUL: 'ก.ค.', M_AUG: 'ส.ค.', M_SEP: 'ก.ย.'
};

// 📌 ฟังก์ชันอ่าน Index จากหัวคอลัมน์อัตโนมัติ
function getColumnMap(sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
        var headerName = String(headers[i]).trim();
        if (headerName !== "") {
            map[headerName] = i; 
        }
    }
    return map;
}

function doGet() {
    var template = HtmlService.createTemplateFromFile('index');
    template.appVersion = APP_VERSION;
    return template.evaluate()
        .setTitle('LPHO BMD : Loei Provincial Public Health Office Budget Management Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { 
    return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

// ==========================================
// 🔐 2. AUTHENTICATION & ADMIN SYSTEM (เพิ่มใหม่)
// ==========================================
function verifyEmailOnly(inputEmail) {
    try {
        if (!inputEmail || String(inputEmail).trim() === "") {
            return { status: 'error', message: 'กรุณาระบุอีเมล' };
        }

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var authSheet = ss.getSheetByName('auth_users');
        if (!authSheet) return { status: 'error', message: 'ไม่พบชีตชื่อ auth_users' };

        var data = authSheet.getDataRange().getValues();
        var searchEmail = String(inputEmail).trim().toLowerCase();
        var userFound = null;

        for (var i = 1; i < data.length; i++) {
            var dbEmail = String(data[i][0] || "").trim().toLowerCase();
            if (dbEmail === searchEmail && dbEmail !== "") {
                userFound = {
                    email: String(data[i][0]),
                    firstName: String(data[i][1]),
                    lastName: String(data[i][2]),
                    fullName: String(data[i][1]) + ' ' + String(data[i][2]),
                    position: String(data[i][3]),
                    department: String(data[i][4]),
                    role: String(data[i][5] || "").trim().toUpperCase(),
                    status: String(data[i][6] || "").trim().toUpperCase(),
                    phone: String(data[i][7])
                };
                break;
            }
        }

        if (userFound) {
            if (userFound.status === 'ACTIVE') {
                try {
                    var logSheet = ss.getSheetByName('auth_users_log');
                    if(!logSheet) logSheet = ss.insertSheet('auth_users_log');
                    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                    logSheet.appendRow([timestamp, userFound.email, userFound.fullName, 'LOGIN_SUCCESS', 'เข้าสู่ระบบด้วยอีเมลสำเร็จ']);
                } catch(e) {}
                return { status: 'success', user: userFound };
            } else {
                return { status: 'error', message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
            }
        } else {
            return { status: 'error', message: 'ไม่พบอีเมล [' + inputEmail + '] ในระบบ' };
        }
    } catch (e) {
        return { status: 'error', message: 'Server Error: ' + e.message };
    }
}

function logAuthAction(email, fullName, action, detail) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var logSheet = ss.getSheetByName('auth_users_log');
        if (!logSheet) {
            logSheet = ss.insertSheet('auth_users_log');
            logSheet.appendRow(['timestamp', 'email', 'full_name', 'action', 'detail']);
        }
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        logSheet.appendRow([timestamp, email, fullName, action, detail]);
    } catch (e) {}
}

function getAllUsers() {
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('auth_users');
        if(!sheet) return [];
        var data = sheet.getDataRange().getValues();
        var users = [];
        for (var i = 1; i < data.length; i++) {
            users.push({
                email: data[i][0], firstName: data[i][1], lastName: data[i][2],
                position: data[i][3], department: data[i][4], role: data[i][5],
                status: data[i][6], phone: data[i][7], rowIndex: i + 1
            });
        }
        return users;
    } catch(e) { return []; }
}

function saveUser(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('auth_users');
        if (!sheet) return {status: 'error', message: 'ไม่พบชีต auth_users'};
        
        var data = sheet.getDataRange().getValues();
        var rowIndex = -1;
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).toLowerCase() === String(form.email).toLowerCase()) {
                rowIndex = i + 1; break;
            }
        }
        var rowData = [form.email, form.firstName, form.lastName, form.position, form.department, form.role, form.status, form.phone];
        if (rowIndex > -1) { 
            sheet.getRange(rowIndex, 1, 1, 8).setValues([rowData]);
            return {status: 'success', message: 'อัปเดตผู้ใช้งานเรียบร้อย'};
        } else { 
            sheet.appendRow(rowData);
            return {status: 'success', message: 'เพิ่มผู้ใช้งานใหม่เรียบร้อย'};
        }
    } catch (e) {
        return {status: 'error', message: e.message};
    } finally { lock.releaseLock(); }
}

function deleteUser(email) {
    try {
        var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('auth_users');
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
                sheet.deleteRow(i + 1);
                return {status: 'success', message: 'ลบผู้ใช้งานเรียบร้อย'};
            }
        }
        return {status: 'error', message: 'ไม่พบอีเมลที่ต้องการลบ'};
    } catch(e) { return {status:'error', message: e.message}; }
}

// ==========================================
// 3. CORE LOGIC (MASTER DATA)
// ==========================================
function getAllMasterDataForClient() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return [];
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return [];
        data.shift(); 
        return data.map(function (r) {
            return {
                id: r[0], category: r[2], order: r[3], dept: r[4],       
                plan: r[5], project: r[6], activity: r[7], subActivity: r[8], budgetType: r[9], budgetSource: r[10],
                approved: parseFloat(String(r[15]).replace(/,/g, '')) || 0, 
                allocated: parseFloat(String(r[16]).replace(/,/g, '')) || 0, 
                spent: parseFloat(String(r[17]).replace(/,/g, '')) || 0,     
                loan: parseFloat(String(r[18]).replace(/,/g, '')) || 0,
                balance: parseFloat(String(r[19]).replace(/,/g, '')) || 0    
            };
        }).filter(function (r) { return r.id && r.project; });
    } catch (e) { return []; }
}

// ==========================================
// 4. DASHBOARD DATA
// ==========================================
function getDashboardData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return { error: "ไม่พบ Sheet" };
        var data = sheet.getDataRange().getValues();
        data.shift();

        var I_DEPT = 4, I_TYPE = 9, I_ALLOC = 16, I_SPENT = 17, I_BAL = 19, I_APPROVE = 15;
        var summary = { moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }, nonMoph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} } };
        var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        
        data.forEach(function (row) {
            var typeVal = String(row[I_TYPE] || "").trim();
            var isNonMoph = (
                typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || 
                typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1
            );
            var target = isNonMoph ? summary.nonMoph : summary.moph;

            target.approved += parseNum(row[I_APPROVE]);
            target.allocated += parseNum(row[I_ALLOC]);
            target.spent += parseNum(row[I_SPENT]);
            target.balance += parseNum(row[I_BAL]);

            var dept = String(row[I_DEPT] || 'ไม่ระบุ').trim();
            if (dept === '') dept = 'ไม่ระบุ';
            if (!target.deptStats[dept]) target.deptStats[dept] = { allocated: 0, spent: 0 };
            target.deptStats[dept].allocated += parseNum(row[I_ALLOC]);
            target.deptStats[dept].spent += parseNum(row[I_SPENT]);
        });
        return summary;
    } catch (e) { return { error: e.message }; }
}

// ==========================================
// 5. SEARCH & YEARLY
// ==========================================
function searchActionPlan(dept, budgetType, quarter, month) {
    var result = getYearlyPlanData(dept, budgetType, quarter, month);
    return { summary: { count: result.summary.projects, approved: result.summary.approved, allocated: result.summary.allocated }, list: result.list };
}

function getYearlyPlanData(deptFilter, typeFilter, quarterFilter, monthFilter) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return { summary: { projects: 0 }, list: [] };
        var data = sheet.getDataRange().getValues();
        data.shift();

        var I_ORDER = 3, I_DEPT = 4, I_PROJ = 6, I_ACT = 7, I_SUB = 8, I_TYPE = 9, I_SOURCE = 10, I_ALLOC = 16, I_SPENT = 17;
        var I_MONTHS = [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37];
        var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
        var summary = { projects: 0, approved: 0, allocated: 0, spent: 0 };
        var list = [];
        var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        
        data.forEach(row => {
            var rowDept = String(row[I_DEPT]).trim();
            var passDept = (deptFilter === "" || deptFilter === null || rowDept === deptFilter);
            var typeVal = String(row[I_TYPE] || "").trim();
            var isNonMoph = (
                typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || 
                typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1
            );
            var isMoph = !isNonMoph;
            var passType = true;
            if (typeFilter === 'MOPH') passType = isMoph;
            else if (typeFilter === 'NONMOPH') passType = isNonMoph;

            var timeline = I_MONTHS.map(idx => (String(row[idx]).trim() !== '') ? 1 : 0);
            var passTime = true;
            if (quarterFilter && quarters[quarterFilter]) {
                if (!quarters[quarterFilter].some(mIdx => timeline[mIdx] === 1)) passTime = false;
            }
            if (monthFilter) {
                if (timeline[parseInt(monthFilter)] !== 1) passTime = false;
            }

            if (passDept && passType && passTime) {
                var actName = String(row[I_ACT]);
                if (row[I_SUB]) actName += " (" + row[I_SUB] + ")";
                var alloc = parseNum(row[I_ALLOC]);
                var spent = parseNum(row[I_SPENT]);
                summary.projects++;
                summary.allocated += alloc; summary.spent += spent;
                list.push({
                    order: row[I_ORDER], dept: rowDept, project: row[I_PROJ], activity: actName,
                    type: row[I_TYPE], budgetSource: row[I_SOURCE],
                    timeline: timeline, allocated: alloc, spent: spent, balance: alloc - spent
                });
            }
        });
        return { summary: summary, list: list };
    } catch (e) { return { error: e.message }; }
}

// ==========================================
// 6. SAVE & UPDATE (Transaction)
// ==========================================
function saveTransaction(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var mSheet = ss.getSheetByName(SHEET_NAME);
        var tSheet = ss.getSheetByName('t_actionplan');
        var mData = mSheet.getDataRange().getValues();
        var idxID = 0; var idxSpent = 17; var rowIndex = -1;
        for (var i = 1; i < mData.length; i++) { if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i + 1; break; } }
        if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการใน Master' };
        
        var currentSpent = (parseFloat(String(mSheet.getRange(rowIndex, idxSpent + 1).getValue()).replace(/,/g, '')) || 0) + parseFloat(form.amount);
        var allocated = parseFloat(String(mSheet.getRange(rowIndex, 17).getValue()).replace(/,/g, '')) || 0; 
        var balanceAfterTx = allocated - currentSpent;
        mSheet.getRange(rowIndex, idxSpent + 1).setValue(currentSpent);
        
        var r = mData[rowIndex - 1];
        tSheet.appendRow([
            new Date(), r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[13], r[14], r[16],
            form.amount, 0, form.txDate, form.expenseType, form.desc, form.remark, "", "", "", balanceAfterTx 
        ]);
        return { status: 'success', message: 'บันทึกเรียบร้อย' };
    } catch (e) { return { status: 'error', message: e.message };
    } finally { lock.releaseLock(); }
}

function deleteTransaction(rowId, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var txSheet = ss.getSheetByName('t_actionplan');
        var mSheet = ss.getSheetByName('m_actionplan');
        var mData = mSheet.getDataRange().getValues();
        var mRow = -1;
        var targetId = String(projectId).trim();
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][0]).trim() === targetId) { mRow = i + 1; break; }
        }
        if (mRow !== -1) {
            var cellSpent = mSheet.getRange(mRow, 18);
            var cellAlloc = mSheet.getRange(mRow, 17);
            // var cellBalance = mSheet.getRange(mRow, 20);
            var curSpent = parseFloat(String(cellSpent.getValue()).replace(/,/g, '')) || 0;
            var allocated = parseFloat(String(cellAlloc.getValue()).replace(/,/g, '')) || 0;
            var amtToReturn = parseFloat(amount) || 0;
            var newSpent = curSpent - amtToReturn;
            if (newSpent < 0) newSpent = 0; 
            var newBalance = allocated - newSpent;
            cellSpent.setValue(newSpent);
            cellBalance.setValue(newBalance);
        }
        txSheet.deleteRow(parseInt(rowId));
        return { status: 'success', message: 'ลบรายการและคืนเงินเรียบร้อย' };
    } catch (e) { return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

function editTransaction(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_actionplan');
        var mSheet = ss.getSheetByName('m_actionplan');
        if (!tSheet || !mSheet) return { status: 'error', message: 'ไม่พบ Sheet ข้อมูล' };
        var tRow = parseInt(form.rowId);
        var checkID = tSheet.getRange(tRow, 2).getValue();
        if (String(checkID) != String(form.projectId)) { return { status: 'error', message: 'ข้อมูลไม่ตรงกัน โปรดรีเฟรช' }; }
        
        var mData = mSheet.getDataRange().getValues();
        var mRowIndex = -1;
        var targetId = String(form.projectId).trim();
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][0]).trim() === targetId) { mRowIndex = i + 1; break; }
        }
        var newBalance = 0;
        if (mRowIndex !== -1) {
            var cellSpent = mSheet.getRange(mRowIndex, 18);
            var cellAlloc = mSheet.getRange(mRowIndex, 17);
            // var cellBalance = mSheet.getRange(mRowIndex, 20);
            var currentSpent = parseFloat(String(cellSpent.getValue()).replace(/,/g, '')) || 0;
            var allocated = parseFloat(String(cellAlloc.getValue()).replace(/,/g, '')) || 0;
            var oldVal = parseFloat(form.oldAmount) || 0;
            var newVal = parseFloat(form.newAmount) || 0;
            var newSpentTotal = currentSpent - oldVal + newVal;
            newBalance = allocated - newSpentTotal;
            cellSpent.setValue(newSpentTotal);
            cellBalance.setValue(newBalance);
        }
        tSheet.getRange(tRow, 16).setValue(form.newAmount);
        tSheet.getRange(tRow, 18).setValue(form.date);
        tSheet.getRange(tRow, 20).setValue(form.desc);
        tSheet.getRange(tRow, 24).setValue(form.reason);
        tSheet.getRange(tRow, 25).setValue(newBalance);
        return { status: 'success', message: 'แก้ไขเรียบร้อย' };
    } catch (e) { return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 7. LOAN FUNCTIONS (Save & Repay)
// ==========================================
function saveLoan(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var mSheet = ss.getSheetByName('m_actionplan');
        var tSheet = ss.getSheetByName('t_loan');
        if (!mSheet) return { status: 'error', message: 'ไม่พบ Sheet: m_actionplan' };
        if (!tSheet) return { status: 'error', message: 'ไม่พบ Sheet: t_loan (กรุณาสร้าง Sheet นี้)' };
        var mData = mSheet.getDataRange().getValues();
        var idxID = 0; var idxLoan = 18; var rowIndex = -1;
        var targetId = String(form.projectId).trim();
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][idxID]).trim() === targetId) { rowIndex = i + 1; break; }
        }
        if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ใน Master (ID: ' + form.projectId + ')' };
        
        var r = mData[rowIndex - 1];
        var cellMasterLoan = mSheet.getRange(rowIndex, idxLoan + 1);
        var currentMasterLoan = parseFloat(String(cellMasterLoan.getValue()).replace(/,/g, '')) || 0;
        var loanAmt = parseFloat(String(form.amount).replace(/,/g, '')) || 0;
        var newMasterLoan = currentMasterLoan + loanAmt;
        var sDate = form.startDate ? new Date(form.startDate) : "";
        var eDate = form.endDate ? new Date(form.endDate) : "";
        
        cellMasterLoan.setValue(newMasterLoan);
        SpreadsheetApp.flush(); 
        
        tSheet.appendRow([
            new Date(), r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[13], r[14], r[16],
            loanAmt, form.loanDate, form.loanType, form.desc, "ยังไม่ดำเนินการ", 0, loanAmt, "", "", "", sDate, eDate 
        ]);
        return { status: 'success', message: 'บันทึกเงินยืมเรียบร้อย' };
    } catch (e) { return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

function updateLoanRepayment(data) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_loan');
        if (!tSheet) throw new Error("ไม่พบ Sheet: t_loan");
        var tData = tSheet.getDataRange().getValues();
        var tDisplayData = tSheet.getDataRange().getDisplayValues();
        var tRowIndex = -1; var projectId = ""; var loanAmount = 0;
        var targetDate = new Date(data.timestamp);
        
        for (var i = 1; i < tData.length; i++) {
            var cellValue = tData[i][0];
            var cellDisplayValue = tDisplayData[i][0]; 
            var isMatch = false;
            if (String(cellDisplayValue).trim() === String(data.timestamp).trim() || String(cellValue).trim() === String(data.timestamp).trim()) {
                isMatch = true;
            } else {
                var cellDate = new Date(cellValue);
                if (!isNaN(cellDate.getTime()) && !isNaN(targetDate.getTime())) {
                    if (Math.abs(cellDate.getTime() - targetDate.getTime()) < 60000) { isMatch = true; }
                }
            }
            if (isMatch) { tRowIndex = i + 1; projectId = tData[i][1]; loanAmount = parseFloat(tData[i][15] || 0); break; }
        }
        if (tRowIndex === -1) throw new Error("ไม่พบรายการกู้ยืม (Timestamp ไม่ตรง)");
        
        var currentPaid = parseFloat(tData[tRowIndex - 1][20] || 0);
        var payAmount = parseFloat(data.paidAmount);
        var newPaid = currentPaid + payAmount;
        var newBalance = loanAmount - newPaid;
        var status = (newBalance <= 0.01) ? "คืนครบ" : "คืนบางส่วน";
        if (newBalance < 0) newBalance = 0;
        
        tSheet.getRange(tRowIndex, 20).setValue(status);
        tSheet.getRange(tRowIndex, 21).setValue(newPaid);
        tSheet.getRange(tRowIndex, 22).setValue(newBalance);
        tSheet.getRange(tRowIndex, 23).setValue(data.payDate);
        
        if (projectId) {
            var mSheet = ss.getSheetByName('m_actionplan');
            if (mSheet) {
                var mData = mSheet.getDataRange().getValues();
                var mRowIndex = -1;
                for (var j = 1; j < mData.length; j++) {
                    if (String(mData[j][0]) == String(projectId)) { mRowIndex = j + 1; break; }
                }
                if (mRowIndex !== -1) {
                    var colAlloc = 17; var colSpent = 18; var colBalance = 20;
                    var cellAlloc = mSheet.getRange(mRowIndex, colAlloc);
                    var allocated = parseFloat(cellAlloc.getValue()) || 0;
                    var cellSpent = mSheet.getRange(mRowIndex, colSpent);
                    var currentSpent = parseFloat(cellSpent.getValue()) || 0;
                    var updatedSpent = currentSpent + payAmount;
                    // var updatedBalance = allocated - updatedSpent;
                    // cellSpent.setValue(updatedSpent);
                    mSheet.getRange(mRowIndex, colBalance).setValue(updatedBalance);
                }
            }
        }
        return { status: 'success', message: 'บันทึกคืนเงินและตัดงบประมาณเรียบร้อย' };
    } catch (e) { return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.toString() };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 8. HISTORY GETTERS
// ==========================================
function getTransactionHistory() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_actionplan');
        if (!tSheet) return [];
        var mSheet = ss.getSheetByName('m_actionplan');
        var masterMap = {};
        if (mSheet) {
            var mData = mSheet.getDataRange().getValues();
            for (var i = 1; i < mData.length; i++) {
                var pid = String(mData[i][0]).trim();
                if (pid) {
                    masterMap[pid] = {
                        allocated: parseFloat(String(mData[i][16]).replace(/,/g, '')) || 0,
                        spent: parseFloat(String(mData[i][17]).replace(/,/g, '')) || 0,    
                        balance: parseFloat(String(mData[i][19]).replace(/,/g, '')) || 0    
                    };
                }
            }
        }
        var data = tSheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseAmount = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };
        var toThaiDate = function (val) {
            if (!val) return "-";
            try {
                var d = new Date(val);
                if (isNaN(d.getTime())) return String(val);
                var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
                return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
            } catch (ex) { return String(val); }
        };
        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[1])) continue;
            var projId = String(row[1]).trim();
            var masterInfo = masterMap[projId] || { allocated: 0, spent: 0, balance: 0 };
            var item = {
                rowId: i + 1, order: row[4], project: row[7], activity: row[8], subActivity: row[9],
                amount: parseAmount(row[15]), date: toThaiDate(row[17]), type: row[18], source: row[11], desc: row[19], id: row[1],
                masterAllocated: masterInfo.allocated, masterSpent: masterInfo.spent, masterBalance: masterInfo.balance
            };
            if (item.amount > 0 || item.order) result.push(item);
            if (result.length >= 100) break;
        }
        return result;
    } catch (e) { return []; }
}

function getLoanHistory() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    var projectMap = {};
    try {
        var mSheet = ss.getSheetByName('m_actionplan');
        if (mSheet) {
            var mData = mSheet.getDataRange().getDisplayValues();
            for (var i = 1; i < mData.length; i++) {
                var pid = String(mData[i][0]).trim();
                if (pid) { projectMap[pid] = { type: mData[i][9] || "-", source: mData[i][10] || "-" }; }
            }
        }
    } catch (e) {}
    var tData = tSheet.getDataRange().getDisplayValues();
    var result = [];
    var parseNum = function (val) { return parseFloat(String(val).replace(/,/g, '')) || 0; };
    var toThaiDate = function (val) {
        if (!val) return "-";
        try {
            var d;
            if (Object.prototype.toString.call(val) === '[object Date]') d = val;
            else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
                var parts = val.split('-');
                d = new Date(parts[0], parts[1] - 1, parts[2]);
            } else { d = new Date(val); }
            if (isNaN(d.getTime())) return String(val);
            var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
            return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
        } catch (ex) { return String(val); }
    };
    for (var i = tData.length - 1; i >= 1; i--) {
        try {
            var row = tData[i];
            if (!row[0] && !row[1]) continue;
            var pid = String(row[1] || "").trim();
            var meta = projectMap[pid] || { type: '-', source: '-' };
            result.push({
                id: row[0], timestamp: row[0], project: row[7], activity: row[8], subActivity: row[9],
                amount: parseNum(row[15]), date: toThaiDate(row[16]), status: row[19], paid: parseNum(row[20]), balance: parseNum(row[21]),
                order: row[4], type: row[17], desc: row[18], budgetType: meta.type, budgetSource: meta.source, dept: row[2]
            });
        } catch (e) {}
    }
    return result;
}

function getHistory(sheetName) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName(sheetName);
        if (!tSheet) return [];
        var data = tSheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseAmount = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };
        var toThaiDate = function (val) {
            if (!val) return "-";
            try {
                var d;
                if (Object.prototype.toString.call(val) === '[object Date]') d = val;
                else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
                    var parts = val.split('-');
                    d = new Date(parts[0], parts[1] - 1, parts[2]);
                } else { d = new Date(val); }
                if (isNaN(d.getTime())) return String(val);
                var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
                return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
            } catch (ex) { return String(val); }
        };
        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[1])) continue;
            var item = {};
            if (sheetName === 't_actionplan') {
                item = {
                    rowId: i + 1, order: row[4], project: row[7], activity: row[8], subActivity: row[9], amount: parseAmount(row[15]),
                    date: toThaiDate(row[17]), type: row[18], source: row[11], desc: row[19], id: row[1]
                };
            } else {
                item = { timestamp: row[0], order: row[4], project: row[7], amount: parseAmount(row[15]), date: toThaiDate(row[16]), status: row[19] };
            }
            if (item.amount > 0 || item.order) result.push(item);
            if (result.length >= 50) break;
        }
        return result;
    } catch (e) { return []; }
}

function searchLoanHistory(criteria) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    var mSheet = ss.getSheetByName('m_actionplan');
    var projectMap = {};
    if (mSheet) {
        var mData = mSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < mData.length; i++) {
            var pid = String(mData[i][0]).trim();
            projectMap[pid] = { type: mData[i][9], source: mData[i][10] };
        }
    }
    var data = tSheet.getDataRange().getDisplayValues();
    var result = [];
    var parseNum = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var match = true;
        if (criteria.order && String(row[4]) != String(criteria.order)) match = false;
        if (match && criteria.project && String(row[7]).indexOf(criteria.project) === -1) match = false;
        if (match) {
            var pid = String(row[1]).trim();
            var meta = projectMap[pid] || { type: '-', source: '-' };
            result.push({
                id: row[0], timestamp: row[0], project: row[7], activity: row[8], subActivity: row[9], amount: parseNum(row[15]),
                date: row[16], status: row[19], paid: parseNum(row[20]), balance: parseNum(row[21]), order: row[4],
                type: row[17], desc: row[18], budgetType: meta.type, budgetSource: meta.source, dept: row[2]
            });
        }
    }
    return result;
}

// ==========================================
// 9. NEW ALLOCATION SYSTEM (Backend)
// ==========================================
function saveAllocation(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var mSheet = ss.getSheetByName('m_actionplan');
        var tAllocSheet = ss.getSheetByName('t_allocate');
        if (!mSheet) return { status: 'error', message: 'ไม่พบ Sheet Master' };
        if (!tAllocSheet) return { status: 'error', message: 'ไม่พบ Sheet t_allocate' };
        
        var mData = mSheet.getDataRange().getValues();
        var rowIndex = -1;
        var targetId = String(form.id).trim();
        var targetAct = form.fullData ? String(form.fullData.activity || "").trim() : "";
        var targetSub = form.fullData ? String(form.fullData.subActivity || "").trim() : "";
        
        for (var i = 1; i < mData.length; i++) {
            var rowId = String(mData[i][0]).trim();
            var rowAct = String(mData[i][7]).trim();
            var rowSub = String(mData[i][8]).trim();
            var isIdMatch = (rowId === targetId);
            var isActMatch = (rowAct === targetAct);
            var isSubMatch = (rowSub === targetSub);
            if (isIdMatch && isActMatch && isSubMatch) { rowIndex = i + 1; break; }
        }
        if (rowIndex === -1) {
            for (var i = 1; i < mData.length; i++) {
                if (String(mData[i][0]).trim() === targetId) { rowIndex = i + 1; break; }
            }
            if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ใน Master (ID: ' + form.id + ')' };
        }
        
        var cellAlloc = mSheet.getRange(rowIndex, 17);
        var rawVal = String(cellAlloc.getValue());
        var currentTotal = parseFloat(rawVal.replace(/,/g, '')) || 0;
        var addAmount = parseFloat(String(form.currentAlloc).replace(/,/g, '')) || 0;
        var newTotal = currentTotal + addAmount;
        cellAlloc.setValue(newTotal);
        SpreadsheetApp.flush(); 
        
        var r = mData[rowIndex - 1];
        var logRow = [
            new Date(), r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[13], r[14],
            newTotal, addAmount, form.date, form.letterNo, form.remark       
        ];
        tAllocSheet.appendRow(logRow);
        return { status: 'success', message: 'บันทึกจัดสรรสำเร็จ (ยอดใหม่: ' + newTotal.toLocaleString() + ')' };
    } catch (e) { return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

function getAllocationHistory() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('t_allocate');
        if (!sheet) return [];
        var data = sheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseNum = (v) => parseFloat(String(v).replace(/,/g, '')) || 0;
        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row[1]) continue; 
            result.push({
                id: row[1], order: row[4], project: row[7], activity: row[8], subActivity: row[9], type: row[10], source: row[11],     
                accumulatedAlloc: parseNum(row[14]), currentAlloc: parseNum(row[15]), date: formatDateThai(row[16]), letterNo: row[17]    
            });
            if (result.length >= 100) break; 
        }
        return result;
    } catch (e) { return []; }
}

function formatDateThai(dateStr) {
    if (!dateStr) return "-";
    try {
        var d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
    } catch (e) { return dateStr; }
}

// ==========================================
// 10. SUMMARY REPORT & DRILL DOWN
// ==========================================
function getSummaryData(quarterFilter, monthFilter) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return { error: "ไม่พบ Sheet ข้อมูล" };

        var data = sheet.getDataRange().getValues();
        var I_DEPT = 4; var I_TYPE = 9; var I_SOURCE = 10; var I_ALLOC = 16; var I_SPENT = 17;
        var I_MONTHS = [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37];
        var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
        var parseNum = function (val) { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        var grandTotal = { allocated: 0, spent: 0, count: 0 };
        var bySource = {};
        var byDeptAll = {}, byDeptMoph = {}, byDeptNon = {};
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var timeline = I_MONTHS.map(function (idx) { return (String(row[idx] || "").trim() !== '') ? 1 : 0; });
            var passTime = true;
            if (quarterFilter && quarters[quarterFilter]) { if (!quarters[quarterFilter].some(function (mIdx) { return timeline[mIdx] === 1; })) passTime = false; }
            if (monthFilter && String(monthFilter) !== "") { if (timeline[parseInt(monthFilter)] !== 1) passTime = false; }

            if (passTime) {
                var alloc = parseNum(row[I_ALLOC]);
                var spent = parseNum(row[I_SPENT]);
                var typeVal = String(row[I_TYPE] || "").trim();
                var deptVal = String(row[I_DEPT] || "ไม่ระบุ").trim();
                if (deptVal === "") deptVal = "ไม่ระบุ";

                var isNonMoph = (
                    typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || 
                    typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1
                );
                var sourceGroup = isNonMoph ? "เงินนอกงบประมาณ (Non-MOPH)" : "งบประมาณ (สป.สธ.)";

                grandTotal.allocated += alloc; grandTotal.spent += spent; grandTotal.count++;
                if (!bySource[sourceGroup]) bySource[sourceGroup] = { allocated: 0, spent: 0, count: 0 };
                bySource[sourceGroup].allocated += alloc; bySource[sourceGroup].spent += spent; bySource[sourceGroup].count++;
                if (!byDeptAll[deptVal]) byDeptAll[deptVal] = { allocated: 0, spent: 0, count: 0 };
                byDeptAll[deptVal].allocated += alloc; byDeptAll[deptVal].spent += spent; byDeptAll[deptVal].count++;
                if (!isNonMoph) { 
                    if (!byDeptMoph[deptVal]) byDeptMoph[deptVal] = { allocated: 0, spent: 0, count: 0 };
                    byDeptMoph[deptVal].allocated += alloc; byDeptMoph[deptVal].spent += spent; byDeptMoph[deptVal].count++;
                } else { 
                    if (!byDeptNon[deptVal]) byDeptNon[deptVal] = { allocated: 0, spent: 0, count: 0 };
                    byDeptNon[deptVal].allocated += alloc; byDeptNon[deptVal].spent += spent; byDeptNon[deptVal].count++;
                }
            }
        }

        var toList = function (obj) {
            var list = [];
            for (var k in obj) { list.push({ name: k, allocated: obj[k].allocated, spent: obj[k].spent, count: obj[k].count }); }
            return list.sort(function (a, b) { return b.allocated - a.allocated; });
        };

        return { grandTotal: grandTotal, bySource: toList(bySource), byDeptAll: toList(byDeptAll), byDeptMoph: toList(byDeptMoph), byDeptNon: toList(byDeptNon) };
    } catch (e) { return { error: e.message }; }
}

function getDeptDetails(deptName, quarterFilter, monthFilter) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return { error: "ไม่พบ Sheet m_actionplan" };
        var data = sheet.getDataRange().getValues();
        var I_DEPT = 4; var I_PROJ = 6; var I_ACT = 7; var I_TYPE = 9; var I_SOURCE = 10;
        var I_APPROVE = 15; var I_ALLOC = 16; var I_SPENT = 17;
        var I_MONTHS = [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37];
        var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
        var parseNum = function (val) { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        var cleanName = function (str) { return String(str).replace(/[\s\/\-_]+/g, "").trim(); };

        var projectsAll = [], projectsMoph = [], projectsNon = [];
        var sumAll = { allocated: 0, spent: 0 }, sumMoph = { allocated: 0, spent: 0 }, sumNon = { allocated: 0, spent: 0 };
        var targetClean = cleanName(deptName);

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var rowDeptRaw = String(row[I_DEPT] || "");
            if (cleanName(rowDeptRaw).indexOf(targetClean) === -1 && targetClean.indexOf(cleanName(rowDeptRaw)) === -1) continue;
            var timeline = I_MONTHS.map(function (idx) { return (String(row[idx] || "").trim() !== '') ? 1 : 0; });
            var passTime = true;
            if (quarterFilter && quarters[quarterFilter]) { if (!quarters[quarterFilter].some(function (mIdx) { return timeline[mIdx] === 1; })) passTime = false; }
            if (monthFilter && String(monthFilter) !== "") { if (timeline[parseInt(monthFilter)] !== 1) passTime = false; }

            if (passTime) {
                var approve = parseNum(row[I_APPROVE]);
                var alloc = parseNum(row[I_ALLOC]);
                var spent = parseNum(row[I_SPENT]);
                var typeVal = String(row[I_TYPE] || "").trim();
                var projObj = {
                    project: String(row[I_PROJ] || "-"),
                    activity: String(row[I_ACT] || "-"),
                    approved: approve, allocated: alloc, spent: spent, balance: alloc - spent,
                    progress: (alloc > 0) ? (spent / alloc * 100) : 0
                };
                projectsAll.push(projObj);
                sumAll.allocated += alloc; sumAll.spent += spent;

                var isNonMoph = (
                    typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 ||
                    typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1
                );
                if (!isNonMoph) { 
                    projectsMoph.push(projObj);
                    sumMoph.allocated += alloc; sumMoph.spent += spent;
                } else { 
                    projectsNon.push(projObj);
                    sumNon.allocated += alloc; sumNon.spent += spent;
                }
            }
        }

        var sortFn = function (a, b) { return b.progress - a.progress; };
        projectsAll.sort(sortFn); projectsMoph.sort(sortFn); projectsNon.sort(sortFn);

        return {
            projectsAll: projectsAll, projectsMoph: projectsMoph, projectsNon: projectsNon,
            sumAll: sumAll, sumMoph: sumMoph, sumNon: sumNon,
            deptName: deptName
        };
    } catch (e) { return { error: "Server Error: " + e.message }; }
}

// ==========================================
// 📌 ฟังก์ชันทดสอบความถูกต้องของชื่อคอลัมน์ (รันเพื่อเช็ค)
// ==========================================
function testDynamicMapping() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    var map = getColumnMap(sheet);
    var keysToTest = Object.keys(COL_NAME);
    
    var errorCount = 0;
    Logger.log("=== 🔍 เริ่มตรวจสอบ Dynamic Mapping ===");
    
    for (var i = 0; i < keysToTest.length; i++) {
        var key = keysToTest[i];
        var expectedName = COL_NAME[key];
        var foundIndex = map[expectedName];
        
        if (foundIndex !== undefined) {
            Logger.log("✅ พบคอลัมน์ [" + expectedName + "] อยู่ที่ Index: " + foundIndex + " (คอลัมน์ที่ " + (foundIndex + 1) + ")");
        } else {
            Logger.log("❌ ไม่พบ! หัวคอลัมน์ [" + expectedName + "] โปรดเช็คตัวสะกด ช่องว่าง หรือการเว้นวรรคในชีตครับ");
            errorCount++;
        }
    }
    
    if (errorCount === 0) {
        Logger.log("\n🎉 ยอดเยี่ยม! ระบบค้นหาคอลัมน์เจอครบ 100% โครงสร้าง Dynamic พร้อมใช้งานแล้ว!");
    } else {
        Logger.log("\n⚠️ พบ " + errorCount + " จุดที่หาไม่เจอ โปรดแก้ไขให้ตรงกันก่อนไปต่อครับ");
    }
}

// ==========================================
// 📌 MASTER PLAN CRUD SYSTEM (Dynamic Mapping)
// ==========================================

// 1. ดึงข้อมูลโครงการทั้งหมดไปแสดงที่ตารางหน้าเว็บ (Admin Panel)
function getAdminMasterPlan() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return { error: "ไม่พบชีต m_actionplan" };
        
        var data = sheet.getDataRange().getValues();
        var list = [];
        
        // เริ่มลูปที่ 1 เพื่อข้ามหัวคอลัมน์ (แถวที่ 1)
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            
            // Col A (Index 0) คือ รหัสโครงการ
            var id = String(row[0] || "").trim();
            if (!id) continue; // ข้ามบรรทัดว่าง
            
            // 🌟 ล็อคเป้าคอลัมน์ตายตัว (A=0, B=1, C=2...) จบปัญหาดึงข้อมูลมั่ว 100%
            list.push({
                id: id,                                    // Col A (0): รหัสโครงการ
                order: String(row[3] || '').trim(),        // Col D (3): ลำดับโครงการ
                dept: String(row[4] || '').trim(),         // Col E (4): กลุ่มงาน/งาน
                proj: String(row[6] || '').trim(),         // Col G (6): โครงการ
                act: String(row[7] || '').trim(),          // Col H (7): กิจกรรมหลัก
                sub: String(row[8] || '').trim(),          // Col I (8): กิจกรรมย่อย
                type: String(row[9] || '').trim(),         // Col J (9): ประเภทงบ
                source: String(row[10] || '').trim(),      // Col K (10): แหล่งงบประมาณ
                
                // 💰 ข้อมูลการเงิน (ลบลูกน้ำ และแปลงเป็นตัวเลข)
                alloc: parseFloat(String(row[15] || '0').replace(/,/g, '')) || 0,   // Col P (15): จัดสรร
                spent: parseFloat(String(row[16] || '0').replace(/,/g, '')) || 0,   // Col Q (16): เบิกจ่าย
                loan: parseFloat(String(row[17] || '0').replace(/,/g, '')) || 0,    // Col R (17): เงินยืม
                balance: parseFloat(String(row[18] || '0').replace(/,/g, '')) || 0, // Col S (18): คงเหลือ
                
                // 📌 สถานะ (Col AM คือคอลัมน์ที่ 39 = Index 38)
                status: String(row[38] || 'ACTIVE').trim().toUpperCase()
            });
        }
        return list;
    } catch (e) {
        return { error: e.message };
    }
}

// 2. คำนวณปีงบประมาณ และ สร้างรหัสโครงการอัตโนมัติ (เช่น P-2569-1074)
function generateNextProjectID() {
    var d = new Date();
    var year = d.getFullYear() + 543; // ค.ศ. -> พ.ศ.
    var month = d.getMonth() + 1;     // 1-12
    if (month >= 10) year += 1;       // ต.ค. เป็นต้นไป ปัดเป็นปีงบถัดไป
    
    var fy = year.toString();
    var prefix = "P-" + fy + "-";
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    var map = getColumnMap(sheet);
    var data = sheet.getDataRange().getValues();
    
    var maxNum = 0;
    for (var i = 1; i < data.length; i++) {
        var currentId = String(data[i][map[COL_NAME.ID]]).trim();
        if (currentId.indexOf(prefix) === 0) {
            var numStr = currentId.replace(prefix, "");
            var num = parseInt(numStr, 10);
            if (!isNaN(num) && num > maxNum) {
                maxNum = num;
            }
        }
    }
    
    var nextNumStr = (maxNum + 1).toString();
    while (nextNumStr.length < 4) nextNumStr = "0" + nextNumStr; // เติม 0 ให้ครบ 4 หลัก
    
    return prefix + nextNumStr;
}

// 3. บันทึกข้อมูล (เพิ่มโครงการใหม่ หรือ อัปเดตโครงการเดิม)
// 3. บันทึกข้อมูล (เพิ่มโครงการใหม่ หรือ อัปเดตโครงการเดิม)
function saveMasterPlan(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); // ป้องกันคนกดเซฟชนกัน
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        var map = getColumnMap(sheet);
        
        // 📌 ระบบจัดการเดือน: แปลงข้อความ "ต.ค., พ.ย." ให้เป็น Array
        var selectedMonths = [];
        if (payload.months) {
            selectedMonths = payload.months.split(',').map(function(m) { return m.trim(); });
        }
        var checkMonth = function(monthName) { return selectedMonths.indexOf(monthName) !== -1 ? 1 : ""; };

        if (payload.id) { // กรณีแก้ไข (UPDATE)
            var data = sheet.getDataRange().getValues();
            var rowIndex = -1;
            for (var i = 1; i < data.length; i++) {
                if (String(data[i][map[COL_NAME.ID]]) === String(payload.id)) { rowIndex = i + 1; break; }
            }
            if (rowIndex > -1) {
                sheet.getRange(rowIndex, map[COL_NAME.ORDER] + 1).setValue(payload.order);
                sheet.getRange(rowIndex, map[COL_NAME.DEPT] + 1).setValue(payload.dept);
                sheet.getRange(rowIndex, map[COL_NAME.PLAN] + 1).setValue(payload.plan); 
                sheet.getRange(rowIndex, map[COL_NAME.PROJ] + 1).setValue(payload.proj);
                sheet.getRange(rowIndex, map[COL_NAME.ACT] + 1).setValue(payload.act);
                sheet.getRange(rowIndex, map[COL_NAME.SUB] + 1).setValue(payload.sub);
                sheet.getRange(rowIndex, map[COL_NAME.TYPE] + 1).setValue(payload.type);
                sheet.getRange(rowIndex, map[COL_NAME.SOURCE] + 1).setValue(payload.source);
                sheet.getRange(rowIndex, map[COL_NAME.APPROVE] + 1).setValue(payload.approve); 
                sheet.getRange(rowIndex, map[COL_NAME.ALLOC] + 1).setValue(payload.alloc);
                sheet.getRange(rowIndex, map[COL_NAME.STATUS] + 1).setValue(payload.status);
                sheet.getRange(rowIndex, map[COL_NAME.REMARK] + 1).setValue(payload.remark);

                // อัปเดต 12 เดือน
                sheet.getRange(rowIndex, map[COL_NAME.M_OCT] + 1).setValue(checkMonth('ต.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_NOV] + 1).setValue(checkMonth('พ.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_DEC] + 1).setValue(checkMonth('ธ.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JAN] + 1).setValue(checkMonth('ม.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_FEB] + 1).setValue(checkMonth('ก.พ.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_MAR] + 1).setValue(checkMonth('มี.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_APR] + 1).setValue(checkMonth('เม.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_MAY] + 1).setValue(checkMonth('พ.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JUN] + 1).setValue(checkMonth('มิ.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JUL] + 1).setValue(checkMonth('ก.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_AUG] + 1).setValue(checkMonth('ส.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_SEP] + 1).setValue(checkMonth('ก.ย.'));

                return { status: 'success', message: "อัปเดตข้อมูลสำเร็จ" };
            }
        } else { // กรณีเพิ่มใหม่ (INSERT)
            var newRow = new Array(sheet.getLastColumn()).fill(""); 
            var newId = generateNextProjectID(); 
            
            newRow[map[COL_NAME.ID]] = newId;
            newRow[map[COL_NAME.ORDER]] = payload.order;
            newRow[map[COL_NAME.DEPT]] = payload.dept;
            newRow[map[COL_NAME.PLAN]] = payload.plan;
            newRow[map[COL_NAME.PROJ]] = payload.proj;
            newRow[map[COL_NAME.ACT]] = payload.act;
            newRow[map[COL_NAME.SUB]] = payload.sub;
            newRow[map[COL_NAME.TYPE]] = payload.type;
            newRow[map[COL_NAME.SOURCE]] = payload.source;
            newRow[map[COL_NAME.APPROVE]] = payload.approve;
            newRow[map[COL_NAME.ALLOC]] = payload.alloc;
            newRow[map[COL_NAME.STATUS]] = payload.status;
            newRow[map[COL_NAME.REMARK]] = payload.remark;
            
            newRow[map[COL_NAME.M_OCT]] = checkMonth('ต.ค.'); newRow[map[COL_NAME.M_NOV]] = checkMonth('พ.ย.');
            newRow[map[COL_NAME.M_DEC]] = checkMonth('ธ.ค.'); newRow[map[COL_NAME.M_JAN]] = checkMonth('ม.ค.');
            newRow[map[COL_NAME.M_FEB]] = checkMonth('ก.พ.'); newRow[map[COL_NAME.M_MAR]] = checkMonth('มี.ค.');
            newRow[map[COL_NAME.M_APR]] = checkMonth('เม.ย.'); newRow[map[COL_NAME.M_MAY]] = checkMonth('พ.ค.');
            newRow[map[COL_NAME.M_JUN]] = checkMonth('มิ.ย.'); newRow[map[COL_NAME.M_JUL]] = checkMonth('ก.ค.');
            newRow[map[COL_NAME.M_AUG]] = checkMonth('ส.ค.'); newRow[map[COL_NAME.M_SEP]] = checkMonth('ก.ย.');

            sheet.appendRow(newRow); 
            return { status: 'success', message: 'สร้างโครงการใหม่สำเร็จ รหัส: ' + newId, newId: newId };
        }
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// 15 March 2026
// ==========================================
// 📌 ดึงข้อมูลโครงการทั้งหมดไปแสดงที่หน้าจัดการแผนงาน
// ==========================================
function getAdminMasterPlan() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return { error: "ไม่พบชีต m_actionplan" };

        var map = getColumnMap(sheet); // 🌟 กลับมาใช้ Dynamic Column ท่าไม้ตาย!
        var data = sheet.getDataRange().getValues();
        var list = [];
        
        // เริ่มลูปที่ 1 เพื่อข้ามหัวคอลัมน์
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            
            // ใช้ map ค้นหา ID อัตโนมัติ
            var id = String(row[map[COL_NAME.ID]] || "").trim();
            if (!id) continue; // ข้ามบรรทัดว่าง
            
            // 🌟 แมปข้อมูลด้วยตัวแปรแบบ Dynamic 100%
            list.push({
                id: id,
                order: String(row[map[COL_NAME.ORDER]] || '').trim(),
                dept: String(row[map[COL_NAME.DEPT]] || '').trim(),
                proj: String(row[map[COL_NAME.PROJ]] || '').trim(),
                act: String(row[map[COL_NAME.ACT]] || '').trim(),
                sub: String(row[map[COL_NAME.SUB]] || '').trim(),
                type: String(row[map[COL_NAME.TYPE]] || '').trim(),
                source: String(row[map[COL_NAME.SOURCE]] || '').trim(),
                
                // 💰 ข้อมูลการเงิน
                alloc: parseFloat(String(row[map[COL_NAME.ALLOC]] || '0').replace(/,/g, '')) || 0,
                spent: parseFloat(String(row[map[COL_NAME.SPENT]] || '0').replace(/,/g, '')) || 0,
                loan: parseFloat(String(row[map[COL_NAME.LOAN]] || '0').replace(/,/g, '')) || 0,
                
                // ✅ แก้ไข: ใช้ COL_NAME.BAL ให้ตรงกับที่ประกาศไว้
                balance: parseFloat(String(row[map[COL_NAME.BAL]] || '0').replace(/,/g, '')) || 0,
                
                // 📌 สถานะ
                status: String(row[map[COL_NAME.STATUS]] || 'ACTIVE').trim().toUpperCase()
            });
        }
        return list;
    } catch (e) {
        return { error: e.message };
    }
}

// ==========================================
// 📌 เปลี่ยนสถานะทีละหลายรายการ (Bulk Update) & คืนเงินลงถัง
// ==========================================
function submitBulkUpdate(ids, newStatus, remark) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("m_actionplan");
    const logSheet = ss.getSheetByName("log_refunded_budget"); // ชีตถังรับเงิน
    
    if (!sheet) return { success: false, message: "หาชีต m_actionplan ไม่พบ" };
    
    const data = sheet.getDataRange().getValues();
    const timestamp = new Date();
    const logsToAppend = [];
    
    // วนลูปเช็คข้อมูลทีละแถว
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let rowId = row[1]; // รหัสโครงการคอลัมน์ B
      
      // ถ้า ID ของแถวนี้ ตรงกับที่แอดมินติ๊ก Checkbox มา
      if (ids.includes(rowId)) {
        
        // 1. อัปเดตสถานะลงชีต m_actionplan (สมมติสถานะอยู่คอลัมน์ T = 20, หมายเหตุคอลัมน์ U = 21)
        sheet.getRange(i + 1, 39).setValue(newStatus); // แก้ไขเลข 20 ให้ตรงกับคอลัมน์สถานะ
        if (remark) {
          sheet.getRange(i + 1, 40).setValue(remark); // แก้ไขเลข 21 ให้ตรงกับคอลัมน์หมายเหตุ
        }
        
        // 2. 🔥 Logic โยนเงินลงถัง (ทำเฉพาะเมื่อเลือก REFUNDED และเป็น เงินนอก สป.)
        let type = row[10]; // K: ประเภทงบ
        let source = row[11]; // L: แหล่งงบ
        
        if (newStatus === "REFUNDED" && (type === "เงินนอก สป." || type === "NONMOPH" || String(source).includes("เงินบำรุง"))) {
          let balance = row[18] === '' ? 0 : Number(row[18]); // S: ยอดคงเหลือล่าสุด
          
          // ถ้ามีเงินเหลือให้คืน
          if (balance > 0) {
            logsToAppend.push([
              timestamp, // A: วันที่
              rowId,     // B: รหัสโครงการ
              row[2],    // C: ปีงบประมาณ 
              row[3],    // D: หมวด
              row[4],    // E: ลำดับโครงการ
              row[5],    // F: กลุ่มงาน
              row[6],    // G: แผนงาน
              row[7],    // H: โครงการ
              row[8],    // I: กิจกรรมหลัก
              row[9],    // J: กิจกรรมย่อย
              row[10],   // K: ประเภทงบ
              row[11],   // L: แหล่งงบประมาณ
              row[12],   // M: รหัสงบประมาณ
              row[13],   // N: รหัสกิจกรรม
              row[14],   // O: อนุมัติตามแผน
              row[15] === '' ? 0 : Number(row[15]), // P: จัดสรร
              row[16] === '' ? 0 : Number(row[16]), // Q: เบิกจ่าย
              row[17] === '' ? 0 : Number(row[17]), // R: เงินยืม
              balance,   // S: ยอดเงินส่งคืน (คงเหลือ)
              remark || "ส่งคืนจากการทำ Bulk Action" // T: หมายเหตุ
            ]);
          }
        }
      }
    }
    
    // 3. นำข้อมูลเงินคืนไปเขียนลงถังรวดเดียว (Batch Insert)
    if (logsToAppend.length > 0 && logSheet) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logsToAppend[0].length).setValues(logsToAppend);
    }
    
    return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 📌 ฟังก์ชันดึงข้อมูลตัวเลือกสำหรับ Modal เพิ่มโครงการ
// ==========================================
function getModalDropdownData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID); // หรือ getActiveSpreadsheet()
        
        // ฟังก์ชันช่วยกวาดข้อมูลจากคอลัมน์ A (ข้ามหัวตารางแถว 1) แล้วตัดตัวซ้ำทิ้ง
        var fetchColumnData = function(sheetName) {
            var sheet = ss.getSheetByName(sheetName);
            if (!sheet) return [];
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return [];
            var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            return [...new Set(data.map(function(r) { return String(r[0]).trim(); }).filter(String))];
        };

        return {
            depts: fetchColumnData('c_deparment'),
            types: fetchColumnData('c_budget_type'),
            plans: fetchColumnData('c_plan')
        };
    } catch (e) {
        return { error: e.message };
    }
}

// ==========================================
// 📌 นำเข้าข้อมูล Master Plan แบบกลุ่ม (Batch Import)
// ==========================================
function importMasterPlanBatch(dataArray) {
  var lock = LockService.getScriptLock();
  try {
    // ล็อคสคริปต์ 15 วินาที ป้องกันคนอัปโหลดพร้อมกัน
    lock.waitLock(15000); 
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME); // ชีต m_actionplan
    var data = sheet.getDataRange().getValues();
    var headers = data[0]; // หัวตารางหลังบ้าน
    
    // 1. หาปีงบประมาณจากข้อมูลแรกเพื่อนำไปสร้าง รหัสโครงการ
    var importYear = String(dataArray[0]["ปีงบประมาณ"] || new Date().getFullYear() + 543).trim();
    var maxRunning = 0;
    
    // 2. วนลูปหา ID ล่าสุดของปีงบประมาณนั้น (เช่น P-2569-XXXX)
    var idColIndex = headers.indexOf(COL_NAME.ID);
    for (var i = 1; i < data.length; i++) {
       var idVal = String(data[i][idColIndex] || "");
       if (idVal.indexOf("P-" + importYear + "-") === 0) {
           var numPart = parseInt(idVal.split("-")[2], 10);
           if (!isNaN(numPart) && numPart > maxRunning) {
               maxRunning = numPart;
           }
       }
    }

    var newRows = [];
    // ตัวแปรสำหรับจับคู่เดือนที่ผู้ใช้พิมพ์ "ต.ค., พ.ย." ให้ตรงกับหัวคอลัมน์
    var monthMap = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];

    // 3. จัดเตรียมข้อมูลสำหรับเขียนลงชีต
    for (var j = 0; j < dataArray.length; j++) {
       var rowObj = dataArray[j];
       maxRunning++; // บวก Running Number ขึ้น 1 อัตโนมัติ
       var runningStr = ("000" + maxRunning).slice(-4); // เติม 0 ข้างหน้าให้ครบ 4 หลัก
       var newId = "P-" + importYear + "-" + runningStr;
       
       var newRow = new Array(headers.length).fill(""); // สร้าง Array ว่างๆ ความยาวเท่าจำนวนคอลัมน์
       
       // ฟังก์ชันย่อยสำหรับเขียนค่าลงคอลัมน์ให้ตรงช่อง
       function setVal(colName, val) {
           var idx = headers.indexOf(colName);
           if (idx > -1) newRow[idx] = val;
       }

       // --- 📌 ทำการ Mapping ข้อมูล (เอาข้อมูลจากหน้าบ้าน หยอดลงหลังบ้าน) ---
       setVal(COL_NAME.ID, newId);
       setVal(COL_NAME.ORDER, maxRunning);
       setVal("ปีงบประมาณ", rowObj["ปีงบประมาณ"]);
       setVal(COL_NAME.DEPT, rowObj["กลุ่มงาน/งาน"]);
       setVal(COL_NAME.PROJ, rowObj["โครงการ"]);
       setVal(COL_NAME.ACT, rowObj["กิจกรรมหลัก"]);
       setVal(COL_NAME.SUB, rowObj["กิจกรรมย่อย"]);
       setVal(COL_NAME.TYPE, rowObj["ประเภทงบ"]);
       setVal(COL_NAME.SOURCE, rowObj["แหล่งงบประมาณ"]);
       setVal("ผู้รับผิดชอบ", rowObj["ผู้รับผิดชอบ"]);
       setVal("รหัสงบประมาณ", rowObj["รหัสงบประมาณ"]);
       setVal("รหัสกิจกรรม", rowObj["รหัสกิจกรรม"]);
       
       // คำนวณยอดเงินเบื้องต้น
       var approveAmt = parseFloat(rowObj["อนุมัติตามแผน"]) || 0;
       var allocAmt = parseFloat(rowObj["จัดสรร"]) || 0;
       setVal(COL_NAME.APPROVE, approveAmt);
       setVal(COL_NAME.ALLOC, allocAmt);
       setVal(COL_NAME.SPENT, 0); // เริ่มต้นเบิกจ่ายเป็น 0
       setVal(COL_NAME.LOAN, 0);  // เริ่มต้นเงินยืมเป็น 0
       setVal(COL_NAME.BAL, allocAmt); // คงเหลือตั้งต้นเท่ากับยอดจัดสรร
       setVal(COL_NAME.STATUS, "ACTIVE");

       // --- 📌 การจัดการคอลัมน์เดือน (ถ้ามี เช่น "ต.ค., พ.ย.") ---
       var monthsStr = String(rowObj["เดือนที่ดำเนินการ"] || "");
       var monthsArr = monthsStr.split(",").map(function(item) { return item.trim(); });
       
       for (var m = 0; m < monthsArr.length; m++) {
           var mName = monthsArr[m];
           // เช็คว่ามีชื่อเดือนนี้อยู่ใน monthMap ไหม (กันคนพิมพ์ผิด)
           if (monthMap.indexOf(mName) > -1) {
               setVal(mName, 1); // ไปใส่เลข 1 ที่คอลัมน์ชื่อเดือนนั้นๆ
           }
       }

       newRows.push(newRow); // เก็บเข้าห่อใหญ่รอส่ง
    }

    // 4. บันทึกก้อนข้อมูลลงชีตในรอบเดียว (เสี้ยววินาที!)
    if (newRows.length > 0) {
       var startRow = sheet.getLastRow() + 1;
       sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
    }

    return { success: true, count: newRows.length };

  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock(); // ปลดล็อค
  }
}