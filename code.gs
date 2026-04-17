// ==========================================
// 1. CONFIGURATION & SETUP
// ==========================================
var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ';
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = 'Version : 690416-612-77';

// ==========================================
// 📌 1.1 DYNAMIC COLUMN MAPPING (แผนที่หัวคอลัมน์)
// ==========================================
var COL_NAME = {
    ID: 'รหัสโครงการ',
    YEAR: 'ปีงบประมาณ',
    CAT: 'หมวด',
    ORDER: 'ลำดับโครงการ',
    DEPT: 'กลุ่มงาน/งาน',
    PLAN: 'แผนงาน',
    PROJ: 'โครงการ',
    ACT: 'กิจกรรมหลัก',
    SUB: 'กิจกรรมย่อย',
    TYPE: 'ประเภทงบ',
    SOURCE: 'แหล่งงบประมาณ',
    BUDGET_CODE: 'รหัสงบประมาณ', // 👈 เพิ่มใหม่
    ACT_CODE: 'รหัสกิจกรรม',      // 👈 เพิ่มใหม่
    APPROVE: 'อนุมัติตามแผน',
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

// ==========================================
// 📌 HELPER FUNCTION: สร้าง Map เพื่อหา Index ของคอลัมน์จากชื่อหัวคอลัมน์
// ==========================================
function getColumnMap(sheet) {
    var map = {};
    // ดึงข้อมูลแถวแรก (หัวคอลัมน์) มาทั้งหมด
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (var i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim() !== "") {
            // จับคู่ 'ชื่อหัวคอลัมน์' : Index (0, 1, 2, ...)
            map[headers[i].toString().trim()] = i; 
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
// 📌 UNIVERSAL MAPPER: ตัวแปลงข้อมูลมาตรฐาน (Phase 2)
// ==========================================
function mapProjectRow(row, map) {
    var parseNum = function(val) { 
        var v = parseFloat(String(val).replace(/,/g, '')); 
        return isNaN(v) ? 0 : v; 
    };

    return {
        id: String(row[map[COL_NAME.ID]] || "").trim(),
        order: String(row[map[COL_NAME.ORDER]] || "").trim(),
        dept: String(row[map[COL_NAME.DEPT]] || "").trim(),
        proj: String(row[map[COL_NAME.PROJ]] || "").trim(),
        act: String(row[map[COL_NAME.ACT]] || "").trim(),
        subAct: String(row[map[COL_NAME.SUB]] || "").trim(),
        type: String(row[map[COL_NAME.TYPE]] || "").trim(),
        source: String(row[map[COL_NAME.SOURCE]] || "").trim(),
        budgetCode: String(row[map[COL_NAME.BUDGET_CODE]] || "").trim(),
        actCode: String(row[map[COL_NAME.ACT_CODE]] || "").trim(),
        
        approved: parseNum(row[map[COL_NAME.APPROVE]]),
        allocated: parseNum(row[map[COL_NAME.ALLOC]]),
        spent: parseNum(row[map[COL_NAME.SPENT]]),
        loan: parseNum(row[map[COL_NAME.LOAN]]),
        balance: parseNum(row[map[COL_NAME.BAL]]),
        
        status: String(row[map[COL_NAME.STATUS]] || 'ACTIVE').trim().toUpperCase(),
        
        timeline: [
            row[map[COL_NAME.M_OCT]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_NOV]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_DEC]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JAN]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_FEB]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_MAR]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_APR]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_MAY]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JUN]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JUL]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_AUG]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_SEP]] == 1 ? 1 : 0
        ]
    };
}

// ==========================================
// 📌 ดึงข้อมูล Master Plan (ใช้ท่อส่งมาตรฐานเดียว)
// ==========================================
function getAllMasterDataForClient() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return [];

        var map = getColumnMap(sheet);
        var data = sheet.getDataRange().getValues();
        var list = [];
        
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var id = String(row[map[COL_NAME.ID]] || "").trim();
            if (!id) continue; // ข้ามบรรทัดว่าง
            
            // 🌟 เรียกใช้ Mapper ตัวเดียวจบ ได้ข้อมูลสะอาดและมาตรฐาน 100%
            list.push(mapProjectRow(row, map));
        }
        return list;
    } catch (e) { 
        return { error: e.message }; 
    }
}

// แอดมินและ Client ใช้ข้อมูลมาตรฐานเดียวกัน ไม่ต้องเขียนโค้ดซ้ำ
function getAdminMasterPlan() {
    return getAllMasterDataForClient();
}

// ==========================================
// 4. DASHBOARD DATA
// ==========================================
function getDashboardData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return { error: "ไม่พบ Sheet" };

        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var map = {};
        for (var i = 0; i < headers.length; i++) {
            if (headers[i]) map[String(headers[i]).trim()] = i;
        }

        var I_DEPT = map['กลุ่มงาน/งาน'] !== undefined ? map['กลุ่มงาน/งาน'] : 4;
        var I_TYPE = map['ประเภทงบ'] !== undefined ? map['ประเภทงบ'] : 9;
        var I_SOURCE = map['แหล่งงบประมาณ'] !== undefined ? map['แหล่งงบประมาณ'] : 10;
        var I_STATUS = map['สถานะ'] !== undefined ? map['สถานะ'] : 14; // 🎯 เพิ่มคอลัมน์สถานะ

        var I_APPROVE = map['อนุมัติตามแผน'] !== undefined ? map['อนุมัติตามแผน'] : 15;
        var I_ALLOC = map['จัดสรร'] !== undefined ? map['จัดสรร'] : 16;
        var I_SPENT = map['เบิกจ่าย'] !== undefined ? map['เบิกจ่าย'] : 17;
        var I_BAL = map['คงเหลือ (ไม่รวมเงินยืม)'] !== undefined ? map['คงเหลือ (ไม่รวมเงินยืม)'] : 19;

        var summary = {
            moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} },
            loeiFund: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }
        };

        var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };

        var data = sheet.getDataRange().getValues();
        data.shift(); // ข้ามหัวตาราง

        data.forEach(function (row) {
            var typeVal = String(row[I_TYPE] || "").trim();
            var sourceVal = String(row[I_SOURCE] || "").trim();
            var statusVal = String(row[I_STATUS] || "").trim().toUpperCase(); // 🎯 ดึงค่าสถานะ (ทำเป็นตัวพิมพ์ใหญ่)
            var dept = String(row[I_DEPT] || 'ไม่ระบุ').trim();
            if (dept === '') dept = 'ไม่ระบุ';

            // โครงการต้องเป็น ACTIVE เท่านั้น ถึงจะนำมาคำนวณ!
            if (statusVal === 'ACTIVE') {

                // ----------------------------------------------------
                // ส่วนที่ 1: งบประมาณใน สป.สธ. (MOPH)
                // ----------------------------------------------------
                var isOldNonMoph = (
                    typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 ||
                    typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1
                );

                if (!isOldNonMoph) {
                    summary.moph.approved += parseNum(row[I_APPROVE]);
                    summary.moph.allocated += parseNum(row[I_ALLOC]);
                    summary.moph.spent += parseNum(row[I_SPENT]);
                    summary.moph.balance += parseNum(row[I_BAL]);

                    if (!summary.moph.deptStats[dept]) summary.moph.deptStats[dept] = { allocated: 0, spent: 0 };
                    summary.moph.deptStats[dept].allocated += parseNum(row[I_ALLOC]);
                    summary.moph.deptStats[dept].spent += parseNum(row[I_SPENT]);
                }

                // ----------------------------------------------------
                // 🎯 ส่วนที่ 2: งบประมาณเงินนอก สป. (เฉพาะเงินบำรุง สสจ.เลย)
                // ----------------------------------------------------
                var cleanType = typeVal.replace(/\s+/g, '');
                var cleanSource = sourceVal.replace(/\s+/g, '');

                // เช็คแค่ 2 ข้อ เพราะดัก ACTIVE ไว้ด้านบนแล้ว
                var isLoeiFund = (
                    cleanType === 'เงินนอกสป.' &&
                    cleanSource === 'เงินบำรุงสสจ.เลย'
                );

                if (isLoeiFund) {
                    summary.loeiFund.approved += parseNum(row[I_APPROVE]);
                    summary.loeiFund.allocated += parseNum(row[I_ALLOC]);
                    summary.loeiFund.spent += parseNum(row[I_SPENT]);
                    summary.loeiFund.balance += parseNum(row[I_BAL]);

                    if (!summary.loeiFund.deptStats[dept]) summary.loeiFund.deptStats[dept] = { allocated: 0, spent: 0 };
                    summary.loeiFund.deptStats[dept].allocated += parseNum(row[I_ALLOC]);
                    summary.loeiFund.deptStats[dept].spent += parseNum(row[I_SPENT]);
                }
            }
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
function saveTransaction(payload) {
    try {
        var lock = LockService.getScriptLock();
        lock.waitLock(10000);

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetT = ss.getSheetByName('t_actionplan');
        var sheetM = ss.getSheetByName('m_actionplan');
        
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);

        // 🔍 1. ค้นหาแถวของโครงการเป้าหมายใน m_actionplan
        var targetRow = -1;
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(payload.projectId).trim()) {
                targetRow = i;
                break;
            }
        }

        if (targetRow === -1) {
            return { success: false, message: 'ไม่พบรหัสโครงการนี้ในฐานข้อมูลหลัก (ID: ' + payload.projectId + ')' };
        }

        // 📊 2. ดึงค่ายอดเงินปัจจุบัน
        var spentIndex = map[COL_NAME.SPENT];
        var balIndex = map[COL_NAME.BAL];

        var currentSpent = Number(mData[targetRow][spentIndex]) || 0;
        var currentBal = Number(mData[targetRow][balIndex]) || 0;
        var amount = Number(payload.amount) || 0;

        // ดักจับความปลอดภัย ห้ามเบิกเกินงบ
        if (amount > currentBal) { 
            return { success: false, message: 'ยอดเงินคงเหลือไม่เพียงพอสำหรับการเบิกจ่าย' };
        }

        // 💾 3. อัปเดตตาราง Master (บวกยอดเบิกจ่าย, ลบยอดคงเหลือ)
        sheetM.getRange(targetRow + 1, spentIndex + 1).setValue(currentSpent + amount);
        sheetM.getRange(targetRow + 1, balIndex + 1).setValue(currentBal - amount);

        // 📋 4. ดึงข้อมูล Year, Cat, Plan จากตาราง Master มาเติมให้สมบูรณ์
        var mYear = map[COL_NAME.YEAR] !== undefined ? mData[targetRow][map[COL_NAME.YEAR]] : "";
        var mCat = map[COL_NAME.CAT] !== undefined ? mData[targetRow][map[COL_NAME.CAT]] : "";
        var mPlan = map[COL_NAME.PLAN] !== undefined ? mData[targetRow][map[COL_NAME.PLAN]] : "";

        // 🎯 5. สร้าง Array แถวใหม่สำหรับตารางประวัติ (t_actionplan) โยนลงช่องให้เป๊ะ!
        var newRowT = [];
        newRowT.length = 30; // จองพื้นที่ว่างไว้ 30 คอลัมน์กันเหนียว
        newRowT.fill('');
        
        newRowT[0] = new Date();                     // A: ประทับเวลา
        newRowT[1] = payload.projectId;              // B: รหัสโครงการ
        newRowT[2] = mYear;                          // C: ปีงบประมาณ
        newRowT[3] = mCat;                           // D: หมวด
        newRowT[4] = payload.order;                  // E: ลำดับ
        newRowT[5] = payload.dept;                   // F: กลุ่มงาน
        newRowT[6] = mPlan;                          // G: แผนงาน
        newRowT[7] = payload.proj;                   // H: โครงการ
        newRowT[8] = payload.act;                    // I: กิจกรรมหลัก
        newRowT[9] = payload.sub;                    // J: กิจกรรมย่อย
        newRowT[10] = payload.type;                  // K: ประเภทงบ
        newRowT[11] = payload.source;                // L: แหล่งงบ
        
        newRowT[15] = amount;                        // P: ยอดเบิกจ่าย
        newRowT[16] = currentBal - amount;           // Q: คงเหลือ
        newRowT[17] = formatToStorageDate(payload.txDate); // R: วันที่เบิกจ่าย (แปลงเป็น format ชีต)
        newRowT[18] = payload.expenseType || '';     // S: หมวดรายจ่าย
        newRowT[19] = payload.desc || payload.remark || ''; // T: รายละเอียด/หมายเหตุ
        
        newRowT[25] = formatToStorageDate(payload.startDate) || ''; // Z: เริ่มดำเนินการ (สำหรับ Gantt)
        newRowT[26] = formatToStorageDate(payload.endDate) || '';   // AA: สิ้นสุดดำเนินการ (สำหรับ Gantt)

        // 🚀 6. บันทึกประวัติและปลดล็อค
        sheetT.appendRow(newRowT);
        SpreadsheetApp.flush();
        lock.releaseLock();
        
        return { success: true, count: 1 };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    }
}

// ==========================================
// 📌 ฟังก์ชันลบประวัติเบิกจ่ายและดึงยอดเงินคืน (Delete Transaction)
// ==========================================
function deleteTransaction(rowIndex, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetT = ss.getSheetByName('t_actionplan');
        var sheetM = ss.getSheetByName('m_actionplan');

        // 1. ตรวจสอบว่าแถวที่จะลบในประวัติ มีรหัสโครงการตรงกันหรือไม่ (ป้องกันการลบผิดแถว)
        var tProj = String(sheetT.getRange(rowIndex, 2).getValue()).trim();
        if (tProj !== String(projectId).trim()) {
            return { success: false, message: 'ข้อมูลไม่ตรงกัน ไม่สามารถลบได้ (อาจมีการเปลี่ยนแปลงข้อมูลก่อนหน้า โปรดรีเฟรช)' };
        }

        // 2. หายอดเงินของโครงการนั้น เพื่อดึงเงินคืนกลับเข้ากระเป๋า
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;
        
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                targetRowM = i + 1; 
                break;
            }
        }

        if (targetRowM > -1) {
            var spentIdx = map[COL_NAME.SPENT] + 1;
            var balIdx = map[COL_NAME.BAL] + 1;
            
            var currentSpent = Number(sheetM.getRange(targetRowM, spentIdx).getValue()) || 0;
            var currentBal = Number(sheetM.getRange(targetRowM, balIdx).getValue()) || 0;
            var refundAmount = Number(amount) || 0;

            var newSpent = currentSpent - refundAmount;
            if (newSpent < 0) newSpent = 0; // กันติดลบ

            // 🎯 คำนวณให้เสร็จสรรพใน Script ไม่ต้องง้อสูตรในชีต!
            sheetM.getRange(targetRowM, spentIdx).setValue(newSpent);
            sheetM.getRange(targetRowM, balIdx).setValue(currentBal + refundAmount);
        }

        // 3. ลบแถวประวัติทิ้ง
        sheetT.deleteRow(rowIndex);

        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
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
            var currentSpent = parseFloat(String(cellSpent.getValue()).replace(/,/g, '')) || 0;
            
            var oldVal = parseFloat(form.oldAmount) || 0;
            var newVal = parseFloat(form.newAmount) || 0;
            
            // คำนวณยอดเบิกจ่ายสะสมใหม่
            var newSpentTotal = currentSpent - oldVal + newVal;
            
            // บันทึกเฉพาะยอดเบิกจ่าย สูตรใน Sheet จะจัดการยอดคงเหลือให้เอง
            cellSpent.setValue(newSpentTotal); 
        }
        tSheet.getRange(tRow, 16).setValue(form.newAmount);
        tSheet.getRange(tRow, 18).setValue(formatToStorageDate(form.date));
        tSheet.getRange(tRow, 20).setValue(form.desc);
        tSheet.getRange(tRow, 24).setValue(form.reason);
        tSheet.getRange(tRow, 25).setValue(newBalance);
        return { status: 'success', message: 'แก้ไขเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
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
        var sDate = form.startDate ? formatToStorageDate(form.startDate) : "";
        var eDate = form.endDate ? formatToStorageDate(form.endDate) : "";

        cellMasterLoan.setValue(newMasterLoan);
        SpreadsheetApp.flush();
        tSheet.appendRow([
            new Date(), r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[13], r[14], r[16],
            loanAmt, formatToStorageDate(form.loanDate), form.loanType, form.desc, "ยังไม่ดำเนินการ", 0, loanAmt, "", "", "", sDate, eDate
        ]);
        return { status: 'success', message: 'บันทึกเงินยืมเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
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
        tSheet.getRange(tRowIndex, 23).setValue(formatToStorageDate(data.payDate));

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
    } catch (e) {
        return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.toString() };
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
        
        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[1])) continue;
            var projId = String(row[1]).trim();
            var masterInfo = masterMap[projId] || { allocated: 0, spent: 0, balance: 0 };
            var item = {
                rowId: i + 1, 
                order: row[4], 
                project: row[7], 
                activity: row[8], 
                subActivity: row[9],
                amount: parseAmount(row[15]), 
                date: formatToThaiUI(row[17]), // ✅ อัปเดตวันที่
                type: row[18], 
                source: row[11], 
                desc: row[19], 
                id: row[1],
                masterAllocated: masterInfo.allocated, 
                masterSpent: masterInfo.spent, 
                masterBalance: masterInfo.balance
            };
            if (item.amount > 0 || item.order) result.push(item);
            if (result.length >= 100) break;
        }
        return result;
    } catch (e) { return []; }
}

// --- ดึงประวัติเงินยืม ---
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
    } catch (e) { }
    var tData = tSheet.getDataRange().getDisplayValues();
    var result = [];
    var parseNum = function (val) { return parseFloat(String(val).replace(/,/g, '')) || 0; };
    
    for (var i = tData.length - 1; i >= 1; i--) {
        try {
            var row = tData[i];
            if (!row[0] && !row[1]) continue;
            var pid = String(row[1] || "").trim();
            var meta = projectMap[pid] || { type: '-', source: '-' };
            result.push({
                id: row[0], 
                timestamp: row[0], 
                projectId: row[1], // 👈 เพิ่มบรรทัดนี้! เพื่อส่งรหัสโครงการ (P-2569-xxxx)
                rowId: i + 1,      // 👈 เพิ่มบรรทัดนี้! เพื่อส่งเลขแถวสำหรับให้ปุ่มลบทำงาน
                project: row[7], 
                activity: row[8], 
                subActivity: row[9],
                amount: parseNum(row[15]), 
                date: formatToThaiUI(row[16]), 
                status: row[19], 
                paid: parseNum(row[20]), 
                balance: parseNum(row[21]),
                order: row[4], 
                type: row[17], 
                desc: row[18], 
                budgetType: meta.type, 
                budgetSource: meta.source, 
                dept: row[2]
             });
        } catch (e) { }
    }
    return result;
}

// --- ฟังก์ชันดึงประวัติรวม (Generic) ---
function getHistory(sheetName) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName(sheetName);
        if (!tSheet) return [];
        var data = tSheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseAmount = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };
        
        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[1])) continue;
            var item = {};
            if (sheetName === 't_actionplan') {
                item = {
                    rowId: i + 1, order: row[4], project: row[7], activity: row[8], subActivity: row[9], amount: parseAmount(row[15]),
                    date: formatToThaiUI(row[17]), // ✅ อัปเดตวันที่
                    type: row[18], source: row[11], desc: row[19], id: row[1]
                };
            } else {
                item = { timestamp: row[0], order: row[4], project: row[7], amount: parseAmount(row[15]), date: formatToThaiUI(row[16]), status: row[19] };
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
            newTotal, addAmount, formatToStorageDate(form.date), form.letterNo, form.remark
        ];
        tAllocSheet.appendRow(logRow);
        return { status: 'success', message: 'บันทึกจัดสรรสำเร็จ (ยอดใหม่: ' + newTotal.toLocaleString() + ')' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

// --- ดึงประวัติจัดสรรงบประมาณ ---
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
                id: row[1], 
                order: row[4], 
                dept: row[5],
                project: row[7], 
                activity: row[8], 
                subActivity: row[9], 
                type: row[10], 
                source: row[11],
                accumulatedAlloc: parseNum(row[14]), 
                currentAlloc: parseNum(row[15]), 
                date: formatToThaiUI(row[16]), // ✅ อัปเดตวันที่
                letterNo: row[17]
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

// /**
//  * 📌 ฟังก์ชันดึงข้อมูลสรุปรายงาน (Summary)
//  * อัปเกรดให้ใช้ Standard Keys และระบบ Filter ล่าสุด
//  */
// function getSummaryData(typeFilter, sourceFilter, deptFilter) {
//     try {
//         // 🌟 เรียกใช้ฟังก์ชันดึงข้อมูลมาตรฐานที่เราทำไว้ใน Phase 2
//         var data = getAllMasterDataForClient(); 
//         if (!data || data.length === 0) return { status: 'empty', message: 'ไม่มีข้อมูลโครงการ' };

//         var grandTotal = { allocated: 0, spent: 0, balance: 0, count: 0 };
//         var byDeptFiltered = {};
//         var bySource = {};
        
//         // 🎯 ตะกร้า 6 ใบสำหรับตารางวิเคราะห์งบประมาณ
//         var customSourceTable = {
//             'สป.สธ.': { allocated: 0, spent: 0 },
//             'เงิน นอก สป.': { allocated: 0, spent: 0 },
//             'เงินบำรุง สสจ.เลย': { allocated: 0, spent: 0 },
//             'เงินบำรุง HRD': { allocated: 0, spent: 0 },
//             'งบพัฒนากลุ่มจังหวัด': { allocated: 0, spent: 0 },
//             'งบพัฒนาจังหวัด': { allocated: 0, spent: 0 }
//         };

//         data.forEach(function(row) {
//             if (row.status !== 'ACTIVE') return;

//             var alloc = Number(row.allocated || 0);
//             var spent = Number(row.spent || 0);
//             var typeVal = String(row.type || "").trim();
//             var sourceVal = String(row.source || "").trim();
//             var actCodeVal = String(row.actCode || "").trim();
//             var deptVal = String(row.dept || "ไม่ระบุ").trim();

//             var cType = typeVal.replace(/\s+/g, '');
//             var cSource = sourceVal.replace(/\s+/g, '');

//             // 🌟 Logic การจัดกลุ่ม 6 ตะกร้า
//             if (cType === 'สป.สธ.') {
//                 customSourceTable['สป.สธ.'].allocated += alloc;
//                 customSourceTable['สป.สธ.'].spent += spent;
//             } else if (cType === 'เงินนอกสป.') {
//                 if (cSource === 'งบพัฒนากลุ่มจังหวัด') {
//                     customSourceTable['งบพัฒนากลุ่มจังหวัด'].allocated += alloc;
//                     customSourceTable['งบพัฒนากลุ่มจังหวัด'].spent += spent;
//                 } else if (cSource === 'งบพัฒนาจังหวัด') {
//                     customSourceTable['งบพัฒนาจังหวัด'].allocated += alloc;
//                     customSourceTable['งบพัฒนาจังหวัด'].spent += spent;
//                 } else if (cSource === 'เงินบำรุงสสจ.เลย') {
//                     if (actCodeVal.includes('HRD')) { // กรองผ่านรหัสกิจกรรม
//                         customSourceTable['เงินบำรุง HRD'].allocated += alloc;
//                         customSourceTable['เงินบำรุง HRD'].spent += spent;
//                     } else {
//                         customSourceTable['เงินบำรุง สสจ.เลย'].allocated += alloc;
//                         customSourceTable['เงินบำรุง สสจ.เลย'].spent += spent;
//                     }
//                 } else {
//                     customSourceTable['เงิน นอก สป.'].allocated += alloc;
//                     customSourceTable['เงิน นอก สป.'].spent += spent;
//                 }
//             }

//             // 🎯 กรองตามตัวเลือก Filter ในหน้าจอ
//             var passType = !typeFilter || cType === typeFilter.replace(/\s+/g, '');
//             var passSource = !sourceFilter || cSource === sourceFilter.replace(/\s+/g, '');
//             var passDept = !deptFilter || deptVal === deptFilter;

//             if (passType && passSource && passDept) {
//                 grandTotal.allocated += alloc;
//                 grandTotal.spent += spent;
//                 grandTotal.balance += (row.balance || 0);
//                 grandTotal.count++;

//                 if (!byDeptFiltered[deptVal]) byDeptFiltered[deptVal] = { allocated: 0, spent: 0 };
//                 byDeptFiltered[deptVal].allocated += alloc;
//                 byDeptFiltered[deptVal].spent += spent;
//             }

//             // จัดกลุ่มสำหรับวาดกราฟ Pie
//             var group = typeVal.includes('เงินนอก') ? "เงินนอกงบประมาณ" : "งบประมาณ (สป.สธ.)";
//             if (!bySource[group]) bySource[group] = { allocated: 0 };
//             bySource[group].allocated += alloc;
//         });

//         // แปลง Object เป็น Array เพื่อส่งให้ JavaScript หน้าบ้าน
//         var toList = (obj) => Object.keys(obj).map(k => ({ name: k, allocated: obj[k].allocated, spent: obj[k].spent }))
//                             .sort((a, b) => b.allocated - a.allocated);

//         return {
//             status: 'success',
//             grandTotal: grandTotal,
//             customSourceTable: customSourceTable,
//             byDeptFiltered: toList(byDeptFiltered),
//             bySource: toList(bySource)
//         };
//     } catch (e) { return { status: 'error', message: e.message }; }
// }

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
function saveMasterPlan(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        var map = getColumnMap(sheet);

        var selectedMonths = [];
        if (payload.months) {
            selectedMonths = payload.months.split(',').map(function (m) { return m.trim(); });
        }
        var checkMonth = function (monthName) { return selectedMonths.indexOf(monthName) !== -1 ? 1 : ""; };

        if (payload.id) { // ✏️ กรณีแก้ไข (UPDATE)
            var data = sheet.getDataRange().getValues();
            var rowIndex = -1;
            for (var i = 1; i < data.length; i++) {
                if (String(data[i][map[COL_NAME.ID]]) === String(payload.id)) { rowIndex = i + 1; break; }
            }
            if (rowIndex > -1) {
                if (map[COL_NAME.YEAR] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.YEAR] + 1).setValue(payload.year);
                if (map[COL_NAME.CAT] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.CAT] + 1).setValue(payload.cat);
                
                sheet.getRange(rowIndex, map[COL_NAME.ORDER] + 1).setValue(payload.order);
                sheet.getRange(rowIndex, map[COL_NAME.DEPT] + 1).setValue(payload.dept);
                sheet.getRange(rowIndex, map[COL_NAME.PLAN] + 1).setValue(payload.plan);
                sheet.getRange(rowIndex, map[COL_NAME.PROJ] + 1).setValue(payload.proj);
                sheet.getRange(rowIndex, map[COL_NAME.ACT] + 1).setValue(payload.act);
                sheet.getRange(rowIndex, map[COL_NAME.SUB] + 1).setValue(payload.sub);
                sheet.getRange(rowIndex, map[COL_NAME.TYPE] + 1).setValue(payload.type);
                sheet.getRange(rowIndex, map[COL_NAME.SOURCE] + 1).setValue(payload.source);

                if (map[COL_NAME.BUDGET_CODE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.BUDGET_CODE] + 1).setValue(payload.budgetCode);
                if (map[COL_NAME.ACT_CODE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.ACT_CODE] + 1).setValue(payload.actCode);          
                
                sheet.getRange(rowIndex, map[COL_NAME.APPROVE] + 1).setValue(payload.approve);
                // 🌟 แก้ไขเป็น payload.allocated ให้ตรงกับหน้าบ้าน
                sheet.getRange(rowIndex, map[COL_NAME.ALLOC] + 1).setValue(payload.allocated);
                sheet.getRange(rowIndex, map[COL_NAME.STATUS] + 1).setValue(payload.status);
                if (map[COL_NAME.REMARK] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.REMARK] + 1).setValue(payload.remark);

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
        } else { // ➕ กรณีเพิ่มใหม่ (INSERT)
            var newRow = new Array(sheet.getLastColumn()).fill("");
            var newId = generateNextProjectID();

            newRow[map[COL_NAME.ID]] = newId;
            
            if (map[COL_NAME.YEAR] !== undefined) newRow[map[COL_NAME.YEAR]] = payload.year;
            if (map[COL_NAME.CAT] !== undefined) newRow[map[COL_NAME.CAT]] = payload.cat;

            newRow[map[COL_NAME.ORDER]] = payload.order;
            newRow[map[COL_NAME.DEPT]] = payload.dept;
            newRow[map[COL_NAME.PLAN]] = payload.plan;
            newRow[map[COL_NAME.PROJ]] = payload.proj;
            newRow[map[COL_NAME.ACT]] = payload.act;
            newRow[map[COL_NAME.SUB]] = payload.sub;
            newRow[map[COL_NAME.TYPE]] = payload.type;
            newRow[map[COL_NAME.SOURCE]] = payload.source;

            if (map[COL_NAME.BUDGET_CODE] !== undefined) newRow[map[COL_NAME.BUDGET_CODE]] = payload.budgetCode; 
            if (map[COL_NAME.ACT_CODE] !== undefined) newRow[map[COL_NAME.ACT_CODE]] = payload.actCode;          

            newRow[map[COL_NAME.APPROVE]] = payload.approve;
            // 🌟 แก้ไขเป็น payload.allocated ให้ตรงกับหน้าบ้าน
            newRow[map[COL_NAME.ALLOC]] = payload.allocated;
            newRow[map[COL_NAME.STATUS]] = payload.status;
            if (map[COL_NAME.REMARK] !== undefined) newRow[map[COL_NAME.REMARK]] = payload.remark;

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
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var fetchColumnData = function (sheetName) {
            var sheet = ss.getSheetByName(sheetName);
            if (!sheet) return [];
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return [];
            var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            return [...new Set(data.map(function (r) { return String(r[0]).trim(); }).filter(String))];
        };
        
        return {
            cats: fetchColumnData('c_category'),   // 👈 เพิ่มใหม่: ดึงข้อมูลหมวด
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
            var monthsArr = monthsStr.split(",").map(function (item) { return item.trim(); });

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

// ==========================================
// 🛠️ ONE-TIME SCRIPT: อัปเดตรูปแบบวันที่ในฐานข้อมูล
// ==========================================
function migrateAllDateFormats() {
  var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // กำหนดชีตและตำแหน่งคอลัมน์ที่มีวันที่ (เริ่มนับ A=1, B=2...)
  var targetSheets = [
    { name: 't_actionplan', cols: [18] }, // R = 18
    { name: 't_loan', cols: [17, 23, 26, 27] }, // Q=17, W=23, Z=26, AA=27
    { name: 't_allocate', cols: [17] } // Q = 17
  ];

  var totalUpdated = 0;

  for (var s = 0; s < targetSheets.length; s++) {
    var sheetInfo = targetSheets[s];
    var sheet = ss.getSheetByName(sheetInfo.name);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue; // ไม่มีข้อมูลให้ข้ามไป

    for (var c = 0; c < sheetInfo.cols.length; c++) {
      var colIndex = sheetInfo.cols[c];
      var range = sheet.getRange(2, colIndex, lastRow - 1, 1);
      var values = range.getValues();

      for (var r = 0; r < values.length; r++) {
        var val = values[r][0];
        if (val !== "" && val !== null) {
          values[r][0] = convertToStandardDate_(val);
          totalUpdated++;
        }
      }
      
      // บังคับ Format เซลล์ให้เป็น Text เพื่อล็อกไม่ให้ Google Sheets แปลงกลับเป็น M/D/Y อัตโนมัติ
      range.setNumberFormat("@"); 
      // บันทึกค่ากลับลงชีต
      range.setValues(values);
    }
  }
  
  SpreadsheetApp.flush();
  Logger.log("✅ แปลงรูปแบบวันที่เรียบร้อยแล้ว จำนวน " + totalUpdated + " เซลล์");
}

// 🔧 ฟังก์ชันย่อยสำหรับแปลงวันที่ (ไม่ต้องกด Run ตัวนี้)
function convertToStandardDate_(val) {
  try {
    var d;
    
    // 1. ถ้าเป็น Date Object อยู่แล้ว
    if (Object.prototype.toString.call(val) === '[object Date]') {
      d = val;
    } 
    // 2. ถ้าเป็น String แบบ YYYY-MM-DD
    else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
      var parts = val.split('-');
      d = new Date(parts[0], parseInt(parts[1]) - 1, parts[2]);
    } 
    // 3. ถ้ามี / อยู่แล้ว (อาจจะตรงฟอร์แมตแล้ว)
    else if (typeof val === 'string' && val.indexOf('/') > -1) {
      // เช็คว่าถ้าเป็น DD/MM/YYYY หรือ D/M/YYYY อยู่แล้ว ให้ปล่อยผ่านได้เลย
      if (val.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) return val;
      d = new Date(val);
    } 
    // 4. กรณีอื่นๆ
    else {
      d = new Date(val);
    }

    if (!isNaN(d.getTime())) {
      var dd = d.getDate(); // ไม่เติม 0 ข้างหน้าเพื่อให้ออกมาเป็น 15/1/2026 ตามที่ท่านต้องการ (ไม่ใช่ 15/01/2026)
      var mm = d.getMonth() + 1;
      var yyyy = d.getFullYear();
      return dd + '/' + mm + '/' + yyyy;
    }
  } catch (e) {
    // ถ้าจับ Format ไม่ได้จริงๆ ให้คืนค่าเดิมกลับไป
  }
  return val;
}

// ==========================================
// 📅 DATE UTILITY FUNCTIONS (จัดการรูปแบบวันที่)
// ==========================================

// ==========================================
// 📅 1. ส่วนกลาง: จัดการรูปแบบวันที่ (มาตรฐานใหม่)
// ==========================================

// แปลงจากหน้าเว็บ (YYYY-MM-DD) -> (DD/MM/YYYY) เพื่อบันทึกลง Sheet
function formatToStorageDate(dateStr) {
    if (!dateStr) return "";
    try {
        var parts = dateStr.split('-');
        if (parts.length === 3) {
            return parseInt(parts[2], 10) + '/' + parseInt(parts[1], 10) + '/' + parts[0];
        }
    } catch(e) {}
    return String(dateStr);
}

// แปลงจาก Sheet -> (15 ม.ค. 2569) เพื่อแสดงผลบนหน้าเว็บ
function formatToThaiUI(dateVal) {
    if (!dateVal) return "-";
    try {
        var d;
        var strVal = String(dateVal).trim();
        
        if (Object.prototype.toString.call(dateVal) === '[object Date]') {
            d = dateVal;
        } else if (strVal.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { 
            var parts = strVal.split('/');
            d = new Date(parts[2], parseInt(parts[1], 10) - 1, parts[0]);
        } else if (strVal.match(/^\d{4}-\d{2}-\d{2}$/)) { 
            var parts = strVal.split('-');
            d = new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        } else {
            d = new Date(dateVal);
        }

        if (isNaN(d.getTime())) return strVal;

        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
    } catch (ex) {
        return String(dateVal);
    }
}

// ==========================================
// 📌 1. ฟังก์ชันล้างหนี้ / คืนเงินยืม (Repayment)
// ==========================================
function saveRepayment(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetM = ss.getSheetByName('m_actionplan');
        
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(payload.projectId).trim()) {
                targetRowM = i + 1; 
                break;
            }
        }

        if (targetRowM > -1) {
            var spentIdx = map[COL_NAME.SPENT] + 1;
            var loanIdx = map[COL_NAME.LOAN] + 1;

            var curSpent = Number(sheetM.getRange(targetRowM, spentIdx).getValue()) || 0;
            var curLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;

            var actualSpent = Number(payload.actualSpent) || 0;
            var returnedAmount = Number(payload.returnedAmount) || 0;
            var totalClear = actualSpent + returnedAmount;

            if (totalClear > curLoan) {
                return { success: false, message: 'ยอดเคลียร์บิล+เงินทอน มากกว่ายอดเงินยืมคงค้างในระบบ' };
            }

            // 🎯 อัปเดตเฉพาะยอด "เบิกจ่าย" และ "เงินยืม" (ปล่อยคอลัมน์ คงเหลือ ให้สูตรคำนวณเอง)
            sheetM.getRange(targetRowM, spentIdx).setValue(curSpent + actualSpent);
            sheetM.getRange(targetRowM, loanIdx).setValue(curLoan - totalClear);
            
            SpreadsheetApp.flush();
            return { success: true };
        } else {
            return { success: false, message: 'ไม่พบรหัสโครงการนี้ในระบบ' };
        }
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 2. ฟังก์ชันยกเลิก/ลบรายการเงินยืม (Delete Loan)
// ==========================================
function deleteLoanBackend(rowIndex, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        // *เช็คชื่อชีตประวัติยืมเงินให้ตรงกับของเจ้านายด้วยนะครับ
        var sheetL = ss.getSheetByName('t_loan') || ss.getSheetByName('t_actionplan'); 
        var sheetM = ss.getSheetByName('m_actionplan');

        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;
        
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                targetRowM = i + 1; 
                break;
            }
        }

        if (targetRowM > -1) {
            var loanIdx = map[COL_NAME.LOAN] + 1;
            var currentLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;
            var refundAmount = Number(amount) || 0;

            // 🎯 หักยอดออกจากช่องเงินยืมอย่างเดียว (ปล่อยคอลัมน์ คงเหลือ ให้สูตรคำนวณเอง)
            var newLoan = currentLoan - refundAmount;
            if (newLoan < 0) newLoan = 0; // กันติดลบ
            sheetM.getRange(targetRowM, loanIdx).setValue(newLoan);
        }

        // ลบแถวประวัติทิ้ง
        if(sheetL) sheetL.deleteRow(rowIndex);

        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}