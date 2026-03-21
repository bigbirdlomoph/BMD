// ==========================================
// AUTHENTICATION & ROLE MANAGEMENT (EMAIL ONLY)
// ==========================================
var AUTH_SHEET = 'auth_users';
var AUTH_LOG_SHEET = 'auth_users_log';

// 1. ฟังก์ชันยืนยันตัวตนด้วยการรับค่า Email ที่ผู้ใช้พิมพ์เข้ามาเช็ค
function verifyEmailOnly(inputEmail) {
  try {
    if (!inputEmail || inputEmail.trim() === "") {
      return { status: 'error', message: 'กรุณาระบุอีเมล' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var authSheet = ss.getSheetByName(AUTH_SHEET);
    if (!authSheet) return { status: 'error', message: 'ไม่พบฐานข้อมูลผู้ใช้งาน' };

    var data = authSheet.getDataRange().getValues();
    var userFound = null;
    var searchEmail = String(inputEmail).trim().toLowerCase();

    // วนลูปหา Email (เริ่มที่ 1 ข้ามหัวตาราง)
    for (var i = 1; i < data.length; i++) {
      var dbEmail = String(data[i][0] || "").trim().toLowerCase();

      if (dbEmail === searchEmail && dbEmail !== "") {
        userFound = {
          email: data[i][0],
          firstName: data[i][1],
          lastName: data[i][2],
          fullName: data[i][1] + ' ' + data[i][2],
          position: data[i][3],
          department: data[i][4],
          role: String(data[i][5] || "").trim().toUpperCase(),
          status: String(data[i][6] || "").trim().toUpperCase(),
          phone: data[i][7]
        };
        break;
      }
    }

    if (userFound) {
      if (userFound.status === 'ACTIVE') {
        logAuthAction(userFound.email, userFound.fullName, 'LOGIN_SUCCESS', 'ยืนยันตัวตนด้วย Email สำเร็จ');
        return { status: 'success', user: userFound };
      } else {
        logAuthAction(userFound.email, userFound.fullName, 'LOGIN_FAIL', 'บัญชีถูกระงับ');
        return { status: 'error', message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
      }
    } else {
      logAuthAction(inputEmail, 'Unknown', 'LOGIN_FAIL', 'ไม่พบอีเมลนี้ในระบบ');
      return { status: 'error', message: 'ปฏิเสธการเข้าถึง! ไม่พบอีเมลนี้ในระบบ' };
    }
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// 2. บันทึก Log การเข้าใช้งาน
function logAuthAction(email, fullName, action, detail) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(AUTH_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(AUTH_LOG_SHEET);
      logSheet.appendRow(['timestamp', 'email', 'full_name', 'action', 'detail']);
    }
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([timestamp, email, fullName, action, detail]);
  } catch (e) { console.log("Log Auth Error: " + e); }
}

// 3. ป้องกันการเข้าถึง API (เช็คจากตัวแปรที่ฝั่งหน้าบ้านส่งมา)
function requirePermission(clientEmail) {
  if(!clientEmail) throw new Error('Permission Denied: เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
}
function requireAdmin() {
  // สำหรับฟังก์ชันเฉพาะแอดมิน
}

// ==========================================
// ADMIN PANEL FUNCTIONS
// ==========================================
function getAllUsers() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(AUTH_SHEET);
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
}

function saveUser(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(AUTH_SHEET);
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
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(AUTH_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
      sheet.deleteRow(i + 1);
      return {status: 'success', message: 'ลบผู้ใช้งานเรียบร้อย'};
    }
  }
  return {status: 'error', message: 'ไม่พบอีเมลที่ต้องการลบ'};
}
