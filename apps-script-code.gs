// ============================================================
//  aisFix — Google Apps Script (Code.gs)
//  วิธีใช้งาน: วางโค้ดนี้ใน Google Apps Script ของคุณ แล้ว Deploy เป็น Web App
//  - Execute as: Me (บัญชีของคุณ)
//  - Who has access: Anyone (ทุกคนรวมถึงผู้ใช้นอกระบบ)
// ============================================================

const SHEET_USERS = 'Users';
const SHEET_DATA = 'Data';
const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── GET API ────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  let result;
  
  try {
    if (action === 'getUserData') {
      const username = e.parameter.username || '';
      result = getUserData(username);
    } else {
      result = { status: 'error', message: 'ไม่พบ Action ที่ระบุ หรือรูปแบบ GET ไม่ถูกต้อง' };
    }
  } catch (err) {
    result = { status: 'error', message: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST API ───────────────────────────────────────────────
function doPost(e) {
  let body, result;
  
  try {
    // แก้ไขปัญหา CORS ด้วยการรองรับทั้ง application/json และ text/plain
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else {
      body = {};
    }
    
    const action = body.action || '';
    
    if (action === 'login') {
      result = loginUser(body.username, body.password);
    } else if (action === 'addData') {
      result = addDigitNumber(body.username, body.digitNumber);
    } else {
      result = { status: 'error', message: 'ไม่พบ Action ที่ระบุ หรือรูปแบบ POST ไม่ถูกต้อง' };
    }
  } catch (err) {
    result = { status: 'error', message: 'การประมวลผลล้มเหลว: ' + err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── HELPER: ดึงชีต Users (ถ้าไม่มีจะสร้างให้พร้อมข้อมูลเริ่มต้น) ───
function getUsersSheet() {
  let sheet = SS.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = SS.insertSheet(SHEET_USERS);
    // สร้าง Header
    sheet.appendRow(['username', 'password', 'name', 'createdAt']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f1f5f9');
    
    // ใส่ User เริ่มต้นเพื่อพร้อมใช้งานทันที
    const defaultUsers = [
      ['admin', '1234', 'ผู้ดูแลระบบ (Admin)', new Date().toISOString()],
      ['user1', '1234', 'ผู้ใช้งานทั่วไป 1 (User 1)', new Date().toISOString()],
      ['user2', '1234', 'ผู้ใช้งานทั่วไป 2 (User 2)', new Date().toISOString()]
    ];
    defaultUsers.forEach(function(row) {
      sheet.appendRow(row);
    });
  }
  return sheet;
}

// ─── HELPER: ดึงชีต Data (ถ้าไม่มีจะสร้างให้) ───────────────────
function getDataSheet() {
  let sheet = SS.getSheetByName(SHEET_DATA);
  if (!sheet) {
    sheet = SS.insertSheet(SHEET_DATA);
    // สร้าง Header
    sheet.appendRow(['id', 'username', 'digitNumber', 'timestamp']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e2e8f0');
  }
  return sheet;
}

// ─── ฟังก์ชันตรวจสอบสิทธิ์การเข้าสู่ระบบ (Login) ───────────────────
function loginUser(username, password) {
  if (!username || !password) {
    return { status: 'error', message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
  }
  
  const sheet = getUsersSheet();
  const values = sheet.getDataRange().getValues();
  
  // วนลูปตรวจสอบข้อมูล (ข้ามแถวหัวตาราง)
  for (let i = 1; i < values.length; i++) {
    const dbUsername = String(values[i][0]).trim().toLowerCase();
    const dbPassword = String(values[i][1]).trim();
    const dbName = String(values[i][2]);
    
    if (dbUsername === String(username).trim().toLowerCase() && dbPassword === String(password).trim()) {
      return { 
        status: 'ok', 
        user: { 
          username: dbUsername, 
          name: dbName 
        } 
      };
    }
  }
  
  return { status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// ─── ฟังก์ชันเพิ่มตัวเลข 10 หลัก (ฟังก์ชัน 1) ──────────────────────
function addDigitNumber(username, digitNumber) {
  if (!username) {
    return { status: 'error', message: 'ไม่พบข้อมูลผู้ใช้ กรุณาเข้าสู่ระบบใหม่' };
  }
  
  // ตรวจสอบความถูกต้องของเลข 10 หลัก
  const numberStr = String(digitNumber).trim();
  const digitRegex = /^\d{10}$/;
  if (!digitRegex.test(numberStr)) {
    return { status: 'error', message: 'ข้อมูลต้องเป็นตัวเลข 10 หลักเท่านั้น' };
  }
  
  const sheet = getDataSheet();
  const nextId = sheet.getLastRow(); // แถวสุดท้ายคือ ID ปัจจุบัน
  
  // บันทึกวันและเวลาปัจจุบันในรูปแบบไทย
  const timezone = Session.getScriptTimeZone() || 'Asia/Bangkok';
  const timestampStr = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([
    nextId,
    username.toLowerCase(),
    "'" + numberStr,
    timestampStr
  ]);
  
  return { status: 'ok', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
}

// ─── ฟังก์ชันดึงประวัติข้อมูลเฉพาะของ User นั้นๆ ─────────────────
function getUserData(username) {
  if (!username) {
    return { status: 'error', message: 'ไม่พบชื่อผู้ใช้สำหรับการดึงประวัติ' };
  }
  
  const sheet = getDataSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { status: 'ok', data: [] };
  }
  
  const userRows = [];
  const searchUser = String(username).trim().toLowerCase();
  
  // ค้นหารายการจากท้ายตารางขึ้นไป เพื่อแสดงรายการล่าสุดก่อน
  for (let i = values.length - 1; i >= 1; i--) {
    const dbUsername = String(values[i][1]).trim().toLowerCase();
    if (dbUsername === searchUser) {
      let digitVal = String(values[i][2]).trim();
      if (digitVal.length === 9) {
        digitVal = '0' + digitVal;
      }
      userRows.push({
        id: values[i][0],
        digitNumber: digitVal,
        timestamp: values[i][3]
      });
    }
  }
  
  return { status: 'ok', data: userRows };
}
