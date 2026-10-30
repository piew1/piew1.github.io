/**
 * ============================================================================
 * ระบบยืม-คืนสินค้า (Loan/Return Management System)
 * Backend: Google Apps Script (Web App) + Google Sheets เป็นฐานข้อมูล
 * ----------------------------------------------------------------------------
 * วิธีใช้งานไฟล์นี้:
 * 1. สร้าง Google Sheet ใหม่ แล้วเปิด Extensions > Apps Script
 * 2. วางโค้ดทั้งหมดนี้ทับไฟล์ Code.gs เดิม
 * 3. รันฟังก์ชัน setupSheets() ด้วยตนเองหนึ่งครั้ง (เพื่อสร้างชีต/หัวคอลัมน์/ข้อมูลตัวอย่าง)
 * 4. Deploy เป็น Web App (Execute as "Me", Who has access "Anyone")
 * 5. นำ URL ของ Web App ไปใส่ในตัวแปร API_URL ที่ไฟล์ index.html
 * ----------------------------------------------------------------------------
 * หมายเหตุ: ห้ามตั้งชื่อฟังก์ชันขึ้นต้น/ลงท้ายด้วย underscore ตามข้อกำหนด
 * ============================================================================
 */

// ============================================================================
// ส่วนที่ 1: ทางเข้า Web App (doGet / doPost) และตัวจัดเส้นทาง (Router)
// ============================================================================

/**
 * doGet: ใช้สำหรับคำขอ GET เช่น ตรวจสอบสถานะ API และ PWA manifest
 * (ตาม requirement 4.1: หาก e.parameter.page === 'manifest' ให้คืนค่า Manifest JSON)
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    if (params.page === 'manifest') {
      return ContentService.createTextOutput(JSON.stringify(getManifestJson()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (!params.action) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'ระบบยืม-คืนสินค้า API พร้อมทำงาน' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var result = routeAction(params.action, params);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * doPost: ใช้สำหรับคำขอหลักของระบบทั้งหมด
 * Client ส่งมาแบบ Content-Type: text/plain (เพื่อเลี่ยง CORS preflight) โดย body เป็น JSON string
 */
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || (e && e.parameter && e.parameter.action);
    var result = routeAction(action, body);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** ตัวจัดเส้นทางกลางสำหรับทุก action ของระบบ */
function routeAction(action, payload) {
  payload = payload || {};
  switch (action) {
    // ---------- Auth ----------
    case 'login': return login(payload);
    case 'logout': return logout(payload);
    case 'changePassword': return changePassword(payload);
    case 'requestPasswordReset': return requestPasswordReset(payload);
    case 'resetPassword': return resetPassword(payload);
    case 'getCurrentUser': return getCurrentUser(payload);

    // ---------- สินค้า / คู่ค้า ----------
    case 'checkProduct': return checkProduct(payload);
    case 'getProducts': return getProducts(payload);
    case 'getPartners': return getPartners(payload);
    case 'importProducts': return importProducts(payload);

    // ---------- รายการยืม-คืน ----------
    case 'saveTransaction': return saveTransaction(payload);
    case 'searchTransactions': return searchTransactions(payload);
    case 'sendSearchNotification': return sendSearchResultsNotification(payload);
    case 'generateTransactionPdf': return generateTransactionPdf(payload);
    case 'getNotificationBadge': return getNotificationBadge(payload);

    // ---------- Dashboard ----------
    case 'getDashboardData': return getDashboardData(payload);
    case 'saveUserPreference': return saveUserPreference(payload);
    case 'getUserPreference': return getUserPreference(payload);

    // ---------- ตั้งค่า ----------
    case 'getPublicSettings': return getPublicSettings(payload);
    case 'getAppSettings': return getAppSettings(payload);
    case 'saveAppSettings': return saveAppSettings(payload);
    case 'testNotificationChannel': return testNotificationChannel(payload);
    case 'triggerAutoNotifications': return triggerAutoNotifications(payload);

    // ---------- ผู้ใช้งาน ----------
    case 'getUsers': return getUsers(payload);
    case 'saveUser': return saveUser(payload);
    case 'setUserStatus': return setUserStatus(payload);
    case 'deleteUser': return deleteUser(payload);
    case 'getActivityLogs': return getActivityLogs(payload);

    // ---------- Setup ----------
    case 'setupSheets': return setupSheets();

    default: return { success: false, message: 'ไม่รู้จัก action: ' + action };
  }
}

// ============================================================================
// ส่วนที่ 2: การตั้งค่าโครงสร้างชีต (Setup)
// ============================================================================

/**
 * setupSheets: สร้างชีตและหัวคอลัมน์อัตโนมัติถ้ายังไม่มี พร้อมข้อมูลตัวอย่าง
 * รันได้เองด้วยมือครั้งแรก และจะถูกเรียกซ้ำอัตโนมัติเมื่อระบบตรวจไม่พบชีตที่ต้องใช้
 */
function setupSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ===== ชีต Products (ทะเบียนสินค้า) =====
    var products = ss.getSheetByName('Products');
    if (!products) {
      products = ss.insertSheet('Products');
      products.appendRow(['รหัส', 'ชื่อสินค้า', 'ประเภท', 'จำนวน', 'หน่วย', 'QR']);
      // สูตรสร้าง QR Code อัตโนมัติจากรหัสสินค้าในคอลัมน์ A
      products.getRange('F1').setFormula('={"QR Code";ARRAYFORMULA(IF(A2:A="","",IMAGE("https://quickchart.io/qr?text="&A2:A)))}');
      [
        ['P001', 'เก้าอี้สำนักงาน', 'เฟอร์นิเจอร์', 20, 'ตัว'],
        ['P002', 'โต๊ะประชุม', 'เฟอร์นิเจอร์', 5, 'ตัว'],
        ['P003', 'โปรเจคเตอร์', 'อุปกรณ์ไฟฟ้า', 8, 'เครื่อง'],
        ['P004', 'สายไฟ HDMI', 'อุปกรณ์ไฟฟ้า', 50, 'เส้น'],
        ['P005', 'เครื่องพิมพ์เลเซอร์', 'อุปกรณ์สำนักงาน', 6, 'เครื่อง']
      ].forEach(function (row) { products.appendRow(row); });
      products.setFrozenRows(1);
    }

    // ===== ชีต Partners (ผู้ค้า/ลูกค้า) =====
    var partners = ss.getSheetByName('Partners');
    if (!partners) {
      partners = ss.insertSheet('Partners');
      partners.appendRow(['รหัส', 'ชื่อ', 'ประเภท']);
      [
        ['V001', 'บริษัท ทูลส์ไทย จำกัด', 'ผู้ค้า'],
        ['V002', 'ห้างหุ้นส่วน รุ่งเรืองวัสดุ', 'ผู้ค้า'],
        ['C001', 'บจก. สมชายก่อสร้าง', 'ลูกค้า'],
        ['C002', 'ร้านสมหญิงค้าวัสดุ', 'ลูกค้า'],
        ['C003', 'คุณประภาส ใจดี', 'ลูกค้า']
      ].forEach(function (row) { partners.appendRow(row); });
      partners.setFrozenRows(1);
    }

    // ===== ชีต Transactions (ประวัติรับคืน/เบิกออก) =====
    var transactions = ss.getSheetByName('Transactions');
    if (!transactions) {
      transactions = ss.insertSheet('Transactions');
      transactions.appendRow([
        'วันเวลาบันทึก', 'รหัสสินค้า', 'ชื่อสินค้า', 'ประเภทสินค้า', 'รับคืน-เบิกออก', 'จำนวน',
        'เอกสารอ้างอิง', 'ประเภทคู่ค้า', 'รหัสผู้ค้า-ลูกค้า', 'ชื่อผู้ค้า-ลูกค้า',
        'วันหมดอายุ/กำหนดคืน', 'หมายเหตุ', 'ผู้บันทึก', 'สภาพ'
      ]);
      transactions.setFrozenRows(1);
    }

    // ===== ชีต Settings (ค่าคอนฟิกระบบ key/value) =====
    var settings = ss.getSheetByName('Settings');
    if (!settings) {
      settings = ss.insertSheet('Settings');
      settings.appendRow(['key', 'value']);
      [
        ['companyName', 'บริษัท ตัวอย่าง จำกัด'],
        ['companyAddress', '123 ถนนตัวอย่าง แขวงตัวอย่าง เขตตัวอย่าง กรุงเทพฯ 10000'],
        ['logoUrl', ''],
        ['lineToken', ''],
        ['lineTarget', ''],
        ['telegramToken', ''],
        ['telegramChatId', ''],
        ['notifyEmail', ''],
        ['lowStockThreshold', '5'],
        ['nearDueDays', '7'],
        ['autoDailyNotify', 'false'],
        ['autoWeeklyNotify', 'false'],
        ['autoMonthlyNotify', 'false'],
        ['autoLowStockNotify', 'false']
      ].forEach(function (row) { settings.appendRow(row); });
      settings.setFrozenRows(1);
    }

    // ===== ชีต Users (บัญชีผู้ใช้งาน) =====
    var users = ss.getSheetByName('Users');
    if (!users) {
      users = ss.insertSheet('Users');
      users.appendRow([
        'UserID', 'Username', 'Email', 'PasswordHash', 'Salt', 'Role', 'Status',
        'FailedAttempts', 'LockoutUntil', 'ResetToken', 'ResetTokenExpiry',
        'MustChangePassword', 'LastLoginDate', 'CreatedDate'
      ]);
      // สร้างบัญชี admin เริ่มต้น: admin / Admin@1234 (ต้องบังคับเปลี่ยนรหัสผ่านทันทีที่ล็อกอินครั้งแรก)
      var salt = generateSalt();
      var hash = hashPassword('Admin@1234', salt);
      users.appendRow(['U0001', 'admin', 'admin@example.com', hash, salt, 'admin', 'Approved', 0, '', '', '', true, '', new Date()]);
      users.setFrozenRows(1);
    }

    // ===== ชีต ActivityLogs (บันทึกกิจกรรมทุกอย่าง) =====
    var logs = ss.getSheetByName('ActivityLogs');
    if (!logs) {
      logs = ss.insertSheet('ActivityLogs');
      logs.appendRow(['เวลา', 'ผู้ใช้', 'การกระทำ', 'รายละเอียด']);
      logs.setFrozenRows(1);
    }

    return { success: true, message: 'ตั้งค่าโครงสร้างชีตเรียบร้อยแล้ว (บัญชีเริ่มต้น admin / Admin@1234)' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการตั้งค่าชีต: ' + err.message };
  }
}

/** ตรวจสอบว่าชีตที่จำเป็นครบหรือไม่ ถ้าไม่ครบให้เรียก setupSheets() อัตโนมัติ */
function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var required = ['Products', 'Partners', 'Transactions', 'Settings', 'Users', 'ActivityLogs'];
  var missing = required.some(function (name) { return !ss.getSheetByName(name); });
  if (missing) setupSheets();
}

/** ดึงชีตแบบปลอดภัย ถ้าไม่พบจะพยายามสร้างใหม่อัตโนมัติ */
function getSheetSafe(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

// ============================================================================
// ส่วนที่ 3: ความปลอดภัย (Hashing / Session / Authorization)
// ============================================================================

/** สุ่ม salt แบบสุ่มมากกว่า 16 ไบต์ ต่อผู้ใช้ 1 คน */
function generateSalt() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

/** แปลง byte array เป็น hex string (จัดการค่า signed byte ของ Apps Script) */
function bytesToHex(bytes) {
  return bytes.map(function (b) {
    var v = b < 0 ? b + 256 : b;
    var h = v.toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}

/**
 * hashPassword: SHA-256 + salt พร้อม key stretching อย่างน้อย 1000 รอบ
 * เพื่อป้องกัน brute-force / rainbow table (bcrypt ใช้ไม่ได้ใน Apps Script)
 */
function hashPassword(password, salt) {
  var input = salt + String(password);
  var rounds = 1000;
  for (var i = 0; i < rounds; i++) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
    input = bytesToHex(digest) + salt;
  }
  var finalDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  return bytesToHex(finalDigest);
}

/** ตรวจสอบ session token จาก CacheService (TTL 1 ชั่วโมง) */
function validateSession(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/** ตรวจสิทธิ์: โยน error ถ้าไม่มี session หรือ role ไม่ตรงที่อนุญาต */
function requireRole(session, roles) {
  if (!session) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  if (roles && roles.indexOf(session.role) === -1) throw new Error('ไม่มีสิทธิ์เข้าถึงฟังก์ชันนี้');
}

/** บันทึก log กิจกรรมแบบเงียบ ไม่ให้กระทบ flow หลักหากล้มเหลว */
function logActivityInternal(username, action, detail) {
  try {
    var sheet = getSheetSafe('ActivityLogs');
    sheet.appendRow([new Date(), username || '', action || '', detail || '']);
  } catch (e) { /* ignore */ }
}

// ============================================================================
// ส่วนที่ 4: Authentication และการจัดการผู้ใช้
// ============================================================================

/** เข้าสู่ระบบ: ตรวจสถานะบัญชี, lockout, รหัสผ่าน แล้วออก session token (UUID) */
function login(payload) {
  try {
    ensureSheets();
    var email = (payload.email || '').toString().trim().toLowerCase();
    var password = (payload.password || '').toString();
    if (!email || !password) return { success: false, message: 'กรุณากรอกอีเมลและรหัสผ่าน' };

    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    var now = new Date();

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if ((row[idx.Email] || '').toString().trim().toLowerCase() !== email) continue;

      var lockoutUntil = row[idx.LockoutUntil];
      if (lockoutUntil && new Date(lockoutUntil) > now) {
        return { success: false, message: 'บัญชีถูกล็อกชั่วคราวจากการพยายามเข้าสู่ระบบผิดพลาด กรุณาลองใหม่ภายหลัง' };
      }
      if (row[idx.Status] !== 'Approved') {
        return { success: false, message: 'บัญชีนี้ยังไม่ได้รับการอนุมัติ หรือถูกระงับการใช้งาน' };
      }

      var salt = row[idx.Salt];
      var hash = hashPassword(password, salt);
      if (hash !== row[idx.PasswordHash]) {
        var attempts = Number(row[idx.FailedAttempts] || 0) + 1;
        sheet.getRange(r + 1, idx.FailedAttempts + 1).setValue(attempts);
        if (attempts >= 5) {
          var lockUntil = new Date(now.getTime() + 15 * 60000);
          sheet.getRange(r + 1, idx.LockoutUntil + 1).setValue(lockUntil);
          logActivityInternal(row[idx.Username], 'LOGIN_LOCKED', 'บัญชีถูกล็อก 15 นาที จากการพยายามเข้าสู่ระบบผิด 5 ครั้ง');
          return { success: false, message: 'กรอกรหัสผ่านผิดครบ 5 ครั้ง บัญชีถูกล็อก 15 นาที' };
        }
        return { success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
      }

      // เข้าสู่ระบบสำเร็จ: รีเซ็ตตัวนับ, บันทึกเวลา, ออก session token
      sheet.getRange(r + 1, idx.FailedAttempts + 1).setValue(0);
      sheet.getRange(r + 1, idx.LockoutUntil + 1).setValue('');
      sheet.getRange(r + 1, idx.LastLoginDate + 1).setValue(now);

      var token = Utilities.getUuid();
      var mustChange = row[idx.MustChangePassword] === true || row[idx.MustChangePassword] === 'TRUE' || row[idx.MustChangePassword] === 'true';
      var userInfo = {
        userId: row[idx.UserID],
        username: row[idx.Username],
        email: row[idx.Email],
        role: row[idx.Role],
        mustChangePassword: mustChange
      };
      CacheService.getScriptCache().put('session_' + token, JSON.stringify(userInfo), 3600); // TTL 1 ชั่วโมง
      logActivityInternal(userInfo.username, 'LOGIN', 'เข้าสู่ระบบสำเร็จ');
      return { success: true, token: token, user: userInfo };
    }
    return { success: false, message: 'ไม่พบบัญชีผู้ใช้นี้' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function buildHeaderIndex(headerRow) {
  var idx = {};
  headerRow.forEach(function (h, i) { idx[h] = i; });
  return idx;
}

/** ออกจากระบบ: ลบ session token ออกจาก CacheService ทันที */
function logout(payload) {
  try {
    if (payload.token) CacheService.getScriptCache().remove('session_' + payload.token);
    return { success: true, message: 'ออกจากระบบเรียบร้อย' };
  } catch (err) { return { success: false, message: err.message }; }
}

function getCurrentUser(payload) {
  var session = validateSession(payload.token);
  if (!session) return { success: false, message: 'เซสชันหมดอายุ' };
  return { success: true, user: session };
}

/** เปลี่ยนรหัสผ่าน (ใช้ตอนบังคับเปลี่ยนรหัสผ่านครั้งแรก หรือเปลี่ยนเองภายหลัง) */
function changePassword(payload) {
  try {
    var session = validateSession(payload.token);
    if (!session) return { success: false, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    for (var r = 1; r < data.length; r++) {
      if (data[r][idx.UserID] !== session.userId) continue;
      var salt = data[r][idx.Salt];
      var oldHash = hashPassword(payload.oldPassword || '', salt);
      if (oldHash !== data[r][idx.PasswordHash]) return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      if (!payload.newPassword || payload.newPassword.length < 8) return { success: false, message: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร' };
      var newSalt = generateSalt();
      sheet.getRange(r + 1, idx.PasswordHash + 1).setValue(hashPassword(payload.newPassword, newSalt));
      sheet.getRange(r + 1, idx.Salt + 1).setValue(newSalt);
      sheet.getRange(r + 1, idx.MustChangePassword + 1).setValue(false);
      logActivityInternal(session.username, 'CHANGE_PASSWORD', 'เปลี่ยนรหัสผ่านสำเร็จ');
      return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (err) { return { success: false, message: err.message }; }
}

/** ขอรีเซ็ตรหัสผ่าน: สร้าง ResetToken อายุ 1 ชั่วโมง และส่งทางอีเมลเท่านั้น (ไม่แสดงบนหน้าจอ) */
function requestPasswordReset(payload) {
  try {
    ensureSheets();
    var email = (payload.email || '').toString().trim().toLowerCase();
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    for (var r = 1; r < data.length; r++) {
      if ((data[r][idx.Email] || '').toString().trim().toLowerCase() !== email) continue;
      var token = Utilities.getUuid();
      var expiry = new Date(new Date().getTime() + 60 * 60 * 1000);
      sheet.getRange(r + 1, idx.ResetToken + 1).setValue(token);
      sheet.getRange(r + 1, idx.ResetTokenExpiry + 1).setValue(expiry);
      var body = 'รหัสสำหรับตั้งรหัสผ่านใหม่ของคุณคือ:\n\n' + token +
        '\n\nรหัสนี้จะหมดอายุภายใน 1 ชั่วโมง\nหากคุณไม่ได้ร้องขอ กรุณาเพิกเฉยต่ออีเมลฉบับนี้';
      MailApp.sendEmail(email, 'รีเซ็ตรหัสผ่าน - ระบบยืม-คืนสินค้า', body);
      logActivityInternal(data[r][idx.Username], 'REQUEST_RESET', 'ร้องขอรีเซ็ตรหัสผ่าน');
      return { success: true, message: 'ระบบได้ส่งรหัสสำหรับตั้งรหัสผ่านใหม่ไปยังอีเมลของคุณแล้ว' };
    }
    // ไม่บอกว่าพบหรือไม่พบอีเมล เพื่อความปลอดภัย
    return { success: true, message: 'หากอีเมลนี้มีอยู่ในระบบ จะได้รับอีเมลสำหรับรีเซ็ตรหัสผ่าน' };
  } catch (err) { return { success: false, message: err.message }; }
}

/** ตั้งรหัสผ่านใหม่ด้วย ResetToken */
function resetPassword(payload) {
  try {
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    var token = (payload.token || '').toString().trim();
    if (!token) return { success: false, message: 'กรุณากรอกรหัสยืนยัน' };
    for (var r = 1; r < data.length; r++) {
      if (data[r][idx.ResetToken] !== token) continue;
      var expiry = data[r][idx.ResetTokenExpiry];
      if (!expiry || new Date(expiry) < new Date()) return { success: false, message: 'รหัสหมดอายุแล้ว กรุณาขอรหัสใหม่' };
      if (!payload.newPassword || payload.newPassword.length < 8) return { success: false, message: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร' };
      var newSalt = generateSalt();
      sheet.getRange(r + 1, idx.PasswordHash + 1).setValue(hashPassword(payload.newPassword, newSalt));
      sheet.getRange(r + 1, idx.Salt + 1).setValue(newSalt);
      sheet.getRange(r + 1, idx.ResetToken + 1).setValue('');
      sheet.getRange(r + 1, idx.ResetTokenExpiry + 1).setValue('');
      sheet.getRange(r + 1, idx.MustChangePassword + 1).setValue(false);
      sheet.getRange(r + 1, idx.FailedAttempts + 1).setValue(0);
      sheet.getRange(r + 1, idx.LockoutUntil + 1).setValue('');
      logActivityInternal(data[r][idx.Username], 'RESET_PASSWORD', 'ตั้งรหัสผ่านใหม่ผ่านระบบลืมรหัสผ่าน');
      return { success: true, message: 'ตั้งรหัสผ่านใหม่สำเร็จ กรุณาเข้าสู่ระบบ' };
    }
    return { success: false, message: 'รหัสยืนยันไม่ถูกต้อง' };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
// ส่วนที่ 5: การจัดการผู้ใช้งาน (เฉพาะ admin)
// ============================================================================

function getUsers(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    var list = [];
    for (var r = 1; r < data.length; r++) {
      if (!data[r][idx.UserID]) continue;
      list.push({
        userId: data[r][idx.UserID], username: data[r][idx.Username], email: data[r][idx.Email],
        role: data[r][idx.Role], status: data[r][idx.Status],
        lastLogin: data[r][idx.LastLoginDate], createdDate: data[r][idx.CreatedDate]
      });
    }
    return { success: true, data: list };
  } catch (err) { return { success: false, message: err.message }; }
}

function saveUser(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    var username = (payload.username || '').toString().trim();
    var email = (payload.email || '').toString().trim().toLowerCase();
    var role = payload.role;
    if (!username || !email || !role) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบถ้วน' };
    if (['admin', 'staff', 'salesperson', 'manager'].indexOf(role) === -1) return { success: false, message: 'บทบาทไม่ถูกต้อง' };

    var existingRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (payload.userId && data[r][idx.UserID] === payload.userId) { existingRow = r; break; }
      if (!payload.userId && (data[r][idx.Email] || '').toString().toLowerCase() === email) { existingRow = r; break; }
    }

    if (existingRow >= 0) {
      sheet.getRange(existingRow + 1, idx.Username + 1).setValue(username);
      sheet.getRange(existingRow + 1, idx.Email + 1).setValue(email);
      sheet.getRange(existingRow + 1, idx.Role + 1).setValue(role);
      if (payload.status) sheet.getRange(existingRow + 1, idx.Status + 1).setValue(payload.status);
      if (payload.newPassword) {
        var salt = generateSalt();
        sheet.getRange(existingRow + 1, idx.PasswordHash + 1).setValue(hashPassword(payload.newPassword, salt));
        sheet.getRange(existingRow + 1, idx.Salt + 1).setValue(salt);
        sheet.getRange(existingRow + 1, idx.MustChangePassword + 1).setValue(true);
      }
      logActivityInternal(session.username, 'UPDATE_USER', 'แก้ไขผู้ใช้ ' + username);
      return { success: true, message: 'บันทึกผู้ใช้เรียบร้อย' };
    } else {
      var newId = 'U' + Utilities.getUuid().slice(0, 8);
      var tempPassword = payload.newPassword || ('Temp@' + Math.floor(1000 + Math.random() * 9000));
      var salt2 = generateSalt();
      sheet.appendRow([newId, username, email, hashPassword(tempPassword, salt2), salt2, role, 'Approved', 0, '', '', '', true, '', new Date()]);
      logActivityInternal(session.username, 'CREATE_USER', 'สร้างผู้ใช้ใหม่ ' + username + ' (' + role + ')');
      return { success: true, message: 'สร้างผู้ใช้สำเร็จ รหัสผ่านชั่วคราวคือ: ' + tempPassword, tempPassword: tempPassword };
    }
  } catch (err) { return { success: false, message: err.message }; }
}

function setUserStatus(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    for (var r = 1; r < data.length; r++) {
      if (data[r][idx.UserID] !== payload.userId) continue;
      sheet.getRange(r + 1, idx.Status + 1).setValue(payload.status);
      if (payload.status === 'Approved') {
        sheet.getRange(r + 1, idx.FailedAttempts + 1).setValue(0);
        sheet.getRange(r + 1, idx.LockoutUntil + 1).setValue('');
      }
      logActivityInternal(session.username, 'SET_USER_STATUS', 'เปลี่ยนสถานะผู้ใช้ ' + data[r][idx.Username] + ' เป็น ' + payload.status);
      return { success: true, message: 'อัปเดตสถานะเรียบร้อย' };
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (err) { return { success: false, message: err.message }; }
}

function deleteUser(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    if (payload.userId === session.userId) return { success: false, message: 'ไม่สามารถลบบัญชีตัวเองได้' };
    var sheet = getSheetSafe('Users');
    var data = sheet.getDataRange().getValues();
    var idx = buildHeaderIndex(data[0]);
    for (var r = 1; r < data.length; r++) {
      if (data[r][idx.UserID] !== payload.userId) continue;
      var uname = data[r][idx.Username];
      sheet.deleteRow(r + 1);
      logActivityInternal(session.username, 'DELETE_USER', 'ลบผู้ใช้ ' + uname);
      return { success: true, message: 'ลบผู้ใช้เรียบร้อย' };
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (err) { return { success: false, message: err.message }; }
}

function getActivityLogs(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var sheet = getSheetSafe('ActivityLogs');
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var r = Math.max(1, data.length - 500); r < data.length; r++) {
      if (!data[r][0]) continue;
      list.push({ time: data[r][0], user: data[r][1], action: data[r][2], detail: data[r][3] });
    }
    list.reverse();
    return { success: true, data: list };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
// ส่วนที่ 6: สินค้า (Products) และคู่ค้า (Partners)
// ============================================================================

/** checkProduct: ค้นหารหัสสินค้าในชีต Products (ใช้ตอนสแกน QR/Barcode) */
function checkProduct(payload) {
  try {
    ensureSheets();
    var code = (payload.code || '').toString().trim();
    if (!code) return { success: false, found: false, message: 'กรุณาระบุรหัสสินค้า' };
    var sheet = getSheetSafe('Products');
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if ((data[r][0] || '').toString().trim().toLowerCase() === code.toLowerCase()) {
        return {
          success: true, found: true, code: data[r][0], name: data[r][1],
          category: data[r][2], quantity: data[r][3], unit: data[r][4]
        };
      }
    }
    return { success: true, found: false, message: 'ไม่พบรหัสสินค้านี้ในระบบ' };
  } catch (err) { return { success: false, found: false, message: 'เกิดข้อผิดพลาด: ' + err.message }; }
}

/**
 * getProducts: สำหรับหน้า "ค้นหายอดคงเหลือ" ใช้ cache 30 วินาที
 * เพื่อรองรับการเรียกพร้อมกันจากผู้ใช้จำนวนมาก (โดยเฉพาะ Sales) ให้เร็วและลดภาระชีต
 */
function getProducts(payload) {
  try {
    ensureSheets();
    var query = (payload.query || '').toString().trim().toLowerCase();
    var cache = CacheService.getScriptCache();
    var cached = cache.get('products_list');
    var list;
    if (cached) {
      list = JSON.parse(cached);
    } else {
      var sheet = getSheetSafe('Products');
      var data = sheet.getDataRange().getValues();
      list = [];
      for (var r = 1; r < data.length; r++) {
        if (!data[r][0]) continue;
        list.push({ code: data[r][0], name: data[r][1], category: data[r][2], quantity: data[r][3], unit: data[r][4] });
      }
      cache.put('products_list', JSON.stringify(list), 30);
    }
    if (query) {
      list = list.filter(function (p) {
        return p.code.toString().toLowerCase().indexOf(query) !== -1 || p.name.toString().toLowerCase().indexOf(query) !== -1;
      });
    }
    return { success: true, data: list.slice(0, 200) };
  } catch (err) { return { success: false, message: err.message }; }
}

function getPartners(payload) {
  try {
    ensureSheets();
    var type = payload.type;
    var sheet = getSheetSafe('Partners');
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var r = 1; r < data.length; r++) {
      if (!data[r][0]) continue;
      if (type && data[r][2] !== type) continue;
      list.push({ code: data[r][0], name: data[r][1], type: data[r][2] });
    }
    return { success: true, data: list };
  } catch (err) { return { success: false, message: err.message }; }
}

/** importProducts: นำเข้าสินค้าจาก CSV/TXT (UTF-8) โดย admin เท่านั้น */
function importProducts(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var text = (payload.csvText || '').toString();
    if (!text.trim()) return { success: false, message: 'ไม่มีข้อมูลสำหรับนำเข้า' };
    var lines = text.split(/\r\n|\n/).filter(function (l) { return l.trim() !== ''; });
    var sheet = getSheetSafe('Products');
    var existing = sheet.getDataRange().getValues();
    var existingCodes = {};
    for (var e = 1; e < existing.length; e++) {
      if (existing[e][0]) existingCodes[existing[e][0].toString().trim().toLowerCase()] = e + 1;
    }
    var startIdx = 0;
    if (lines[0] && lines[0].toLowerCase().indexOf('รหัส') !== -1) startIdx = 1;
    var added = 0, updated = 0;
    for (var i = startIdx; i < lines.length; i++) {
      var cols = lines[i].indexOf('\t') !== -1 ? lines[i].split('\t') : lines[i].split(',');
      var code = (cols[0] || '').trim();
      if (!code) continue;
      var name = (cols[1] || '').trim();
      var category = (cols[2] || '').trim();
      var qty = Number(cols[3] || 0);
      var unit = (cols[4] || '').trim();
      var key = code.toLowerCase();
      if (existingCodes[key]) {
        sheet.getRange(existingCodes[key], 2, 1, 4).setValues([[name, category, qty, unit]]);
        updated++;
      } else {
        sheet.appendRow([code, name, category, qty, unit]);
        existingCodes[key] = sheet.getLastRow();
        added++;
      }
    }
    CacheService.getScriptCache().remove('products_list');
    logActivityInternal(session.username, 'IMPORT_PRODUCTS', 'นำเข้าสินค้า เพิ่ม ' + added + ' รายการ อัปเดต ' + updated + ' รายการ');
    return { success: true, message: 'นำเข้าข้อมูลสำเร็จ เพิ่มใหม่ ' + added + ' รายการ อัปเดต ' + updated + ' รายการ' };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
// ส่วนที่ 7: รายการยืม-คืน (Transactions)
// ============================================================================

/** saveTransaction: บันทึกใบยืม/คืน (รองรับหลายรายการสินค้าในใบเดียว) */
function saveTransaction(payload) {
  var lock = LockService.getScriptLock();
  try {
    ensureSheets();
    var session = validateSession(payload.token);
    requireRole(session, ['admin', 'staff']);

    var type = (payload.type || '').toString().trim();
    var referenceNo = (payload.referenceNo || '').toString().trim();
    var partnerType = (payload.partnerType || '').toString().trim();
    var partnerCode = (payload.partnerCode || '').toString().trim();
    var items = payload.items;

    // ---- ตรวจสอบฝั่ง Server (กันกรณี Client ถูกแก้ไข/ข้าม validation) ----
    if (['รับคืน', 'เบิกออก'].indexOf(type) === -1) return { success: false, message: 'กรุณาเลือกประเภทรายการ รับคืน หรือ เบิกออก' };
    if (!referenceNo) return { success: false, message: 'กรุณาระบุเอกสารอ้างอิง' };
    if (['ผู้ค้า', 'ลูกค้า'].indexOf(partnerType) === -1) return { success: false, message: 'กรุณาเลือกประเภทผู้ค้า-ลูกค้า' };
    if (!partnerCode) return { success: false, message: 'กรุณาระบุรหัสผู้ค้า-ลูกค้า' };
    if (!items || !items.length) return { success: false, message: 'กรุณาระบุรายการสินค้าอย่างน้อย 1 รายการ' };

    lock.waitLock(15000); // กันการเขียนชนกันเมื่อมีคนบันทึกพร้อมกัน

    var productsSheet = getSheetSafe('Products');
    var productData = productsSheet.getDataRange().getValues();
    var productMap = {};
    for (var p = 1; p < productData.length; p++) {
      if (!productData[p][0]) continue;
      productMap[productData[p][0].toString().trim().toLowerCase()] = {
        name: productData[p][1], category: productData[p][2], rowIndex: p + 1,
        quantity: Number(productData[p][3] || 0), unit: productData[p][4]
      };
    }

    var partnersSheet = getSheetSafe('Partners');
    var partnerData = partnersSheet.getDataRange().getValues();
    var partnerName = payload.partnerName || '';
    for (var pp = 1; pp < partnerData.length; pp++) {
      if ((partnerData[pp][0] || '').toString().trim().toLowerCase() === partnerCode.toLowerCase()) {
        partnerName = partnerData[pp][1];
        break;
      }
    }

    // ตรวจสอบรายการทั้งหมดก่อนเขียนจริง
    for (var chk = 0; chk < items.length; chk++) {
      var ci = items[chk];
      var ccode = (ci.code || '').toString().trim();
      if (!ccode) return { success: false, message: 'รายการที่ ' + (chk + 1) + ' ไม่มีรหัสสินค้า' };
      if (!productMap[ccode.toLowerCase()]) return { success: false, message: 'ไม่พบรหัสสินค้า ' + ccode + ' ในระบบ' };
      if (!Number(ci.qty) || Number(ci.qty) <= 0) return { success: false, message: 'จำนวนของรายการ ' + ccode + ' ต้องมากกว่า 0' };
    }

    var txSheet = getSheetSafe('Transactions');
    var now = new Date();
    var savedRows = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var code = item.code.toString().trim();
      var qty = Number(item.qty);
      var prod = productMap[code.toLowerCase()];
      var condition = item.condition || 'ปกติ';

      txSheet.appendRow([
        now, code, prod.name, prod.category, type, qty, referenceNo,
        partnerType, partnerCode, partnerName, payload.expiryDate || '', payload.note || '',
        session.username, condition
      ]);

      // ปรับจำนวนคงเหลือ: เบิกออก = ลดสต็อก, รับคืน = เพิ่มสต็อก
      var newQty = type === 'เบิกออก' ? prod.quantity - qty : prod.quantity + qty;
      productsSheet.getRange(prod.rowIndex, 4).setValue(newQty);
      prod.quantity = newQty;

      savedRows.push({ code: code, name: prod.name, category: prod.category, qty: qty, unit: prod.unit, condition: condition });
    }

    CacheService.getScriptCache().remove('products_list'); // ล้าง cache ยอดคงเหลือให้ทันสมัย
    logActivityInternal(session.username, 'SAVE_TRANSACTION', 'บันทึก' + type + ' เอกสาร ' + referenceNo + ' จำนวน ' + items.length + ' รายการ');

    return {
      success: true,
      message: 'บันทึกข้อมูลสำเร็จ',
      document: {
        timestamp: now, type: type, referenceNo: referenceNo, partnerType: partnerType,
        partnerCode: partnerCode, partnerName: partnerName, expiryDate: payload.expiryDate || '',
        note: payload.note || '', recordedBy: session.username, items: savedRows
      }
    };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

/** searchTransactions: ค้นหารายการยืม-คืนตามเงื่อนไขต่างๆ พร้อมผลรวมจำนวน */
function searchTransactions(payload) {
  try {
    ensureSheets();
    var session = validateSession(payload.token);
    requireRole(session, ['admin', 'staff', 'manager']);
    var sheet = getSheetSafe('Transactions');
    var data = sheet.getDataRange().getValues();
    var result = [];
    var dateFrom = payload.dateFrom ? new Date(payload.dateFrom) : null;
    var dateTo = payload.dateTo ? new Date(payload.dateTo) : null;
    if (dateTo) dateTo.setHours(23, 59, 59, 999);

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[1]) continue;
      var rowDate = new Date(row[0]);
      if (dateFrom && rowDate < dateFrom) continue;
      if (dateTo && rowDate > dateTo) continue;
      if (payload.type && row[4] !== payload.type) continue;
      if (payload.category && row[3] !== payload.category) continue;
      if (payload.partnerQuery) {
        var pq = payload.partnerQuery.toString().toLowerCase();
        if (row[8].toString().toLowerCase().indexOf(pq) === -1 && row[9].toString().toLowerCase().indexOf(pq) === -1) continue;
      }
      if (payload.productQuery) {
        var prq = payload.productQuery.toString().toLowerCase();
        if (row[1].toString().toLowerCase().indexOf(prq) === -1 && row[2].toString().toLowerCase().indexOf(prq) === -1) continue;
      }
      result.push({
        timestamp: row[0], code: row[1], name: row[2], category: row[3], type: row[4], qty: row[5],
        referenceNo: row[6], partnerType: row[7], partnerCode: row[8], partnerName: row[9],
        expiryDate: row[10], note: row[11], recordedBy: row[12], condition: row[13]
      });
    }
    var sum = result.reduce(function (s, x) { return s + Number(x.qty || 0); }, 0);
    return { success: true, data: result, sum: sum, count: result.length };
  } catch (err) { return { success: false, message: err.message }; }
}

/** ส่งผลการค้นหาไปแจ้งเตือน (LINE/Telegram ตามที่ตั้งค่าไว้) */
function sendSearchResultsNotification(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin', 'staff', 'manager']);
    var message = payload.message || 'ผลการค้นหารายการยืม-คืนจากระบบ';
    sendLineMessage(message);
    sendTelegramMessage(message);
    logActivityInternal(session.username, 'SEND_NOTIFICATION', 'ส่งผลการค้นหาไปแจ้งเตือน');
    return { success: true, message: 'ส่งการแจ้งเตือนเรียบร้อยแล้ว (ตรวจสอบว่าได้ตั้งค่าช่องทางไว้แล้ว)' };
  } catch (err) { return { success: false, message: err.message }; }
}

/** นับจำนวนรายการใกล้ครบกำหนดคืน สำหรับ Notification Badge บนเมนู */
function getNotificationBadge(payload) {
  try {
    var session = validateSession(payload.token);
    if (!session) return { success: false, message: 'เซสชันหมดอายุ' };
    ensureSheets();
    var settings = getSettingsMap();
    var nearDueDays = Number(settings.nearDueDays || 7);
    var sheet = getSheetSafe('Transactions');
    var data = sheet.getDataRange().getValues();
    var outstandingMap = {};
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[1]) continue;
      var key = row[1] + '|' + row[8] + '|' + row[6];
      if (!outstandingMap[key]) outstandingMap[key] = { qty: 0, expiryDate: row[10] };
      if (row[4] === 'เบิกออก') outstandingMap[key].qty += Number(row[5] || 0); else outstandingMap[key].qty -= Number(row[5] || 0);
      outstandingMap[key].expiryDate = row[10];
    }
    var now = new Date(); var count = 0;
    Object.keys(outstandingMap).forEach(function (k) {
      var o = outstandingMap[k];
      if (o.qty > 0 && o.expiryDate) {
        var due = new Date(o.expiryDate);
        if (!isNaN(due.getTime())) {
          var diff = Math.ceil((due - now) / 86400000);
          if (diff <= nearDueDays) count++;
        }
      }
    });
    return { success: true, count: count };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
// ส่วนที่ 8: สร้างใบยืม/คืน เป็น PDF (HTML Template แบบฝัง string ใน Code.gs)
// ============================================================================

/** คืนค่า HTML Template (string) สำหรับใบยืม/คืน ใช้ scriptlet <?= ?> ของ HtmlService (escape อัตโนมัติ) */
function buildPdfHtmlTemplate() {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  /* รองรับภาษาไทย: กำหนด font-family fallback หลายตัวไม่พึ่ง Google Fonts อย่างเดียว */
  body { font-family: 'Sarabun', 'Noto Sans Thai', 'Leelawadee UI', 'Tahoma', sans-serif; font-size: 14px; color: #222; padding: 24px; }
  .header { display:flex; align-items:center; border-bottom:2px solid #444; padding-bottom:10px; margin-bottom:16px; }
  .logo { height:60px; margin-right:16px; }
  .company h2 { margin:0; font-size:18px; }
  .company p { margin:2px 0; font-size:12px; color:#555; }
  h1 { text-align:center; font-size:20px; margin:10px 0 18px; letter-spacing:1px; }
  table { width:100%; border-collapse:collapse; margin-top:12px; }
  th, td { border:1px solid #888; padding:6px 8px; font-size:13px; text-align:left; }
  th { background:#eee; }
  .meta { display:flex; justify-content:space-between; margin-top:10px; font-size:13px; }
  .meta div { width:48%; line-height:1.6; }
  .signs { display:flex; justify-content:space-between; margin-top:70px; }
  .sign-box { width:45%; text-align:center; font-size:13px; }
  .sign-line { margin-top:50px; border-top:1px solid #333; padding-top:4px; }
</style>
</head>
<body>
  <div class="header">
    <? if (logoDataUri) { ?><img class="logo" src="<?= logoDataUri ?>"><? } ?>
    <div class="company"><h2><?= companyName ?></h2><p><?= companyAddress ?></p></div>
  </div>
  <h1><?= docType ?></h1>
  <div class="meta">
    <div>
      <b>เลขที่เอกสารอ้างอิง:</b> <?= referenceNo ?><br>
      <b>วันที่:</b> <?= dateStr ?><br>
      <b><?= partnerLabel ?>:</b> <?= partnerCode ?> - <?= partnerName ?>
    </div>
    <div style="text-align:right">
      <b>กำหนดคืน/วันหมดอายุ:</b> <?= expiryDate ?><br>
      <b>ผู้บันทึก:</b> <?= recordedBy ?>
    </div>
  </div>
  <table>
    <thead><tr><th>ลำดับ</th><th>รหัส</th><th>ชื่อสินค้า</th><th>จำนวน</th><th>หน่วย</th><th>สภาพ</th></tr></thead>
    <tbody>
    <? for (var i = 0; i < items.length; i++) { var it = items[i]; ?>
      <tr><td><?= i + 1 ?></td><td><?= it.code ?></td><td><?= it.name ?></td><td><?= it.qty ?></td><td><?= it.unit ?></td><td><?= it.condition ?></td></tr>
    <? } ?>
    </tbody>
  </table>
  <p><b>หมายเหตุ:</b> <?= note ?></p>
  <div class="signs">
    <? if (isLoan) { ?>
      <div class="sign-box">ลงชื่อ.............................................<div class="sign-line">(ผู้ยืม) ..................................... วันที่ ................</div></div>
      <div class="sign-box">ลงชื่อ.............................................<div class="sign-line">(ผู้อนุมัติ) ..................................... วันที่ ................</div></div>
    <? } else { ?>
      <div class="sign-box">ลงชื่อ.............................................<div class="sign-line">(ผู้คืน) ..................................... วันที่ ................</div></div>
      <div class="sign-box">ลงชื่อ.............................................<div class="sign-line">(ผู้รับคืน) ..................................... วันที่ ................</div></div>
    <? } ?>
  </div>
</body>
</html>`;
}

/** generateTransactionPdf: สร้างใบยืม/คืนเป็น PDF แล้วคืนเป็น base64 (admin, staff เท่านั้น) */
function generateTransactionPdf(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin', 'staff']);
    var doc = payload.document;
    if (!doc) return { success: false, message: 'ไม่พบข้อมูลเอกสาร' };

    var settings = getSettingsMap();
    var logoDataUri = '';
    if (settings.logoUrl) {
      try {
        var resp = UrlFetchApp.fetch(settings.logoUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          var blob = resp.getBlob();
          logoDataUri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
        }
      } catch (e) { /* ข้ามถ้าโหลดโลโก้ไม่สำเร็จ */ }
    }

    var template = HtmlService.createTemplate(buildPdfHtmlTemplate());
    template.docType = doc.type === 'เบิกออก' ? 'ใบยืม / เบิกสินค้า' : 'ใบคืนสินค้า';
    template.companyName = settings.companyName || '';
    template.companyAddress = settings.companyAddress || '';
    template.logoDataUri = logoDataUri;
    template.referenceNo = doc.referenceNo || '';
    template.dateStr = Utilities.formatDate(new Date(doc.timestamp || new Date()), Session.getScriptTimeZone() || 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    template.partnerLabel = doc.partnerType === 'ผู้ค้า' ? 'ผู้ค้า' : 'ลูกค้า';
    template.partnerCode = doc.partnerCode || '';
    template.partnerName = doc.partnerName || '';
    template.expiryDate = doc.expiryDate || '-';
    template.note = doc.note || '-';
    template.items = doc.items || [];
    template.recordedBy = doc.recordedBy || '';
    template.isLoan = doc.type === 'เบิกออก';

    var pdfBlob = template.evaluate().getBlob().getAs('application/pdf');
    var filename = (doc.type === 'เบิกออก' ? 'ใบยืม_' : 'ใบคืน_') + doc.referenceNo + '.pdf';
    var base64 = Utilities.base64Encode(pdfBlob.getBytes());

    logActivityInternal(session.username, 'GENERATE_PDF', 'พิมพ์ ' + template.docType + ' เอกสาร ' + doc.referenceNo);
    return { success: true, filename: filename, base64: base64 };
  } catch (err) {
    return { success: false, message: 'สร้าง PDF ไม่สำเร็จ: ' + err.message };
  }
}

// ============================================================================
// ส่วนที่ 9: Dashboard สำหรับผู้บริหาร (admin, manager เท่านั้น - manager อ่านอย่างเดียว)
// ============================================================================

function getDashboardData(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin', 'manager']);
    ensureSheets();
    var settings = getSettingsMap();
    var nearDueDays = Number(settings.nearDueDays || 7);
    var lowStockThreshold = Number(settings.lowStockThreshold || 5);

    var sheet = getSheetSafe('Transactions');
    var data = sheet.getDataRange().getValues();
    var dateFrom = payload.dateFrom ? new Date(payload.dateFrom) : new Date(new Date().setDate(new Date().getDate() - 30));
    var dateTo = payload.dateTo ? new Date(payload.dateTo) : new Date();
    dateTo.setHours(23, 59, 59, 999);
    var periodLength = dateTo.getTime() - dateFrom.getTime();
    var prevFrom = new Date(dateFrom.getTime() - periodLength - 1);
    var prevTo = new Date(dateFrom.getTime() - 1);
    var query = (payload.query || '').toString().toLowerCase();

    var current = [], previous = [], allRows = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[1]) continue;
      var rowDate = new Date(row[0]);
      var record = {
        timestamp: row[0], code: row[1], name: row[2], category: row[3], type: row[4], qty: Number(row[5] || 0),
        referenceNo: row[6], partnerType: row[7], partnerCode: row[8], partnerName: row[9],
        expiryDate: row[10], note: row[11], recordedBy: row[12], condition: row[13]
      };
      if (query) {
        var hay = (record.name + ' ' + record.code + ' ' + record.partnerName + ' ' + record.partnerCode).toLowerCase();
        if (hay.indexOf(query) === -1) continue;
      }
      allRows.push(record);
      if (rowDate >= dateFrom && rowDate <= dateTo) current.push(record);
      else if (rowDate >= prevFrom && rowDate <= prevTo) previous.push(record);
    }

    function summarize(rows) {
      var loanDocs = {}, returnDocs = {}, loanQty = 0, returnQty = 0, damaged = 0;
      rows.forEach(function (x) {
        if (x.type === 'เบิกออก') { loanDocs[x.referenceNo] = 1; loanQty += x.qty; }
        else { returnDocs[x.referenceNo] = 1; returnQty += x.qty; }
        if (x.condition && x.condition !== 'ปกติ') damaged += x.qty;
      });
      return { loanDocCount: Object.keys(loanDocs).length, returnDocCount: Object.keys(returnDocs).length, loanQty: loanQty, returnQty: returnQty, damaged: damaged };
    }
    var curSum = summarize(current), prevSum = summarize(previous);

    // ยอดค้างยืม: รวม เบิกออก - รับคืน ต่อ (รหัส+คู่ค้า+เอกสารอ้างอิง)
    var outstandingMap = {};
    allRows.forEach(function (x) {
      var key = x.code + '|' + x.partnerCode + '|' + x.referenceNo;
      if (!outstandingMap[key]) outstandingMap[key] = { qty: 0, expiryDate: x.expiryDate, name: x.name, code: x.code, partnerName: x.partnerName, referenceNo: x.referenceNo, timestamp: x.timestamp };
      if (x.type === 'เบิกออก') outstandingMap[key].qty += x.qty; else outstandingMap[key].qty -= x.qty;
      if (new Date(x.timestamp) > new Date(outstandingMap[key].timestamp)) { outstandingMap[key].expiryDate = x.expiryDate; outstandingMap[key].timestamp = x.timestamp; }
    });
    var outstandingList = Object.keys(outstandingMap).map(function (k) { return outstandingMap[k]; }).filter(function (o) { return o.qty > 0; });
    var outstandingQtyTotal = outstandingList.reduce(function (s, o) { return s + o.qty; }, 0);

    var now = new Date(), nearDue = [], overdue = [], overdueDaysSum = 0, overdueCount = 0;
    outstandingList.forEach(function (o) {
      if (!o.expiryDate) return;
      var due = new Date(o.expiryDate);
      if (isNaN(due.getTime())) return;
      var diffDays = Math.ceil((due.getTime() - now.getTime()) / 86400000);
      if (diffDays < 0) { overdue.push(Object.assign({}, o, { overdueDays: -diffDays })); overdueDaysSum += (-diffDays); overdueCount++; }
      else if (diffDays <= nearDueDays) nearDue.push(Object.assign({}, o, { dueInDays: diffDays }));
    });

    // อัตราคืนตรงเวลา: เทียบวันที่คืนจริงกับกำหนดคืนของรายการเบิกออกคู่กัน
    var loanByKey = {};
    allRows.forEach(function (x) { if (x.type === 'เบิกออก') loanByKey[x.code + '|' + x.partnerCode + '|' + x.referenceNo] = x; });
    var onTimeReturnCount = 0, totalReturnWithDueCount = 0;
    allRows.forEach(function (x) {
      if (x.type !== 'รับคืน') return;
      var loan = loanByKey[x.code + '|' + x.partnerCode + '|' + x.referenceNo];
      if (loan && loan.expiryDate) {
        totalReturnWithDueCount++;
        if (new Date(x.timestamp) <= new Date(loan.expiryDate)) onTimeReturnCount++;
      }
    });
    var onTimeRate = totalReturnWithDueCount ? Math.round((onTimeReturnCount / totalReturnWithDueCount) * 1000) / 10 : null;

    var productAgg = {}, partnerAgg = {};
    current.forEach(function (x) {
      if (x.type !== 'เบิกออก') return;
      productAgg[x.name] = (productAgg[x.name] || 0) + x.qty;
      partnerAgg[x.partnerName] = (partnerAgg[x.partnerName] || 0) + x.qty;
    });
    function topN(obj, n) {
      return Object.keys(obj).map(function (k) { return { name: k, qty: obj[k] }; }).sort(function (a, b) { return b.qty - a.qty; }).slice(0, n);
    }

    var productsSheet = getSheetSafe('Products');
    var productsData = productsSheet.getDataRange().getValues();
    var lowStock = [];
    for (var pr = 1; pr < productsData.length; pr++) {
      if (!productsData[pr][0]) continue;
      var qty = Number(productsData[pr][3] || 0);
      if (qty <= lowStockThreshold) lowStock.push({ code: productsData[pr][0], name: productsData[pr][1], quantity: qty, unit: productsData[pr][4] });
    }

    function pctChange(cur, prev) { if (!prev) return cur ? 100 : 0; return Math.round(((cur - prev) / prev) * 1000) / 10; }

    var kpi = {
      loanDocCount: curSum.loanDocCount, loanDocCountChange: pctChange(curSum.loanDocCount, prevSum.loanDocCount),
      returnDocCount: curSum.returnDocCount, returnDocCountChange: pctChange(curSum.returnDocCount, prevSum.returnDocCount),
      outstandingQty: outstandingQtyTotal,
      onTimeRate: onTimeRate,
      overdueCount: overdueCount,
      overdueAvgDays: overdueCount ? Math.round((overdueDaysSum / overdueCount) * 10) / 10 : 0,
      damagedOrLost: curSum.damaged,
      lowStockCount: lowStock.length,
      outOfStockQty: outstandingQtyTotal
    };

    return {
      success: true, kpi: kpi, nearDue: nearDue.slice(0, 50), overdue: overdue.slice(0, 50),
      topProducts: topN(productAgg, 10), topPartners: topN(partnerAgg, 10),
      lowStock: lowStock.slice(0, 50), transactions: current
    };
  } catch (err) { return { success: false, message: err.message }; }
}

/** บันทึกการตั้งค่ากราฟ/การแสดงผลต่อผู้ใช้ (localStorage ฝั่ง client + สำรองในชีต Settings) */
function saveUserPreference(payload) {
  try {
    var session = validateSession(payload.token);
    if (!session) return { success: false, message: 'เซสชันหมดอายุ' };
    var key = 'pref_' + session.username + '_' + (payload.key || 'default');
    setSettingValue(key, JSON.stringify(payload.value || {}));
    return { success: true, message: 'บันทึกการตั้งค่าเรียบร้อย' };
  } catch (err) { return { success: false, message: err.message }; }
}

function getUserPreference(payload) {
  try {
    var session = validateSession(payload.token);
    if (!session) return { success: false, message: 'เซสชันหมดอายุ' };
    var key = 'pref_' + session.username + '_' + (payload.key || 'default');
    var map = getSettingsMap();
    var val = map[key];
    return { success: true, value: val ? JSON.parse(val) : null };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
// ส่วนที่ 10: ตั้งค่าระบบ (Settings) และการแจ้งเตือน
// ============================================================================

function getSettingsMap() {
  ensureSheets();
  var sheet = getSheetSafe('Settings');
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    map[data[r][0]] = data[r][1];
  }
  return map;
}

function setSettingValue(key, value) {
  var sheet = getSheetSafe('Settings');
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] === key) { sheet.getRange(r + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

/** การตั้งค่าสาธารณะ (ชื่อบริษัท/โลโก้) ใช้แสดงหัวเว็บได้โดยไม่ต้อง login */
function getPublicSettings(payload) {
  try {
    var map = getSettingsMap();
    return { success: true, data: { companyName: map.companyName || '', logoUrl: map.logoUrl || '' } };
  } catch (err) { return { success: false, message: err.message }; }
}

function getAppSettings(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    return { success: true, data: getSettingsMap() };
  } catch (err) { return { success: false, message: err.message }; }
}

function saveAppSettings(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var settings = payload.settings || {};
    Object.keys(settings).forEach(function (key) { setSettingValue(key, settings[key]); });
    logActivityInternal(session.username, 'SAVE_SETTINGS', 'บันทึกการตั้งค่าระบบ');
    return { success: true, message: 'บันทึกการตั้งค่าเรียบร้อย' };
  } catch (err) { return { success: false, message: err.message }; }
}

function testNotificationChannel(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var channel = payload.channel;
    var message = 'ทดสอบการแจ้งเตือนจากระบบยืม-คืนสินค้า เวลา ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    var result;
    if (channel === 'line') result = sendLineMessage(message);
    else if (channel === 'telegram') result = sendTelegramMessage(message);
    else if (channel === 'email') result = sendEmailMessage('ทดสอบการแจ้งเตือน', message);
    else return { success: false, message: 'ไม่รู้จักช่องทางแจ้งเตือน' };
    return result;
  } catch (err) { return { success: false, message: err.message }; }
}

/** เปิดใช้งาน Trigger อัตโนมัติ (รายวัน/สัปดาห์/เดือน/สต็อกต่ำ) - admin เท่านั้น */
function triggerAutoNotifications(payload) {
  try {
    var session = validateSession(payload.token);
    requireRole(session, ['admin']);
    var handlers = ['runDailyNotification', 'runWeeklyNotification', 'runMonthlyNotification', 'runLowStockCheck'];
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('runDailyNotification').timeBased().everyDays(1).atHour(8).create();
    ScriptApp.newTrigger('runWeeklyNotification').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
    ScriptApp.newTrigger('runMonthlyNotification').timeBased().onMonthDay(1).atHour(8).create();
    ScriptApp.newTrigger('runLowStockCheck').timeBased().everyHours(6).create();
    logActivityInternal(session.username, 'SETUP_TRIGGERS', 'เปิดใช้งาน trigger แจ้งเตือนอัตโนมัติ');
    return { success: true, message: 'ตั้งค่าแจ้งเตือนอัตโนมัติเรียบร้อยแล้ว' };
  } catch (err) { return { success: false, message: err.message }; }
}

/** ส่งข้อความผ่าน LINE Messaging API (Push Message) */
function sendLineMessage(message) {
  try {
    var settings = getSettingsMap();
    if (!settings.lineToken || !settings.lineTarget) return { success: false, message: 'ยังไม่ได้ตั้งค่า LINE Token/Target' };
    var url = 'https://api.line.me/v2/bot/message/push';
    var payload = { to: settings.lineTarget, messages: [{ type: 'text', text: message }] };
    var options = {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + settings.lineToken },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    };
    var resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() === 200) return { success: true, message: 'ส่ง LINE สำเร็จ' };
    return { success: false, message: 'ส่ง LINE ไม่สำเร็จ: ' + resp.getContentText() };
  } catch (err) { return { success: false, message: 'LINE error: ' + err.message }; }
}

/** ส่งข้อความผ่าน Telegram Bot API */
function sendTelegramMessage(message) {
  try {
    var settings = getSettingsMap();
    if (!settings.telegramToken || !settings.telegramChatId) return { success: false, message: 'ยังไม่ได้ตั้งค่า Telegram Token/Chat ID' };
    var url = 'https://api.telegram.org/bot' + settings.telegramToken + '/sendMessage';
    var options = { method: 'post', contentType: 'application/x-www-form-urlencoded', payload: { chat_id: settings.telegramChatId, text: message }, muteHttpExceptions: true };
    var resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() === 200) return { success: true, message: 'ส่ง Telegram สำเร็จ' };
    return { success: false, message: 'ส่ง Telegram ไม่สำเร็จ: ' + resp.getContentText() };
  } catch (err) { return { success: false, message: 'Telegram error: ' + err.message }; }
}

/** ส่งอีเมลแจ้งเตือนผ่าน MailApp */
function sendEmailMessage(subject, body) {
  try {
    var settings = getSettingsMap();
    if (!settings.notifyEmail) return { success: false, message: 'ยังไม่ได้ตั้งค่าอีเมลผู้รับแจ้งเตือน' };
    MailApp.sendEmail(settings.notifyEmail, subject, body);
    return { success: true, message: 'ส่งอีเมลสำเร็จ' };
  } catch (err) { return { success: false, message: 'Email error: ' + err.message }; }
}

function buildPeriodSummaryMessage(sinceDate) {
  var sheet = getSheetSafe('Transactions');
  var data = sheet.getDataRange().getValues();
  var loan = 0, ret = 0;
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    var d = new Date(data[r][0]);
    if (d >= sinceDate) { if (data[r][4] === 'เบิกออก') loan += Number(data[r][5] || 0); else ret += Number(data[r][5] || 0); }
  }
  return 'สรุปยอดเบิกออก ' + loan + ' ชิ้น / รับคืน ' + ret + ' ชิ้น';
}

/** งานอัตโนมัติ: ตั้งค่าผ่าน Time-driven Trigger โดย triggerAutoNotifications() */
function runDailyNotification() {
  var settings = getSettingsMap();
  if (settings.autoDailyNotify !== 'true') return;
  var since = new Date(); since.setHours(0, 0, 0, 0);
  var msg = 'สรุปรายวัน: ' + buildPeriodSummaryMessage(since);
  sendLineMessage(msg); sendTelegramMessage(msg); sendEmailMessage('สรุปรายวัน - ระบบยืม-คืนสินค้า', msg);
}

function runWeeklyNotification() {
  var settings = getSettingsMap();
  if (settings.autoWeeklyNotify !== 'true') return;
  var since = new Date(); since.setDate(since.getDate() - 7);
  var msg = 'สรุปรายสัปดาห์: ' + buildPeriodSummaryMessage(since);
  sendLineMessage(msg); sendTelegramMessage(msg); sendEmailMessage('สรุปรายสัปดาห์ - ระบบยืม-คืนสินค้า', msg);
}

function runMonthlyNotification() {
  var settings = getSettingsMap();
  if (settings.autoMonthlyNotify !== 'true') return;
  var since = new Date(); since.setMonth(since.getMonth() - 1);
  var msg = 'สรุปรายเดือน: ' + buildPeriodSummaryMessage(since);
  sendLineMessage(msg); sendTelegramMessage(msg); sendEmailMessage('สรุปรายเดือน - ระบบยืม-คืนสินค้า', msg);
}

function runLowStockCheck() {
  var settings = getSettingsMap();
  if (settings.autoLowStockNotify !== 'true') return;
  var threshold = Number(settings.lowStockThreshold || 5);
  var sheet = getSheetSafe('Products');
  var data = sheet.getDataRange().getValues();
  var low = [];
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] && Number(data[r][3] || 0) <= threshold) low.push(data[r][0] + ' - ' + data[r][1] + ' (คงเหลือ ' + data[r][3] + ' ' + data[r][4] + ')');
  }
  if (low.length) {
    var msg = 'แจ้งเตือนสินค้าคงเหลือต่ำกว่ากำหนด:\n' + low.join('\n');
    sendLineMessage(msg); sendTelegramMessage(msg); sendEmailMessage('แจ้งเตือนสินค้าใกล้หมด - ระบบยืม-คืนสินค้า', msg);
  }
}

// ============================================================================
// ส่วนที่ 11: PWA Manifest (routing ผ่าน doGet แทนไฟล์ manifest.json แยก)
// ============================================================================

function getManifestJson() {
  // ไอคอนสร้างเป็น SVG data URI ภายใน เพื่อไม่ต้องพึ่งไฟล์ภาพแยกต่างหาก
  var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512">' +
    '<rect width="512" height="512" rx="90" fill="%23b79ced"/>' +
    '<text x="50%" y="58%" font-size="260" text-anchor="middle" fill="white" font-family="sans-serif">\uD83D\uDCE6</text></svg>';
  var dataUri = 'data:image/svg+xml,' + svg;
  return {
    name: 'ระบบยืม-คืนสินค้า',
    short_name: 'ยืม-คืน',
    start_url: '.',
    display: 'standalone',
    background_color: '#f6f1ff',
    theme_color: '#b79ced',
    icons: [
      { src: dataUri, sizes: '192x192', type: 'image/svg+xml' },
      { src: dataUri, sizes: '512x512', type: 'image/svg+xml' }
    ]
  };
}
