/**
 * ============================================================================
 *  ระบบยืม-คืนสินค้า (Borrow / Return System) - Backend
 *  Google Apps Script + Google Sheets (ทำหน้าที่เป็น REST-like JSON API)
 * ============================================================================
 *  วิธีใช้งานไฟล์นี้: ดูขั้นตอน Deploy แบบละเอียดในคำตอบ/README ที่แนบมาด้วย
 *
 *  โครงสร้างชีตที่ใช้งาน (สร้างอัตโนมัติด้วย setupSheets()):
 *   - Products      : รหัส | ชื่อสินค้า | ประเภท | จำนวน | หน่วย | QR
 *   - Partners      : รหัส | ชื่อ | ประเภท
 *   - Transactions  : วันเวลาบันทึก | รหัสสินค้า | ชื่อสินค้า | ประเภทสินค้า |
 *                     รับคืน-เบิกออก | จำนวน | เอกสารอ้างอิง | ประเภทคู่ค้า |
 *                     รหัสผู้ค้า-ลูกค้า | ชื่อผู้ค้า-ลูกค้า | วันหมดอายุ |
 *                     หมายเหตุ | ผู้บันทึก | เลขที่ใบ
 *   - Settings      : key | value
 *   - Users         : UserID | Username | Email | PasswordHash | Salt | Role |
 *                     Status | LastLoginDate | FailedAttempts | LockoutUntil |
 *                     ResetToken | ResetTokenExpiry | ForceChangePassword | CreatedDate
 *   - ActivityLogs  : เวลา | UserID | ผู้ใช้ | กิจกรรม | รายละเอียด
 *
 *  สิทธิ์การใช้งาน (บังคับที่ Server ทุกจุด):
 *   - admin        : ใช้งานได้ทุกฟังก์ชัน / จัดการผู้ใช้ / ตั้งค่าระบบ
 *   - staff        : สแกน/บันทึกรายการยืม-คืน, พิมพ์ PDF, ค้นหารายการ, ค้นหาคงเหลือ
 *   - salesperson  : ค้นหายอดคงเหลือ เท่านั้น
 *   - manager      : ค้นหารายการยืม-คืน, ค้นหายอดคงเหลือ, Dashboard (อ่านอย่างเดียว)
 * ============================================================================
 */

// ========================= ค่าคงที่หลักของระบบ =========================
const SHEET_PRODUCTS = 'Products';
const SHEET_PARTNERS = 'Partners';
const SHEET_TRANSACTIONS = 'Transactions';
const SHEET_SETTINGS = 'Settings';
const SHEET_USERS = 'Users';
const SHEET_LOGS = 'ActivityLogs';

const PRODUCTS_HEADERS = ['รหัส', 'ชื่อสินค้า', 'ประเภท', 'จำนวน', 'หน่วย', 'QR'];
const PARTNERS_HEADERS = ['รหัส', 'ชื่อ', 'ประเภท'];
const TRANSACTIONS_HEADERS = ['วันเวลาบันทึก', 'รหัสสินค้า', 'ชื่อสินค้า', 'ประเภทสินค้า', 'รับคืน-เบิกออก', 'จำนวน', 'เอกสารอ้างอิง', 'ประเภทคู่ค้า', 'รหัสผู้ค้า-ลูกค้า', 'ชื่อผู้ค้า-ลูกค้า', 'วันหมดอายุ', 'หมายเหตุ', 'ผู้บันทึก', 'เลขที่ใบ'];
const USERS_HEADERS = ['UserID', 'Username', 'Email', 'PasswordHash', 'Salt', 'Role', 'Status', 'LastLoginDate', 'FailedAttempts', 'LockoutUntil', 'ResetToken', 'ResetTokenExpiry', 'ForceChangePassword', 'CreatedDate'];
const LOGS_HEADERS = ['เวลา', 'UserID', 'ผู้ใช้', 'กิจกรรม', 'รายละเอียด'];

const SESSION_TTL = 60 * 60;      // อายุ session token ใน CacheService (วินาที) = 1 ชั่วโมง
const MAX_FAILED_ATTEMPTS = 5;    // จำนวนครั้งที่ล็อกอินผิดได้ก่อนถูกล็อก
const LOCKOUT_MINUTES = 15;       // ระยะเวลาล็อกบัญชี (นาที)
const HASH_ROUNDS = 1000;         // จำนวนรอบ key-stretching ของการแฮชรหัสผ่าน

// ============================================================================
//  ส่วนที่ 1: Web App Entry Point (doGet / doPost) + Router
// ============================================================================

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    // เส้นทางสำหรับ PWA manifest.json (ตามข้อกำหนด ห้ามสร้างไฟล์ manifest.json จริง)
    if (params.page === 'manifest') {
      return ContentService.createTextOutput(JSON.stringify(getManifest()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (params.action) {
      const result = routeAction(params.action, params, params.token);
      return jsonOutput(result);
    }
    return jsonOutput({ success: true, message: 'ระบบยืม-คืนสินค้า API พร้อมใช้งาน' });
  } catch (err) {
    return jsonOutput({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

function doPost(e) {
  try {
    let body = {};
    // ฝั่ง client ส่งเป็น text/plain (JSON string) เพื่อเลี่ยง CORS preflight
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = body.action;
    const data = body.data || {};
    const token = body.token || data.token;
    const result = routeAction(action, data, token);
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ตัวกลางกระจายคำสั่ง (action) ไปยังฟังก์ชันที่เกี่ยวข้องทั้งหมด
function routeAction(action, data, token) {
  data = data || {};
  token = token || data.token;
  switch (action) {
    // ---------- Auth ----------
    case 'login': return loginUser(data.username, data.password);
    case 'logout': return logoutUser(token);
    case 'checkSession': return validateSessionResponse(token);
    case 'changePassword': return changePassword(token, data.oldPassword, data.newPassword);
    case 'forgotPassword': return forgotPassword(data.email);
    case 'resetPassword': return resetPassword(data);
    // ---------- สแกน/บันทึกรายการยืม-คืน ----------
    case 'checkProduct': return checkProduct(data.code, token);
    case 'saveTransaction': return saveTransaction(data, token);
    case 'getPartners': return getPartners(data.type, token);
    case 'generatePDF': return generateTransactionPDF(data, token);
    // ---------- ค้นหา ----------
    case 'searchTransactions': return searchTransactions(data, token);
    case 'getStockBalance': return getStockBalance(data.query, token);
    case 'sendSearchResultNotification': return sendSearchResultNotification(data, token);
    // ---------- Dashboard ----------
    case 'getDashboardData': return getDashboardData(data, token);
    case 'getDashboardPrefs': return getDashboardPrefs(token);
    case 'saveDashboardPrefs': return saveDashboardPrefs(data, token);
    // ---------- แจ้งเตือน / Badge ----------
    case 'getNotificationBadge': return getNotificationBadgeCount(token);
    case 'testNotification': return testNotification(data.channel, token);
    // ---------- ตั้งค่าระบบ ----------
    case 'getSettings': return getSettingsForClient(token);
    case 'saveSettings': return saveSettings(data, token);
    // ---------- จัดการผู้ใช้ ----------
    case 'listUsers': return listUsers(token);
    case 'createUser': return createUserAccount(data, token);
    case 'updateUser': return updateUserAccount(data, token);
    case 'setUserStatus': return setUserStatus(data.userId, data.status, token);
    // ---------- Activity Logs / Import ----------
    case 'getActivityLogs': return getActivityLogs(data, token);
    case 'importProducts': return importProductsCSV(data.csv, token);
    default: return { success: false, message: 'ไม่รู้จักคำสั่ง (action="' + action + '")' };
  }
}

function getManifest() {
  return {
    name: 'ระบบยืม-คืนสินค้า',
    short_name: 'ยืม-คืน',
    start_url: '.',
    scope: '.',
    display: 'standalone',
    background_color: '#f3eefc',
    theme_color: '#9b8cf2',
    lang: 'th',
    icons: [
      { src: 'https://cdn-icons-png.flaticon.com/192/3144/3144456.png', sizes: '192x192', type: 'image/png' },
      { src: 'https://cdn-icons-png.flaticon.com/512/3144/3144456.png', sizes: '512x512', type: 'image/png' }
    ]
  };
}

// ============================================================================
//  ส่วนที่ 2: Sheet Helper / setupSheets()
// ============================================================================

function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }

// ดึงชีตตามชื่อ ถ้ายังไม่มีให้สร้างใหม่พร้อมหัวคอลัมน์ (เรียกอัตโนมัติทุกครั้งที่ไม่พบชีต)
function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function headerIndexMap(headers) {
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });
  return map;
}

/**
 * setupSheets(): สร้างชีตทั้งหมดพร้อมหัวคอลัมน์ + ข้อมูลตัวอย่าง + ผู้ใช้ admin เริ่มต้น
 * รันเองครั้งแรกได้จากตัวแก้ไขสคริปต์ (เลือกฟังก์ชันนี้แล้วกด Run)
 */
function setupSheets() {
  try {
    // ---- Products (พร้อมสูตร QR Code ในคอลัมน์ F) ----
    const products = getOrCreateSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    if (products.getLastRow() < 2) {
      products.getRange('F1').setFormula('={"QR Code";ARRAYFORMULA(IF(A2:A="","",IMAGE("https://quickchart.io/qr?text="&A2:A)))}');
      products.getRange(2, 1, 5, 5).setValues([
        ['P001', 'เครื่องพิมพ์ Laser', 'อุปกรณ์สำนักงาน', 10, 'เครื่อง'],
        ['P002', 'โน้ตบุ๊ก Dell', 'อุปกรณ์ไอที', 15, 'เครื่อง'],
        ['P003', 'เก้าอี้สำนักงาน', 'เฟอร์นิเจอร์', 30, 'ตัว'],
        ['P004', 'โต๊ะประชุม', 'เฟอร์นิเจอร์', 8, 'ตัว'],
        ['P005', 'โปรเจคเตอร์', 'อุปกรณ์ไอที', 5, 'เครื่อง']
      ]);
    }
    // ---- Partners ----
    const partners = getOrCreateSheet(SHEET_PARTNERS, PARTNERS_HEADERS);
    if (partners.getLastRow() < 2) {
      partners.getRange(2, 1, 5, 3).setValues([
        ['C001', 'บริษัท เอบีซี จำกัด', 'ลูกค้า'],
        ['C002', 'ร้านค้าปลีกไทย', 'ลูกค้า'],
        ['V001', 'บริษัท ซัพพลาย จำกัด', 'ผู้ค้า'],
        ['V002', 'ห้างหุ้นส่วนไอที', 'ผู้ค้า'],
        ['C003', 'คุณสมชาย ใจดี', 'ลูกค้า']
      ]);
    }
    // ---- Transactions ----
    getOrCreateSheet(SHEET_TRANSACTIONS, TRANSACTIONS_HEADERS);
    // ---- Settings (ค่าเริ่มต้น) ----
    const settings = getOrCreateSheet(SHEET_SETTINGS, ['key', 'value']);
    if (settings.getLastRow() < 2) {
      settings.getRange(2, 1, 13, 2).setValues([
        ['CompanyName', 'บริษัท ตัวอย่าง จำกัด'],
        ['CompanyAddress', '123 ถนนตัวอย่าง แขวง/ตำบล เขต/อำเภอ จังหวัด 10000'],
        ['LogoUrl', ''],
        ['LineToken', ''],
        ['LineTarget', ''],
        ['TelegramToken', ''],
        ['TelegramChatId', ''],
        ['EmailTo', ''],
        ['NotifyDaily', 'false'],
        ['NotifyWeekly', 'false'],
        ['NotifyMonthly', 'false'],
        ['LowStockThreshold', '5'],
        ['NearExpiryDays', '3']
      ]);
    }
    // ---- Users (สร้างบัญชี admin เริ่มต้น) ----
    const users = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    if (users.getLastRow() < 2) {
      const salt = generateSalt();
      const hash = hashPasswordHash('Admin@1234', salt);
      // ForceChangePassword = true เพื่อบังคับเปลี่ยนรหัสผ่านทันทีที่ล็อกอินครั้งแรก
      users.appendRow(['U' + new Date().getTime(), 'admin', 'admin@example.com', hash, salt, 'admin', 'Approved', '', 0, '', '', '', true, new Date()]);
    }
    // ---- ActivityLogs ----
    getOrCreateSheet(SHEET_LOGS, LOGS_HEADERS);
    return { success: true, message: 'ตั้งค่าชีตทั้งหมดเรียบร้อยแล้ว (admin / Admin@1234 ต้องเปลี่ยนรหัสผ่านทันทีหลังล็อกอินครั้งแรก)' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาดในการตั้งค่า: ' + err.message };
  }
}

// ============================================================================
//  ส่วนที่ 3: Authentication / Session / Security
// ============================================================================

function generateSalt() {
  // สุ่ม salt อย่างน้อย 16 byte ต่อผู้ใช้ (ใช้ 2 UUID ต่อกันให้ยาวและสุ่มมากขึ้น)
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

// แฮชรหัสผ่านด้วย SHA-256 + salt และทำ key stretching อย่างน้อย 1000 รอบ
function hashPasswordHash(password, salt) {
  let value = String(password) + String(salt);
  for (let i = 0; i < HASH_ROUNDS; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value + salt);
    value = bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
  }
  return value;
}

function loginUser(username, password) {
  try {
    if (!username || !password) return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const data = sh.getDataRange().getValues();
    const idx = headerIndexMap(data[0]);
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = String(row[idx.Username]).trim();
      const rowEmail = String(row[idx.Email]).trim();
      if (rowUsername.toLowerCase() !== username.trim().toLowerCase() && rowEmail.toLowerCase() !== username.trim().toLowerCase()) continue;

      const now = new Date();
      const lockoutUntil = row[idx.LockoutUntil];
      if (lockoutUntil && new Date(lockoutUntil) > now) {
        return { success: false, message: 'บัญชีถูกล็อกชั่วคราวจากการเข้าสู่ระบบผิดพลาดหลายครั้ง กรุณาลองใหม่หลัง ' + Utilities.formatDate(new Date(lockoutUntil), 'Asia/Bangkok', 'HH:mm') };
      }
      if (row[idx.Status] !== 'Approved') {
        return { success: false, message: 'บัญชีนี้ยังไม่ได้รับอนุมัติ หรือถูกระงับการใช้งาน' };
      }
      const salt = row[idx.Salt];
      const hash = hashPasswordHash(password, salt);
      if (hash === row[idx.PasswordHash]) {
        // ---- ล็อกอินสำเร็จ: รีเซ็ตตัวนับผิดพลาด + สร้าง session token ----
        sh.getRange(i + 1, idx.FailedAttempts + 1).setValue(0);
        sh.getRange(i + 1, idx.LockoutUntil + 1).setValue('');
        sh.getRange(i + 1, idx.LastLoginDate + 1).setValue(now);
        const token = Utilities.getUuid();
        CacheService.getScriptCache().put('session_' + token, JSON.stringify({
          userId: row[idx.UserID], username: rowUsername, role: row[idx.Role]
        }), SESSION_TTL);
        logActivity(row[idx.UserID], rowUsername, 'เข้าสู่ระบบ', 'เข้าสู่ระบบสำเร็จ');
        return {
          success: true, token: token, userId: row[idx.UserID], username: rowUsername,
          role: row[idx.Role], forceChangePassword: (row[idx.ForceChangePassword] === true || row[idx.ForceChangePassword] === 'TRUE')
        };
      } else {
        const attempts = (Number(row[idx.FailedAttempts]) || 0) + 1;
        sh.getRange(i + 1, idx.FailedAttempts + 1).setValue(attempts);
        if (attempts >= MAX_FAILED_ATTEMPTS) {
          const lockUntil = new Date(now.getTime() + LOCKOUT_MINUTES * 60000);
          sh.getRange(i + 1, idx.LockoutUntil + 1).setValue(lockUntil);
          logActivity(row[idx.UserID], rowUsername, 'บัญชีถูกล็อก', 'เข้าสู่ระบบผิดครบ ' + MAX_FAILED_ATTEMPTS + ' ครั้ง');
          return { success: false, message: 'กรอกรหัสผ่านผิดครบ ' + MAX_FAILED_ATTEMPTS + ' ครั้ง บัญชีถูกล็อก ' + LOCKOUT_MINUTES + ' นาที' };
        }
        return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (เหลือโอกาสอีก ' + (MAX_FAILED_ATTEMPTS - attempts) + ' ครั้ง)' };
      }
    }
    return { success: false, message: 'ไม่พบบัญชีผู้ใช้นี้ในระบบ' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function validateSession(token) {
  if (!token) return { valid: false };
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session_' + token);
  if (!raw) return { valid: false };
  const obj = JSON.parse(raw);
  cache.put('session_' + token, raw, SESSION_TTL); // ต่ออายุ session เมื่อมีการใช้งาน
  return { valid: true, userId: obj.userId, username: obj.username, role: obj.role };
}

function validateSessionResponse(token) {
  const s = validateSession(token);
  if (!s.valid) return { success: false, message: 'Session หมดอายุ กรุณาเข้าสู่ระบบใหม่' };
  return { success: true, userId: s.userId, username: s.username, role: s.role };
}

// ตรวจสอบสิทธิ์: ต้อง login และ role ต้องอยู่ใน roles ที่กำหนด (roles=null คือ login แล้วใช้ได้ทุก role)
function requireRole(token, roles) {
  const session = validateSession(token);
  if (!session.valid) return { ok: false, response: { success: false, message: 'กรุณาเข้าสู่ระบบใหม่ (session หมดอายุหรือไม่ถูกต้อง)' } };
  if (roles && roles.indexOf(session.role) === -1) {
    return { ok: false, response: { success: false, message: 'คุณไม่มีสิทธิ์ใช้งานส่วนนี้' } };
  }
  return { ok: true, session: session };
}

function logoutUser(token) {
  try {
    const session = validateSession(token);
    CacheService.getScriptCache().remove('session_' + token);
    if (session.valid) logActivity(session.userId, session.username, 'ออกจากระบบ', '');
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function changePassword(token, oldPassword, newPassword) {
  const auth = requireRole(token, null);
  if (!auth.ok) return auth.response;
  try {
    if (!newPassword || newPassword.length < 8) return { success: false, message: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร' };
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const vals = sh.getDataRange().getValues();
    const idx = headerIndexMap(vals[0]);
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][idx.UserID]) !== String(auth.session.userId)) continue;
      const forced = (vals[i][idx.ForceChangePassword] === true || vals[i][idx.ForceChangePassword] === 'TRUE');
      const oldHash = hashPasswordHash(oldPassword, vals[i][idx.Salt]);
      if (!forced && oldHash !== vals[i][idx.PasswordHash]) {
        return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      }
      const salt = generateSalt();
      const hash = hashPasswordHash(newPassword, salt);
      sh.getRange(i + 1, idx.PasswordHash + 1).setValue(hash);
      sh.getRange(i + 1, idx.Salt + 1).setValue(salt);
      sh.getRange(i + 1, idx.ForceChangePassword + 1).setValue(false);
      logActivity(auth.session.userId, auth.session.username, 'เปลี่ยนรหัสผ่าน', '');
      return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (err) { return { success: false, message: err.message }; }
}

function forgotPassword(email) {
  try {
    if (!email) return { success: false, message: 'กรุณาระบุอีเมล' };
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const data = sh.getDataRange().getValues();
    const idx = headerIndexMap(data[0]);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.Email]).toLowerCase() !== String(email).toLowerCase()) continue;
      const resetToken = Utilities.getUuid();
      const expiry = new Date(Date.now() + 3600000); // หมดอายุใน 1 ชั่วโมง
      sh.getRange(i + 1, idx.ResetToken + 1).setValue(resetToken);
      sh.getRange(i + 1, idx.ResetTokenExpiry + 1).setValue(expiry);
      // ส่งรหัสยืนยันทางอีเมลเท่านั้น (ห้ามแสดงบนหน้าจอ)
      MailApp.sendEmail(email, 'รีเซ็ตรหัสผ่าน - ระบบยืม-คืนสินค้า',
        'รหัสยืนยันสำหรับตั้งรหัสผ่านใหม่ของคุณคือ: ' + resetToken +
        '\nรหัสนี้จะหมดอายุภายใน 1 ชั่วโมง\nกรุณานำรหัสนี้ไปกรอกในหน้า "ลืมรหัสผ่าน" ของระบบ');
      logActivity(data[i][idx.UserID], data[i][idx.Username], 'ขอรีเซ็ตรหัสผ่าน', '');
      return { success: true, message: 'ส่งรหัสยืนยันไปยังอีเมลแล้ว กรุณาตรวจสอบกล่องจดหมาย' };
    }
    // ไม่เปิดเผยว่าไม่พบอีเมล เพื่อป้องกัน user enumeration
    return { success: true, message: 'หากอีเมลนี้มีอยู่ในระบบ คุณจะได้รับรหัสยืนยันทางอีเมล' };
  } catch (err) { return { success: false, message: err.message }; }
}

function resetPassword(data) {
  try {
    const email = data.email, resetToken = data.resetToken, newPassword = data.newPassword;
    if (!email || !resetToken || !newPassword) return { success: false, message: 'กรอกข้อมูลไม่ครบถ้วน' };
    if (newPassword.length < 8) return { success: false, message: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร' };
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const vals = sh.getDataRange().getValues();
    const idx = headerIndexMap(vals[0]);
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][idx.Email]).toLowerCase() !== String(email).toLowerCase()) continue;
      if (!vals[i][idx.ResetToken] || String(vals[i][idx.ResetToken]) !== String(resetToken)) {
        return { success: false, message: 'รหัสยืนยันไม่ถูกต้อง' };
      }
      if (!vals[i][idx.ResetTokenExpiry] || new Date(vals[i][idx.ResetTokenExpiry]) < new Date()) {
        return { success: false, message: 'รหัสยืนยันหมดอายุแล้ว กรุณาขอรหัสใหม่' };
      }
      const salt = generateSalt();
      const hash = hashPasswordHash(newPassword, salt);
      sh.getRange(i + 1, idx.PasswordHash + 1).setValue(hash);
      sh.getRange(i + 1, idx.Salt + 1).setValue(salt);
      sh.getRange(i + 1, idx.ResetToken + 1).setValue('');
      sh.getRange(i + 1, idx.ResetTokenExpiry + 1).setValue('');
      sh.getRange(i + 1, idx.ForceChangePassword + 1).setValue(false);
      sh.getRange(i + 1, idx.FailedAttempts + 1).setValue(0);
      sh.getRange(i + 1, idx.LockoutUntil + 1).setValue('');
      logActivity(vals[i][idx.UserID], vals[i][idx.Username], 'ตั้งรหัสผ่านใหม่ (ลืมรหัสผ่าน)', '');
      return { success: true, message: 'ตั้งรหัสผ่านใหม่สำเร็จ กรุณาเข้าสู่ระบบ' };
    }
    return { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
//  ส่วนที่ 4: Products / Partners / Transactions (สแกน + บันทึก)
// ============================================================================

function checkProduct(code, token) {
  const auth = requireRole(token, ['admin', 'staff']);
  if (!auth.ok) return auth.response;
  try {
    if (!code) return { success: false, found: false, message: 'กรุณาระบุรหัสสินค้า' };
    const sh = getOrCreateSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(code).trim().toLowerCase()) {
        return {
          success: true, found: true, code: data[i][0], name: data[i][1],
          category: data[i][2], quantity: Number(data[i][3]) || 0, unit: data[i][4] || '-'
        };
      }
    }
    return { success: true, found: false, message: 'ไม่พบรหัสสินค้านี้ในระบบ' };
  } catch (err) {
    return { success: false, found: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  }
}

function getPartners(type, token) {
  const auth = requireRole(token, ['admin', 'staff', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_PARTNERS, PARTNERS_HEADERS);
    const data = sh.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (type && data[i][2] !== type) continue;
      list.push({ code: data[i][0], name: data[i][1], type: data[i][2] });
    }
    return { success: true, partners: list };
  } catch (err) { return { success: false, message: err.message }; }
}

/**
 * saveTransaction: บันทึกใบยืม/คืน 1 ใบ ซึ่งมีได้หลายรายการสินค้า (items[])
 * ตรวจสอบ field บังคับซ้ำฝั่ง Server ทุกจุด + กันบันทึกชนกันด้วย LockService
 */
function saveTransaction(data, token) {
  const auth = requireRole(token, ['admin', 'staff']);
  if (!auth.ok) return auth.response;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!data) return { success: false, message: 'ไม่มีข้อมูลที่ส่งมา' };
    const type = data.type;
    const docRef = String(data.docRef || '').trim();
    const partnerType = data.partnerType;
    const partnerCode = String(data.partnerCode || '').trim();
    const partnerName = String(data.partnerName || '').trim();
    const expiryDate = String(data.expiryDate || '').trim();
    const note = String(data.note || '').trim();
    const items = Array.isArray(data.items) ? data.items : [];

    // ---- ตรวจสอบ field บังคับ (ฝั่ง Server) ----
    const errors = [];
    if (type !== 'รับคืน' && type !== 'เบิกออก') errors.push('กรุณาเลือก "รับคืน-เบิกออก"');
    if (!docRef) errors.push('กรุณากรอกเอกสารอ้างอิง');
    if (partnerType !== 'ผู้ค้า' && partnerType !== 'ลูกค้า') errors.push('กรุณาเลือกประเภทผู้ค้า-ลูกค้า');
    if (!partnerCode) errors.push('กรุณากรอกรหัสผู้ค้า-ลูกค้า');
    if (!items.length) errors.push('กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ');
    items.forEach(function (it, i) {
      if (!it || !it.code) errors.push('รายการที่ ' + (i + 1) + ': กรุณาระบุรหัสสินค้า');
      if (!it || !(Number(it.qty) > 0)) errors.push('รายการที่ ' + (i + 1) + ': จำนวนต้องมากกว่า 0');
    });
    if (errors.length) return { success: false, message: errors.join(', ') };

    // ---- ตรวจสอบรหัสสินค้ามีจริงทุกตัวก่อนบันทึก ----
    const productSheet = getOrCreateSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const productData = productSheet.getDataRange().getValues();
    const productIndex = {};
    for (let i = 1; i < productData.length; i++) productIndex[String(productData[i][0]).trim()] = i;

    const resolvedItems = [];
    for (let i = 0; i < items.length; i++) {
      const code = String(items[i].code).trim();
      if (!productIndex.hasOwnProperty(code)) {
        return { success: false, message: 'ไม่พบรหัสสินค้า "' + code + '" ในระบบ (รายการที่ ' + (i + 1) + ')' };
      }
      const rowIdx = productIndex[code];
      resolvedItems.push({
        code: code, name: productData[rowIdx][1], category: productData[rowIdx][2],
        qty: Number(items[i].qty), unit: productData[rowIdx][4] || '-', rowIdx: rowIdx
      });
    }

    const now = new Date();
    const docNo = 'DOC' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
    const timestampStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

    const txSheet = getOrCreateSheet(SHEET_TRANSACTIONS, TRANSACTIONS_HEADERS);
    const rowsToAppend = [];
    resolvedItems.forEach(function (it) {
      rowsToAppend.push([timestampStr, it.code, it.name, it.category, type, it.qty, docRef,
        partnerType, partnerCode, partnerName, expiryDate, note, auth.session.username, docNo]);
      // ปรับปรุงจำนวนคงเหลือในชีต Products (เบิกออก = ลด, รับคืน = เพิ่ม)
      const currentQty = Number(productData[it.rowIdx][3]) || 0;
      const newQty = type === 'รับคืน' ? currentQty + it.qty : currentQty - it.qty;
      productSheet.getRange(it.rowIdx + 1, 4).setValue(newQty);
      productData[it.rowIdx][3] = newQty;
    });
    txSheet.getRange(txSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    CacheService.getScriptCache().remove('products_cache_v1'); // ล้างแคชสินค้าหลังมีการเปลี่ยนแปลง

    logActivity(auth.session.userId, auth.session.username, 'บันทึกรายการ' + type, docNo + ' (' + resolvedItems.length + ' รายการ)');

    return {
      success: true, docNo: docNo, timestamp: timestampStr, type: type, docRef: docRef,
      partnerType: partnerType, partnerCode: partnerCode, partnerName: partnerName,
      expiryDate: expiryDate, note: note, items: resolvedItems, recordedBy: auth.session.username
    };
  } catch (err) {
    return { success: false, message: 'บันทึกไม่สำเร็จ: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
//  ส่วนที่ 5: สร้าง PDF ใบยืม/คืน (HTML Template ฝังเป็น string ในไฟล์นี้)
// ============================================================================

function buildPdfTemplateString() {
  // หมายเหตุ: ใช้ <?= ?> (auto-escape) ทุกจุดเพื่อป้องกัน HTML Injection
  // ฟอนต์ใช้ fallback หลายชั้นให้แสดงภาษาไทยได้แน่นอนบน PDF engine ของ GAS
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:"Sarabun","Noto Sans Thai","Leelawadee UI",Tahoma,Arial,sans-serif;font-size:13px;color:#222;}' +
    '.header-table{width:100%;border-collapse:collapse;}' +
    '.company-name{font-size:16px;font-weight:bold;}' +
    '.company-address{font-size:11px;color:#555;}' +
    '.doc-title{font-size:18px;font-weight:bold;color:#5b3fd1;}' +
    '.info-table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px;}' +
    '.info-table td{padding:3px 4px;}' +
    '.item-table{width:100%;border-collapse:collapse;margin-top:10px;font-size:12px;}' +
    '.item-table th,.item-table td{border:1px solid #999;padding:5px;}' +
    '.item-table th{background:#efe9ff;}' +
    '.sign-table{width:100%;margin-top:70px;font-size:12px;}' +
    '.sign-table td{width:50%;text-align:center;}' +
    '</style></head><body>' +
    '<table class="header-table"><tr>' +
    '<td style="width:80px;"><? if (logoDataUri) { ?><img src="<?= logoDataUri ?>" style="width:65px;height:65px;"><? } ?></td>' +
    '<td><div class="company-name"><?= companyName ?></div><div class="company-address"><?= companyAddress ?></div></td>' +
    '<td style="text-align:right;"><div class="doc-title"><?= docTitle ?></div><div>เลขที่: <?= docNo ?></div><div>วันที่: <?= docDate ?></div></td>' +
    '</tr></table><hr>' +
    '<table class="info-table">' +
    '<tr><td><b>ผู้ยืม/คู่ค้า:</b></td><td><?= partnerName ?> (<?= partnerCode ?>)</td><td><b>เอกสารอ้างอิง:</b></td><td><?= docRef ?></td></tr>' +
    '<tr><td><b>ประเภท:</b></td><td><?= partnerType ?></td><td><b>กำหนดคืน:</b></td><td><?= expiryDate ?></td></tr>' +
    '</table>' +
    '<table class="item-table"><thead><tr><th>ลำดับ</th><th>รหัส</th><th>ชื่อสินค้า</th><th>จำนวน</th><th>หน่วย</th><th>สภาพ</th></tr></thead><tbody>' +
    '<? for (var i = 0; i < items.length; i++) { ?>' +
    '<tr><td style="text-align:center;"><?= i + 1 ?></td><td><?= items[i].code ?></td><td><?= items[i].name ?></td>' +
    '<td style="text-align:center;"><?= items[i].qty ?></td><td style="text-align:center;"><?= items[i].unit ?></td><td style="text-align:center;"><?= items[i].condition ?></td></tr>' +
    '<? } ?>' +
    '</tbody></table>' +
    '<p><b>หมายเหตุ:</b> <?= note ?></p>' +
    '<table class="sign-table"><tr>' +
    '<td>ลงชื่อ .......................................................<br>(<?= signer1Label ?>)<br>วันที่ ........./........./.........</td>' +
    '<td>ลงชื่อ .......................................................<br>(<?= signer2Label ?>)<br>วันที่ ........./........./.........</td>' +
    '</tr></table>' +
    '</body></html>';
}

function generateTransactionPDF(data, token) {
  const auth = requireRole(token, ['admin', 'staff']); // เฉพาะ admin, staff
  if (!auth.ok) return auth.response;
  try {
    const docNo = data.docNo;
    if (!docNo) return { success: false, message: 'ไม่พบเลขที่ใบ' };
    const sh = getOrCreateSheet(SHEET_TRANSACTIONS, TRANSACTIONS_HEADERS);
    const rows = sh.getDataRange().getValues();
    const idx = headerIndexMap(rows[0]);
    const items = [];
    let header = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idx['เลขที่ใบ']]) === String(docNo)) {
        if (!header) header = rows[i];
        items.push({ code: rows[i][idx['รหัสสินค้า']], name: rows[i][idx['ชื่อสินค้า']], qty: rows[i][idx['จำนวน']], unit: '-', condition: 'ปกติ' });
      }
    }
    if (!header) return { success: false, message: 'ไม่พบข้อมูลใบนี้' };

    const settings = getSettingsMap();
    let logoDataUri = '';
    try {
      if (settings.LogoUrl) {
        const resp = UrlFetchApp.fetch(settings.LogoUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          const blob = resp.getBlob();
          logoDataUri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
        }
      }
    } catch (e) { /* ไม่มีโลโก้ก็ไม่เป็นไร */ }

    const type = header[idx['รับคืน-เบิกออก']];
    const tpl = HtmlService.createTemplate(buildPdfTemplateString());
    tpl.companyName = settings.CompanyName || 'บริษัท ตัวอย่าง จำกัด';
    tpl.companyAddress = settings.CompanyAddress || '-';
    tpl.logoDataUri = logoDataUri;
    tpl.docTitle = type === 'เบิกออก' ? 'ใบเบิกออก / ใบยืม' : 'ใบรับคืน';
    tpl.docNo = docNo;
    tpl.docDate = header[idx['วันเวลาบันทึก']];
    tpl.partnerName = header[idx['ชื่อผู้ค้า-ลูกค้า']];
    tpl.partnerCode = header[idx['รหัสผู้ค้า-ลูกค้า']];
    tpl.partnerType = header[idx['ประเภทคู่ค้า']];
    tpl.docRef = header[idx['เอกสารอ้างอิง']];
    tpl.expiryDate = header[idx['วันหมดอายุ']] || '-';
    tpl.note = header[idx['หมายเหตุ']] || '-';
    tpl.items = items;
    tpl.signer1Label = type === 'เบิกออก' ? 'ผู้ยืม' : 'ผู้คืน';
    tpl.signer2Label = type === 'เบิกออก' ? 'ผู้อนุมัติ' : 'ผู้รับคืน';

    const pdfBlob = tpl.evaluate().getAs('application/pdf').setName(docNo + '.pdf');
    const base64 = Utilities.base64Encode(pdfBlob.getBytes());
    logActivity(auth.session.userId, auth.session.username, 'พิมพ์ใบยืม-คืน (PDF)', docNo);
    return { success: true, filename: docNo + '.pdf', base64: base64 };
  } catch (err) {
    return { success: false, message: 'สร้าง PDF ไม่สำเร็จ: ' + err.message };
  }
}

// ============================================================================
//  ส่วนที่ 6: วันที่ / ตัวช่วยแปลงรูปแบบ
// ============================================================================

function parseThaiDateTime(str) {
  try {
    if (str instanceof Date) return str;
    const parts = String(str).trim().split(' ');
    const dmy = parts[0].split('/');
    const hms = (parts[1] || '00:00:00').split(':');
    return new Date(Number(dmy[2]), Number(dmy[1]) - 1, Number(dmy[0]), Number(hms[0] || 0), Number(hms[1] || 0), Number(hms[2] || 0));
  } catch (e) { return new Date(0); }
}

function parseDMY(str) {
  try {
    if (!str) return null;
    str = String(str).trim();
    if (str.indexOf('-') !== -1) { const p = str.split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
    if (str.indexOf('/') !== -1) { const p = str.split('/'); return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])); }
    return null;
  } catch (e) { return null; }
}

// ============================================================================
//  ส่วนที่ 7: ค้นหารายการยืม-คืน / ค้นหายอดคงเหลือ (เน้นความเร็ว)
// ============================================================================

// ดึงรายการ Transactions ตามเงื่อนไข filter (ใช้ภายใน ไม่ตรวจสิทธิ์ซ้ำ)
function getFilteredTransactionRows(filters) {
  filters = filters || {};
  const sh = getOrCreateSheet(SHEET_TRANSACTIONS, TRANSACTIONS_HEADERS);
  const data = sh.getDataRange().getValues();
  const dateFrom = filters.dateFrom ? parseDMY(filters.dateFrom) : null;
  const dateTo = filters.dateTo ? parseDMY(filters.dateTo) : null;
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const rowDate = parseThaiDateTime(r[0]);
    if (dateFrom && rowDate < dateFrom) continue;
    if (dateTo) { const end = new Date(dateTo.getTime() + 86399000); if (rowDate > end) continue; }
    if (filters.type && r[4] !== filters.type) continue;
    if (filters.partner) {
      const p = String(filters.partner).toLowerCase();
      if (String(r[8]).toLowerCase().indexOf(p) === -1 && String(r[9]).toLowerCase().indexOf(p) === -1) continue;
    }
    if (filters.product) {
      const p = String(filters.product).toLowerCase();
      if (String(r[1]).toLowerCase().indexOf(p) === -1 && String(r[2]).toLowerCase().indexOf(p) === -1) continue;
    }
    if (filters.category) {
      const p = String(filters.category).toLowerCase();
      if (String(r[3]).toLowerCase().indexOf(p) === -1) continue;
    }
    rows.push({
      timestamp: r[0], code: r[1], name: r[2], category: r[3], type: r[4], qty: Number(r[5]) || 0,
      docRef: r[6], partnerType: r[7], partnerCode: r[8], partnerName: r[9], expiryDate: r[10],
      note: r[11], recordedBy: r[12], docNo: r[13]
    });
  }
  return rows;
}

function searchTransactions(filters, token) {
  const auth = requireRole(token, ['admin', 'staff', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    const rows = getFilteredTransactionRows(filters);
    rows.sort(function (a, b) { return parseThaiDateTime(b.timestamp) - parseThaiDateTime(a.timestamp); });
    const sum = rows.reduce(function (s, r) { return s + r.qty; }, 0);
    const page = Number(filters.page) || 1;
    const pageSize = Number(filters.pageSize) || 20;
    const paged = rows.slice((page - 1) * pageSize, page * pageSize);
    return { success: true, rows: paged, total: rows.length, sum: sum, page: page, pageSize: pageSize };
  } catch (err) { return { success: false, message: err.message }; }
}

function sendSearchResultNotification(data, token) {
  const auth = requireRole(token, ['admin', 'staff', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    const message = data.message || 'ผลการค้นหารายการยืม-คืนสินค้า';
    const settings = getSettingsMap();
    const results = {};
    if (settings.LineToken) results.line = sendLineMessage(message);
    if (settings.TelegramToken) results.telegram = sendTelegramMessage(message);
    if (settings.EmailTo) results.email = sendEmailNotification('ผลการค้นหารายการยืม-คืน', message);
    logActivity(auth.session.userId, auth.session.username, 'ส่งแจ้งเตือนผลการค้นหา', '');
    return { success: true, results: results };
  } catch (err) { return { success: false, message: err.message }; }
}

// แคชรายการสินค้าไว้ 30 วินาที เพื่อรองรับการเรียกพร้อมกันจำนวนมาก (เช่น ทีม Sales)
function getProductsCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('products_cache_v1');
  if (cached) return JSON.parse(cached);
  const sh = getOrCreateSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({ code: String(data[i][0]), name: String(data[i][1]), category: String(data[i][2]), quantity: Number(data[i][3]) || 0, unit: String(data[i][4] || '-') });
  }
  cache.put('products_cache_v1', JSON.stringify(list), 30);
  return list;
}

function getStockBalance(query, token) {
  const auth = requireRole(token, ['admin', 'staff', 'salesperson', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    const list = getProductsCached();
    const q = String(query || '').trim().toLowerCase();
    const result = !q ? list.slice(0, 50) : list.filter(function (p) {
      return p.code.toLowerCase().indexOf(q) !== -1 || p.name.toLowerCase().indexOf(q) !== -1;
    }).slice(0, 50);
    return { success: true, items: result };
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
//  ส่วนที่ 8: Dashboard (admin, manager เท่านั้น - manager อ่านอย่างเดียว)
// ============================================================================

function getDashboardData(filters, token) {
  const auth = requireRole(token, ['admin', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    filters = filters || {};
    const rows = getFilteredTransactionRows(filters);
    const settings = getSettingsMap();
    const nearDueDays = Number(filters.nearDueDays || settings.NearExpiryDays || 3);
    const lowStockThreshold = Number(settings.LowStockThreshold || 5);
    const today = new Date(); today.setHours(0, 0, 0, 0);

    const docsIn = new Set(), docsOut = new Set();
    let qtyIn = 0, qtyOut = 0, damagedLost = 0;
    const productAgg = {}, partnerAgg = {};
    const outstandingMap = {};
    rows.forEach(function (r) {
      if (r.type === 'รับคืน') { docsIn.add(r.docNo); qtyIn += r.qty; }
      else { docsOut.add(r.docNo); qtyOut += r.qty; }
      const noteStr = String(r.note || '');
      if (noteStr.indexOf('ชำรุด') !== -1 || noteStr.indexOf('สูญหาย') !== -1) damagedLost += r.qty;
      productAgg[r.name] = (productAgg[r.name] || 0) + r.qty;
      partnerAgg[r.partnerName] = (partnerAgg[r.partnerName] || 0) + r.qty;
      const key = r.code + '|' + r.partnerCode;
      if (!outstandingMap[key]) outstandingMap[key] = { code: r.code, name: r.name, partnerCode: r.partnerCode, partnerName: r.partnerName, qty: 0, expiryDate: '' };
      outstandingMap[key].qty += (r.type === 'เบิกออก' ? r.qty : -r.qty);
      if (r.type === 'เบิกออก' && r.expiryDate) outstandingMap[key].expiryDate = r.expiryDate;
    });
    const outstandingList = Object.keys(outstandingMap).map(function (k) { return outstandingMap[k]; }).filter(function (o) { return o.qty > 0; });

    let overdueCount = 0, overdueQty = 0, overdueDaysTotal = 0;
    const nearDueList = [], overdueList = [];
    outstandingList.forEach(function (o) {
      if (!o.expiryDate) return;
      const due = parseDMY(o.expiryDate);
      if (!due) return;
      const diffDays = Math.round((due - today) / 86400000);
      if (diffDays < 0) {
        overdueCount++; overdueQty += o.qty; overdueDaysTotal += (-diffDays);
        overdueList.push(Object.assign({}, o, { daysOverdue: -diffDays }));
      } else if (diffDays <= nearDueDays) {
        nearDueList.push(Object.assign({}, o, { daysLeft: diffDays }));
      }
    });
    const totalDueItems = outstandingList.filter(function (o) { return o.expiryDate; }).length;
    const onTimeCount = totalDueItems - overdueCount;
    const onTimeRate = totalDueItems > 0 ? Math.round((onTimeCount / totalDueItems) * 1000) / 10 : 100;

    const topProducts = Object.keys(productAgg).map(function (k) { return { name: k, qty: productAgg[k] }; }).sort(function (a, b) { return b.qty - a.qty; }).slice(0, 10);
    const topPartners = Object.keys(partnerAgg).map(function (k) { return { name: k, qty: partnerAgg[k] }; }).sort(function (a, b) { return b.qty - a.qty; }).slice(0, 10);
    const lowStock = getProductsCached().filter(function (p) { return p.quantity <= lowStockThreshold; });

    // เปรียบเทียบกับช่วงเวลาเดียวกันก่อนหน้า
    let prevCompare = null;
    if (filters.dateFrom && filters.dateTo) {
      const from = parseDMY(filters.dateFrom), to = parseDMY(filters.dateTo);
      const durationMs = to - from;
      const prevTo = new Date(from.getTime() - 86400000);
      const prevFrom = new Date(prevTo.getTime() - durationMs);
      const prevRows = getFilteredTransactionRows(Object.assign({}, filters, {
        dateFrom: Utilities.formatDate(prevFrom, 'Asia/Bangkok', 'yyyy-MM-dd'),
        dateTo: Utilities.formatDate(prevTo, 'Asia/Bangkok', 'yyyy-MM-dd')
      }));
      let prevIn = 0, prevOut = 0;
      prevRows.forEach(function (r) { if (r.type === 'รับคืน') prevIn += r.qty; else prevOut += r.qty; });
      prevCompare = {
        qtyIn: prevIn, qtyOut: prevOut,
        qtyInChangePct: prevIn > 0 ? Math.round(((qtyIn - prevIn) / prevIn) * 1000) / 10 : null,
        qtyOutChangePct: prevOut > 0 ? Math.round(((qtyOut - prevOut) / prevOut) * 1000) / 10 : null
      };
    }

    return {
      success: true,
      kpi: {
        docsIn: docsIn.size, docsOut: docsOut.size, qtyIn: qtyIn, qtyOut: qtyOut,
        outstandingItems: outstandingList.length,
        outstandingQty: outstandingList.reduce(function (s, o) { return s + o.qty; }, 0),
        onTimeRate: onTimeRate, overdueCount: overdueCount, overdueQty: overdueQty,
        avgOverdueDays: overdueCount > 0 ? Math.round((overdueDaysTotal / overdueCount) * 10) / 10 : 0,
        damagedLost: damagedLost, lowStockCount: lowStock.length
      },
      prevCompare: prevCompare,
      nearDueList: nearDueList.slice(0, 50), overdueList: overdueList.slice(0, 50),
      topProducts: topProducts, topPartners: topPartners, lowStock: lowStock.slice(0, 50),
      rawRows: rows.slice(0, 3000) // ให้ client นำไปสร้างกราฟ Chart.js แบบ dynamic เอง
    };
  } catch (err) { return { success: false, message: err.message }; }
}

// บันทึก/อ่านการตั้งค่ากราฟส่วนตัว (เฉพาะ admin เท่านั้นที่บันทึกลงชีตได้ - manager ใช้ localStorage)
function saveDashboardPrefs(data, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_SETTINGS, ['key', 'value']);
    const key = 'dashpref_' + auth.session.username;
    const vals = sh.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < vals.length; i++) { if (vals[i][0] === key) { foundRow = i + 1; break; } }
    const value = JSON.stringify(data.prefs || {});
    if (foundRow > 0) sh.getRange(foundRow, 2).setValue(value); else sh.appendRow([key, value]);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function getDashboardPrefs(token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const map = getSettingsMap();
    const key = 'dashpref_' + auth.session.username;
    return { success: true, prefs: map[key] ? JSON.parse(map[key]) : null };
  } catch (err) { return { success: false, message: err.message }; }
}

function getNotificationBadgeCount(token) {
  const auth = requireRole(token, ['admin', 'staff', 'manager']);
  if (!auth.ok) return auth.response;
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('badge_cache_v1');
    if (cached) return JSON.parse(cached);
    const settings = getSettingsMap();
    const lowStockThreshold = Number(settings.LowStockThreshold || 5);
    const low = getProductsCached().filter(function (p) { return p.quantity <= lowStockThreshold; }).length;
    const rows = getFilteredTransactionRows({});
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const outstandingMap = {};
    rows.forEach(function (r) {
      const key = r.code + '|' + r.partnerCode;
      if (!outstandingMap[key]) outstandingMap[key] = { qty: 0, expiryDate: '' };
      outstandingMap[key].qty += (r.type === 'เบิกออก' ? r.qty : -r.qty);
      if (r.type === 'เบิกออก' && r.expiryDate) outstandingMap[key].expiryDate = r.expiryDate;
    });
    let overdue = 0;
    Object.keys(outstandingMap).forEach(function (k) {
      const o = outstandingMap[k];
      if (o.qty > 0 && o.expiryDate) { const due = parseDMY(o.expiryDate); if (due && due < today) overdue++; }
    });
    const result = { success: true, count: low + overdue, low: low, overdue: overdue };
    cache.put('badge_cache_v1', JSON.stringify(result), 30);
    return result;
  } catch (err) { return { success: false, message: err.message }; }
}

// ============================================================================
//  ส่วนที่ 9: การแจ้งเตือน LINE / Telegram / Email + งานอัตโนมัติ
// ============================================================================

function getSettingsMap() {
  const sh = getOrCreateSheet(SHEET_SETTINGS, ['key', 'value']);
  const data = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) map[data[i][0]] = data[i][1]; }
  return map;
}

function getSettingsForClient(token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  return { success: true, settings: getSettingsMap() };
}

function saveSettings(data, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_SETTINGS, ['key', 'value']);
    const sheetData = sh.getDataRange().getValues();
    const keyRow = {};
    for (let i = 1; i < sheetData.length; i++) keyRow[sheetData[i][0]] = i + 1;
    Object.keys(data.settings || {}).forEach(function (key) {
      const value = data.settings[key];
      if (keyRow[key]) sh.getRange(keyRow[key], 2).setValue(value); else sh.appendRow([key, value]);
    });
    logActivity(auth.session.userId, auth.session.username, 'แก้ไขตั้งค่าระบบ', JSON.stringify(Object.keys(data.settings || {})));
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function sendLineMessage(message) {
  const s = getSettingsMap();
  if (!s.LineToken || !s.LineTarget) return { success: false, message: 'ยังไม่ได้ตั้งค่า LINE Messaging API' };
  try {
    const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + s.LineToken },
      payload: JSON.stringify({ to: s.LineTarget, messages: [{ type: 'text', text: message }] }),
      muteHttpExceptions: true
    });
    return { success: resp.getResponseCode() === 200, message: resp.getContentText() };
  } catch (err) { return { success: false, message: err.message }; }
}

function sendTelegramMessage(message) {
  const s = getSettingsMap();
  if (!s.TelegramToken || !s.TelegramChatId) return { success: false, message: 'ยังไม่ได้ตั้งค่า Telegram Bot' };
  try {
    const url = 'https://api.telegram.org/bot' + s.TelegramToken + '/sendMessage';
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: s.TelegramChatId, text: message }), muteHttpExceptions: true
    });
    return { success: resp.getResponseCode() === 200, message: resp.getContentText() };
  } catch (err) { return { success: false, message: err.message }; }
}

function sendEmailNotification(subject, body) {
  const s = getSettingsMap();
  if (!s.EmailTo) return { success: false, message: 'ยังไม่ได้ตั้งค่าอีเมลผู้รับ' };
  try {
    MailApp.sendEmail(s.EmailTo, subject, body);
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function testNotification(channel, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  const msg = 'ทดสอบการแจ้งเตือนจากระบบยืม-คืนสินค้า เวลา ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  if (channel === 'line') return sendLineMessage(msg);
  if (channel === 'telegram') return sendTelegramMessage(msg);
  if (channel === 'email') return sendEmailNotification('ทดสอบระบบแจ้งเตือน', msg);
  return { success: false, message: 'ไม่รู้จักช่องทางการแจ้งเตือน' };
}

// สร้าง Time-driven Trigger (รันฟังก์ชันนี้ 1 ครั้งด้วยมือจากตัวแก้ไขสคริปต์ เพื่อเปิดใช้งานแจ้งเตือนอัตโนมัติ)
function createTimeTriggers() {
  deleteExistingTriggers();
  ScriptApp.newTrigger('dailyNotificationJob').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('lowStockNotificationJob').timeBased().everyDays(1).atHour(9).create();
  return { success: true, message: 'ตั้งค่า Trigger อัตโนมัติเรียบร้อย' };
}
function deleteExistingTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyNotificationJob' || t.getHandlerFunction() === 'lowStockNotificationJob') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
function dailyNotificationJob() {
  const s = getSettingsMap();
  const today = new Date();
  if (String(s.NotifyDaily) === 'true') sendPeriodSummary('รายวัน', 1);
  if (String(s.NotifyWeekly) === 'true' && today.getDay() === 1) sendPeriodSummary('รายสัปดาห์', 7);
  if (String(s.NotifyMonthly) === 'true' && today.getDate() === 1) sendPeriodSummary('รายเดือน', 30);
}
function sendPeriodSummary(label, days) {
  const to = new Date(), from = new Date(Date.now() - days * 86400000);
  const rows = getFilteredTransactionRows({
    dateFrom: Utilities.formatDate(from, 'Asia/Bangkok', 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(to, 'Asia/Bangkok', 'yyyy-MM-dd')
  });
  let qtyIn = 0, qtyOut = 0;
  rows.forEach(function (r) { if (r.type === 'รับคืน') qtyIn += r.qty; else qtyOut += r.qty; });
  const message = 'สรุป' + label + ': เบิกออก ' + qtyOut + ' ชิ้น, รับคืน ' + qtyIn + ' ชิ้น (' + rows.length + ' รายการ)';
  sendLineMessage(message); sendTelegramMessage(message); sendEmailNotification('สรุป' + label + ' - ระบบยืม-คืนสินค้า', message);
}
function lowStockNotificationJob() {
  const s = getSettingsMap();
  const threshold = Number(s.LowStockThreshold || 5);
  const low = getProductsCached().filter(function (p) { return p.quantity <= threshold; });
  if (low.length === 0) return;
  const message = 'แจ้งเตือนสินค้าใกล้หมด: ' + low.map(function (p) { return p.name + '(' + p.quantity + ' ' + p.unit + ')'; }).join(', ');
  sendLineMessage(message); sendTelegramMessage(message); sendEmailNotification('แจ้งเตือนสินค้าใกล้หมด - ระบบยืม-คืนสินค้า', message);
}

// ============================================================================
//  ส่วนที่ 10: จัดการผู้ใช้งาน (admin เท่านั้น)
// ============================================================================

function listUsers(token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const data = sh.getDataRange().getValues();
    const idx = headerIndexMap(data[0]);
    const list = [];
    for (let i = 1; i < data.length; i++) {
      list.push({
        userId: data[i][idx.UserID], username: data[i][idx.Username], email: data[i][idx.Email],
        role: data[i][idx.Role], status: data[i][idx.Status], lastLogin: data[i][idx.LastLoginDate]
      });
    }
    return { success: true, users: list };
  } catch (err) { return { success: false, message: err.message }; }
}

function createUserAccount(data, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    if (!data.username || !data.email || !data.password || !data.role) return { success: false, message: 'กรอกข้อมูลไม่ครบถ้วน' };
    if (['admin', 'staff', 'salesperson', 'manager'].indexOf(data.role) === -1) return { success: false, message: 'บทบาทไม่ถูกต้อง' };
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][1]).toLowerCase() === data.username.toLowerCase() || String(vals[i][2]).toLowerCase() === data.email.toLowerCase()) {
        return { success: false, message: 'มีชื่อผู้ใช้หรืออีเมลนี้อยู่แล้วในระบบ' };
      }
    }
    const salt = generateSalt();
    const hash = hashPasswordHash(data.password, salt);
    const userId = 'U' + new Date().getTime();
    sh.appendRow([userId, data.username, data.email, hash, salt, data.role, 'Approved', '', 0, '', '', '', true, new Date()]);
    logActivity(auth.session.userId, auth.session.username, 'สร้างผู้ใช้ใหม่', data.username + ' (' + data.role + ')');
    return { success: true };
  } catch (err) { return { success: false, message: err.message }; }
}

function updateUserAccount(data, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_USERS, USERS_HEADERS);
    const vals = sh.getDataRange().getValues();
    const idx = headerIndexMap(vals[0]);
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][idx.UserID]) !== String(data.userId)) continue;
      if (data.role) sh.getRange(i + 1, idx.Role + 1).setValue(data.role);
      if (data.status) sh.getRange(i + 1, idx.Status + 1).setValue(data.status);
      logActivity(auth.session.userId, auth.session.username, 'แก้ไขผู้ใช้', String(data.userId));
      return { success: true };
    }
    return { success: false, message: 'ไม่พบผู้ใช้' };
  } catch (err) { return { success: false, message: err.message }; }
}

function setUserStatus(userId, status, token) {
  return updateUserAccount({ userId: userId, status: status }, token);
}

// ============================================================================
//  ส่วนที่ 11: Activity Logs / Import Products (CSV/TXT, UTF-8)
// ============================================================================

function logActivity(userId, username, action, detail) {
  try {
    const sh = getOrCreateSheet(SHEET_LOGS, LOGS_HEADERS);
    sh.appendRow([Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'), userId, username, action, detail]);
  } catch (e) { /* ไม่ให้ log ล้มเหลวกระทบการทำงานหลัก */ }
}

function getActivityLogs(data, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    const sh = getOrCreateSheet(SHEET_LOGS, LOGS_HEADERS);
    const vals = sh.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < vals.length; i++) rows.push({ time: vals[i][0], userId: vals[i][1], username: vals[i][2], action: vals[i][3], detail: vals[i][4] });
    rows.reverse();
    const page = Number(data.page) || 1, pageSize = Number(data.pageSize) || 20;
    return { success: true, rows: rows.slice((page - 1) * pageSize, page * pageSize), total: rows.length };
  } catch (err) { return { success: false, message: err.message }; }
}

function importProductsCSV(csvText, token) {
  const auth = requireRole(token, ['admin']);
  if (!auth.ok) return auth.response;
  try {
    if (!csvText) return { success: false, message: 'ไม่มีข้อมูล CSV' };
    const lines = String(csvText).replace(/\r/g, '').split('\n').filter(function (l) { return l.trim() !== ''; });
    if (lines.length < 2) return { success: false, message: 'ไฟล์ไม่มีข้อมูล (ต้องมีหัวตารางและข้อมูลอย่างน้อย 1 แถว)' };
    const sh = getOrCreateSheet(SHEET_PRODUCTS, PRODUCTS_HEADERS);
    const data = sh.getDataRange().getValues();
    const codeIndex = {};
    for (let i = 1; i < data.length; i++) codeIndex[String(data[i][0]).trim()] = i + 1;
    let added = 0, updated = 0;
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(',').map(function (c) { return c.trim(); });
      if (!cols[0]) continue;
      const code = cols[0], name = cols[1] || '', category = cols[2] || '', qty = Number(cols[3]) || 0, unit = cols[4] || '';
      if (codeIndex[code]) { sh.getRange(codeIndex[code], 1, 1, 5).setValues([[code, name, category, qty, unit]]); updated++; }
      else { sh.appendRow([code, name, category, qty, unit, '']); added++; }
    }
    CacheService.getScriptCache().remove('products_cache_v1');
    logActivity(auth.session.userId, auth.session.username, 'นำเข้าสินค้าจาก CSV', 'เพิ่ม ' + added + ' / แก้ไข ' + updated);
    return { success: true, added: added, updated: updated };
  } catch (err) { return { success: false, message: err.message }; }
}
