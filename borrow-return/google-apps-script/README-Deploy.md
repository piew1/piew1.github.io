# วิธี Deploy ระบบยืม-คืนสินค้า (Google Apps Script + GitHub Pages)

ไฟล์ที่ต้องใช้มีแค่ 2 ไฟล์:
- `Code.gs` → นำไปวางใน Google Apps Script (Backend + API)
- `index.html` → นำไปโฮสต์บน GitHub Pages (Frontend)

## ขั้นตอนที่ 1: สร้าง Google Sheet (Database)
1. ไปที่ https://sheets.google.com สร้างสเปรดชีตใหม่ ตั้งชื่อ เช่น "ระบบยืม-คืนสินค้า DB"
2. เมนู **ส่วนขยาย (Extensions) → Apps Script** เพื่อเปิดตัวแก้ไขสคริปต์ที่ผูกกับชีตนี้

## ขั้นตอนที่ 2: ใส่สคริปต์ Code.gs
1. ลบโค้ดเริ่มต้นในไฟล์ `Code.gs` ที่เปิดมาทั้งหมด
2. คัดลอกโค้ดทั้งหมดจากไฟล์ `google-apps-script/Code.gs` ในโปรเจกต์นี้ วางแทน
3. กด Ctrl+S (หรือไอคอนบันทึก) เพื่อบันทึกโปรเจกต์สคริปต์

## ขั้นตอนที่ 3: รัน setupSheets() ครั้งแรก
1. ที่แถบด้านบนของตัวแก้ไขสคริปต์ เลือกฟังก์ชัน `setupSheets` จาก dropdown
2. กด **Run (▶)**
3. ครั้งแรกจะมีหน้าต่างขอสิทธิ์ (Authorization) → เลือกบัญชี Google ของคุณ → กด "Advanced/ขั้นสูง" → "Go to ... (unsafe)" → Allow
4. กลับไปที่ Google Sheet จะเห็นชีต `Products`, `Partners`, `Transactions`, `Settings`, `Users`, `ActivityLogs` ถูกสร้างพร้อมข้อมูลตัวอย่าง และบัญชี admin เริ่มต้น (`admin` / `Admin@1234`)

## ขั้นตอนที่ 4: (ไม่บังคับ) เปิดใช้งานแจ้งเตือนอัตโนมัติ
1. เลือกฟังก์ชัน `createTimeTriggers` จาก dropdown แล้วกด Run หนึ่งครั้ง
2. จะมีการสร้าง Time-driven Trigger ให้รัน `dailyNotificationJob` ทุกวัน 08:00 น. และ `lowStockNotificationJob` ทุกวัน 09:00 น. ตามเวลาของสคริปต์
3. เปิด/ปิดการแจ้งเตือนรายวัน/สัปดาห์/เดือน และตั้งเกณฑ์สินค้าใกล้หมด ได้ที่หน้า "ตั้งค่า" ในเว็บ (หลัง deploy)

## ขั้นตอนที่ 5: Deploy เป็น Web App
1. มุมขวาบนของตัวแก้ไขสคริปต์ กด **Deploy → New deployment**
2. คลิกไอคอนเฟือง ข้าง "Select type" → เลือก **Web app**
3. ตั้งค่า:
   - Description: ระบบยืม-คืนสินค้า v1
   - Execute as: **Me (บัญชีของคุณ)**
   - Who has access: **Anyone**
4. กด **Deploy**
5. ระบบจะขอ Authorize อีกครั้ง (ถ้ายังไม่เคย) → อนุญาตทั้งหมด
6. คัดลอก **Web app URL** ที่ได้ (รูปแบบ `https://script.google.com/macros/s/XXXXXXXX/exec`)

> หมายเหตุ: ทุกครั้งที่แก้ไข `Code.gs` ภายหลัง ต้องกด **Deploy → Manage deployments → แก้ไข (ไอคอนดินสอ) → เปลี่ยน Version เป็น "New version" → Deploy** เพื่อให้โค้ดที่แก้ไขมีผลจริงกับ URL เดิม

## ขั้นตอนที่ 6: ใส่ API_URL ใน index.html
1. เปิดไฟล์ `google-apps-script/index.html`
2. ค้นหาบรรทัด:
   ```js
   const API_URL = 'https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/exec';
   ```
3. แทนที่ด้วย Web app URL ที่ได้จากขั้นตอนที่ 5
4. (ไม่บังคับ) ใส่ URL โลโก้ที่ตัวแปร `LOGO_URL` ถ้าต้องการแสดงโลโก้ระบบ

## ขั้นตอนที่ 7: อัปโหลดขึ้น GitHub Pages
1. สร้าง repository ใหม่บน GitHub เช่น `borrow-return-system`
2. อัปโหลดไฟล์ `index.html` (ที่แก้ไข API_URL แล้ว) ไว้ที่ root ของ repository
3. ไปที่ **Settings → Pages**
4. Source: เลือก **Deploy from a branch** → Branch: `main` → Folder: `/ (root)` → Save
5. รอสักครู่ จะได้ URL เช่น `https://<username>.github.io/borrow-return-system/`
6. เปิด URL นั้นในมือถือ/คอมพิวเตอร์ เพื่อทดสอบใช้งานระบบ

## ขั้นตอนที่ 8: ทดสอบระบบ
1. เข้าสู่ระบบด้วย `admin` / `Admin@1234` → ระบบจะบังคับให้ตั้งรหัสผ่านใหม่ทันที
2. ไปที่เมนู "ตั้งค่า → ผู้ใช้งาน" สร้างบัญชี staff/salesperson/manager เพิ่มตามต้องการ
3. ไปที่เมนู "สแกน/บันทึกยืม-คืน" → อนุญาตสิทธิ์กล้อง → ลองสแกน QR ของสินค้า (เปิดคอลัมน์ QR ในชีต Products แล้วใช้มือถืออีกเครื่องสแกนภาพ QR บนจอ หรือพิมพ์รหัส เช่น `P001` ที่ช่องรหัสแล้วกด "ตรวจสอบ")
4. ทดสอบ **บันทึกข้อมูล** แล้วกด **พิมพ์ใบยืม/คืน (PDF)**
5. ทดสอบฟอนต์ไทยใน PDF: เปิดไฟล์ PDF ที่ดาวน์โหลดมา ตรวจว่าอักษรไทยในหัวใบ/ตารางรายการ/ช่องเซ็นชื่อแสดงผลถูกต้อง ไม่เป็นสี่เหลี่ยม (กล่องโหว่)
   - หากฟอนต์ไทยแสดงผลไม่สวย/ไม่ครบ ให้ลองปรับ `font-family` ใน `buildPdfTemplateString()` ภายใน `Code.gs` เช่นสลับลำดับเป็น `"Noto Sans Thai","Leelawadee UI","Sarabun",Tahoma,sans-serif` แล้ว deploy เวอร์ชันใหม่อีกครั้ง เนื่องจาก PDF engine ของ Google Apps Script render จากฟอนต์ระบบ ไม่ใช่ดึง Google Fonts ตรง ๆ เหมือนหน้าเว็บ
6. ทดสอบหน้า "ค้นหารายการยืม-คืน" และ "ค้นหายอดคงเหลือ" ด้วยบัญชี staff/salesperson/manager ตามสิทธิ์
7. ทดสอบหน้า "Dashboard" ด้วยบัญชี admin/manager (manager ต้องไม่เห็นปุ่มแก้ไข/ตั้งค่าใด ๆ)
8. ทดสอบตั้งค่า LINE/Telegram/Email ที่หน้า "ตั้งค่า" แล้วกดปุ่ม "ทดสอบส่ง" แต่ละช่องทาง

## ข้อจำกัดที่ควรทราบ
- PWA: เนื่องจากข้อกำหนดอนุญาตไฟล์ได้แค่ 2 ไฟล์ (ไม่มีไฟล์ Service Worker แยก) ระบบจึงรองรับ `manifest.json` (ผ่าน endpoint `?page=manifest`) และ meta tag สำหรับ "Add to Home Screen" แต่จะไม่สามารถทำงานแบบออฟไลน์เต็มรูปแบบ (offline caching) ได้ เนื่องจาก Service Worker ต้องเป็นไฟล์ .js แยกต่างหากที่เสิร์ฟจากโดเมนเดียวกัน
- Session token เก็บในตัวแปร JavaScript ล้วน (ไม่ใช้ localStorage/sessionStorage) ตามข้อกำหนดความปลอดภัย ดังนั้นเมื่อรีโหลดหน้าเว็บ ผู้ใช้ต้องเข้าสู่ระบบใหม่ทุกครั้ง (เป็นพฤติกรรมที่ตั้งใจออกแบบไว้)
- หากเปิดเว็บแล้ว manifest ไม่โหลด (บางเบราว์เซอร์บังคับ CORS กับ manifest cross-origin) จะไม่กระทบการใช้งานหลักของระบบ เป็นเพียงความสามารถ "ติดตั้งเป็นแอป" ที่อาจไม่สมบูรณ์ 100% บนบาง browser
