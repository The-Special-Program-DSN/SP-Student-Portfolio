# SP Student Portfolio

เว็บสำหรับนักเรียนโครงการห้องเรียนพิเศษ โรงเรียนเทพศิรินทร์ นนทบุรี  
Frontend: **GitHub Pages** · Backend: **Google Apps Script** · Database: **Google Sheets**

## ตั้งค่า Google Apps Script Backend
1. เปิด Google Sheet ฐานข้อมูล
2. ไปที่ **Extensions → Apps Script**
3. วางโค้ดจาก `backend/Code.gs`
4. เปิด **Project Settings → Script Properties** และเพิ่ม:
   - `ADMIN_USER`
   - `ADMIN_PASSWORD`
   - `YEAR_CHANGE_CODE`
5. Deploy เป็น **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. คัดลอก URL ที่ลงท้ายด้วย `/exec`
7. ใส่ URL ลงใน `js/config.js`

> ห้ามใส่รหัสผ่านจริงลงใน GitHub repository

## GitHub Pages
ไปที่ **Settings → Pages → Deploy from a branch → main / (root)**

## หมายเหตุด้านข้อมูล
ระบบมีข้อมูลส่วนบุคคล เช่น เลขบัตรประชาชนและวันเดือนปีเกิด จึงไม่ควรตั้ง Google Sheet เป็น Public
