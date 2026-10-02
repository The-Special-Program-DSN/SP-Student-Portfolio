# SP Student Portfolio

ระบบสำหรับนักเรียนโครงการห้องเรียนพิเศษ โรงเรียนเทพศิรินทร์ นนทบุรี

- Frontend: GitHub Pages
- Backend: Google Apps Script
- Database: Google Sheets

## Backend
ใช้ไฟล์:
- `backend/Code.gs`
- `backend/Helpers.gs`

นำทั้ง 2 ไฟล์ไปใส่ใน Apps Script โปรเจกต์เดียวกัน แล้วตั้ง Script Properties:
- `ADMIN_USER`
- `ADMIN_PASSWORD`
- `YEAR_CHANGE_CODE`

จากนั้น Deploy เป็น Web app:
- Execute as: Me
- Who has access: Anyone

นำ URL ที่ลงท้ายด้วย `/exec` ไปใส่ใน `js/config.js`.

## GitHub Pages
เปิด Settings → Pages แล้วเลือก:
- Deploy from a branch
- Branch: `main`
- Folder: `/(root)`

> Repository นี้ไม่ควรเก็บรหัสผ่านหรือรหัสยืนยันไว้ใน source code
> Google Sheet มีข้อมูลส่วนบุคคล จึงไม่ควรตั้ง Sheet เป็นสาธารณะ
