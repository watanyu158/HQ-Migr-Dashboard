# SDA Installation Dashboard

Web App แสดง Dashboard ติดตาม SDA Installation สร้างจาก Excel ต้นฉบับ

---

## โครงสร้างโปรเจกต์

```
sda-app/
├── backend/
│   ├── server.js                   ← Express API
│   ├── SDA_Installation_Plan_V2.xlsx  ← Excel ต้นฉบับ
│   └── package.json
└── frontend/
    └── index.html                  ← Dashboard (Single Page)
```

---

## วิธี Deploy บน Render.com (ฟรี)

### ขั้นตอนที่ 1 — เตรียม GitHub

1. สร้าง repo ใหม่บน GitHub (เช่น `sda-dashboard`)
2. อัพโหลดทั้ง `backend/` และ `frontend/` ขึ้น repo

```
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR_NAME/sda-dashboard.git
git push -u origin main
```

### ขั้นตอนที่ 2 — Deploy Backend (Web Service)

1. ไป https://render.com → **New → Web Service**
2. เชื่อม GitHub repo
3. ตั้งค่า:
   - **Root Directory**: `backend`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Instance Type**: Free
4. กด **Create Web Service**
5. รอ ~2 นาที → ได้ URL เช่น `https://sda-api.onrender.com`

### ขั้นตอนที่ 3 — แก้ URL ใน Frontend

เปิด `frontend/index.html` แก้บรรทัด:
```js
: 'https://YOUR-RENDER-API-URL.onrender.com';
```
เปลี่ยนเป็น URL จริง เช่น:
```js
: 'https://sda-api.onrender.com';
```

### ขั้นตอนที่ 4 — Deploy Frontend (Static Site)

1. Render.com → **New → Static Site**
2. เชื่อม GitHub repo เดิม
3. ตั้งค่า:
   - **Root Directory**: `frontend`
   - **Publish Directory**: `.`  (หรือเว้นว่าง)
4. กด **Create Static Site**
5. ได้ URL เช่น `https://sda-dashboard.onrender.com`

---

## API Endpoints

| Method | Endpoint | คำอธิบาย |
|--------|----------|-----------|
| GET | `/api/summary` | ภาพรวม + progress แต่ละ Fabric |
| GET | `/api/devices` | รายการอุปกรณ์ (รองรับ filter + pagination) |
| GET | `/api/devices?fabric=D1-041&status=Success` | filter ตาม Fabric+Status |
| GET | `/api/fabric/:name` | รายละเอียด Fabric เดียว |
| GET | `/api/filters` | ค่า dropdown สำหรับ filter |
| POST | `/api/update` | อัพเดทสถานะการติดตั้งกลับ Excel |
| GET | `/health` | ตรวจสอบสถานะ server |

---

## ข้อจำกัด Render Free Tier

- Backend หยุดทำงานหลังไม่มีคนใช้ **15 นาที** (cold start ~30 วินาที)
- ข้อมูลที่เขียนกลับ Excel **หายเมื่อ redeploy** (ไม่มี persistent storage)
- แนะนำ: เก็บข้อมูลแยกใน **Google Sheets API** หรือ **Supabase** หากต้องการถาวร

---

## ทดสอบ Local

```bash
cd backend
npm install
node server.js
# เปิด frontend/index.html ด้วย browser
```
