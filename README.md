# HQ Migration Progress Dashboard

Dashboard สำหรับ project HQ (สำนักงานใหญ่)

## Structure
```
hq-app/
├── backend/
│   ├── server.js         ← Express API (port 3001)
│   ├── package.json
│   └── SAT_Progress.xlsx ← Excel fallback
└── frontend/
    └── index.html        ← Single-page dashboard
```

## Deploy on Render.com

### Backend (Web Service)
- Root: `backend`
- Build: `npm install`
- Start: `node server.js`
- Env: `SHAREPOINT_URL` = your SharePoint share link + `&download=1`
- Service name: `hq-migr-api`

### Frontend (Static Site)
- Root: `frontend`
- Publish: `.`
- Service name: `hq-migr-progress`

## API Endpoints
- `GET  /api/summary` — overview + weekly + daily data
- `GET  /api/devices` — device list (filter: location, status)
- `POST /api/cache/refresh` — force re-fetch from SharePoint
- `GET  /health` — health check
