# RBU Detention Report Generator

A full-stack MERN web application for Ramdeobaba University that reads an Excel attendance sheet and generates a professionally formatted Word document containing all detention tables.

## Features

- Drag-and-drop Excel upload (`.xls` / `.xlsx`)
- Processes complex multi-row "pouring attendance" Excel format
- Generates **Table I** – Overall attendance < 75%
- Generates **Table II** – Condonation applicants (60–75%)
- Generates **Table III(A)** – Course-wise detention list
- Generates **Table III(B)** – Student-wise detention list
- Outputs a `.docx` matching the official RBU format
- Clean, responsive admin UI

---

## Project Structure

```
attendance-system/
├── client/               ← React + Vite + Tailwind CSS
│   └── src/
│       ├── App.jsx
│       └── services/api.js
└── server/               ← Node.js + Express
    ├── server.js
    ├── routes/reportRoutes.js
    ├── controllers/reportController.js
    └── services/
        ├── excelParser.js          ← Excel parsing logic
        ├── attendanceProcessor.js  ← Detention rule engine
        └── wordGenerator.js        ← .docx generator
```

---

## Quick Start

### Backend

```bash
cd server
npm install
node server.js        # starts on port 5000
```

### Frontend (development)

```bash
cd client
npm install
npm run dev           # starts on port 5173, proxies /api → :5000
```

### Frontend (production build)

```bash
cd client
npm run build
# Serve dist/ with any static server, or let Express serve it
```

---

## API

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/generate` | Upload Excel + metadata, returns filename + stats |
| `GET`  | `/api/download/:filename` | Download generated .docx |

### POST /api/generate (multipart/form-data)

| Field | Type | Required |
|-------|------|----------|
| `file` | Excel file | ✅ |
| `examName` | string | ✅ |
| `schoolName` | string | ✅ |
| `programme` | string | ✅ |
| `semester` | string | ✅ |
| `date` | string | ✅ |
| `condonation` | comma-separated seat numbers | ❌ |

---

## Attendance Rules

| Rule | Threshold |
|------|-----------|
| Overall detention | < 75% |
| Condonation window | 60% – 74.99% |
| Subject-wise detention | < 60% in any subject |

---

## Excel Format Expected

Row 0: Headers — `Roll No. | Unique Id | Seat No | Student name | OverAll Attendance | [subjects...]`

Each student may span 2–3 rows (pouring attendance). The parser handles:
- Merged rows
- `"pouring Attendance :28.0/ 33 (84.85)"` format
- `"10 / 24 ( 41.67 % )"` format
- Empty separator rows
- `"97/274 (35.40)"` overall attendance format

---

## Tech Stack

- **Frontend**: React 18, Vite, Tailwind CSS, Axios
- **Backend**: Node.js, Express, Multer, XLSX, docx
