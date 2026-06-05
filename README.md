# RBU Detention Report Generator

A full-stack web application for Ramdeobaba University (RBU) that parses Excel attendance sheets and generates professionally formatted Word documents (`.docx`) containing official university detention tables.

---

## Features

- **Drag-and-Drop Excel Upload** (`.xls` / `.xlsx`).
- **Support for Multiple Academic Years & Branches**:
  - Year 1 (Generic ECS)
  - Year 2 (ECS 4th Semester)
  - Year 3 (ECS 6th Semester)
  - Year 4 (ECS 8th Semester)
  - BMED Year 2 (Biomedical 4th Semester)
  - BMED Year 3 (Biomedical 6th Semester)
- **Advanced Excel Parsing**: Robust row-merging engine that automatically reconciles multi-row vertical merges, empty spacer rows, and complex "pouring attendance" data.
- **Dynamic Course Mapping**: Upload optional `.docx` course mapping tables to overwrite default syllabus mappings on the fly.
- **Course Mapping Persistence**: Uploaded course mappings are cached locally on the server. Subsequent report generations for that program/semester reuse the saved mappings automatically without requiring another upload.
- **Rules & Detention Tables Engine**:
  - **Table I (Overall Detention)**: Students with aggregate attendance under 75%.
  - **Table II (Condonation)**: Students with aggregate attendance between 60% and 75% who applied for condonation.
  - **Table III (Subject-wise Detention)**: Students detained in specific courses (attendance < 60%).
    - *Note: For 3rd and 4th years, students with overall attendance >= 75% are automatically exempted from Table III.*
- **Customized Administrative Layouts**: Outputs docx files formatted with `Times New Roman`, proper header margins, college headers, and dual HOD/Dean signature blocks conforming to RBU specifications.

---

## Project Structure

```
detention-report-system/
├── client/                     ← React + Vite + Tailwind CSS Frontend
│   └── src/
│       ├── App.jsx             ← Primary user interface
│       ├── App2.jsx            ← Year 2 UI container
│       ├── App3.jsx            ← Year 3 UI container
│       ├── App4.jsx            ← Year 4 UI container
│       ├── AppBMED2.jsx        ← BMED Year 2 UI container
│       ├── AppBMED3.jsx        ← BMED Year 3 UI container
│       ├── LandingPage.jsx     ← Initial landing page
│       ├── BranchSelectionPage.jsx ← Year/branch selection menu
│       └── services/api.js      ← Client endpoints caller
└── server/                     ← Node.js + Express Backend
    ├── server.js               ← Main app entrypoint
    ├── routes/
    │   └── reportRoutes.js     ← Express router mapping routes to controllers
    ├── controllers/
    │   ├── reportController.js      ← Y1 controller
    │   ├── reportController2.js     ← Y2 controller
    │   ├── reportController3.js     ← Y3 controller
    │   ├── reportController4.js     ← Y4 controller
    │   ├── reportControllerBMED2.js ← BMED Y2 controller
    │   └── reportControllerBMED3.js ← BMED Y3 controller
    └── services/
        ├── excelParser.js           ← Merged row Excel parsing engine
        ├── attendanceProcessor.js   ← Core detention rules processor
        ├── attendanceProcessorY2.js  ← Year 2 rules processor
        ├── attendanceProcessorY3.js  ← Year 3 rules processor
        ├── attendanceProcessorY4.js  ← Year 4 rules processor
        ├── attendanceProcessorBMED2.js ← BMED Year 2 rules processor
        ├── attendanceProcessorBMED3.js ← BMED Year 3 rules processor
        ├── docxMappingParser.js     ← mammoth-based .docx mapping parsing
        ├── mappingPersistence.js    ← JSON-based persistent mapping cache
        ├── wordGenerator.js         ← Monolith Word document builder
        └── wordGeneratorY[2-4].js   ← Program/Year specific Word builders
```

---

## Quick Start

### 1. Start the Backend Server
```bash
cd server
npm install
node server.js        # Server starts on port 5000
```

### 2. Start the Frontend Dev Client
```bash
cd client
npm install
npm run dev           # Client starts on port 5173, proxying /api to port 5000
```

### 3. Production Build
```bash
cd client
npm run build
# Serves dist/ statically from production server
```

---

## API Router Endpoints

All route handlers expect `multipart/form-data` uploads containing Excel sheets and the metadata variables.

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/generate` | Generates 1st Year detention report |
| `POST` | `/api/y2/generate` | Generates 2nd Year (ECS) report |
| `POST` | `/api/y3/generate` | Generates 3rd Year (ECS) report |
| `POST` | `/api/y4/generate` | Generates 4th Year (ECS) report |
| `POST` | `/api/bmed2/generate` | Generates 2nd Year (BMED) report |
| `POST` | `/api/bmed3/generate` | Generates 3rd Year (BMED) report |
| `GET`  | `/api/download/:filename` | Downloads generated `.docx` file from cache |

### Fields expected in `POST` payload:

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `files` | file list (Excel) | ✅ | Raw student attendance Excel sheets |
| `mappingFile` | file (.docx) | ❌ | Dynamic Course Mapping file |
| `examName` | string | ✅ | Examination name (e.g. `Summer-2026`) |
| `schoolName` | string | ✅ | School name (e.g. `School of ECS`) |
| `programme` | string | ✅ | Degree programme (e.g. `B.Tech ECS`) |
| `semester` | string | ✅ | Semester (e.g. `Semester VIII`) |
| `date` | string | ✅ | Report generation date (e.g. `2026-06-05`) |
| `condonation` | string | ❌ | Comma-separated list of approved Seat/Roll Numbers |

---

## Core Detention Guidelines

| Parameter | Threshold | Scope |
|------|-----------|-------|
| Overall Detention (Table I) | < 75% | All Years |
| Condonation Band (Table II) | 60% – 74.99% | All Years |
| Subject Detention (Table III) | < 60% in subject | Year 1 & 2 (Always)<br>Year 3 & 4 (Exempted if Overall >= 75%) |
