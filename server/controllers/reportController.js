/**
 * reportController.js
 * Handles upload, generate, and download endpoints.
 * Supports multiple Excel files (one per batch/division) — results are merged
 * into a single combined report across all four tables.
 */

const path  = require('path');
const fs    = require('fs');
const { parseExcel }                    = require('../services/excelParser');
const { processAttendance, COURSE_MAP } = require('../services/attendanceProcessor');
const { generateWord }                  = require('../services/wordGenerator');
const { parseWordMapping, parseDocxMapping }              = require('../services/docxMappingParser');
const { getMapping, saveMapping }       = require('../services/mappingPersistence');

// ── Canonical course order matching SHORTFORM.docx (SN 1-9) ──────────────────
const SHORTFORM_ORDER = [
  '25EE01TP0201::Theory',    // SN 1 – MI (L)
  '25EE01TP0201::Practical', // SN 1 – MI (P)
  '25HS03TH0213::Theory',    // SN 2 – CLA
  '25EE01TP0202::Theory',    // SN 3 – PPS (L)
  '25EE01TP0202::Practical', // SN 3 – PPS (P)
  '25EE01TH0203::Theory',    // SN 4 – AIML
  '25EE01TP0204::Theory',    // SN 5 – CAO (L)
  '25EE01TP0204::Practical', // SN 5 – CAO (P)
  '25HS02TP0201::Theory',    // SN 6 – EPC (L)
  '25HS02TP0201::Practical', // SN 6 – EPC (P)
  '25HS02TH0203-1::Theory',  // SN 7 – FLIC
  '25EE01PR0205::Practical', // SN 8 – MP1
  '25HS04PR0201::Practical', // SN 9 – HFW
];

function shortformIndex(code, type, activeShortformOrder) {
  const order = activeShortformOrder || SHORTFORM_ORDER;
  const idx = order.indexOf(`${code}::${type}`);
  return idx === -1 ? order.length : idx;
}

function shortformCodeIndex(code, activeShortformOrder) {
  const order = activeShortformOrder || SHORTFORM_ORDER;
  for (let i = 0; i < order.length; i++) {
    if (order[i].startsWith(`${code}::`)) return i;
  }
  return order.length;
}

const GENERATED_DIR = path.join(__dirname, '..', 'generated');

// ── Merge helpers ─────────────────────────────────────────────────────────────

/**
 * Merge two tableIIIA arrays (each item = { code, fullName, rows[] }).
 * Courses with the same code have their rows concatenated; duplicate seatNos are
 * de-duplicated by keeping the entry with the most information.
 */
function mergeTableIIIA(a, b, activeShortformOrder) {
  const map = {};
  for (const item of [...a, ...b]) {
    const key = item.code;
    if (!map[key]) {
      map[key] = { code: item.code, fullName: item.fullName, rows: [] };
    }
    // Merge rows: if a row for the same student already exists, combine pct strings
    for (const row of item.rows) {
      const existing = map[key].rows.find(r => r.seatNo === row.seatNo);
      if (existing) {
        // Combine if not already combined
        if (existing.pct !== row.pct) {
          existing.pct  = `${existing.pct}/${row.pct}`;
          existing.type = 'Theory/Practical';
        }
      } else {
        map[key].rows.push({ ...row });
      }
    }
  }
  // Re-sort by SHORTFORM.docx order (SN 1-9)
  return Object.values(map)
    .sort((x, y) => shortformCodeIndex(x.code, activeShortformOrder) - shortformCodeIndex(y.code, activeShortformOrder));
}

/**
 * POST /api/generate
 * Body (multipart/form-data):
 *   files[]     – One or more Excel files (one per batch/division)
 *   examName    – string
 *   schoolName  – string
 *   programme   – string
 *   semester    – string
 *   date        – string
 *   condonation – comma-separated seat numbers (optional)
 *   mappingFile - optional .docx mapping file
 */
async function generateReport(req, res) {
  try {
    const files = req.files ? req.files['files'] : null;
    const mappingFile = req.files ? req.files['mappingFile']?.[0] : null;

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No Excel files uploaded.' });
    }

    const { examName, schoolName, programme, semester, date, condonation } = req.body;

    if (!examName || !schoolName || !programme || !semester || !date) {
      return res.status(400).json({ error: 'All metadata fields are required.' });
    }

    // Parse optional dynamic mapping docx
    let dynamicMapping = null;
    if (mappingFile) {
      dynamicMapping = await parseDocxMapping(mappingFile.buffer);
      if (dynamicMapping) {
        saveMapping('y1', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('y1');
    }
    const activeShortformOrder = dynamicMapping ? dynamicMapping.shortformOrder : SHORTFORM_ORDER;

    // Parse condonation seat list
    const condonationSeats = new Set(
      (condonation || '')
        .split(',')
        .map(s => s.trim())
        .filter(Boolean)
    );

    // ── Process each file and accumulate merged results ───────────────────────
    let mergedTableI   = [];
    let mergedTableII  = [];
    let mergedTableIIIA = [];
    let mergedTableIIIB = [];
    let totalStudents  = 0;

    for (const file of files) {
      // 1. Parse Excel
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReport] No students found in file: ${file.originalname}`);
        continue;
      }

      // 2. Process attendance rules
      const { tableI, tableII, tableIIIA, tableIIIB } =
        processAttendance(students, subjectNames, condonationSeats, dynamicMapping);

      // 3. Accumulate
      mergedTableI    = mergedTableI.concat(tableI);
      mergedTableII   = mergedTableII.concat(tableII);
      mergedTableIIIA = mergeTableIIIA(mergedTableIIIA, tableIIIA, activeShortformOrder);
      mergedTableIIIB = mergedTableIIIB.concat(tableIIIB);
      totalStudents  += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    const parseRoll = (rollNo = '') => {
      const m = rollNo.match(/^([A-Za-z]+\d+)[_-](\d+)$/);
      return m
        ? { section: m[1].toUpperCase(), num: parseInt(m[2], 10) }
        : { section: rollNo.toUpperCase(), num: 0 };
    };
    const sortByRoll = (a, b) => {
      const ra = parseRoll(a.rollNo);
      const rb = parseRoll(b.rollNo);
      if (ra.section < rb.section) return -1;
      if (ra.section > rb.section) return  1;
      return ra.num - rb.num;           // numeric comparison within same section
    };
    mergedTableI.sort(sortByRoll);
    mergedTableII.sort(sortByRoll);
    mergedTableIIIB.sort(sortByRoll);

    // Sort students within each Course Code in Table III(A) by Roll No
    for (const course of mergedTableIIIA) {
      course.rows.sort(sortByRoll);
    }

    // ── Compute combined stats ────────────────────────────────────────────────
    const stats = {
      totalStudents,
      tableICount:      mergedTableI.length,
      tableIICount:     mergedTableII.length,
      detainedStudents: mergedTableIIIB.length,
      detainedSubjects: mergedTableIIIA.reduce((acc, c) => acc + c.rows.length, 0),
    };

    // ── Generate Word doc ─────────────────────────────────────────────────────
    const timestamp  = Date.now();
    const filename   = `detention_report_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWord(
      { examName, schoolName, programme, semester, date },
      {
        tableI:   mergedTableI,
        tableII:  mergedTableII,
        tableIIIA: mergedTableIIIA,
        tableIIIB: mergedTableIIIB,
      },
      outputPath
    );

    res.json({ success: true, filename, stats });

  } catch (err) {
    console.error('[generateReport]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

/**
 * GET /api/download/:filename
 */
function downloadReport(req, res) {
  const { filename } = req.params;

  // Basic security: no path traversal
  if (filename.includes('/') || filename.includes('..')) {
    return res.status(400).json({ error: 'Invalid filename.' });
  }

  const filePath = path.join(GENERATED_DIR, filename);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found.' });
  }

  res.download(filePath, filename, err => {
    if (err) {
      console.error('[downloadReport]', err);
      res.status(500).json({ error: 'Download failed.' });
    }
  });
}

module.exports = { generateReport, downloadReport };
