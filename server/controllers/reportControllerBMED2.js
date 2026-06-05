/**
 * reportControllerBMED2.js
 * Handles upload, generate, and download endpoints for 2nd Year Biomedical Engineering (4th Semester).
 * Supports multiple Excel files (one per section/division) — results are merged
 * into a single combined report across all four tables.
 *
 * Excel file format (IV SEM BME.xls, IV SEM BME_1.xls):
 *   Col 0: Roll No.    (e.g. "21")
 *   Col 1: Unique Id
 *   Col 2: Seat No     (e.g. "UBEE26RS40004")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance
 *   Col 5+: Subject columns (HAPE II, FAIML, SPA, BMECH, MIOT, PEIoT, INC, FED, IE, OE II, FAIML LAB, SPA LAB, DF3D, BCC, IE Lab)
 *
 * Sorting rules:
 *   - Table I & II:   sorted by Exam Seat No
 *   - Table III(A):   sorted by SN from Biomedical_Sem4_Course_Mapping.docx
 *   - Table III(B):   sorted by Exam Seat No within merged result
 */

const path = require('path');
const fs   = require('fs');
const { parseExcel }              = require('../services/excelParser');   // same column structure as ECS Y2-Y4
const { processAttendanceBMED2 }  = require('../services/attendanceProcessorBMED2');
const { generateWordBMED2 }       = require('../services/wordGeneratorBMED2');
const { parseDocxMapping }        = require('../services/docxMappingParser');
const { getMapping, saveMapping }   = require('../services/mappingPersistence');

const GENERATED_DIR = path.join(__dirname, '..', 'generated');

// ── Sort helpers ──────────────────────────────────────────────────────────────
function sortBySeatNo(a, b) {
  const sa = (a.seatNo || a.rollNo || '').toUpperCase();
  const sb = (b.seatNo || b.rollNo || '').toUpperCase();
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/**
 * Merge two tableIIIA arrays sorted by SN.
 * Courses with the same code have their student rows merged (no duplicates).
 */
function mergeTableIIIA(a, b) {
  const map = {};
  for (const item of [...a, ...b]) {
    const key = item.code;
    if (!map[key]) {
      map[key] = { code: item.code, fullName: item.fullName, sn: item.sn, rows: [] };
    }
    for (const row of item.rows) {
      const existing = map[key].rows.find(r => (r.seatNo || r.rollNo) === (row.seatNo || row.rollNo));
      if (!existing) {
        map[key].rows.push({ ...row });
      }
    }
  }
  return Object.values(map).sort((x, y) => x.sn - y.sn);
}

/**
 * POST /api/bmed2/generate
 */
async function generateReportBMED2(req, res) {
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
        saveMapping('bmed2', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('bmed2');
    }

    const condonationSeats = new Set(
      (condonation || '').split(',').map(s => s.trim()).filter(Boolean)
    );

    let mergedTableI    = [];
    let mergedTableII   = [];
    let mergedTableIIIA = [];
    let mergedTableIIIB = [];
    let totalStudents   = 0;

    for (const file of files) {
      // Biomedical Excel has same structure as ECS (Roll No, Unique Id, Seat No, Name, Overall)
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReportBMED2] No students in: ${file.originalname}`);
        continue;
      }

      const { tableI, tableII, tableIIIA, tableIIIB } =
        processAttendanceBMED2(students, subjectNames, condonationSeats, dynamicMapping);

      mergedTableI    = mergedTableI.concat(tableI);
      mergedTableII   = mergedTableII.concat(tableII);
      mergedTableIIIA = mergeTableIIIA(mergedTableIIIA, tableIIIA);
      mergedTableIIIB = mergedTableIIIB.concat(tableIIIB);
      totalStudents  += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    // ── Deduplicate Table I by seatNo ─────────────────────────────────────────
    const seenI = new Set();
    mergedTableI = mergedTableI.filter(r => {
      const key = r.seatNo || r.rollNo;
      if (seenI.has(key)) return false;
      seenI.add(key);
      return true;
    });

    // ── Sort Table I & II by Seat No ──────────────────────────────────────────
    mergedTableI.sort(sortBySeatNo);
    mergedTableII.sort(sortBySeatNo);

    // ── Sort Table III(B) by Seat No ──────────────────────────────────────────
    mergedTableIIIB.sort(sortBySeatNo);

    // ── Sort rows within each Table III(A) course by Seat No ─────────────────
    for (const course of mergedTableIIIA) {
      course.rows.sort(sortBySeatNo);
    }

    const stats = {
      totalStudents,
      tableICount:      mergedTableI.length,
      tableIICount:     mergedTableII.length,
      detainedStudents: mergedTableIIIB.length,
      detainedSubjects: mergedTableIIIA.reduce((acc, c) => acc + c.rows.length, 0),
    };

    // Ensure output directory exists
    if (!fs.existsSync(GENERATED_DIR)) {
      fs.mkdirSync(GENERATED_DIR, { recursive: true });
    }

    const timestamp  = Date.now();
    const filename   = `detention_report_bmed2_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWordBMED2(
      { examName, schoolName, programme, semester, date },
      {
        tableI:    mergedTableI,
        tableII:   mergedTableII,
        tableIIIA: mergedTableIIIA,
        tableIIIB: mergedTableIIIB,
      },
      outputPath
    );

    res.json({ success: true, filename, stats });

  } catch (err) {
    console.error('[generateReportBMED2]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

module.exports = { generateReportBMED2 };
