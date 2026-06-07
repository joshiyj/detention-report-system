/**
 * reportController2.js
 * Handles upload, generate, and download endpoints for 2nd Year (4th Semester).
 * Supports multiple Excel files (one per section/division) — results are merged
 * into a single combined report across all four tables.
 *
 * Excel file format (A updated 23.xls, B updated 23.xls, etc.):
 *   Col 0: Roll No.    (e.g. "B-10")
 *   Col 1: Unique Id
 *   Col 2: Seat No     (e.g. "UECS26RS40022")   ← used for sorting Tables I & II
 *   Col 3: Student name
 *   Col 4: OverAll Attendance
 *   Col 5+: Subject columns (ESD, OS, DAA, AIML, MC, CMT, EVAC, BFS, IPD, WS, INC, FED, IE, OE II, ESD LAB, OS LAB, AIML LAB, SL I, BCC, IE Lab)
 *   Null subject cell = student NOT enrolled in that course → skip
 *
 * Sorting rules:
 *   - Table I & II:     sorted by Exam Seat No
 *   - Table III(A):     sorted by SN from Master_Course_Mapping_Table.docx
 *   - Table III(B):     sorted by Exam Seat No within merged result
 */

const path = require('path');
const fs   = require('fs');
const { parseExcel }               = require('../services/excelParser');       // reuse 1st year parser (same column structure)
const { processAttendanceY2new }   = require('../services/attendanceProcessorY2');
const { generateWordY2 }           = require('../services/wordGeneratorY2');
const { parseDocxMapping }          = require('../services/docxMappingParser');
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
 * Courses with the same code have their student rows merged.
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
 * POST /api/y2/generate
 */
async function generateReport2(req, res) {
  try {
    const files = req.files ? req.files['files'] : null;
    const mappingFile = req.files ? req.files['mappingFile']?.[0] : null;

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No Excel files uploaded.' });
    }

    const { examName, schoolName, programme, semester, date, condonation, coordinatorName, hodName, branchName } = req.body;

    if (!examName || !schoolName || !programme || !semester || !date) {
      return res.status(400).json({ error: 'All metadata fields are required.' });
    }

    // Parse optional dynamic mapping docx
    let dynamicMapping = null;
    if (mappingFile) {
      dynamicMapping = await parseDocxMapping(mappingFile.buffer);
      if (dynamicMapping) {
        saveMapping('y2', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('y2');
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
      // Parse using 1st-year parser — same 5-column structure (Roll No, Unique Id, Seat No, Name, Overall)
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReport2] No students in: ${file.originalname}`);
        continue;
      }

      const { tableI, tableII, tableIIIA, tableIIIB } =
        processAttendanceY2new(students, subjectNames, condonationSeats, dynamicMapping);

      mergedTableI    = mergedTableI.concat(tableI);
      mergedTableII   = mergedTableII.concat(tableII);
      mergedTableIIIA = mergeTableIIIA(mergedTableIIIA, tableIIIA);
      mergedTableIIIB = mergedTableIIIB.concat(tableIIIB);
      totalStudents  += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    // ── Sort Table I & II by Seat No ──────────────────────────────────────────
    mergedTableI.sort(sortBySeatNo);
    mergedTableII.sort(sortBySeatNo);

    // ── Sort Table III(B) by Seat No ──────────────────────────────────────────
    mergedTableIIIB.sort(sortBySeatNo);

    // ── Sort rows within each Table III(A) course by Seat No ─────────────────
    for (const course of mergedTableIIIA) {
      course.rows.sort(sortBySeatNo);
    }
    // tableIIIA itself is already sorted by SN (done in processor + mergeTableIIIA)

    const stats = {
      totalStudents,
      tableICount:      mergedTableI.length,
      tableIICount:     mergedTableII.length,
      detainedStudents: mergedTableIIIB.length,
      detainedSubjects: mergedTableIIIA.reduce((acc, c) => acc + c.rows.length, 0),
    };

    const timestamp  = Date.now();
    const filename   = `detention_report_y2_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWordY2(
      { examName, schoolName, programme, semester, date, coordinatorName, hodName, branchName },
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
    console.error('[generateReport2]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

module.exports = { generateReport2 };
