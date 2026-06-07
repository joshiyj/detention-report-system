/**
 * reportController3.js
 * Handles upload, generate, and download endpoints for 3rd Year (6th Semester) ECS.
 * Supports multiple Excel files (one per section/division).
 *
 * Excel file format (VI A updated 22.xls, etc.):
 *   Col 0: Roll No.    (e.g. "A-67")
 *   Col 1: Unique Id
 *   Col 2: Seat No     (e.g. "ENCS26RS6020")  ← used for sorting Tables I, II & III
 *   Col 3: Student name
 *   Col 4: OverAll Attendance
 *   Col 5+: Subject columns
 *   Null subject cell = student NOT enrolled in that course → skip
 *
 * Sorting rules:
 *   - Tables I, II & III: sorted by Exam Seat No
 */

const path = require('path');
const fs   = require('fs');
const { parseExcel }             = require('../services/excelParser');     // reuse 1st-year parser (same 5-col structure)
const { processAttendanceY3 }    = require('../services/attendanceProcessorY3');
const { generateWordY3 }         = require('../services/wordGeneratorY3');
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
 * POST /api/y3/generate
 */
async function generateReport3(req, res) {
  try {
    const files = req.files ? req.files['files'] : null;
    const mappingFile = req.files ? req.files['mappingFile']?.[0] : null;

    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No Excel files uploaded.' });
    }

    const { examName, schoolName, programme, semester, date, condonation, coordinatorName, hodName, branchName, departmentName } = req.body;

    if (!examName || !schoolName || !programme || !semester || !date) {
      return res.status(400).json({ error: 'All metadata fields are required.' });
    }

    // Parse optional dynamic mapping docx
    let dynamicMapping = null;
    if (mappingFile) {
      dynamicMapping = await parseDocxMapping(mappingFile.buffer);
      if (dynamicMapping) {
        saveMapping('y3', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('y3');
    }

    const condonationSeats = new Set(
      (condonation || '').split(',').map(s => s.trim()).filter(Boolean)
    );

    let mergedTableI   = [];
    let mergedTableII  = [];
    let mergedTableIII = [];
    let totalStudents  = 0;

    for (const file of files) {
      // Parse using 1st-year parser (same 5-column leading structure)
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReport3] No students in: ${file.originalname}`);
        continue;
      }

      const { tableI, tableII, tableIII } =
        processAttendanceY3(students, subjectNames, condonationSeats, dynamicMapping);

      mergedTableI   = mergedTableI.concat(tableI);
      mergedTableII  = mergedTableII.concat(tableII);
      mergedTableIII = mergedTableIII.concat(tableIII);
      totalStudents += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    // ── Sort all tables by Seat No ─────────────────────────────────────────────
    mergedTableI.sort(sortBySeatNo);
    mergedTableII.sort(sortBySeatNo);
    mergedTableIII.sort(sortBySeatNo);

    const stats = {
      totalStudents,
      tableICount:      mergedTableI.length,
      tableIICount:     mergedTableII.length,
      detainedStudents: mergedTableIII.length,
      detainedCourses:  mergedTableIII.reduce((acc, s) => acc + s.courses.length, 0),
    };

    const timestamp  = Date.now();
    const filename   = `detention_report_y3_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWordY3(
      { examName, schoolName, programme, semester, date, coordinatorName, hodName, branchName, departmentName },
      {
        tableI:   mergedTableI,
        tableII:  mergedTableII,
        tableIII: mergedTableIII,
      },
      outputPath
    );

    res.json({ success: true, filename, stats });

  } catch (err) {
    console.error('[generateReport3]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

module.exports = { generateReport3 };
