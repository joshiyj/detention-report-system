/**
 * reportControllerBMED3.js
 * Handles upload, generate, and download endpoints for 3rd Year Biomedical Engineering (6th Semester).
 * Supports multiple Excel files (one per section/division) — results are merged
 * into a single combined report across all three tables.
 *
 * Excel file format (vi bme.xls):
 *   Col 0: Roll No.    (e.g. "38")
 *   Col 1: Unique Id
 *   Col 2: Seat No     (e.g. "BMU26RS6032")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance
 *   Col 5+: Subject columns
 *
 * Sorting rules:
 *   - Table I, II, & III: sorted by Exam Seat No
 */

const path = require('path');
const fs   = require('fs');
const { parseExcel }              = require('../services/excelParser');
const { processAttendanceBMED3 }  = require('../services/attendanceProcessorBMED3');
const { generateWordBMED3 }       = require('../services/wordGeneratorBMED3');
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
 * POST /api/bmed3/generate
 */
async function generateReportBMED3(req, res) {
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
        saveMapping('bmed3', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('bmed3');
    }

    const condonationSeats = new Set(
      (condonation || '').split(',').map(s => s.trim()).filter(Boolean)
    );

    let mergedTableI   = [];
    let mergedTableII  = [];
    let mergedTableIII = [];
    let totalStudents  = 0;

    for (const file of files) {
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReportBMED3] No students in: ${file.originalname}`);
        continue;
      }

      const { tableI, tableII, tableIII } =
        processAttendanceBMED3(students, subjectNames, condonationSeats, dynamicMapping);

      mergedTableI   = mergedTableI.concat(tableI);
      mergedTableII  = mergedTableII.concat(tableII);
      mergedTableIII = mergedTableIII.concat(tableIII);
      totalStudents += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    // Deduplicate mergedTableI by seatNo or rollNo
    const seenI = new Set();
    mergedTableI = mergedTableI.filter(r => {
      const key = r.seatNo || r.rollNo;
      if (seenI.has(key)) return false;
      seenI.add(key);
      return true;
    });

    // Deduplicate mergedTableII by seatNo or rollNo
    const seenII = new Set();
    mergedTableII = mergedTableII.filter(r => {
      const key = r.seatNo || r.rollNo;
      if (seenII.has(key)) return false;
      seenII.add(key);
      return true;
    });

    // Deduplicate mergedTableIII by seatNo or rollNo (merge courses if duplicate seats occur, though unlikely in a single run)
    const seenIII = new Map();
    for (const s of mergedTableIII) {
      const key = s.seatNo || s.rollNo;
      if (!seenIII.has(key)) {
        seenIII.set(key, { ...s, courses: [...s.courses] });
      } else {
        const existing = seenIII.get(key);
        // Merge courses without duplicates
        const existingCodes = new Set(existing.courses.map(c => c.code));
        for (const c of s.courses) {
          if (!existingCodes.has(c.code)) {
            existing.courses.push(c);
          }
        }
      }
    }
    mergedTableIII = Array.from(seenIII.values());

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

    // Ensure output directory exists
    if (!fs.existsSync(GENERATED_DIR)) {
      fs.mkdirSync(GENERATED_DIR, { recursive: true });
    }

    const timestamp  = Date.now();
    const filename   = `detention_report_bmed3_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWordBMED3(
      { examName, schoolName, programme, semester, date },
      {
        tableI:   mergedTableI,
        tableII:  mergedTableII,
        tableIII: mergedTableIII,
      },
      outputPath
    );

    res.json({ success: true, filename, stats });

  } catch (err) {
    console.error('[generateReportBMED3]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

module.exports = { generateReportBMED3 };
