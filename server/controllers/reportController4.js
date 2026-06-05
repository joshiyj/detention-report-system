/**
* reportController4.js
* Handles upload, generate, and download endpoints for 4th Year (8th Semester) ECS.
* Supports multiple Excel files (one per section/division).
*
* Sorting rules:
*   - Preserves original raw merge order (no alphabetical sorting).
*/

const path = require('path');
const fs   = require('fs');
const { parseExcel }             = require('../services/excelParser');
const { processAttendanceY4 }    = require('../services/attendanceProcessorY4');
const { generateWordY4 }         = require('../services/wordGeneratorY4');
const { parseDocxMapping }        = require('../services/docxMappingParser');
const { getMapping, saveMapping }   = require('../services/mappingPersistence');

const GENERATED_DIR = path.join(__dirname, '..', 'generated');

/**
* POST /api/y4/generate
*/
async function generateReport4(req, res) {
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
        saveMapping('y4', dynamicMapping);
      }
    } else {
      dynamicMapping = getMapping('y4');
    }

    const condonationIds = new Set(
      (condonation || '').split(',').map(s => s.trim()).filter(Boolean)
    );

    let mergedTableI   = [];
    let mergedTableII  = [];
    let mergedTableIII = [];
    let totalStudents  = 0;

    for (const file of files) {
      const { subjectNames, students } = parseExcel(file.buffer);

      if (!students.length) {
        console.warn(`[generateReport4] No students in: ${file.originalname}`);
        continue;
      }

      const { tableI, tableII, tableIII } =
        processAttendanceY4(students, subjectNames, condonationIds, dynamicMapping);

      mergedTableI   = mergedTableI.concat(tableI);
      mergedTableII  = mergedTableII.concat(tableII);
      mergedTableIII = mergedTableIII.concat(tableIII);
      totalStudents += students.length;
    }

    if (totalStudents === 0) {
      return res.status(400).json({ error: 'No student records found in any of the uploaded files.' });
    }

    // ── Deduplicate tables by unique Seat No / Roll No ───────────────────────
    const seenI = new Set();
    const uniqueI = [];
    for (const item of mergedTableI) {
      const id = item.seatNo || item.rollNo;
      if (id && !seenI.has(id)) {
        seenI.add(id);
        uniqueI.push(item);
      }
    }
    mergedTableI = uniqueI;

    const seenII = new Set();
    const uniqueII = [];
    for (const item of mergedTableII) {
      const id = item.seatNo || item.rollNo;
      if (id && !seenII.has(id)) {
        seenII.add(id);
        uniqueII.push(item);
      }
    }
    mergedTableII = uniqueII;

    const seenIII = new Set();
    const uniqueIII = [];
    for (const item of mergedTableIII) {
      const id = item.seatNo || item.rollNo;
      if (id && !seenIII.has(id)) {
        seenIII.add(id);
        uniqueIII.push(item);
      }
    }
    mergedTableIII = uniqueIII;

    // Note: No sorting by Seat No for Year 4 to match the reference document's original sheet order.

    const stats = {
      totalStudents,
      tableICount:      mergedTableI.length,
      tableIICount:     mergedTableII.length,
      detainedStudents: mergedTableIII.length,
      detainedCourses:  mergedTableIII.reduce((acc, s) => acc + s.courses.length, 0),
    };

    const timestamp  = Date.now();
    const filename   = `detention_report_y4_${timestamp}.docx`;
    const outputPath = path.join(GENERATED_DIR, filename);

    await generateWordY4(
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
    console.error('[generateReport4]', err);
    res.status(500).json({ error: err.message || 'Internal server error.' });
  }
}

module.exports = { generateReport4 };
