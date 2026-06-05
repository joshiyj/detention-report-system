/**
 * excelParser2.js
 * Parses 2nd Year (4th Semester) RBU attendance Excel sheets.
 *
 * Column structure:
 *   Col 0: Roll No       (e.g. "O1_1")
 *   Col 1: Student Name
 *   Col 2: Overall Attendance  (e.g. "198/248 (79.84%)")
 *   Col 3..N: Subject columns  (ESD, ESD Lab, OS, OS Lab, DAA, AIML, AIML Lab, ...)
 *
 * Differences from 1st year parser:
 *   - No "Unique Id" column (col 1 = student name directly)
 *   - No "Seat No" column (rollNo is used as the unique key)
 *   - Overall attendance format: "198/248 (79.84%)" — note the "%" inside parens
 *   - Subject values same format: "32 / 34 (94.12%)" or pouring Attendance lines
 */

const XLSX = require('xlsx');

// Regex: extract the numeric percentage from "(79.84%)" or "(79.84)"
const PCT_REGEX = /\(\s*([\d.]+)\s*%?\s*\)/;

function extractPct(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  if (!s || s === 'NaN') return null;
  const m = s.match(PCT_REGEX);
  if (m) return parseFloat(m[1]);
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function extractOverallPct(val) {
  if (!val) return null;
  const s = String(val).trim();
  // Handles: "198/248 (79.84%)" or "198/248 (79.84)"
  const m = s.match(/\(\s*([\d.]+)\s*%?\s*\)/);
  if (m) return parseFloat(m[1]);
  return null;
}

/**
 * Parse a 2nd year Excel file buffer.
 * @param {Buffer} buffer
 * @returns {{ subjectNames: string[], students: StudentRecord[] }}
 */
function parseExcel2(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

  if (raw.length < 2) throw new Error('Excel file appears empty or has no data rows.');

  // Row 0 = header
  const headerRow = raw[0].map(h => (h ? String(h).trim() : ''));

  // Col 0 = Roll No, Col 1 = Student Name, Col 2 = Overall Attendance
  // Col 3+ = subject columns
  const subjectCols = headerRow.slice(3).map((name, i) => ({
    name: name || `Subject_${i}`,
    colIdx: i + 3,
  }));

  const students = [];
  let i = 1; // start after header

  while (i < raw.length) {
    const row = raw[i];

    // Skip blank rows
    if (!row || row.every(c => !c || String(c).trim() === '')) {
      i++;
      continue;
    }

    const rollNo = row[0] ? String(row[0]).trim() : null;
    if (!rollNo) { i++; continue; }

    const name        = row[1] ? String(row[1]).trim() : '';
    const seatNo      = '';   // 2nd year sheets have no Seat No column; use empty
    const div         = rollNo.split('_')[0] || ''; // e.g. "O1", "O2"
    const overallPct  = extractOverallPct(row[2]);

    // Collect all rows belonging to this student
    // (multiple rows possible due to "pouring Attendance" merged cells)
    const studentRows = [row];
    let j = i + 1;
    while (j < raw.length) {
      const nr = raw[j];
      if (!nr || nr.every(c => !c || String(c).trim() === '')) { j++; continue; }
      if (nr[0] && String(nr[0]).trim()) break; // new student starts
      studentRows.push(nr);
      j++;
    }

    // Detect "pouring Attendance" across all rows for this student
    let hasPouringAttendance = false;
    for (const sr of studentRows) {
      for (const sc of subjectCols) {
        const cellVal = sr[sc.colIdx];
        if (cellVal && String(cellVal).toLowerCase().includes('pouring attendance')) {
          hasPouringAttendance = true;
          break;
        }
      }
      if (hasPouringAttendance) break;
    }

    // Merge subject attendance across all rows (first non-null wins per column)
    const subjectAttendance = {};
    for (const sc of subjectCols) {
      let bestPct = null;
      for (const sr of studentRows) {
        const pct = extractPct(sr[sc.colIdx]);
        if (pct !== null) {
          bestPct = pct;
          break; // primary row preferred
        }
      }
      // Fallback: try subsequent rows (pouring attendance rows)
      if (bestPct === null && studentRows.length > 1) {
        for (const sr of studentRows.slice(1)) {
          const pct = extractPct(sr[sc.colIdx]);
          if (pct !== null) { bestPct = pct; break; }
        }
      }
      subjectAttendance[sc.name] = bestPct;
    }

    students.push({
      rollNo,
      seatNo,
      name,
      div,
      overallPct,
      hasPouringAttendance,
      subjects: subjectAttendance,
    });

    i = j;
  }

  return { subjectNames: subjectCols.map(s => s.name), students };
}

module.exports = { parseExcel2 };
