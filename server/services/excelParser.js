/**
 * excelParser.js
 * Parses the RBU attendance Excel sheet.
 *
 * Excel structure quirks:
 *  - Each student may span 2-3 rows (due to "pouring attendance" merged cells)
 *  - Empty rows separate students
 *  - Attendance values: "10 / 24 ( 41.67 % )" OR "pouring Attendance :28.0/ 33 (84.85)"
 *  - Overall attendance: "134/256 (52.34)"
 *  - Some cells may be NaN / undefined
 */

const XLSX = require('xlsx');

// Regex to extract percentage from various formats
const PCT_REGEX = /\(\s*([\d.]+)\s*%?\s*\)/;

function extractPct(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  if (!s || s === 'NaN') return null;
  const m = s.match(PCT_REGEX);
  if (m) return parseFloat(m[1]);
  // plain number fallback
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function extractOverallPct(val) {
  if (!val) return null;
  const s = String(val).trim();
  // Format: "134/256 (52.34)"
  const m = s.match(/\((\s*[\d.]+\s*)\)/);
  if (m) return parseFloat(m[1]);
  return null;
}

/**
 * Parse the Excel file buffer and return normalized student records.
 * @param {Buffer} buffer
 * @returns {{ headers: string[], students: StudentRecord[] }}
 */
function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  // Read as array of arrays (raw, no header inference)
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

  if (raw.length < 2) throw new Error('Excel file appears empty or has no data rows.');

  // Row 0 is header
  const headerRow = raw[0].map(h => (h ? String(h).trim() : ''));

  // Subject columns: everything after col 4 (index 4 = OverAll Attendance)
  // Cols: 0=Roll No, 1=Unique Id, 2=Seat No, 3=Student name, 4=OverAll, 5..N = subjects
  const subjectCols = headerRow.slice(5).map((name, i) => ({ name: name || `Subject_${i}`, colIdx: i + 5 }));

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

    const seatNo = row[2] ? String(row[2]).trim() : '';
    const name   = row[3] ? String(row[3]).trim() : '';
    const div    = rollNo.split('_')[0] || ''; // e.g. "O3"
    const overallPct = extractOverallPct(row[4]);

    // Build subject attendance map – may span multiple rows
    // Collect rows belonging to this student (until next non-empty Roll No or end)
    const studentRows = [row];
    let j = i + 1;
    while (j < raw.length) {
      const nr = raw[j];
      if (!nr || nr.every(c => !c || String(c).trim() === '')) { j++; continue; }
      // If next row has a Roll No value, it's a new student
      if (nr[0] && String(nr[0]).trim()) break;
      studentRows.push(nr);
      j++;
    }

    // Detect "pouring Attendance" in any subject cell across all rows for this student
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

    // Merge subject attendance across all rows for this student
    // Use highest-quality value per column (prefer non-null, non-pouring first row; then pouring rows)
    const subjectAttendance = {};
    for (const sc of subjectCols) {
      let bestPct = null;
      for (const sr of studentRows) {
        const pct = extractPct(sr[sc.colIdx]);
        if (pct !== null) {
          bestPct = pct;
          break; // first non-null wins (primary row preferred)
        }
      }
      // If still null, try 2nd row entries (merged pouring cells)
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

    i = j; // advance past this student's rows
  }

  return { subjectNames: subjectCols.map(s => s.name), students };
}

module.exports = { parseExcel };
