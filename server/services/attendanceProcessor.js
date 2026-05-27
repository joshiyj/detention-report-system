/**
 * attendanceProcessor.js
 * Applies attendance rules and builds datasets for all four tables.
 * Includes Theory/Practical type splitting and course code/name mapping.
 */

const OVERALL_THRESHOLD = 75;   // below this → detained overall
const CONDONATION_MIN   = 60;   // condonation window lower bound
const SUBJECT_THRESHOLD = 60;   // below this per subject → detained in that subject

// ── Course code + full name mapping (from SHORTFORM.docx) ─────────────────────
// Key = Excel column header abbreviation (exactly as in row 0 of the sheet)
// Maps each abbreviation to: { code, fullName, type }
const COURSE_MAP = {
  'MI (L)':   { code: '25EE01TP0201', fullName: 'Microcontroller and Interfacing',             type: 'Theory'    },
  'MI (P)':   { code: '25EE01TP0201', fullName: 'Microcontroller and Interfacing',             type: 'Practical' },
  'CLA':      { code: '25HS03TH0213', fullName: 'Calculus & Linear Algebra',                   type: 'Theory'    },
  'PPS (L)':  { code: '25EE01TP0202', fullName: 'Programming for Problem Solving',             type: 'Theory'    },
  'PPS (P)':  { code: '25EE01TP0202', fullName: 'Programming for Problem Solving',             type: 'Practical' },
  'AIML':     { code: '25EE01TH0203', fullName: 'AI and ML Essentials',                        type: 'Theory'    },
  'CAO (L)':  { code: '25EE01TP0204', fullName: 'Computer Architecture and Organization',      type: 'Theory'    },
  'CAO (P)':  { code: '25EE01TP0204', fullName: 'Computer Architecture and Organization',      type: 'Practical' },
  'EPC (L)':  { code: '25HS02TP0201', fullName: 'English for Professional Communication',      type: 'Theory'    },
  'EPC (P)':  { code: '25HS02TP0201', fullName: 'English for Professional Communication',      type: 'Practical' },
  'FLIC':     { code: '25HS02TH0203-1', fullName: 'Foundational Literature of Indian Civilization', type: 'Theory' },
  'MP1':      { code: '25EE01PR0205', fullName: 'Mini Project-I',                              type: 'Practical' },
  'HFW':      { code: '25HS04PR0201', fullName: 'Health-Fitness-Wellbeing',                    type: 'Practical' },
  'INC':      { code: 'INC',          fullName: 'INC',                                         type: 'Theory'    },
};

/**
 * Resolve a subject abbreviation to its course info.
 * Falls back gracefully if not found in the map.
 */
function resolveCourse(abbr) {
  const mapped = COURSE_MAP[abbr.trim()];
  if (mapped) return mapped;
  // Fallback: derive type from suffix
  const type = abbr.includes('(P)') ? 'Practical' : 'Theory';
  return { code: abbr, fullName: abbr, type };
}

/**
 * Process students and return all table datasets.
 *
 * @param {Array}  students        - from excelParser
 * @param {Array}  subjectNames    - ordered list of subject column names
 * @param {Set}    condonationSeats - seat numbers that applied for condonation
 * @returns {{ tableI, tableII, tableIIIA, tableIIIB, stats }}
 */
function processAttendance(students, subjectNames, condonationSeats = new Set()) {

  // Table I: overall < 75%
  const tableI = [];

  // Table II: 60 <= overall <= 75 AND has pouring attendance in at least one subject
  const tableII = [];

  // Table III(A): course-wise – key = "CODE::type", value = [{ div, seatNo, name, pct, type, code, fullName }]
  // Using "courseCode::type" as key to separate Theory and Practical streams
  const tableIIIA_map = {};

  // Table III(B): student-wise
  const tableIIIB = [];

  for (const s of students) {
    const overall = s.overallPct;

    // ── Table I ──────────────────────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      tableI.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall !== null ? overall.toFixed(2) : 'N/A',
      });
    }

    // ── Table II ─────────────────────────────────────────────────────────────
    if (
      overall !== null &&
      overall >= CONDONATION_MIN &&
      overall <= OVERALL_THRESHOLD &&
      s.hasPouringAttendance
    ) {
      tableII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall.toFixed(2),
      });
    }

    // ── Table III – subjects below 60% ────────────────────────────────────────
    // Each subject column is kept separate (Theory vs Practical are different columns)
    const detainedEntries = []; // { code, fullName, type, pct }

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct !== null && pct !== undefined && pct < SUBJECT_THRESHOLD) {
        const { code, fullName, type } = resolveCourse(subj);
        const entry = { abbr: subj, code, fullName, type, pct };
        detainedEntries.push(entry);

        // III(A) course-wise map — key groups by course code + type
        const mapKey = `${code}::${type}`;
        if (!tableIIIA_map[mapKey]) {
          tableIIIA_map[mapKey] = { code, fullName, type, rows: [] };
        }
        tableIIIA_map[mapKey].rows.push({
          div:    s.div,
          rollNo: s.rollNo,
          seatNo: s.seatNo,
          name:   s.name,
          pct:    pct.toFixed(2),
        });
      }
    }

    if (detainedEntries.length > 0) {
      tableIIIB.push({
        div:     s.div,
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall !== null ? overall.toFixed(2) : 'N/A',
        // Each entry has: { code, fullName, type, pct }
        subjects: detainedEntries.map(d => ({
          code:     d.code,
          fullName: d.fullName,
          type:     d.type,
          pct:      d.pct.toFixed(2),
        })),
      });
    }
  }

  // Convert III(A) map to ordered array, sorted by course code then type
  const tableIIIA = Object.values(tableIIIA_map)
    .sort((a, b) => {
      if (a.code < b.code) return -1;
      if (a.code > b.code) return  1;
      // Theory before Practical within same code
      if (a.type === 'Theory' && b.type === 'Practical') return -1;
      if (a.type === 'Practical' && b.type === 'Theory') return  1;
      return 0;
    });

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIIIB.length,
    detainedSubjects: Object.values(tableIIIA_map).reduce((a, v) => a + v.rows.length, 0),
  };

  return { tableI, tableII, tableIIIA, tableIIIB, stats };
}

module.exports = { processAttendance, COURSE_MAP };
