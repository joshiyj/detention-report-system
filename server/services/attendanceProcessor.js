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

// ── Canonical course order as per SHORTFORM.docx (SN 1-9) ────────────────────
// Each entry is "CODE::Type" matching the keys used in tableIIIA_map.
// Theory is listed before Practical for the same course (matching the document).
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

/**
 * Returns the SHORTFORM sort index for a course entry.
 * Unknown courses are placed at the end.
 */
function shortformIndex(code, type, activeShortformOrder) {
  const order = activeShortformOrder || SHORTFORM_ORDER;
  const key = `${code}::${type}`;
  const idx = order.indexOf(key);
  return idx === -1 ? order.length : idx;
}

/**
 * Returns the SHORTFORM sort index for a course by code alone
 * (uses the first appearance of the code in SHORTFORM_ORDER).
 */
function shortformCodeIndex(code, activeShortformOrder) {
  const order = activeShortformOrder || SHORTFORM_ORDER;
  for (let i = 0; i < order.length; i++) {
    if (order[i].startsWith(`${code}::`)) return i;
  }
  return order.length;
}

/**
 * Resolve a subject abbreviation to its course info.
 * Falls back gracefully if not found in the map.
 */
function resolveCourse(abbr, activeCourseMap) {
  const map = activeCourseMap || COURSE_MAP;
  const mapped = map[abbr.trim()];
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
 * @param {Object} dynamicMapping  - parsed docx course mapping (optional)
 * @returns {{ tableI, tableII, tableIIIA, tableIIIB, stats }}
 */
function processAttendance(students, subjectNames, condonationSeats = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP;
  const activeShortformOrder = dynamicMapping ? dynamicMapping.shortformOrder : SHORTFORM_ORDER;

  // Table I: overall < 75%
  const tableI = [];

  // Table II: 60 <= overall < 75 AND has pouring attendance in at least one subject
  const tableII = [];

  // Table III(A): course-wise – key = course code only.
  // Per-course, a studentMap (keyed by seatNo) accumulates Theory + Practical pcts
  // so each student gets ONE row showing both side-by-side.
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
      overall < OVERALL_THRESHOLD &&
      s.hasPouringAttendance
    ) {
      tableII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall.toFixed(2),
      });
    }

    // ── Table III — exempt students with overall >= 75% ──────────────────────
    // A student whose overall attendance is 75% or more must not appear in
    // Table III(A) or Table III(B), even if an individual course is below 60%.
    if (overall !== null && overall >= OVERALL_THRESHOLD) {
      continue;
    }

    // ── Table III – course detention (component < 60%) ────────────────────────
    const allCourseData = {}; // code → { code, fullName, theoryPct, practicalPct }
    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct === null || pct === undefined) continue;
      const { code, fullName, type } = resolveCourse(subj, activeCourseMap);
      if (!allCourseData[code]) {
        allCourseData[code] = { code, fullName, theoryPct: null, practicalPct: null };
      }
      if (type === 'Theory') allCourseData[code].theoryPct = pct;
      else                   allCourseData[code].practicalPct = pct;
    }

    // Determine which courses are detained: any component (Theory or Practical) < 60%
    const detainedCourses = {};

    for (const cd of Object.values(allCourseData)) {
      const theoryDetained = cd.theoryPct !== null && cd.theoryPct < SUBJECT_THRESHOLD;
      const practicalDetained = cd.practicalPct !== null && cd.practicalPct < SUBJECT_THRESHOLD;

      if (theoryDetained || practicalDetained) {
        detainedCourses[cd.code] = {
          code: cd.code,
          fullName: cd.fullName,
          theoryPct: theoryDetained ? cd.theoryPct : null,
          practicalPct: practicalDetained ? cd.practicalPct : null,
        };
      }
    }

    // Feed detained courses into tableIIIA_map
    for (const dc of Object.values(detainedCourses)) {
      const mapKey = dc.code;
      if (!tableIIIA_map[mapKey]) {
        tableIIIA_map[mapKey] = { code: dc.code, fullName: dc.fullName, studentMap: {} };
      }
      const studentKey = s.seatNo || s.rollNo;
      if (!tableIIIA_map[mapKey].studentMap[studentKey]) {
        tableIIIA_map[mapKey].studentMap[studentKey] = {
          div: s.div, rollNo: s.rollNo, seatNo: s.seatNo, name: s.name,
          theoryPct: dc.theoryPct !== null ? dc.theoryPct.toFixed(2) : null,
          practicalPct: dc.practicalPct !== null ? dc.practicalPct.toFixed(2) : null,
        };
      }
    }

    // Build tableIIIB from detained courses
    if (Object.keys(detainedCourses).length > 0) {
      const subjects = Object.values(detainedCourses)
        .map(dc => ({
          code:     dc.code,
          fullName: dc.fullName,
          type: dc.theoryPct !== null && dc.practicalPct !== null
                ? 'Theory/Practical'
                : dc.theoryPct !== null ? 'Theory' : 'Practical',
          // pct includes the % sign already (e.g. "42.86 %" or "42.86 %/9.09 %")
          pct: dc.theoryPct !== null && dc.practicalPct !== null
               ? `${dc.theoryPct.toFixed(2)} %/${dc.practicalPct.toFixed(2)} %`
               : `${(dc.theoryPct ?? dc.practicalPct).toFixed(2)} %`,
        }))
        .sort((a, b) => shortformCodeIndex(a.code, activeShortformOrder) - shortformCodeIndex(b.code, activeShortformOrder));

      tableIIIB.push({
        div:      s.div,
        rollNo:   s.rollNo,
        seatNo:   s.seatNo,
        name:     s.name,
        overall:  overall !== null ? overall.toFixed(2) : 'N/A',
        subjects,
      });
    }
  }

  // Convert III(A) map → ordered array; Theory+Practical merged into one row per student
  const tableIIIA = Object.values(tableIIIA_map)
    .map(c => ({
      code:     c.code,
      fullName: c.fullName,
      rows: Object.values(c.studentMap)
        .map(s => ({
          rollNo: s.rollNo,
          seatNo: s.seatNo,
          name:   s.name,
          div:    s.div,
          // Type: show both if detained in both, otherwise whichever applies
          type: s.theoryPct !== null && s.practicalPct !== null
                ? 'Theory/Practical'
                : s.theoryPct !== null ? 'Theory' : 'Practical',
          // Attendance: "T%/P%" if both, single value otherwise
          pct: s.theoryPct !== null && s.practicalPct !== null
               ? `${s.theoryPct} %/${s.practicalPct} %`
               : `${s.theoryPct ?? s.practicalPct} %`,
        })),
    }))
    // Drop courses that have no student rows remaining
    .filter(c => c.rows.length > 0)
    .sort((a, b) => shortformCodeIndex(a.code, activeShortformOrder) - shortformCodeIndex(b.code, activeShortformOrder));

  // III-B subjects are already sorted by shortformCodeIndex inside the map above

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIIIB.length,
    detainedSubjects: tableIIIA.reduce((a, c) => a + c.rows.length, 0),
  };

  return { tableI, tableII, tableIIIA, tableIIIB, stats };
}

module.exports = { processAttendance, COURSE_MAP };
