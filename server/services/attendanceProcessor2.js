/**
 * attendanceProcessor2.js
 * Processes 2nd Year (4th Semester) attendance sheets.
 * Same detention rules as 1st year:
 *   - Overall < 75%  → detained overall (Table I)
 *   - 60% ≤ Overall < 75% with pouring attendance → condonation (Table II)
 *   - Per-subject (or avg of Theory+Practical) < 60% → subject detention (Table III)
 *
 * Sheet structure (Semester_IV):
 *   Col 0: Roll No  (e.g. "O1_1")
 *   Col 1: Student Name
 *   Col 2: Overall Attendance  (e.g. "198/248 (79.84%)")
 *   Col 3+: Subject columns
 */

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── 4th Semester Course Map ────────────────────────────────────────────────────
// Key = exact column header from Excel sheet
// Value = { code, fullName, type }
const COURSE_MAP_Y2 = {
  'ESD':                        { code: '24EE01TH0401',    fullName: 'Embedded System Design',                          type: 'Theory'    },
  'ESD Lab':                    { code: '24EE01PR0401',    fullName: 'Embedded System Design Lab',                      type: 'Practical' },
  'OS':                         { code: '24EE01TH0402',    fullName: 'Operating System',                                type: 'Theory'    },
  'OS Lab':                     { code: '24EE01PR0402',    fullName: 'Operating System Lab',                            type: 'Practical' },
  'DAA':                        { code: '24EE01TH0403',    fullName: 'Design and Analysis of Algorithms',               type: 'Theory'    },
  'AIML':                       { code: '24EE01TH0404',    fullName: 'Fundamentals of AI and Machine Learning',         type: 'Theory'    },
  'AIML Lab':                   { code: '24EE01PR0404',    fullName: 'Fundamentals of AI and Machine Learning Lab',     type: 'Practical' },
  'MDM-2':                      { code: '24EE01TH0405',    fullName: 'MDM-2',                                           type: 'Theory'    },
  'OE (IOT/MOOCs)':             { code: '24EEOEC01TH0406', fullName: 'Open Elective-II/MOOCs',                          type: 'Theory'    },
  'Software Lab Practice-I':    { code: '24EE01PR0407',    fullName: 'Software Laboratory Practice-I',                  type: 'Practical' },
  'Basic Competitive Coding':   { code: '24EE01PR0408',    fullName: 'Basic Competitive Coding',                        type: 'Practical' },
  'Innovations & Entrepreneurship': { code: '24SM01TH0401', fullName: 'Innovations and Entrepreneurship',              type: 'Theory'    },
  'I&E Practical':              { code: '24SM01PR0401',    fullName: 'Innovations and Entrepreneurship Lab',            type: 'Practical' },
};

// ── Canonical 4th Sem course order (per PDF SN order) ─────────────────────────
// Each entry: "CODE::Type"
const SHORTFORM_ORDER_Y2 = [
  '24EE01TH0401::Theory',      // SN 1  – ESD (Theory)
  '24EE01PR0401::Practical',   // SN 2  – ESD Lab
  '24EE01TH0402::Theory',      // SN 3  – OS (Theory)
  '24EE01PR0402::Practical',   // SN 4  – OS Lab
  '24EE01TH0403::Theory',      // SN 5  – DAA
  '24EE01TH0404::Theory',      // SN 6  – AIML (Theory)
  '24EE01PR0404::Practical',   // SN 7  – AIML Lab
  '24EE01TH0405::Theory',      // SN 8  – MDM-2
  '24EEOEC01TH0406::Theory',   // SN 9  – OE (IOT/MOOCs)
  '24EE01PR0407::Practical',   // SN 10 – Software Lab Practice-I
  '24EE01PR0408::Practical',   // SN 11 – Basic Competitive Coding
  '24SM01TH0401::Theory',      // SN 12 – I&E (Theory)
  '24SM01PR0401::Practical',   // SN 13 – I&E Practical
];

function shortformIndexY2(code, type) {
  const key = `${code}::${type}`;
  const idx = SHORTFORM_ORDER_Y2.indexOf(key);
  return idx === -1 ? SHORTFORM_ORDER_Y2.length : idx;
}

function shortformCodeIndexY2(code) {
  for (let i = 0; i < SHORTFORM_ORDER_Y2.length; i++) {
    if (SHORTFORM_ORDER_Y2[i].startsWith(`${code}::`)) return i;
  }
  return SHORTFORM_ORDER_Y2.length;
}

/**
 * Resolve subject column header to course info.
 */
function resolveCourseY2(abbr) {
  const mapped = COURSE_MAP_Y2[abbr.trim()];
  if (mapped) return mapped;
  const type = abbr.toLowerCase().includes('lab') || abbr.toLowerCase().includes('practical')
    ? 'Practical' : 'Theory';
  return { code: abbr, fullName: abbr, type };
}

/**
 * processAttendanceY2
 * @param {Array} students    – from excelParser2
 * @param {Array} subjectNames – ordered list of subject column headers
 * @param {Set}   condonationSeats – seat/roll numbers applied for condonation
 */
function processAttendanceY2(students, subjectNames, condonationSeats = new Set()) {
  const tableI    = [];
  const tableII   = [];
  const tableIIIA_map = {};  // code → { code, fullName, studentMap }
  const tableIIIB = [];

  for (const s of students) {
    const overall = s.overallPct;

    // ── Table I: overall < 75% ──────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      tableI.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,   // may be empty for 2nd year sheets
        name:    s.name,
        overall: overall.toFixed(2),
      });
    }

    // ── Table II: 60% ≤ overall < 75% AND has pouring attendance ────────────
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

    // ── Table III – per-subject detention with avg rule ──────────────────────
    // Collect ALL subject percentages (including those >= 60%) to apply avg rule
    const allCourseData = {}; // code → { code, fullName, theoryPct, practicalPct }
    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct === null || pct === undefined) continue;
      const { code, fullName, type } = resolveCourseY2(subj);
      if (!allCourseData[code]) {
        allCourseData[code] = { code, fullName, theoryPct: null, practicalPct: null };
      }
      if (type === 'Theory') allCourseData[code].theoryPct = pct;
      else                   allCourseData[code].practicalPct = pct;
    }

    // Determine detained courses applying avg rule
    const detainedCourses = {};
    for (const cd of Object.values(allCourseData)) {
      const hasBoth = cd.theoryPct !== null && cd.practicalPct !== null;
      if (hasBoth) {
        const avg = (cd.theoryPct + cd.practicalPct) / 2;
        if (avg < SUBJECT_THRESHOLD) {
          detainedCourses[cd.code] = { ...cd };
        }
      } else {
        const singlePct = cd.theoryPct ?? cd.practicalPct;
        if (singlePct < SUBJECT_THRESHOLD) {
          detainedCourses[cd.code] = { ...cd };
        }
      }
    }

    // Feed detained courses → tableIIIA_map
    for (const dc of Object.values(detainedCourses)) {
      if (!tableIIIA_map[dc.code]) {
        tableIIIA_map[dc.code] = { code: dc.code, fullName: dc.fullName, studentMap: {} };
      }
      const studentKey = s.seatNo || s.rollNo;
      if (!tableIIIA_map[dc.code].studentMap[studentKey]) {
        tableIIIA_map[dc.code].studentMap[studentKey] = {
          div:    s.div,
          rollNo: s.rollNo,
          seatNo: s.seatNo,
          name:   s.name,
          theoryPct:    dc.theoryPct    !== null ? dc.theoryPct.toFixed(2)    : null,
          practicalPct: dc.practicalPct !== null ? dc.practicalPct.toFixed(2) : null,
        };
      }
    }

    // Build tableIIIB from detained courses
    if (Object.keys(detainedCourses).length > 0) {
      // Sort subjects per SHORTFORM order
      const subjects = Object.values(detainedCourses)
        .sort((a, b) => shortformCodeIndexY2(a.code) - shortformCodeIndexY2(b.code))
        .map(dc => {
          // For Table III(B), show individual component attendances as separate rows
          // (one row per component when both exist and avg is detained)
          const rows = [];
          if (dc.theoryPct !== null) {
            rows.push({
              code:     dc.code,
              fullName: dc.fullName,
              type:     'Theory',
              pct:      dc.theoryPct.toFixed(2),
            });
          }
          if (dc.practicalPct !== null) {
            rows.push({
              code:     dc.code,
              fullName: dc.fullName,
              type:     'Practical',
              pct:      dc.practicalPct.toFixed(2),
            });
          }
          if (rows.length === 0) {
            rows.push({ code: dc.code, fullName: dc.fullName, type: 'Theory', pct: '0.00' });
          }
          return rows;
        })
        .flat();

      tableIIIB.push({
        div:     s.div,
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall !== null ? overall.toFixed(2) : 'N/A',
        subjects,
      });
    }
  }

  // Convert tableIIIA_map → ordered array
  const tableIIIA = Object.values(tableIIIA_map)
    .map(c => ({
      code:     c.code,
      fullName: c.fullName,
      rows: Object.values(c.studentMap)
        .filter(s => {
          // Re-apply avg filter in case student was inserted via one-component path
          if (s.theoryPct !== null && s.practicalPct !== null) {
            const avg = (parseFloat(s.theoryPct) + parseFloat(s.practicalPct)) / 2;
            return avg < SUBJECT_THRESHOLD;
          }
          return true;
        })
        .map(s => ({
          rollNo: s.rollNo,
          seatNo: s.seatNo,
          name:   s.name,
          div:    s.div,
          // Attendance: show individual values for Theory and Practical columns
          theoryPct:    s.theoryPct,
          practicalPct: s.practicalPct,
          // Combined pct string for display (used by word generator if needed)
          pct: s.theoryPct !== null && s.practicalPct !== null
               ? `${s.theoryPct} % / ${s.practicalPct} %`
               : `${s.theoryPct ?? s.practicalPct} %`,
        })),
    }))
    .filter(c => c.rows.length > 0)
    .sort((a, b) => shortformCodeIndexY2(a.code) - shortformCodeIndexY2(b.code));

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIIIB.length,
    detainedSubjects: tableIIIA.reduce((a, c) => a + c.rows.length, 0),
  };

  return { tableI, tableII, tableIIIA, tableIIIB, stats };
}

module.exports = { processAttendanceY2, COURSE_MAP_Y2, SHORTFORM_ORDER_Y2, shortformCodeIndexY2 };
