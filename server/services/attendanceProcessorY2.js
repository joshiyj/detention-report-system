/**
 * attendanceProcessorY2.js
 * Processes 2nd Year (4th Semester) attendance sheets based on Master_Course_Mapping_Table.docx.
 *
 * Excel structure (A updated 23.xls, B updated 23.xls, C updated 23.xls, updated dd.xls):
 *   Col 0: Roll No.     (e.g. "B-10")
 *   Col 1: Unique Id    (email)
 *   Col 2: Seat No      (e.g. "UECS26RS40022")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance  (e.g. "165/232 (71.12)")
 *   Col 5+: Subject columns (ESD, OS, DAA, AIML, MC, CMT, EVAC, BFS, IPD, WS, INC, FED, IE, OE II, ESD LAB, OS LAB, AIML LAB, SL I, BCC, IE Lab)
 *
 * Rules:
 *   - Null / blank subject cell → student NOT enrolled in that course (skip entirely)
 *   - Overall < 75% → Table I
 *   - 60% ≤ Overall < 75% AND has pouring attendance → Table II
 *   - Subject/avg < 60% → subject detention in Table III
 *   - If Theory+Practical both exist for a course: apply avg rule
 *   - Sort: Table I & II by Seat No; Table III(A) & III(B) by SN from Master Course Mapping
 */

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── Master Course Mapping (from Master_Course_Mapping_Table.docx) ─────────────
// SN order is the canonical sort order for Table III(A) and III(B).
// Key = EXACT column header abbreviation in the Excel file.
const COURSE_MAP_Y2 = {
  'ESD':      { sn: 1,  code: '24EE01TH0401',       fullName: 'Embedded System Design',                        type: 'Theory'    },
  'ESD LAB':  { sn: 2,  code: '24EE01PR0401',       fullName: 'Embedded System Design Lab',                    type: 'Practical' },
  'OS':       { sn: 3,  code: '24EE01TH0402',       fullName: 'Operating System',                              type: 'Theory'    },
  'OS LAB':   { sn: 4,  code: '24EE01PR0402',       fullName: 'Operating System Lab',                          type: 'Practical' },
  'DAA':      { sn: 5,  code: '24EE01TH0403',       fullName: 'Design and Analysis of Algorithms',             type: 'Theory'    },
  'AIML':     { sn: 6,  code: '24EE01TH0404',       fullName: 'Fundamentals of AI and Machine Learning',       type: 'Theory'    },
  'AIML LAB': { sn: 7,  code: '24EE01PR0404',       fullName: 'Fundamentals of AI and Machine Learning Lab',   type: 'Practical' },
  'MC':       { sn: 8,  code: '24EE05TH0408-2',     fullName: 'Mobile Communication',                          type: 'Theory'    },
  'CMT':      { sn: 9,  code: '24ES01TH0404',       fullName: 'Construction Materials and Technology',         type: 'Theory'    },
  'EVAC':     { sn: 10, code: '24EE07TH0409',       fullName: 'EV Architecture and Components',                type: 'Theory'    },
  'BFS':      { sn: 11, code: '24SM02TH0401',       fullName: 'Banking and Financial Services',                type: 'Theory'    },
  'IPD':      { sn: 12, code: '24SM03TH0401',       fullName: 'Innovation and Product Development',            type: 'Theory'    },
  'WS':       { sn: 13, code: '24ES03TH0407',       fullName: 'Warfare System',                                type: 'Theory'    },
  'INC':      { sn: 999, code: 'INC',                fullName: 'INC',                                           type: 'Theory'    },
  'FED':      { sn: 14,  code: '24CS01TH0407-1',     fullName: 'Front End Development',                         type: 'Theory'    },

  'OE II':    { sn: 15, code: '24EEOEC01TH0406',    fullName: 'Open Elective-II / MOOCs',                      type: 'Theory'    },
  'SL I':     { sn: 16, code: '24EE01PR0407',       fullName: 'Software Laboratory Practice-I',                type: 'Practical' },
  'BCC':      { sn: 17, code: '24EE01PR0408',       fullName: 'Basic Competitive Coding',                      type: 'Practical' },
  'IE':       { sn: 18, code: '24SM01TH0401',       fullName: 'Innovations and Entrepreneurship',              type: 'Theory'    },
  'IE Lab':   { sn: 19, code: '24SM01PR0401',       fullName: 'Innovations and Entrepreneurship Lab',          type: 'Practical' },
};

// ── Theory-Practical pairs (same code) ────────────────────────────────────────
// For average-rule: if a course has both Theory+Practical components,
// use the average to decide detention.
// In 4th sem mapping, Theory+Practical pairs:
//   ESD (sn 1) + ESD LAB (sn 2) → same physical course (24EE01TH0401 / 24EE01PR0401)
//   OS (sn 3) + OS LAB (sn 4)   → 24EE01TH0402 / 24EE01PR0402
//   AIML (sn 6) + AIML LAB (sn 7) → 24EE01TH0404 / 24EE01PR0404
//   IE (sn 18) + IE Lab (sn 19) → 24SM01TH0401 / 24SM01PR0401
// All other courses are standalone Theory-only or Practical-only.

// Canonical SN sort order for Table III(A) and III(B):
// Each entry: "CODE::type" — but we group by SN, not by code, since FED/INC overlap.
// We use SN from the master map for sorting.
function getSN(abbr) {
  const entry = COURSE_MAP_Y2[abbr];
  return entry ? entry.sn : 999;
}

function getCodeSN(code) {
  const entries = Object.values(COURSE_MAP_Y2).filter(e => e.code === code);
  return entries.length > 0 ? entries[0].sn : 999;
}

/**
 * Resolve a subject abbreviation to course info.
 * Returns null if abbreviation is not in COURSE_MAP_Y2 (unknown subject).
 */
function resolveCourseY2new(abbr, activeCourseMap) {
  const map = activeCourseMap || COURSE_MAP_Y2;
  const mapped = map[abbr ? abbr.trim() : ''];
  if (mapped) return mapped;
  // Unknown column — generic fallback
  const type = abbr && (abbr.toLowerCase().includes('lab') || abbr.toLowerCase().includes('practical'))
    ? 'Practical' : 'Theory';
  return { sn: 999, code: abbr, fullName: abbr, type };
}

/**
 * processAttendanceY2new
 * @param {Array} students      — from excelParser (with seatNo present)
 * @param {Array} subjectNames  — ordered list of subject column headers
 * @param {Set}   condonationSeats — seat numbers applied for condonation
 * @param {Object} dynamicMapping — parsed docx course mapping (optional)
 */
function processAttendanceY2new(students, subjectNames, condonationSeats = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP_Y2;
  const tableI         = [];
  const tableII        = [];
  const tableIIIA_map  = {};  // code → { code, fullName, sn, studentMap }
  const tableIIIB      = [];

  for (const s of students) {
    const overall = s.overallPct;

    // ── Table I: overall < 75% ────────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      tableI.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall.toFixed(2),
      });
    }

    // ── Table II: 60% ≤ overall < 75% AND has pouring attendance ─────────────
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

    // ── Table III — per-subject detention ─────────────────────────────────────
    // If a student has overall attendance >= 75%, they are exempt from subject detention.
    if (overall !== null && overall >= OVERALL_THRESHOLD) {
      continue;
    }

    // Step 1: Collect all subject pct values, grouped by course code (to apply avg rule for T+P pairs)
    // Blank cells (null) = not enrolled → skip entirely
    const allCourseData = {}; // code → { code, fullName, sn, theoryPct, practicalPct, theoryAbbr, practicalAbbr }

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      // null = not enrolled in this course → skip
      if (pct === null || pct === undefined) continue;

      const { sn, code, fullName, type } = resolveCourseY2new(subj, activeCourseMap);
      if (!allCourseData[code]) {
        allCourseData[code] = { code, fullName, sn, theoryPct: null, practicalPct: null };
      }
      if (type === 'Theory')    allCourseData[code].theoryPct    = pct;
      else                      allCourseData[code].practicalPct = pct;
    }

    // Step 2: Apply detention threshold + avg rule
    const detainedCourses = {};

    for (const cd of Object.values(allCourseData)) {
      const hasBoth = cd.theoryPct !== null && cd.practicalPct !== null;

      if (hasBoth) {
        const avg = (cd.theoryPct + cd.practicalPct) / 2;
        if (avg < SUBJECT_THRESHOLD) {
          detainedCourses[cd.code] = { ...cd };
        }
        // avg >= 60%: NOT detained → skip
      } else {
        const singlePct = cd.theoryPct ?? cd.practicalPct;
        if (singlePct < SUBJECT_THRESHOLD) {
          detainedCourses[cd.code] = { ...cd };
        }
      }
    }

    // Step 3: Feed into tableIIIA_map
    for (const dc of Object.values(detainedCourses)) {
      if (!tableIIIA_map[dc.code]) {
        tableIIIA_map[dc.code] = { code: dc.code, fullName: dc.fullName, sn: dc.sn, studentMap: {} };
      }
      const studentKey = s.seatNo || s.rollNo;
      if (!tableIIIA_map[dc.code].studentMap[studentKey]) {
        tableIIIA_map[dc.code].studentMap[studentKey] = {
          rollNo:       s.rollNo,
          seatNo:       s.seatNo,
          name:         s.name,
          div:          s.div,
          theoryPct:    dc.theoryPct    !== null ? dc.theoryPct.toFixed(2)    : null,
          practicalPct: dc.practicalPct !== null ? dc.practicalPct.toFixed(2) : null,
        };
      }
    }

    // Step 4: Build tableIIIB
    if (Object.keys(detainedCourses).length > 0) {
      // Sort subjects by SN order
      const subjects = Object.values(detainedCourses)
        .sort((a, b) => a.sn - b.sn)
        .map(dc => ({
          code:     dc.code,
          fullName: dc.fullName,
          sn:       dc.sn,
          type: dc.theoryPct !== null && dc.practicalPct !== null
                ? 'Theory/Practical'
                : dc.theoryPct !== null ? 'Theory' : 'Practical',
          pct: dc.theoryPct !== null && dc.practicalPct !== null
               ? `${dc.theoryPct.toFixed(2)} % / ${dc.practicalPct.toFixed(2)} %`
               : `${(dc.theoryPct ?? dc.practicalPct).toFixed(2)} %`,
        }));

      tableIIIB.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        div:     s.div,
        overall: overall !== null ? overall.toFixed(2) : 'N/A',
        subjects,
      });
    }
  }

  // ── Convert tableIIIA_map → ordered array ────────────────────────────────────
  const tableIIIA = Object.values(tableIIIA_map)
    .map(c => ({
      code:     c.code,
      fullName: c.fullName,
      sn:       c.sn,
      rows: Object.values(c.studentMap)
        .filter(st => {
          // Re-apply avg filter
          if (st.theoryPct !== null && st.practicalPct !== null) {
            const avg = (parseFloat(st.theoryPct) + parseFloat(st.practicalPct)) / 2;
            return avg < SUBJECT_THRESHOLD;
          }
          return true;
        })
        .map(st => ({
          rollNo:       st.rollNo,
          seatNo:       st.seatNo,
          name:         st.name,
          div:          st.div,
          theoryPct:    st.theoryPct,
          practicalPct: st.practicalPct,
          type: st.theoryPct !== null && st.practicalPct !== null
                ? 'Theory/Practical'
                : st.theoryPct !== null ? 'Theory' : 'Practical',
          pct: st.theoryPct !== null && st.practicalPct !== null
               ? `${st.theoryPct} % / ${st.practicalPct} %`
               : `${st.theoryPct ?? st.practicalPct} %`,
        })),
    }))
    .filter(c => c.rows.length > 0)
    .sort((a, b) => a.sn - b.sn);  // Sort by SN from Master Course Mapping

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIIIB.length,
    detainedSubjects: tableIIIA.reduce((acc, c) => acc + c.rows.length, 0),
  };

  return { tableI, tableII, tableIIIA, tableIIIB, stats };
}

module.exports = { processAttendanceY2new, COURSE_MAP_Y2 };
