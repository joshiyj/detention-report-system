/**
 * attendanceProcessorBMED2.js
 * Processes 2nd Year (4th Semester) Biomedical Engineering attendance sheets.
 *
 * Excel structure (IV SEM BME.xls, IV SEM BME_1.xls):
 *   Col 0: Roll No.     (e.g. "21")
 *   Col 1: Unique Id    (email)
 *   Col 2: Seat No      (e.g. "UBEE26RS40004")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance  (e.g. "146/208 (70.19)")
 *   Col 5+: Subject columns:
 *     HAPE II, FAIML, SPA, BMECH, MIOT, PEIoT, INC, FED, IE, OE II,
 *     FAIML LAB, SPA LAB, DF3D, BCC, IE Lab
 *
 * Rules:
 *   - Null / blank subject cell → student NOT enrolled in that course (skip)
 *   - Overall < 75% → Table I
 *   - 60% ≤ Overall < 75% AND has pouring attendance → Table II
 *   - Subject/avg < 60% → subject detention in Table III
 *   - If Theory+Practical both exist for a course: apply avg rule
 *   - Sort: Table I & II by Seat No; Table III(A) & III(B) by SN
 */

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── Master Course Mapping (from Biomedical_Sem4_Course_Mapping.docx) ──────────
// Key = EXACT column header abbreviation in the Excel file.
const COURSE_MAP_BMED2 = {
  'HAPE II':    { sn: 1,  code: '24EE03TH0401',     fullName: 'Human Anatomy and Physiology for Engineers-II', type: 'Theory'    },
  'FAIML':      { sn: 2,  code: '24EE03TH0402',     fullName: 'Fundamentals of AI and Machine Learning',       type: 'Theory'    },
  'FAIML LAB':  { sn: 3,  code: '24EE03PR0402',     fullName: 'Fundamentals of AI and Machine Learning Lab',   type: 'Practical' },
  'SPA':        { sn: 4,  code: '24EE03TH0403',     fullName: 'Signals Processing and Analysis',               type: 'Theory'    },
  'SPA LAB':    { sn: 5,  code: '24EE03PR0403',     fullName: 'Signals Processing and Analysis Lab',           type: 'Practical' },
  'DF3D':       { sn: 6,  code: '24EE03PR0404',     fullName: 'Introduction to Digital Fabrication & 3D Printing', type: 'Practical' },
  'MIOT':       { sn: 7,  code: '24EE05TH0408-1',   fullName: 'Microcontrollers and IoT Applications',         type: 'Theory'    },
  'PEIoT':      { sn: 8,  code: '24EE01TH0405-2',   fullName: 'Programming for Environmental IoT',             type: 'Theory'    },
  'FED':        { sn: 9,  code: '24CS01TH0407-1',   fullName: 'Front End Development',                         type: 'Theory'    },
  'OE II':      { sn: 10, code: '24EEOEC03TH0406',  fullName: 'Open Elective-II / MOOCs',                      type: 'Theory'    },
  'BMECH':      { sn: 11, code: '24EE03TH0407',     fullName: 'Biomechanics',                                  type: 'Theory'    },
  'BCC':        { sn: 12, code: '24EE03PR0408',     fullName: 'Basic Competitive Coding',                      type: 'Practical' },
  'IE':         { sn: 13, code: '24SM03TH0401',     fullName: 'Innovations and Entrepreneurship',              type: 'Theory'    },
  'IE Lab':     { sn: 14, code: '24SM03PR0401',     fullName: 'Innovations and Entrepreneurship Lab',          type: 'Practical' },
  'INC':        { sn: 999, code: 'INC',             fullName: 'INC',                                           type: 'Theory'    },
};

/**
 * Resolve a subject abbreviation to course info.
 * Returns a generic fallback if not in COURSE_MAP_BMED2.
 */
function resolveCourse(abbr, activeCourseMap) {
  const map = activeCourseMap || COURSE_MAP_BMED2;
  const mapped = map[abbr ? abbr.trim() : ''];
  if (mapped) return mapped;
  const type = abbr && (abbr.toLowerCase().includes('lab') || abbr.toLowerCase().includes('practical'))
    ? 'Practical' : 'Theory';
  return { sn: 999, code: abbr, fullName: abbr, type };
}

/**
 * processAttendanceBMED2
 * @param {Array} students      — from excelParser (with seatNo present)
 * @param {Array} subjectNames  — ordered list of subject column headers
 * @param {Set}   condonationSeats — seat numbers applied for condonation
 * @param {Object} dynamicMapping — parsed docx course mapping (optional)
 */
function processAttendanceBMED2(students, subjectNames, condonationSeats = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP_BMED2;
  const tableI        = [];
  const tableII       = [];
  const tableIIIA_map = {};  // code → { code, fullName, sn, studentMap }
  const tableIIIB     = [];

  // Deduplicate Table I by seatNo
  const tableISeenSeats = new Set();

  for (const s of students) {
    const overall = s.overallPct;

    // ── Table I: overall < 75% ────────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      const key = s.seatNo || s.rollNo;
      if (!tableISeenSeats.has(key)) {
        tableISeenSeats.add(key);
        tableI.push({
          rollNo:  s.rollNo,
          seatNo:  s.seatNo,
          name:    s.name,
          overall: overall.toFixed(2),
        });
      }
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
    // Blank cells (null) = not enrolled → skip
    const allCourseData = {}; // code → { code, fullName, sn, theoryPct, practicalPct }

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct === null || pct === undefined) continue;

      const { sn, code, fullName, type } = resolveCourse(subj, activeCourseMap);
      if (!allCourseData[code]) {
        allCourseData[code] = { code, fullName, sn, theoryPct: null, practicalPct: null };
      }
      if (type === 'Theory')    allCourseData[code].theoryPct    = pct;
      else                      allCourseData[code].practicalPct = pct;
    }

    // Apply detention threshold + avg rule
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

    // Feed into tableIIIA_map
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

    // Build tableIIIB
    if (Object.keys(detainedCourses).length > 0) {
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

  // ── Convert tableIIIA_map → ordered array ───────────────────────────────────
  const tableIIIA = Object.values(tableIIIA_map)
    .map(c => ({
      code:     c.code,
      fullName: c.fullName,
      sn:       c.sn,
      rows: Object.values(c.studentMap)
        .filter(st => {
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
    .sort((a, b) => a.sn - b.sn);

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIIIB.length,
    detainedSubjects: tableIIIA.reduce((acc, c) => acc + c.rows.length, 0),
  };

  return { tableI, tableII, tableIIIA, tableIIIB, stats };
}

module.exports = { processAttendanceBMED2, COURSE_MAP_BMED2 };
