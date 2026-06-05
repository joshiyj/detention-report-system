/**
 * attendanceProcessorY3.js
 * Processes 3rd Year (6th Semester) ECS attendance sheets.
 *
 * Excel structure (VI A updated 22.xls, etc.):
 *   Col 0: Roll No.     (e.g. "A-67")
 *   Col 1: Unique Id    (email, e.g. "ahsana@rknec.edu")
 *   Col 2: Seat No      (e.g. "ENCS26RS6020")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance  (e.g. "126/177 (71.19)")
 *   Col 5+: Subject columns
 *
 * Rules:
 *   - Null / blank subject cell → student NOT enrolled in that course (skip)
 *   - Overall < 75% → Table I
 *   - 60% ≤ Overall < 75% AND has pouring attendance → Table II
 *   - Subject/avg < 60% → subject detention in Table III
 *   - Sort: Table I & II by Seat No; Table III by Seat No
 *
 * Course Mapping: Semester_VI_Course_Mapping.docx (SN 1-19)
 */

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── Semester VI Course Mapping ─────────────────────────────────────────────────
// Key = EXACT column header in Excel file
const COURSE_MAP_Y3 = {
  'DBMS':      { sn: 1,  code: 'ECST6001',    fullName: 'Database Management System',               type: 'Theory'    },
  'DBMS LAB':  { sn: 2,  code: 'ECSP6001',    fullName: 'Database Management System Lab',           type: 'Practical' },
  'SVV':       { sn: 3,  code: 'ECST6002',    fullName: 'System Verilog for Verification',          type: 'Theory'    },
  'SVV LAB':   { sn: 4,  code: 'ECSP6002',    fullName: 'System Verilog for Verification Lab',      type: 'Practical' },
  'DL-I':      { sn: 5,  code: 'ECST6003-1',  fullName: 'Deep Learning-I',                          type: 'Theory'    },
  'DL-I LAB':  { sn: 6,  code: 'ECSP6003-1',  fullName: 'Deep Learning-I Lab',                      type: 'Practical' },
  'CVLSI':     { sn: 7,  code: 'ECST6003-2',  fullName: 'C Based VLSI Design',                      type: 'Theory'    },
  'CVLSI LAB': { sn: 8,  code: 'ECSP6003-2',  fullName: 'C Based VLSI Design Lab',                  type: 'Practical' },
  'DMW':       { sn: 9,  code: 'ECST6003-4',  fullName: 'Data Mining and Warehousing',              type: 'Theory'    },
  'DMW LAB':   { sn: 10, code: 'ECSP6003-4',  fullName: 'Data Mining and Warehousing Lab',          type: 'Practical' },
  'NLP':       { sn: 11, code: 'ECST6004-1',  fullName: 'Natural Language Processing',              type: 'Theory'    },
  'NLP LAB':   { sn: 12, code: 'ECSP6004-1',  fullName: 'Natural Language Processing Lab',          type: 'Practical' },
  'DT':        { sn: 13, code: 'ECST6004-2',  fullName: 'Design for Testability',                   type: 'Theory'    },
  'DT LAB':    { sn: 14, code: 'ECSP6004-2',  fullName: 'Design for Testability Lab',               type: 'Practical' },
  'BDWI':      { sn: 15, code: 'ECST6004-4',  fullName: 'Big Data Web Intelligence',                type: 'Theory'    },
  'BDWI LAB':  { sn: 16, code: 'ECSP6004-4',  fullName: 'Big Data Web Intelligence Lab',            type: 'Practical' },
  'DHV':       { sn: 17, code: 'ECST6005',    fullName: 'Data Handling and Visualization',          type: 'Theory'    },
  'SL-II':     { sn: 18, code: 'ECSP6006',    fullName: 'Software Laboratory-II',                   type: 'Practical' },
  'PR-I':      { sn: 19, code: 'ECSP6007',    fullName: 'Project-I',                                type: 'Practical' },
  // INC is not in the master mapping → unknown subject, assign high SN
  'INC':       { sn: 999, code: 'INC',         fullName: 'INC',                                      type: 'Theory'    },
};

function resolveCourseY3(abbr, activeCourseMap) {
  const trimmed = abbr ? abbr.trim() : '';
  const map = activeCourseMap || COURSE_MAP_Y3;
  const mapped = map[trimmed];
  if (mapped) return mapped;
  // Unknown column — fallback
  const type = trimmed.toLowerCase().includes('lab') ? 'Practical' : 'Theory';
  return { sn: 999, code: trimmed, fullName: trimmed, type };
}

/**
 * processAttendanceY3
 * @param {Array} students      — from excelParser (with seatNo present)
 * @param {Array} subjectNames  — ordered list of subject column headers
 * @param {Set}   condonationSeats — seat numbers applied for condonation
 * @param {Object} dynamicMapping — parsed docx course mapping (optional)
 */
function processAttendanceY3(students, subjectNames, condonationSeats = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP_Y3;
  const tableI    = [];
  const tableII   = [];
  const tableIII  = []; // Student-wise: each entry has detained courses list

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
    // Step 1: Collect all subject pct values, grouped by course code (to apply avg rule for T+P pairs)
    // Blank cells (null) = not enrolled → skip entirely
    const allCourseData = {}; // code → { code, fullName, sn, theoryPct, practicalPct }

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      // null = not enrolled in this course → skip
      if (pct === null || pct === undefined) continue;

      const { sn, code, fullName, type } = resolveCourseY3(subj, activeCourseMap);
      if (!allCourseData[code]) {
        allCourseData[code] = { code, fullName, sn, theoryPct: null, practicalPct: null };
      }
      if (type === 'Theory')    allCourseData[code].theoryPct    = pct;
      else                      allCourseData[code].practicalPct = pct;
    }

    // Step 2: Apply detention threshold + avg rule
    const detainedCourses = [];

    for (const cd of Object.values(allCourseData)) {
      const hasBoth = cd.theoryPct !== null && cd.practicalPct !== null;

      if (hasBoth) {
        const avg = (cd.theoryPct + cd.practicalPct) / 2;
        if (avg < SUBJECT_THRESHOLD) {
          detainedCourses.push({ ...cd });
        }
      } else {
        const singlePct = cd.theoryPct ?? cd.practicalPct;
        if (singlePct < SUBJECT_THRESHOLD) {
          detainedCourses.push({ ...cd });
        }
      }
    }

    // Sort detained courses by SN
    detainedCourses.sort((a, b) => a.sn - b.sn);

    if (detainedCourses.length > 0 && (overall === null || overall < OVERALL_THRESHOLD)) {
      tableIII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall !== null ? overall.toFixed(2) : 'N/A',
        courses: detainedCourses.map(dc => ({
          code:     dc.code,
          fullName: dc.fullName,
          sn:       dc.sn,
          type: dc.theoryPct !== null && dc.practicalPct !== null
                ? 'Theory/Practical'
                : dc.theoryPct !== null ? 'Theory' : 'Practical',
          pct: dc.theoryPct !== null && dc.practicalPct !== null
               ? `${dc.theoryPct.toFixed(2)} % / ${dc.practicalPct.toFixed(2)} %`
               : `${(dc.theoryPct ?? dc.practicalPct).toFixed(2)}`,
        })),
      });
    }
  }

  const stats = {
    totalStudents:    students.length,
    tableICount:      tableI.length,
    tableIICount:     tableII.length,
    detainedStudents: tableIII.length,
    detainedCourses:  tableIII.reduce((acc, s) => acc + s.courses.length, 0),
  };

  return { tableI, tableII, tableIII, stats };
}

module.exports = { processAttendanceY3, COURSE_MAP_Y3 };
