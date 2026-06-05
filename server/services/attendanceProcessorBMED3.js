/**
 * attendanceProcessorBMED3.js
 * Processes 3rd Year (6th Semester) Biomedical Engineering attendance sheets.
 *
 * Excel structure (vi bme.xls):
 *   Col 0: Roll No.     (e.g. "38")
 *   Col 1: Unique Id    (email, e.g. "vyasss_2@rknec.edu")
 *   Col 2: Seat No      (e.g. "BMU26RS6032")
 *   Col 3: Student name
 *   Col 4: OverAll Attendance  (e.g. "129/201 (64.18)")
 *   Col 5+: Subject columns (BMICRO, BPPD, MLH, BIP, OOPS, RTOS, SSD, INC, BML LAB, BIP LAB, PR-II)
 *
 * Rules:
 *   - Null / blank subject cell → student NOT enrolled in that course (skip)
 *   - Overall < 75% → Table I
 *   - 60% ≤ Overall < 75% AND has pouring attendance → Table II
 *   - Subject/avg < 60% → subject detention in Table III
 *   - If Theory+Practical both exist for a course: apply avg rule
 *     - Pair MLH (BMT6003) & BML LAB (BMP6003)
 *     - Pair BIP (BMT6004) & BIP LAB (BMP6004)
 *   - Sort: Table I & II by Seat No; Table III by Seat No
 */

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── Semester VI Course Mapping (from Biomedical_Semester_VI_Course_Mapping.docx) ──
// Key = EXACT column header in Excel file
const COURSE_MAP_BMED3 = {
  'BMICRO':  { sn: 1,  code: 'BMT6001',   fullName: 'Biomedical Microsystems',                     type: 'Theory'    },
  'BPPD':    { sn: 2,  code: 'BMT6002',   fullName: 'Biomedical Product and Prototype Design',     type: 'Theory'    },
  'MLH':     { sn: 3,  code: 'BMT6003',   fullName: 'Machine Learning for Healthcare',             type: 'Theory'    },
  'MLH LAB': { sn: 4,  code: 'BMP6003',   fullName: 'Machine Learning for Healthcare Lab',         type: 'Practical' },
  'BML LAB': { sn: 4,  code: 'BMP6003',   fullName: 'Machine Learning for Healthcare Lab',         type: 'Practical' }, // alias to map excel header BML LAB
  'BIP':     { sn: 5,  code: 'BMT6004',   fullName: 'Biomedical Image Processing',                 type: 'Theory'    },
  'BIP LAB': { sn: 6,  code: 'BMP6004',   fullName: 'Biomedical Image Processing Lab',             type: 'Practical' },
  'MB':      { sn: 7,  code: 'BMT6005-1', fullName: 'Molecular Biology',                           type: 'Theory'    },
  'BNT':     { sn: 8,  code: 'BMT6005-2', fullName: 'Bionanotechnology',                           type: 'Theory'    },
  'FOR':     { sn: 9,  code: 'BMT6005-3', fullName: 'Fundamentals of Robotics',                    type: 'Theory'    },
  'OOPS':    { sn: 10, code: 'BMT6005-4', fullName: 'Object Oriented Programming',                 type: 'Theory'    },
  'ABI':     { sn: 11, code: 'BMT6006-1', fullName: 'Advanced Bioinformatics',                     type: 'Theory'    },
  'RME':     { sn: 12, code: 'BMT6006-2', fullName: 'Reliability of Medical Equipments',               type: 'Theory'    },
  'RTOS':    { sn: 13, code: 'BMT6006-3', fullName: 'RTOS for Embedded System',                        type: 'Theory'    },
  'TM':      { sn: 14, code: 'BMT6006-4', fullName: 'Telemedicine',                                    type: 'Theory'    },
  'SSD':     { sn: 15, code: 'BMT6007',   fullName: 'Soft Skill Development',                      type: 'Theory'    },
  'PR-II':   { sn: 16, code: 'BMP6008',   fullName: 'Project-II',                                  type: 'Practical' },
  'INC':     { sn: 999, code: 'INC',      fullName: 'INC',                                         type: 'Theory'    },
};

function resolveCourseBMED3(abbr, activeCourseMap) {
  const trimmed = abbr ? abbr.trim() : '';
  const map = activeCourseMap || COURSE_MAP_BMED3;
  const mapped = map[trimmed];
  if (mapped) return mapped;
  // Unknown column — fallback
  const type = trimmed.toLowerCase().includes('lab') ? 'Practical' : 'Theory';
  return { sn: 999, code: trimmed, fullName: trimmed, type };
}

/**
 * Normalizes course codes for Theory-Practical pairing.
 * E.g., BMT6003 & BMP6003 both map to BM6003 for pairing.
 */
function getPairBaseCode(code) {
  if (!code) return '';
  // Match prefix like BMT or BMP or ECST or ECSP and replace the T/P with a neutral tag
  return code.replace(/^([A-Z]{2,3})[TP]/i, '$1');
}

function isStudentCondoned(student, condonationInputSet) {
  if (!condonationInputSet || condonationInputSet.size === 0) return false;
  
  const studentSeat = (student.seatNo || '').trim().toLowerCase();
  const studentRoll = (student.rollNo || '').trim().toLowerCase();
  const studentName = (student.name || '').trim().toLowerCase();

  for (const input of condonationInputSet) {
    const cleanInput = input.trim().toLowerCase();
    if (!cleanInput) continue;
    
    // Check exact seat number match
    if (studentSeat === cleanInput) return true;
    
    // Check exact roll number match
    if (studentRoll === cleanInput) return true;

    // Check if name contains the input or vice versa
    if (studentName.includes(cleanInput) || cleanInput.includes(studentName)) return true;
  }
  return false;
}

/**
 * processAttendanceBMED3
 * @param {Array} students      — from excelParser (with seatNo present)
 * @param {Array} subjectNames  — ordered list of subject column headers
 * @param {Set}   condonationSeats — seat numbers / names / roll numbers applied for condonation
 * @param {Object} dynamicMapping — parsed docx course mapping (optional)
 */
function processAttendanceBMED3(students, subjectNames, condonationSeats = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP_BMED3;
  const tableI    = [];
  const tableII   = [];
  const tableIII  = []; // Student-wise: each entry has detained courses list

  const seenI = new Set();

  for (const s of students) {
    const overall = s.overallPct;

    // ── Table I: overall < 75% ────────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      const key = s.seatNo || s.rollNo;
      if (!seenI.has(key)) {
        seenI.add(key);
        tableI.push({
          rollNo:  s.rollNo,
          seatNo:  s.seatNo,
          name:    s.name,
          overall: overall.toFixed(2),
        });
      }
    }

    // ── Table II: 60% ≤ overall < 75% AND has pouring attendance and matches condonation ──
    const matchesCondonation = isStudentCondoned(s, condonationSeats);
    if (
      overall !== null &&
      overall >= CONDONATION_MIN &&
      overall < OVERALL_THRESHOLD &&
      s.hasPouringAttendance &&
      (condonationSeats.size === 0 || matchesCondonation)
    ) {
      tableII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overall.toFixed(2),
      });
    }

    // ── Table III — per-subject detention ─────────────────────────────────────
    // Step 1: Collect all subject pct values, grouped by course pair base code
    // Blank cells (null) = not enrolled → skip entirely
    const allCoursePairs = {}; // baseCode → { baseCode, theory: { code, fullName, sn, pct }, practical: { code, fullName, sn, pct } }

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct === null || pct === undefined) continue;

      const { sn, code, fullName, type } = resolveCourseBMED3(subj, activeCourseMap);
      const baseCode = getPairBaseCode(code);

      if (!allCoursePairs[baseCode]) {
        allCoursePairs[baseCode] = { baseCode, theory: null, practical: null };
      }

      const item = { code, fullName, sn, pct };
      if (type === 'Theory') {
        allCoursePairs[baseCode].theory = item;
      } else {
        allCoursePairs[baseCode].practical = item;
      }
    }

    // Step 2: Apply detention threshold + avg rule
    const detainedCourses = [];

    for (const cp of Object.values(allCoursePairs)) {
      const hasBoth = cp.theory !== null && cp.practical !== null;

      if (hasBoth) {
        const avg = (cp.theory.pct + cp.practical.pct) / 2;
        if (avg < SUBJECT_THRESHOLD) {
          detainedCourses.push(cp.theory);
          detainedCourses.push(cp.practical);
        }
      } else {
        const singleItem = cp.theory ?? cp.practical;
        if (singleItem.pct < SUBJECT_THRESHOLD) {
          detainedCourses.push(singleItem);
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
          pct:      dc.pct.toFixed(2),
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

module.exports = { processAttendanceBMED3, COURSE_MAP_BMED3 };
