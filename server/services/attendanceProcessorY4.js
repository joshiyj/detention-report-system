/**
* attendanceProcessorY4.js
* Processes 4th Year (8th Semester) ECS attendance sheets.
*
* Rules:
*   - Null / blank subject cell → student NOT enrolled in that course (skip)
*   - Overall < 75% → Table I
*   - 60% ≤ Overall < 75% AND (Roll No or Seat No condoned) → Table II
*   - Subject < 60% → subject detention in Table III
*   - Preserves original raw merge order from the sheets (no alphabetical sorting).
*/

const OVERALL_THRESHOLD = 75;
const CONDONATION_MIN   = 60;
const SUBJECT_THRESHOLD = 60;

// ── Semester VIII Course Mapping ───────────────────────────────────────────────
// Key = EXACT column header in Excel file
const COURSE_MAP_Y4 = {
  'VSP':       { sn: 1, code: 'ECST407-1', fullName: 'VLSI Signal Processing',               type: 'Theory' },
  'CPI':       { sn: 2, code: 'ECST407-2', fullName: 'Cybersecurity and Privacy in IoT',     type: 'Theory' },
  'GAN':       { sn: 3, code: 'ECST407-3', fullName: 'Generative Adversarial Network',       type: 'Theory' },
  'GAD':       { sn: 3, code: 'ECST407-3', fullName: 'Generative Adversarial Network',       type: 'Theory' },
  'BDW':       { sn: 4, code: 'ECST407-4', fullName: 'Big data web intelligence',            type: 'Theory' },
  'BDWI':      { sn: 4, code: 'ECST407-4', fullName: 'Big data web intelligence',            type: 'Theory' },
  'NE':        { sn: 5, code: 'ECST408-1', fullName: 'Nano electronics',                     type: 'Theory' },
  'NANO':      { sn: 5, code: 'ECST408-1', fullName: 'Nano electronics',                     type: 'Theory' },
  'AV':        { sn: 6, code: 'ECST408-2', fullName: 'Autonomous Vehicle',                    type: 'Theory' },
  'RL':        { sn: 7, code: 'ECST408-3', fullName: 'Reinforcement Learning',               type: 'Theory' },
  'BIFO':      { sn: 8, code: 'ECST408-4', fullName: 'Bioinformatics',                       type: 'Theory' },
  'PROJ-III':  { sn: 9, code: 'ECSP409',   fullName: 'Project III',                          type: 'Practical' },
  'INC':       { sn: 999, code: 'INC',     fullName: 'INC',                                  type: 'Theory' },
};

function resolveCourseY4(abbr, activeCourseMap) {
  const trimmed = abbr ? abbr.trim() : '';
  const map = activeCourseMap || COURSE_MAP_Y4;
  const mapped = map[trimmed];
  if (mapped) return mapped;
  // Unknown column — fallback
  const type = trimmed.toLowerCase().includes('lab') || trimmed.toLowerCase().includes('proj') ? 'Practical' : 'Theory';
  return { sn: 999, code: trimmed, fullName: trimmed, type };
}

/**
* processAttendanceY4
* @param {Array} students      — from excelParser (with seatNo and overallRaw present)
* @param {Array} subjectNames  — ordered list of subject column headers
* @param {Set}   condonationIds — Roll Numbers or Seat Numbers applied for condonation
* @param {Object} dynamicMapping — parsed docx course mapping (optional)
*/
function processAttendanceY4(students, subjectNames, condonationIds = new Set(), dynamicMapping = null) {
  const activeCourseMap = dynamicMapping ? dynamicMapping.courseMap : COURSE_MAP_Y4;
  const tableI   = [];
  const tableII  = [];
  const tableIII = []; // Student-wise

  for (const s of students) {
    const overall = s.overallPct;
    // Format: "48/67 (71.64)" or fallback to percentage string if overallRaw not present
    const overallDisplay = s.overallRaw || (overall !== null ? `${overall.toFixed(2)}%` : 'N/A');

    // ── Table I: overall < 75% ────────────────────────────────────────────────
    if (overall !== null && overall < OVERALL_THRESHOLD) {
      tableI.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overallDisplay,
      });
    }

    // ── Table II: 60% ≤ overall < 75% AND applied for condonation ─────────────
    const hasCondonation = condonationIds.has(s.rollNo) || condonationIds.has(s.seatNo);
    if (
      overall !== null &&
      overall >= CONDONATION_MIN &&
      overall < OVERALL_THRESHOLD &&
      hasCondonation
    ) {
      tableII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overallDisplay,
      });
    }

    // ── Table III — per-subject detention ─────────────────────────────────────
    // For Year 4, subjects are theory-only or practical-only, no paired components.
    // If a subject's attendance is not null and is < 60%, they are detained in it.
    const detainedCourses = [];

    for (const subj of subjectNames) {
      const pct = s.subjects[subj];
      if (pct === null || pct === undefined) continue; // not enrolled / empty

      if (pct < SUBJECT_THRESHOLD) {
        const { sn, code, fullName } = resolveCourseY4(subj, activeCourseMap);
        detainedCourses.push({
          sn,
          code,
          fullName,
          pct: pct.toFixed(2) + ' %',
        });
      }
    }

    // Sort detained courses by SN (syllabus sequence order)
    detainedCourses.sort((a, b) => a.sn - b.sn);

    if (detainedCourses.length > 0 && (overall === null || overall < OVERALL_THRESHOLD)) {
      tableIII.push({
        rollNo:  s.rollNo,
        seatNo:  s.seatNo,
        name:    s.name,
        overall: overallDisplay,
        courses: detainedCourses,
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

module.exports = { processAttendanceY4, COURSE_MAP_Y4 };
