/**
 * wordGeneratorY2.js
 * Generates the 2nd Year (4th Semester) RBU Detention List .docx.
 * Matches the reference document: 23_04_26_IV SEM 25-26 FINAL 10_37AM.docx
 *
 * Table I:    Exam Seat No | Name | Aggregate Attendance %
 * Table II:   Exam Seat No | Name | Aggregate Attendance %
 * Table III(A): SN | Course Code | Course Name | Sec/Div. | Exam Seat No | Name | Attendance (%)
 * Table III(B): SN | Sec/Div. | Exam Seat No | Name | Overall Attendance (%) | Course Code | Course Name | Attendance (%)
 *
 * Sorting:
 *   - Tables I & II: sorted by Seat No (caller's responsibility before passing in)
 *   - Table III(A): sorted by SN from Master Course Mapping (done in processor)
 *   - Table III(B): sorted by SN of the FIRST detained subject per student (done by caller/processor)
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  VerticalMergeType,
} = require('docx');
const fs = require('fs');

// ── Border styles ─────────────────────────────────────────────────────────────
const BORDER      = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDERS  = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'auto' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'auto' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'auto' },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  insideVertical:   { style: BorderStyle.NONE, size: 0, color: 'auto' },
};

function formatDate(dateStr) {
  if (!dateStr) return '';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) return dateStr;
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    const [y, m, d] = parts;
    return `${d}/${m}/${y}`;
  }
  return dateStr;
}

function buildSignatureTable(coordinatorName, hodName, branchName) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 4513, type: WidthType.DXA },
            borders: NO_BORDERS,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [normal(coordinatorName, 22)],
                spacing: { before: 800, after: 40 },
              }),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [normal('Academic Coordinator', 22)],
              }),
            ],
          }),
          new TableCell({
            width: { size: 4513, type: WidthType.DXA },
            borders: NO_BORDERS,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [normal(hodName, 22)],
                spacing: { before: 800, after: 40 },
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [normal(`H.O.D, ${branchName}`, 22)],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ── Typography helpers ────────────────────────────────────────────────────────
const bold   = (text, size = 20) => new TextRun({ text: String(text ?? ''), bold: true,  size, font: 'Times New Roman' });
const normal = (text, size = 20) => new TextRun({ text: String(text ?? ''), bold: false, size, font: 'Times New Roman' });
const italic = (text, size = 20) => new TextRun({ text: String(text ?? ''), italics: true, size, font: 'Times New Roman' });

function para(children, alignment = AlignmentType.LEFT, spacingAfter = 80) {
  return new Paragraph({ alignment, children, spacing: { after: spacingAfter } });
}
function centerPara(children, spacingAfter = 80) {
  return para(children, AlignmentType.CENTER, spacingAfter);
}
function sectionHeading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [bold(text, 22)],
  });
}

// ── Cell builders ─────────────────────────────────────────────────────────────
function headerCell(text, width, colSpan = 1, rowSpan = 1) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    columnSpan: colSpan,
    rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [bold(text, 18)],
    })],
  });
}

function dataCell(text, width, alignment = AlignmentType.CENTER, verticalMerge = undefined) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    verticalMerge,
    children: [new Paragraph({
      alignment,
      children: [normal(text ?? '', 18)],
    })],
  });
}

// ── Table I & II ──────────────────────────────────────────────────────────────
// Cols: Exam Seat No | Name | Aggregate Attendance %
const T12_SEAT   = 2500;
const T12_NAME   = 4526;
const T12_ATT    = 2000;
const T12_WIDTHS = [T12_SEAT, T12_NAME, T12_ATT];

function buildTableI_II(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Exam Seat No',          T12_SEAT),
      headerCell('Name',                  T12_NAME),
      headerCell('Aggregate Attendance %', T12_ATT),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map((r) => new TableRow({
        children: [
          dataCell(r.seatNo || r.rollNo || '', T12_SEAT),
          dataCell(r.name ?? '',               T12_NAME, AlignmentType.LEFT),
          dataCell(r.overall ?? '',            T12_ATT),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('—', T12_SEAT),
        dataCell('No students', T12_NAME, AlignmentType.LEFT),
        dataCell('—', T12_ATT),
      ] })];

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T12_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table III(A) ──────────────────────────────────────────────────────────────
// Cols: SN | Course Code | Course Name | Sec/Div. | Exam Seat No | Name | Attendance (%)
const T3A_SN   = 500;
const T3A_CODE = 1700;
const T3A_CNAME = 2200;
const T3A_DIV   = 700;
const T3A_SEAT  = 1500;
const T3A_STU   = 1726;
const T3A_ATT   = 700;
// sum = 500+1700+2200+700+1500+1726+700 = 9026
const T3A_WIDTHS = [T3A_SN, T3A_CODE, T3A_CNAME, T3A_DIV, T3A_SEAT, T3A_STU, T3A_ATT];

function buildTableIIIA(courses) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',             T3A_SN),
      headerCell('Course Code',    T3A_CODE),
      headerCell('Course Name',    T3A_CNAME),
      headerCell('Sec/Div.',       T3A_DIV),
      headerCell('Exam Seat No',   T3A_SEAT),
      headerCell('Name',           T3A_STU),
      headerCell('Attendance (%)', T3A_ATT),
    ],
  });

  const dataRows = [];
  if (courses.length === 0) {
    dataRows.push(new TableRow({
      children: [
        dataCell('—', T3A_SN),
        dataCell('No detained students',
          T3A_CODE + T3A_CNAME + T3A_DIV + T3A_SEAT + T3A_STU + T3A_ATT,
          AlignmentType.CENTER),
      ],
    }));
  } else {
    let sn = 1;
    for (const c of courses) {
      // Sort rows within each course by Seat No
      const sortedRows = [...c.rows].sort((a, b) => {
        const sa = (a.seatNo || a.rollNo || '').toUpperCase();
        const sb = (b.seatNo || b.rollNo || '').toUpperCase();
        return sa < sb ? -1 : sa > sb ? 1 : 0;
      });

      sortedRows.forEach((r, idx) => {
        const isFirst = idx === 0;
        const rowCount = sortedRows.length;
        const vMerge = rowCount > 1
          ? (isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE)
          : undefined;

        dataRows.push(new TableRow({
          children: [
            dataCell(isFirst ? String(sn) : '',      T3A_SN,   AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? c.code : '',           T3A_CODE, AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? c.fullName : '',       T3A_CNAME, AlignmentType.LEFT,  vMerge),
            dataCell(r.div   ?? '',                   T3A_DIV),
            dataCell(r.seatNo || r.rollNo || '',      T3A_SEAT),
            dataCell(r.name  ?? '',                   T3A_STU, AlignmentType.LEFT),
            dataCell(r.pct   ?? '',                   T3A_ATT),
          ],
        }));
      });
      sn++;
    }
  }

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T3A_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table III(B) ──────────────────────────────────────────────────────────────
// Cols: SN | Sec/Div. | Exam Seat No | Name | Overall Att (%) | Course Code | Course Name | Attendance (%)
const T3B_SN      = 500;
const T3B_DIV     = 700;
const T3B_SEAT    = 1500;
const T3B_NAME    = 1700;
const T3B_OVERALL = 800;
const T3B_CODE    = 1200;
const T3B_CNAME   = 1626;
const T3B_ATT     = 1000;
// sum = 500+700+1500+1700+800+1200+1626+1000 = 9026
const T3B_WIDTHS = [T3B_SN, T3B_DIV, T3B_SEAT, T3B_NAME, T3B_OVERALL, T3B_CODE, T3B_CNAME, T3B_ATT];

function buildTableIIIB(studentList) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',                      T3B_SN),
      headerCell('Sec/Div.',                T3B_DIV),
      headerCell('Exam Seat No',            T3B_SEAT),
      headerCell('Name',                    T3B_NAME),
      headerCell('Overall Attendance (%)',  T3B_OVERALL),
      headerCell('Course Code',             T3B_CODE),
      headerCell('Course Name',             T3B_CNAME),
      headerCell('Attendance (%)',          T3B_ATT),
    ],
  });

  const dataRows = [];
  if (studentList.length === 0) {
    dataRows.push(new TableRow({
      children: T3B_WIDTHS.map((w, i) =>
        i === 0 ? dataCell('—', w) : dataCell(i === 2 ? 'No detained students' : '', w)
      ),
    }));
  } else {
    let sn = 1;
    for (const s of studentList) {
      s.subjects.forEach((sub, idx) => {
        const isFirst = idx === 0;
        const subCount = s.subjects.length;
        const vMerge = subCount > 1
          ? (isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE)
          : undefined;

        dataRows.push(new TableRow({
          children: [
            dataCell(isFirst ? String(sn) : '',              T3B_SN,      AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? (s.div || '') : '',           T3B_DIV,     AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? (s.seatNo || s.rollNo) : '',  T3B_SEAT,    AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? s.name : '',                   T3B_NAME,    AlignmentType.LEFT,   vMerge),
            dataCell(isFirst ? s.overall : '',                T3B_OVERALL, AlignmentType.CENTER, vMerge),
            dataCell(sub.code     ?? '',                      T3B_CODE),
            dataCell(sub.fullName ?? '',                      T3B_CNAME,   AlignmentType.LEFT),
            dataCell(sub.pct      ?? '',                      T3B_ATT),
          ],
        }));
      });
      sn++;
    }
  }

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T3B_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Main generator ────────────────────────────────────────────────────────────
async function generateWordY2(meta, { tableI, tableII, tableIIIA, tableIIIB }, outputPath) {
  const { examName, schoolName, programme, semester, date } = meta;
  const coordinatorName = meta.coordinatorName || 'Dr. Anju Gupta';
  const hodName          = meta.hodName || 'Dr. N.P. Narkhede';
  const branchName       = meta.branchName || 'EN';

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Times New Roman', size: 22 } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
        },
      },
      children: [

        // ── University heading ─────────────────────────────────────────────────
        centerPara([bold('RAMDEOBABA UNIVERSITY, NAGPUR', 28)], 40),
        centerPara([bold('DETENTION LIST', 24)], 40),
        centerPara([bold(examName, 24)], 40),

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 120 },
          children: [normal(`Date: ${formatDate(date)}`, 22)],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 120 },
          children: [
            normal('School: ', 22),
            normal(schoolName, 22),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 120 },
          children: [
            normal('Programme: ', 22),
            bold(programme, 22),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 180 },
          children: [
            normal('Semester: ', 22),
            bold(semester, 22),
          ],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 180 },
          children: [normal('Submitted to Vice-Chancellor for Approval through Dean Academics', 22)],
        }),

        // ── Table I ───────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 120 },
          children: [
            normal('Table I is the list of students having aggregate attendance less than 75%.', 22),
          ],
        }),
        new Paragraph({
          spacing: { after: 180 },
          children: [
            normal('As per the ordinances/ regulations of the university these students should be detained in the ', 22),
            bold(examName, 22),
            normal(' examination in all courses.', 22),
          ],
        }),
        sectionHeading('Table I'),
        buildTableI_II(tableI),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table II ──────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 180 },
          children: [
            normal('However Table II, is the list of students having aggregate attendance between 60% and 75% and have applied for condonation of attendance. The application forms and medical certificates/other relevant documents have been verified. After due consideration and it is recommended that these students may be permitted to appear in all courses in the ', 22),
            bold(examName, 22),
            normal(' examinations, since the absence was due to circumstances beyond the control of the students.', 22),
          ],
        }),
        sectionHeading('Table II'),
        buildTableI_II(tableII),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(A) ──────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 180 },
          children: [
            normal('Table III is the list of courses along with the student details and the attendance (%) in the courses in which they are recommended to be detained.', 22),
          ],
        }),
        sectionHeading('Table III(A) – Course wise Detention list'),
        buildTableIIIA(tableIIIA),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(B) ──────────────────────────────────────────────────────
        sectionHeading('Table III (B) – Student wise Detention list'),
        buildTableIIIB(tableIIIB),
        new Paragraph({ spacing: { after: 240 }, children: [] }),

        // ── Signature footer ──────────────────────────────────────────────────
        buildSignatureTable(coordinatorName, hodName, branchName),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  return outputPath;
}

module.exports = { generateWordY2 };
