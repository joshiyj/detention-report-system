/**
 * wordGenerator.js
 * Generates the RBU Detention List .docx using the docx npm package.
 * Matches the official RBU format as closely as possible.
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, PageOrientation, UnderlineType, VerticalMergeType,
} = require('docx');
const fs = require('fs');

// ── Shared border styles ──────────────────────────────────────────────────────
const BORDER = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

// ── Typography helpers ────────────────────────────────────────────────────────
const bold   = (text, size = 20) => new TextRun({ text, bold: true,  size, font: 'Times New Roman' });
const normal = (text, size = 20) => new TextRun({ text, bold: false, size, font: 'Times New Roman' });
const italic = (text, size = 20) => new TextRun({ text, bold: false, italics: true, size, font: 'Times New Roman' });

function para(children, alignment = AlignmentType.LEFT, spacingAfter = 80) {
  return new Paragraph({ alignment, children, spacing: { after: spacingAfter } });
}

function centerPara(children, spacingAfter = 80) {
  return para(children, AlignmentType.CENTER, spacingAfter);
}

// ── Table cell builders ───────────────────────────────────────────────────────
function headerCell(text, width, colSpan = 1, rowSpan = 1) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    columnSpan: colSpan,
    rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [bold(text, 19)],
    })],
  });
}

function dataCell(text, width, alignment = AlignmentType.CENTER, verticalMerge = undefined) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    verticalMerge,
    children: [new Paragraph({
      alignment,
      children: [normal(text ?? '', 19)],
    })],
  });
}

function emptyCell(width, verticalMerge = undefined) {
  return dataCell('', width, AlignmentType.CENTER, verticalMerge);
}

// ── Section heading ───────────────────────────────────────────────────────────
function sectionHeading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [bold(text, 22)],
  });
}

// ── Table I ───────────────────────────────────────────────────────────────────
// Cols: Roll No. | Seat No | Student name | Aggregate Attendance %
// Total width = 9026 DXA (A4 with ~1" margins)
const T1_WIDTHS = [1126, 1800, 3900, 2200]; // sum = 9026

function buildTableI(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Roll No.',              T1_WIDTHS[0]),
      headerCell('Seat No',               T1_WIDTHS[1]),
      headerCell('Student name',          T1_WIDTHS[2]),
      headerCell('Aggregate Attendance %',T1_WIDTHS[3]),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map(r => new TableRow({
        children: [
          dataCell(r.rollNo  ?? '',        T1_WIDTHS[0]),
          dataCell(r.seatNo  ?? '',        T1_WIDTHS[1]),
          dataCell(r.name    ?? '',        T1_WIDTHS[2], AlignmentType.LEFT),
          dataCell(r.overall ?? '',        T1_WIDTHS[3]),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('—',           T1_WIDTHS[0]),
        dataCell('—',           T1_WIDTHS[1]),
        dataCell('No students', T1_WIDTHS[2], AlignmentType.LEFT),
        dataCell('—',           T1_WIDTHS[3]),
      ] })];

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T1_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table II ──────────────────────────────────────────────────────────────────
// Cols: Roll No. | Seat No | Student name | Aggregate Attendance %
const T2_WIDTHS = [1126, 1800, 3900, 2200]; // sum = 9026

function buildTableII(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Roll No.',              T2_WIDTHS[0]),
      headerCell('Seat No',               T2_WIDTHS[1]),
      headerCell('Student name',          T2_WIDTHS[2]),
      headerCell('Aggregate Attendance %',T2_WIDTHS[3]),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map(r => new TableRow({
        children: [
          dataCell(r.rollNo  ?? '',        T2_WIDTHS[0]),
          dataCell(r.seatNo  ?? '',        T2_WIDTHS[1]),
          dataCell(r.name    ?? '',        T2_WIDTHS[2], AlignmentType.LEFT),
          dataCell(r.overall ?? '',        T2_WIDTHS[3]),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('—',           T2_WIDTHS[0]),
        dataCell('—',           T2_WIDTHS[1]),
        dataCell('No students', T2_WIDTHS[2], AlignmentType.LEFT),
        dataCell('—',           T2_WIDTHS[3]),
      ] })];

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T2_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table III(A) – Course wise ────────────────────────────────────────────────────
// Cols: SN | Course Code | Course Name | Type | Sec/Div. | Exam Seat No | Name | Attendance (%)
// Two-row header: first row has merged "Student Details" spanning cols 5-8
const T3A_SN   = 400;
const T3A_CODE = 1300;
const T3A_NAME = 1900;
const T3A_TYPE = 900;
const T3A_DIV  = 700;
const T3A_SEAT = 1400;
const T3A_STU  = 1626;
const T3A_ATT  = 800;
// sum = 400+1300+1900+900+700+1400+1626+800 = 9026
const T3A_WIDTHS = [T3A_SN, T3A_CODE, T3A_NAME, T3A_TYPE, T3A_DIV, T3A_SEAT, T3A_STU, T3A_ATT];

function buildTableIIIA(courses) {
  // Header row 1: SN | Course Code | Course Name | Student Details (span 4)
  const studentDetailsWidth = T3A_TYPE + T3A_DIV + T3A_SEAT + T3A_STU + T3A_ATT;
  const hRow1 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',              T3A_SN,   1, 2),
      headerCell('Course Code',     T3A_CODE, 1, 2),
      headerCell('Course Name',     T3A_NAME, 1, 2),
      headerCell('Student Details', studentDetailsWidth, 5, 1),
    ],
  });

  // Header row 2: Type | Sec/Div. | Exam Seat No | Name | Attendance (%)
  const hRow2 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Type',           T3A_TYPE),
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
        dataCell('No detained students', T3A_CODE + T3A_NAME + T3A_TYPE + T3A_DIV + T3A_SEAT + T3A_STU + T3A_ATT, AlignmentType.LEFT),
      ],
    }));
  } else {
    let sn = 1;
    for (const c of courses) {
      c.rows.forEach((r, idx) => {
        const vMerge = c.rows.length > 1
          ? (idx === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE)
          : undefined;

        dataRows.push(new TableRow({
          children: [
            dataCell(idx === 0 ? String(sn) : '',   T3A_SN,   AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? c.code : '',       T3A_CODE, AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? c.fullName : '',   T3A_NAME, AlignmentType.LEFT,   vMerge),
            dataCell(c.type,    T3A_TYPE),
            dataCell(r.rollNo,  T3A_DIV),
            dataCell(r.seatNo,  T3A_SEAT),
            dataCell(r.name,    T3A_STU, AlignmentType.LEFT),
            dataCell(r.pct + ' %', T3A_ATT),
          ],
        }));
      });
      sn++;
    }
  }

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T3A_WIDTHS,
    rows: [hRow1, hRow2, ...dataRows],
  });
}

// ── Table III(B) – Student wise ─────────────────────────────────────────────────────────
// SN | Sec/Div | Exam Seat No | Name | Overall Att | Course Code | Course Name | Type | Attendance
const T3B_SN      = 300;
const T3B_DIV     = 600;
const T3B_SEAT    = 1300;
const T3B_STUN    = 1600;
const T3B_OVERALL = 800;
const T3B_CODE    = 1100;
const T3B_CNAME   = 1626; // adjusted so total = 9026
const T3B_TYPE    = 700;
const T3B_ATT     = 1000;
// sum = 300+600+1300+1600+800+1100+1626+700+1000 = 9026
const T3B_WIDTHS = [T3B_SN, T3B_DIV, T3B_SEAT, T3B_STUN, T3B_OVERALL, T3B_CODE, T3B_CNAME, T3B_TYPE, T3B_ATT];

function buildTableIIIB(studentList) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',                    T3B_SN),
      headerCell('Sec/Div.',              T3B_DIV),
      headerCell('Exam Seat No',          T3B_SEAT),
      headerCell('Name',                  T3B_STUN),
      headerCell('Overall Attendance',    T3B_OVERALL),
      headerCell('Course Code',           T3B_CODE),
      headerCell('Course Name',           T3B_CNAME),
      headerCell('Type',                  T3B_TYPE),
      headerCell('Attendance (%)',         T3B_ATT),
    ],
  });

  const dataRows = [];
  if (studentList.length === 0) {
    dataRows.push(new TableRow({
      children: T3B_WIDTHS.map((w, i) => i === 0 ? dataCell('—', w) : dataCell(i === 1 ? 'No detained students' : '', w)),
    }));
  } else {
    let sn = 1;
    for (const s of studentList) {
      s.subjects.forEach((sub, idx) => {
        const vMerge = s.subjects.length > 1
          ? (idx === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE)
          : undefined;

        dataRows.push(new TableRow({
          children: [
            dataCell(idx === 0 ? String(sn) : '',   T3B_SN,      AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? s.rollNo : '',     T3B_DIV,     AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? s.seatNo : '',     T3B_SEAT,    AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? s.name : '',       T3B_STUN,    AlignmentType.LEFT,   vMerge),
            dataCell(idx === 0 ? s.overall : '',    T3B_OVERALL, AlignmentType.CENTER, vMerge),
            dataCell(sub.code,     T3B_CODE),
            dataCell(sub.fullName, T3B_CNAME, AlignmentType.LEFT),
            dataCell(sub.type,     T3B_TYPE),
            dataCell(sub.pct + ' %', T3B_ATT),
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
async function generateWord(meta, { tableI, tableII, tableIIIA, tableIIIB }, outputPath) {
  const { examName, schoolName, programme, semester, date } = meta;

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Times New Roman', size: 20 } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
        },
      },
      children: [

        // ── University heading ────────────────────────────────────────────────
        centerPara([bold('RAMDEOBABA UNIVERSITY, NAGPUR', 28)], 40),
        centerPara([bold('DETENTION LIST', 26)], 40),
        centerPara([bold(examName, 24)], 40),
        centerPara([normal(`Date: ${date}`, 20)], 80),

        // School / Programme / Semester
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 40 },
          children: [normal(`School: ${schoolName}`, 20)],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 160 },
          children: [
            normal(`Programme: ${programme}`, 20),
            new TextRun({ text: '     ', font: 'Times New Roman', size: 20 }),
            normal(`Semester: ${semester}`, 20),
          ],
        }),

        // Submitted to VC line
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 160 },
          children: [italic('Submitted to Vice-Chancellor for Approval through Dean Academics', 19)],
        }),

        // ── Table I ───────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Table I is the list of students having aggregate attendance less than 75%.', 19),
          ],
        }),
        new Paragraph({
          spacing: { after: 100 },
          children: [
            normal('As per the ordinances/ regulations of the university these students should be detained in the ', 19),
            bold(examName, 19),
            normal(' examination in all courses.', 19),
          ],
        }),
        sectionHeading('Table I'),
        buildTableI(tableI),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table II ──────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('However Table II, is the list of students having aggregate attendance between 60% and 75% and have applied for condonation of attendance. The application forms and medical certificates/other relevant documents have been verified. After due consideration and it is recommended that these students may be permitted to appear in all courses in the ', 19),
            bold(examName, 19),
            normal(' examinations, since the absence was due to circumstances beyond the control of the students.', 19),
          ],
        }),
        sectionHeading('Table II'),
        buildTableII(tableII),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(A) ──────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Table III is the list of courses along with the student details and the attendance(%) in the courses in which they are recommended to be detained.', 19),
          ],
        }),
        sectionHeading('Table III (A) – Course wise Detention list'),
        buildTableIIIA(tableIIIA),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(B) ──────────────────────────────────────────────────────
        sectionHeading('Table III (B) – Student wise Detention list'),
        buildTableIIIB(tableIIIB),
        new Paragraph({ spacing: { after: 240 }, children: [] }),

        // ── Signature footer ──────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Prepared by:', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t\t', size: 20 }),
            normal('H.O.D.', 20),
          ],
          tabStops: [{ type: 'right', position: 9026 }],
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [
            normal('Name & Signature', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t', size: 20 }),
            normal('Name & Signature', 20),
          ],
        }),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  return outputPath;
}

module.exports = { generateWord };
