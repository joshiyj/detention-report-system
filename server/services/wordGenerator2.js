/**
 * wordGenerator2.js
 * Generates the 2nd Year (4th Semester) RBU Detention List .docx.
 * Matches the reference document format (23_04_26_IV SEM 25-26 FINAL 10_37AM.docx).
 *
 * Table III(A) columns: SN | Course Code | Course Name | Sec/Div. | Exam Seat No | Name | Attendance (%)
 * Table III(B) columns: SN | Sec/Div | Exam Seat No | Name | Overall Att (%) | Course Code | Course Name | Attendance (%)
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  VerticalMergeType,
} = require('docx');
const fs = require('fs');

// ── Border styles ─────────────────────────────────────────────────────────────
const BORDER     = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

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

// ── Table I ───────────────────────────────────────────────────────────────────
// Cols: Exam Seat No | Name | Aggregate Attendance %
// For 2nd year, Seat No may be empty — use Roll No instead
const T1_WIDTHS = [1500, 4526, 3000]; // sum = 9026

function buildTableI2(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Exam Seat No',            T1_WIDTHS[0]),
      headerCell('Name',                    T1_WIDTHS[1]),
      headerCell('Aggregate Attendance %',  T1_WIDTHS[2]),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map(r => new TableRow({
        children: [
          dataCell(r.seatNo || r.rollNo || '', T1_WIDTHS[0]),
          dataCell(r.name   ?? '',              T1_WIDTHS[1], AlignmentType.LEFT),
          dataCell(r.overall ?? '',             T1_WIDTHS[2]),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('—',           T1_WIDTHS[0]),
        dataCell('No students', T1_WIDTHS[1], AlignmentType.LEFT),
        dataCell('—',           T1_WIDTHS[2]),
      ] })];

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T1_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table II ──────────────────────────────────────────────────────────────────
function buildTableII2(rows) {
  return buildTableI2(rows); // identical columns for 2nd year
}

// ── Table III(A) – Course wise ────────────────────────────────────────────────
// Cols: SN | Course Code | Course Name | Sec/Div. | Exam Seat No | Name | Attendance (%)
// (No "Type" column in 2nd year reference doc)
const T3A_SN   = 500;
const T3A_CODE = 1500;
const T3A_NAME = 2000;
const T3A_DIV  = 800;
const T3A_SEAT = 1500;
const T3A_STU  = 1726;
const T3A_ATT  = 1000;
// sum = 500+1500+2000+800+1500+1726+1000 = 9026
const T3A_WIDTHS = [T3A_SN, T3A_CODE, T3A_NAME, T3A_DIV, T3A_SEAT, T3A_STU, T3A_ATT];

function buildTableIIIA2(courses) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',              T3A_SN),
      headerCell('Course Code',     T3A_CODE),
      headerCell('Course Name',     T3A_NAME),
      headerCell('Sec/Div.',        T3A_DIV),
      headerCell('Exam Seat No',    T3A_SEAT),
      headerCell('Name',            T3A_STU),
      headerCell('Attendance (%)',  T3A_ATT),
    ],
  });

  const dataRows = [];
  if (courses.length === 0) {
    dataRows.push(new TableRow({
      children: [
        dataCell('—', T3A_SN),
        dataCell('No detained students',
          T3A_CODE + T3A_NAME + T3A_DIV + T3A_SEAT + T3A_STU + T3A_ATT,
          AlignmentType.LEFT),
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
            dataCell(idx === 0 ? String(sn) : '',     T3A_SN,   AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? c.code : '',         T3A_CODE, AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? c.fullName : '',     T3A_NAME, AlignmentType.LEFT,   vMerge),
            dataCell(r.div    ?? '',                  T3A_DIV),
            dataCell(r.seatNo || r.rollNo || '',      T3A_SEAT),
            dataCell(r.name   ?? '',                  T3A_STU,  AlignmentType.LEFT),
            dataCell(r.pct    ?? '',                  T3A_ATT),
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

// ── Table III(B) – Student wise ───────────────────────────────────────────────
// Cols: SN | Sec/Div | Exam Seat No | Name | Overall Att (%) | Course Code | Course Name | Attendance (%)
const T3B_SN      = 500;
const T3B_DIV     = 700;
const T3B_SEAT    = 1500;
const T3B_STUN    = 1600;
const T3B_OVERALL = 800;
const T3B_CODE    = 1300;
const T3B_CNAME   = 1626;
const T3B_ATT     = 1000;
// sum = 500+700+1500+1600+800+1300+1626+1000 = 9026
const T3B_WIDTHS = [T3B_SN, T3B_DIV, T3B_SEAT, T3B_STUN, T3B_OVERALL, T3B_CODE, T3B_CNAME, T3B_ATT];

function buildTableIIIB2(studentList) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',                      T3B_SN),
      headerCell('Sec/Div.',                T3B_DIV),
      headerCell('Exam Seat No',            T3B_SEAT),
      headerCell('Name',                    T3B_STUN),
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
        i === 0 ? dataCell('—', w) : dataCell(i === 1 ? 'No detained students' : '', w)
      ),
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
            dataCell(idx === 0 ? String(sn) : '',             T3B_SN,      AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? (s.div || s.rollNo) : '',    T3B_DIV,     AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? (s.seatNo || s.rollNo) : '', T3B_SEAT,    AlignmentType.CENTER, vMerge),
            dataCell(idx === 0 ? s.name : '',                 T3B_STUN,    AlignmentType.LEFT,   vMerge),
            dataCell(idx === 0 ? s.overall : '',              T3B_OVERALL, AlignmentType.CENTER, vMerge),
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

// ── Main Word generator ───────────────────────────────────────────────────────
async function generateWord2(meta, { tableI, tableII, tableIIIA, tableIIIB }, outputPath) {
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
        buildTableI2(tableI),
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
        buildTableII2(tableII),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(A) ──────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Table III is the list of courses along with the student details and the attendance (%) in the courses in which they are recommended to be detained.', 19),
          ],
        }),
        sectionHeading('Table III(A) – Course wise Detention list'),
        buildTableIIIA2(tableIIIA),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III(B) ──────────────────────────────────────────────────────
        sectionHeading('Table III (B) – Student wise Detention list'),
        buildTableIIIB2(tableIIIB),
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

module.exports = { generateWord2 };
