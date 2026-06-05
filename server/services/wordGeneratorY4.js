/**
* wordGeneratorY4.js
* Generates the 4th Year (8th Semester) ECS Detention List .docx.
* Matches the reference document: VIII Sem A and B Detention.docx
*
* Table I:   Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
*            (No SN column)
* Table II:  Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
*            (No SN column)
* Table III: Examination Seat No./ Student Unique No | Name of the students | Aggregate Attendance | Course Code | Attendance (%)
*            (No SN, No Course Name columns. Grouped student-wise with vertical merges)
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
function headerCell(text, width, colSpan = 1) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    columnSpan: colSpan,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [bold(text, 18)],
    })],
  });
}

function dataCell(text, width, alignment = AlignmentType.CENTER, verticalMerge = undefined, colSpan = 1) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    verticalMerge,
    columnSpan: colSpan,
    children: [new Paragraph({
      alignment,
      children: [normal(text ?? '', 18)],
    })],
  });
}

// ── Table I & II ──────────────────────────────────────────────────────────────
// Cols: Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
// Widths sum = 9026
const T12_SEAT   = 3000;
const T12_NAME   = 4300;
const T12_ATT    = 1726;
const T12_WIDTHS = [T12_SEAT, T12_NAME, T12_ATT];

function buildTableI_II(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Student ID/Roll No.',   T12_SEAT),
      headerCell('Name of the students',    T12_NAME),
      headerCell('Aggregate Attendance (%)', T12_ATT),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map(r => new TableRow({
        children: [
          dataCell(r.seatNo || r.rollNo || '',   T12_SEAT),
          dataCell(r.name ?? '',                 T12_NAME, AlignmentType.LEFT),
          dataCell(r.overall ?? '',              T12_ATT),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('', T12_SEAT),
        dataCell('Nil', T12_NAME, AlignmentType.CENTER),
        dataCell('', T12_ATT),
      ] })];

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T12_WIDTHS,
    rows: [headerRow, ...dataRows],
  });
}

// ── Table III ─────────────────────────────────────────────────────────────────
// Grouped student-wise
// Cols: Examination Seat No./ Student Unique No | Name of the students | Aggregate Attendance | Course Code | Course Name | Attendance (%)
// Widths sum = 2000 + 2600 + 1000 + 1000 + 1426 + 1000 = 9026
const T3_SEAT    = 2000;
const T3_NAME    = 2600;
const T3_OVERALL = 1000;
const T3_CODE    = 1000;
const T3_CNAME   = 1426;
const T3_ATT     = 1000;
const T3_WIDTHS  = [T3_SEAT, T3_NAME, T3_OVERALL, T3_CODE, T3_CNAME, T3_ATT];

function buildTableIII(studentList) {
  // Row 1 Header (with merged gridSpan cell)
  const headerRow1 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Examination Seat No./ Student Unique No', T3_SEAT),
      headerCell('Name of the students', T3_NAME),
      headerCell('Aggregate Attendance', T3_OVERALL),
      headerCell('Courses in which the student is recommended for detention', T3_CODE + T3_CNAME + T3_ATT, 3),
    ],
  });

  // Row 2 Header
  const headerRow2 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('', T3_SEAT),
      headerCell('', T3_NAME),
      headerCell('', T3_OVERALL),
      headerCell('Course Code', T3_CODE),
      headerCell('Course Name', T3_CNAME),
      headerCell('Attendance (%)', T3_ATT),
    ],
  });

  const dataRows = [];

  if (studentList.length === 0) {
    dataRows.push(new TableRow({
      children: [
        dataCell('', T3_SEAT),
        dataCell('NIl', T3_NAME, AlignmentType.CENTER),
        dataCell('', T3_OVERALL),
        dataCell('', T3_CODE),
        dataCell('', T3_CNAME),
        dataCell('', T3_ATT),
      ],
    }));
  } else {
    for (const s of studentList) {
      const sortedCourses = [...s.courses].sort((a, b) => a.sn - b.sn);

      sortedCourses.forEach((c, idx) => {
        const isFirst  = idx === 0;
        const rowCount = sortedCourses.length;
        const vMerge   = rowCount > 1
          ? (isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE)
          : undefined;

        dataRows.push(new TableRow({
          children: [
            dataCell(isFirst ? (s.seatNo || s.rollNo) : '',  T3_SEAT,    AlignmentType.CENTER, vMerge),
            dataCell(isFirst ? s.name : '',                   T3_NAME,    AlignmentType.LEFT,   vMerge),
            dataCell(isFirst ? s.overall : '',                T3_OVERALL, AlignmentType.CENTER, vMerge),
            dataCell(c.code ?? '',                            T3_CODE,    AlignmentType.CENTER),
            dataCell(c.fullName ?? '',                        T3_CNAME,   AlignmentType.LEFT),
            dataCell(c.pct ?? '',                             T3_ATT,     AlignmentType.CENTER),
          ],
        }));
      });
    }
  }

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T3_WIDTHS,
    rows: [headerRow1, headerRow2, ...dataRows],
  });
}

// ── Main generator ────────────────────────────────────────────────────────────
async function generateWordY4(meta, { tableI, tableII, tableIII }, outputPath) {
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
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
        },
      },
      children: [
        // ── Heading matching college reference doc ─────────────────────────────
        centerPara([bold('SHRI RAMDEOBABA COLLEGE OF ENGINEERING AND MANAGEMENT, NAGPUR', 22)], 40),
        centerPara([bold(schoolName || 'Department of Electronics Engineering', 22)], 80),
        
        // Date block right aligned
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 120 },
          children: [normal(`Date: ${date}`, 20)],
        }),

        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 160 },
          children: [bold('Submitted to Principal for Approval through Dean Academics', 20)],
        }),

        // ── Table I description ───────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 80 },
          children: [
            normal(`Table I is the list of students of programme ${programme} ${semester}  having aggregate attendance less than 75%.`, 19),
          ],
        }),
        new Paragraph({
          spacing: { after: 100 },
          children: [
            normal('As per the ordinances/ regulations of the college these students should be detained in the ', 19),
            bold(examName, 19),
            normal(' examination in all courses.', 19),
          ],
        }),
        
        sectionHeading('Table I'),
        buildTableI_II(tableI),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table II description ──────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 80 },
          children: [
            normal(`However Table II, is the list of students of programme ${programme} ${semester}  having aggregate attendance between 60% and 75% and has applied for condonation of attendance. The application forms and medical certificates / other relevant documents have been attached herewith. After due consideration and as per the regulation R 17, it is recommended that these students may be permitted to appear in all courses in the `, 19),
            bold(examName, 19),
            normal(' examinations. Since the absence was due to circumstances beyond the control of the students.', 19),
          ],
        }),

        sectionHeading('Table II'),
        buildTableI_II(tableII),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III description ─────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 80 },
          children: [
            normal(`Table III is the list of students of programme ${programme} ${semester}  along with the course code and the attendance in the courses in which they are recommended to be detained in `, 19),
            bold(examName, 19),
            normal(' examination as per regulation R 17.', 19),
          ],
        }),

        sectionHeading('Table III'),
        buildTableIII(tableIII),
        new Paragraph({ spacing: { after: 280 }, children: [] }),

        // ── Signature footer ──────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Class Incharge', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t', size: 20 }),
            normal('Dr. Nitin Narkhede', 20),
          ],
          tabStops: [{ type: 'right', position: 9026 }],
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [
            normal('', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t', size: 20 }),
            normal('HoD', 20),
          ],
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [
            normal('', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t', size: 20 }),
            normal('Department of Electronics Engineering', 20),
          ],
        }),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  return outputPath;
}

module.exports = { generateWordY4 };
