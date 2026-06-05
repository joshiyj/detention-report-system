/**
 * wordGeneratorY3.js
 * Generates the 3rd Year (6th Semester) ECS Detention List .docx.
 * Matches the reference document: VI SEM DETENTION UPDATED 22_4_2026.docx
 *
 * Table I:   Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
 * Table II:  Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
 * Table III: Examination Seat No. | Name of the students | Aggregate Attendance | Course Code | Attendance (%)
 *            (Student-wise: multiple rows per student for multiple detained courses,
 *             with course code and name listed together in the course column)
 *
 * Sorting:
 *   - Tables I & II: sorted by Seat No (caller's responsibility)
 *   - Table III:     sorted by Seat No (caller's responsibility)
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
function headerCell(text, width, colSpan = 1, rowSpan = 1) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
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
// Cols: SN | Student ID/Roll No. | Name of the students | Aggregate Attendance (%)
const T12_SN     = 500;
const T12_SEAT   = 2000;
const T12_NAME   = 4800;
const T12_ATT    = 1726;
// sum = 9026
const T12_WIDTHS = [T12_SN, T12_SEAT, T12_NAME, T12_ATT];

function buildTableI_II(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      headerCell('SN',                      T12_SN),
      headerCell('Student ID / Seat No.',   T12_SEAT),
      headerCell('Name of the Students',    T12_NAME),
      headerCell('Aggregate Attendance (%)', T12_ATT),
    ],
  });

  const dataRows = rows.length > 0
    ? rows.map((r, idx) => new TableRow({
        children: [
          dataCell(String(idx + 1),              T12_SN),
          dataCell(r.seatNo || r.rollNo || '',   T12_SEAT),
          dataCell(r.name ?? '',                 T12_NAME, AlignmentType.LEFT),
          dataCell(r.overall ?? '',              T12_ATT),
        ],
      }))
    : [new TableRow({ children: [
        dataCell('—', T12_SN),
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

// ── Table III ─────────────────────────────────────────────────────────────────
// Student-wise detention list with all detained courses listed per student
// Cols: Examination Seat No./ Student Unique No | Name of the students | Aggregate Attendance | Course Code (Course Name + Code) | Attendance (%)
const T3_SEAT    = 2500;
const T3_NAME    = 2500;
const T3_OVERALL = 1200;
const T3_CODE    = 1826;
const T3_ATT     = 1000;
// sum = 2500+2500+1200+1826+1000 = 9026
const T3_WIDTHS  = [T3_SEAT, T3_NAME, T3_OVERALL, T3_CODE, T3_ATT];

function buildTableIII(studentList) {
  const hRow1 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Examination Seat No./\rStudent Unique No',         T3_SEAT, 1, 2),
      headerCell('Name of the students',                            T3_NAME, 1, 2),
      headerCell('Aggregate Attendance',                            T3_OVERALL, 1, 2),
      headerCell('Courses in which the student is recommended for detention', T3_CODE + T3_ATT, 2, 1),
    ],
  });

  const hRow2 = new TableRow({
    tableHeader: true,
    children: [
      headerCell('Course Code',                  T3_CODE),
      headerCell('Attendance (%)',               T3_ATT),
    ],
  });

  const dataRows = [];

  if (studentList.length === 0) {
    dataRows.push(new TableRow({
      children: [
        dataCell('—', T3_SEAT),
        dataCell('No detained students', T3_NAME, AlignmentType.LEFT),
        dataCell('—', T3_OVERALL),
        dataCell('—', T3_CODE),
        dataCell('—', T3_ATT),
      ],
    }));
  } else {
    for (const s of studentList) {
      // Sort courses by SN within this student
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
            dataCell(c.fullName ? `${c.fullName} ${c.code}` : (c.code ?? ''), T3_CODE, AlignmentType.CENTER),
            dataCell(c.pct ?? '',                             T3_ATT,     AlignmentType.CENTER),
          ],
        }));
      });
    }
  }

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: T3_WIDTHS,
    rows: [hRow1, hRow2, ...dataRows],
  });
}

// ── Main generator ────────────────────────────────────────────────────────────
async function generateWordY3(meta, { tableI, tableII, tableIII }, outputPath) {
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

        // ── University heading ─────────────────────────────────────────────────
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
            normal(`Table I is the list of students of programme ${programme} ${semester} having aggregate attendance less than 75%.`, 19),
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
        buildTableI_II(tableI),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table II ──────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal(`However Table II, is the list of students of programme ${programme} ${semester} having aggregate attendance between 60% and 75% and have applied for condonation of attendance. The application forms and medical certificates / other relevant documents have been attached herewith. After due consideration and as per the regulation R 17, it is recommended that these students may be permitted to appear in all courses in the `, 19),
            bold(examName, 19),
            normal(' examinations. Since the absence was due to circumstances beyond the control of the students.', 19),
          ],
        }),
        sectionHeading('Table II'),
        buildTableI_II(tableII),
        new Paragraph({ spacing: { after: 160 }, children: [] }),

        // ── Table III ─────────────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal(`Table III is the list of students of programme ${programme} ${semester} along with the course code and the attendance in the courses in which they are recommended to be detained in `, 19),
            bold(examName, 19),
            normal(' examination as per regulation R 17.', 19),
          ],
        }),
        sectionHeading('Table III – Student wise Detention list'),
        buildTableIII(tableIII),
        new Paragraph({ spacing: { after: 240 }, children: [] }),

        // ── Signature footer ──────────────────────────────────────────────────
        new Paragraph({
          spacing: { after: 60 },
          children: [
            normal('Academic Coordinator, EN', 20),
            new TextRun({ text: '\t\t\t\t\t\t\t\t', size: 20 }),
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

module.exports = { generateWordY3 };
