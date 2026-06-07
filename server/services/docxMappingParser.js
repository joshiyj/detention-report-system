/**
 * docxMappingParser.js
 * Utility to parse an uploaded course mapping document (.docx) on the fly.
 * Uses mammoth to convert to clean HTML and robust regex to extract table rows.
 */

const mammoth = require('mammoth');

/**
 * Strips HTML tags and decodes basic entities.
 */
function cleanText(str) {
  if (!str) return '';
  return str
    .replace(/<[^>]*?>/g, ' ') // replace tags with spaces
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Parses the converted HTML tables to build COURSE_MAP and SHORTFORM_ORDER.
 * Matches SN, Course Type, Code, Course, and Abbreviation columns.
 *
 * @param {string} html - HTML returned by mammoth
 * @returns {{ courseMap: Object, shortformOrder: string[] } | null}
 */
function parseHtmlMappingTable(html) {
  const courseMap = {};
  const shortformOrder = [];

  // Find the first table
  const tableStart = html.indexOf('<table');
  if (tableStart === -1) return null;

  let depth = 0;
  let inRow = false;
  let inCell = false;
  
  let currentRawRow = [];
  const rows = [];
  let currentCellContent = '';
  
  let pos = tableStart;
  while (pos < html.length) {
    if (html.substring(pos, pos + 6) === '<table') {
      depth++;
      const closeIdx = html.indexOf('>', pos);
      if (closeIdx !== -1) {
        pos = closeIdx + 1;
      } else {
        pos += 6;
      }
      continue;
    }
    if (html.substring(pos, pos + 7) === '</table') {
      depth--;
      const closeIdx = html.indexOf('>', pos);
      if (closeIdx !== -1) {
        pos = closeIdx + 1;
      } else {
        pos += 7;
      }
      if (depth === 0) {
        break; // End of outermost table
      }
      continue;
    }
    
    if (depth === 1) {
      if (html.substring(pos, pos + 3) === '<tr') {
        inRow = true;
        currentRawRow = [];
        pos += 3;
        continue;
      }
      if (html.substring(pos, pos + 4) === '</tr') {
        inRow = false;
        if (currentRawRow.length > 0) {
          rows.push(currentRawRow);
        }
        pos += 4;
        continue;
      }
      
      if (inRow) {
        if (html.substring(pos, pos + 3) === '<td') {
          inCell = true;
          currentCellContent = '';
          const closeTag = html.indexOf('>', pos);
          if (closeTag !== -1) {
            pos = closeTag + 1;
          } else {
            pos += 3;
          }
          continue;
        }
        if (html.substring(pos, pos + 4) === '</td') {
          inCell = false;
          currentRawRow.push(currentCellContent);
          pos += 4;
          continue;
        }
      }
    }
    
    if (inCell) {
      currentCellContent += html[pos];
    }
    pos++;
  }

  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const cells = rows[rIdx];
    if (cells.length < 5) continue; // Must have SN, Course Type, Code, Course, Abbreviation

    const sn = cleanText(cells[0]);
    const courseType = cleanText(cells[1]);
    const code = cleanText(cells[2]);
    const course = cleanText(cells[3]);
    
    // Fifth cell contains abbreviation(s) potentially separated by <p> tags
    const cell5Html = cells[4];
    const pRegex = /<p[\s\S]*?>([\s\S]*?)<\/p>/g;
    const rawAbbrs = [];
    let pMatch;
    while ((pMatch = pRegex.exec(cell5Html)) !== null) {
      const txt = cleanText(pMatch[1]);
      if (txt) rawAbbrs.push(txt);
    }
    
    // Fallback if no paragraphs are found (just clean the cell)
    if (rawAbbrs.length === 0) {
      const txt = cleanText(cell5Html);
      if (txt) rawAbbrs.push(txt);
    }

    // Skip header/empty/total rows
    if (
      !code ||
      code === 'Code' ||
      sn === 'SN' ||
      courseType === 'Course Type' ||
      course.toUpperCase().includes('TOTAL')
    ) {
      continue;
    }

    // Process abbreviation(s)
    for (const abbr of rawAbbrs) {
      // High-precision type classification
      let type = 'Theory';
      const abbrLower = abbr.toLowerCase();
      const typeLower = courseType.toLowerCase();
      const codeLower = code.toLowerCase();
      const courseLower = course.toLowerCase();

      if (
        abbrLower.includes('(p)') || 
        abbrLower.includes('lab') || 
        abbrLower.includes('practical') ||
        typeLower.includes('lab') || 
        typeLower.includes('practical') || 
        typeLower.includes('pr') ||
        typeLower.includes('pro') ||
        codeLower.includes('pr') ||
        courseLower.includes(' lab') ||
        courseLower.includes('practical') ||
        courseLower.includes('project')
      ) {
        type = 'Practical';
      } else if (
        abbrLower.includes('(l)') || 
        abbrLower.includes('theory') || 
        typeLower.includes('th') ||
        codeLower.includes('th')
      ) {
        type = 'Theory';
      }

      // Sn sorting integer (fallback to 999)
      const snInt = parseInt(sn, 10) || 999;

      courseMap[abbr] = {
        sn: snInt,
        code,
        fullName: course,
        type
      };

      const key = `${code}::${type}`;
      if (!shortformOrder.includes(key)) {
        shortformOrder.push(key);
      }
    }
  }

  return { courseMap, shortformOrder };
}

/**
 * Parses a docx file buffer and returns unified courseMap and shortformOrder.
 *
 * @param {Buffer} buffer - Uploaded docx buffer
 * @returns {Promise<{ courseMap: Object, shortformOrder: string[] } | null>}
 */
async function parseDocxMapping(buffer) {
  try {
    const result = await mammoth.convertToHtml({ buffer });
    return parseHtmlMappingTable(result.value);
  } catch (e) {
    console.error('[parseDocxMapping] Error parsing docx:', e);
    return null;
  }
}

module.exports = { parseDocxMapping };
