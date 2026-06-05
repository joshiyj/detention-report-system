const fs = require('fs');
const path = require('path');

const MAPPINGS_FILE_PATH = path.join(__dirname, 'saved_mappings.json');

/**
 * Retrieves a saved mapping for the given academic year/programme key.
 * @param {string} yearKey - Key identifying the report category (e.g. 'y1', 'y3')
 * @returns {Object|null} - Saved mapping containing courseMap and shortformOrder, or null if not found
 */
function getMapping(yearKey) {
  if (!fs.existsSync(MAPPINGS_FILE_PATH)) {
    return null;
  }
  try {
    const data = fs.readFileSync(MAPPINGS_FILE_PATH, 'utf8');
    const allMappings = JSON.parse(data);
    return allMappings[yearKey] || null;
  } catch (err) {
    console.error(`[mappingPersistence] Error reading saved mappings for ${yearKey}:`, err);
    return null;
  }
}

/**
 * Saves a parsed mapping for the given academic year/programme key.
 * @param {string} yearKey - Key identifying the report category (e.g. 'y1', 'y3')
 * @param {Object} mapping - Mapped object containing courseMap and shortformOrder
 */
function saveMapping(yearKey, mapping) {
  if (!mapping) return;
  let allMappings = {};
  if (fs.existsSync(MAPPINGS_FILE_PATH)) {
    try {
      const data = fs.readFileSync(MAPPINGS_FILE_PATH, 'utf8');
      allMappings = JSON.parse(data);
    } catch (err) {
      console.error(`[mappingPersistence] Error parsing saved mappings, initializing empty:`, err);
    }
  }
  allMappings[yearKey] = mapping;
  try {
    fs.writeFileSync(MAPPINGS_FILE_PATH, JSON.stringify(allMappings, null, 2), 'utf8');
    console.log(`[mappingPersistence] Successfully saved mapping for key: ${yearKey}`);
  } catch (err) {
    console.error(`[mappingPersistence] Error writing saved mappings for ${yearKey}:`, err);
  }
}

module.exports = {
  getMapping,
  saveMapping
};
