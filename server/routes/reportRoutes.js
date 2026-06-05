const express  = require('express');
const multer   = require('multer');
const { generateReport, downloadReport } = require('../controllers/reportController');
const { generateReport2 } = require('../controllers/reportController2');
const { generateReport3 } = require('../controllers/reportController3');
const { generateReport4 } = require('../controllers/reportController4');
const { generateReportBMED2 } = require('../controllers/reportControllerBMED2');
const { generateReportBMED3 } = require('../controllers/reportControllerBMED3');

const router = express.Router();

// Use memory storage – no temp files on disk
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter(req, file, cb) {
    const ext = require('path').extname(file.originalname).toLowerCase();
    if (file.fieldname === 'mappingFile') {
      if (ext === '.docx') {
        cb(null, true);
      } else {
        cb(new Error('Only .docx course mapping files are allowed.'));
      }
    } else {
      const allowed = ['.xls', '.xlsx'];
      if (allowed.includes(ext)) {
        cb(null, true);
      } else {
        cb(new Error('Only .xls and .xlsx attendance files are allowed.'));
      }
    }
  },
});

const uploadFields = upload.fields([
  { name: 'files', maxCount: 10 },
  { name: 'mappingFile', maxCount: 1 },
]);

router.post('/generate',           uploadFields, generateReport);
router.post('/y2/generate',        uploadFields, generateReport2);
router.post('/y3/generate',        uploadFields, generateReport3);
router.post('/y4/generate',        uploadFields, generateReport4);
router.post('/bmed2/generate',     uploadFields, generateReportBMED2);
router.post('/bmed3/generate',     uploadFields, generateReportBMED3);
router.get( '/download/:filename', downloadReport);

module.exports = router;

