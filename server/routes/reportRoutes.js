const express  = require('express');
const multer   = require('multer');
const { generateReport, downloadReport } = require('../controllers/reportController');

const router = express.Router();

// Use memory storage – no temp files on disk
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter(req, file, cb) {
    const allowed = ['.xls', '.xlsx'];
    const ext = require('path').extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only .xls and .xlsx files are allowed.'));
    }
  },
});

router.post('/generate',           upload.array('files', 10), generateReport);
router.get( '/download/:filename', downloadReport);

module.exports = router;
