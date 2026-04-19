const express = require('express');
const router = express.Router();
const quantitativeController = require('../controllers/quantitativeController');
const { verifyToken } = require('../middleware/authMiddleware');

// Route to get logs for an employee (with optional period query)
router.get('/:employeeId', verifyToken, quantitativeController.getLogs);

// Route to create or update a quantitative log
router.post('/', verifyToken, quantitativeController.upsertLog);

module.exports = router;
