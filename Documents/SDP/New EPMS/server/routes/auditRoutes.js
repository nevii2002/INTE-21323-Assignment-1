const express = require('express');
const router = express.Router();
const auditController = require('../controllers/auditController');
const { verifyToken, isAdmin } = require('../middleware/authMiddleware');

// Route is protected: must be logged in AND must be An Admin
router.get('/', verifyToken, isAdmin, auditController.getAuditLogs);

module.exports = router;
