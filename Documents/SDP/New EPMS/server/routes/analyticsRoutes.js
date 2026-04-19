const express = require('express');
const router = express.Router();
const analyticsController = require('../controllers/analyticsController');
const { verifyToken, isAdminOrManager } = require('../middleware/authMiddleware');

router.get('/dashboard', verifyToken, isAdminOrManager, analyticsController.getDashboardData);

module.exports = router;
