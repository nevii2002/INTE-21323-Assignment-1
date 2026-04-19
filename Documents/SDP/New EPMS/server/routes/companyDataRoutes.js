const express = require('express');
const router = express.Router();
const companyDataController = require('../controllers/companyDataController');
const { verifyToken, isAdminOrManager } = require('../middleware/authMiddleware');

router.get('/metrics', verifyToken, companyDataController.getMetrics);
router.post('/metrics', verifyToken, isAdminOrManager, companyDataController.createMetric);
router.put('/metrics/:id', verifyToken, isAdminOrManager, companyDataController.updateMetric);
router.delete('/metrics/:id', verifyToken, isAdminOrManager, companyDataController.deleteMetric);

router.get('/logs', verifyToken, companyDataController.getLogs);
router.post('/logs', verifyToken, isAdminOrManager, companyDataController.upsertLog);

module.exports = router;
