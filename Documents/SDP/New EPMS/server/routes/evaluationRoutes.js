const express = require('express');
const router = express.Router();
const evaluationController = require('../controllers/evaluationController');
const { verifyToken, isAdminOrManager } = require('../middleware/authMiddleware');

router.use(verifyToken);

router.post('/', isAdminOrManager, evaluationController.createEvaluation);
router.get('/', evaluationController.getEvaluations); // Employees can see their own, Managers can see all
router.get('/stats', evaluationController.getDashboardStats);
router.get('/:id', evaluationController.getEvaluationById);

module.exports = router;
