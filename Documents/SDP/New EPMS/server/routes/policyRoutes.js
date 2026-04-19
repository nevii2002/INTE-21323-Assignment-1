const express = require('express');
const router = express.Router();
const policyController = require('../controllers/policyController');
const { uploadPolicy } = require('../middleware/uploadMiddleware');
const { verifyToken, isAdminOrManager } = require('../middleware/authMiddleware');

// GET all policies (accessible to Employee, Manager, Admin)
router.get('/', verifyToken, policyController.getPolicies);

// POST a new policy (accessible to Manager, Admin)
router.post('/upload', verifyToken, isAdminOrManager, uploadPolicy.single('policyDoc'), policyController.uploadPolicy);

// DELETE a policy (accessible to Manager, Admin)
router.delete('/:id', verifyToken, isAdminOrManager, policyController.deletePolicy);

module.exports = router;
