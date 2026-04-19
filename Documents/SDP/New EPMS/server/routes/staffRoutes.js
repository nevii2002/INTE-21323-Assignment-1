const express = require('express');
const router = express.Router();
const staffController = require('../controllers/staffController');

const { verifyToken, isAdmin, isAdminOrManager } = require('../middleware/authMiddleware');

// Apply authentication to all routes
router.use(verifyToken);

// Admin only routes for management (Except GET / which handles its own role scoping)
router.get('/', staffController.getAllStaff);
router.get('/:id', isAdminOrManager, staffController.getStaffById);
router.post('/', isAdmin, staffController.createStaff);
router.put('/:id', isAdmin, staffController.updateStaff);
router.delete('/:id', isAdmin, staffController.deleteStaff);

// KPI Assignments
router.get('/my-kpis', verifyToken, staffController.getMyKPIs);
router.get('/:id/kpis', isAdminOrManager, staffController.getEmployeeKPIs);
router.post('/:id/kpis', isAdmin, staffController.assignKPIs);

module.exports = router;
