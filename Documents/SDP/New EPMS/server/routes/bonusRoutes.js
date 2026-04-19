const express = require('express');
const router = express.Router();
const bonusController = require('../controllers/bonusController');
const { verifyToken, isAdmin, isAdminOrManager } = require('../middleware/authMiddleware');

// Admin or Manager can award bonuses
router.post('/', verifyToken, isAdminOrManager, bonusController.createBonus);

// Get all bonuses (Admin/HR view)
router.get('/', verifyToken, isAdmin, bonusController.getAllBonuses);

// Get bonuses for a specific employee
// Users can see their own, Managers/Admins can see others
router.get('/employee/:employeeId', verifyToken, bonusController.getBonusesByEmployee);

// Delete a bonus (Admin only)
router.delete('/:id', verifyToken, isAdmin, bonusController.deleteBonus);

module.exports = router;
