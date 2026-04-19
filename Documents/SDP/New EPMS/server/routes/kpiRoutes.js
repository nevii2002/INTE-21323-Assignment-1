const express = require('express');
const router = express.Router();
const kpiController = require('../controllers/kpiController');

// TODO: Add Auth/Admin middleware protection later
router.get('/', kpiController.getAllKPIs);
router.post('/', kpiController.createKPI);
router.put('/:id', kpiController.updateKPI);
router.delete('/:id', kpiController.deleteKPI);

module.exports = router;
