const { QuantitativeLog, KPI, User } = require('../models');

exports.upsertLog = async (req, res) => {
    try {
        const { employeeId, kpiId, period, actualValue } = req.body;

        if (!employeeId || !kpiId || !period || actualValue === undefined) {
            return res.status(400).json({ message: 'Missing required fields' });
        }

        // Check if log exists for this period
        let log = await QuantitativeLog.findOne({
            where: { employeeId, kpiId, period }
        });

        if (log) {
            log.actualValue = actualValue;
            await log.save();
        } else {
            log = await QuantitativeLog.create({
                employeeId,
                kpiId,
                period,
                actualValue
            });
        }

        res.status(200).json({ message: 'Metric logged successfully', log });
    } catch (error) {
        console.error('Upsert Log Error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getLogs = async (req, res) => {
    try {
        const { employeeId } = req.params;
        const { period } = req.query; // optional filter

        let whereClause = { employeeId };
        if (period) {
            whereClause.period = period;
        }

        const logs = await QuantitativeLog.findAll({
            where: whereClause,
            include: [{ model: KPI, attributes: ['id', 'title', 'targetValue', 'unit'] }]
        });

        res.json(logs);
    } catch (error) {
        console.error('Get Logs Error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
