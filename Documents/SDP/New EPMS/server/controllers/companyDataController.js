const { CompanyMetric, CompanyMetricLog } = require('../models');

// Fetch all metrics
exports.getMetrics = async (req, res) => {
    try {
        const metrics = await CompanyMetric.findAll({ order: [['id', 'ASC']] });
        res.json(metrics);
    } catch (error) {
        console.error("Get Metrics Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

// Create a new metric
exports.createMetric = async (req, res) => {
    try {
        const { name, description, unit } = req.body;
        const newMetric = await CompanyMetric.create({ name, description, unit });
        res.status(201).json(newMetric);
    } catch (error) {
        console.error("Create Metric Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

// Fetch logs for a specific period
exports.getLogs = async (req, res) => {
    try {
        const { period } = req.query; // format: YYYY-MM
        if (!period) return res.status(400).json({ message: 'Period is required' });

        const logs = await CompanyMetricLog.findAll({
            where: { period },
            include: [{ model: CompanyMetric, as: 'metric' }]
        });
        res.json(logs);
    } catch (error) {
        console.error("Get Logs Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

// Create or update a log for a specific metric and period
exports.upsertLog = async (req, res) => {
    try {
        const { metricId, period, value } = req.body;
        if (!metricId || !period || value === undefined) {
            return res.status(400).json({ message: 'Missing required fields' });
        }

        // Check if log already exists
        let log = await CompanyMetricLog.findOne({ where: { metricId, period } });
        if (log) {
            log.value = value;
            await log.save();
        } else {
            log = await CompanyMetricLog.create({ metricId, period, value });
        }

        res.status(200).json(log);
    } catch (error) {
        console.error("Upsert Log Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
// Update a metric
exports.updateMetric = async (req, res) => {
    try {
        const { id } = req.params;
        const { name, description, unit } = req.body;
        const metric = await CompanyMetric.findByPk(id);
        if (!metric) return res.status(404).json({ message: 'Metric not found' });
        await metric.update({ name, description, unit });
        res.json(metric);
    } catch (error) {
        console.error('Update Metric Error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

// Delete a metric and all its logs
exports.deleteMetric = async (req, res) => {
    try {
        const { id } = req.params;
        await CompanyMetricLog.destroy({ where: { metricId: id } });
        await CompanyMetric.destroy({ where: { id } });
        res.json({ message: 'Metric deleted successfully' });
    } catch (error) {
        console.error('Delete Metric Error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
