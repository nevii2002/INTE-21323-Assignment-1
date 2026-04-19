const KPI = require('../models/KPI');
const { logAction } = require('../utils/auditLogger');

exports.getAllKPIs = async (req, res) => {
    try {
        const kpis = await KPI.findAll();
        res.json(kpis);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.createKPI = async (req, res) => {
    try {
        const { title, description, unit, weight, targetValue, dataSource, role } = req.body;
        const newKPI = await KPI.create({
            title,
            description,
            category: 'Quantitative',
            unit,
            weight,
            targetValue,
            dataSource,
            role
        });

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'CREATE', 'KPI', { kpiId: newKPI.id, title: newKPI.title });
        }

        res.status(201).json(newKPI);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.updateKPI = async (req, res) => {
    try {
        const { id } = req.params;
        const { title, description, unit, weight, targetValue, dataSource, role } = req.body;

        const kpi = await KPI.findByPk(id);
        if (!kpi) return res.status(404).json({ message: 'KPI not found' });

        kpi.title = title;
        kpi.description = description;
        kpi.category = 'Quantitative';
        kpi.unit = unit;
        kpi.weight = weight;
        kpi.targetValue = targetValue;
        kpi.dataSource = dataSource;
        kpi.role = role;

        await kpi.save();

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'UPDATE', 'KPI', { kpiId: kpi.id, title: kpi.title });
        }

        res.json(kpi);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.deleteKPI = async (req, res) => {
    try {
        const { id } = req.params;
        const kpi = await KPI.findByPk(id);
        if (!kpi) return res.status(404).json({ message: 'KPI not found' });

        await kpi.destroy();

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'DELETE', 'KPI', { kpiId: id, title: kpi.title });
        }

        res.json({ message: 'KPI deleted successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
