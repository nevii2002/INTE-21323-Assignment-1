const { User, Evaluation, Bonus } = require('../models');
const { Sequelize } = require('sequelize');

exports.getDashboardData = async (req, res) => {
    try {
        // 1. Employee Job Category Distribution (Pie Chart)
        const jobCategories = await User.findAll({
            attributes: ['jobCategory', [Sequelize.fn('COUNT', Sequelize.col('id')), 'count']],
            group: ['jobCategory']
        });

        // 2. User Roles Distribution (Pie Chart)
        const roles = await User.findAll({
            attributes: ['role', [Sequelize.fn('COUNT', Sequelize.col('id')), 'count']],
            group: ['role']
        });

        // 3. Evaluation Status Overview (Bar Chart)
        const evaluationStatuses = await Evaluation.findAll({
            attributes: ['status', [Sequelize.fn('COUNT', Sequelize.col('id')), 'count']],
            group: ['status']
        });

        // 4. Bonus Allocation by Reason (Bar Chart)
        const bonusesByReason = await Bonus.findAll({
            attributes: ['reason', [Sequelize.fn('SUM', Sequelize.col('amount')), 'totalAmount']],
            group: ['reason']
        });

        res.json({
            jobCategories,
            roles,
            evaluationStatuses,
            bonusesByReason
        });
    } catch (error) {
        console.error("Error fetching analytics data:", error);
        res.status(500).json({ message: 'Server error' });
    }
};
