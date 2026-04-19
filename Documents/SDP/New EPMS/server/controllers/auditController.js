const { AuditLog, User } = require('../models');

/**
 * Fetch all audit logs (Admin only)
 */
exports.getAuditLogs = async (req, res) => {
    try {
        const filters = {};

        // Optional filtering by query params
        if (req.query.action) filters.action = req.query.action;
        if (req.query.resource) filters.resource = req.query.resource;

        const logs = await AuditLog.findAll({
            where: filters,
            include: [
                {
                    model: User,
                    as: 'user',
                    attributes: ['id', 'username', 'email', 'role']
                }
            ],
            order: [['createdAt', 'DESC']]
        });

        res.status(200).json(logs);
    } catch (error) {
        console.error('Error fetching audit logs:', error);
        res.status(500).json({ message: 'Failed to retrieve audit logs.', error });
    }
};
