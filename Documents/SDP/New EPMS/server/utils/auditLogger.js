const { AuditLog } = require('../models');

/**
 * Helper function to log audit actions.
 * @param {number} userId - The ID of the user performing the action.
 * @param {string} action - 'CREATE', 'UPDATE', 'DELETE', or other verbs.
 * @param {string} resource - The name of the entity being acted upon (e.g., 'KPI', 'Staff', 'Evaluation').
 * @param {object} details - Any relevant context or json summary of what was changed.
 */
const logAction = async (userId, action, resource, details) => {
    try {
        await AuditLog.create({
            userId,
            action,
            resource,
            details: details ? JSON.stringify(details) : null
        });
    } catch (error) {
        console.error('Failed to write audit log:', error);
    }
};

module.exports = {
    logAction
};
