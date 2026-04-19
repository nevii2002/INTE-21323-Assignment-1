const sequelize = require('../config/database');
const User = require('./User');
const KPI = require('./KPI');
const Evaluation = require('./Evaluation');
const EvaluationDetail = require('./EvaluationDetail');
const EmployeeKPI = require('./EmployeeKPI');
const Bonus = require('./Bonus');
const Policy = require('./Policy');
const QuantitativeLog = require('./QuantitativeLog');
const AuditLog = require('./AuditLog');
const CompanyMetric = require('./CompanyMetric');
const CompanyMetricLog = require('./CompanyMetricLog');

// Associations
User.hasMany(Evaluation, { foreignKey: 'employeeId', as: 'evaluationsReceived' });
User.hasMany(Evaluation, { foreignKey: 'evaluatorId', as: 'evaluationsGiven' });
Evaluation.belongsTo(User, { foreignKey: 'employeeId', as: 'employee' });
Evaluation.belongsTo(User, { foreignKey: 'evaluatorId', as: 'evaluator' });

Evaluation.hasMany(EvaluationDetail, { foreignKey: 'evaluationId', as: 'details' });
EvaluationDetail.belongsTo(Evaluation, { foreignKey: 'evaluationId' });

EvaluationDetail.belongsTo(KPI, { foreignKey: 'kpiId', as: 'kpi' });

// Employee Custom KPIs
User.belongsToMany(KPI, { through: EmployeeKPI, foreignKey: 'employeeId', as: 'assignedKPIs' });
KPI.belongsToMany(User, { through: EmployeeKPI, foreignKey: 'kpiId', as: 'assignedEmployees' });

// Quantitative Logs
User.hasMany(QuantitativeLog, { foreignKey: 'employeeId', as: 'quantitativeLogs' });
QuantitativeLog.belongsTo(User, { foreignKey: 'employeeId' });

KPI.hasMany(QuantitativeLog, { foreignKey: 'kpiId', as: 'quantitativeLogs' });
QuantitativeLog.belongsTo(KPI, { foreignKey: 'kpiId' });

// Bonus Associations
User.hasMany(Bonus, { foreignKey: 'employeeId', as: 'bonuses' });
Bonus.belongsTo(User, { foreignKey: 'employeeId', as: 'employee' });

// Policy Associations
User.hasMany(Policy, { foreignKey: 'uploadedBy', as: 'policies' });
Policy.belongsTo(User, { foreignKey: 'uploadedBy', as: 'uploader' });

// AuditLog Associations
User.hasMany(AuditLog, { foreignKey: 'userId', as: 'auditLogs' });
AuditLog.belongsTo(User, { foreignKey: 'userId', as: 'user' });

// Company Metric Associations
CompanyMetric.hasMany(CompanyMetricLog, { foreignKey: 'metricId', as: 'logs' });
CompanyMetricLog.belongsTo(CompanyMetric, { foreignKey: 'metricId', as: 'metric' });

module.exports = {
    sequelize,
    User,
    KPI,
    Evaluation,
    EvaluationDetail,
    EmployeeKPI,
    Bonus,
    Policy,
    QuantitativeLog,
    AuditLog,
    CompanyMetric,
    CompanyMetricLog
};
