const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const CompanyMetricLog = sequelize.define('CompanyMetricLog', {
    id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
    },
    period: {
        type: DataTypes.STRING,
        allowNull: false
    },
    value: {
        type: DataTypes.FLOAT,
        allowNull: false
    }
});

module.exports = CompanyMetricLog;
