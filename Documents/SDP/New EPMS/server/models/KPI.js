const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const KPI = sequelize.define('KPI', {
    id: {
        type: DataTypes.INTEGER,
        autoIncrement: true,
        primaryKey: true
    },
    title: {
        type: DataTypes.STRING,
        allowNull: false
    },
    description: {
        type: DataTypes.STRING,
        allowNull: true
    },
    category: {
        type: DataTypes.ENUM('Qualitative', 'Quantitative'),
        allowNull: false,
        comment: 'Qualitative = Evaluation Form, Quantitative = Data Driven'
    },
    type: {
        type: DataTypes.ENUM('EVALUATION', 'BONUS'),
        defaultValue: 'EVALUATION',
        allowNull: false
    },
    unit: {
        type: DataTypes.STRING,
        allowNull: true,
        comment: 'e.g., %, Count, Currency, or Rating (1-5)'
    },
    weight: {
        type: DataTypes.FLOAT,
        defaultValue: 0,
        allowNull: false
    },
    targetValue: {
        type: DataTypes.FLOAT,
        allowNull: true,
        comment: 'Default target for quantitative metrics'
    },
    dataSource: {
        type: DataTypes.STRING,
        allowNull: true,
        comment: 'Source of truth for this metric (e.g. Fiverr, YouTube)'
    },
    role: {
        type: DataTypes.STRING,
        allowNull: true,
        comment: 'Specific role this KPI applies to (optional)'
    }
});

module.exports = KPI;
