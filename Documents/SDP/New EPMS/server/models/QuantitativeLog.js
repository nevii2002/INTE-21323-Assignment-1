const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const QuantitativeLog = sequelize.define('QuantitativeLog', {
    id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true
    },
    employeeId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: 'Users',
            key: 'id'
        }
    },
    kpiId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: 'KPIs',
            key: 'id'
        }
    },
    period: {
        type: DataTypes.STRING, // e.g., '2024-03'
        allowNull: false
    },
    actualValue: {
        type: DataTypes.FLOAT,
        allowNull: false,
        defaultValue: 0
    }
}, {
    timestamps: true,
    indexes: [
        {
            unique: true,
            fields: ['employeeId', 'kpiId', 'period']
        }
    ]
});

module.exports = QuantitativeLog;
