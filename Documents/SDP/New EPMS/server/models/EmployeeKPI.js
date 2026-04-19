const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const EmployeeKPI = sequelize.define('EmployeeKPI', {
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
    customWeight: {
        type: DataTypes.FLOAT,
        allowNull: false,
        comment: 'Custom weight assigned to this employee for this KPI'
    },
    customBonus: {
        type: DataTypes.FLOAT,
        allowNull: true,
        defaultValue: 0,
        comment: 'Custom bonus amount allocated for this KPI'
    }
});

module.exports = EmployeeKPI;
