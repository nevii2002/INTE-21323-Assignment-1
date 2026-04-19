const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');
const User = require('./User');

const Bonus = sequelize.define('Bonus', {
    id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true,
    },
    employeeId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: User,
            key: 'id'
        }
    },
    amount: {
        type: DataTypes.FLOAT,
        allowNull: false,
        validate: {
            min: 0
        }
    },
    reason: {
        type: DataTypes.STRING,
        allowNull: false
    },
    dateGiven: {
        type: DataTypes.DATEONLY,
        defaultValue: DataTypes.NOW,
        allowNull: false
    },
    status: {
        type: DataTypes.ENUM('Pending', 'Paid'),
        defaultValue: 'Paid' // Assuming usually recorded when decided
    }
}, {
    timestamps: true
});

module.exports = Bonus;
