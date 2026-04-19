const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');
const User = require('./User');

const Evaluation = sequelize.define('Evaluation', {
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
    evaluatorId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: User,
            key: 'id'
        }
    },
    type: {
        type: DataTypes.ENUM('Self', 'Manager', 'Peer'),
        allowNull: false
    },
    period: {
        type: DataTypes.STRING, // e.g., '2024-01'
        allowNull: false
    },
    status: {
        type: DataTypes.ENUM('Draft', 'Submitted'),
        defaultValue: 'Draft'
    },
    comments: {
        type: DataTypes.TEXT,
        allowNull: true
    }
}, {
    timestamps: true
});

module.exports = Evaluation;
