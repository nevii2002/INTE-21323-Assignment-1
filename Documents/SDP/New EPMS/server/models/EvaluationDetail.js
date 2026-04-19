const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');
const Evaluation = require('./Evaluation');
const KPI = require('./KPI');

const EvaluationDetail = sequelize.define('EvaluationDetail', {
    id: {
        type: DataTypes.INTEGER,
        primaryKey: true,
        autoIncrement: true,
    },
    evaluationId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: Evaluation,
            key: 'id'
        }
    },
    kpiId: {
        type: DataTypes.INTEGER,
        allowNull: false,
        references: {
            model: KPI,
            key: 'id'
        }
    },
    rating: {
        type: DataTypes.INTEGER,
        allowNull: false,
        validate: {
            min: 1,
            max: 5
        }
    },
    comment: {
        type: DataTypes.TEXT,
        allowNull: true
    }
}, {
    timestamps: true
});

module.exports = EvaluationDetail;
