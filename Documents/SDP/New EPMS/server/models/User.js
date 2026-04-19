const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const User = sequelize.define('User', {
    id: {
        type: DataTypes.INTEGER,
        autoIncrement: true,
        primaryKey: true
    },
    username: {
        type: DataTypes.STRING,
        allowNull: false,
        unique: true
    },
    email: {
        type: DataTypes.STRING,
        allowNull: false,
        unique: true,
        validate: {
            isEmail: true
        }
    },
    password: {
        type: DataTypes.STRING,
        allowNull: true // Changed to allow pre-registration
    },
    status: {
        type: DataTypes.ENUM('Active', 'Pending'),
        defaultValue: 'Pending',
        allowNull: false
    },
    role: {
        type: DataTypes.ENUM('Admin', 'HR', 'Manager', 'Employee', 'CEO'),
        defaultValue: 'Employee',
        allowNull: false
    },
    jobCategory: {
        type: DataTypes.ENUM('Full time', 'Hourly', 'Remote', 'Intern'),
        defaultValue: 'Full time',
        allowNull: true
    },
    jobDescription: {
        type: DataTypes.TEXT,
        allowNull: true
    },
    responsibilities: {
        type: DataTypes.TEXT,
        allowNull: true
    },
    basicSalary: {
        type: DataTypes.FLOAT,
        defaultValue: 0,
        allowNull: true
    },
    bonusAmount: {
        type: DataTypes.FLOAT,
        defaultValue: 0,
        allowNull: true
    },
    mobileNumber: {
        type: DataTypes.STRING,
        allowNull: true
    },
    position: {
        type: DataTypes.STRING,
        allowNull: true
    },
    department: {
        type: DataTypes.STRING,
        allowNull: true
    },
    startDate: {
        type: DataTypes.DATEONLY,
        allowNull: true
    },
    profilePicture: {
        type: DataTypes.STRING,
        allowNull: true
    }
});

module.exports = User;
