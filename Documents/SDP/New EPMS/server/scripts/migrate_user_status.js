const sequelize = require('../config/database');
const { User } = require('../models');

async function migrate() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        // Alter table to add status and change password constraint
        await sequelize.sync({ alter: true });

        console.log('User table altered successfully.');

        // Update all existing users to be 'Active'
        await User.update({ status: 'Active' }, { where: {} });
        console.log('Existing users marked as Active.');

    } catch (error) {
        console.error('Migration failed:', error);
    } finally {
        await sequelize.close();
    }
}

migrate();
