const { sequelize } = require('../models');

async function sync() {
    try {
        await sequelize.authenticate();
        await sequelize.sync({ alter: true });
        console.log('Database synced successfully.');
    } catch (error) {
        console.error('Error syncing database:', error);
    } finally {
        await sequelize.close();
    }
}

sync();
