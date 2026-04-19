const sequelize = require('../config/database');
const { User } = require('../models');

async function rebuild() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        // 1. Rename existing table
        try {
            await sequelize.query('ALTER TABLE Users RENAME TO Users_old;');
            console.log('Renamed Users to Users_old');
        } catch (e) {
            console.log('Table might not exist or already renamed:', e.message);
        }

        // 2. Sync to create NEW table with correct schema (password nullable)
        await User.sync({ force: true });
        console.log('Created new Users table');

        // 3. Copy data
        // We need to be careful about columns matching.
        // Users_old has: id, username, email, password, role, jobCategory, basicSalary, bonusAmount, mobileNumber, position, createdAt, updatedAt
        // New Users has: ... + status

        // We select columns that exist in old table
        await sequelize.query(`
            INSERT INTO Users (id, username, email, password, role, jobCategory, basicSalary, bonusAmount, mobileNumber, position, createdAt, updatedAt, status)
            SELECT id, username, email, password, role, jobCategory, basicSalary, bonusAmount, mobileNumber, position, createdAt, updatedAt, 'Active'
            FROM Users_old;
        `);
        console.log('Copied data from Users_old to Users (defaulting status to Active)');

        // 4. Drop old table
        await sequelize.query('DROP TABLE Users_old;');
        console.log('Dropped Users_old');

    } catch (error) {
        console.error('Rebuild failed:', error);
    } finally {
        await sequelize.close();
    }
}

rebuild();
