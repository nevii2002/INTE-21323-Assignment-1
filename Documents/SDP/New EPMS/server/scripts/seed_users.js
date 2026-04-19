const { sequelize, User } = require('../models');
const bcrypt = require('bcryptjs');

async function seedUsers() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        const passwordHash = await bcrypt.hash('password123', 10);

        // Demo Admin
        const [admin, createdAdmin] = await User.findOrCreate({
            where: { email: 'admin@techznap.com' },
            defaults: {
                username: 'Admin User',
                password: passwordHash,
                role: 'Admin',
                jobCategory: 'Full time'
            }
        });
        console.log(createdAdmin ? 'Created Admin: admin@techznap.com' : 'Admin already exists');

        // Demo Employee
        const [emp, createdEmp] = await User.findOrCreate({
            where: { email: 'employee@techznap.com' },
            defaults: {
                username: 'John Employee',
                password: passwordHash,
                role: 'Employee',
                jobCategory: 'Full time',
                position: 'Software Engineer',
                mobileNumber: '+94 77 123 4567'
            }
        });
        console.log(createdEmp ? 'Created Employee: employee@techznap.com' : 'Employee already exists');

    } catch (err) {
        console.error('Error seeding users:', err);
    } finally {
        await sequelize.close();
    }
}

seedUsers();
