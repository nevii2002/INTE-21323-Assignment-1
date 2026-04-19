const { sequelize, User } = require('./models');
const bcrypt = require('bcryptjs');

async function seedUsers() {
    try {
        await sequelize.authenticate();
        await sequelize.sync();

        const hashedPassword = await bcrypt.hash('password123', 10);

        await User.create({
            username: 'Admin User',
            email: 'admin@techznap.com',
            password: hashedPassword,
            role: 'Admin',
            jobCategory: 'Management'
        });

        await User.create({
            username: 'John Doe',
            email: 'john@techznap.com',
            password: hashedPassword,
            role: 'Employee',
            jobCategory: 'Full time'
        });

        console.log('Users seeded successfully!');
    } catch (err) {
        // console.error('Seeding failed:', err); 
        // Ignore duplicate errors if they exist
    } finally {
        await sequelize.close();
    }
}

seedUsers();
