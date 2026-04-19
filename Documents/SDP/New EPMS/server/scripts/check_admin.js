const { User } = require('../models');
const sequelize = require('../config/database');

async function checkAdmin() {
    try {
        await sequelize.authenticate();
        const admin = await User.findOne({ where: { email: 'admin@techznap.com' } });
        if (admin) {
            console.log('Admin user exists:', admin.email);
            console.log('Role:', admin.role);
            console.log('Status:', admin.status);
        } else {
            console.log('Admin user NOT found!');
        }
    } catch (e) {
        console.error(e);
    } finally {
        await sequelize.close();
    }
}
checkAdmin();
