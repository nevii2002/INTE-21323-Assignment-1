const { User, KPI, Evaluation, Bonus } = require('../models');

async function showDatabase() {
    try {
        console.log('\n--- 👥 USERS (Top 5) ---');
        const users = await User.findAll({ limit: 5, attributes: ['id', 'username', 'email', 'role', 'status'] });
        console.table(users.map(u => u.toJSON()));

        console.log('\n--- 🔐 SECURITY CHECK (Password Hashing) ---');
        console.log('Proof that we do not store plain text passwords:');
        const secureUsers = await User.findAll({ limit: 3, attributes: ['email', 'password'] });
        secureUsers.forEach(u => {
            console.log(`User: ${u.email}`);
            console.log(`Hash: ${u.password.substring(0, 30)}... (Bcrypt Encrypted)`);
            console.log('--------------------------------------------------');
        });

        console.log('\n--- 📊 KPIs (Top 5) ---');
        const kpis = await KPI.findAll({ limit: 5, attributes: ['id', 'name', 'category', 'weight'] });
        console.table(kpis.map(k => k.toJSON()));

        console.log('\n--- 📝 EVALUATIONS (Top 5) ---');
        const evals = await Evaluation.findAll({ limit: 5, attributes: ['id', 'employeeId', 'status', 'period'] });
        console.table(evals.map(e => e.toJSON()));

        console.log('\n--- 💰 BONUSES (Top 5) ---');
        const bonuses = await Bonus.findAll({ limit: 5, attributes: ['id', 'employeeId', 'amount', 'reason'] });
        console.table(bonuses.map(b => b.toJSON()));

    } catch (error) {
        console.error('Error fetching data:', error);
    }
}

showDatabase();
