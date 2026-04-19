const sequelize = require('../config/database');
const { User, KPI, Evaluation, EvaluationDetail, Bonus, EmployeeKPI } = require('../models');

async function restore() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        const users = await User.findAll({ where: { role: 'Employee' } });
        console.log(`Found ${users.length} employees.`);

        const kpis = await KPI.findAll({ where: { type: 'EVALUATION' } });
        if (kpis.length === 0) {
            console.log("No KPIs found! Please run seed_new_kpis.js first.");
            return;
        }

        for (const user of users) {
            // 1. Restore Assignments if missing
            const existingAssignments = await EmployeeKPI.count({ where: { employeeId: user.id } });
            if (existingAssignments === 0) {
                // Assign 5 random KPIs
                const toAssign = kpis.slice(0, 5).map(k => ({
                    employeeId: user.id,
                    kpiId: k.id,
                    customWeight: 20, // 20 * 5 = 100
                    customBonus: 0
                }));
                await EmployeeKPI.bulkCreate(toAssign);
                console.log(`Restored assignments for ${user.username}`);
            }

            // 2. Add an Evaluation (if missing)
            const existingEval = await Evaluation.count({ where: { employeeId: user.id } });
            if (existingEval === 0) {
                const evaluation = await Evaluation.create({
                    employeeId: user.id,
                    evaluatorId: 1, // Assuming admin is ID 1
                    type: 'Annual',
                    period: '2024',
                    status: 'Submitted',
                    comments: 'Restored Demo Data'
                });

                // Add details (random ratings)
                const details = kpis.slice(0, 5).map(k => ({
                    evaluationId: evaluation.id,
                    kpiId: k.id,
                    rating: Math.floor(Math.random() * 3) + 3, // 3, 4, or 5
                    comment: 'Good performance'
                }));
                await EvaluationDetail.bulkCreate(details);
                console.log(`Restored evaluation for ${user.username}`);
            }

            // 3. Add Bonuses (if missing)
            const existingBonus = await Bonus.count({ where: { employeeId: user.id } });
            if (existingBonus === 0) {
                await Bonus.create({
                    employeeId: user.id,
                    amount: Math.floor(Math.random() * 10000) + 5000,
                    reason: 'Performance Bonus (Restored)',
                    dateGiven: new Date()
                });
                console.log(`Restored bonus for ${user.username}`);
            }
        }

        console.log('Data restoration complete.');

    } catch (error) {
        console.error('Restore failed:', error);
    } finally {
        await sequelize.close();
    }
}

restore();
