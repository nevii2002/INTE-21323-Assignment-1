const { User, Evaluation, EvaluationDetail } = require('./models');
const sequelize = require('./config/database');

const seedUnderperforming = async () => {
    try {
        await sequelize.query('PRAGMA foreign_keys = OFF;');

        const currentYear = new Date().getFullYear();
        const currentMonth = new Date().getMonth() + 1;
        const period = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;

        const manager = await User.findOne({ where: { role: 'Manager' } });
        if (!manager) { console.log('No manager found!'); process.exit(1); }

        // Get 3 employee IDs
        const employees = await User.findAll({ where: { role: 'Employee' }, limit: 3 });

        for (const emp of employees) {
            await Evaluation.destroy({ where: { employeeId: emp.id, period, type: 'Manager' } });

            const evaluation = await Evaluation.create({
                employeeId: emp.id,
                evaluatorId: manager.id,
                period,
                type: 'Manager',
                status: 'Completed',
                summaryScore: 1,
                feedback: 'Significant improvement needed in overall performance.',
                dateVisible: new Date()
            });

            // Insert details directly without FK checks
            await sequelize.query(`
                INSERT INTO EvaluationDetails (evaluationId, kpiId, rating, comment, createdAt, updatedAt)
                VALUES (${evaluation.id}, 2, 1, 'Needs improvement.', datetime('now'), datetime('now'))
            `);

            console.log(`Seeded low evaluation for ${emp.username} (id=${emp.id})`);
        }

        console.log('Done! Refresh the admin dashboard to see the underperforming section.');
        process.exit(0);
    } catch (err) {
        console.error('Error:', err.message);
        process.exit(1);
    }
};

seedUnderperforming();
