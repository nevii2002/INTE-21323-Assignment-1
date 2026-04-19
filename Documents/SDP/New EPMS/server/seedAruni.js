const { User, Evaluation } = require('./models');
const sequelize = require('./config/database');

(async () => {
    try {
        await sequelize.query('PRAGMA foreign_keys = OFF;');
        const year = new Date().getFullYear();
        const month = String(new Date().getMonth() + 1).padStart(2, '0');
        const period = `${year}-${month}`;
        const aruniId = 24;

        const manager = await User.findOne({ where: { role: 'Manager' } });

        // Remove any existing Manager evaluation for Aruni this month
        await Evaluation.destroy({ where: { employeeId: aruniId, period, type: 'Manager' } });

        // Create a perfect 5-star evaluation
        const ev = await Evaluation.create({
            employeeId: aruniId,
            evaluatorId: manager.id,
            period,
            type: 'Manager',
            status: 'Completed',
            summaryScore: 5,
            feedback: 'Outstanding performance! Consistently exceeds all targets. Highly recommended for Employee of the Month.',
            dateVisible: new Date()
        });

        // Insert 3 detail rows with rating=5
        const kpiIds = [2, 3, 4];
        for (const kpiId of kpiIds) {
            await sequelize.query(
                `INSERT INTO EvaluationDetails (evaluationId, kpiId, rating, comment, createdAt, updatedAt)
                 VALUES (${ev.id}, ${kpiId}, 5, 'Exceptional performance.', datetime('now'), datetime('now'))`
            );
        }

        console.log(`Done! Aruni (id=${aruniId}) now has top scores for ${period}.`);
        process.exit(0);
    } catch (e) {
        console.error('Error:', e.message);
        process.exit(1);
    }
})();
