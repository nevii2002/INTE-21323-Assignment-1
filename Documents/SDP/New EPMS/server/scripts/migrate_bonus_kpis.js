const sequelize = require('../config/database');
const KPI = require('../models/KPI');

const migrateAndSeed = async () => {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        // 1. Add Column via raw query (Safety check if exists skipped for simplicity, try/catch handles it)
        try {
            await sequelize.queryInterface.addColumn('KPIs', 'type', {
                type: require('sequelize').DataTypes.ENUM('EVALUATION', 'BONUS'),
                defaultValue: 'EVALUATION',
                allowNull: false
            });
            console.log('Added "type" column to KPIs table.');
        } catch (err) {
            console.log('"type" column might already exist or error:', err.message);
        }

        // 2. Update existing to EVALUATION
        await KPI.update({ type: 'EVALUATION' }, { where: {} });

        // 3. Seed Bonus Projects
        const bonusProjects = [
            { title: 'YouTube ad sense revenue', targetValue: 1000, weight: 0.20, unit: 'Currency', category: 'Quantitative' },
            { title: 'YouTube collab revenue', targetValue: 500, weight: 0.30, unit: 'Currency', category: 'Quantitative' },
            { title: 'Fiverr - Graphic Design', targetValue: 10, weight: 0.25, unit: 'Count', category: 'Quantitative' },
            { title: 'Fiverr - Video Editing', targetValue: 5, weight: 0.25, unit: 'Count', category: 'Quantitative' }
        ];

        for (const proj of bonusProjects) {
            const [kpi, created] = await KPI.findOrCreate({
                where: { title: proj.title },
                defaults: {
                    ...proj,
                    description: 'Bonus Project',
                    type: 'BONUS',
                    dataSource: 'Manual'
                }
            });

            if (!created) {
                await kpi.update({ type: 'BONUS' }); // Ensure type is set if exists
            }
        }

        console.log('Bonus Projects seeded successfully.');
        process.exit(0);
    } catch (error) {
        console.error('Migration failed:', error);
        process.exit(1);
    }
};

migrateAndSeed();
