const { sequelize, KPI } = require('./models');

async function seedKpis() {
    try {
        await sequelize.authenticate();
        await sequelize.sync();

        // Check if KPIs exist
        const count = await KPI.count();
        if (count > 0) {
            console.log('KPIs already exist. Skipping seed.');
            return;
        }

        // Qualitative
        await KPI.create({
            title: 'Team Collaboration',
            description: 'Participates in team meetings and helps others.',
            category: 'Qualitative',
            weight: 20
        });

        await KPI.create({
            title: 'Communication Skills',
            description: 'Clear and concise communication.',
            category: 'Qualitative',
            weight: 10
        });

        // Quantitative
        await KPI.create({
            title: 'Sales Revenue',
            category: 'Quantitative',
            unit: 'Currency',
            weight: 40,
            targetValue: 50000,
            dataSource: 'Salesforce'
        });

        await KPI.create({
            title: 'Projects Completed',
            category: 'Quantitative',
            unit: 'Count',
            weight: 30,
            targetValue: 10,
            dataSource: 'Jira'
        });

        console.log('KPIs seeded successfully!');
    } catch (err) {
        console.error('Seeding failed:', err);
    } finally {
        await sequelize.close();
    }
}

seedKpis();
