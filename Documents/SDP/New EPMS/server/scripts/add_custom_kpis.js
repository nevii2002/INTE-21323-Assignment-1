const { sequelize, KPI } = require('../models');

const newKPIs = [
    {
        title: 'YouTube Watch Time',
        description: 'Total watch time accumulated on the channel.',
        targetValue: 16000,
        unit: 'Hours',
        weight: 20,
        category: 'Quantitative',
        dataSource: 'YouTube Studio',
        role: 'All'
    },
    {
        title: 'Cash Collected from Collaborations',
        description: 'Revenue generated specifically from brand collaborations.',
        targetValue: 1000,
        unit: 'USD',
        weight: 15,
        category: 'Quantitative',
        dataSource: 'Finance Records',
        role: 'All'
    },
    {
        title: 'Fiverr Income (Graphic Design)',
        description: 'Earnings from Graphic Design gigs.',
        targetValue: 1750,
        unit: 'USD',
        weight: 20,
        category: 'Quantitative',
        dataSource: 'Fiverr Earnings',
        role: 'All'
    },
    {
        title: 'Fiverr Income (Video & Animation)',
        description: 'Earnings from Video Editing & Animation gigs.',
        targetValue: 1000,
        unit: 'USD',
        weight: 15,
        category: 'Quantitative',
        dataSource: 'Fiverr Earnings',
        role: 'All'
    }
];

async function seed() {
    try {
        await sequelize.authenticate();
        await sequelize.sync({ alter: true }); // Sync enabled to add new column

        for (const kpi of newKPIs) {
            await KPI.create(kpi);
            console.log(`Added KPI: ${kpi.title}`);
        }

        console.log('All custom KPIs added successfully.');
    } catch (error) {
        console.error('Error adding KPIs:', error);
    } finally {
        await sequelize.close();
    }
}

seed();
