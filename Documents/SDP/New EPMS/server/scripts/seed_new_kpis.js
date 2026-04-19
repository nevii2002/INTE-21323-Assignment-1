const { sequelize, KPI } = require('../models');

const newKpis = [
    // Strategic & Financial
    { title: 'Total New Orders', description: 'Number of new business generated.', category: 'Quantitative', unit: 'Count', weight: 10, targetValue: 50 },
    { title: 'Net Profit', description: 'Financial impact and profitability of projects.', category: 'Quantitative', unit: 'Currency', weight: 15, targetValue: 10000 },
    { title: 'Average Order Value', description: 'Mean revenue generated per order.', category: 'Quantitative', unit: 'Currency', weight: 10, targetValue: 500 },
    { title: 'Return on Investment (ROI)', description: 'Efficiency and financial gain of resources.', category: 'Quantitative', unit: 'Percentage', weight: 10, targetValue: 120 },

    // Operational & Efficiency
    { title: 'Order Completion Rate', description: 'Percentage of successfully delivered orders.', category: 'Quantitative', unit: 'Percentage', weight: 10, targetValue: 95 },
    { title: 'On-Time Delivery Rate', description: 'Ability to meet project deadlines.', category: 'Quantitative', unit: 'Percentage', weight: 10, targetValue: 98 },
    { title: 'Cancellation Rate', description: 'Percentage of failed/cancelled tasks.', category: 'Quantitative', unit: 'Percentage', weight: 5, targetValue: 2 },
    { title: 'Average Response Time', description: 'Speed of communication with stakeholders.', category: 'Quantitative', unit: 'Hours', weight: 5, targetValue: 2 },

    // Customer Satisfaction & Quality
    { title: 'Customer Satisfaction Score', description: 'Direct ratings from clients (x/5).', category: 'Quantitative', unit: 'Score (1-5)', weight: 10, targetValue: 4.5 },
    { title: 'Repeat Customer Rate', description: 'Frequency of returning clients.', category: 'Quantitative', unit: 'Percentage', weight: 10, targetValue: 30 },
    { title: 'Quality of Work Product', description: 'Subjective rating of accuracy and thoroughness.', category: 'Qualitative', unit: 'Rating', weight: 10 },
    { title: 'Dispute Rate', description: 'Tasks resulting in formal complaints.', category: 'Quantitative', unit: 'Percentage', weight: 5, targetValue: 1 },

    // Service & Engagement
    { title: 'Disputes Converted to Satisfied', description: 'Problem-solving skills in conflicts.', category: 'Quantitative', unit: 'Count', weight: 5, targetValue: 5 },
    { title: 'Tips or Recognition', description: 'External positive feedback or rewards.', category: 'Quantitative', unit: 'Currency/Count', weight: 5, targetValue: 100 },
    { title: 'Adaptability & Learning', description: 'Speed of mastering new technologies.', category: 'Qualitative', unit: 'Rating', weight: 10 }
];

async function seed() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        for (const kpiData of newKpis) {
            const [kpi, created] = await KPI.findOrCreate({
                where: { title: kpiData.title },
                defaults: kpiData
            });
            if (created) console.log(`Created: ${kpi.title}`);
            else console.log(`Exists: ${kpi.title}`);
        }

        console.log('Done.');
    } catch (err) {
        console.error('Error:', err);
    } finally {
        await sequelize.close();
    }
}

seed();
