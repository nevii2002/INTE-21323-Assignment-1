const bcrypt = require('bcryptjs');
const { User, KPI, EmployeeKPI, Evaluation, EvaluationDetail, QuantitativeLog, CompanyMetric, CompanyMetricLog } = require('./models');
const sequelize = require('./config/database');

const seedData = async () => {
    try {
        console.log('--- Starting Seeding Process based on Employee Performance Sheet ---');
        await sequelize.query('PRAGMA foreign_keys = OFF;');

        // 1. Create Admins and Managers
        const salt = await bcrypt.genSalt(10);
        const hashedPassword = await bcrypt.hash('password123', salt);

        const admin = await User.findOrCreate({
            where: { email: 'admin@techznap.com' },
            defaults: { username: 'Admin User', role: 'Admin', password: hashedPassword }
        });
        const adminId = admin[0].id;

        const manager = await User.findOrCreate({
            where: { email: 'manager@techznap.com' },
            defaults: { username: 'Sales Manager', role: 'Manager', password: hashedPassword }
        });
        const managerId = manager[0].id;

        // 2. Create 5 Employees
        const employeesData = [
            { email: 'alice@techznap.com', username: 'Alice Smith', position: 'Sales Representative' },
            { email: 'bob@techznap.com', username: 'Bob Johnson', position: 'Support Specialist' },
            { email: 'charlie@techznap.com', username: 'Charlie Brown', position: 'Software Engineer' },
            { email: 'david@techznap.com', username: 'David Lee', position: 'Marketing Associate' },
            { email: 'eve@techznap.com', username: 'Eve Adams', position: 'HR Coordinator' }
        ];

        const employeeIds = [];
        for (const emp of employeesData) {
            const user = await User.findOrCreate({
                where: { email: emp.email },
                defaults: { username: emp.username, role: 'Employee', password: hashedPassword }
            });
            employeeIds.push(user[0].id);
        }

        // 3. Create KPIs based on the Google Sheet "Performance Indicators"
        const kpiTopics = [
            { title: 'Job Knowledge', description: 'Understands job duties and responsibilities.', weight: 20 },
            { title: 'Work Quality', description: 'Maintains high standards of accuracy and quality.', weight: 20 },
            { title: 'Productivity', description: 'Meets assigned targets and goals.', weight: 30, unit: 'Tasks' },
            { title: 'Teamwork', description: 'Collaborates effectively with peers.', weight: 15 },
            { title: 'Communication', description: 'Communicates clearly and professionally.', weight: 15 }
        ];

        const kpis = [];
        for (const kt of kpiTopics) {
            const kpi = await KPI.findOrCreate({
                where: { title: kt.title },
                defaults: { description: kt.description, category: 'Quantitative', weight: kt.weight, unit: kt.unit || 'Score 1-10', targetValue: 10 }
            });
            kpis.push(kpi[0]);
        }

        // 4. Assign KPIs to Employees and add Log Actuals and Evaluations
        const currentYear = new Date().getFullYear();
        const currentMonth = new Date().getMonth() + 1;
        const periodStr = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;

        for (const empId of employeeIds) {
            let totalWeight = 0;
            // EmployeeKPI assignments
            for (const kpi of kpis) {
                console.log(`Trying to assign KPI ${kpi?.id} to Employee ${empId}`);
                try {
                    await EmployeeKPI.findOrCreate({
                        where: { employeeId: empId, kpiId: kpi.id },
                        defaults: { customWeight: kpi.weight }
                    });
                } catch (fkErr) {
                    console.log(`Skipping EmployeeKPI assignment due to localized FK error: ${fkErr.message}`);
                }
                totalWeight += kpi.weight;

                // Quantitative Logs (Performance Sheet Table entries)
                const randomActual = Math.floor(Math.random() * 5) + 5; // Score 5-10
                await QuantitativeLog.findOrCreate({
                    where: { employeeId: empId, kpiId: kpi.id, period: periodStr },
                    defaults: { actualValue: randomActual }
                });
            }

            // Evaluations Table (from manager to employee)
            const evaluation = await Evaluation.create({
                employeeId: empId,
                evaluatorId: managerId,
                period: periodStr,
                type: 'Manager',
                status: 'Completed',
                summaryScore: Math.floor(Math.random() * 3) + 3, // 3 to 5
                feedback: 'Great performance this month! Continuing to show solid job knowledge as rated in our standard sheets.',
                dateVisible: new Date()
            });

            for (const kpi of kpis) {
                await EvaluationDetail.create({
                    evaluationId: evaluation.id,
                    kpiId: kpi.id,
                    rating: Math.floor(Math.random() * 3) + 3,
                    comments: 'Good effort.'
                });
            }
        }

        // 5. Create Company Measures from the Performance Survey Sheet concepts
        const compMetrics = [
            { name: 'Employee Satisfaction Score (eNPS)', description: 'Monthly survey results for overall satisfaction.', unit: '%' },
            { name: 'Training Completion Rate', description: 'Percentage of staff that finished pro development training.', unit: '%' },
            { name: 'Customer Satisfaction (CSAT)', description: 'Average rating from customer feedback loops.', unit: '%' }
        ];

        for (const cm of compMetrics) {
            const metric = await CompanyMetric.findOrCreate({
                where: { name: cm.name },
                defaults: { description: cm.description, unit: cm.unit }
            });

            // Log data for this month
            await CompanyMetricLog.upsert({
                metricId: metric[0].id,
                period: periodStr,
                value: Math.floor(Math.random() * 30) + 70 // 70 to 100
            });
            // And last month
            const lastMonth = currentMonth === 1 ? 12 : currentMonth - 1;
            const lastYear = currentMonth === 1 ? currentYear - 1 : currentYear;
            await CompanyMetricLog.upsert({
                metricId: metric[0].id,
                period: `${lastYear}-${String(lastMonth).padStart(2, '0')}`,
                value: Math.floor(Math.random() * 30) + 60
            });
        }

        console.log('--- Database Seeding Complete! ---');
        process.exit(0);

    } catch (err) {
        console.error('Seeding failed:', err);
        process.exit(1);
    }
};

seedData();
