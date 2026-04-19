const { sequelize, User, KPI, EmployeeKPI } = require('../models');

async function testAssignment() {
    try {
        await sequelize.authenticate();
        console.log('DB Connected.');

        // 1. Get or Create User
        let user = await User.findOne({ where: { email: 'test_assign@example.com' } });
        if (!user) {
            user = await User.create({
                username: 'Test Assignee',
                email: 'test_assign@example.com',
                password: 'password',
                role: 'Employee'
            });
        }
        console.log(`User ID: ${user.id}`);

        // 2. Get Quantitative KPIs
        const kpis = await KPI.findAll({ where: { category: 'Quantitative' }, limit: 2 });
        if (kpis.length < 2) {
            console.log('Not enough Quantitative KPIs to test.');
            return;
        }
        const kpi1 = kpis[0];
        const kpi2 = kpis[1];
        console.log(`Testing with KPIs: ${kpi1.title} (ID: ${kpi1.id}), ${kpi2.title} (ID: ${kpi2.id})`);

        // 3. Assign Custom Weights (Total 100)
        await EmployeeKPI.destroy({ where: { employeeId: user.id } }); // Clear previous

        const assignments = [
            { employeeId: user.id, kpiId: kpi1.id, customWeight: 30 },
            { employeeId: user.id, kpiId: kpi2.id, customWeight: 70 }
        ];

        await EmployeeKPI.bulkCreate(assignments);
        console.log('Assignments Created: 30% / 70%');

        // 4. Verify Fetch (Simulate ManagerEvaluation logic)
        const fetchedUser = await User.findByPk(user.id, {
            include: [{
                model: KPI,
                as: 'assignedKPIs',
                through: { attributes: ['customWeight'] }
            }]
        });

        console.log('Fetched Assigned KPIs:');
        fetchedUser.assignedKPIs.forEach(k => {
            console.log(`- ${k.title}: Custom Weight = ${k.EmployeeKPI.customWeight}% (Default: ${k.weight}%)`);
        });

        const w1 = fetchedUser.assignedKPIs.find(k => k.id === kpi1.id).EmployeeKPI.customWeight;
        const w2 = fetchedUser.assignedKPIs.find(k => k.id === kpi2.id).EmployeeKPI.customWeight;

        if (w1 === 30 && w2 === 70) {
            console.log('SUCCESS: Custom weights verified.');
        } else {
            console.error('FAILURE: Weights do not match.');
        }

    } catch (error) {
        console.error('Error:', error);
    } finally {
        await sequelize.close();
    }
}

testAssignment();
