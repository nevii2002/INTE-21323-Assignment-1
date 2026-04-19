const { User, KPI, EmployeeKPI, sequelize } = require('../models');

async function testKPIAssignment() {
    try {
        await sequelize.authenticate();
        console.log('Database connected.');

        // 1. Setup Data: Get a User and a KPI
        const user = await User.findOne({ where: { role: 'Employee' } });
        const kpi = await KPI.findOne();

        if (!user || !kpi) {
            console.error('Test Failed: Need at least one Employee and one KPI in the DB.');
            process.exit(1);
        }

        console.log(`\nTesting Assignment for User: ${user.username} (ID: ${user.id})`);
        console.log(`Assigning KPI: ${kpi.title} (ID: ${kpi.id})`);

        // 2. Simulate Payload
        const assignments = [{
            kpiId: kpi.id,
            weight: 25,
            customBonus: 100
        }];

        // 3. Perform Assignment (Mimicking Controller Logic)
        // Clear existing for clean test
        await EmployeeKPI.destroy({ where: { employeeId: user.id } });

        const newAssignments = assignments.map(a => ({
            employeeId: user.id,
            kpiId: a.kpiId,
            customWeight: a.weight,
            customBonus: a.customBonus || 0
        }));

        await EmployeeKPI.bulkCreate(newAssignments);

        // 4. Verify in Database
        const verification = await EmployeeKPI.findOne({
            where: { employeeId: user.id, kpiId: kpi.id }
        });

        if (verification) {
            console.log('\n✅ SUCCESS: KPI Assigned Successfully!');
            console.log('-----------------------------------');
            console.log(`Employee ID: ${verification.employeeId}`);
            console.log(`KPI ID:      ${verification.kpiId}`);
            console.log(`Weight:      ${verification.customWeight}`);
            console.log(`Bonus:       ${verification.customBonus}`);
            console.log('-----------------------------------');
        } else {
            console.error('\n❌ FAILURE: KPI Record not found in EmployeeKPI table.');
        }

    } catch (error) {
        console.error('Test Error:', error);
    } finally {
        await sequelize.close();
    }
}

testKPIAssignment();
