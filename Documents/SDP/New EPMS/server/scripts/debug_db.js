const sequelize = require('../config/database');
const { User, Bonus, Evaluation, KPI } = require('../models');

async function check() {
    try {
        await sequelize.authenticate();

        const userCount = await User.count();
        console.log(`Users (New): ${userCount}`);

        try {
            const [results] = await sequelize.query("SELECT count(*) as count FROM Users_old");
            console.log(`Users_old: ${results[0].count}`);
        } catch (e) {
            console.log("Users_old table does not exist.");
        }

        const bonusCount = await Bonus.count();
        console.log(`Bonuses: ${bonusCount}`);

        const evalCount = await Evaluation.count();
        console.log(`Evaluations: ${evalCount}`);

        const kpiCount = await KPI.count();
        console.log(`KPIs: ${kpiCount}`);

    } catch (error) {
        console.error('Check failed:', error);
    } finally {
        await sequelize.close();
    }
}

check();
