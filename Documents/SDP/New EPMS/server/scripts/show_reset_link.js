// const axios = require('axios'); // Not needed for local controller simulation

async function triggerPasswordReset() {
    try {
        console.log('Sending Password Reset Request...');
        // Assuming the server is running on port 5000 based on standard practice in this repo
        // If not, this might fail, but the user is likely running the server separately.
        // However, since I can't see the user's running server console, 
        // I will use a direct call to the controller logic if possible, 
        // OR better yet, I will run a script that imports the logic to FORCE the log to appear HERE.

        // OPTION 2: Import controller logic to show log in THIS terminal
        // We need to mock req/res for the controller
        const authController = require('../controllers/authController');
        const { User, sequelize } = require('../models');

        await sequelize.authenticate();

        // Find a user to reset
        const user = await User.findOne();
        if (!user) {
            console.log('No user found to reset.');
            return;
        }

        const req = { body: { email: user.email } };
        const res = {
            json: (data) => console.log('Response JSON:', data),
            status: (code) => ({ json: (data) => console.log(`Response ${code}:`, data) })
        };

        console.log(`\n--- Simulating Request for ${user.email} ---`);
        await authController.forgotPassword(req, res);
        console.log('-------------------------------------------');

    } catch (error) {
        console.error('Error:', error);
    }
}

triggerPasswordReset();
