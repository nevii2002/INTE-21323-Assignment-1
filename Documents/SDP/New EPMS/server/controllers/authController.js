const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const nodemailer = require('nodemailer');
const User = require('../models/User');

// --- Nodemailer Setup (Using Ethereal for testing) ---
let transporter;

async function initMailer() {
    try {
        let testAccount = await nodemailer.createTestAccount();
        transporter = nodemailer.createTransport({
            host: "smtp.ethereal.email",
            port: 587,
            secure: false, // true for 465, false for other ports
            auth: {
                user: testAccount.user, // generated ethereal user
                pass: testAccount.pass, // generated ethereal password
            },
        });
        console.log(`✉️ Mailer initialized (Ethereal test account).`);
    } catch (err) {
        console.error("Failed to initialize mailer", err);
    }
}
initMailer();
// -----------------------------------------------------

exports.register = async (req, res) => {
    try {
        const { username, email, password, mobileNumber, position } = req.body;

        const existingUser = await User.findOne({ where: { email } });

        // CASE 1: No User Found -> Error (Invite Only)
        if (!existingUser) {
            return res.status(400).json({ message: 'Email not found. Please ask your administrator to invite you first.' });
        }

        // CASE 2: User Found but already Active -> Error
        if (existingUser.status === 'Active' && existingUser.password) {
            return res.status(400).json({ message: 'Account already active. Please login.' });
        }

        // CASE 3: User Found (Pending) -> Activate
        // Note: We might want to verify username matches, but let's trust email for now and update username if provided.
        const hashedPassword = await bcrypt.hash(password, 10);

        // Update the existing record
        existingUser.password = hashedPassword;
        existingUser.status = 'Active';
        if (username) existingUser.username = username; // Allow updating name
        if (mobileNumber) existingUser.mobileNumber = mobileNumber;
        if (position) existingUser.position = position;

        await existingUser.save();

        res.status(200).json({
            message: 'Account activated successfully',
            user: { id: existingUser.id, username: existingUser.username, email: existingUser.email, role: existingUser.role, profilePicture: existingUser.profilePicture }
        });

    } catch (error) {
        console.error("Register Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.login = async (req, res) => {
    try {
        const { email, password } = req.body;

        const user = await User.findOne({ where: { email } });
        if (!user) {
            return res.status(400).json({ message: 'Invalid credentials' });
        }

        const isMatch = await bcrypt.compare(password, user.password);
        if (!isMatch) {
            return res.status(400).json({ message: 'Invalid credentials' });
        }

        const token = jwt.sign({ userId: user.id, role: user.role }, process.env.JWT_SECRET || 'techznap_secret_key', { expiresIn: '1d' });

        res.json({ token, user: { id: user.id, username: user.username, email: user.email, role: user.role, profilePicture: user.profilePicture } });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.forgotPassword = async (req, res) => {
    try {
        const { email } = req.body;
        const user = await User.findOne({ where: { email } });

        if (!user) {
            // For security, don't reveal if user exists, but here we can be helpful for dev
            return res.status(404).json({ message: 'User not found' });
        }

        // Generate a temporary token valid for 15 minutes
        const resetToken = jwt.sign({ userId: user.id }, process.env.JWT_SECRET || 'techznap_secret_key', { expiresIn: '15m' });

        // Link points to the frontend React server port 3000
        const resetLink = `http://localhost:3000/reset-password/${resetToken}`;

        try {
            let info = await transporter.sendMail({
                from: '"Techznap HR System" <hr@techznap.com>',
                to: email, // The user's email
                subject: "Password Reset Request",
                html: `
                    <div style="font-family: Arial, sans-serif; max-w-md; margin: 0 auto; padding: 20px;">
                        <h2>Password Reset Request</h2>
                        <p>Hi ${user.username || 'User'},</p>
                        <p>We received a request to reset your password for the Techznap Employee Performance Management System.</p>
                        <p>Click the button below to reset your password. This link is valid for 15 minutes.</p>
                        <a href="${resetLink}" style="display: inline-block; padding: 10px 20px; background-color: #2563eb; color: #ffffff; text-decoration: none; border-radius: 5px; margin: 20px 0;">Reset Password</a>
                        <p>If you did not request a password reset, please ignore this email.</p>
                        <hr>
                        <p style="color: #6b7280; font-size: 12px;">This is an automated message. Please do not reply.</p>
                    </div>
                `,
            });
            console.log("Message sent: %s", info.messageId);
            console.log("Preview URL: %s", nodemailer.getTestMessageUrl(info));
        } catch (mailError) {
            console.error("Error sending email:", mailError);
        }

        res.json({ message: 'If an account exists, a reset link has been sent to your email.' });
    } catch (error) {
        console.error("Forgot Password Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.resetPassword = async (req, res) => {
    try {
        const { token, newPassword } = req.body;

        // Verify Token
        let decoded;
        try {
            decoded = jwt.verify(token, process.env.JWT_SECRET || 'techznap_secret_key');
        } catch (err) {
            return res.status(400).json({ message: 'Invalid or expired token' });
        }

        const user = await User.findByPk(decoded.userId);
        if (!user) {
            return res.status(400).json({ message: 'User not found' });
        }

        // Update Password
        const hashedPassword = await bcrypt.hash(newPassword, 10);
        user.password = hashedPassword;
        await user.save();

        res.json({ message: 'Password has been reset successfully. Please login.' });
    } catch (error) {
        console.error("Reset Password Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getProfile = async (req, res) => {
    try {
        const userId = req.user.userId;
        const user = await User.findByPk(userId, {
            attributes: { exclude: ['password'] }
        });

        if (!user) return res.status(404).json({ message: 'User not found' });
        res.json(user);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.updateProfile = async (req, res) => {
    try {
        const userId = req.user.userId;
        const { username, mobileNumber, position } = req.body;

        const user = await User.findByPk(userId);
        if (!user) return res.status(404).json({ message: 'User not found' });

        if (username) user.username = username;
        if (mobileNumber !== undefined) user.mobileNumber = mobileNumber;
        if (position !== undefined) user.position = position;

        await user.save();

        // Return updated user without password
        const updatedUser = user.toJSON();
        delete updatedUser.password;

        res.json({ message: 'Profile updated successfully', user: updatedUser });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.changePassword = async (req, res) => {
    try {
        const userId = req.user.userId;
        const { currentPassword, newPassword } = req.body;

        const user = await User.findByPk(userId);
        if (!user) return res.status(404).json({ message: 'User not found' });

        // Verify current password
        const isMatch = await bcrypt.compare(currentPassword, user.password);
        if (!isMatch) {
            return res.status(400).json({ message: 'Incorrect current password' });
        }

        // Hash and save new password
        const hashedPassword = await bcrypt.hash(newPassword, 10);
        user.password = hashedPassword;
        await user.save();

        res.json({ message: 'Password changed successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.uploadProfilePicture = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: 'No file uploaded' });
        }

        const userId = req.user.userId;
        const user = await User.findByPk(userId);

        if (!user) {
            return res.status(404).json({ message: 'User not found' });
        }

        // The file is saved to uploads/profiles/ by multer
        const fileUrl = `/uploads/profiles/${req.file.filename}`;

        user.profilePicture = fileUrl;
        await user.save();

        res.json({
            message: 'Profile picture updated successfully',
            profilePicture: fileUrl
        });
    } catch (error) {
        console.error("Profile Picture Upload Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
