const User = require('../models/User');
const KPI = require('../models/KPI');
const EmployeeKPI = require('../models/EmployeeKPI');
const bcrypt = require('bcryptjs');
const { logAction } = require('../utils/auditLogger');

exports.getAllStaff = async (req, res) => {
    try {
        const staff = await User.findAll({
            attributes: { exclude: ['password'] }
        });

        // Security: Strip private compensation data for regular Employees viewing the Team Overview
        const userRole = req.user.role;
        if (userRole === 'Employee') {
            const sanitizedStaff = staff.map(s => {
                const sObj = s.toJSON();
                delete sObj.basicSalary;
                delete sObj.bonusAmount;
                return sObj;
            });
            return res.json(sanitizedStaff);
        }

        res.json(staff);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getStaffById = async (req, res) => {
    try {
        const { id } = req.params;
        const user = await User.findByPk(id, {
            attributes: { exclude: ['password'] }
        });
        if (!user) return res.status(404).json({ message: 'User not found' });
        res.json(user);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.createStaff = async (req, res) => {
    try {
        const { username, email, role, jobCategory, jobDescription, responsibilities } = req.body;

        const existingUser = await User.findOne({ where: { email } });
        if (existingUser) {
            return res.status(400).json({ message: 'User already exists' });
        }

        // Create Pending User (No password)
        const newUser = await User.create({
            username,
            email,
            password: null, // No password yet
            role: role || 'Employee',
            jobCategory: jobCategory || 'Full time',
            jobDescription: jobDescription || '',
            responsibilities: responsibilities || '',
            status: 'Pending'
        });

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'CREATE', 'User', { userId: newUser.id, username: newUser.username });
        }

        res.status(201).json({ message: 'Staff invited successfully. They can now register.', user: newUser });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.updateStaff = async (req, res) => {
    try {
        const { id } = req.params;
        const { username, email, role, jobCategory, basicSalary, bonusAmount, jobDescription, responsibilities, position, department, startDate } = req.body;

        const user = await User.findByPk(id);
        if (!user) return res.status(404).json({ message: 'User not found' });

        user.username = username || user.username;
        user.email = email || user.email;
        user.role = role || user.role;
        user.jobCategory = jobCategory || user.jobCategory;
        if (jobDescription !== undefined) user.jobDescription = jobDescription;
        if (responsibilities !== undefined) user.responsibilities = responsibilities;
        if (basicSalary !== undefined) user.basicSalary = basicSalary;
        if (bonusAmount !== undefined) user.bonusAmount = bonusAmount;
        if (position !== undefined) user.position = position;
        if (department !== undefined) user.department = department;
        if (startDate !== undefined) user.startDate = startDate || null;

        await user.save();

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'UPDATE', 'User', { userId: user.id, username: user.username });
        }

        res.json(user);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.deleteStaff = async (req, res) => {
    try {
        const { id } = req.params;
        const user = await User.findByPk(id);
        if (!user) return res.status(404).json({ message: 'User not found' });

        await user.destroy();

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'DELETE', 'User', { userId: id, username: user.username });
        }

        res.json({ message: 'User deleted successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

// --- KPI Assignment Logic ---

exports.getEmployeeKPIs = async (req, res) => {
    try {
        const { id } = req.params;
        const user = await User.findByPk(id, {
            include: [{
                model: KPI,
                as: 'assignedKPIs',
                through: { attributes: ['customWeight', 'customBonus'] }
            }]
        });

        if (!user) return res.status(404).json({ message: 'User not found' });

        res.json(user.assignedKPIs);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.assignKPIs = async (req, res) => {
    try {
        const { id } = req.params;
        const { assignments } = req.body; // Array of { kpiId, weight, customBonus }

        const user = await User.findByPk(id);
        if (!user) return res.status(404).json({ message: 'User not found' });

        // NEW VALIDATION: Check if any KPI is over-allocated (Sum of weights across ALL employees > 100)
        // This is complex because we are updating THIS user's assignments.

        // 1. Get all assignments for the KPIs involved in this update, EXCLUDING the current user
        // But first, let's just proceed with a loop check for safety.
        if (assignments && assignments.length > 0) {
            for (const assignment of assignments) {
                const kpiId = assignment.kpiId;
                const newWeight = parseFloat(assignment.weight) || 0;

                // Find sum of weights for this KPI from ALL OTHER employees
                const otherAssignments = await EmployeeKPI.findAll({
                    where: {
                        kpiId: kpiId,
                        employeeId: { [require('sequelize').Op.ne]: id } // Not current user
                    }
                });

                const currentTotal = otherAssignments.reduce((sum, a) => sum + (a.customWeight || 0), 0);

                if (currentTotal + newWeight > 100) {
                    // Fetch KPI name for better error
                    const kpi = await KPI.findByPk(kpiId);
                    return res.status(400).json({
                        message: `Validation Error: KPI "${kpi ? kpi.title : kpiId}" is over-allocated. Current total: ${currentTotal}%. Adding ${newWeight}% exceeds 100%.`
                    });
                }
            }
        }

        // If validation passes, save.
        await EmployeeKPI.destroy({ where: { employeeId: id } });

        if (assignments && assignments.length > 0) {
            const newAssignments = assignments.map(a => ({
                employeeId: id,
                kpiId: a.kpiId,
                customWeight: a.weight,
                customBonus: a.customBonus || 0
            }));
            await EmployeeKPI.bulkCreate(newAssignments);
        }

        if (req.user && req.user.userId) {
            await logAction(req.user.userId, 'UPDATE', 'EmployeeKPIs', { targetEmployeeId: id, totalAssignments: assignments ? assignments.length : 0 });
        }

        res.json({ message: 'KPIs assigned successfully' });
    } catch (error) {
        console.error("Assignment Error", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getMyKPIs = async (req, res) => {
    try {
        const userId = req.user.userId;
        const user = await User.findByPk(userId, {
            include: [{
                model: KPI,
                as: 'assignedKPIs',
                through: { attributes: ['customWeight', 'customBonus'] }
            }]
        });

        if (!user) return res.status(404).json({ message: 'User not found' });

        res.json(user.assignedKPIs);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
