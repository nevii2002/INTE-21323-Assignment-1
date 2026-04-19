const { Bonus, User } = require('../models');

exports.createBonus = async (req, res) => {
    try {
        const { employeeId, amount, reason, dateGiven } = req.body;

        const user = await User.findByPk(employeeId);
        if (!user) {
            return res.status(404).json({ message: 'Employee not found' });
        }

        const bonus = await Bonus.create({
            employeeId,
            amount,
            reason,
            dateGiven: dateGiven || new Date()
        });

        // Optionally update the aggregate bonusAmount in User table if used for current period cache
        // But the requirements say "View history", so this table is the source of truth.
        // We might want to increment the user's bonusAmount cache if that field is still used for something specific.
        // For now, let's keep them separate or use this as the record.

        // If we want to keep the User.bonusAmount as a "Total Lifetime Bonus" or "Current Pending Bonus", we can update it.
        // Let's treat User.bonusAmount as potentially legacy or a dashboard cache, but we'll update it for consistency if needed.
        // For now, let's just save the record.

        res.status(201).json({ message: 'Bonus awarded successfully', bonus });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getBonusesByEmployee = async (req, res) => {
    try {
        const { employeeId } = req.params;
        const bonuses = await Bonus.findAll({
            where: { employeeId },
            order: [['dateGiven', 'DESC']]
        });
        res.json(bonuses);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getAllBonuses = async (req, res) => {
    try {
        const bonuses = await Bonus.findAll({
            include: [{
                model: User,
                as: 'employee',
                attributes: ['id', 'username', 'email']
            }],
            order: [['createdAt', 'DESC']]
        });
        res.json(bonuses);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.deleteBonus = async (req, res) => {
    try {
        const { id } = req.params;
        const bonus = await Bonus.findByPk(id);
        if (!bonus) {
            return res.status(404).json({ message: 'Bonus record not found' });
        }
        await bonus.destroy();
        res.json({ message: 'Bonus removed successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
