const { Evaluation, EvaluationDetail, User, KPI } = require('../models');

exports.createEvaluation = async (req, res) => {
    try {
        const { employeeId, type, period, details, comments } = req.body;
        const evaluatorId = req.user.userId; // From authMiddleware

        // Check if evaluation already exists for this period
        const existing = await Evaluation.findOne({
            where: { employeeId, period, type }
        });

        if (existing) {
            return res.status(400).json({ message: 'Evaluation already exists for this period.' });
        }

        // Create Evaluation
        const evaluation = await Evaluation.create({
            employeeId,
            evaluatorId,
            type,
            period,
            status: 'Submitted', // Direct submission for now
            comments
        });

        // Create Details
        if (details && details.length > 0) {
            const detailRecords = details.map(d => ({
                evaluationId: evaluation.id,
                kpiId: d.kpiId,
                rating: d.rating,
                comment: d.comment
            }));
            await EvaluationDetail.bulkCreate(detailRecords);
        }

        res.status(201).json({ message: 'Evaluation submitted successfully', evaluation });
    } catch (error) {
        console.error('Create Eval Error:', error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getEvaluations = async (req, res) => {
    try {
        const { role, userId } = req.user;
        let whereClause = {};

        // If employee, see own received evaluations
        if (role === 'Employee') {
            whereClause = { employeeId: userId, status: 'Submitted' };
        }
        // If manager/admin, can see all (or filter by query params)
        else {
            // Optional: Filter by specific employee if requested
            if (req.query.employeeId) {
                whereClause.employeeId = req.query.employeeId;
            }
        }

        const evaluations = await Evaluation.findAll({
            where: whereClause,
            include: [
                { model: User, as: 'employee', attributes: ['id', 'username', 'email'] },
                { model: User, as: 'evaluator', attributes: ['id', 'username'] }
            ],
            order: [['createdAt', 'DESC']]
        });

        res.json(evaluations);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getEvaluationById = async (req, res) => {
    try {
        const evaluation = await Evaluation.findByPk(req.params.id, {
            include: [
                { model: User, as: 'employee', attributes: ['id', 'username'] },
                { model: User, as: 'evaluator', attributes: ['id', 'username'] },
                {
                    model: EvaluationDetail,
                    as: 'details',
                    include: [{ model: KPI, as: 'kpi', attributes: ['id', 'title', 'description', 'weight'] }]
                }
            ]
        });

        if (!evaluation) return res.status(404).json({ message: 'Evaluation not found' });

        res.json(evaluation);
    } catch (error) {
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};

exports.getDashboardStats = async (req, res) => {
    try {
        const { User, KPI, EvaluationDetail, Evaluation, QuantitativeLog } = require('../models');
        const role = req.user.role;
        const userId = req.user.userId;

        // --- Calculate Employee of the Month (Composite Score) ---
        // 1. Determine most recent active period to calculate EOTM for
        const latestQuant = await QuantitativeLog.findOne({ order: [['createdAt', 'DESC']] });
        const latestEval = await Evaluation.findOne({ order: [['createdAt', 'DESC']] });

        let targetPeriod = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
        if (latestQuant && latestEval) {
            targetPeriod = (new Date(latestQuant.createdAt) > new Date(latestEval.createdAt)) ? latestQuant.period : latestEval.period;
        } else if (latestQuant) targetPeriod = latestQuant.period;
        else if (latestEval) targetPeriod = latestEval.period;

        const employees = await User.findAll({
            where: { role: 'Employee' },
            attributes: ['id', 'username', 'profilePicture', 'position']
        });

        let allScores = [];

        for (const emp of employees) {
            // 1. Manager Score (Avg Rating / 5 * 100)
            const managerEvals = await EvaluationDetail.findAll({
                include: [{
                    model: Evaluation,
                    where: { employeeId: emp.id, type: 'Manager', period: targetPeriod },
                    attributes: []
                }],
                attributes: [[require('sequelize').fn('AVG', require('sequelize').col('rating')), 'avgRating']],
                raw: true
            });
            const managerAvg = managerEvals[0]?.avgRating ? parseFloat(managerEvals[0].avgRating) : 0;
            const managerScore = (managerAvg / 5) * 100;

            // 2. Quantitative Score (Weighted Calculation)
            const quantLogs = await QuantitativeLog.findAll({
                where: { employeeId: emp.id, period: targetPeriod },
                include: [{ model: KPI }]
            });

            let quantScore = 0;
            if (quantLogs.length > 0) {
                let totalWeightedScore = 0;
                let totalWeights = 0;

                quantLogs.forEach(log => {
                    const target = (log.KPI && log.KPI.targetValue > 0) ? log.KPI.targetValue : 1;
                    const actual = log.actualValue || 0;
                    const weight = (log.KPI && log.KPI.weight > 0) ? log.KPI.weight : 0;

                    const score = (actual / target) * 100;
                    totalWeightedScore += (score * weight);
                    totalWeights += weight;
                });

                // If weighting is not configured properly, default to a simple average
                quantScore = totalWeights > 0 ? (totalWeightedScore / totalWeights) :
                    (quantLogs.reduce((acc, log) => acc + (((log.actualValue || 0) / ((log.KPI && log.KPI.targetValue > 0) ? log.KPI.targetValue : 1)) * 100), 0) / quantLogs.length);
            }

            // Composite Score logic
            let compositeScore = 0;
            if (managerScore > 0 && quantScore > 0) compositeScore = (managerScore + quantScore) / 2;
            else if (managerScore > 0) compositeScore = managerScore;
            else if (quantScore > 0) compositeScore = quantScore;

            if (compositeScore > 0) {
                allScores.push({
                    id: emp.id,
                    username: emp.username,
                    profilePicture: emp.profilePicture,
                    position: emp.position,
                    compositeScore,
                    avgRating: (compositeScore / 20).toFixed(1)
                });
            }
        }

        let employeeOfTheMonth = null;
        let underperformingEmployees = [];

        if (allScores.length > 0) {
            allScores.sort((a, b) => b.compositeScore - a.compositeScore);
            employeeOfTheMonth = allScores[0];

            // Score < 50% is underperforming
            underperformingEmployees = allScores.filter(e => e.compositeScore < 50).reverse();
        }

        // --- Return Data Based on Role ---
        if (role === 'Employee') {
            const totalEvaluations = await Evaluation.count({ where: { employeeId: userId } });
            const totalKPIs = await KPI.count();

            const aggregations = await EvaluationDetail.findAll({
                include: [{ model: Evaluation, where: { employeeId: userId }, attributes: [] }],
                attributes: [[require('sequelize').fn('AVG', require('sequelize').col('rating')), 'avgRating']],
                raw: true
            });
            const avgRating = aggregations[0]?.avgRating ? parseFloat(aggregations[0].avgRating) : 0;
            const avgPerformance = Math.round((avgRating / 5) * 100);

            const distribution = await EvaluationDetail.findAll({
                include: [{ model: Evaluation, where: { employeeId: userId }, attributes: [] }],
                attributes: ['rating', [require('sequelize').fn('COUNT', require('sequelize').col('rating')), 'count']],
                group: ['rating'],
                order: [['rating', 'ASC']],
                raw: true
            });
            const ratingDistribution = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
            distribution.forEach(d => { ratingDistribution[d.rating] = parseInt(d.count); });

            return res.json({ totalStaff: totalEvaluations, totalKPIs, avgPerformance, ratingDistribution, employeeOfTheMonth });
        }

        // --- Admin / Manager Stats ---
        const totalStaff = await User.count({ where: { role: 'Employee' } });
        const totalKPIs = await KPI.count();

        const aggregations = await EvaluationDetail.findAll({
            attributes: [[require('sequelize').fn('AVG', require('sequelize').col('rating')), 'avgRating']],
            raw: true
        });
        const avgRating = aggregations[0]?.avgRating ? parseFloat(aggregations[0].avgRating) : 0;
        const avgPerformance = Math.round((avgRating / 5) * 100);

        const distribution = await EvaluationDetail.findAll({
            attributes: ['rating', [require('sequelize').fn('COUNT', require('sequelize').col('rating')), 'count']],
            group: ['rating'],
            order: [['rating', 'ASC']],
            raw: true
        });
        const ratingDistribution = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
        distribution.forEach(d => { ratingDistribution[d.rating] = parseInt(d.count); });

        res.json({ totalStaff, totalKPIs, avgPerformance, ratingDistribution, employeeOfTheMonth, underperformingEmployees });
    } catch (error) {
        console.error("Dashboard Stats Error:", error);
        res.status(500).json({ message: 'Server error', error: error.message });
    }
};
