const jwt = require('jsonwebtoken');

const verifyToken = (req, res, next) => {
    const token = req.headers['authorization']?.split(' ')[1];

    if (!token) {
        return res.status(403).json({ message: 'No token provided' });
    }

    jwt.verify(token, process.env.JWT_SECRET || 'techznap_secret_key', (err, decoded) => {
        if (err) {
            return res.status(401).json({ message: 'Unauthorized' });
        }
        req.user = decoded; // { userId, role, ... }
        next();
    });
};

const isAdmin = (req, res, next) => {
    if (req.user.role !== 'Admin') {
        return res.status(403).json({ message: 'Require Admin Role' });
    }
    next();
};

const isAdminOrManager = (req, res, next) => {
    if (req.user.role !== 'Admin' && req.user.role !== 'Manager') {
        return res.status(403).json({ message: 'Require Admin or Manager Role' });
    }
    next();
};

module.exports = {
    verifyToken,
    isAdmin,
    isAdminOrManager
};
