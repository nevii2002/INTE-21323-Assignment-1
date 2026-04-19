const { Policy, User } = require('../models');
const fs = require('fs');
const path = require('path');

// @desc    Upload a new policy
// @route   POST /api/policies/upload
// @access  Admin/Manager
exports.uploadPolicy = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: 'No file uploaded' });
        }

        const { title } = req.body;
        if (!title) {
            // Clean up the uploaded file if title is missing
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ message: 'Title is required' });
        }

        const policy = await Policy.create({
            title,
            fileName: req.file.originalname,
            filePath: req.file.path.replace(/\\/g, '/'), // Normalize path for windows
            uploadedBy: req.user.userId
        });

        res.status(201).json({
            message: 'Policy uploaded successfully',
            policy
        });
    } catch (error) {
        if (req.file) {
            fs.unlinkSync(req.file.path);
        }
        console.error('Upload policy error:', error);
        res.status(500).json({ message: 'Server error while uploading policy' });
    }
};

// @desc    Get all policies
// @route   GET /api/policies
// @access  Private
exports.getPolicies = async (req, res) => {
    try {
        const policies = await Policy.findAll({
            include: [{ model: User, as: 'uploader', attributes: ['username', 'email'] }],
            order: [['createdAt', 'DESC']]
        });

        res.json(policies);
    } catch (error) {
        console.error('Get policies error:', error);
        res.status(500).json({ message: 'Server error while fetching policies' });
    }
};

// @desc    Delete a policy
// @route   DELETE /api/policies/:id
// @access  Admin/Manager
exports.deletePolicy = async (req, res) => {
    try {
        const policy = await Policy.findByPk(req.params.id);

        if (!policy) {
            return res.status(404).json({ message: 'Policy not found' });
        }

        // Delete file from filesystem
        const fullPath = path.join(__dirname, '..', policy.filePath);
        if (fs.existsSync(fullPath)) {
            fs.unlinkSync(fullPath);
        }

        await policy.destroy();

        res.json({ message: 'Policy deleted successfully' });
    } catch (error) {
        console.error('Delete policy error:', error);
        res.status(500).json({ message: 'Server error while deleting policy' });
    }
};
