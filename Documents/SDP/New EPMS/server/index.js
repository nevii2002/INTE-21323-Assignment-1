const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const path = require('path');
const sequelize = require('./config/database');
const { User, Evaluation, EvaluationDetail } = require('./models');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

const authRoutes = require('./routes/authRoutes');
const kpiRoutes = require('./routes/kpiRoutes');
const staffRoutes = require('./routes/staffRoutes');
const evaluationRoutes = require('./routes/evaluationRoutes');
const bonusRoutes = require('./routes/bonusRoutes');
const policyRoutes = require('./routes/policyRoutes');
const analyticsRoutes = require('./routes/analyticsRoutes');
const quantitativeRoutes = require('./routes/quantitativeRoutes');
const auditRoutes = require('./routes/auditRoutes');
const companyDataRoutes = require('./routes/companyDataRoutes');
const helmet = require('helmet');

app.use(cors());
app.use(express.json());

// Security Headers (Allow iframe embedding from frontend)
app.use((req, res, next) => {
  res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
  res.setHeader('Content-Security-Policy', "frame-ancestors 'self' http://localhost:3000");
  res.setHeader('X-Frame-Options', 'ALLOW-FROM http://localhost:3000');
  next();
});
// Serve uploaded files statically
app.use(express.static(path.join(__dirname, 'uploads')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

app.use('/api/auth', authRoutes);
app.use('/api/kpis', kpiRoutes);
app.use('/api/staff', staffRoutes);
app.use('/api/evaluations', evaluationRoutes);
app.use('/api/bonuses', bonusRoutes);
app.use('/api/policies', policyRoutes);
app.use('/api/analytics', analyticsRoutes);
app.use('/api/quantitative', quantitativeRoutes);
app.use('/api/audit-logs', auditRoutes);
app.use('/api/company-data', companyDataRoutes);

app.get('/', (req, res) => {
  res.send('SEPES API is running');
});

// Database synchronization and Server Start
sequelize.sync({ alter: false })
  .then(() => {
    console.log('Database connected and synced');
    app.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  })
  .catch(err => {
    console.error('Unable to connect to the database:', err);
  });
