import React from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import MainLayout from './layouts/MainLayout';
import Login from './pages/auth/Login';
import Register from './pages/auth/Register';
import ForgotPassword from './pages/auth/ForgotPassword';
import ResetPassword from './pages/auth/ResetPassword';
import Dashboard from './pages/common/Dashboard';
import PlaceholderPage from './pages/common/PlaceholderPage';
import Policies from './pages/common/Policies';
import EmployeeKPITracking from './pages/admin/EmployeeKPITracking';
import KPIFramework from './pages/admin/KPIFramework';
import ManagerDataEntry from './pages/admin/ManagerDataEntry';
import StaffManagement from './pages/admin/StaffManagement';
import StaffProfile from './pages/admin/StaffProfile';
import Analytics from './pages/admin/Analytics';
import SelfEvaluation from './pages/evaluation/SelfEvaluation';
import ManagerEvaluation from './pages/evaluation/ManagerEvaluation';
import PeerEvaluation from './pages/evaluation/PeerEvaluation';
import Evaluation360 from './pages/evaluation/Evaluation360';
import BonusConfiguration from './pages/admin/BonusConfiguration';
import BonusCalculation from './pages/admin/BonusCalculation';
import Bonus from './pages/admin/Bonus';
import Reports from './pages/admin/Reports';
import AuditLogs from './pages/admin/AuditLogs';

import MyBonuses from './pages/user/MyBonuses';
import MyEvaluations from './pages/user/MyEvaluations';
import MyKPIs from './pages/user/MyKPIs';
import TeamOverview from './pages/user/TeamOverview';
import ProfileSettings from './pages/common/ProfileSettings';
import { AuthProvider } from './context/AuthContext';
import PrivateRoute from './components/PrivateRoute';

function App() {
  return (
    <AuthProvider>
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<Navigate to="/login" replace />} />
          <Route path="/login" element={<Login />} />
          <Route path="/register" element={<Register />} />
          <Route path="/forgot-password" element={<ForgotPassword />} />
          <Route path="/reset-password/:token" element={<ResetPassword />} />

          {/* Admin Routes - Protected */}
          <Route element={<PrivateRoute allowedRoles={['Admin', 'Manager']} />}>
            <Route path="/admin" element={<MainLayout role="admin" />}>
              <Route path="dashboard" element={<Dashboard role="Admin" />} />
              <Route path="staff" element={<StaffManagement />} />
              <Route path="staff/:id" element={<StaffProfile />} />
              <Route path="kpi-framework" element={<KPIFramework />} />
              <Route path="employee-tracking" element={<EmployeeKPITracking />} />
              <Route path="data-entry" element={<ManagerDataEntry />} />
              <Route path="360-evaluation" element={<Evaluation360 />} />
              <Route path="manager-evaluation" element={<ManagerEvaluation />} />
              <Route path="self-evaluation" element={<SelfEvaluation />} />
              <Route path="peer-evaluation" element={<PeerEvaluation />} />
              <Route path="bonus" element={<Bonus />}>
                <Route path="configuration" element={<BonusConfiguration />} />
                <Route path="calculation" element={<BonusCalculation />} />
                <Route index element={<Navigate to="configuration" replace />} />
              </Route>
              <Route path="policies" element={<Policies />} />
              <Route path="reports" element={<Reports />} />
              <Route path="analytics" element={<Analytics />} />
              <Route path="audit-logs" element={<AuditLogs />} />
              <Route path="profile" element={<ProfileSettings />} />
              <Route index element={<Navigate to="dashboard" replace />} />
            </Route>
          </Route>

          {/* User Routes - Protected */}
          <Route element={<PrivateRoute allowedRoles={['Employee']} />}>
            <Route path="/user" element={<MainLayout role="user" />}>
              <Route path="dashboard" element={<Dashboard role="Employee" />} />
              <Route path="my-kpis" element={<MyKPIs />} />
              <Route path="my-bonuses" element={<MyBonuses />} />
              <Route path="evaluations" element={<MyEvaluations />} />
              <Route path="360-evaluation" element={<Evaluation360 />} />
              <Route path="self-evaluation" element={<SelfEvaluation />} />
              <Route path="peer-evaluation" element={<PeerEvaluation />} />
              <Route path="policies" element={<Policies />} />
              <Route path="team" element={<TeamOverview />} />
              <Route path="profile" element={<ProfileSettings />} />
              <Route index element={<Navigate to="dashboard" replace />} />
            </Route>
          </Route>

        </Routes>
      </BrowserRouter>
    </AuthProvider>
  );
}

export default App;
