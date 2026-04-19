import React from 'react';
import { Outlet, NavLink, useNavigate, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import {
    LayoutDashboard,
    Users,
    Settings,
    FileText,
    BarChart2,
    LogOut,
    Award,
    BookOpen,
    UserCheck,
    DollarSign,
    Shield
} from 'lucide-react';

const MainLayout = ({ role = 'user' }) => {
    const navigate = useNavigate();
    const location = useLocation();
    const { user, logout } = useAuth();

    const handleLogout = () => {
        logout();
        navigate('/login');
    };

    const adminLinks = [
        { path: '/admin/dashboard', label: 'Dashboard', icon: LayoutDashboard },
        { path: '/admin/staff', label: 'Staff Management', icon: Users },
        { path: '/admin/kpi-framework', label: 'KPI Framework', icon: Settings },
        { path: '/admin/employee-tracking', label: 'Employee Logs', icon: Users },
        { path: '/admin/data-entry', label: 'Company Data', icon: BarChart2 },
        { path: '/admin/360-evaluation', label: '360 Evaluation', icon: FileText },
        { path: '/admin/bonus', label: 'Bonus', icon: Award },
        { path: '/admin/policies', label: 'HR Policies', icon: BookOpen },
        { path: '/admin/reports', label: 'Reports', icon: FileText },
        { path: '/admin/audit-logs', label: 'Audit Logs', icon: Shield },
        { path: '/admin/analytics', label: 'Analytics', icon: BarChart2 },
        { path: '/admin/profile', label: 'Profile Settings', icon: Settings },
    ];

    const userLinks = [
        { path: '/user/dashboard', label: 'Dashboard', icon: LayoutDashboard },
        { path: '/user/my-kpis', label: 'My KPIs', icon: Award },
        { path: '/user/my-bonuses', label: 'My Bonuses', icon: DollarSign },
        { path: '/user/evaluations', label: 'My Evaluations', icon: FileText },
        { path: '/user/360-evaluation', label: '360 Evaluation', icon: UserCheck },
        { path: '/user/policies', label: 'HR Policies', icon: BookOpen },
        { path: '/user/team', label: 'Team Overview', icon: Users },
        { path: '/user/profile', label: 'Profile Settings', icon: Settings },
    ];

    const links = role === 'admin' ? adminLinks : userLinks;

    return (
        <div className="flex h-screen bg-gray-50">
            {/* Sidebar */}
            <aside className="w-64 bg-white border-r border-gray-200 flex flex-col">
                <div className="p-6 border-b border-gray-100 flex flex-col items-center">
                    <img src="/logo.png" alt="Techznap Logo" className="h-10 mb-2" />
                    <h1 className="text-lg font-bold text-gray-800 text-center leading-tight">Employee Performance Evaluation System</h1>
                    <p className="text-[10px] text-blue-500 mt-2 uppercase tracking-wider">{role} Portal</p>
                </div>

                <nav className="flex-1 p-4 space-y-1 overflow-y-auto">
                    {links.map((link) => {
                        const Icon = link.icon;
                        const isActive = location.pathname.startsWith(link.path);
                        return (
                            <NavLink
                                key={link.path}
                                to={link.path}
                                className={({ isActive }) =>
                                    `flex items-center px-4 py-3 text-sm font-medium rounded-lg transition-colors ${isActive
                                        ? 'bg-blue-50 text-blue-600'
                                        : 'text-gray-600 hover:bg-gray-50 hover:text-gray-900'
                                    }`
                                }
                            >
                                <Icon className={`w-5 h-5 mr-3 ${isActive ? 'text-blue-500' : 'text-gray-400'}`} />
                                {link.label}
                            </NavLink>
                        );
                    })}
                </nav>

                <div className="p-4 border-t border-gray-100">
                    <button
                        onClick={handleLogout}
                        className="flex items-center w-full px-4 py-3 text-sm font-medium text-red-600 rounded-lg hover:bg-red-50 transition-colors"
                    >
                        <LogOut className="w-5 h-5 mr-3" />
                        Sign Out
                    </button>
                </div>
            </aside>

            {/* Main Content */}
            <main className="flex-1 overflow-y-auto">
                <header className="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-8">
                    <h2 className="text-lg font-semibold text-gray-800">
                        {links.find(l => location.pathname.startsWith(l.path))?.label || 'Dashboard'}
                    </h2>
                    <div className="flex items-center space-x-4">
                        <button
                            onClick={() => navigate(`/${role}/profile`)}
                            title="Profile Settings"
                            className="w-8 h-8 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 font-bold hover:bg-blue-200 focus:outline-none transition-colors cursor-pointer overflow-hidden shadow-sm"
                        >
                            {user?.profilePicture ? (
                                <img src={`http://localhost:5000${user.profilePicture}`} alt="Avatar" className="w-full h-full object-cover" />
                            ) : (
                                role === 'admin' ? 'A' : 'U'
                            )}
                        </button>
                    </div>
                </header>
                <div className="p-8">
                    <Outlet />
                </div>
            </main>
        </div>
    );
};

export default MainLayout;
