import React, { useEffect } from 'react';
import { Outlet, NavLink, useLocation, useNavigate } from 'react-router-dom';
import { Settings, Calculator } from 'lucide-react';

const Bonus = () => {
    const location = useLocation();
    const navigate = useNavigate();

    // Redirect root /admin/bonus to configuration
    useEffect(() => {
        if (location.pathname === '/admin/bonus' || location.pathname === '/admin/bonus/') {
            navigate('configuration', { replace: true });
        }
    }, [location.pathname, navigate]);

    const tabs = [
        { path: 'configuration', label: 'Configuration', icon: Settings },
        { path: 'calculation', label: 'Calculation', icon: Calculator },
    ];

    return (
        <div className="flex flex-col h-full bg-gray-50">
            {/* Header Tabs */}
            <div className="bg-white border-b border-gray-200 px-6 pt-4">
                <div className="flex space-x-1">
                    {tabs.map(tab => (
                        <NavLink
                            key={tab.path}
                            to={tab.path}
                            className={({ isActive }) =>
                                `flex items-center px-4 py-3 text-sm font-medium rounded-t-lg border-b-2 transition-colors ${isActive
                                    ? 'border-blue-500 text-blue-600 bg-blue-50'
                                    : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                                }`
                            }
                        >
                            <tab.icon className="w-4 h-4 mr-2" />
                            {tab.label}
                        </NavLink>
                    ))}
                </div>
            </div>

            {/* Content Content - Rendered via Outlet */}
            <div className="flex-1 overflow-auto">
                <Outlet />
            </div>
        </div>
    );
};

export default Bonus;
