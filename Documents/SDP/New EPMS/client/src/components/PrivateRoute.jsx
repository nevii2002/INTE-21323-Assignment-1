import React from 'react';
import { Navigate, Outlet } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';

const PrivateRoute = ({ allowedRoles }) => {
    const { user } = useAuth();

    if (!user) {
        return <Navigate to="/login" replace />;
    }

    if (allowedRoles && !allowedRoles.includes(user.role)) {
        // Redirect to appropriate dashboard based on actual role if trying to access unauthorized route
        if (user.role === 'Admin' || user.role === 'Manager') {
            return <Navigate to="/admin/dashboard" replace />;
        } else {
            return <Navigate to="/user/dashboard" replace />;
        }
    }

    return <Outlet />;
};

export default PrivateRoute;
