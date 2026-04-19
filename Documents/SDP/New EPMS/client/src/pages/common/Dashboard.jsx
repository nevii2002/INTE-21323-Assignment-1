import React, { useState, useEffect } from 'react';
import api from '../../api/axios';

import { Users, Activity, Award } from 'lucide-react';

const Dashboard = ({ role }) => {
    const [stats, setStats] = useState({ totalStaff: 0, totalKPIs: 0, avgPerformance: 0 });
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        const fetchStats = async () => {
            try {
                // If role is user, maybe fetch different stats?
                // For now, assuming this endpoint returns general stats used by Admin.
                // If user, we might want "My Targets" instead of "Total Employees".
                const response = await api.get('/evaluations/stats');
                setStats(response.data);
            } catch (error) {
                console.error('Failed to fetch dashboard stats', error);
            } finally {
                setLoading(false);
            }
        };

        fetchStats();
    }, [role]);

    return (
        <div>
            <h3 className="text-xl font-medium text-gray-900 mb-6">Overview</h3>

            <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 lg:grid-cols-3 mb-8">
                {/* Card 1 */}
                <div className="bg-white overflow-hidden shadow rounded-lg">
                    <div className="p-5">
                        <div className="flex items-center">
                            <div className="flex-shrink-0 bg-blue-500 rounded-md p-3">
                                <Users className="h-6 w-6 text-white" />
                            </div>
                            <div className="ml-5 w-0 flex-1">
                                <dl>
                                    <dt className="text-sm font-medium text-gray-500 truncate">
                                        {role === 'Admin' || role === 'Manager' ? 'Total Employees' : 'My Evaluations'}
                                    </dt>
                                    <dd>
                                        <div className="text-lg font-medium text-gray-900">
                                            {loading ? '...' : stats.totalStaff}
                                        </div>
                                    </dd>
                                </dl>
                            </div>
                        </div>
                    </div>
                </div>

                {/* Card 2 */}
                <div className="bg-white overflow-hidden shadow rounded-lg">
                    <div className="p-5">
                        <div className="flex items-center">
                            <div className="flex-shrink-0 bg-green-500 rounded-md p-3">
                                <Activity className="h-6 w-6 text-white" />
                            </div>
                            <div className="ml-5 w-0 flex-1">
                                <dl>
                                    <dt className="text-sm font-medium text-gray-500 truncate">Avg. Performance</dt>
                                    <dd>
                                        <div className="text-lg font-medium text-gray-900">
                                            {loading ? '...' : `${stats.avgPerformance}%`}
                                        </div>
                                    </dd>
                                </dl>
                            </div>
                        </div>
                    </div>
                </div>

                {/* Card 3 */}
                <div className="bg-white overflow-hidden shadow rounded-lg">
                    <div className="p-5">
                        <div className="flex items-center">
                            <div className="flex-shrink-0 bg-purple-500 rounded-md p-3">
                                <Award className="h-6 w-6 text-white" />
                            </div>
                            <div className="ml-5 w-0 flex-1">
                                <dl>
                                    <dt className="text-sm font-medium text-gray-500 truncate">Active KPIs</dt>
                                    <dd>
                                        <div className="text-lg font-medium text-gray-900">
                                            {loading ? '...' : stats.totalKPIs}
                                        </div>
                                    </dd>
                                </dl>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            {/* Employee of the Month Banner */}
            {!loading && stats.employeeOfTheMonth && (
                <div className="bg-gradient-to-r from-yellow-400 via-yellow-500 to-orange-500 shadow-lg rounded-lg p-6 mb-8 flex flex-col sm:flex-row items-center justify-between transform transition-all hover:scale-[1.01] overflow-hidden relative">
                    {/* Decorative Background Elements */}
                    <div className="absolute top-0 right-0 -mt-4 -mr-4 w-32 h-32 bg-white opacity-10 rounded-full blur-2xl"></div>
                    <div className="absolute bottom-0 left-10 -mb-4 w-24 h-24 bg-white opacity-10 rounded-full blur-xl"></div>

                    <div className="flex items-center mb-4 sm:mb-0 relative z-10 w-full sm:w-auto justify-center sm:justify-start">
                        <div className="relative flex-shrink-0">
                            <div className="h-20 w-20 rounded-full border-4 border-white shadow-md overflow-hidden bg-white flex items-center justify-center">
                                {stats.employeeOfTheMonth.profilePicture ? (
                                    <img src={`http://localhost:5000${stats.employeeOfTheMonth.profilePicture}`} alt="Avatar" className="h-full w-full object-cover" />
                                ) : (
                                    <span className="text-3xl font-bold text-yellow-500">{stats.employeeOfTheMonth?.username?.charAt(0).toUpperCase() || 'E'}</span>
                                )}
                            </div>
                            <div className="absolute -bottom-2 -right-2 bg-white rounded-full p-1.5 shadow-md">
                                <Award className="h-6 w-6 text-yellow-500" />
                            </div>
                        </div>
                        <div className="ml-6 text-white text-center sm:text-left">
                            <h2 className="text-xs font-bold uppercase tracking-widest text-yellow-100 mb-1 flex items-center justify-center sm:justify-start">
                                <Award className="h-4 w-4 mr-1 inline" /> Employee of the Month
                            </h2>
                            <h3 className="text-2xl font-bold drop-shadow-sm">{stats.employeeOfTheMonth.username}</h3>
                            <p className="text-yellow-50 font-medium mt-1">{stats.employeeOfTheMonth.position || 'Valued Team Member'}</p>
                        </div>
                    </div>
                    <div className="bg-white/20 backdrop-blur-md rounded-xl p-4 text-center text-white border border-white/40 shadow-inner relative z-10 mt-4 sm:mt-0 w-full sm:w-48">
                        <div className="text-xs font-semibold uppercase tracking-wider text-yellow-100 mb-1">Average Rating</div>
                        <div className="text-3xl font-bold flex items-center justify-center drop-shadow-md">
                            {stats.employeeOfTheMonth.avgRating} <span className="text-base font-medium ml-1 text-yellow-100 mt-1">/ 5.0</span>
                        </div>
                    </div>
                </div>
            )}

            {/* Underperforming Employees Banner (Admin & Manager Only) */}
            {!loading && (role === 'Admin' || role === 'Manager') && stats.underperformingEmployees && stats.underperformingEmployees.length > 0 && (
                <div className="bg-white shadow rounded-lg p-6 mb-8 border-l-4 border-red-500">
                    <h4 className="text-lg font-bold text-gray-900 mb-4 flex items-center">
                        <Activity className="w-5 h-5 mr-2 text-red-500" />
                        Needs Attention (Underperforming Employees)
                    </h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                        {stats.underperformingEmployees.map(emp => (
                            <div key={emp.id} className="border border-gray-200 rounded-lg p-4 flex items-center bg-red-50/30">
                                <div className="h-12 w-12 rounded-full border-2 border-red-100 bg-white flex items-center justify-center overflow-hidden mr-4">
                                    {emp.profilePicture ? (
                                        <img src={`http://localhost:5000${emp.profilePicture}`} alt="Avatar" className="h-full w-full object-cover" />
                                    ) : (
                                        <span className="font-bold text-red-500">{emp.username?.charAt(0).toUpperCase()}</span>
                                    )}
                                </div>
                                <div>
                                    <h5 className="font-semibold text-gray-900">{emp.username}</h5>
                                    <p className="text-xs text-red-600 font-medium pb-1">{emp.position || 'Employee'}</p>
                                    <span className="px-2 py-0.5 rounded bg-red-100 text-red-800 text-[10px] font-bold">
                                        Score: {emp.compositeScore.toFixed(1)}% | {emp.avgRating}/5
                                    </span>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>
            )}

            {/* Analytics Section - CSS Bar Chart */}
            <div className="bg-white shadow rounded-lg p-6">
                <h4 className="text-lg font-medium text-gray-900 mb-4">
                    {role === 'Admin' || role === 'Manager' ? 'Organization Performance Distribution' : 'My Performance Distribution'}
                </h4>
                {loading ? (
                    <div className="text-center text-gray-400 py-10">Loading analytics...</div>
                ) : (
                    <div className="space-y-4">
                        {[5, 4, 3, 2, 1].map((star) => {
                            const count = stats.ratingDistribution ? stats.ratingDistribution[star] : 0;
                            // Calculate simple max for scaling (avoid /0)
                            const maxVal = Math.max(...Object.values(stats.ratingDistribution || { 0: 1 }), 1);
                            const widthPercent = (count / maxVal) * 100;

                            return (
                                <div key={star} className="flex items-center">
                                    <div className="w-16 text-sm text-gray-600 font-medium">{star} Stars</div>
                                    <div className="flex-1 h-4 bg-gray-100 rounded-full mx-2 overflow-hidden">
                                        <div
                                            className={`h-full rounded-full ${star >= 4 ? 'bg-green-500' : star === 3 ? 'bg-yellow-400' : 'bg-red-500'
                                                }`}
                                            style={{ width: `${widthPercent}%`, transition: 'width 1s ease-in-out' }}
                                        ></div>
                                    </div>
                                    <div className="w-12 text-sm text-gray-500 text-right">{count}</div>
                                </div>
                            );
                        })}
                    </div>
                )}
                <p className="mt-4 text-xs text-center text-gray-400">
                    {role === 'Admin' || role === 'Manager' ? 'Distribution of all evaluation scores across the organization.' : 'Distribution of all your personal evaluation scores.'}
                </p>
            </div>
        </div>
    );
};

export default Dashboard;
