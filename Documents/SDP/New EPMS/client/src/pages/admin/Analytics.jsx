import React, { useState, useEffect } from 'react';
import axios from 'axios';
import {
    PieChart, Pie, Cell, Tooltip, Legend,
    BarChart, Bar, XAxis, YAxis, CartesianGrid, ResponsiveContainer
} from 'recharts';
import { BarChart2, PieChart as PieChartIcon, Activity, TrendingUp } from 'lucide-react';

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d'];

const Analytics = () => {
    const [data, setData] = useState({
        jobCategories: [],
        roles: [],
        evaluationStatuses: [],
        bonusesByReason: []
    });
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        fetchAnalyticsData();
    }, []);

    const fetchAnalyticsData = async () => {
        try {
            const res = await axios.get('http://localhost:5000/api/analytics/dashboard', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setData(res.data);
            setLoading(false);
        } catch (error) {
            console.error("Error fetching analytics data:", error);
            setLoading(false);
        }
    };

    if (loading) {
        return (
            <div className="flex justify-center items-center h-full min-h-[400px]">
                <Activity className="animate-spin text-blue-600 w-12 h-12" />
            </div>
        );
    }

    // Prepare data for UI
    const jobData = data.jobCategories.map(item => ({ name: item.jobCategory, value: Number(item.count) }));
    const roleData = data.roles.map(item => ({ name: item.role, value: Number(item.count) }));
    const evalData = data.evaluationStatuses.map(item => ({ name: item.status, count: Number(item.count) }));
    const bonusData = data.bonusesByReason.map(item => ({ name: item.reason, amount: Number(item.totalAmount) }));

    return (
        <div className="p-6 bg-gray-50 min-h-screen">
            <h2 className="text-2xl font-bold text-gray-800 mb-6 flex items-center">
                <BarChart2 className="mr-2 text-blue-600" />
                System Analytics
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                {/* Job Categories Pie Chart */}
                <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
                        <PieChartIcon className="w-5 h-5 mr-2 text-indigo-500" />
                        Employees by Job Category
                    </h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie
                                    data={jobData}
                                    cx="50%"
                                    cy="50%"
                                    labelLine={false}
                                    outerRadius={80}
                                    fill="#8884d8"
                                    dataKey="value"
                                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                                >
                                    {jobData.map((entry, index) => (
                                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                    ))}
                                </Pie>
                                <Tooltip />
                                <Legend />
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Roles Pie Chart */}
                <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
                        <PieChartIcon className="w-5 h-5 mr-2 text-green-500" />
                        Users by Access Role
                    </h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie
                                    data={roleData}
                                    cx="50%"
                                    cy="50%"
                                    outerRadius={80}
                                    fill="#82ca9d"
                                    dataKey="value"
                                    label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                                >
                                    {roleData.map((entry, index) => (
                                        <Cell key={`cell-${index}`} fill={COLORS[(index + 2) % COLORS.length]} />
                                    ))}
                                </Pie>
                                <Tooltip />
                                <Legend />
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Evaluation Status Bar Chart */}
                <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
                        <BarChart2 className="w-5 h-5 mr-2 text-blue-500" />
                        Performance Evaluations Status
                    </h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart
                                data={evalData}
                                margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                            >
                                <CartesianGrid strokeDasharray="3 3" />
                                <XAxis dataKey="name" />
                                <YAxis allowDecimals={false} />
                                <Tooltip />
                                <Legend />
                                <Bar dataKey="count" name="Evaluations" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>

                {/* Bonuses Bar Chart */}
                <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
                        <TrendingUp className="w-5 h-5 mr-2 text-yellow-500" />
                        Total Bonus Allocation by Reason
                    </h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart
                                data={bonusData}
                                margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                            >
                                <CartesianGrid strokeDasharray="3 3" />
                                <XAxis dataKey="name" />
                                <YAxis />
                                <Tooltip formatter={(value) => `LKR ${value.toLocaleString()}`} />
                                <Legend />
                                <Bar dataKey="amount" name="Total Amount (LKR)" fill="#f59e0b" radius={[4, 4, 0, 0]} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default Analytics;
