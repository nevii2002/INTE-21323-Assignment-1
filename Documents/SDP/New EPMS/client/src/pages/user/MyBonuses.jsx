import React, { useState, useEffect } from 'react';
import api from '../../api/axios';
import { useAuth } from '../../context/AuthContext';
import { Award, DollarSign, Calendar, TrendingUp } from 'lucide-react';

const MyBonuses = () => {
    const { user } = useAuth();
    const [targets, setTargets] = useState([]);
    const [history, setHistory] = useState([]);
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        const fetchData = async () => {
            if (!user) return;
            try {
                // Fetch KPI Targets (Potential Bonus)
                const kpiResponse = await api.get(`/staff/${user.id}/kpis`);
                setTargets(kpiResponse.data);

                // Fetch Bonus History (Actual Earned)
                const historyResponse = await api.get(`/bonuses/employee/${user.id}`);
                setHistory(historyResponse.data);
            } catch (error) {
                console.error("Failed to fetch bonus data", error);
            } finally {
                setIsLoading(false);
            }
        };
        fetchData();
    }, [user]);

    const totalPotential = targets.reduce((sum, t) => sum + (t.EmployeeKPI?.customBonus || 0), 0);
    const totalEarned = history.reduce((sum, h) => sum + parseFloat(h.amount), 0);

    if (isLoading) return <div className="p-8 text-center text-gray-500">Loading bonus data...</div>;

    return (
        <div className="space-y-8">
            {/* Header Stats */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-gradient-to-r from-blue-500 to-blue-600 rounded-lg shadow-lg p-6 text-white">
                    <div className="flex items-center justify-between">
                        <div>
                            <p className="text-blue-100 text-sm font-medium uppercase tracking-wider">Total Earned Bonus</p>
                            <h3 className="text-3xl font-bold mt-1">Rs. {totalEarned.toLocaleString()}</h3>
                        </div>
                        <div className="p-3 bg-white/20 rounded-full">
                            <DollarSign className="w-8 h-8 text-white" />
                        </div>
                    </div>
                </div>

                <div className="bg-white rounded-lg shadow p-6 border border-gray-100">
                    <div className="flex items-center justify-between">
                        <div>
                            <p className="text-gray-500 text-sm font-medium uppercase tracking-wider">Potential Bonus (Targets)</p>
                            <h3 className="text-3xl font-bold mt-1 text-gray-800">Rs. {totalPotential.toLocaleString()}</h3>
                        </div>
                        <div className="p-3 bg-green-50 rounded-full">
                            <TrendingUp className="w-8 h-8 text-green-600" />
                        </div>
                    </div>
                </div>
            </div>

            {/* Bonus Targets Section */}
            <div className="bg-white shadow rounded-lg overflow-hidden border border-gray-200">
                <div className="px-6 py-4 border-b border-gray-200 bg-gray-50">
                    <h3 className="text-lg font-medium text-gray-900 flex items-center">
                        <Award className="w-5 h-5 mr-2 text-blue-600" />
                        Bonus Targets
                    </h3>
                    <p className="text-xs text-gray-500 mt-1">Achieve these targets to unlock bonuses.</p>
                </div>
                <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Project / KPI</th>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Category</th>
                                <th className="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Target</th>
                                {/* WEIGHT HIDDEN AS REQUESTED */}
                                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Potential Bonus</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                            {targets.length === 0 ? (
                                <tr><td colSpan="4" className="px-6 py-4 text-center text-gray-500 text-sm">No targets assigned yet.</td></tr>
                            ) : (
                                targets.map((t) => (
                                    <tr key={t.id} className="hover:bg-gray-50">
                                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{t.title}</td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                                            <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${t.category === 'Quantitative' ? 'bg-green-100 text-green-800' : 'bg-blue-100 text-blue-800'}`}>
                                                {t.category}
                                            </span>
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-center text-gray-500">
                                            {t.targetValue} {t.unit}
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-right font-bold text-green-600">
                                            {t.EmployeeKPI?.customBonus ? `Rs. ${t.EmployeeKPI.customBonus.toLocaleString()}` : '-'}
                                        </td>
                                    </tr>
                                ))
                            )}
                        </tbody>
                    </table>
                </div>
            </div>

            {/* Bonus History Section */}
            <div className="bg-white shadow rounded-lg overflow-hidden border border-gray-200">
                <div className="px-6 py-4 border-b border-gray-200 bg-gray-50">
                    <h3 className="text-lg font-medium text-gray-900 flex items-center">
                        <Calendar className="w-5 h-5 mr-2 text-purple-600" />
                        Payment History
                    </h3>
                </div>
                <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Date</th>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Reason</th>
                                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Amount</th>
                                <th className="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                            {history.length === 0 ? (
                                <tr><td colSpan="4" className="px-6 py-4 text-center text-gray-500 text-sm">No bonus payments recorded yet.</td></tr>
                            ) : (
                                history.map((h) => (
                                    <tr key={h.id} className="hover:bg-gray-50">
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                            {new Date(h.dateGiven).toLocaleDateString()}
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{h.reason}</td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-right font-bold text-gray-900">
                                            Rs. {parseFloat(h.amount).toLocaleString()}
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-center">
                                            <span className="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-green-100 text-green-800">
                                                Paid
                                            </span>
                                        </td>
                                    </tr>
                                ))
                            )}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export default MyBonuses;
