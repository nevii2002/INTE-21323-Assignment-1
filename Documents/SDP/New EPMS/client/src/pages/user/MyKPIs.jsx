import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Target, PieChart, TrendingUp, Award, Save } from 'lucide-react';
import { useAuth } from '../../context/AuthContext';

const MyKPIs = () => {
    const { user } = useAuth();
    const [kpis, setKpis] = useState([]);
    const [loading, setLoading] = useState(true);

    const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
    const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
    const [logs, setLogs] = useState({});
    const [message, setMessage] = useState('');
    const [saving, setSaving] = useState(false);

    useEffect(() => {
        const fetchKpis = async () => {
            try {
                // Ensure we hit the 'my-kpis' endpoint we just created
                const res = await axios.get('http://localhost:5000/api/staff/my-kpis', {
                    headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
                });
                setKpis(res.data);
            } catch (error) {
                console.error("Error fetching KPIs:", error);
            } finally {
                setLoading(false);
            }
        };
        fetchKpis();
    }, []);

    // Fetch logs when month/period changes
    useEffect(() => {
        if (!user || kpis.length === 0) return;

        const fetchLogs = async () => {
            try {
                const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;
                const logsRes = await axios.get(`http://localhost:5000/api/quantitative/${user.id}?period=${period}`, {
                    headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
                });
                const newLogs = {};
                logsRes.data.forEach(log => {
                    newLogs[log.kpiId] = {
                        actualValue: log.actualValue,
                        updatedAt: log.updatedAt
                    };
                });
                setLogs(newLogs);
            } catch (err) {
                console.error('Failed to load actuals', err);
            }
        };
        fetchLogs();
    }, [selectedYear, selectedMonth, kpis, user]);

    const handleSave = async (kpiId) => {
        setSaving(true);
        setMessage('');
        try {
            const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;
            const actualValue = parseFloat(logs[kpiId]?.actualValue);
            if (isNaN(actualValue)) {
                setMessage('Please enter a valid number.');
                setSaving(false);
                return;
            }

            const res = await axios.post('http://localhost:5000/api/quantitative', {
                employeeId: user.id,
                kpiId,
                period,
                actualValue
            }, {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });

            // Update log with accurate Date Last Updated from server response
            setLogs(prev => ({
                ...prev,
                [kpiId]: { actualValue, updatedAt: res.data?.log?.updatedAt || new Date().toISOString() }
            }));

            setMessage('Metric successfully logged!');
            setTimeout(() => setMessage(''), 3000);
        } catch (err) {
            setMessage('Failed to save metric.');
        } finally {
            setSaving(false);
        }
    };

    const evaluationKpis = kpis.filter(k => k.type === 'EVALUATION');
    const bonusKpis = kpis.filter(k => k.type === 'BONUS');

    // Calculate Overall KPI Score based on Evaluation KPIs
    const overallScore = evaluationKpis.reduce((total, kpi) => {
        const actual = parseFloat(logs[kpi.id]?.actualValue) || 0;
        const target = parseFloat(kpi.targetValue) || 1;
        const weight = parseFloat(kpi.EmployeeKPI?.customWeight) || parseFloat(kpi.weight) || 0;

        let scorePercent = (actual / target) * 100;
        if (scorePercent > 100) scorePercent = 100;

        return total + (scorePercent * (weight / 100));
    }, 0);

    const getBonusRemaining = (kpi) => {
        const actual = parseFloat(logs[kpi.id]?.actualValue) || 0;
        const target = parseFloat(kpi.targetValue) || 0;
        const rem = target - actual;
        return rem > 0 ? rem : 0;
    };

    return (
        <div className="p-6 bg-gray-50 min-h-screen">
            <div className="flex flex-col md:flex-row md:items-center justify-between mb-6">
                <h2 className="text-2xl font-bold text-gray-800 flex items-center">
                    <Target className="mr-2 text-blue-600" />
                    My Targets & KPIs
                </h2>

                {/* Period Selector */}
                <div className="flex space-x-3 items-center mt-4 md:mt-0">
                    <span className="text-sm text-gray-600 font-medium mr-2">Track Month:</span>
                    <select
                        className="p-2 border border-gray-300 rounded-md shadow-sm bg-white text-sm"
                        value={selectedYear}
                        onChange={(e) => setSelectedYear(parseInt(e.target.value))}
                    >
                        <option value={2025}>2025</option>
                        <option value={2026}>2026</option>
                    </select>

                    <select
                        className="p-2 border border-gray-300 rounded-md shadow-sm bg-white text-sm"
                        value={selectedMonth}
                        onChange={(e) => setSelectedMonth(parseInt(e.target.value))}
                    >
                        {Array.from({ length: 12 }, (_, i) => (
                            <option key={i + 1} value={i + 1}>
                                {new Date(2025, i, 1).toLocaleString('default', { month: 'short' })}
                            </option>
                        ))}
                    </select>
                </div>
            </div>

            {message && (
                <div className="mb-6 p-4 rounded-lg flex items-center shadow-sm bg-green-50 text-green-700 font-medium">
                    {message}
                </div>
            )}

            {loading ? (
                <div className="text-center py-10">Loading targets...</div>
            ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">

                    {/* Performance Section */}
                    <div className="bg-white p-6 rounded-lg shadow-md border-t-4 border-blue-500">
                        <div className="flex items-center justify-between mb-4">
                            <h3 className="text-xl font-semibold text-gray-800 flex items-center">
                                <PieChart className="w-5 h-5 mr-2 text-blue-500" />
                                Performance Evaluation Weights
                            </h3>
                            <div className="flex space-x-3">
                                <span className="text-xs font-bold bg-green-100 text-green-800 px-3 py-1.5 rounded flex items-center shadow-sm">
                                    Overall Score: {overallScore.toFixed(1)}%
                                </span>
                                <span className="text-xs font-bold bg-blue-100 text-blue-800 px-3 py-1.5 rounded flex items-center shadow-sm">
                                    Total Weights: {evaluationKpis.reduce((sum, k) => sum + (k.EmployeeKPI?.customWeight || parseFloat(k.weight) || 0), 0)}%
                                </span>
                            </div>
                        </div>
                        <p className="text-sm text-gray-500 mb-4">
                            Your performance review score is calculated based on these weighted metrics. Log your actual data periodically.
                        </p>

                        <div className="space-y-4">
                            {evaluationKpis.length === 0 ? (
                                <p className="text-gray-400 italic text-center">No performance KPIs assigned.</p>
                            ) : (
                                evaluationKpis.map(kpi => (
                                    <div key={kpi.id} className="border border-gray-100 rounded-lg p-4 hover:bg-gray-50 transition shadow-sm">
                                        <div className="flex justify-between items-start mb-1">
                                            <span className="font-medium text-gray-700 text-base">{kpi.title}</span>
                                            <span className="font-bold text-blue-600 bg-blue-50 px-2 py-0.5 rounded text-sm">{kpi.EmployeeKPI?.customWeight || kpi.weight}%</span>
                                        </div>
                                        <p className="text-xs text-gray-500 mb-2">{kpi.description}</p>

                                        <div className="text-xs font-semibold text-blue-700 bg-blue-50 px-2 py-1 rounded inline-block mb-3">
                                            Target: {kpi.targetValue} {kpi.unit}
                                        </div>

                                        {logs[kpi.id]?.updatedAt && (
                                            <div className="text-[11px] text-gray-400 mb-1 flex items-center">
                                                Last Updated: {new Date(logs[kpi.id].updatedAt).toLocaleString()}
                                            </div>
                                        )}

                                        <div className="flex items-center gap-2 mt-2 pt-3 border-t border-gray-100">
                                            <input
                                                type="number"
                                                className="w-24 border border-gray-300 rounded px-2 py-1.5 text-sm font-bold text-gray-800 bg-white"
                                                placeholder={`Actual`}
                                                value={logs[kpi.id]?.actualValue !== undefined ? logs[kpi.id].actualValue : ''}
                                                onChange={(e) => setLogs({ ...logs, [kpi.id]: { ...logs[kpi.id], actualValue: e.target.value } })}
                                            />
                                            <span className="text-xs text-gray-500 font-medium">{kpi.unit}</span>
                                            <button
                                                onClick={() => handleSave(kpi.id)}
                                                disabled={saving || logs[kpi.id]?.actualValue === undefined || logs[kpi.id]?.actualValue === ''}
                                                className="ml-auto bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white px-3 py-1.5 rounded flex items-center transition-colors text-xs font-semibold"
                                                title="Save Metric"
                                            >
                                                <Save className="w-4 h-4 mr-1" />
                                                Save
                                            </button>
                                        </div>
                                    </div>
                                ))
                            )}
                        </div>
                    </div>

                    {/* Bonus Section */}
                    <div className="bg-white p-6 rounded-lg shadow-md border-t-4 border-purple-500">
                        <div className="flex items-center justify-between mb-4">
                            <h3 className="text-xl font-semibold text-gray-800 flex items-center">
                                <TrendingUp className="w-5 h-5 mr-2 text-purple-500" />
                                Bonus Opportunities
                            </h3>
                        </div>
                        <p className="text-sm text-gray-500 mb-4">
                            Active projects or targets that contribute to your financial bonuses.
                        </p>

                        <div className="space-y-4">
                            {bonusKpis.length === 0 ? (
                                <p className="text-gray-400 italic text-center">No bonus targets assigned.</p>
                            ) : (
                                bonusKpis.map(kpi => (
                                    <div key={kpi.id} className="bg-purple-50 border border-purple-100 rounded-lg p-4 shadow-sm">
                                        <div className="flex justify-between items-start mb-2">
                                            <div className="flex items-center">
                                                <Award className="w-5 h-5 text-purple-600 mr-2" />
                                                <span className="font-semibold text-purple-900 text-base">{kpi.title}</span>
                                            </div>
                                            <span className="font-bold text-purple-700 bg-purple-100 px-2 py-0.5 rounded text-sm">
                                                Share: {kpi.EmployeeKPI?.customWeight || kpi.weight}%
                                            </span>
                                        </div>
                                        <p className="text-xs text-purple-800 opacity-80 mb-3">{kpi.description}</p>

                                        <div className="flex items-center text-xs font-semibold text-purple-700 bg-white px-2 py-1.5 rounded w-max mb-1 border border-purple-100">
                                            Target: {kpi.targetValue} {kpi.unit}
                                        </div>
                                        <div className="flex items-center text-[11px] font-bold text-orange-600 bg-orange-50 px-2 py-1 rounded w-max mb-3 border border-orange-100">
                                            To Complete: {getBonusRemaining(kpi)} {kpi.unit}
                                        </div>

                                        {logs[kpi.id]?.updatedAt && (
                                            <div className="text-[11px] text-gray-500 mb-1 flex items-center">
                                                Last Updated: {new Date(logs[kpi.id].updatedAt).toLocaleString()}
                                            </div>
                                        )}

                                        <div className="flex items-center gap-2 mt-2 pt-3 border-t border-purple-200">
                                            <input
                                                type="number"
                                                className="w-24 border border-purple-300 rounded px-2 py-1.5 text-sm font-bold text-gray-800 bg-white"
                                                placeholder={`Actual`}
                                                value={logs[kpi.id]?.actualValue !== undefined ? logs[kpi.id].actualValue : ''}
                                                onChange={(e) => setLogs({ ...logs, [kpi.id]: { ...logs[kpi.id], actualValue: e.target.value } })}
                                            />
                                            <span className="text-xs text-purple-600 font-medium">{kpi.unit}</span>
                                            <button
                                                onClick={() => handleSave(kpi.id)}
                                                disabled={saving || logs[kpi.id]?.actualValue === undefined || logs[kpi.id]?.actualValue === ''}
                                                className="ml-auto bg-purple-600 hover:bg-purple-700 disabled:bg-gray-300 text-white px-3 py-1.5 rounded flex items-center transition-colors text-xs font-semibold"
                                                title="Save Metric"
                                            >
                                                <Save className="w-4 h-4 mr-1" />
                                                Save
                                            </button>
                                        </div>
                                    </div>
                                ))
                            )}
                        </div>
                    </div>

                </div>
            )}
        </div>
    );
};

export default MyKPIs;
