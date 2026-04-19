import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Calculator, DollarSign, Award, AlertCircle, CheckCircle } from 'lucide-react';
import api from '../../api/axios';

const BonusCalculation = () => {
    const [selectedEmployee, setSelectedEmployee] = useState('');
    const [employees, setEmployees] = useState([]);
    const [rows, setRows] = useState([]);
    const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
    const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [isLoading, setIsLoading] = useState(false);

    useEffect(() => {
        const fetchEmployees = async () => {
            try {
                const response = await api.get('/staff');
                const staff = response.data.filter(u => u.role === 'Employee');
                setEmployees(staff);
                if (staff.length > 0) setSelectedEmployee(staff[0].id);
            } catch (error) { console.error(error); }
        };
        fetchEmployees();
    }, []);

    useEffect(() => {
        if (!selectedEmployee) return;

        const fetchData = async () => {
            setIsLoading(true);
            try {
                // Fetch Employee's KPIs
                const profileRes = await api.get(`/staff/${selectedEmployee}`);
                const assignedKPIs = profileRes.data.assignedKPIs || [];
                const quantKPIs = assignedKPIs.filter(k => k.category === 'Quantitative');

                // Fetch their logged Quantitative metrics for this period
                const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;
                const logsRes = await api.get(`/quantitative/${selectedEmployee}?period=${period}`);

                const logsMap = {};
                logsRes.data.forEach(l => {
                    logsMap[l.kpiId] = l.actualValue;
                });

                // Generate Rows
                const generatedRows = quantKPIs.map(kpi => {
                    return {
                        id: kpi.id,
                        project: kpi.title,
                        target: kpi.targetValue,
                        weight: kpi.EmployeeKPI?.customWeight || kpi.weight,
                        bonusAt100: kpi.EmployeeKPI?.customBonus || 20000,
                        actual: logsMap[kpi.id] !== undefined ? logsMap[kpi.id] : ''
                    };
                });

                setRows(generatedRows);
            } catch (err) {
                console.error('Failed to load bonus data', err);
            } finally {
                setIsLoading(false);
            }
        };

        fetchData();
    }, [selectedEmployee, selectedYear, selectedMonth]);

    const handleActualChange = (id, value) => {
        setRows(prev => prev.map(r => r.id === id ? { ...r, actual: value } : r));
    };

    const calculateResult = (row) => {
        const actual = parseFloat(row.actual) || 0;
        const target = parseFloat(row.target) || 1;
        const bonusPotential = parseFloat(row.bonusAt100) || 0;

        const achievementRatio = (actual / target) * 100;
        let payoutPercentage = 0;
        let label = 'Not Met';
        let labelColor = 'text-red-500';
        let icon = AlertCircle;

        if (achievementRatio >= 120) {
            payoutPercentage = 1.5; // 150%
            label = 'Exceeded 🚀';
            labelColor = 'text-purple-600';
            icon = Award;
        } else if (achievementRatio >= 100) {
            payoutPercentage = 1.0; // 100%
            label = 'Met ✅';
            labelColor = 'text-green-600';
            icon = CheckCircle;
        } else if (achievementRatio >= 85) {
            payoutPercentage = 0.5; // 50%
            label = 'Partially Met ⚠️';
            labelColor = 'text-yellow-600';
            icon = AlertCircle;
        } else {
            payoutPercentage = 0; // 0%
            label = 'Not Met ❌';
            labelColor = 'text-red-500';
            icon = AlertCircle;
        }

        const earned = bonusPotential * payoutPercentage * (row.weight / 100);

        return { earned, label, labelColor, achievementRatio, payoutPercentage };
    };

    const totalBonus = rows.reduce((sum, row) => sum + calculateResult(row).earned, 0);

    const formatCurrency = (amount) => {
        return new Intl.NumberFormat('en-LK', { style: 'currency', currency: 'LKR' }).format(amount);
    };

    const handleSubmitBonus = async () => {
        if (!selectedEmployee || totalBonus === 0) {
            alert('Please select an employee and ensure there is a calculated bonus.');
            return;
        }

        const periodName = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;

        if (!window.confirm(`Are you sure you want to award a bonus of ${formatCurrency(totalBonus)} to this employee for ${periodName}?`)) {
            return;
        }

        setIsSubmitting(true);
        try {
            await api.post('/bonuses', {
                employeeId: selectedEmployee,
                amount: totalBonus,
                reason: `Monthly Bonus - ${periodName}`
            });
            alert('Bonus awarded successfully! The employee can now view this in their portal.');
        } catch (error) {
            console.error('Failed to submit bonus:', error);
            alert('Failed to submit bonus.');
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <div className="p-6 bg-gray-50 min-h-screen font-sans">
            <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <h2 className="text-2xl font-bold text-gray-800 flex items-center">
                        <Calculator className="mr-2 text-blue-600" />
                        Bonus Calculation
                    </h2>
                    <p className="text-sm text-gray-500 mt-1">
                        Bonus Logic: &lt;85% (0%), 85-99% (50%), 100-119% (100%), &ge;120% (150%)
                    </p>
                </div>

                <div className="flex flex-wrap items-center gap-4">
                    <div className="flex items-center">
                        <label className="mr-2 text-sm font-semibold text-gray-600">Employee:</label>
                        <select
                            className="p-2 border rounded shadow-sm bg-white min-w-[200px]"
                            value={selectedEmployee}
                            onChange={(e) => setSelectedEmployee(e.target.value)}
                        >
                            {employees.map(e => <option key={e.id} value={e.id}>{e.username}</option>)}
                        </select>
                    </div>

                    <div className="flex items-center gap-2">
                        <select
                            className="p-2 border rounded shadow-sm bg-white"
                            value={selectedYear}
                            onChange={(e) => setSelectedYear(parseInt(e.target.value))}
                        >
                            <option value={2025}>2025</option>
                            <option value={2026}>2026</option>
                        </select>
                        <select
                            className="p-2 border rounded shadow-sm bg-white"
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
            </div>

            {/* Scoring Logic Display */}
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4 mb-6">
                <h3 className="text-sm font-bold text-blue-800 mb-2">KPI Scoring Logic (Achievement vs Target)</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <table className="min-w-full text-xs text-left">
                        <thead>
                            <tr className="border-b border-blue-200">
                                <th className="py-1">Ratio (Achieved / Target)</th>
                                <th className="py-1">Score</th>
                                <th className="py-1">Level</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td>&lt; 85%</td><td>0</td><td>❌ Not Met</td></tr>
                            <tr><td>85% - 99%</td><td>50</td><td>⚠️ Partially Met</td></tr>
                            <tr><td>100% - 119%</td><td>100</td><td>✅ Met</td></tr>
                            <tr><td>≥ 120%</td><td>150</td><td>🚀 Exceeded</td></tr>
                        </tbody>
                    </table>
                    <div>
                        <h4 className="text-xs font-bold text-blue-800 mb-1">Bonus Eligibility (Final Weighted Score)</h4>
                        <ul className="text-xs space-y-1 text-blue-900">
                            <li>&lt; 70: No Bonus ❌</li>
                            <li>70 - 99: 50% Bonus</li>
                            <li>100 - 119: 100% Bonus</li>
                            <li>≥ 120: 150% Bonus</li>
                        </ul>
                    </div>
                </div>
            </div>

            <div className="bg-white rounded-lg shadow-lg overflow-hidden border border-gray-200">
                <div className="grid grid-cols-12 gap-4 px-6 py-4 bg-gray-50 border-b border-gray-200 text-xs font-bold text-gray-700 uppercase tracking-wider items-center">
                    <div className="col-span-3">Project / KPI</div>
                    <div className="col-span-1 text-center">Target</div>
                    <div className="col-span-1 text-center">Weight</div>
                    <div className="col-span-2 text-center text-blue-600">Actual</div>
                    <div className="col-span-2 text-center">Achievement</div>
                    <div className="col-span-1 text-center">Payout %</div>
                    <div className="col-span-2 text-right text-green-600">Earned Bonus</div>
                </div>

                {isLoading ? (
                    <div className="p-8 text-center text-gray-500">Loading metrics...</div>
                ) : rows.length === 0 ? (
                    <div className="p-8 text-center text-gray-400 italic">No Quantitative KPIs assigned to this employee. Please assign them in Staff Management.</div>
                ) : (
                    <div className="divide-y divide-gray-100">
                        {rows.map((row) => {
                            const { earned, label, labelColor, achievementRatio, payoutPercentage } = calculateResult(row);

                            return (
                                <div key={row.id} className="grid grid-cols-12 gap-4 px-6 py-4 items-center hover:bg-gray-50 transition-colors">
                                    <div className="col-span-3 font-medium text-gray-800 text-sm">{row.project}</div>
                                    <div className="col-span-1 text-center text-gray-600 text-sm">{row.target}</div>
                                    <div className="col-span-1 text-center text-gray-400 text-xs">{row.weight}%</div>

                                    <div className="col-span-2">
                                        <input
                                            type="number"
                                            value={row.actual}
                                            onChange={(e) => handleActualChange(row.id, e.target.value)}
                                            className="w-full text-center border-b-2 border-blue-200 focus:border-blue-500 focus:outline-none bg-transparent font-bold text-blue-800 py-1"
                                            placeholder="-"
                                        />
                                    </div>

                                    <div className={`col-span-2 text-center text-xs font-bold ${labelColor}`}>
                                        {row.actual !== '' ? (
                                            <div className="flex flex-col items-center">
                                                <span>{achievementRatio.toFixed(0)}%</span>
                                                <span className="text-[10px] opacity-80">{label}</span>
                                            </div>
                                        ) : <span className="text-gray-300">-</span>}
                                    </div>

                                    <div className="col-span-1 text-center text-sm font-semibold text-gray-600">
                                        {row.actual !== '' ? `${(payoutPercentage * 100).toFixed(0)}%` : '-'}
                                    </div>

                                    <div className="col-span-2 text-right font-bold text-green-600 text-lg">
                                        {formatCurrency(earned)}
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                )}

                <div className="bg-gray-50 px-6 py-6 border-t border-gray-200 flex justify-end items-center">
                    <div className="text-right">
                        <p className="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-1">Total Bonus Payout</p>
                        <p className="text-3xl font-extrabold text-indigo-600 flex items-center justify-end">
                            {formatCurrency(totalBonus)}
                        </p>
                        <button
                            onClick={handleSubmitBonus}
                            disabled={isSubmitting || totalBonus === 0 || rows.length === 0}
                            className="mt-4 bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-6 rounded shadow-lg transition duration-200 flex items-center disabled:opacity-50"
                        >
                            <DollarSign className="w-5 h-5 mr-2" />
                            {isSubmitting ? 'Submitting...' : 'Submit Bonus'}
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default BonusCalculation;
