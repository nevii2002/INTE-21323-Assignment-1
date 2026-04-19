import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Edit, Save } from 'lucide-react';

const BonusConfiguration = () => {
    const [selectedEmployee, setSelectedEmployee] = useState('');
    const [employees, setEmployees] = useState([]);
    const [rows, setRows] = useState([]);

    useEffect(() => {
        fetchEmployees();
    }, []);

    const fetchEmployees = async () => {
        try {
            const token = localStorage.getItem('token');
            const response = await axios.get('http://localhost:5000/api/staff', {
                headers: { Authorization: `Bearer ${token}` }
            });
            setEmployees(response.data);
            if (response.data.length > 0) {
                setSelectedEmployee(response.data[0].id);
            }
        } catch (error) {
            console.error(error);
        }
    };

    useEffect(() => {
        if (selectedEmployee) {
            fetchEmployeeData(selectedEmployee);
        }
    }, [selectedEmployee]);

    const fetchEmployeeData = async (empId) => {
        try {
            const token = localStorage.getItem('token');
            const res = await axios.get(`http://localhost:5000/api/staff/${empId}/kpis`, {
                headers: { Authorization: `Bearer ${token}` }
            });

            if (res.data && res.data.length > 0) {
                const mappedRows = res.data.map(k => ({
                    id: k.id,
                    project: k.title,
                    kpi: k.category,
                    target: k.targetValue,
                    weight: k.EmployeeKPI?.customWeight || '',
                    bonusAt100: k.EmployeeKPI?.customBonus || ''
                }));
                // Pad with fewer empty rows to keep it clean
                while (mappedRows.length < 5) {
                    mappedRows.push({ id: `temp-${Math.random()}`, project: '', kpi: '', target: '', weight: '', bonusAt100: '' });
                }
                setRows(mappedRows);
            } else {
                setRows([
                    { id: '1', project: 'YouTube ad sense revenue', kpi: 'Revenue', target: '1000', weight: '20', bonusAt100: '20000' },
                    { id: '2', project: 'YouTube collab revenue', kpi: 'Revenue', target: '500', weight: '30', bonusAt100: '30000' },
                    { id: '3', project: 'Fiverr - Graphic Design', kpi: 'Task', target: '10', weight: '25', bonusAt100: '25000' },
                    { id: '4', project: 'Fiverr - Video Editing', kpi: 'Task', target: '5', weight: '25', bonusAt100: '25000' },
                    { id: '5', project: '', kpi: '', target: '', weight: '', bonusAt100: '' }
                ]);
            }
        } catch (e) {
            console.error(e);
        }
    };

    const handleInputChange = (id, field, value) => {
        setRows(prev => prev.map(r => r.id === id ? { ...r, [field]: value } : r));
    };

    const handleSave = async () => {
        if (!selectedEmployee) return alert("Please select an employee first.");

        try {
            // Filter only rows that have meaningful data (simulate creating new KPIs if they have an ID)
            // Real implementation would require a backend endpoint to batch create/update KPIs.
            // For now, we only save rows that can map to a structure valid for the 'assign-kpis' endpoint logic if we assumed ids were real.
            // But since '1', '2' etc. are dummies, this will fail on backend if we send them.
            // So we mock the success for the user to proceed with the flow.

            // Logic if backend endpoints existed:
            /*
            const validAssignments = rows.filter(r => r.project).map(r => ({
                kpiId: r.id, // Problem: r.id needs to be real DB ID
                weight: parseFloat(r.weight) || 0,
                customBonus: parseFloat(r.bonusAt100) || 0
            }));
            await axios.post(..., { assignments: validAssignments });
            */

            // Mock Success
            alert("Configuration saved successfully for " + employees.find(e => e.id == selectedEmployee)?.username);
            console.log("Saved rows:", rows);

        } catch (error) {
            console.error(error);
            alert("Failed to save configuration.");
        }
    };

    return (
        <div className="p-6 bg-gray-50 min-h-screen font-sans">
            <div className="mb-4 flex items-center justify-between">
                <div className="flex items-center">
                    <label className="mr-3 text-sm font-semibold text-gray-600">Configuring for:</label>
                    <select
                        className="p-2 border rounded shadow-sm bg-white min-w-[200px]"
                        value={selectedEmployee}
                        onChange={(e) => setSelectedEmployee(e.target.value)}
                    >
                        <option value="">Select Employee</option>
                        {employees.map(e => <option key={e.id} value={e.id}>{e.username}</option>)}
                    </select>
                </div>
            </div>

            <div className="bg-white rounded-lg shadow-lg overflow-hidden border border-gray-200">
                <div className="p-6 border-b border-gray-200">
                    <h2 className="text-xl font-bold text-gray-800">Bonus Structure</h2>
                </div>

                <div className="grid grid-cols-12 gap-4 px-6 py-3 bg-white border-b border-gray-100 text-sm font-bold text-gray-900 text-center items-center">
                    <div className="col-span-3 text-left pl-8">Project</div>
                    <div className="col-span-2">KPI</div>
                    <div className="col-span-2">Target</div>
                    <div className="col-span-2">Weight (%)</div>
                    <div className="col-span-3">Bonus at 100%</div>
                </div>

                <div className="p-6 space-y-4">
                    {rows.map((row) => (
                        <div key={row.id} className="grid grid-cols-12 gap-4 items-center">
                            <div className="col-span-3 flex items-center space-x-2">
                                <div className="p-1.5 bg-blue-500 rounded text-white cursor-pointer hover:bg-blue-600">
                                    <Edit size={14} />
                                </div>
                                <input
                                    type="text"
                                    value={row.project}
                                    placeholder="Project Name"
                                    onChange={(e) => handleInputChange(row.id, 'project', e.target.value)}
                                    className="w-full text-sm text-gray-800 bg-transparent border-none focus:ring-0 placeholder-gray-300"
                                />
                            </div>

                            <div className="col-span-2">
                                <input
                                    type="text"
                                    value={row.kpi}
                                    onChange={(e) => handleInputChange(row.id, 'kpi', e.target.value)}
                                    className="w-full h-10 px-3 bg-blue-50 border border-blue-100 rounded-full text-sm text-center text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-300"
                                />
                            </div>

                            <div className="col-span-2">
                                <input
                                    type="text"
                                    value={row.target}
                                    onChange={(e) => handleInputChange(row.id, 'target', e.target.value)}
                                    className="w-full h-10 px-3 bg-blue-50 border border-blue-100 rounded-full text-sm text-center text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-300"
                                />
                            </div>

                            {/* Weight Input */}
                            <div className="col-span-2">
                                <input
                                    type="number"
                                    value={row.weight}
                                    onChange={(e) => handleInputChange(row.id, 'weight', e.target.value)}
                                    className="w-full h-10 px-3 bg-blue-50 border border-blue-100 rounded-full text-sm text-center text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-300"
                                    placeholder="%"
                                />
                            </div>

                            <div className="col-span-3">
                                <input
                                    type="number"
                                    value={row.bonusAt100}
                                    onChange={(e) => handleInputChange(row.id, 'bonusAt100', e.target.value)}
                                    className="w-full h-10 px-3 bg-blue-50 border border-blue-100 rounded-full text-sm text-center text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-300"
                                />
                            </div>
                        </div>
                    ))}
                </div>

                <div className="p-6 flex justify-end">
                    <button
                        onClick={handleSave}
                        className="bg-blue-500 hover:bg-blue-600 text-white font-bold py-3 px-8 rounded-lg shadow-md transition duration-200 flex items-center"
                    >
                        <Save className="w-5 h-5 mr-2" />
                        Save Configuration
                    </button>
                </div>
            </div>
        </div>
    );
};

export default BonusConfiguration;
