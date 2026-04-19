import React, { useState, useEffect } from 'react';
import { Save, BarChart2, Plus, Pencil, Trash2 } from 'lucide-react';
import api from '../../api/axios';

const ManagerDataEntry = () => {
    const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
    const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);

    const [metrics, setMetrics] = useState([]);
    const [logs, setLogs] = useState({});

    const [saving, setSaving] = useState(false);
    const [message, setMessage] = useState('');

    // New / Edit Measure state
    const [showModal, setShowModal] = useState(false);
    const [editingMetric, setEditingMetric] = useState(null); // null = add mode
    const [modalData, setModalData] = useState({ name: '', description: '', unit: '' });

    const fetchMetricsAndLogs = async () => {
        try {
            const metricsRes = await api.get('/company-data/metrics');
            setMetrics(metricsRes.data);

            const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;
            const logsRes = await api.get(`/company-data/logs?period=${period}`);

            const newLogs = {};
            logsRes.data.forEach(log => {
                newLogs[log.metricId] = log.value;
            });
            setLogs(newLogs);
        } catch (err) {
            console.error('Failed to load company data', err);
        }
    };

    useEffect(() => {
        fetchMetricsAndLogs();
    }, [selectedYear, selectedMonth]);

    const handleSaveTrack = async (metricId) => {
        setSaving(true);
        setMessage('');
        try {
            const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;
            const value = parseFloat(logs[metricId]);
            if (isNaN(value)) { setSaving(false); return; }
            await api.post('/company-data/logs', { metricId, period, value });
            setMessage('Value saved successfully');
            setTimeout(() => setMessage(''), 3000);
        } catch (err) {
            setMessage('Failed to save value');
        } finally {
            setSaving(false);
        }
    };

    const openAddModal = () => {
        setEditingMetric(null);
        setModalData({ name: '', description: '', unit: '' });
        setShowModal(true);
    };

    const openEditModal = (metric) => {
        setEditingMetric(metric);
        setModalData({ name: metric.name, description: metric.description || '', unit: metric.unit || '' });
        setShowModal(true);
    };

    const handleSaveModal = async () => {
        if (!modalData.name) return;
        try {
            if (editingMetric) {
                // Update
                const res = await api.put(`/company-data/metrics/${editingMetric.id}`, modalData);
                setMetrics(metrics.map(m => m.id === editingMetric.id ? res.data : m));
            } else {
                // Create
                const res = await api.post('/company-data/metrics', modalData);
                setMetrics([...metrics, res.data]);
            }
            setShowModal(false);
        } catch (err) {
            console.error('Failed to save measure:', err);
            alert('Failed to save measure. Please try again.');
        }
    };

    const handleDelete = async (metricId) => {
        if (!window.confirm('Are you sure you want to delete this measure and all its historical data logs?')) return;
        try {
            await api.delete(`/company-data/metrics/${metricId}`);
            setMetrics(metrics.filter(m => m.id !== metricId));
            const updatedLogs = { ...logs };
            delete updatedLogs[metricId];
            setLogs(updatedLogs);
        } catch (err) {
            console.error('Failed to delete metric:', err);
            alert('Failed to delete measure.');
        }
    };

    return (
        <div>
            <div className="flex justify-between items-center mb-6">
                <div>
                    <h2 className="text-2xl font-bold text-gray-900 flex items-center">
                        <BarChart2 className="w-6 h-6 mr-3 text-blue-600" />
                        Company Data Entry
                    </h2>
                    <p className="text-sm text-gray-500">Log raw, overarching business metrics for the organization.</p>
                </div>
                <button
                    onClick={openAddModal}
                    className="bg-blue-600 text-white px-4 py-2 rounded-lg flex items-center hover:bg-blue-700 font-medium"
                >
                    <Plus className="w-5 h-5 mr-2" />
                    New Measure
                </button>
            </div>

            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
                <div className="flex space-x-4 items-center">
                    <div className="text-sm font-semibold text-gray-600 mr-2">Select Period:</div>
                    <select
                        className="w-full md:w-auto p-2 border border-gray-300 rounded-md shadow-sm bg-gray-50 text-sm font-medium"
                        value={selectedYear}
                        onChange={(e) => setSelectedYear(parseInt(e.target.value))}
                    >
                        <option value={2024}>2024</option>
                        <option value={2025}>2025</option>
                        <option value={2026}>2026</option>
                        <option value={2027}>2027</option>
                    </select>

                    <select
                        className="w-full md:w-auto p-2 border border-gray-300 rounded-md shadow-sm bg-gray-50 text-sm font-medium"
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
                <div className="mb-6 p-4 rounded-lg bg-green-50 text-green-700 font-medium border border-green-200 shadow-sm">{message}</div>
            )}

            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
                {metrics.length === 0 ? (
                    <div className="text-center py-10 text-gray-400 italic">
                        No company measures defined yet. Click 'New Measure' to add one.
                    </div>
                ) : (
                    <div className="space-y-4">
                        {metrics.map(metric => (
                            <div key={metric.id} className="border border-gray-100 rounded-lg p-5 border-l-4 border-l-green-500 bg-gray-50/50 flex flex-col md:flex-row md:items-center justify-between shadow-sm hover:shadow-md transition-shadow">
                                <div className="flex-1 pr-4">
                                    <h3 className="font-bold text-gray-800 text-lg">{metric.name}</h3>
                                    {metric.description && <p className="text-sm text-gray-500 mt-1">{metric.description}</p>}
                                </div>
                                <div className="mt-4 md:mt-0 flex items-center space-x-2">
                                    {/* Value input */}
                                    <div className="flex items-center space-x-2 bg-white p-2 rounded-lg border border-gray-200 shadow-sm">
                                        <input
                                            type="number"
                                            className="w-32 border border-gray-200 rounded-md px-3 py-2 font-bold text-gray-800 focus:outline-none focus:ring-2 focus:ring-green-500 bg-gray-50"
                                            placeholder="Value"
                                            value={logs[metric.id] !== undefined ? logs[metric.id] : ''}
                                            onChange={(e) => setLogs({ ...logs, [metric.id]: e.target.value })}
                                        />
                                        {metric.unit && <span className="text-xs font-bold text-gray-500 px-2 w-12">{metric.unit}</span>}
                                        <button
                                            onClick={() => handleSaveTrack(metric.id)}
                                            disabled={saving || logs[metric.id] === undefined || logs[metric.id] === ''}
                                            className="bg-green-600 hover:bg-green-700 disabled:bg-gray-300 text-white p-2.5 rounded-md flex items-center justify-center transition-colors shadow-sm"
                                            title="Save Value"
                                        >
                                            <Save className="w-4 h-4" />
                                        </button>
                                    </div>

                                    {/* Edit button */}
                                    <button
                                        onClick={() => openEditModal(metric)}
                                        className="p-2.5 rounded-md border border-gray-200 bg-white text-blue-600 hover:bg-blue-50 shadow-sm transition-colors"
                                        title="Edit Measure"
                                    >
                                        <Pencil className="w-4 h-4" />
                                    </button>

                                    {/* Delete button */}
                                    <button
                                        onClick={() => handleDelete(metric.id)}
                                        className="p-2.5 rounded-md border border-red-100 bg-white text-red-500 hover:bg-red-50 shadow-sm transition-colors"
                                        title="Delete Measure"
                                    >
                                        <Trash2 className="w-4 h-4" />
                                    </button>
                                </div>
                            </div>
                        ))}
                    </div>
                )}
            </div>

            {/* Add / Edit Modal */}
            {showModal && (
                <div className="fixed z-10 inset-0 overflow-y-auto">
                    <div className="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
                        <div className="fixed inset-0 transition-opacity" aria-hidden="true">
                            <div className="absolute inset-0 bg-gray-500 opacity-75"></div>
                        </div>
                        <span className="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
                        <div className="inline-block align-bottom bg-white rounded-lg px-4 pt-5 pb-4 text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full sm:p-6">
                            <div>
                                <h3 className="text-lg leading-6 font-medium text-gray-900 mb-4">
                                    {editingMetric ? 'Edit Measure' : 'Add New Company Measure'}
                                </h3>
                                <div className="space-y-4">
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Metric Name</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-green-500 focus:border-green-500 sm:text-sm"
                                            value={modalData.name}
                                            onChange={(e) => setModalData({ ...modalData, name: e.target.value })}
                                            placeholder="e.g. Total Revenue"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Description</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-green-500 focus:border-green-500 sm:text-sm"
                                            value={modalData.description}
                                            onChange={(e) => setModalData({ ...modalData, description: e.target.value })}
                                            placeholder="Optional details about this metric"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Unit</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-green-500 focus:border-green-500 sm:text-sm"
                                            value={modalData.unit}
                                            onChange={(e) => setModalData({ ...modalData, unit: e.target.value })}
                                            placeholder="e.g. USD, Count, %"
                                        />
                                    </div>
                                </div>
                            </div>
                            <div className="mt-5 sm:mt-6 sm:grid sm:grid-cols-2 sm:gap-3 sm:grid-flow-row-dense">
                                <button
                                    type="button"
                                    onClick={handleSaveModal}
                                    className="w-full inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-green-600 text-base font-medium text-white hover:bg-green-700 focus:outline-none sm:col-start-2 sm:text-sm"
                                >
                                    {editingMetric ? 'Update' : 'Save'}
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setShowModal(false)}
                                    className="mt-3 w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 sm:mt-0 sm:col-start-1 sm:text-sm"
                                >
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

export default ManagerDataEntry;
