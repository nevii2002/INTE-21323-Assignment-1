import React, { useState, useEffect } from 'react';
import { Plus, Trash2, Edit2, FileText, Target, Save, BarChart2 } from 'lucide-react';
import api from '../../api/axios';

// Mock data removed


const KPIFramework = () => {
    const [kpis, setKpis] = useState([]);
    const [showModal, setShowModal] = useState(false);
    const [newKPI, setNewKPI] = useState({ title: '', description: '', category: 'Quantitative', weight: 0, unit: '', targetValue: 0, dataSource: '', role: 'All' });
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);


    // Fetch KPIs from API
    useEffect(() => {
        const fetchKPIs = async () => {
            setIsLoading(true);
            try {
                const response = await api.get('/kpis');
                setKpis(response.data);
                setError(null);
            } catch (err) {
                console.error('Failed to fetch KPIs:', err);
                setError('Failed to load KPIs. Please try again later.');
            } finally {
                setIsLoading(false);
            }
        };

        fetchKPIs();
    }, []);



    const handleAddKPI = async () => {
        if (!newKPI.title || !newKPI.weight) return;

        try {
            const kpiPayload = { ...newKPI, category: 'Quantitative' };
            const response = await api.post('/kpis', kpiPayload);
            setKpis([...kpis, response.data]);
            setShowModal(false);
            setNewKPI({ title: '', description: '', category: 'Quantitative', weight: 0, unit: '', targetValue: 0, dataSource: '', role: 'All' });
        } catch (err) {
            console.error('Failed to add KPI:', err);
            alert('Failed to add KPI. Please try again.');
        }
    };

    const handleDeleteKPI = async (id) => {
        if (!window.confirm('Are you sure you want to delete this KPI?')) return;
        try {
            await api.delete(`/kpis/${id}`);
            setKpis(kpis.filter(k => k.id !== id));
        } catch (err) {
            console.error('Failed to delete KPI:', err);
            alert('Failed to delete KPI.');
        }
    };

    // We now treat everything natively as quantitative, regardless of DB past states
    const filteredKPIs = kpis;

    return (
        <div>
            <div className="flex justify-between items-center mb-6">
                <div>
                    <h2 className="text-2xl font-bold text-gray-900">KPI Framework</h2>
                    <p className="text-sm text-gray-500">Define how employees are evaluated.</p>
                </div>
                <div className="flex space-x-3 items-center">
                    <button
                        onClick={() => setShowModal(true)}
                        className="bg-blue-600 text-white px-4 py-2 rounded-lg flex items-center hover:bg-blue-700"
                    >
                        <Plus className="w-5 h-5 mr-2" />
                        Add KPI
                    </button>
                </div>
            </div>

            {error && <div className="mb-4 text-red-600 bg-red-50 p-3 rounded">{error}</div>}

            {/* List */}
            <div className="bg-white shadow overflow-hidden sm:rounded-md">
                {isLoading ? (
                    <div className="p-6 text-center text-gray-500">Loading KPIs...</div>
                ) : (
                    <ul className="divide-y divide-gray-200">
                        {filteredKPIs.length === 0 ? (
                            <li className="px-6 py-12 text-center text-gray-400">
                                No KPIs defined for this category yet.
                            </li>
                        ) : (
                            filteredKPIs.map((kpi) => (
                                <li key={kpi.id}>
                                    <div className="px-4 py-4 flex items-center sm:px-6">
                                        <div className="min-w-0 flex-1 sm:flex sm:items-center sm:justify-between">
                                            <div>
                                                <div className="flex text-sm">
                                                    <p className="font-medium text-blue-600 truncate">{kpi.title}</p>
                                                    <p className="ml-1 flex-shrink-0 font-normal text-gray-500">
                                                        - Weight: {kpi.weight}%
                                                    </p>
                                                </div>
                                                <div className="mt-2 flex flex-col">
                                                    <div className="flex items-center text-sm text-gray-500">
                                                        {kpi.description}
                                                    </div>
                                                    {kpi.dataSource && (
                                                        <div className="text-xs text-blue-500 mt-1">
                                                            Source: {kpi.dataSource}
                                                        </div>
                                                    )}
                                                </div>
                                            </div>
                                            <div className="ml-4 flex-shrink-0">
                                                <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                                                    Target: {kpi.targetValue} {kpi.unit}
                                                </span>
                                            </div>
                                        </div>
                                        <div className="ml-5 flex-shrink-0">
                                            <button className="text-gray-400 hover:text-gray-500 mr-3">
                                                <Edit2 className="w-5 h-5" />
                                            </button>
                                            <button onClick={() => handleDeleteKPI(kpi.id)} className="text-red-400 hover:text-red-500">
                                                <Trash2 className="w-5 h-5" />
                                            </button>
                                        </div>
                                    </div>
                                </li>
                            ))
                        )}
                    </ul>
                )}
            </div>



            {/* Modal */}
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
                                    Add KPI
                                </h3>
                                <div className="space-y-4">
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Title</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.title}
                                            onChange={(e) => setNewKPI({ ...newKPI, title: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Description</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.description}
                                            onChange={(e) => setNewKPI({ ...newKPI, description: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Weight (%)</label>
                                        <input
                                            type="number"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.weight}
                                            onChange={(e) => setNewKPI({ ...newKPI, weight: e.target.value })}
                                        />
                                    </div>

                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Unit</label>
                                        <input
                                            type="text"
                                            placeholder="e.g., USD, Count, %"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.unit}
                                            onChange={(e) => setNewKPI({ ...newKPI, unit: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Target Value</label>
                                        <input
                                            type="number"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.targetValue}
                                            onChange={(e) => setNewKPI({ ...newKPI, targetValue: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Data Source</label>
                                        <input
                                            type="text"
                                            placeholder="e.g. Fiverr Dashboard"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={newKPI.dataSource}
                                            onChange={(e) => setNewKPI({ ...newKPI, dataSource: e.target.value })}
                                        />
                                    </div>

                                </div>
                            </div>
                            <div className="mt-5 sm:mt-6 sm:grid sm:grid-cols-2 sm:gap-3 sm:grid-flow-row-dense">
                                <button
                                    type="button"
                                    onClick={handleAddKPI}
                                    className="w-full inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-blue-600 text-base font-medium text-white hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 sm:col-start-2 sm:text-sm"
                                >
                                    Save
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setShowModal(false)}
                                    className="mt-3 w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 sm:mt-0 sm:col-start-1 sm:text-sm"
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

export default KPIFramework;
