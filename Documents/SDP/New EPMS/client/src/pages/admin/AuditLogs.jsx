import React, { useState, useEffect } from 'react';
import axios from 'axios';

const AuditLogs = () => {
    const [logs, setLogs] = useState([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);

    // Filters
    const [actionFilter, setActionFilter] = useState('');
    const [resourceFilter, setResourceFilter] = useState('');

    const fetchLogs = async () => {
        try {
            setLoading(true);
            const token = localStorage.getItem('token');

            const config = {
                headers: { Authorization: `Bearer ${token}` },
                params: {}
            };

            if (actionFilter) config.params.action = actionFilter;
            if (resourceFilter) config.params.resource = resourceFilter;

            const response = await axios.get('http://localhost:5000/api/audit-logs', config);
            setLogs(response.data);
            setLoading(false);
        } catch (err) {
            console.error('Failed to fetch audit logs', err);
            setError('Failed to fetch audit logs');
            setLoading(false);
        }
    };

    useEffect(() => {
        fetchLogs();
    }, [actionFilter, resourceFilter]);

    const formatDate = (dateString) => {
        return new Date(dateString).toLocaleString();
    };

    return (
        <div className="p-6">
            <h1 className="text-2xl font-bold mb-4 text-gray-800">System Audit Logs</h1>
            <p className="text-gray-600 mb-6">Track critical administrative actions in the system.</p>

            {/* Filters */}
            <div className="flex gap-4 mb-6 bg-white p-4 rounded shadow-sm border border-gray-100">
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Action</label>
                    <select
                        className="border border-gray-300 rounded p-2 text-sm w-48"
                        value={actionFilter}
                        onChange={(e) => setActionFilter(e.target.value)}
                    >
                        <option value="">All Actions</option>
                        <option value="CREATE">CREATE</option>
                        <option value="UPDATE">UPDATE</option>
                        <option value="DELETE">DELETE</option>
                    </select>
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Resource</label>
                    <select
                        className="border border-gray-300 rounded p-2 text-sm w-48"
                        value={resourceFilter}
                        onChange={(e) => setResourceFilter(e.target.value)}
                    >
                        <option value="">All Resources</option>
                        <option value="User">User</option>
                        <option value="KPI">KPI</option>
                        <option value="EmployeeKPIs">EmployeeKPIs</option>
                    </select>
                </div>
                <div className="flex items-end">
                    <button
                        onClick={fetchLogs}
                        className="bg-blue-600 text-white px-4 py-2 rounded text-sm hover:bg-blue-700"
                    >
                        Refresh
                    </button>
                </div>
            </div>

            {/* Logs Table */}
            {loading ? (
                <div className="flex justify-center p-10">
                    <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                </div>
            ) : error ? (
                <div className="bg-red-50 text-red-600 p-4 rounded border border-red-200">
                    {error}
                </div>
            ) : logs.length === 0 ? (
                <div className="bg-gray-50 text-center p-10 rounded text-gray-500 border border-gray-200">
                    No audit logs found matching criteria.
                </div>
            ) : (
                <div className="bg-white rounded shadow-sm overflow-hidden border border-gray-200">
                    <div className="overflow-x-auto">
                        <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50">
                                <tr>
                                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Date & Time</th>
                                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">User (ID)</th>
                                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Action</th>
                                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Resource</th>
                                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Details</th>
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                                {logs.map((log) => (
                                    <tr key={log.id} className="hover:bg-gray-50">
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                                            {formatDate(log.createdAt)}
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                                            {log.user ? `${log.user.username} (#${log.user.id})` : `System (#${log.userId})`}
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm">
                                            <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full 
                        ${log.action === 'CREATE' ? 'bg-green-100 text-green-800' : ''}
                        ${log.action === 'UPDATE' ? 'bg-blue-100 text-blue-800' : ''}
                        ${log.action === 'DELETE' ? 'bg-red-100 text-red-800' : ''}
                      `}>
                                                {log.action}
                                            </span>
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                                            {log.resource}
                                        </td>
                                        <td className="px-6 py-4 text-sm text-gray-500 max-w-xs truncate">
                                            <div className="font-mono text-xs bg-gray-50 p-1 rounded border border-gray-100 break-all overflow-hidden" title={log.details}>
                                                {log.details || '-'}
                                            </div>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            )}
        </div>
    );
};

export default AuditLogs;
