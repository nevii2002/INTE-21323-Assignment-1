import React, { useState, useEffect } from 'react';
import StarRating from '../../components/StarRating';
import api from '../../api/axios';
import { useAuth } from '../../context/AuthContext';
import { useNavigate } from 'react-router-dom';

const PeerEvaluation = () => {
    const { user } = useAuth();
    const navigate = useNavigate();
    const [employees, setEmployees] = useState([]);
    const [selectedEmployee, setSelectedEmployee] = useState('');
    const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
    const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
    const [kpis, setKpis] = useState([]);
    const [ratings, setRatings] = useState({});
    const [globalComment, setGlobalComment] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [successMessage, setSuccessMessage] = useState('');
    const [errorMessage, setErrorMessage] = useState('');

    useEffect(() => {
        const initData = async () => {
            setIsLoading(true);
            try {
                // Fetch Staff and filter out the current user
                const staffRes = await api.get('/staff');
                setEmployees(staffRes.data.filter(u => u.id !== user.id));

                // Fetch Qualitative KPIs for peer evaluations
                const kpiRes = await api.get('/kpis');
                setKpis(kpiRes.data.filter(k => k.category === 'Qualitative'));
            } catch (err) {
                console.error('Failed to init evaluation data', err);
                setErrorMessage('Failed to load necessary data.');
            } finally {
                setIsLoading(false);
            }
        };
        if (user?.id) {
            initData();
        }
    }, [user]);

    const handleRatingChange = (kpiId, rating) => {
        setRatings(prev => ({
            ...prev,
            [kpiId]: rating
        }));
    };

    const handleSubmit = async () => {
        if (!selectedEmployee) return alert('Please select a peer.');

        setSuccessMessage('');
        setErrorMessage('');
        setIsLoading(true);

        const period = `${selectedYear}-${String(selectedMonth).padStart(2, '0')}`;

        const details = Object.entries(ratings).map(([kpiId, rating]) => ({
            kpiId: parseInt(kpiId),
            rating: rating,
            comment: ''
        }));

        const payload = {
            employeeId: selectedEmployee,
            type: 'Peer',
            period,
            details,
            comments: globalComment
        };

        try {
            await api.post('/evaluations', payload);
            setSuccessMessage('Peer Evaluation submitted successfully!');
            setRatings({});
            setGlobalComment('');
            setSelectedEmployee('');
        } catch (err) {
            console.error('Submit Error:', err);
            setErrorMessage(err.response?.data?.message || 'Failed to submit evaluation.');
        } finally {
            setIsLoading(false);
        }
    };

    if (isLoading && kpis.length === 0) return <div className="p-8 text-center">Loading...</div>;

    // Helper for back button depending on role
    const goBack = () => {
        const isManagerOrAdmin = user?.role === 'Manager' || user?.role === 'Admin';
        navigate(isManagerOrAdmin ? '/admin/360-evaluation' : '/user/360-evaluation');
    };

    return (
        <div className="bg-gray-50 min-h-screen py-8 px-4 font-sans">
            <div className="max-w-3xl mx-auto bg-white rounded-lg shadow-lg overflow-hidden">
                <div className="p-6 border-b border-gray-100 flex justify-between items-start">
                    <div>
                        <h2 className="text-xl font-bold text-gray-800">Peer Evaluation Form</h2>
                        <p className="mt-1 text-sm text-gray-500 font-medium text-purple-600">Provide constructive feedback to your peers</p>
                    </div>
                    <button onClick={goBack} className="text-sm font-medium text-gray-500 hover:text-gray-700">
                        &larr; Back to Hub
                    </button>
                </div>

                <div className="p-6">
                    {successMessage && <div className="mb-4 bg-green-100 text-green-700 p-3 rounded">{successMessage}</div>}
                    {errorMessage && <div className="mb-4 bg-red-100 text-red-700 p-3 rounded">{errorMessage}</div>}

                    {/* Selectors */}
                    <div className="flex flex-wrap gap-4 mb-6">
                        <select
                            value={selectedEmployee}
                            onChange={(e) => setSelectedEmployee(e.target.value)}
                            className="bg-purple-500 hover:bg-purple-600 text-white text-sm font-medium py-2 px-4 rounded-md outline-none cursor-pointer border-r-8 border-transparent"
                        >
                            <option value="" className="bg-white text-gray-800">Select Peer</option>
                            {employees.map(e => (
                                <option key={e.id} value={e.id} className="bg-white text-gray-800">{e.username} ({e.role})</option>
                            ))}
                        </select>

                        <select
                            value={selectedYear}
                            onChange={(e) => setSelectedYear(e.target.value)}
                            className="bg-purple-500 hover:bg-purple-600 text-white text-sm font-medium py-2 px-4 rounded-md outline-none cursor-pointer border-r-8 border-transparent"
                        >
                            <option value="2025" className="bg-white text-gray-800">2025</option>
                            <option value="2026" className="bg-white text-gray-800">2026</option>
                            <option value="2027" className="bg-white text-gray-800">2027</option>
                        </select>

                        <select
                            value={selectedMonth}
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="bg-purple-500 hover:bg-purple-600 text-white text-sm font-medium py-2 px-4 rounded-md outline-none cursor-pointer border-r-8 border-transparent"
                        >
                            {Array.from({ length: 12 }, (_, i) => {
                                const m = i + 1;
                                const name = new Date(2025, i, 1).toLocaleString('default', { month: 'long' });
                                return <option key={m} value={m} className="bg-white text-gray-800">{name}</option>
                            })}
                        </select>
                    </div>

                    <div className="space-y-6">
                        {kpis.length === 0 ? (
                            <p className="text-gray-500 text-center">No Qualitative KPIs found for evaluation.</p>
                        ) : (
                            kpis.map((kpi) => (
                                <div key={kpi.id} className="border border-gray-200 rounded-lg p-4">
                                    <div className="flex justify-between items-start mb-2">
                                        <div>
                                            <h3 className="font-bold text-gray-800 text-md">{kpi.title}</h3>
                                        </div>
                                        <StarRating
                                            rating={ratings[kpi.id] || 0}
                                            onRatingChange={(r) => handleRatingChange(kpi.id, r)}
                                        />
                                    </div>
                                    <p className="text-sm text-gray-600 mb-0">{kpi.description}</p>
                                </div>
                            ))
                        )}

                        <div className="border border-gray-200 rounded-lg p-4">
                            <h3 className="font-bold text-gray-800 text-md mb-2">Peer Comments</h3>
                            <textarea
                                rows={3}
                                className="w-full border border-gray-300 rounded-md p-2 text-sm focus:ring-purple-500 focus:border-purple-500 outline-none"
                                placeholder="Share your positive feedback and areas for growth for this peer..."
                                value={globalComment}
                                onChange={(e) => setGlobalComment(e.target.value)}
                            />
                        </div>

                        <div className="flex justify-end pt-4">
                            <button
                                onClick={handleSubmit}
                                disabled={isLoading}
                                className="bg-purple-500 hover:bg-purple-600 text-white font-bold py-2 px-6 rounded-lg shadow transition duration-200 flex items-center disabled:bg-gray-400"
                            >
                                {isLoading ? 'Submitting...' : 'Submit Peer Evaluation'}
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default PeerEvaluation;
