import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Star, Eye, Calendar, UserCheck } from 'lucide-react';

const MyEvaluations = () => {
    const [evaluations, setEvaluations] = useState([]);
    const [loading, setLoading] = useState(true);
    const [selectedEval, setSelectedEval] = useState(null); // For modal

    useEffect(() => {
        fetchEvaluations();
    }, []);

    const fetchEvaluations = async () => {
        try {
            const res = await axios.get('http://localhost:5000/api/evaluations', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setEvaluations(res.data);
        } catch (error) {
            console.error("Error fetching evaluations:", error);
        } finally {
            setLoading(false);
        }
    };

    // Helper to fetch details when clicking view
    const handleViewDetails = async (evalId) => {
        try {
            const res = await axios.get(`http://localhost:5000/api/evaluations/${evalId}`, {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setSelectedEval(res.data);
        } catch (error) {
            console.error("Error fetching details:", error);
        }
    };

    const closeModal = () => setSelectedEval(null);

    return (
        <div className="p-6 bg-gray-50 min-h-screen">
            <h2 className="text-2xl font-bold text-gray-800 mb-6 flex items-center">
                <UserCheck className="mr-2 text-blue-600" />
                My Performance Evaluations
            </h2>

            {loading ? (
                <div className="text-center py-10">Loading...</div>
            ) : evaluations.length === 0 ? (
                <div className="bg-white p-6 rounded-lg shadow text-center text-gray-500">
                    No evaluations found.
                </div>
            ) : (
                <div className="grid gap-4">
                    {evaluations.map(ev => (
                        <div key={ev.id} className="bg-white p-4 rounded-lg shadow-sm border border-gray-200 flex justify-between items-center transition hover:shadow-md">
                            <div className="flex items-center gap-4">
                                <div className="bg-blue-100 p-3 rounded-full text-blue-600">
                                    <Calendar className="w-6 h-6" />
                                </div>
                                <div>
                                    <h3 className="font-semibold text-gray-800 text-lg">Period: {ev.period}</h3>
                                    <p className="text-sm text-gray-500">Evaluator: {ev.evaluator ? ev.evaluator.username : 'Unknown'}</p>
                                </div>
                            </div>
                            <div className="flex items-center gap-4">
                                <span className={`px-3 py-1 rounded-full text-sm font-medium ${ev.status === 'Submitted' ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'
                                    }`}>
                                    {ev.status}
                                </span>
                                <button
                                    onClick={() => handleViewDetails(ev.id)}
                                    className="flex items-center bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded text-sm font-medium transition"
                                >
                                    <Eye className="w-4 h-4 mr-2" />
                                    View Details
                                </button>
                            </div>
                        </div>
                    ))}
                </div>
            )}

            {/* Detail Modal */}
            {selectedEval && (
                <div className="fixed inset-0 bg-black bg-opacity-50 flex justify-center items-center z-50 p-4">
                    <div className="bg-white rounded-lg shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
                        <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50 rounded-t-lg">
                            <div>
                                <h3 className="text-xl font-bold text-gray-900">Evaluation Details</h3>
                                <p className="text-sm text-gray-500">{selectedEval.period}</p>
                            </div>
                            <div className="flex items-center space-x-4">
                                <div className="bg-blue-100 text-blue-800 px-3 py-1 rounded-full text-sm font-bold border border-blue-200 shadow-sm">
                                    Summary Score: {(selectedEval.details?.reduce((sum, d) => sum + d.rating, 0) / (selectedEval.details?.length || 1)).toFixed(1)}/5
                                </div>
                                <button onClick={closeModal} className="text-gray-400 hover:text-gray-600 text-2xl">&times;</button>
                            </div>
                        </div>

                        <div className="p-6 space-y-6">
                            {/* Comments Section */}
                            {selectedEval.comments && (
                                <div className="bg-blue-50 p-4 rounded-lg border border-blue-100">
                                    <h4 className="font-semibold text-blue-900 mb-1">Manager Comments</h4>
                                    <p className="text-blue-800 text-sm italic">"{selectedEval.comments}"</p>
                                </div>
                            )}

                            {/* KPI Breakdown */}
                            <div>
                                <h4 className="font-bold text-gray-800 mb-4 border-b pb-2">KPI Breakdown</h4>
                                <div className="space-y-4">
                                    {selectedEval.details?.map(detail => (
                                        <div key={detail.id} className="bg-white border border-gray-200 p-4 rounded-lg">
                                            <div className="flex justify-between items-start mb-2">
                                                <div className="flex-1">
                                                    <h5 className="font-semibold text-gray-800">{detail.kpi?.title || `KPI #${detail.kpiId}`}</h5>
                                                    <p className="text-xs text-gray-500 mt-1">{detail.kpi?.description}</p>
                                                </div>
                                                <div className="flex items-center bg-gray-50 px-3 py-1 rounded-full border">
                                                    <Star className="w-4 h-4 text-yellow-400 fill-current mr-1" />
                                                    <span className="font-bold text-gray-700">{detail.rating}/5</span>
                                                </div>
                                            </div>
                                            {detail.comment && (
                                                <p className="text-sm text-gray-600 mt-2 border-l-2 border-gray-300 pl-2">
                                                    {detail.comment}
                                                </p>
                                            )}
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </div>

                        <div className="p-6 border-t border-gray-100 flex justify-end">
                            <button
                                onClick={closeModal}
                                className="bg-gray-200 hover:bg-gray-300 text-gray-800 font-medium px-6 py-2 rounded transition"
                            >
                                Close
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

export default MyEvaluations;
