import React, { useState } from 'react';
import StarRating from '../../components/StarRating';
import api from '../../api/axios';

const SelfEvaluation = () => {
    // Hardcoded KPIs to match the specific requirement
    const specificKPIs = [
        { id: 1, title: "Achieving Project Deadlines", description: "Consistently meets project deadlines and delivers tasks on time, demonstrating strong time management skills." },
        { id: 2, title: "Quality of Work Product", description: "Produces high-quality work that is accurate, thorough, and meets expectations, contributing to overall project success." },
        { id: 3, title: "Communication & Collaboration", description: "Effectively communicates ideas and information to team members and stakeholders, fostering a collaborative and productive work environment." },
        { id: 4, title: "Problem-Solving & Innovation", description: "Identifies challenges proactively, analyzes complex situations, and develops creative, effective solutions, driving continuous improvement." },
        { id: 5, title: "Adaptability & Learning", description: "Adapts quickly to new technologies, processes, and changing priorities, demonstrating a strong willingness to learn and grow professionally." }
    ];

    const [ratings, setRatings] = useState({});
    const [globalComment, setGlobalComment] = useState('');
    const [submitted, setSubmitted] = useState(false);
    const [loading, setLoading] = useState(false);

    const handleRatingChange = (kpiId, rating) => {
        setRatings(prev => ({
            ...prev,
            [kpiId]: rating
        }));
    };

    const calculateAverage = () => {
        const values = Object.values(ratings);
        if (values.length === 0) return 0;
        const sum = values.reduce((a, b) => a + b, 0);
        return (sum / specificKPIs.length).toFixed(1); // Divide by total KPIs to treat unrated as 0? Or divide by rated? Usually by total for "final mark". Let's do rated to be nicer, or total to be accurate. Let's do rated for now as form validates later? No, usually out of 5 for the whole form. Let's do: sum / 5 (total items).
    };

    const currentScore = calculateAverage();

    const handleSubmit = async () => {
        setLoading(true);
        // Simulate API call for now or map to backend structure
        setTimeout(() => {
            setLoading(false);
            setSubmitted(true);
        }, 1000);
    };

    if (submitted) {
        return (
            <div className="flex justify-center items-center h-screen bg-gray-50">
                <div className="bg-white p-8 rounded-lg shadow-lg text-center max-w-md w-full border-t-4 border-blue-500">
                    <h2 className="text-2xl font-bold text-gray-800 mb-2">Evaluation Submitted</h2>
                    <p className="text-gray-600 mb-4">Your self-evaluation has been successfully recorded.</p>

                    <div className="bg-blue-50 rounded-lg p-4 mb-6">
                        <p className="text-sm text-blue-600 font-semibold uppercase tracking-wider">Your Final Mark</p>
                        <p className="text-4xl font-extrabold text-blue-900">{currentScore} <span className="text-lg text-blue-400">/ 5.0</span></p>
                    </div>

                    <button onClick={() => setSubmitted(false)} className="mt-2 text-blue-600 hover:text-blue-800 font-medium">Return to Form</button>
                </div>
            </div>
        );
    }

    return (
        <div className="bg-gray-50 min-h-screen py-8 px-4 font-sans">
            <div className="max-w-3xl mx-auto bg-white rounded-lg shadow-lg overflow-hidden">
                <div className="p-6 border-b border-gray-100 flex justify-between items-center">
                    <div>
                        <h2 className="text-xl font-bold text-gray-800">Employee Self-Evaluation Form</h2>
                        <p className="mt-1 text-sm text-gray-500">Rate your performance on the displayed KPIs.</p>
                    </div>
                    <div className="text-right">
                        <p className="text-xs text-gray-400 uppercase font-bold">Current Score</p>
                        <p className="text-2xl font-bold text-indigo-600">{currentScore}</p>
                    </div>
                </div>

                <div className="p-6 space-y-6">
                    {specificKPIs.map((kpi) => (
                        <div key={kpi.id} className="border border-gray-200 rounded-lg p-4">
                            <div className="flex justify-between items-start mb-2">
                                <h3 className="font-bold text-gray-800 text-md">{kpi.title}</h3>
                                <StarRating
                                    rating={ratings[kpi.id] || 0}
                                    onRatingChange={(r) => handleRatingChange(kpi.id, r)}
                                />
                            </div>
                            <p className="text-sm text-gray-600 mb-0">{kpi.description}</p>
                        </div>
                    ))}

                    <div className="border border-gray-200 rounded-lg p-4">
                        <h3 className="font-bold text-gray-800 text-md mb-2">Additional Comments</h3>
                        <p className="text-xs text-gray-500 mb-2">Overall, I believe my performance has been strong, particularly in problem-solving. I aim to improve my project deadline consistency in the next quarter.</p>
                        <textarea
                            rows={3}
                            className="w-full border border-gray-300 rounded-md p-2 text-sm focus:ring-blue-500 focus:border-blue-500 outline-none"
                            placeholder="Type your comments here..."
                            value={globalComment}
                            onChange={(e) => setGlobalComment(e.target.value)}
                        />
                    </div>

                    <div className="flex justify-between items-center pt-4 border-t border-gray-100">
                        <div>
                            <span className="text-sm font-medium text-gray-500">Final Mark: </span>
                            <span className="text-xl font-bold text-gray-900">{currentScore} / 5.0</span>
                        </div>
                        <button
                            onClick={handleSubmit}
                            disabled={loading}
                            className="bg-blue-500 hover:bg-blue-600 text-white font-bold py-2 px-6 rounded-lg shadow transition duration-200 flex items-center"
                        >
                            {loading ? 'Submitting...' : 'Submit Report'}
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default SelfEvaluation;
