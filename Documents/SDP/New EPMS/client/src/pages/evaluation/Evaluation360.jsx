import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../../context/AuthContext';
import { UserCheck, Users, Briefcase } from 'lucide-react';

const Evaluation360 = () => {
    const navigate = useNavigate();
    const { user } = useAuth();

    // Base route depending on role (e.g. /admin or /user)
    const isManagerOrAdmin = user?.role === 'Manager' || user?.role === 'Admin';
    const basePath = isManagerOrAdmin ? '/admin' : '/user';

    return (
        <div className="p-6 max-w-6xl mx-auto">
            <div className="mb-8">
                <h1 className="text-3xl font-bold text-gray-900">360° Evaluation Hub</h1>
                <p className="mt-2 text-gray-600">Access your Self, Peer, and Manager evaluations. Continuous feedback helps us all grow.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">

                {/* Self Evaluation Card */}
                <div
                    onClick={() => navigate(`${basePath}/self-evaluation`)}
                    className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 cursor-pointer hover:shadow-md hover:border-blue-300 transition-all flex flex-col items-center text-center group"
                >
                    <div className="w-16 h-16 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                        <UserCheck className="w-8 h-8" />
                    </div>
                    <h3 className="text-xl font-bold text-gray-800 mb-2">Self Evaluation</h3>
                    <p className="text-sm text-gray-500">Reflect on your own performance, achievements, and areas for improvement.</p>
                </div>

                {/* Peer Evaluation Card */}
                <div
                    onClick={() => navigate(`${basePath}/peer-evaluation`)}
                    className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 cursor-pointer hover:shadow-md hover:border-purple-300 transition-all flex flex-col items-center text-center group"
                >
                    <div className="w-16 h-16 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                        <Users className="w-8 h-8" />
                    </div>
                    <h3 className="text-xl font-bold text-gray-800 mb-2">Peer Evaluation</h3>
                    <p className="text-sm text-gray-500">Provide constructive feedback to your team members and peers.</p>
                </div>

                {/* Manager Evaluation Card (Only visible to Managers/Admins) */}
                {isManagerOrAdmin && (
                    <div
                        onClick={() => navigate('/admin/manager-evaluation')}
                        className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 cursor-pointer hover:shadow-md hover:border-green-300 transition-all flex flex-col items-center text-center group"
                    >
                        <div className="w-16 h-16 rounded-full bg-green-50 text-green-600 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                            <Briefcase className="w-8 h-8" />
                        </div>
                        <h3 className="text-xl font-bold text-gray-800 mb-2">Manager Evaluation</h3>
                        <p className="text-sm text-gray-500">Evaluate the performance of your direct reports and subordinates.</p>
                    </div>
                )}
            </div>
        </div>
    );
};

export default Evaluation360;
