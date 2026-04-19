import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Users, Phone, Mail, Briefcase, Award } from 'lucide-react';

const TeamOverview = () => {
    const [team, setTeam] = useState([]);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        fetchTeam();
    }, []);

    const fetchTeam = async () => {
        try {
            const res = await axios.get('http://localhost:5000/api/staff', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setTeam(res.data);
            setLoading(false);
        } catch (error) {
            console.error("Error fetching team data:", error);
            setLoading(false);
        }
    };

    if (loading) {
        return (
            <div className="flex justify-center items-center h-full min-h-[400px]">
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
            </div>
        );
    }

    return (
        <div className="p-6 bg-gray-50 min-h-screen">
            <h2 className="text-2xl font-bold text-gray-800 mb-6 flex items-center">
                <Users className="mr-2 text-blue-600" />
                Team Overview
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {team.map((member) => (
                    <div key={member.id} className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 flex flex-col items-center hover:shadow-md transition-shadow">
                        {member.profilePicture ? (
                            <img
                                src={`http://localhost:5000${member.profilePicture}`}
                                alt={member.username}
                                className="w-20 h-20 rounded-full object-cover border-4 border-white shadow-sm mb-4"
                            />
                        ) : (
                            <div className="w-20 h-20 bg-blue-100 rounded-full flex items-center justify-center text-blue-600 text-2xl font-bold mb-4 border-4 border-white shadow-sm">
                                {member.username.charAt(0).toUpperCase()}
                            </div>
                        )}

                        <h3 className="text-lg font-semibold text-gray-900 mb-1">{member.username}</h3>
                        <div className="px-3 py-1 bg-blue-50 text-blue-700 text-xs font-medium rounded-full mb-4">
                            {member.role}
                        </div>

                        <div className="w-full space-y-3 mt-2 border-t pt-4">
                            <div className="flex items-center text-sm text-gray-600">
                                <Mail className="w-4 h-4 mr-3 text-gray-400" />
                                <span className="truncate">{member.email}</span>
                            </div>

                            {member.mobileNumber && (
                                <div className="flex items-center text-sm text-gray-600">
                                    <Phone className="w-4 h-4 mr-3 text-gray-400" />
                                    <span>{member.mobileNumber}</span>
                                </div>
                            )}

                            {member.position && (
                                <div className="flex items-center text-sm text-gray-600">
                                    <Award className="w-4 h-4 mr-3 text-gray-400" />
                                    <span>{member.position}</span>
                                </div>
                            )}

                            {member.jobCategory && (
                                <div className="flex items-center text-sm text-gray-600">
                                    <Briefcase className="w-4 h-4 mr-3 text-gray-400" />
                                    <span>{member.jobCategory}</span>
                                </div>
                            )}
                        </div>
                    </div>
                ))}
            </div>

            {team.length === 0 && (
                <div className="text-center py-12 bg-white rounded-lg border border-gray-200">
                    <Users className="mx-auto h-12 w-12 text-gray-400" />
                    <h3 className="mt-2 text-sm font-medium text-gray-900">No team members found</h3>
                    <p className="mt-1 text-sm text-gray-500">Wait for the admin to invite users.</p>
                </div>
            )}
        </div>
    );
};

export default TeamOverview;
