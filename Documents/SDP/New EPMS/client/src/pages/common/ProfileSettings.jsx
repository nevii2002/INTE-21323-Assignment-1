import React, { useState, useEffect, useRef } from 'react';
import axios from 'axios';
import { UserCircle, Mail, Phone, Briefcase, Lock, Shield, CheckCircle, Save, Camera } from 'lucide-react';
import { useAuth } from '../../context/AuthContext';

const ProfileSettings = () => {
    const { user, updateUser } = useAuth();
    const [profile, setProfile] = useState(null);
    const [loading, setLoading] = useState(true);

    // Tab State
    const [activeTab, setActiveTab] = useState('profile');

    // Upload State
    const fileInputRef = useRef(null);
    const [uploadingAvatar, setUploadingAvatar] = useState(false);

    // Forms State
    const [editData, setEditData] = useState({
        username: '',
        mobileNumber: '',
        position: ''
    });
    const [passwordData, setPasswordData] = useState({
        currentPassword: '',
        newPassword: '',
        confirmPassword: ''
    });

    // Alert State
    const [message, setMessage] = useState(null);
    const [error, setError] = useState(null);

    useEffect(() => {
        fetchProfile();
    }, []);

    const fetchProfile = async () => {
        try {
            const res = await axios.get('http://localhost:5000/api/auth/profile', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setProfile(res.data);
            setEditData({
                username: res.data.username || '',
                mobileNumber: res.data.mobileNumber || '',
                position: res.data.position || ''
            });
            setLoading(false);
        } catch (error) {
            console.error("Error fetching profile:", error);
            setError("Failed to load profile data.");
            setLoading(false);
        }
    };

    const handleProfileSubmit = async (e) => {
        e.preventDefault();
        setMessage(null);
        setError(null);
        try {
            const res = await axios.put('http://localhost:5000/api/auth/profile', editData, {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setMessage(res.data.message);
            setProfile(res.data.user);
        } catch (error) {
            setError(error.response?.data?.message || 'Error updating profile');
        }
    };

    const handlePasswordSubmit = async (e) => {
        e.preventDefault();
        setMessage(null);
        setError(null);

        if (passwordData.newPassword !== passwordData.confirmPassword) {
            return setError("New passwords do not match.");
        }

        try {
            const res = await axios.put('http://localhost:5000/api/auth/change-password', {
                currentPassword: passwordData.currentPassword,
                newPassword: passwordData.newPassword
            }, {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setMessage(res.data.message);
            setPasswordData({ currentPassword: '', newPassword: '', confirmPassword: '' });
        } catch (error) {
            setError(error.response?.data?.message || 'Error changing password');
        }
    };

    const handleAvatarUpload = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('avatar', file);

        setUploadingAvatar(true);
        setError(null);
        setMessage(null);

        try {
            const res = await axios.post('http://localhost:5000/api/auth/profile-picture', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data',
                    Authorization: `Bearer ${localStorage.getItem('token')}`
                }
            });
            setMessage(res.data.message);
            setProfile(prev => ({ ...prev, profilePicture: res.data.profilePicture }));
            updateUser({ profilePicture: res.data.profilePicture });
        } catch (error) {
            setError(error.response?.data?.message || 'Error uploading profile picture');
        } finally {
            setUploadingAvatar(false);
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
        <div className="p-6 max-w-5xl mx-auto">
            <h2 className="text-2xl font-bold text-gray-800 mb-6 flex items-center">
                <UserCircle className="mr-3 text-blue-600" />
                Profile & Settings
            </h2>

            <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden">
                <div className="flex flex-col md:flex-row">

                    {/* Left Sidebar Menu */}
                    <div className="w-full md:w-64 bg-gray-50 border-r border-gray-200">
                        <div className="p-6 text-center border-b border-gray-200">
                            <div className="relative w-24 h-24 mx-auto mb-4 group">
                                {profile?.profilePicture ? (
                                    <img
                                        src={`http://localhost:5000${profile.profilePicture}`}
                                        alt="Profile"
                                        className="w-24 h-24 rounded-full object-cover border-4 border-white shadow-sm"
                                    />
                                ) : (
                                    <div className="w-24 h-24 bg-blue-100 rounded-full flex items-center justify-center text-blue-600 text-3xl font-bold border-4 border-white shadow-sm">
                                        {profile?.username?.charAt(0).toUpperCase()}
                                    </div>
                                )}

                                {/* Upload Overlay */}
                                <div
                                    className="absolute inset-0 bg-black bg-opacity-50 rounded-full flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity cursor-pointer flex-col"
                                    onClick={() => fileInputRef.current?.click()}
                                >
                                    {uploadingAvatar ? (
                                        <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-white"></div>
                                    ) : (
                                        <>
                                            <Camera className="w-6 h-6 text-white mb-1" />
                                            <span className="text-white text-xs font-medium">Update</span>
                                        </>
                                    )}
                                </div>
                                <input
                                    type="file"
                                    ref={fileInputRef}
                                    onChange={handleAvatarUpload}
                                    accept="image/*"
                                    className="hidden"
                                />
                            </div>
                            <h3 className="font-semibold text-gray-800 truncate">{profile?.username}</h3>
                            <p className="text-sm text-gray-500 truncate">{profile?.role}</p>
                        </div>
                        <nav className="p-4 space-y-2">
                            <button
                                onClick={() => { setActiveTab('profile'); setMessage(null); setError(null); }}
                                className={`w-full flex items-center px-4 py-2 text-sm font-medium rounded-md transition-colors ${activeTab === 'profile' ? 'bg-blue-100 text-blue-700' : 'text-gray-600 hover:bg-white hover:text-gray-900'
                                    }`}
                            >
                                <UserCircle className="w-4 h-4 mr-3" />
                                Edit Profile
                            </button>
                            <button
                                onClick={() => { setActiveTab('password'); setMessage(null); setError(null); }}
                                className={`w-full flex items-center px-4 py-2 text-sm font-medium rounded-md transition-colors ${activeTab === 'password' ? 'bg-blue-100 text-blue-700' : 'text-gray-600 hover:bg-white hover:text-gray-900'
                                    }`}
                            >
                                <Lock className="w-4 h-4 mr-3" />
                                Change Password
                            </button>
                        </nav>
                    </div>

                    {/* Right Content Area */}
                    <div className="flex-1 p-6 md:p-8">
                        {message && (
                            <div className="mb-6 p-4 bg-green-50 border-l-4 border-green-500 text-green-700 flex items-center">
                                <CheckCircle className="w-5 h-5 mr-2" />
                                {message}
                            </div>
                        )}
                        {error && (
                            <div className="mb-6 p-4 bg-red-50 border-l-4 border-red-500 text-red-700 flex items-center">
                                <Shield className="w-5 h-5 mr-2" />
                                {error}
                            </div>
                        )}

                        {activeTab === 'profile' && (
                            <div>
                                <h3 className="text-lg font-semibold text-gray-800 mb-6">Personal Information</h3>

                                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Email Address (Read-only)</label>
                                        <div className="relative">
                                            <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                <Mail className="h-4 w-4 text-gray-400" />
                                            </div>
                                            <input
                                                type="email"
                                                disabled
                                                value={profile?.email || ''}
                                                className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md bg-gray-100 text-gray-500 sm:text-sm"
                                            />
                                        </div>
                                    </div>

                                    <div>
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Job Category (Read-only)</label>
                                        <div className="relative">
                                            <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                <Briefcase className="h-4 w-4 text-gray-400" />
                                            </div>
                                            <input
                                                type="text"
                                                disabled
                                                value={profile?.jobCategory || ''}
                                                className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md bg-gray-100 text-gray-500 sm:text-sm"
                                            />
                                        </div>
                                    </div>
                                </div>

                                {profile?.jobDescription && (
                                    <div className="mb-6">
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Job Description</label>
                                        <div className="bg-gray-50 p-4 rounded-md border border-gray-200 text-sm text-gray-600 whitespace-pre-wrap">
                                            {profile.jobDescription}
                                        </div>
                                    </div>
                                )}

                                {profile?.responsibilities && (
                                    <div className="mb-8">
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Key Responsibilities</label>
                                        <div className="bg-gray-50 p-4 rounded-md border border-gray-200 text-sm text-gray-600 whitespace-pre-wrap">
                                            {profile.responsibilities}
                                        </div>
                                    </div>
                                )}

                                <form onSubmit={handleProfileSubmit} className="space-y-6">
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                        <div>
                                            <label className="block text-sm font-medium text-gray-700 mb-1">Full Name</label>
                                            <div className="relative">
                                                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                    <UserCircle className="h-4 w-4 text-gray-400" />
                                                </div>
                                                <input
                                                    type="text"
                                                    required
                                                    value={editData.username}
                                                    onChange={(e) => setEditData({ ...editData, username: e.target.value })}
                                                    className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                />
                                            </div>
                                        </div>

                                        <div>
                                            <label className="block text-sm font-medium text-gray-700 mb-1">Mobile Number</label>
                                            <div className="relative">
                                                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                    <Phone className="h-4 w-4 text-gray-400" />
                                                </div>
                                                <input
                                                    type="text"
                                                    value={editData.mobileNumber}
                                                    onChange={(e) => setEditData({ ...editData, mobileNumber: e.target.value })}
                                                    className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                    placeholder="+1 (555) 000-0000"
                                                />
                                            </div>
                                        </div>

                                        <div className="md:col-span-2">
                                            <label className="block text-sm font-medium text-gray-700 mb-1">Position / Job Title</label>
                                            <div className="relative">
                                                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                    <Briefcase className="h-4 w-4 text-gray-400" />
                                                </div>
                                                <input
                                                    type="text"
                                                    value={editData.position}
                                                    onChange={(e) => setEditData({ ...editData, position: e.target.value })}
                                                    className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                    placeholder="e.g. Senior Software Engineer"
                                                />
                                            </div>
                                        </div>
                                    </div>

                                    <div className="flex justify-end pt-4 border-t border-gray-100">
                                        <button
                                            type="submit"
                                            className="inline-flex items-center px-4 py-2 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500"
                                        >
                                            <Save className="w-4 h-4 mr-2" />
                                            Save Changes
                                        </button>
                                    </div>
                                </form>
                            </div>
                        )}

                        {activeTab === 'password' && (
                            <div className="max-w-md">
                                <h3 className="text-lg font-semibold text-gray-800 mb-2">Change Password</h3>
                                <p className="text-sm text-gray-500 mb-6">Ensure your account is using a long, random password to stay secure.</p>

                                <form onSubmit={handlePasswordSubmit} className="space-y-6">
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Current Password</label>
                                        <input
                                            type="password"
                                            required
                                            value={passwordData.currentPassword}
                                            onChange={(e) => setPasswordData({ ...passwordData, currentPassword: e.target.value })}
                                            className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                        />
                                    </div>

                                    <div>
                                        <label className="block text-sm font-medium text-gray-700 mb-1">New Password</label>
                                        <input
                                            type="password"
                                            required
                                            minLength="6"
                                            value={passwordData.newPassword}
                                            onChange={(e) => setPasswordData({ ...passwordData, newPassword: e.target.value })}
                                            className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                        />
                                    </div>

                                    <div>
                                        <label className="block text-sm font-medium text-gray-700 mb-1">Confirm New Password</label>
                                        <input
                                            type="password"
                                            required
                                            value={passwordData.confirmPassword}
                                            onChange={(e) => setPasswordData({ ...passwordData, confirmPassword: e.target.value })}
                                            className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                        />
                                    </div>

                                    <div className="flex justify-start">
                                        <button
                                            type="submit"
                                            className="inline-flex items-center px-4 py-2 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500"
                                        >
                                            Update Password
                                        </button>
                                    </div>
                                </form>
                            </div>
                        )}

                    </div>
                </div>
            </div>
        </div>
    );
};

export default ProfileSettings;
