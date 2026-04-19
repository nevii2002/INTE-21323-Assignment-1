import React, { useState, useEffect } from 'react';
import { Plus, Edit2, Trash2, Mail, User, Shield, Briefcase, Eye } from 'lucide-react';
import { useNavigate } from 'react-router-dom';
import api from '../../api/axios';

// Mock Data removed


const StaffManagement = () => {
    const [staff, setStaff] = useState([]);
    const [showModal, setShowModal] = useState(false);
    const [formData, setFormData] = useState({ username: '', email: '', password: '', role: 'Employee', jobCategory: 'Full time', jobDescription: '', responsibilities: '' });
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);
    const [isEditing, setIsEditing] = useState(false);
    const [currentUserId, setCurrentUserId] = useState(null);
    const navigate = useNavigate();

    useEffect(() => {
        const fetchStaff = async () => {
            setIsLoading(true);
            try {
                const response = await api.get('/staff');
                setStaff(response.data);
            } catch (err) {
                console.error('Failed to fetch staff:', err);
                setError('Failed to load staff list.');
            } finally {
                setIsLoading(false);
            }
        };

        fetchStaff();
    }, []);

    const handleAddStaff = async () => {
        if (!formData.username || !formData.email) {
            alert('Please fill in all required fields.');
            return;
        }

        try {
            const response = await api.post('/staff', formData);
            setStaff([...staff, response.data.user]);
            closeModal();
        } catch (err) {
            console.error('Failed to add staff:', err);
            alert(err.response?.data?.message || 'Failed to add staff member.');
        }
    };

    const handleUpdateStaff = async () => {
        if (!formData.username || !formData.email) {
            alert('Please fill in required fields.');
            return;
        }

        try {
            const response = await api.put(`/staff/${currentUserId}`, formData);
            // Update local state
            setStaff(staff.map(u => u.id === currentUserId ? response.data : u));
            closeModal();
            alert('Employee details updated successfully.');
        } catch (err) {
            console.error('Failed to update staff:', err);
            alert(err.response?.data?.message || 'Failed to update staff.');
        }
    };

    const handleDelete = async (id) => {
        if (!window.confirm('Are you sure you want to delete this employee?')) return;
        try {
            await api.delete(`/staff/${id}`);
            setStaff(staff.filter(u => u.id !== id));
        } catch (err) {
            console.error('Failed to delete staff:', err);
            alert('Failed to delete staff member.');
        }
    };

    const handleEdit = (person) => {
        setFormData({
            username: person.username,
            email: person.email,
            role: person.role,
            jobCategory: person.jobCategory,
            jobDescription: person.jobDescription || '',
            responsibilities: person.responsibilities || '',
            password: '' // Password not executed on update
        });
        setIsEditing(true);
        setCurrentUserId(person.id);
        setShowModal(true);
    };

    const openAddModal = () => {
        setFormData({ username: '', email: '', password: '', role: 'Employee', jobCategory: 'Full time', jobDescription: '', responsibilities: '' });
        setIsEditing(false);
        setCurrentUserId(null);
        setShowModal(true);
    };

    const closeModal = () => {
        setShowModal(false);
        setFormData({ username: '', email: '', password: '', role: 'Employee', jobCategory: 'Full time', jobDescription: '', responsibilities: '' });
        setIsEditing(false);
        setCurrentUserId(null);
    };

    const handleViewProfile = (id) => {
        navigate(`/admin/staff/${id}`);
    };

    return (
        <div>
            <div className="flex justify-between items-center mb-6">
                <div>
                    <h2 className="text-2xl font-bold text-gray-900">Staff Management</h2>
                    <p className="text-sm text-gray-500">Manage employee accounts and roles.</p>
                </div>
                <button
                    onClick={openAddModal}
                    className="bg-blue-600 text-white px-4 py-2 rounded-lg flex items-center hover:bg-blue-700"
                >
                    <Plus className="w-5 h-5 mr-2" />
                    Add Employee
                </button>
            </div>

            {error && <div className="mb-4 text-red-600 bg-red-50 p-3 rounded">{error}</div>}

            <div className="bg-white shadow overflow-hidden sm:rounded-lg">
                {isLoading ? (
                    <div className="p-6 text-center text-gray-500">Loading Staff...</div>
                ) : (
                    <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Contact</th>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Role</th>
                                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Job Category</th>
                                <th className="relative px-6 py-3">
                                    <span className="sr-only">Actions</span>
                                </th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                            {staff.map((person) => (
                                <tr key={person.id}>
                                    <td className="px-6 py-4 whitespace-nowrap">
                                        <div className="flex items-center">
                                            <div className="flex-shrink-0 h-10 w-10 bg-gray-100 rounded-full flex items-center justify-center">
                                                <User className="h-5 w-5 text-gray-500" />
                                            </div>
                                            <div className="ml-4">
                                                <div className="text-sm font-medium text-gray-900">{person.username}</div>
                                            </div>
                                        </div>
                                    </td>
                                    <td className="px-6 py-4 whitespace-nowrap">
                                        <div className="flex items-center text-sm text-gray-500">
                                            <Mail className="h-4 w-4 mr-2" />
                                            {person.email}
                                        </div>
                                    </td>
                                    <td className="px-6 py-4 whitespace-nowrap">
                                        <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full 
                                            ${person.role === 'Admin' ? 'bg-purple-100 text-purple-800' :
                                                person.role === 'Manager' ? 'bg-orange-100 text-orange-800' : 'bg-green-100 text-green-800'}`}>
                                            {person.role}
                                        </span>
                                    </td>
                                    <td className="px-6 py-4 whitespace-nowrap">
                                        <div className="flex items-center text-sm text-gray-500">
                                            <Briefcase className="h-4 w-4 mr-2" />
                                            {person.jobCategory}
                                        </div>
                                    </td>
                                    <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                                        <button onClick={() => handleViewProfile(person.id)} className="text-gray-500 hover:text-gray-900 mr-4" title="View Profile">
                                            <Eye className="w-5 h-5" />
                                        </button>
                                        <button onClick={() => handleEdit(person)} className="text-blue-600 hover:text-blue-900 mr-4" title="Edit">
                                            <Edit2 className="w-5 h-5" />
                                        </button>
                                        <button onClick={() => handleDelete(person.id)} className="text-red-600 hover:text-red-900" title="Delete">
                                            <Trash2 className="w-5 h-5" />
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                )}
            </div>

            {/* Modal */}
            {showModal && (
                <div className="fixed z-10 inset-0 overflow-y-auto">
                    <div className="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
                        <div className="fixed inset-0 transition-opacity" aria-hidden="true">
                            <div className="absolute inset-0 bg-gray-500 opacity-75"></div>
                        </div>
                        <div className="inline-block align-bottom bg-white rounded-lg px-4 pt-5 pb-4 text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full sm:p-6">
                            <div>
                                <h3 className="text-lg leading-6 font-medium text-gray-900 mb-4">{isEditing ? 'Edit Employee' : 'Add Employee'}</h3>
                                <div className="space-y-4">
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Full Name</label>
                                        <input
                                            type="text"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={formData.username}
                                            onChange={(e) => setFormData({ ...formData, username: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Email</label>
                                        <input
                                            type="email"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={formData.email}
                                            onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                                        />
                                    </div>
                                    {/* Password field removed for invite-only flow */}
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                            <label className="block text-sm font-medium text-gray-700">Role</label>
                                            <select
                                                className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                value={formData.role}
                                                onChange={(e) => setFormData({ ...formData, role: e.target.value })}
                                            >
                                                <option value="Employee">Employee</option>
                                                <option value="Manager">Manager</option>
                                                <option value="HR">HR</option>
                                                <option value="Admin">Admin</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-sm font-medium text-gray-700">Job Category</label>
                                            <select
                                                className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                value={formData.jobCategory}
                                                onChange={(e) => setFormData({ ...formData, jobCategory: e.target.value })}
                                            >
                                                <option value="Full time">Full time</option>
                                                <option value="Hourly">Hourly</option>
                                                <option value="Remote">Remote</option>
                                                <option value="Intern">Intern</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Job Description</label>
                                        <textarea
                                            rows="2"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={formData.jobDescription}
                                            onChange={(e) => setFormData({ ...formData, jobDescription: e.target.value })}
                                            placeholder="Overview of the role..."
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">Key Responsibilities</label>
                                        <textarea
                                            rows="3"
                                            className="mt-1 block w-full border border-gray-300 rounded-md shadow-sm py-2 px-3 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                            value={formData.responsibilities}
                                            onChange={(e) => setFormData({ ...formData, responsibilities: e.target.value })}
                                            placeholder="List of primary responsibilities..."
                                        />
                                    </div>
                                </div>
                            </div>
                            <div className="mt-5 sm:mt-6 sm:grid sm:grid-cols-2 sm:gap-3 sm:grid-flow-row-dense">
                                <button
                                    type="button"
                                    onClick={isEditing ? handleUpdateStaff : handleAddStaff}
                                    className="w-full inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-blue-600 text-base font-medium text-white hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 sm:col-start-2 sm:text-sm"
                                >
                                    {isEditing ? 'Update' : 'Save'}
                                </button>
                                <button
                                    type="button"
                                    onClick={closeModal}
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

export default StaffManagement;
