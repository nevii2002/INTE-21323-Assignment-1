import React, { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { User, Mail, Briefcase, Phone, MapPin, Calendar, ArrowLeft, DollarSign, Award, Trash2, Save } from 'lucide-react';
import api from '../../api/axios';

// ------ Editable Job Details sub-component ------
const EditableJobDetails = ({ staffId, staff, onSaved }) => {
    const [form, setForm] = useState({
        position: staff.position || '',
        department: staff.department || '',
        jobCategory: staff.jobCategory || 'Full time',
        startDate: staff.startDate || '',
        basicSalary: staff.basicSalary || '',
        jobDescription: staff.jobDescription || '',
        responsibilities: staff.responsibilities || '',
    });
    const [saving, setSaving] = useState(false);
    const [msg, setMsg] = useState('');

    const handleChange = (e) => setForm({ ...form, [e.target.name]: e.target.value });

    const handleSave = async () => {
        setSaving(true);
        try {
            const res = await api.put(`/staff/${staffId}`, form);
            onSaved(res.data.user || form);
            setMsg('Saved successfully!');
            setTimeout(() => setMsg(''), 3000);
        } catch (e) {
            setMsg('Failed to save. Please try again.');
        } finally { setSaving(false); }
    };

    return (
        <div className="space-y-4">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Position / Job Title</label>
                    <input name="position" value={form.position} onChange={handleChange} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" placeholder="e.g. Sales Representative" />
                </div>
                <div>
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Department</label>
                    <input name="department" value={form.department} onChange={handleChange} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" placeholder="e.g. Operations" />
                </div>
                <div>
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Employment Type</label>
                    <select name="jobCategory" value={form.jobCategory} onChange={handleChange} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm">
                        <option>Full time</option>
                        <option>Hourly</option>
                        <option>Remote</option>
                        <option>Intern</option>
                    </select>
                </div>
                <div>
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Start Date</label>
                    <input type="date" name="startDate" value={form.startDate} onChange={handleChange} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" />
                </div>
                <div>
                    <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Basic Salary (LKR)</label>
                    <input type="number" name="basicSalary" value={form.basicSalary} onChange={handleChange} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" placeholder="0.00" />
                </div>
            </div>
            <div>
                <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Job Description</label>
                <textarea name="jobDescription" value={form.jobDescription} onChange={handleChange} rows={3} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" placeholder="Describe the primary purpose of this role..." />
            </div>
            <div>
                <label className="block text-xs font-medium text-gray-500 uppercase tracking-wider mb-1">Key Responsibilities</label>
                <textarea name="responsibilities" value={form.responsibilities} onChange={handleChange} rows={3} className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm" placeholder="List main responsibilities, one per line..." />
            </div>
            <div className="flex items-center justify-between pt-2">
                {msg && <span className={`text-sm font-medium ${msg.includes('Failed') ? 'text-red-600' : 'text-green-600'}`}>{msg}</span>}
                <button onClick={handleSave} disabled={saving} className="ml-auto flex items-center bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 text-white px-4 py-2 rounded-md text-sm font-medium">
                    <Save className="w-4 h-4 mr-2" />{saving ? 'Saving...' : 'Save Job Details'}
                </button>
            </div>
        </div>
    );
};
// ------------------------------------------------

const StaffProfile = () => {
    const { id } = useParams();
    const navigate = useNavigate();
    const [staff, setStaff] = useState(null);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);
    const [assignedKpis, setAssignedKpis] = useState([]);
    const [allKpis, setAllKpis] = useState([]);
    const [activeTab, setActiveTab] = useState('Profile');
    const [isSaving, setIsSaving] = useState(false);
    const [salary, setSalary] = useState(0);
    const [bonuses, setBonuses] = useState([]);
    const [newBonus, setNewBonus] = useState({ amount: '', reason: '' });


    useEffect(() => {
        const fetchStaffDetails = async () => {
            try {
                const response = await api.get(`/staff/${id}`);
                setStaff({
                    ...response.data,
                    joinDate: new Date(response.data.createdAt).toLocaleDateString()
                });
                setSalary(response.data.basicSalary || 0);


                // Fetch Assigned KPIs
                const kpiResponse = await api.get(`/staff/${id}/kpis`);
                // Fetch All KPIs (to select from)
                const allKpiResponse = await api.get(`/kpis`);

                setAllKpis(allKpiResponse.data);

                // Merge assigned weight with KPI data
                const assignments = kpiResponse.data.map(ak => ({
                    kpiId: ak.id,
                    weight: ak.EmployeeKPI?.customWeight || ak.weight,

                }));
                // Wait, the backend returns KPIs with 'EmployeeKPI' join data?
                // Let's check controller: `through: { attributes: ['customWeight'] }`
                // Yes, standard Sequelize many-to-many return structure.

                setAssignedKpis(assignments);

                // Fetch Bonuses
                const bonusResponse = await api.get(`/bonuses/employee/${id}`);
                setBonuses(bonusResponse.data);

            } catch (err) {
                console.error('Failed to fetch staff details:', err);
                setError('Failed to load profile.');
            } finally {
                setIsLoading(false);
            }
        };

        fetchStaffDetails();
    }, [id]);

    const handleWeightChange = (kpiId, weight) => {
        const newAssignments = [...assignedKpis];
        const index = newAssignments.findIndex(a => a.kpiId === kpiId);
        if (index > -1) {
            newAssignments[index].weight = parseFloat(weight);
        } else {
            // Add if not exists (checked)
            newAssignments.push({ kpiId, weight: parseFloat(weight) });
        }
        setAssignedKpis(newAssignments);
    };

    const toggleKpi = (kpiId, defaultWeight) => {
        const index = assignedKpis.findIndex(a => a.kpiId === kpiId);
        if (index > -1) {
            // Remove
            setAssignedKpis(assignedKpis.filter(a => a.kpiId !== kpiId));
        } else {
            // Add with default weight
            setAssignedKpis([...assignedKpis, { kpiId, weight: defaultWeight }]);
        }
    };

    const handleSaveAssignments = async () => {
        // Validation: Sum must be exactly 100% for Evaluation KPIs only
        const evalKpis = assignedKpis.filter(a => {
            const kpi = allKpis.find(k => k.id === a.kpiId);
            return kpi && kpi.type === 'EVALUATION';
        });

        if (evalKpis.length > 0) {
            const totalEvalWeight = evalKpis.reduce((sum, item) => sum + (parseFloat(item.weight) || 0), 0);
            if (Math.abs(totalEvalWeight - 100) > 0.01) {
                alert(`Evaluation KPI weights must sum to exactly 100%. Current total: ${totalEvalWeight}%`);
                return;
            }
        }

        setIsSaving(true);
        try {
            await api.post(`/staff/${id}/kpis`, { assignments: assignedKpis });
            alert('Assignments updated successfully!');
        } catch (err) {
            console.error('Failed to save assignments:', err);
            // Display specific validation error from backend if available
            alert(err.response?.data?.message || 'Failed to save assignments.');
        } finally {
            setIsSaving(false);
        }
    };

    const handleAddBonus = async (e) => {
        e.preventDefault();
        if (!newBonus.amount || !newBonus.reason) return;

        try {
            const response = await api.post('/bonuses', {
                employeeId: id,
                amount: parseFloat(newBonus.amount),
                reason: newBonus.reason
            });
            setBonuses([response.data.bonus, ...bonuses]);
            setNewBonus({ amount: '', reason: '' });
            alert('Bonus awarded successfully!');
        } catch (err) {
            console.error('Failed to add bonus:', err);
            alert('Failed to award bonus.');
        }
    };

    const handleDeleteBonus = async (bonusId) => {
        if (!window.confirm('Are you sure you want to remove this bonus record?')) return;
        try {
            await api.delete(`/bonuses/${bonusId}`);
            setBonuses(bonuses.filter(b => b.id !== bonusId));
        } catch (err) {
            console.error('Failed to delete bonus:', err);
            alert('Failed to delete bonus.');
        }
    };

    if (isLoading) return <div className="p-6 text-center text-gray-500">Loading Profile...</div>;
    if (error) return <div className="p-6 text-center text-red-500">{error}</div>;
    if (!staff) return <div className="p-6 text-center text-gray-500">User not found</div>;

    const quantitativeKpis = allKpis.filter(k => k.category === 'Quantitative');

    return (
        <div>
            <div className="mb-6">
                <button
                    onClick={() => navigate('/admin/staff')}
                    className="flex items-center text-gray-600 hover:text-gray-900"
                >
                    <ArrowLeft className="w-5 h-5 mr-2" />
                    Back to Staff List
                </button>
            </div>

            <div className="bg-white shadow overflow-hidden sm:rounded-lg">
                <div className="px-4 py-5 sm:px-6 flex justify-between items-center border-b border-gray-200">
                    <div>
                        <h3 className="text-lg leading-6 font-medium text-gray-900">Employee Profile</h3>
                        <p className="mt-1 max-w-2xl text-sm text-gray-500">{staff.username}</p>
                    </div>
                    <div className="flex space-x-4">
                        <button
                            onClick={() => setActiveTab('Profile')}
                            className={`px-3 py-2 text-sm font-medium rounded-md ${activeTab === 'Profile' ? 'bg-blue-100 text-blue-700' : 'text-gray-500 hover:text-gray-700'}`}
                        >
                            Profile Details
                        </button>
                        <button
                            onClick={() => setActiveTab('KPIs')}
                            className={`px-3 py-2 text-sm font-medium rounded-md ${activeTab === 'KPIs' ? 'bg-blue-100 text-blue-700' : 'text-gray-500 hover:text-gray-700'}`}
                        >
                            KPI Assignments
                        </button>
                        <button
                            onClick={() => setActiveTab('Bonuses')}
                            className={`px-3 py-2 text-sm font-medium rounded-md ${activeTab === 'Bonuses' ? 'bg-blue-100 text-blue-700' : 'text-gray-500 hover:text-gray-700'}`}
                        >
                            Bonuses
                        </button>
                    </div>
                </div>

                {activeTab === 'Profile' ? (
                    <div className="border-t border-gray-200 px-4 py-5 sm:p-6">
                        {/* Personal Info */}
                        <h4 className="text-base font-semibold text-gray-800 mb-4 pb-2 border-b">Personal Information</h4>
                        <dl className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-4 mb-8">
                            <div><dt className="text-xs font-medium text-gray-500 uppercase tracking-wider flex items-center"><User className="w-3 h-3 mr-1" />Full Name</dt><dd className="mt-1 text-sm font-semibold text-gray-900">{staff.username}</dd></div>
                            <div><dt className="text-xs font-medium text-gray-500 uppercase tracking-wider flex items-center"><Mail className="w-3 h-3 mr-1" />Email</dt><dd className="mt-1 text-sm font-semibold text-gray-900">{staff.email}</dd></div>
                            <div><dt className="text-xs font-medium text-gray-500 uppercase tracking-wider flex items-center"><Phone className="w-3 h-3 mr-1" />Mobile</dt><dd className="mt-1 text-sm font-semibold text-gray-900">{staff.mobileNumber || '—'}</dd></div>
                            <div><dt className="text-xs font-medium text-gray-500 uppercase tracking-wider flex items-center"><Calendar className="w-3 h-3 mr-1" />Account Created</dt><dd className="mt-1 text-sm font-semibold text-gray-900">{staff.joinDate}</dd></div>
                        </dl>

                        {/* Job Info */}
                        <h4 className="text-base font-semibold text-gray-800 mb-4 pb-2 border-b">Job Details</h4>
                        <EditableJobDetails staffId={id} staff={staff} onSaved={(updated) => setStaff(s => ({ ...s, ...updated }))} />
                    </div>
                ) : activeTab === 'KPIs' ? (

                    <div className="p-6">
                        {/* Section 1: Evaluation KPIs */}
                        <div className="mb-6">
                            <h4 className="text-lg font-medium text-gray-900 border-b pb-2 mb-4">Performance Evaluation KPIs</h4>
                            <p className="text-sm text-gray-500 mb-4">Assign weights for performance review. Total must equal 100%.</p>

                            <div className="space-y-4">
                                {allKpis.filter(k => k.type === 'EVALUATION').length === 0 ? (
                                    <p className="text-gray-500 italic">No Evaluation KPIs defined.</p>
                                ) : (
                                    allKpis.filter(k => k.type === 'EVALUATION').map(kpi => {
                                        const isAssigned = assignedKpis.some(a => a.kpiId === kpi.id);
                                        const currentWeight = assignedKpis.find(a => a.kpiId === kpi.id)?.weight || 0;

                                        return (
                                            <div key={kpi.id} className={`flex items-center justify-between p-4 rounded-lg border ${isAssigned ? 'border-blue-500 bg-blue-50' : 'border-gray-200'}`}>
                                                <div className="flex items-center">
                                                    <input
                                                        type="checkbox"
                                                        checked={isAssigned}
                                                        onChange={() => toggleKpi(kpi.id, kpi.weight)}
                                                        className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mr-4"
                                                    />
                                                    <div>
                                                        <p className="text-sm font-medium text-gray-900">{kpi.title}</p>
                                                        <p className="text-xs text-gray-500">{kpi.description}</p>
                                                        <p className="text-xs text-gray-400">Default: {kpi.weight}% | Target: {kpi.targetValue} {kpi.unit}</p>
                                                    </div>
                                                </div>
                                                {isAssigned && (
                                                    <div className="flex items-center">
                                                        <label className="text-sm text-gray-700 mr-2">Weight (%):</label>
                                                        <input
                                                            type="number"
                                                            value={currentWeight}
                                                            onChange={(e) => handleWeightChange(kpi.id, e.target.value)}
                                                            className="w-20 border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                                                        />
                                                    </div>
                                                )}
                                            </div>
                                        );
                                    })
                                )}
                            </div>

                            <div className="mt-4 flex justify-end">
                                <span className={`text-sm font-bold ${assignedKpis
                                    .filter(a => allKpis.find(k => k.id === a.kpiId)?.type === 'EVALUATION')
                                    .reduce((s, i) => s + (parseFloat(i.weight) || 0), 0) === 100
                                    ? 'text-green-600'
                                    : 'text-red-600'
                                    }`}>
                                    Total Evaluation Weight: {assignedKpis
                                        .filter(a => allKpis.find(k => k.id === a.kpiId)?.type === 'EVALUATION')
                                        .reduce((s, i) => s + (parseFloat(i.weight) || 0), 0)}%
                                </span>
                            </div>
                        </div>

                        {/* Section 2: Bonus Projects */}
                        <div className="mb-6 pt-6 border-t border-gray-200">
                            <h4 className="text-lg font-medium text-gray-900 border-b pb-2 mb-4">Bonus Projects Assignments</h4>
                            <p className="text-sm text-gray-500 mb-4">Assign this employee to specific revenue projects. Weight represents their share of the project bonus pool.</p>

                            <div className="space-y-4">
                                {allKpis.filter(k => k.type === 'BONUS').length === 0 ? (
                                    <p className="text-gray-500 italic">No Bonus Projects defined.</p>
                                ) : (
                                    allKpis.filter(k => k.type === 'BONUS').map(kpi => {
                                        const isAssigned = assignedKpis.some(a => a.kpiId === kpi.id);
                                        const currentWeight = assignedKpis.find(a => a.kpiId === kpi.id)?.weight || 0;

                                        return (
                                            <div key={kpi.id} className={`flex items-center justify-between p-4 rounded-lg border ${isAssigned ? 'border-purple-500 bg-purple-50' : 'border-gray-200'}`}>
                                                <div className="flex items-center">
                                                    <input
                                                        type="checkbox"
                                                        checked={isAssigned}
                                                        onChange={() => toggleKpi(kpi.id, 0)} // Default 0 for bonus projects to force manual entry? Or keep default
                                                        className="h-4 w-4 text-purple-600 focus:ring-purple-500 border-gray-300 rounded mr-4"
                                                    />
                                                    <div>
                                                        <p className="text-sm font-medium text-gray-900">{kpi.title}</p>
                                                        <p className="text-xs text-gray-500">{kpi.unit === 'Currency' ? 'Revenue Based' : 'Task Based'}</p>
                                                    </div>
                                                </div>
                                                {isAssigned && (
                                                    <div className="flex items-center">
                                                        <label className="text-sm text-gray-700 mr-2">Share (%):</label>
                                                        <input
                                                            type="number"
                                                            value={currentWeight}
                                                            onChange={(e) => handleWeightChange(kpi.id, e.target.value)}
                                                            className="w-20 border-gray-300 rounded-md shadow-sm focus:ring-purple-500 focus:border-purple-500 sm:text-sm"
                                                        />
                                                    </div>
                                                )}
                                            </div>
                                        );
                                    })
                                )}
                            </div>
                        </div>

                        <div className="mt-6 flex justify-end border-t pt-4">
                            <button
                                onClick={handleSaveAssignments}
                                disabled={isSaving}
                                className="inline-flex justify-center py-2 px-6 border border-transparent shadow-sm text-sm font-bold rounded-md text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:bg-gray-400"
                            >
                                {isSaving ? 'Saving...' : 'Save All Assignments'}
                            </button>
                        </div>
                    </div>
                ) : (
                    <div className="p-6">
                        <div className="mb-6">
                            <h4 className="text-lg font-medium text-gray-900 mb-4 flex items-center">
                                <Award className="w-5 h-5 mr-2 text-yellow-500" />
                                Bonus History & Awards
                            </h4>

                            {/* Add Bonus Form */}
                            <form onSubmit={handleAddBonus} className="bg-gray-50 p-4 rounded-lg border border-gray-200 mb-6">
                                <h5 className="text-sm font-medium text-gray-700 mb-3">Award New Bonus</h5>
                                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                    <div>
                                        <label className="block text-xs text-gray-500 mb-1">Amount</label>
                                        <div className="relative">
                                            <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                                                <DollarSign className="h-4 w-4 text-gray-400" />
                                            </div>
                                            <input
                                                type="number"
                                                required
                                                min="0"
                                                value={newBonus.amount}
                                                onChange={(e) => setNewBonus({ ...newBonus, amount: e.target.value })}
                                                className="pl-8 focus:ring-blue-500 focus:border-blue-500 block w-full sm:text-sm border-gray-300 rounded-md"
                                                placeholder="0.00"
                                            />
                                        </div>
                                    </div>
                                    <div className="md:col-span-2 flex space-x-2">
                                        <div className="flex-grow">
                                            <label className="block text-xs text-gray-500 mb-1">Reason</label>
                                            <input
                                                type="text"
                                                required
                                                value={newBonus.reason}
                                                onChange={(e) => setNewBonus({ ...newBonus, reason: e.target.value })}
                                                className="focus:ring-blue-500 focus:border-blue-500 block w-full sm:text-sm border-gray-300 rounded-md"
                                                placeholder="e.g. Year End Bonus"
                                            />
                                        </div>
                                        <div className="flex items-end">
                                            <button
                                                type="submit"
                                                className="inline-flex justify-center py-2 px-4 border border-transparent shadow-sm text-sm font-medium rounded-md text-white bg-green-600 hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500"
                                            >
                                                Award
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </form>

                            {/* Bonus List */}
                            <div className="overflow-hidden shadow ring-1 ring-black ring-opacity-5 md:rounded-lg">
                                <table className="min-w-full divide-y divide-gray-300">
                                    <thead className="bg-gray-50">
                                        <tr>
                                            <th scope="col" className="py-3.5 pl-4 pr-3 text-left text-sm font-semibold text-gray-900 sm:pl-6">Date</th>
                                            <th scope="col" className="px-3 py-3.5 text-left text-sm font-semibold text-gray-900">Reason</th>
                                            <th scope="col" className="px-3 py-3.5 text-left text-sm font-semibold text-gray-900">Amount</th>
                                            <th scope="col" className="relative py-3.5 pl-3 pr-4 sm:pr-6">
                                                <span className="sr-only">Delete</span>
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-200 bg-white">
                                        {bonuses.length === 0 ? (
                                            <tr>
                                                <td colSpan="4" className="py-4 text-center text-sm text-gray-500">No bonuses recorded.</td>
                                            </tr>
                                        ) : (
                                            bonuses.map((bonus) => (
                                                <tr key={bonus.id}>
                                                    <td className="whitespace-nowrap py-4 pl-4 pr-3 text-sm font-medium text-gray-900 sm:pl-6">
                                                        {new Date(bonus.dateGiven).toLocaleDateString()}
                                                    </td>
                                                    <td className="whitespace-nowrap px-3 py-4 text-sm text-gray-500">{bonus.reason}</td>
                                                    <td className="whitespace-nowrap px-3 py-4 text-sm text-green-600 font-semibold">
                                                        Rs. {parseFloat(bonus.amount).toLocaleString()}
                                                    </td>
                                                    <td className="relative whitespace-nowrap py-4 pl-3 pr-4 text-right text-sm font-medium sm:pr-6">
                                                        <button
                                                            onClick={() => handleDeleteBonus(bonus.id)}
                                                            className="text-red-600 hover:text-red-900"
                                                        >
                                                            <Trash2 className="w-4 h-4" />
                                                        </button>
                                                    </td>
                                                </tr>
                                            ))
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};

export default StaffProfile;
