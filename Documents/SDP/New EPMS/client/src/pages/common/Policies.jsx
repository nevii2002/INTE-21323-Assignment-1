import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useAuth } from '../../context/AuthContext';
import {
    FileText, Upload, Plus, X, Download, Eye, Trash2, ShieldAlert
} from 'lucide-react';

const Policies = () => {
    const { user } = useAuth();
    const isAdminOrManager = user?.role === 'Admin' || user?.role === 'Manager';

    const [policies, setPolicies] = useState([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState('');

    // Upload Modal State
    const [showUploadModal, setShowUploadModal] = useState(false);
    const [uploading, setUploading] = useState(false);
    const [uploadTitle, setUploadTitle] = useState('');
    const [uploadFile, setUploadFile] = useState(null);

    // PDF Viewer State
    const [viewingPolicy, setViewingPolicy] = useState(null);

    useEffect(() => {
        fetchPolicies();
    }, []);

    const fetchPolicies = async () => {
        try {
            setLoading(true);
            const res = await axios.get('http://localhost:5000/api/policies', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setPolicies(res.data);
            setError('');
        } catch (err) {
            console.error('Error fetching policies:', err);
            setError('Failed to load HR policies. Please try again.');
        } finally {
            setLoading(false);
        }
    };

    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            const file = e.target.files[0];
            if (file.type !== 'application/pdf') {
                alert('Only PDF files are allowed!');
                return;
            }
            if (file.size > 10 * 1024 * 1024) {
                alert('File size must be less than 10MB');
                return;
            }
            setUploadFile(file);
            if (!uploadTitle) {
                // Auto-fill title from filename if empty
                setUploadTitle(file.name.replace('.pdf', '').replace(/[-_]/g, ' '));
            }
        }
    };

    const handleUploadSubmit = async (e) => {
        e.preventDefault();
        if (!uploadTitle || !uploadFile) {
            alert('Please provide both a title and a PDF file.');
            return;
        }

        const formData = new FormData();
        formData.append('title', uploadTitle);
        formData.append('policyDoc', uploadFile);

        try {
            setUploading(true);
            await axios.post('http://localhost:5000/api/policies/upload', formData, {
                headers: {
                    Authorization: `Bearer ${localStorage.getItem('token')}`,
                    'Content-Type': 'multipart/form-data'
                }
            });

            setShowUploadModal(false);
            setUploadTitle('');
            setUploadFile(null);
            fetchPolicies();
        } catch (err) {
            console.error('Error uploading policy:', err);
            alert(err.response?.data?.message || 'Failed to upload policy document');
        } finally {
            setUploading(false);
        }
    };

    const handleDelete = async (id) => {
        if (!window.confirm('Are you sure you want to delete this policy document?')) return;

        try {
            await axios.delete(`http://localhost:5000/api/policies/${id}`, {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            fetchPolicies();
        } catch (err) {
            console.error('Error deleting policy:', err);
            alert('Failed to delete policy');
        }
    };

    const openPdfViewer = (policy) => {
        // Construct the full URL for the PDF
        const normalizedPath = policy.filePath.replace(/\\/g, '/');
        const fileUrl = `http://localhost:5000/${normalizedPath.startsWith('/') ? normalizedPath.slice(1) : normalizedPath}`;

        // Open the PDF in a new browser tab securely
        window.open(fileUrl, '_blank', 'noopener,noreferrer');
    };

    // Convert bytes to MB/KB
    const formatSize = (bytes) => {
        if (!bytes) return 'Unknown size';
        const k = 1024;
        if (bytes < k) return bytes + ' Bytes';
        else if (bytes < k * k) return (bytes / k).toFixed(2) + ' KB';
        else return (bytes / (k * k)).toFixed(2) + ' MB';
    }

    return (
        <div className="space-y-6">
            {/* Header */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                    <h1 className="text-2xl font-bold tracking-tight text-gray-900">HR Policies & Documents</h1>
                    <p className="text-gray-500 mt-1">View and download company policies and handbooks.</p>
                </div>

                {isAdminOrManager && (
                    <button
                        onClick={() => setShowUploadModal(true)}
                        className="inline-flex items-center justify-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:ring-offset-2 transition-colors sm:w-auto w-full font-medium"
                    >
                        <Upload className="w-4 h-4" />
                        Upload Policy
                    </button>
                )}
            </div>

            {/* Content */}
            {loading ? (
                <div className="flex justify-center items-center h-64">
                    <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div>
                </div>
            ) : error ? (
                <div className="bg-red-50 border border-red-200 rounded-lg p-4 flex items-center justify-center text-red-700 gap-2">
                    <ShieldAlert className="w-5 h-5" />
                    <span>{error}</span>
                </div>
            ) : policies.length === 0 ? (
                <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-12 text-center">
                    <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-gray-50 mb-4">
                        <FileText className="w-8 h-8 text-gray-400" />
                    </div>
                    <h3 className="text-lg font-medium text-gray-900 mb-1">No policies found</h3>
                    <p className="text-gray-500">
                        {isAdminOrManager
                            ? "Upload a new HR policy document to get started."
                            : "There are currently no hr policies available to view."}
                    </p>
                </div>
            ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    {policies.map((policy) => (
                        <div
                            key={policy.id}
                            className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden hover:shadow-md transition-shadow group flex flex-col h-full"
                        >
                            <div className="p-6 flex-grow flex flex-col">
                                <div className="flex items-start justify-between mb-4">
                                    <div className="p-3 bg-indigo-50 text-indigo-600 rounded-lg">
                                        <FileText className="w-6 h-6" />
                                    </div>
                                    {isAdminOrManager && (
                                        <button
                                            onClick={() => handleDelete(policy.id)}
                                            className="text-gray-400 hover:text-red-500 transition-colors p-1 opacity-0 group-hover:opacity-100"
                                            title="Delete Policy"
                                        >
                                            <Trash2 className="w-4 h-4" />
                                        </button>
                                    )}
                                </div>
                                <h3 className="text-lg font-semibold text-gray-900 mb-2 line-clamp-2" title={policy.title}>
                                    {policy.title}
                                </h3>
                                <div className="mt-auto space-y-2 text-sm text-gray-500">
                                    <p className="flex justify-between">
                                        <span className="truncate pr-2">File:</span>
                                        <span className="truncate font-medium">{policy.fileName}</span>
                                    </p>
                                    <p className="flex justify-between">
                                        <span>Uploaded:</span>
                                        <span>{new Date(policy.createdAt).toLocaleDateString()}</span>
                                    </p>
                                </div>
                            </div>

                            <div className="bg-gray-50 px-6 py-4 border-t border-gray-100 flex gap-3">
                                <button
                                    onClick={() => openPdfViewer(policy)}
                                    className="flex-1 inline-flex items-center justify-center gap-2 px-3 py-2 text-sm font-medium text-indigo-700 bg-indigo-50 border border-transparent rounded-lg hover:bg-indigo-100 transition-colors"
                                >
                                    <Eye className="w-4 h-4" />
                                    View
                                </button>
                                <a
                                    href={`http://localhost:5000/${policy.filePath.replace(/\\/g, '/').startsWith('/') ? policy.filePath.replace(/\\/g, '/').slice(1) : policy.filePath.replace(/\\/g, '')}`}
                                    download={policy.fileName}
                                    target="_blank"
                                    rel="noreferrer"
                                    className="inline-flex items-center justify-center gap-2 px-3 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                                    title="Download PDF"
                                >
                                    <Download className="w-4 h-4" />
                                </a>
                            </div>
                        </div>
                    ))}
                </div>
            )}

            {/* Upload Modal */}
            {showUploadModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
                    <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in-95 duration-200">
                        <div className="flex justify-between items-center p-6 border-b border-gray-100">
                            <h3 className="text-lg font-semibold text-gray-900">Upload HR Policy</h3>
                            <button
                                onClick={() => setShowUploadModal(false)}
                                className="text-gray-400 hover:text-gray-500 transition-colors"
                            >
                                <X className="w-5 h-5" />
                            </button>
                        </div>

                        <form onSubmit={handleUploadSubmit} className="p-6 space-y-4">
                            <div>
                                <label htmlFor="title" className="block text-sm font-medium text-gray-700 mb-1">
                                    Document Title
                                </label>
                                <input
                                    type="text"
                                    id="title"
                                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-colors"
                                    placeholder="e.g., Employee Handbook 2024"
                                    value={uploadTitle}
                                    onChange={(e) => setUploadTitle(e.target.value)}
                                    required
                                />
                            </div>

                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">
                                    PDF File
                                </label>
                                <div className="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-gray-300 border-dashed rounded-lg hover:bg-gray-50 transition-colors">
                                    <div className="space-y-1 text-center">
                                        <Upload className="mx-auto h-12 w-12 text-gray-400" />
                                        <div className="flex text-sm text-gray-600 justify-center">
                                            <label htmlFor="file-upload" className="relative cursor-pointer bg-white rounded-md font-medium text-indigo-600 hover:text-indigo-500 focus-within:outline-none focus-within:ring-2 focus-within:ring-offset-2 focus-within:ring-indigo-500 p-1">
                                                <span>Upload a file</span>
                                                <input
                                                    id="file-upload"
                                                    name="file-upload"
                                                    type="file"
                                                    className="sr-only"
                                                    accept="application/pdf"
                                                    onChange={handleFileChange}
                                                    required
                                                />
                                            </label>
                                        </div>
                                        <p className="text-xs text-gray-500">PDF up to 10MB</p>
                                    </div>
                                </div>
                                {uploadFile && (
                                    <div className="mt-2 text-sm text-gray-600 flex items-center justify-between bg-indigo-50 px-3 py-2 rounded-lg border border-indigo-100">
                                        <span className="truncate mr-2 text-indigo-700 font-medium">{uploadFile.name}</span>
                                        <button
                                            type="button"
                                            onClick={() => { setUploadFile(null); }}
                                            className="text-indigo-400 hover:text-indigo-600"
                                        >
                                            <X className="w-4 h-4" />
                                        </button>
                                    </div>
                                )}
                            </div>

                            <div className="pt-4 flex justify-end gap-3">
                                <button
                                    type="button"
                                    onClick={() => setShowUploadModal(false)}
                                    className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-colors"
                                >
                                    Cancel
                                </button>
                                <button
                                    type="submit"
                                    disabled={uploading || !uploadFile || !uploadTitle}
                                    className="inline-flex items-center justify-center gap-2 px-4 py-2 text-sm font-medium text-white bg-indigo-600 border border-transparent rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                                >
                                    {uploading ? (
                                        <>
                                            <div className="animate-spin rounded-full h-4 w-4 border-2 border-white/20 border-t-white"></div>
                                            Uploading...
                                        </>
                                    ) : (
                                        <>
                                            <Upload className="w-4 h-4" />
                                            Upload Policy
                                        </>
                                    )}
                                </button>
                            </div>
                        </form>
                    </div>
                </div>
            )}
        </div>
    );
};

export default Policies;
