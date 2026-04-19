import React, { useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import axios from 'axios';
import { Eye, EyeOff } from 'lucide-react';

const Register = () => {
    const navigate = useNavigate();
    const [formData, setFormData] = useState({
        username: '',
        email: '',
        mobileNumber: '',
        position: '',
        password: ''
    });
    const [error, setError] = useState('');
    const [loading, setLoading] = useState(false);
    const [showPassword, setShowPassword] = useState(false);

    const handleChange = (e) => {
        setFormData({ ...formData, [e.target.name]: e.target.value });
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');
        setLoading(true);

        try {
            await axios.post('http://localhost:5000/api/auth/register', formData);
            alert('Registration successful! Please login.');
            navigate('/login');
        } catch (err) {
            setError(err.response?.data?.message || 'Registration failed');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="min-h-screen bg-[#0099ff] flex items-center justify-center font-sans">
            <div className="w-full max-w-4xl grid md:grid-cols-2 gap-8 items-center p-8">

                {/* Left Side - Logo */}
                <div className="hidden md:flex flex-col items-center justify-center text-white">
                    <div className="bg-white p-8 rounded-2xl shadow-xl mb-8">
                        <img src="/logo.png" alt="Techznap Logo" className="w-[300px] object-contain" />
                    </div>
                </div>

                {/* Right Side - Form */}
                <div className="w-full max-w-md bg-white rounded-xl shadow-2xl p-8 mx-auto">
                    <h2 className="text-3xl font-bold text-center text-[#0099ff] mb-8">Register</h2>

                    {error && <div className="mb-4 bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded relative text-sm">{error}</div>}

                    <form onSubmit={handleSubmit} className="space-y-5">

                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Email Address</label>
                            <input
                                type="email"
                                name="email"
                                required
                                value={formData.email}
                                onChange={handleChange}
                                placeholder="name@techznap.com"
                                className="appearance-none block w-full px-4 py-3 border border-[#0099ff] rounded-lg shadow-sm placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#0099ff] focus:border-transparent"
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Full Name</label>
                            <input
                                type="text"
                                name="username"
                                required
                                value={formData.username}
                                onChange={handleChange}
                                placeholder="Ex: Nilan Manoj"
                                className="appearance-none block w-full px-4 py-3 border border-[#0099ff] rounded-lg shadow-sm placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#0099ff] focus:border-transparent"
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Mobile Number</label>
                            <input
                                type="text"
                                name="mobileNumber"
                                required
                                value={formData.mobileNumber}
                                onChange={handleChange}
                                placeholder="+94 XXXXXXX"
                                className="appearance-none block w-full px-4 py-3 border border-[#0099ff] rounded-lg shadow-sm placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#0099ff] focus:border-transparent"
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Position</label>
                            <input
                                type="text"
                                name="position"
                                required
                                value={formData.position}
                                onChange={handleChange}
                                placeholder="Ex: Graphic Designer"
                                className="appearance-none block w-full px-4 py-3 border border-[#0099ff] rounded-lg shadow-sm placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#0099ff] focus:border-transparent"
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">Password</label>
                            <div className="relative">
                                <input
                                    type={showPassword ? "text" : "password"}
                                    name="password"
                                    required
                                    value={formData.password}
                                    onChange={handleChange}
                                    placeholder="••••••••••"
                                    className="appearance-none block w-full px-4 py-3 border border-[#0099ff] rounded-lg shadow-sm placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-[#0099ff] focus:border-transparent pr-12"
                                />
                                <button
                                    type="button"
                                    onClick={() => setShowPassword(!showPassword)}
                                    className="absolute inset-y-0 right-0 pr-4 flex items-center text-gray-400 hover:text-gray-600"
                                >
                                    {showPassword ? <EyeOff className="h-5 w-5" /> : <Eye className="h-5 w-5" />}
                                </button>
                            </div>
                        </div>

                        <div className="pt-2">
                            <button
                                type="submit"
                                disabled={loading}
                                className="w-full flex justify-center py-3 px-4 border border-transparent rounded-lg shadow-sm text-sm font-medium text-white bg-[#0099ff] hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 disabled:bg-gray-400 transition-colors"
                            >
                                {loading ? 'Registering...' : 'Register'}
                            </button>
                        </div>

                        <div className="text-center mt-4">
                            <Link to="/login" className="font-medium text-[#0099ff] hover:text-blue-500 underline">
                                Already have an account? Login
                            </Link>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    );
};
export default Register;
