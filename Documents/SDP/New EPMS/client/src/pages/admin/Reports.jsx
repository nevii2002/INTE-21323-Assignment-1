import React, { useState, useEffect } from 'react';
import axios from 'axios';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { FileText, Download, Loader, Filter } from 'lucide-react';
import { useNavigate } from 'react-router-dom';

const Reports = () => {
    const [loading, setLoading] = useState(false);
    const [employees, setEmployees] = useState([]);

    // Filters
    const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
    const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1); // 1-12
    const [selectedEmployee, setSelectedEmployee] = useState('all');

    useEffect(() => {
        fetchEmployees();
    }, []);

    const fetchEmployees = async () => {
        try {
            const res = await axios.get('http://localhost:5000/api/staff', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });
            setEmployees(res.data);
        } catch (error) {
            console.error("Error fetching staff:", error);
        }
    };

    const getMonthName = (monthNum) => {
        const date = new Date();
        date.setMonth(monthNum - 1);
        return date.toLocaleString('default', { month: 'long' });
    };

    const generateBonusReport = async () => {
        setLoading(true);
        try {
            const response = await axios.get('http://localhost:5000/api/bonuses', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });

            // Filter Logic
            const bonuses = response.data.filter(b => {
                if (!b.employee) return false;
                const bonusDate = new Date(b.dateGiven);
                const matchYear = bonusDate.getFullYear() === parseInt(selectedYear);

                // Allow "All Months" if maybe we add that option later, but for now specific month
                const matchMonth = selectedMonth === 'all' ? true : (bonusDate.getMonth() + 1) === parseInt(selectedMonth);

                const matchEmployee = selectedEmployee === 'all' ? true : b.employeeId === parseInt(selectedEmployee);

                return matchYear && matchMonth && matchEmployee;
            });

            if (bonuses.length === 0) {
                alert("No bonus records found for the selected criteria.");
                setLoading(false);
                return;
            }

            const doc = new jsPDF();
            doc.setFontSize(18);
            doc.text('Employee Bonus Report', 14, 22);
            doc.setFontSize(11);
            doc.text(`Period: ${selectedMonth === 'all' ? 'All Months' : getMonthName(selectedMonth)} ${selectedYear}`, 14, 30);
            doc.text(`Generated on: ${new Date().toLocaleDateString()}`, 14, 36);

            autoTable(doc, {
                startY: 45,
                head: [['Date', 'Employee', 'Reason', 'Amount (LKR)']],
                body: bonuses.map(b => [
                    new Date(b.dateGiven).toLocaleDateString(),
                    b.employee.username,
                    b.reason,
                    parseFloat(b.amount).toLocaleString()
                ]),
                theme: 'grid',
                headStyles: { fillColor: [66, 139, 202] },
            });

            doc.save(`bonus_report_${selectedYear}_${selectedMonth}.pdf`);
        } catch (error) {
            console.error('Report Generation Error', error);
            alert('Failed to generate Bonus Report.');
        } finally {
            setLoading(false);
        }
    };

    const generatePerformanceReport = async () => {
        setLoading(true);
        try {
            // Fetch Evaluations instead of just Staff
            const response = await axios.get('http://localhost:5000/api/evaluations', {
                headers: { Authorization: `Bearer ${localStorage.getItem('token')}` }
            });

            // Filter Logic
            const evaluations = response.data.filter(e => {
                if (!e.employee) return false;

                // Period format "YYYY-MM"
                const [eYear, eMonth] = e.period.split('-');
                const matchYear = parseInt(eYear) === parseInt(selectedYear);
                const matchMonth = selectedMonth === 'all' ? true : parseInt(eMonth) === parseInt(selectedMonth);

                const matchEmployee = selectedEmployee === 'all' ? true : e.employeeId === parseInt(selectedEmployee);

                return matchYear && matchMonth && matchEmployee;
            });

            if (evaluations.length === 0) {
                alert("No evaluations found for the selected criteria.");
                setLoading(false);
                return;
            }

            const doc = new jsPDF();
            doc.setFontSize(18);
            doc.text('Performance Evaluation Report', 14, 22);
            doc.setFontSize(11);
            doc.text(`Period: ${selectedMonth === 'all' ? 'All Months' : getMonthName(selectedMonth)} ${selectedYear}`, 14, 30);

            autoTable(doc, {
                startY: 40,
                head: [['Employee', 'Role', 'Period', 'Status', 'Evaluator']],
                body: evaluations.map(e => [
                    e.employee.username,
                    e.employee.role || 'Employee', // User model might not have role here if not included, but usually is
                    e.period,
                    e.status,
                    e.evaluator ? e.evaluator.username : 'Self'
                ]),
                theme: 'striped',
                headStyles: { fillColor: [46, 204, 113] },
            });

            doc.save(`performance_report_${selectedYear}_${selectedMonth}.pdf`);
        } catch (error) {
            console.error('Report Generation Error', error);
            alert('Failed to generate Performance Report.');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="p-6 bg-gray-50 min-h-screen">
            <h2 className="text-2xl font-bold text-gray-800 mb-6 flex items-center">
                <FileText className="mr-2 text-blue-600" />
                System Reports
            </h2>

            {/* Filter Bar */}
            <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200 mb-6 flex flex-wrap gap-4 items-end">
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Year</label>
                    <select
                        value={selectedYear}
                        onChange={(e) => setSelectedYear(e.target.value)}
                        className="block w-32 pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-md border"
                    >
                        <option value="2025">2025</option>
                        <option value="2026">2026</option>
                        <option value="2027">2027</option>
                    </select>
                </div>

                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Month</label>
                    <select
                        value={selectedMonth}
                        onChange={(e) => setSelectedMonth(e.target.value)}
                        className="block w-40 pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-md border"
                    >
                        {/* <option value="all">All Months</option> */}
                        {Array.from({ length: 12 }, (_, i) => i + 1).map(m => (
                            <option key={m} value={m}>{getMonthName(m)}</option>
                        ))}
                    </select>
                </div>

                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Employee</label>
                    <select
                        value={selectedEmployee}
                        onChange={(e) => setSelectedEmployee(e.target.value)}
                        className="block w-64 pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-blue-500 focus:border-blue-500 sm:text-sm rounded-md border"
                    >
                        <option value="all">All Employees</option>
                        {employees.map(emp => (
                            <option key={emp.id} value={emp.id}>{emp.username} ({emp.email})</option>
                        ))}
                    </select>
                </div>

                <div className="pb-2 text-sm text-gray-500 flex items-center">
                    <Filter className="w-4 h-4 mr-1" />
                    Filters apply to PDF generation
                </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Bonus Report Card */}
                <div className="bg-white p-6 rounded-lg shadow-md border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-800 mb-2">Bonus History Report</h3>
                    <p className="text-sm text-gray-500 mb-6">
                        Generates a PDF of bonuses awarded based on the selected filters.
                    </p>
                    <button
                        onClick={generateBonusReport}
                        disabled={loading}
                        className="w-full flex items-center justify-center bg-blue-600 hover:bg-blue-700 text-white font-medium py-2 px-4 rounded transition-colors"
                    >
                        {loading ? <Loader className="animate-spin mr-2 h-5 w-5" /> : <Download className="mr-2 h-5 w-5" />}
                        Download Bonus PDF
                    </button>
                </div>

                {/* Performance Report Card */}
                <div className="bg-white p-6 rounded-lg shadow-md border border-gray-200">
                    <h3 className="text-lg font-semibold text-gray-800 mb-2">Performance Evaluation Report</h3>
                    <p className="text-sm text-gray-500 mb-6">
                        Generates a PDF of completed evaluations for the selected period/employee.
                    </p>
                    <button
                        onClick={generatePerformanceReport}
                        disabled={loading}
                        className="w-full flex items-center justify-center bg-green-600 hover:bg-green-700 text-white font-medium py-2 px-4 rounded transition-colors"
                    >
                        {loading ? <Loader className="animate-spin mr-2 h-5 w-5" /> : <Download className="mr-2 h-5 w-5" />}
                        Download Performance PDF
                    </button>
                </div>
            </div>
        </div>
    );
};

export default Reports;
