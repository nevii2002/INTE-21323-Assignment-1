import React from 'react';

const PlaceholderPage = ({ title }) => {
    return (
        <div className="bg-white shadow rounded-lg p-6">
            <h3 className="text-xl font-medium text-gray-900 mb-4">{title}</h3>
            <div className="border-4 border-dashed border-gray-200 rounded-lg h-64 flex items-center justify-center">
                <span className="text-gray-400">Content for {title} will be implemented here.</span>
            </div>
        </div>
    );
};

export default PlaceholderPage;
