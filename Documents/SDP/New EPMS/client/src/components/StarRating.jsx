import React from 'react';
import { Star } from 'lucide-react';

const StarRating = ({ rating, onRatingChange, readOnly = false }) => {
    return (
        <div className="flex space-x-1">
            {[1, 2, 3, 4, 5].map((star) => (
                <button
                    key={star}
                    type="button"
                    onClick={() => !readOnly && onRatingChange(star)}
                    className={`${readOnly ? 'cursor-default' : 'cursor-pointer'} focus:outline-none`}
                    disabled={readOnly}
                >
                    <Star
                        className={`w-6 h-6 ${star <= rating ? 'fill-yellow-400 text-yellow-400' : 'text-gray-300'
                            }`}
                    />
                </button>
            ))}
        </div>
    );
};

export default StarRating;
