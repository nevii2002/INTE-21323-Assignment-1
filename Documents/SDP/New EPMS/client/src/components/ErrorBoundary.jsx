import React from 'react';

class ErrorBoundary extends React.Component {
    constructor(props) {
        super(props);
        this.state = { hasError: false, error: null, errorInfo: null };
    }

    static getDerivedStateFromError(error) {
        return { hasError: true };
    }

    componentDidCatch(error, errorInfo) {
        this.setState({
            error: error,
            errorInfo: errorInfo
        });
        console.error("Uncaught error:", error, errorInfo);
    }

    render() {
        if (this.state.hasError) {
            return (
                <div style={{ padding: '20px', backgroundColor: '#fee2e2', border: '1px solid #ef4444', borderRadius: '8px', margin: '20px', fontFamily: 'monospace' }}>
                    <h1 style={{ color: '#b91c1c', fontSize: '1.5rem', marginBottom: '10px' }}>Something went wrong.</h1>
                    <details style={{ whiteSpace: 'pre-wrap' }}>
                        <summary>Click for error details</summary>
                        <p style={{ color: '#b91c1c', fontWeight: 'bold' }}>{this.state.error && this.state.error.toString()}</p>
                        <br />
                        <p style={{ color: '#4b5563', fontSize: '0.8rem' }}>{this.state.errorInfo && this.state.errorInfo.componentStack}</p>
                    </details>
                </div>
            );
        }

        return this.props.children;
    }
}

export default ErrorBoundary;
