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
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '2rem', background: '#fee2e2', color: '#991b1b', borderRadius: '8px', margin: '2rem', fontFamily: 'monospace', zIndex: 999999, position: 'relative' }}>
          <h2 style={{ fontSize: '1.5rem', fontWeight: 'bold', marginBottom: '1rem' }}>เกิดข้อผิดพลาดในการแสดงผล (Render Crash)</h2>
          <p style={{ marginBottom: '1rem' }}>กรุณาคัดลอกข้อความด้านล่างส่งให้ AI อัปเดตแก้ไข:</p>
          <div style={{ background: '#fef2f2', padding: '1rem', borderRadius: '4px', border: '1px solid #fca5a5', overflowX: 'auto', whiteSpace: 'pre-wrap' }}>
            <strong>{this.state.error && this.state.error.toString()}</strong>
            <br />
            {this.state.errorInfo && this.state.errorInfo.componentStack}
          </div>
        </div>
      );
    }

    return this.props.children; 
  }
}

export default ErrorBoundary;