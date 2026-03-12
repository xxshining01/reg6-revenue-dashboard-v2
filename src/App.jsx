import React from 'react';
import './App.css';
import Dashboard from './components/Dashboard';
import ErrorBoundary from './ErrorBoundary';

function App() {
  return (
    <div className="app-container">
      <ErrorBoundary>
        <Dashboard />
      </ErrorBoundary>
    </div>
  );
}

export default App;