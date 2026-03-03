import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles/main.css';

const container = document.getElementById('root');
if (!container) throw new Error('Failed to find the root element');

const root = createRoot(container);

const renderApp = () => {
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};

// Initial render
renderApp();

// Handle HMR updates
if (import.meta.hot) {
  import.meta.hot.accept(renderApp);
}

// Global error handling
window.onerror = function(message, source, lineno, colno, error) {
  console.error('Global error:', { message, source, lineno, colno, error });
  return true; // Prevent default error handling
};

// Handle unhandled promise rejections
window.onunhandledrejection = function(event) {
  console.error('Unhandled rejection:', event.reason);
  event.preventDefault();
};