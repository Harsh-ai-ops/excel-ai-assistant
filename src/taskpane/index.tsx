import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { initializeIcons } from '@fluentui/react';
import App from './App';
import './styles.css';

// Initialize Fluent UI icons
initializeIcons();

// Wait for Office.js to be ready before rendering React
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        const container = document.getElementById('root');
        if (container) {
            const root = createRoot(container);
            root.render(<App />);
        }
    } else {
        // Show fallback for non-Excel environments (for testing)
        const container = document.getElementById('root');
        if (container) {
            container.innerHTML = `
        <div style="padding: 20px; font-family: 'Segoe UI', sans-serif;">
          <h2>Excel AI Assistant</h2>
          <p>This add-in requires Microsoft Excel to function.</p>
          <p>Please open this add-in from within Excel.</p>
        </div>
      `;
        }
    }
});
