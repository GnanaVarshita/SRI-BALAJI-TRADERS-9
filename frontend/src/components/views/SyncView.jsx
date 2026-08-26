import React from 'react';
import ConsoleLogs from '../ConsoleLogs';

function SyncView({ isSyncing, logs, consoleRef, onStartSync }) {
  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Gmail PO Sync</h2>
        <p className="subtitle">Connect to Gmail and download purchase orders from target labels.</p>
      </div>

      <div className="card">
        <h2>Operations</h2>
        <p style={{ color: 'var(--text-muted)', marginBottom: '1rem' }}>
          This will scan your target Gmail labels (e.g. <code>CORTEVA/...</code> and <code>NEW GEN/...</code>) for purchase orders sent by <strong>ordersender-prod@ansmtp.ariba.com</strong> and organize their attachments locally.
        </p>
        <button 
          className="primary" 
          onClick={onStartSync} 
          disabled={isSyncing}
          style={{ padding: '1rem', fontSize: '1.1rem' }}
        >
          {isSyncing ? '🔄 Syncing Gmail Attachments...' : '⚡ Sync Gmail Attachments Now'}
        </button>
      </div>

      <ConsoleLogs logs={logs} consoleRef={consoleRef} />
    </div>
  );
}

export default SyncView;
