import React from 'react';

function OperationsPanel({ isSyncing, totalSynced, onStartSync, onOpenFolder, onResetHistory }) {
  return (
    <section className="card">
      <h2>Operations Panel</h2>
      <div className="button-grid">
        <button 
          className="primary" 
          onClick={onStartSync} 
          disabled={isSyncing}
        >
          {isSyncing ? 'Syncing Gmail Attachments...' : 'Sync Gmail Attachments'}
        </button>
        <button onClick={() => onOpenFolder('corteva')}>
          Open Corteva POs Local Folder
        </button>
        <button onClick={() => onOpenFolder('newgen')}>
          Open New Gen POs Local Folder
        </button>
        <button 
          className="danger" 
          onClick={onResetHistory} 
          disabled={isSyncing}
        >
          Reset Sync History Database
        </button>
      </div>
      
      <div style={{ marginTop: '0.5rem', display: 'flex', justifyContent: 'space-between', fontSize: '0.9rem', color: 'var(--text-muted)' }}>
        <span>Total Sync Sessions Saved:</span>
        <strong style={{ color: 'var(--primary-color)' }}>{totalSynced} emails</strong>
      </div>
    </section>
  );
}

export default OperationsPanel;
