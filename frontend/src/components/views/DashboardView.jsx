import React from 'react';

function DashboardView({ isSyncing, totalSynced, onOpenFolder }) {
  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Dashboard Overview</h2>
        <p className="subtitle">Welcome to the PO Automation panel for Sri Balaji Traders.</p>
      </div>

      <div className="stat-cards-grid">
        <div className="stat-card">
          <div className="stat-icon">📥</div>
          <div className="stat-details">
            <h3>Total POs Downloaded</h3>
            <p className="stat-number">{totalSynced}</p>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon">{isSyncing ? '🔄' : '✅'}</div>
          <div className="stat-details">
            <h3>System Status</h3>
            <p className="stat-number status-text">{isSyncing ? 'Syncing...' : 'Idle / Standby'}</p>
          </div>
        </div>
      </div>

      <div className="card" style={{ marginTop: '2rem' }}>
        <h2>Quick Access Folders</h2>
        <p style={{ color: 'var(--text-muted)', marginBottom: '1rem' }}>
          Open your downloaded PO directories directly on your local computer's Windows Explorer:
        </p>
        <div className="button-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '1rem' }}>
          <button className="primary" onClick={() => onOpenFolder('corteva')}>
            📂 Open Corteva POs
          </button>
          <button className="primary" onClick={() => onOpenFolder('newgen')}>
            📂 Open New Gen POs
          </button>
          <button className="primary" onClick={() => onOpenFolder('fmc')}>
            📂 Open FMC POs
          </button>
        </div>
      </div>
    </div>
  );
}

export default DashboardView;
