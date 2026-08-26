import React from 'react';
import CredentialsConfig from '../CredentialsConfig';

function SettingsView({ email, setEmail, password, setPassword, hasPassword, isSyncing, onSave, onResetHistory }) {
  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Configuration & Settings</h2>
        <p className="subtitle">Configure Gmail authentication and clear sync databases.</p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '1.5rem' }}>
        <CredentialsConfig 
          email={email} 
          setEmail={setEmail} 
          password={password} 
          setPassword={setPassword} 
          hasPassword={hasPassword} 
          isSyncing={isSyncing} 
          onSave={onSave} 
        />

        <div className="card">
          <h2>System Maintenance</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '1rem' }}>
            Clear the processed email database if you want to force the sync to re-examine all email attachments.
          </p>
          <button 
            className="danger" 
            onClick={onResetHistory} 
            disabled={isSyncing}
            style={{ width: 'auto', padding: '0.75rem 1.5rem' }}
          >
            ⚠️ Reset Sync History Database
          </button>
        </div>
      </div>
    </div>
  );
}

export default SettingsView;
