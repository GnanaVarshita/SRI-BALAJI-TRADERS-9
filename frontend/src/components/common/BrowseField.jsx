import React from 'react';

function BrowseField({ label, value, onChange, onBrowse, browseLoading, disabled = false }) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <div style={{ display: 'flex', gap: '0.5rem' }}>
        <input 
          type="text" 
          value={value} 
          onChange={onChange} 
          placeholder="Click browse or paste absolute path..."
          required
          disabled={disabled}
          style={{ flex: 1 }}
        />
        <button 
          type="button" 
          onClick={onBrowse} 
          disabled={browseLoading || disabled}
          style={{ width: 'auto', whiteSpace: 'nowrap', backgroundColor: 'var(--primary-color)', color: '#fff', border: 'none' }}
        >
          {browseLoading ? '⌛ Browsing...' : '🔍 Browse PC'}
        </button>
      </div>
    </div>
  );
}

export default BrowseField;
