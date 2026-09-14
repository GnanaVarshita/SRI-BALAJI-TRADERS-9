import React from 'react';

function InvoiceCompanySelector({ company, onCompanyChange, area, onAreaChange, disabled }) {
  const cortevaAreas = ['Nellore', 'Kurnool', 'Suryapet'];
  const fmcAreas = ['Nellore', 'Nandyala', 'Kurnool'];
  const currentAreas = company === 'Corteva' ? cortevaAreas : fmcAreas;

  return (
    <div>
      {/* Company Selector Header */}
      <div style={{ marginBottom: '1.25rem' }}>
        <label className="form-label" style={{ fontSize: '0.92rem', fontWeight: 'bold', marginBottom: '0.6rem', display: 'block' }}>
          Select Client / Organization
        </label>
        <div style={{ display: 'flex', gap: '1rem' }}>
          <button
            type="button"
            className={company === 'Corteva' ? 'primary' : 'secondary'}
            onClick={() => onCompanyChange('Corteva')}
            style={{ flex: 1, padding: '0.75rem', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem', fontWeight: 'bold' }}
            disabled={disabled}
          >
            🌾 Corteva Agriscience
          </button>
          <button
            type="button"
            className={company === 'FMC' ? 'primary' : 'secondary'}
            onClick={() => onCompanyChange('FMC')}
            style={{ flex: 1, padding: '0.75rem', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem', fontWeight: 'bold' }}
            disabled={disabled}
          >
            🌱 FMC (New Gen)
          </button>
        </div>
      </div>

      {/* Area Dropdown Field */}
      <div className="form-field">
        <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
          Operational Area ({company}) <span style={{ color: '#e74c3c' }}>*</span>
        </label>
        <select
          className="input-text"
          style={{ width: '100%', fontSize: '0.95rem', padding: '0.65rem 0.85rem', fontWeight: '600' }}
          value={area}
          onChange={(e) => onAreaChange(e.target.value)}
          disabled={disabled}
        >
          {currentAreas.map((a) => (
            <option key={a} value={a}>
              📍 {a}
            </option>
          ))}
        </select>
        <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: '0.25rem', display: 'block' }}>
          <em>Selected area will be printed directly in the AREA field on the Tax Invoice (Row {company === 'Corteva' ? '16' : '17'}).</em>
        </span>
      </div>
    </div>
  );
}

export default InvoiceCompanySelector;
