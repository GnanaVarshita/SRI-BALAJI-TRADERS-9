import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';

function InvoiceDetailsForm({
  invoiceNumber,
  onInvoiceNumberChange,
  saveFolderPath,
  onSaveFolderPathChange,
  onBrowseFolder,
  browseFolderLoading,
  serviceChargePercent,
  onServiceChargeChange,
  invoiceDate,
  onInvoiceDateChange,
  poValue,
  onPoValueChange,
  requesterName,
  onRequesterNameChange,
  company,
  disabled
}) {
  const [showAdvanced, setShowAdvanced] = useState(false);

  return (
    <>
      {/* Invoice Number Field */}
      <div className="form-field">
        <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
          Invoice Number to be Raised / Updated <span style={{ color: '#e74c3c' }}>*</span>
        </label>
        <input
          type="text"
          className="input-text"
          style={{ width: '100%', fontSize: '0.95rem', padding: '0.65rem 0.85rem', fontWeight: 'bold' }}
          value={invoiceNumber}
          onChange={(e) => onInvoiceNumberChange(e.target.value)}
          placeholder={company === 'Corteva' ? "e.g. SBT26270069 or 69" : "e.g. SBT26270073 or 73"}
          disabled={disabled}
          required
        />
        <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: '0.3rem', display: 'block' }}>
          💡 <em>If an invoice file with this number exists in the save folder, newly received bills will be appended and totals updated. Otherwise, a new invoice is created.</em>
        </span>
      </div>

      {/* Save Folder Path */}
      <BrowseField
        label="Save Folder for Invoices"
        value={saveFolderPath}
        onChange={(e) => onSaveFolderPathChange(e.target.value)}
        onBrowse={onBrowseFolder}
        browseLoading={browseFolderLoading}
        disabled={disabled}
        placeholder="Select destination folder for generated Tax Invoice workbooks..."
      />

      {/* Service Charge Field */}
      <div className="form-field">
        <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
          Agency Service Charge Percentage (%)
        </label>
        <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
          <input
            type="number"
            step="0.1"
            className="input-text"
            style={{ width: '120px', padding: '0.65rem 0.85rem', fontWeight: 'bold' }}
            value={serviceChargePercent}
            onChange={(e) => onServiceChargeChange(e.target.value)}
            disabled={disabled}
          />
          <span style={{ color: 'var(--text-muted)', fontSize: '0.88rem' }}>
            % (Standard: {company === 'Corteva' ? '5.0%' : '4.5%'})
          </span>
        </div>
      </div>

      {/* Advanced Options Toggle */}
      <div style={{ marginTop: '0.5rem' }}>
        <button
          type="button"
          onClick={() => setShowAdvanced(!showAdvanced)}
          style={{
            background: 'none',
            border: 'none',
            color: 'var(--text-accent, #e67e22)',
            padding: 0,
            cursor: 'pointer',
            fontSize: '0.85rem',
            display: 'flex',
            alignItems: 'center',
            gap: '0.4rem',
            fontWeight: '600'
          }}
        >
          <span>{showAdvanced ? '▼ Hide Advanced Invoice Options' : '▶ Show Advanced Invoice Options'}</span>
        </button>

        {showAdvanced && (
          <div style={{
            marginTop: '0.85rem',
            padding: '1rem',
            background: 'rgba(255,255,255,0.02)',
            border: '1px solid rgba(255,255,255,0.07)',
            borderRadius: '6px',
            display: 'flex',
            flexDirection: 'column',
            gap: '0.9rem'
          }}>
            <div className="form-field">
              <label className="form-label" style={{ fontSize: '0.82rem' }}>Custom Invoice Date (DD-MM-YYYY)</label>
              <input
                type="text"
                className="input-text"
                value={invoiceDate}
                onChange={(e) => onInvoiceDateChange(e.target.value)}
                placeholder="DD-MM-YYYY"
                disabled={disabled}
              />
            </div>

            {company === 'Corteva' && (
              <div className="form-field">
                <label className="form-label" style={{ fontSize: '0.82rem' }}>PO Value (Printed in Sheet 4 Summary)</label>
                <input
                  type="number"
                  className="input-text"
                  value={poValue}
                  onChange={(e) => onPoValueChange(e.target.value)}
                  placeholder="250000"
                  disabled={disabled}
                />
              </div>
            )}

            <div className="form-field">
              <label className="form-label" style={{ fontSize: '0.82rem' }}>Requester / Receiver Name (Optional)</label>
              <input
                type="text"
                className="input-text"
                value={requesterName}
                onChange={(e) => onRequesterNameChange(e.target.value)}
                placeholder={company === 'Corteva' ? "e.g. R.Bhaskar" : "e.g. Madhavareddy"}
                disabled={disabled}
              />
            </div>
          </div>
        )}
      </div>
    </>
  );
}

export default InvoiceDetailsForm;
