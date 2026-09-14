import React from 'react';
import BrowseField from '../common/BrowseField';

function InvoicePoSelector({
  tbmSummaryPath,
  onSummaryPathChange,
  onBrowseSummary,
  browseSummaryLoading,
  poNumber,
  onPoNumberChange,
  detectedPOs,
  company,
  disabled
}) {
  return (
    <>
      {/* All-TBMs Summary File */}
      <BrowseField
        label="All-TBMs Summary Excel File (e.g. All-TBMs-Summary.xlsx)"
        value={tbmSummaryPath}
        onChange={(e) => onSummaryPathChange(e.target.value)}
        onBrowse={onBrowseSummary}
        browseLoading={browseSummaryLoading}
        disabled={disabled}
        placeholder="Select All-TBMs Summary workbook containing employee bills..."
      />

      {/* PO Number Mandatory Field */}
      <div className="form-field">
        <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
          PO Number (Mandatory) <span style={{ color: '#e74c3c' }}>*</span>
        </label>
        <input
          type="text"
          list="po-suggestions"
          className="input-text"
          style={{ width: '100%', fontSize: '0.95rem', padding: '0.65rem 0.85rem', fontWeight: '600', color: '#27ae60' }}
          value={poNumber}
          onChange={(e) => onPoNumberChange(e.target.value)}
          placeholder={company === 'Corteva' ? "e.g. 4800108506" : "e.g. 500BB2026018404"}
          disabled={disabled}
          required
        />
        {detectedPOs.length > 0 && (
          <datalist id="po-suggestions">
            {detectedPOs.map((p, idx) => (
              <option key={idx} value={p} />
            ))}
          </datalist>
        )}
        {detectedPOs.length > 0 && (
          <div style={{ marginTop: '0.4rem', display: 'flex', flexWrap: 'wrap', gap: '0.4rem', alignItems: 'center' }}>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>POs in workbook:</span>
            {detectedPOs.map((p, idx) => (
              <button
                key={idx}
                type="button"
                onClick={() => onPoNumberChange(p)}
                style={{
                  padding: '0.18rem 0.55rem',
                  fontSize: '0.75rem',
                  borderRadius: '12px',
                  background: poNumber === p ? 'var(--primary-color, #27ae60)' : 'rgba(255,255,255,0.08)',
                  color: poNumber === p ? '#fff' : 'inherit',
                  border: '1px solid rgba(255,255,255,0.15)',
                  cursor: 'pointer'
                }}
              >
                {p}
              </button>
            ))}
          </div>
        )}
      </div>
    </>
  );
}

export default InvoicePoSelector;
