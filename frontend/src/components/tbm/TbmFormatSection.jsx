import React from 'react';

function TbmFormatSection({
  tbmFolderPath,
  isLoading,
  loadingStep1,
  forceReformat,
  setForceReformat,
  onFormatSheets
}) {
  return (
    <div
      style={{
        border: '1px solid var(--border-color)',
        background: 'var(--bg-hover)',
        padding: '1.15rem',
        borderRadius: '8px',
        display: 'flex',
        flexDirection: 'column',
        gap: '0.65rem',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--primary-color)' }}>
          Step 1: Format TBM Sheets (Second Sheet)
        </h3>
        <span
          style={{
            fontSize: '0.78rem',
            background: 'var(--primary-color)',
            color: '#000',
            padding: '0.15rem 0.5rem',
            borderRadius: '4px',
            fontWeight: 'bold',
          }}
        >
          Step 1
        </span>
      </div>
      <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
        Reads raw summaries from Sheet 1, groups by (PO, Product, Activity), formats dates into{' '}
        <code>DD-MM-YYYY</code>, adds Excel formulas, writes styled tables to{' '}
        <strong>Sheet 2</strong>, and marks Column R as <strong>DONE</strong>. Already formatted files are automatically skipped.
      </p>

      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginTop: '0.1rem' }}>
        <input
          type="checkbox"
          id="forceReformat"
          checked={forceReformat}
          onChange={(e) => setForceReformat(e.target.checked)}
          disabled={isLoading}
          style={{ cursor: 'pointer' }}
        />
        <label htmlFor="forceReformat" style={{ fontSize: '0.8rem', color: 'var(--text-muted)', cursor: 'pointer' }}>
          Force re-format (overwrite files already marked as DONE)
        </label>
      </div>

      <button
        type="button"
        className="secondary"
        onClick={onFormatSheets}
        disabled={isLoading || !tbmFolderPath}
        style={{ marginTop: '0.35rem', padding: '0.75rem', fontWeight: '600' }}
      >
        {loadingStep1
          ? '⏳ Formatting TBM Summaries...'
          : '✨ Step 1: Format All TBM Sheets in-place (Sheet 2)'}
      </button>
    </div>
  );
}

export default TbmFormatSection;
