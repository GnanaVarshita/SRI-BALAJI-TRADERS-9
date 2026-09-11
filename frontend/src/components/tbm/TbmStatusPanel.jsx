import React from 'react';
import ResultPanel from '../common/ResultPanel';

function TbmStatusPanel({
  isLoading,
  loadingStep1,
  loadingStep2,
  step1Result,
  step2Result,
  errorMsg
}) {
  return (
    <div className="card">
      <h2>Execution Status</h2>

      {errorMsg && (
        <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
          ⚠️ {errorMsg}
        </div>
      )}

      {!isLoading && !step1Result && !step2Result && !errorMsg && (
        <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
          <h3>Ready for Processing</h3>
          <p style={{ marginTop: '0.5rem' }}>
            Select the TBM s Summary folder, run <strong>Step 1</strong> to format the Excel sheets on Sheet 2, then run <strong>Step 2</strong> to build the Master Consolidated Summary.
          </p>
        </div>
      )}

      {loadingStep1 && (
        <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
          <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
          <h3>Formatting TBM Excel Sheets (Step 1)...</h3>
          <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
            Extracting raw tables, normalizing dates to DD-MM-YYYY, grouping by (PO, Product, Activity), and writing green styled tables with PO totals to Sheet 2.
          </p>
        </div>
      )}

      {loadingStep2 && (
        <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
          <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
          <h3>Consolidating Master Summary (Step 2)...</h3>
          <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
            Aggregating formatted tables across all TBMs into Master Summary with PO Grand Totals and TBM Amount Summary sheet.
          </p>
        </div>
      )}

      {/* STEP 1 RESULT */}
      {step1Result && !loadingStep1 && (
        <div style={{ marginBottom: '1.5rem' }}>
          <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
            <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Step 1: Formatting Complete</h4>
            <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
              {step1Result.message}
            </p>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Files Processed</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                {step1Result.processedFiles} / {step1Result.totalFiles}
              </div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Activities</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                {step1Result.totalActivities}
              </div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Tables Created</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                {step1Result.totalTables}
              </div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Destination</span>
              <div style={{ fontSize: '0.9rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                Sheet 2 (In-Place)
              </div>
            </div>
          </div>

          {step1Result.details && step1Result.details.length > 0 && (
            <div style={{
              maxHeight: '160px',
              overflowY: 'auto',
              background: 'var(--bg-card)',
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              padding: '0.5rem',
              fontSize: '0.78rem',
            }}>
              {step1Result.details.map((d, idx) => (
                <div key={idx} style={{ padding: '0.25rem 0', borderBottom: idx < step1Result.details.length - 1 ? '1px solid var(--border-color)' : 'none', color: d.success ? 'var(--text-color)' : '#e74c3c' }}>
                  {d.success ? '✓' : '⚠️'} <strong>{d.file}</strong>: {d.message}
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* STEP 2 RESULT */}
      {step2Result && !loadingStep2 && (
        <div>
          <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
            <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Step 2: Consolidation Complete</h4>
            <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
              {step2Result.message}
            </p>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Activities</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.totalActivities}</div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Tables</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.totalTables}</div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Sheets Created</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.sheetsCount}</div>
            </div>
            <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Max Tables / Sheet</span>
              <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>9</div>
            </div>
          </div>

          <ResultPanel result={step2Result} isSummary={true} />
        </div>
      )}
    </div>
  );
}

export default TbmStatusPanel;
