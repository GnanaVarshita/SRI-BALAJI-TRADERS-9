import React from 'react';

function ResultPanel({ result, filePath, isSummary = false }) {
  if (!result) return null;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
      <div className="toast success" style={{ width: '100%', display: 'flex', alignItems: 'flex-start', gap: '0.75rem' }}>
        <span style={{ fontSize: '1.25rem' }}>✅</span>
        <div>
          <strong style={{ display: 'block', fontSize: '1.05rem' }}>
            {isSummary ? 'PO Summary Generated!' : 'Sheets Appended Successfully!'}
          </strong>
          <p style={{ fontSize: '0.85rem', marginTop: '0.25rem' }}>
            {isSummary 
              ? 'A new spreadsheet has been successfully created. The file contains separate product worksheets with live formulas (`SUM`, `SUMPRODUCT`, `With GST`) and tracking actuals tables.'
              : 'The Excel sheet has been updated **in-place**. Product worksheets were created with formatted tables, double-borders, and live formulas.'
            }
          </p>
        </div>
      </div>

      {/* Statistics Card */}
      {!isSummary && result.totals && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <div className="stat-card" style={{ padding: '1rem' }}>
            <div className="stat-details">
              <h3>Total Budget Alloc.</h3>
              <p className="stat-number" style={{ fontSize: '1.3rem' }}>
                ₹{result.totals.totalBudget.toLocaleString()}
              </p>
            </div>
          </div>
          <div className="stat-card" style={{ padding: '1rem' }}>
            <div className="stat-details">
              <h3>Total Activities</h3>
              <p className="stat-number" style={{ fontSize: '1.3rem', color: 'var(--primary-color)' }}>
                {result.totals.totalQty}
              </p>
            </div>
          </div>
        </div>
      )}

      <div style={{ borderTop: '1px solid var(--border-color)', paddingTop: '1.5rem' }}>
        <h4 style={{ color: 'var(--text-gold)', marginBottom: '0.5rem' }}>
          {isSummary ? 'Saved File Location:' : 'Modified File Location:'}
        </h4>
        <code style={{ 
          display: 'block', 
          padding: '0.75rem', 
          backgroundColor: '#f8f9fa', 
          borderRadius: '4px', 
          fontSize: '0.85rem',
          wordBreak: 'break-all',
          border: '1px solid var(--border-color)',
          color: 'var(--text-main)'
        }}>
          {isSummary ? result.outputPath : filePath}
        </code>
        <p style={{ marginTop: '0.75rem', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
          {isSummary 
            ? '💡 Open this sheet directly from your workspace folder to begin tracking invoice expenditures.'
            : '💡 You can now open this file directly from your local folders.'
          }
        </p>
      </div>
    </div>
  );
}

export default ResultPanel;
