import React from 'react';

function DetailsOfBillsResult({ result }) {
  if (!result) {
    return (
      <div
        className="card info-card"
        style={{
          padding: '1.5rem',
          backgroundColor: 'var(--card-bg)',
          color: '#64748b',
          textAlign: 'center',
        }}
      >
        <span style={{ fontSize: '2.5rem', display: 'block', marginBottom: '0.5rem' }}>
          📑
        </span>
        <h4 style={{ margin: '0 0 0.5rem 0', color: 'var(--text-color)' }}>
          Smart Invoice Consolidation
        </h4>
        <p style={{ fontSize: '0.85rem', lineHeight: '1.4', margin: 0 }}>
          Select your <strong>Details of Bills Excel</strong> and the{' '}
          <strong>Invoices Folder</strong>. The system will inspect each invoice workbook,
          parse individual activity and TBM breakdowns, append only unlisted invoices into Sheet
          1, update Row 3 total formulas, and populate IV numbers into your Budget Cards.
        </p>
      </div>
    );
  }

  return (
    <div
      className="card result-card"
      style={{ padding: '1.5rem', backgroundColor: 'var(--card-bg)' }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
        <span style={{ fontSize: '1.5rem' }}>✅</span>
        <div>
          <h3 style={{ margin: 0, color: '#16a34a' }}>Sync Completed</h3>
          <p style={{ margin: 0, fontSize: '0.85rem', color: '#666' }}>{result.message}</p>
        </div>
      </div>

      <div
        className="stats-grid"
        style={{
          display: 'grid',
          gridTemplateColumns: '1fr 1fr',
          gap: '0.75rem',
          marginBottom: '1.25rem',
        }}
      >
        <div
          style={{
            padding: '0.75rem',
            backgroundColor: '#f8fafc',
            borderRadius: '6px',
            border: '1px solid #e2e8f0',
          }}
        >
          <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>
            Invoices Scanned
          </span>
          <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>
            {result.totalInvoicesFound}
          </strong>
        </div>
        <div
          style={{
            padding: '0.75rem',
            backgroundColor: '#f0fdf4',
            borderRadius: '6px',
            border: '1px solid #bbf7d0',
          }}
        >
          <span style={{ fontSize: '0.75rem', color: '#16a34a', display: 'block' }}>
            New Invoices Appended
          </span>
          <strong style={{ fontSize: '1.2rem', color: '#15803d' }}>
            {result.appendedInvoices?.length || 0}
          </strong>
        </div>
        <div
          style={{
            padding: '0.75rem',
            backgroundColor: '#f8fafc',
            borderRadius: '6px',
            border: '1px solid #e2e8f0',
          }}
        >
          <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>
            Activity Rows Added
          </span>
          <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>
            {result.totalRowsAdded || 0}
          </strong>
        </div>
        <div
          style={{
            padding: '0.75rem',
            backgroundColor: '#f8fafc',
            borderRadius: '6px',
            border: '1px solid #e2e8f0',
          }}
        >
          <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>
            Total Rows in Sheet
          </span>
          <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>
            {result.totalRowsInSheet || 0}
          </strong>
        </div>
      </div>

      {result.appendedInvoices && result.appendedInvoices.length > 0 && (
        <div style={{ marginBottom: '1rem' }}>
          <label
            style={{
              fontSize: '0.85rem',
              fontWeight: '600',
              color: '#334155',
              display: 'block',
              marginBottom: '0.4rem',
            }}
          >
            Appended Invoices:
          </label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem' }}>
            {result.appendedInvoices.map((iv, idx) => (
              <span
                key={idx}
                style={{
                  padding: '0.2rem 0.5rem',
                  backgroundColor: '#dcfce7',
                  color: '#15803d',
                  borderRadius: '4px',
                  fontSize: '0.8rem',
                  fontWeight: '600',
                }}
              >
                IV #{iv}
              </span>
            ))}
          </div>
        </div>
      )}

      {result.skippedInvoices && result.skippedInvoices.length > 0 && (
        <div style={{ marginBottom: '1rem' }}>
          <label
            style={{
              fontSize: '0.85rem',
              fontWeight: '600',
              color: '#64748b',
              display: 'block',
              marginBottom: '0.4rem',
            }}
          >
            Already Present (Skipped):
          </label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem' }}>
            {result.skippedInvoices.map((iv, idx) => (
              <span
                key={idx}
                style={{
                  padding: '0.2rem 0.5rem',
                  backgroundColor: '#f1f5f9',
                  color: '#64748b',
                  borderRadius: '4px',
                  fontSize: '0.8rem',
                }}
              >
                IV #{iv}
              </span>
            ))}
          </div>
        </div>
      )}

      {result.cardsSync && (
        <div
          style={{
            padding: '0.75rem',
            backgroundColor: '#eff6ff',
            borderRadius: '6px',
            border: '1px solid #bfdbfe',
            marginBottom: '1rem',
          }}
        >
          <span
            style={{
              fontSize: '0.8rem',
              color: '#1e40af',
              fontWeight: '600',
              display: 'block',
            }}
          >
            Budget Cards Updated: {result.cardsSync.cardsUpdated} card(s)
          </span>
          {result.cardsSync.updatedPos && result.cardsSync.updatedPos.length > 0 && (
            <span
              style={{
                fontSize: '0.75rem',
                color: '#3b82f6',
                marginTop: '0.25rem',
                display: 'block',
              }}
            >
              POs synced: {result.cardsSync.updatedPos.join(', ')}
            </span>
          )}
        </div>
      )}

      <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: '1rem' }}>
        <span
          style={{
            fontSize: '0.8rem',
            color: '#64748b',
            display: 'block',
            marginBottom: '0.25rem',
          }}
        >
          Details of Bills File:
        </span>
        <code
          style={{
            fontSize: '0.8rem',
            wordBreak: 'break-all',
            display: 'block',
            backgroundColor: '#f8fafc',
            padding: '0.4rem',
            borderRadius: '4px',
            border: '1px solid #e2e8f0',
          }}
        >
          {result.detailsExcelPath}
        </code>
      </div>
    </div>
  );
}

export default DetailsOfBillsResult;
