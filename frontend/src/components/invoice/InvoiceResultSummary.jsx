import React from 'react';

function InvoiceResultSummary({ result }) {
  if (!result) return null;

  return (
    <div style={{
      marginTop: '1.5rem',
      padding: '1.25rem',
      background: 'rgba(39, 174, 96, 0.08)',
      border: '1px solid rgba(39, 174, 96, 0.35)',
      borderRadius: '8px'
    }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1rem' }}>
        <h3 style={{ margin: 0, color: '#27ae60', fontSize: '1.1rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <span>✓</span> {result.message}
        </h3>
        <span style={{
          fontSize: '0.78rem',
          padding: '0.2rem 0.6rem',
          borderRadius: '12px',
          background: result.isUpdate ? '#e67e22' : '#27ae60',
          color: '#fff',
          fontWeight: 'bold'
        }}>
          {result.isUpdate ? 'UPDATED' : 'NEW INVOICE'}
        </span>
      </div>

      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
        gap: '0.75rem',
        marginBottom: '1rem'
      }}>
        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '0.6rem', borderRadius: '4px' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>Invoice No</span>
          <strong style={{ fontSize: '0.95rem' }}>{result.invoiceNo}</strong>
        </div>
        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '0.6rem', borderRadius: '4px' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>PO Number</span>
          <strong style={{ fontSize: '0.95rem', color: '#27ae60' }}>{result.poNumber}</strong>
        </div>
        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '0.6rem', borderRadius: '4px' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>Area</span>
          <strong style={{ fontSize: '0.95rem' }}>{result.area || 'General'}</strong>
        </div>
        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '0.6rem', borderRadius: '4px' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>Activities Count</span>
          <strong style={{ fontSize: '0.95rem' }}>{result.totalActivities}</strong>
        </div>
        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '0.6rem', borderRadius: '4px' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>Sub Total (Exc GST)</span>
          <strong style={{ fontSize: '0.95rem' }}>₹{Number(result.subTotalExcGst).toLocaleString('en-IN', { minimumFractionDigits: 2 })}</strong>
        </div>
        <div style={{ background: 'rgba(39, 174, 96, 0.12)', padding: '0.6rem', borderRadius: '4px', border: '1px solid rgba(39,174,96,0.3)' }}>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block' }}>Grand Total (Inc GST)</span>
          <strong style={{ fontSize: '1.05rem', color: '#2ecc71' }}>₹{Number(result.grandTotalIncGst).toLocaleString('en-IN')}</strong>
        </div>
      </div>

      {result.grandTotalWords && (
        <p style={{ fontSize: '0.82rem', margin: '0.5rem 0', color: 'var(--text-muted)' }}>
          <em>Amount in Words:</em> <strong style={{ color: 'var(--text-color)' }}>{result.grandTotalWords}</strong>
        </p>
      )}

      {result.outputPath && (
        <div style={{ marginTop: '0.75rem', fontSize: '0.82rem', wordBreak: 'break-all', color: 'var(--text-muted)' }}>
          📁 <strong>File saved at:</strong> <code>{result.outputPath}</code>
        </div>
      )}
    </div>
  );
}

export default InvoiceResultSummary;
