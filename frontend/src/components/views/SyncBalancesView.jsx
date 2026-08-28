import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function SyncBalancesView() {
  const [cardsSummaryPath, setCardsSummaryPath] = useState('');
  const [tbmSummaryPath, setTbmSummaryPath] = useState('');
  const [serviceChargePercent, setServiceChargePercent] = useState('5');

  const [loading, setLoading] = useState(false);
  const [browseCardsLoading, setBrowseCardsLoading] = useState(false);
  const [browseTbmLoading, setBrowseTbmLoading] = useState(false);
  
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseCardsFile = async () => {
    setBrowseCardsLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setCardsSummaryPath(data.filePath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseCardsLoading(false);
    }
  };

  const handleBrowseTbmFile = async () => {
    setBrowseTbmLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setTbmSummaryPath(data.filePath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseTbmLoading(false);
    }
  };

  const handleSyncBalances = async (e) => {
    if (e) e.preventDefault();
    if (!cardsSummaryPath) {
      setErrorMsg('Please select the PO Cards Summary Excel file.');
      return;
    }
    if (!tbmSummaryPath) {
      setErrorMsg('Please select the Consolidated Master TBM Summary Excel file.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/sync-tbm-cards', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          cardsExcelPath: cardsSummaryPath.trim(),
          tbmSummaryPath: tbmSummaryPath.trim(),
          serviceChargePercent: parseFloat(serviceChargePercent) || 5.0
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to synchronize PO Cards summary.');
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Network error: Failed to contact the backend server.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Sync Spent &amp; Balances</h2>
        <p className="subtitle">
          Synchronize TBM spent details from Consolidated Master TBM Summary into PO summary cards with automatic Service Charges &amp; Balance formulas, reflecting balances in Sheet 1 master overview.
        </p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.45fr 1fr', gap: '2rem' }}>
        {/* Input Form Card */}
        <div className="card">
          <h2>Input Workbooks &amp; Service Charges</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '1.25rem' }}>
            Select the PO Cards Excel and the Consolidated Master TBM Summary Excel.
          </p>

          <form onSubmit={handleSyncBalances} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <BrowseField 
              label="1. PO Cards Summary Excel File (e.g. Nandyala FMC Budget.xlsx)"
              value={cardsSummaryPath}
              onChange={(e) => setCardsSummaryPath(e.target.value)}
              onBrowse={handleBrowseCardsFile}
              browseLoading={browseCardsLoading}
              disabled={loading}
              placeholder="Select PO Summary workbook with cards (e.g. Nandyala FMC Budget.xlsx)..."
            />

            <BrowseField 
              label="2. Consolidated Master TBM Summary Excel File (e.g. NANDYALA-All-TBMs-Summary.xlsx)"
              value={tbmSummaryPath}
              onChange={(e) => setTbmSummaryPath(e.target.value)}
              onBrowse={handleBrowseTbmFile}
              browseLoading={browseTbmLoading}
              disabled={loading}
              placeholder="Select Consolidated Master TBM Summary with TBM Amount Summary sheet..."
            />

            <div className="form-field" style={{ display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
              <div style={{ flex: 1 }}>
                <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
                  Service Charges (%)
                </label>
                <input
                  type="number"
                  step="0.01"
                  min="0"
                  max="100"
                  className="input-text"
                  style={{ width: '100%', fontSize: '0.92rem', padding: '0.55rem 0.75rem' }}
                  value={serviceChargePercent}
                  onChange={(e) => setServiceChargePercent(e.target.value)}
                  placeholder="e.g. 5"
                  disabled={loading}
                />
              </div>
              <div style={{ flex: 1.5, paddingTop: '1.25rem', fontSize: '0.82rem', color: 'var(--text-muted)' }}>
                Default is <strong>5.0%</strong>. Automatically generates Excel formulas:
                <div style={{ marginTop: '0.25rem', fontFamily: 'Consolas, monospace', color: 'var(--primary-color)' }}>
                  Spent + SV Charges = Total IV
                </div>
                <div style={{ fontFamily: 'Consolas, monospace', color: 'var(--primary-color)' }}>
                  Budget - Total IV = Balance
                </div>
              </div>
            </div>

            {/* Info Box */}
            <div style={{
              background: 'var(--bg-hover)',
              border: '1px solid var(--border-color)',
              padding: '0.9rem 1rem',
              borderRadius: '8px',
              fontSize: '0.82rem',
              color: 'var(--text-muted)',
              lineHeight: '1.5'
            }}>
              <strong>📌 Automatic Operations Performed:</strong>
              <ul style={{ margin: '0.4rem 0 0 1.1rem', padding: 0 }}>
                <li>Populates card data rows with TBM spent amounts under matching activity columns.</li>
                <li>Leaves <code>I.V NO</code> and <code>DATE</code> columns blank for invoice verification.</li>
                <li>Generates sum formulas for table totals and right summary blocks.</li>
                <li>Calculates <code>SV Charges</code>, <code>TOTAL IV</code>, and <code>BALANCE</code> per activity.</li>
                <li>Links <code>Spent Budget</code> and <code>Balance</code> in <strong>Sheet 1</strong> master overview.</li>
              </ul>
            </div>

            <button
              type="submit"
              className="primary"
              disabled={loading || !cardsSummaryPath || !tbmSummaryPath}
              style={{ marginTop: '0.5rem', padding: '0.9rem', fontWeight: 'bold' }}
            >
              {loading ? '⏳ Synchronizing PO Cards & Balances...' : '⚖️ Synchronize Spent & Balances'}
            </button>
          </form>
        </div>

        {/* Results Panel Card */}
        <div className="card">
          <h2>Execution Status</h2>
          
          {errorMsg && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ⚠️ {errorMsg}
            </div>
          )}

          {loading ? (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
              <h3>Synchronizing PO Cards &amp; Balances...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Reading TBM expended details, injecting card data rows, and applying formulas for Service Charges ({serviceChargePercent}%), Total IV, and Balances.
              </p>
            </div>
          ) : result ? (
            <div>
              <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
                <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Synchronization Complete</h4>
                <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
                  {result.message}
                </p>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>PO Cards Updated</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {result.updatedCards}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>TBM POs Found</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {result.totalTbmPOs}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Service Charges</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {result.serviceChargePercent}%
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Formulas Created</span>
                  <div style={{ fontSize: '0.88rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    Spent + SV = Total IV
                  </div>
                </div>
              </div>

              <ResultPanel result={result} isSummary={true} />
            </div>
          ) : (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
              <h3>Ready to Synchronize</h3>
              <p style={{ marginTop: '0.5rem' }}>
                Select the PO cards summary workbook and consolidated TBM summary workbook on the left, then click <strong>Synchronize Spent &amp; Balances</strong>.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default SyncBalancesView;
