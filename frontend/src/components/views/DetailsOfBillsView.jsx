import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function DetailsOfBillsView() {
  const [detailsExcelPath, setDetailsExcelPath] = useState('');
  const [invoicesFolderPath, setInvoicesFolderPath] = useState('');
  const [budgetCardsPath, setBudgetCardsPath] = useState('');
  const [financialYear, setFinancialYear] = useState('APRIL 2026 to MARCH 2027');

  const [loading, setLoading] = useState(false);
  const [browseDetailsLoading, setBrowseDetailsLoading] = useState(false);
  const [browseInvoicesLoading, setBrowseInvoicesLoading] = useState(false);
  const [browseCardsLoading, setBrowseCardsLoading] = useState(false);
  
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseDetailsFile = async () => {
    setBrowseDetailsLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setDetailsExcelPath(data.filePath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseDetailsLoading(false);
    }
  };

  const handleBrowseInvoicesFolder = async () => {
    setBrowseInvoicesLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-folder', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.folderPath) {
        setInvoicesFolderPath(data.folderPath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open folder browser dialog.');
    } finally {
      setBrowseInvoicesLoading(false);
    }
  };

  const handleBrowseCardsFile = async () => {
    setBrowseCardsLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setBudgetCardsPath(data.filePath);
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

  const handleSyncDetailsOfBills = async (e) => {
    if (e) e.preventDefault();
    if (!detailsExcelPath) {
      setErrorMsg('Please select or specify the Details of Bills Excel file path.');
      return;
    }
    if (!invoicesFolderPath) {
      setErrorMsg('Please select the folder where raised invoices are stored.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/sync-details-of-bills', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          detailsExcelPath: detailsExcelPath.trim(),
          invoicesFolderPath: invoicesFolderPath.trim(),
          budgetCardsPath: budgetCardsPath.trim() || undefined,
          financialYear: financialYear.trim() || 'APRIL 2026 to MARCH 2027'
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to synchronize Details of Bills.');
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
        <h2>Details of Bills &amp; Cards Sync</h2>
        <p className="subtitle">
          Consolidate raised tax invoices into Master Details of Bills (Sheet 1) matching the official layout, with smart deduplication and automatic synchronization of IV numbers &amp; dates into Budget PO summary cards.
        </p>
      </div>

      <div className="view-content" style={{ display: 'grid', gridTemplateColumns: 'minmax(400px, 1fr) minmax(320px, 1fr)', gap: '1.5rem', alignItems: 'start' }}>
        <div className="card form-card">
          <form onSubmit={handleSyncDetailsOfBills}>
            {/* Details of Bills Excel Path */}
            <BrowseField
              label="1. Details of Bills Master Excel File"
              value={detailsExcelPath}
              onChange={setDetailsExcelPath}
              onBrowse={handleBrowseDetailsFile}
              loading={browseDetailsLoading}
              placeholder="e.g. D:\SRIBALAJITRADERS9\Details of Bills 2026 TO 2027.xlsx"
              helpText="Select existing master file or enter a new path. If empty, the table format and top summary formulas will be created automatically."
              required
            />

            {/* Invoices Folder Path */}
            <BrowseField
              label="2. Raised Invoices Folder"
              value={invoicesFolderPath}
              onChange={setInvoicesFolderPath}
              onBrowse={handleBrowseInvoicesFolder}
              loading={browseInvoicesLoading}
              placeholder="e.g. D:\SRIBALAJITRADERS9\Invoices"
              helpText="Select the folder containing newly created tax invoice Excel (.xlsx) files to be scanned and appended."
              required
            />

            {/* Budget PO Summary Cards Path (Optional) */}
            <BrowseField
              label="3. Budget PO Summary Cards Excel (Optional)"
              value={budgetCardsPath}
              onChange={setBudgetCardsPath}
              onBrowse={handleBrowseCardsFile}
              loading={browseCardsLoading}
              placeholder="e.g. D:\SRIBALAJITRADERS9\Nandyala FMC Budget.xlsx"
              helpText="Optional: If selected, the program will automatically populate the IV Number (Col A) and Date (Col B) in matching PO cards."
            />

            {/* Financial Year / Header Title */}
            <div className="form-group" style={{ marginBottom: '1.25rem' }}>
              <label style={{ display: 'block', fontWeight: '600', marginBottom: '0.4rem', color: 'var(--text-color)' }}>
                Financial Year / Header Title
              </label>
              <input
                type="text"
                className="input-field"
                value={financialYear}
                onChange={(e) => setFinancialYear(e.target.value)}
                placeholder="APRIL 2026 to MARCH 2027"
                style={{ width: '100%' }}
              />
              <span className="help-text" style={{ fontSize: '0.8rem', color: '#666', marginTop: '0.25rem', display: 'block' }}>
                Header displayed across H2:K2 on Sheet 1 of Details of Bills.
              </span>
            </div>

            {errorMsg && (
              <div className="error-banner" style={{ marginTop: '1rem', padding: '0.75rem', backgroundColor: '#fee2e2', color: '#b91c1c', borderRadius: '6px' }}>
                ⚠️ {errorMsg}
              </div>
            )}

            <button
              type="submit"
              className="btn btn-primary"
              disabled={loading}
              style={{ width: '100%', marginTop: '1.5rem', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '0.5rem', padding: '0.75rem' }}
            >
              {loading ? (
                <>
                  <span className="spinner"></span> Synchronizing Details of Bills...
                </>
              ) : (
                <>
                  <span>📑</span> Sync Invoices to Details of Bills &amp; Cards
                </>
              )}
            </button>
          </form>
        </div>

        {/* Results Panel */}
        <div>
          {result ? (
            <div className="card result-card" style={{ padding: '1.5rem', backgroundColor: 'var(--card-bg)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem' }}>
                <span style={{ fontSize: '1.5rem' }}>✅</span>
                <div>
                  <h3 style={{ margin: 0, color: '#16a34a' }}>Sync Completed</h3>
                  <p style={{ margin: 0, fontSize: '0.85rem', color: '#666' }}>{result.message}</p>
                </div>
              </div>

              <div className="stats-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem', marginBottom: '1.25rem' }}>
                <div style={{ padding: '0.75rem', backgroundColor: '#f8fafc', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>Invoices Scanned</span>
                  <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>{result.totalInvoicesFound}</strong>
                </div>
                <div style={{ padding: '0.75rem', backgroundColor: '#f0fdf4', borderRadius: '6px', border: '1px solid #bbf7d0' }}>
                  <span style={{ fontSize: '0.75rem', color: '#16a34a', display: 'block' }}>New Invoices Appended</span>
                  <strong style={{ fontSize: '1.2rem', color: '#15803d' }}>{result.appendedInvoices?.length || 0}</strong>
                </div>
                <div style={{ padding: '0.75rem', backgroundColor: '#f8fafc', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>Activity Rows Added</span>
                  <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>{result.totalRowsAdded || 0}</strong>
                </div>
                <div style={{ padding: '0.75rem', backgroundColor: '#f8fafc', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <span style={{ fontSize: '0.75rem', color: '#64748b', display: 'block' }}>Total Rows in Sheet</span>
                  <strong style={{ fontSize: '1.2rem', color: '#0f172a' }}>{result.totalRowsInSheet || 0}</strong>
                </div>
              </div>

              {result.appendedInvoices && result.appendedInvoices.length > 0 && (
                <div style={{ marginBottom: '1rem' }}>
                  <label style={{ fontSize: '0.85rem', fontWeight: '600', color: '#334155', display: 'block', marginBottom: '0.4rem' }}>
                    Appended Invoices:
                  </label>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem' }}>
                    {result.appendedInvoices.map((iv, idx) => (
                      <span key={idx} style={{ padding: '0.2rem 0.5rem', backgroundColor: '#dcfce7', color: '#15803d', borderRadius: '4px', fontSize: '0.8rem', fontWeight: '600' }}>
                        IV #{iv}
                      </span>
                    ))}
                  </div>
                </div>
              )}

              {result.skippedInvoices && result.skippedInvoices.length > 0 && (
                <div style={{ marginBottom: '1rem' }}>
                  <label style={{ fontSize: '0.85rem', fontWeight: '600', color: '#64748b', display: 'block', marginBottom: '0.4rem' }}>
                    Already Present (Skipped):
                  </label>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem' }}>
                    {result.skippedInvoices.map((iv, idx) => (
                      <span key={idx} style={{ padding: '0.2rem 0.5rem', backgroundColor: '#f1f5f9', color: '#64748b', borderRadius: '4px', fontSize: '0.8rem' }}>
                        IV #{iv}
                      </span>
                    ))}
                  </div>
                </div>
              )}

              {result.cardsSync && (
                <div style={{ padding: '0.75rem', backgroundColor: '#eff6ff', borderRadius: '6px', border: '1px solid #bfdbfe', marginBottom: '1rem' }}>
                  <span style={{ fontSize: '0.8rem', color: '#1e40af', fontWeight: '600', display: 'block' }}>
                    Budget Cards Updated: {result.cardsSync.cardsUpdated} card(s)
                  </span>
                  {result.cardsSync.updatedPos && result.cardsSync.updatedPos.length > 0 && (
                    <span style={{ fontSize: '0.75rem', color: '#3b82f6', marginTop: '0.25rem', display: 'block' }}>
                      POs synced: {result.cardsSync.updatedPos.join(', ')}
                    </span>
                  )}
                </div>
              )}

              <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: '1rem' }}>
                <span style={{ fontSize: '0.8rem', color: '#64748b', display: 'block', marginBottom: '0.25rem' }}>Details of Bills File:</span>
                <code style={{ fontSize: '0.8rem', wordBreak: 'break-all', display: 'block', backgroundColor: '#f8fafc', padding: '0.4rem', borderRadius: '4px', border: '1px solid #e2e8f0' }}>
                  {result.detailsExcelPath}
                </code>
              </div>
            </div>
          ) : (
            <div className="card info-card" style={{ padding: '1.5rem', backgroundColor: 'var(--card-bg)', color: '#64748b', textAlign: 'center' }}>
              <span style={{ fontSize: '2.5rem', display: 'block', marginBottom: '0.5rem' }}>📑</span>
              <h4 style={{ margin: '0 0 0.5rem 0', color: 'var(--text-color)' }}>Smart Invoice Consolidation</h4>
              <p style={{ fontSize: '0.85rem', lineHeight: '1.4', margin: 0 }}>
                Select your <strong>Details of Bills Excel</strong> and the <strong>Invoices Folder</strong>. The system will inspect each invoice workbook, parse individual activity and TBM breakdowns, append only unlisted invoices into Sheet 1, update Row 3 total formulas, and populate IV numbers into your Budget Cards.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default DetailsOfBillsView;
