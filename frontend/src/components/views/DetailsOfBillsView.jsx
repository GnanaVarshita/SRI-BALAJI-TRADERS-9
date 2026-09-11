import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import DetailsOfBillsResult from '../details/DetailsOfBillsResult';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';

function DetailsOfBillsView() {
  const [detailsExcelPath, setDetailsExcelPath] = useState('');
  const [invoicesFolderPath, setInvoicesFolderPath] = useState('');
  const [budgetCardsPath, setBudgetCardsPath] = useState('');
  const [financialYear, setFinancialYear] = useState('APRIL 2026 to MARCH 2027');

  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const detailsBrowser = useFileBrowser();
  const invoicesBrowser = useFileBrowser();
  const cardsBrowser = useFileBrowser();

  const handleBrowseDetailsFile = () => {
    setErrorMsg(null);
    detailsBrowser.browseFile((filePath) => {
      setDetailsExcelPath(filePath);
    });
  };

  const handleBrowseInvoicesFolder = () => {
    setErrorMsg(null);
    invoicesBrowser.browseFolder((folderPath) => {
      setInvoicesFolderPath(folderPath);
    });
  };

  const handleBrowseCardsFile = () => {
    setErrorMsg(null);
    cardsBrowser.browseFile((filePath) => {
      setBudgetCardsPath(filePath);
    });
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
      const data = await api.syncDetailsOfBills({
        detailsExcelPath: detailsExcelPath.trim(),
        invoicesFolderPath: invoicesFolderPath.trim(),
        budgetCardsPath: budgetCardsPath.trim() || undefined,
        financialYear: financialYear.trim() || 'APRIL 2026 to MARCH 2027',
      });
      setResult(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to synchronize Details of Bills.');
    } finally {
      setLoading(false);
    }
  };

  const activeError =
    errorMsg ||
    detailsBrowser.error ||
    invoicesBrowser.error ||
    cardsBrowser.error;

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Details of Bills &amp; Cards Sync</h2>
        <p className="subtitle">
          Consolidate raised tax invoices into Master Details of Bills (Sheet 1) matching the official layout, with smart deduplication and automatic synchronization of IV numbers &amp; dates into Budget PO summary cards.
        </p>
      </div>

      <div
        className="view-content"
        style={{
          display: 'grid',
          gridTemplateColumns: 'minmax(400px, 1fr) minmax(320px, 1fr)',
          gap: '1.5rem',
          alignItems: 'start',
        }}
      >
        <div className="card form-card">
          <form onSubmit={handleSyncDetailsOfBills}>
            {/* Details of Bills Excel Path */}
            <BrowseField
              label="1. Details of Bills Master Excel File"
              value={detailsExcelPath}
              onChange={setDetailsExcelPath}
              onBrowse={handleBrowseDetailsFile}
              loading={detailsBrowser.loading}
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
              loading={invoicesBrowser.loading}
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
              loading={cardsBrowser.loading}
              placeholder="e.g. D:\SRIBALAJITRADERS9\Nandyala FMC Budget.xlsx"
              helpText="Optional: If selected, the program will automatically populate the IV Number (Col A) and Date (Col B) in matching PO cards."
            />

            {/* Financial Year / Header Title */}
            <div className="form-group" style={{ marginBottom: '1.25rem' }}>
              <label
                style={{
                  display: 'block',
                  fontWeight: '600',
                  marginBottom: '0.4rem',
                  color: 'var(--text-color)',
                }}
              >
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
              <span
                className="help-text"
                style={{
                  fontSize: '0.8rem',
                  color: '#666',
                  marginTop: '0.25rem',
                  display: 'block',
                }}
              >
                Header displayed across H2:K2 on Sheet 1 of Details of Bills.
              </span>
            </div>

            {activeError && (
              <div
                className="error-banner"
                style={{
                  marginTop: '1rem',
                  padding: '0.75rem',
                  backgroundColor: '#fee2e2',
                  color: '#b91c1c',
                  borderRadius: '6px',
                }}
              >
                ⚠️ {activeError}
              </div>
            )}

            <button
              type="submit"
              className="btn btn-primary"
              disabled={loading}
              style={{
                width: '100%',
                marginTop: '1.5rem',
                display: 'flex',
                justifyContent: 'center',
                alignItems: 'center',
                gap: '0.5rem',
                padding: '0.75rem',
              }}
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
        <DetailsOfBillsResult result={result} />
      </div>
    </div>
  );
}

export default DetailsOfBillsView;
