import React, { useState, useEffect } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function InvoiceGeneratorView() {
  const [company, setCompany] = useState('Corteva');
  const [area, setArea] = useState('Nellore');
  const [tbmSummaryPath, setTbmSummaryPath] = useState('');
  const [saveFolderPath, setSaveFolderPath] = useState('');
  const [invoiceNumber, setInvoiceNumber] = useState('');
  const [poNumber, setPoNumber] = useState('');
  const [serviceChargePercent, setServiceChargePercent] = useState('5.0');
  const [invoiceDate, setInvoiceDate] = useState(() => {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  });
  
  const [poValue, setPoValue] = useState('250000');
  const [requesterName, setRequesterName] = useState('');
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [detectedPOs, setDetectedPOs] = useState([]);

  const [loading, setLoading] = useState(false);
  const [browseSummaryLoading, setBrowseSummaryLoading] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  const [scanLoading, setScanLoading] = useState(false);

  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  // Available areas based on company
  const cortevaAreas = ['Nellore', 'Kurnool', 'Suryapet'];
  const fmcAreas = ['Nellore', 'Nandyala', 'Kurnool'];
  const currentAreas = company === 'Corteva' ? cortevaAreas : fmcAreas;

  // Update default service charge & area when switching companies
  const handleCompanyChange = (newComp) => {
    setCompany(newComp);
    if (newComp === 'Corteva') {
      setServiceChargePercent('5.0');
      if (!cortevaAreas.includes(area)) {
        setArea('Nellore');
      }
    } else {
      setServiceChargePercent('4.5');
      if (!fmcAreas.includes(area)) {
        setArea('Nellore');
      }
    }
  };

  const handleBrowseSummaryFile = async () => {
    setBrowseSummaryLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setTbmSummaryPath(data.filePath);
        // Auto scan POs
        scanPOs(data.filePath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseSummaryLoading(false);
    }
  };

  const scanPOs = async (filePath) => {
    if (!filePath) return;
    setScanLoading(true);
    try {
      const res = await fetch('/api/scan-pos-in-summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tbmSummaryPath: filePath.trim() })
      });
      const data = await res.json();
      if (res.ok && data.success && data.pos) {
        setDetectedPOs(data.pos);
        if (data.pos.length > 0 && !poNumber) {
          setPoNumber(data.pos[0]);
        }
      }
    } catch (e) {
      console.error('Scan POs error:', e);
    } finally {
      setScanLoading(false);
    }
  };

  const handleBrowseFolder = async () => {
    setBrowseFolderLoading(true);
    setErrorMsg(null);
    try {
      const res = await fetch('/api/browse-folder', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.folderPath) {
        setSaveFolderPath(data.folderPath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open folder browser dialog.');
    } finally {
      setBrowseFolderLoading(false);
    }
  };

  const handleGenerateInvoice = async (e) => {
    if (e) e.preventDefault();
    if (!tbmSummaryPath) {
      setErrorMsg('Please select the All-TBMs Summary Excel file.');
      return;
    }
    if (!saveFolderPath) {
      setErrorMsg('Please select the Save Folder for invoices.');
      return;
    }
    if (!invoiceNumber.trim()) {
      setErrorMsg('Please enter the Invoice Number to be raised/updated (e.g. SBT26270069 or 69).');
      return;
    }
    if (!poNumber.trim()) {
      setErrorMsg('Please enter or select the mandatory PO Number.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/generate-invoices', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          company,
          area,
          tbmSummaryPath: tbmSummaryPath.trim(),
          saveFolderPath: saveFolderPath.trim(),
          invoiceNumber: invoiceNumber.trim(),
          poNumber: poNumber.trim(),
          serviceChargePercent: parseFloat(serviceChargePercent) || (company === 'Corteva' ? 5.0 : 4.5),
          invoiceDate: invoiceDate.trim(),
          poValue: parseFloat(poValue) || 250000,
          requesterName: requesterName.trim() || undefined
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to generate/update invoice.');
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
        <h2>Tax Invoice Generator</h2>
        <p className="subtitle">
          Generate or update formal PO-wise Tax Invoices for Corteva and FMC with Sheet1 (Tax Invoice), Sheet2 (TBM Activity Expenses Breakdown), and Corteva Invoice Summary sheet.
        </p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.45fr 1fr', gap: '2rem' }}>
        {/* Left Form Card */}
        <div className="card">
          {/* Company Selector Header */}
          <div style={{ marginBottom: '1.25rem' }}>
            <label className="form-label" style={{ fontSize: '0.92rem', fontWeight: 'bold', marginBottom: '0.6rem', display: 'block' }}>
              Select Client / Organization
            </label>
            <div style={{ display: 'flex', gap: '1rem' }}>
              <button
                type="button"
                className={company === 'Corteva' ? 'primary' : 'secondary'}
                onClick={() => handleCompanyChange('Corteva')}
                style={{ flex: 1, padding: '0.75rem', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem', fontWeight: 'bold' }}
              >
                🌾 Corteva Agriscience
              </button>
              <button
                type="button"
                className={company === 'FMC' ? 'primary' : 'secondary'}
                onClick={() => handleCompanyChange('FMC')}
                style={{ flex: 1, padding: '0.75rem', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem', fontWeight: 'bold' }}
              >
                🌱 FMC (New Gen)
              </button>
            </div>
          </div>

          <form onSubmit={handleGenerateInvoice} style={{ display: 'flex', flexDirection: 'column', gap: '1.15rem' }}>
            {/* Area Dropdown Field */}
            <div className="form-field">
              <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
                Operational Area ({company}) <span style={{ color: '#e74c3c' }}>*</span>
              </label>
              <select
                className="input-text"
                style={{ width: '100%', fontSize: '0.95rem', padding: '0.65rem 0.85rem', fontWeight: '600' }}
                value={area}
                onChange={(e) => setArea(e.target.value)}
                disabled={loading}
              >
                {currentAreas.map((a) => (
                  <option key={a} value={a}>
                    📍 {a}
                  </option>
                ))}
              </select>
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: '0.25rem', display: 'block' }}>
                <em>Selected area will be printed directly in the AREA field on the Tax Invoice (Row {company === 'Corteva' ? '16' : '17'}).</em>
              </span>
            </div>

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
                onChange={(e) => setInvoiceNumber(e.target.value)}
                placeholder={company === 'Corteva' ? "e.g. SBT26270069 or 69" : "e.g. SBT26270073 or 73"}
                disabled={loading}
                required
              />
              <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: '0.3rem', display: 'block' }}>
                💡 <em>If an invoice file with this number exists in the save folder, newly received bills will be appended and totals updated. Otherwise, a new invoice is created.</em>
              </span>
            </div>

            {/* All-TBMs Summary File (Placed above PO Number) */}
            <BrowseField 
              label="All-TBMs Summary Excel File (e.g. All-TBMs-Summary.xlsx)"
              value={tbmSummaryPath}
              onChange={(e) => {
                setTbmSummaryPath(e.target.value);
                if (e.target.value) scanPOs(e.target.value);
              }}
              onBrowse={handleBrowseSummaryFile}
              browseLoading={browseSummaryLoading || scanLoading}
              disabled={loading}
              placeholder="Select All-TBMs Summary workbook containing employee bills..."
            />

            {/* PO Number Mandatory Field (Directly beneath Summary Picker) */}
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
                onChange={(e) => setPoNumber(e.target.value)}
                placeholder={company === 'Corteva' ? "e.g. 4800108506" : "e.g. 500BB2026018404"}
                disabled={loading}
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
                      onClick={() => setPoNumber(p)}
                      style={{
                        padding: '0.18rem 0.55rem',
                        fontSize: '0.78rem',
                        borderRadius: '4px',
                        border: '1px solid var(--border-color)',
                        background: poNumber === p ? 'var(--primary-color)' : 'var(--bg-hover)',
                        color: poNumber === p ? '#fff' : 'var(--text-color)',
                        cursor: 'pointer',
                        fontWeight: '600'
                      }}
                    >
                      {p}
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* Invoices Save Directory */}
            <BrowseField 
              label="Invoices Save Folder / Directory"
              value={saveFolderPath}
              onChange={(e) => setSaveFolderPath(e.target.value)}
              onBrowse={handleBrowseFolder}
              browseLoading={browseFolderLoading}
              disabled={loading}
              placeholder="Select destination folder to save or update invoice Excel workbooks..."
            />

            {/* Service Charges and Invoice Date Row */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
              <div className="form-field">
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
                  placeholder="e.g. 5.0 or 4.5"
                  disabled={loading}
                />
              </div>

              <div className="form-field">
                <label className="form-label" style={{ fontSize: '0.88rem', fontWeight: '600' }}>
                  Invoice Date (DD-MM-YYYY)
                </label>
                <input
                  type="text"
                  className="input-text"
                  style={{ width: '100%', fontSize: '0.92rem', padding: '0.55rem 0.75rem' }}
                  value={invoiceDate}
                  onChange={(e) => setInvoiceDate(e.target.value)}
                  placeholder="e.g. 20-08-2026"
                  disabled={loading}
                />
              </div>
            </div>

            {/* Optional / Advanced Settings Accordion */}
            <div>
              <button
                type="button"
                onClick={() => setShowAdvanced(!showAdvanced)}
                style={{
                  background: 'none',
                  border: 'none',
                  color: 'var(--primary-color)',
                  fontSize: '0.82rem',
                  cursor: 'pointer',
                  padding: 0,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '0.3rem',
                  fontWeight: '600'
                }}
              >
                {showAdvanced ? '▼ Hide Additional Details' : '▶ Show Additional Details (PO Value, Requester)'}
              </button>

              {showAdvanced && (
                <div style={{
                  marginTop: '0.75rem',
                  padding: '0.85rem',
                  background: 'var(--bg-hover)',
                  borderRadius: '6px',
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: '0.85rem'
                }}>
                  <div className="form-field">
                    <label className="form-label" style={{ fontSize: '0.8rem' }}>PO Value (₹)</label>
                    <input
                      type="number"
                      className="input-text"
                      style={{ width: '100%', fontSize: '0.85rem', padding: '0.45rem' }}
                      value={poValue}
                      onChange={(e) => setPoValue(e.target.value)}
                      placeholder="e.g. 250000"
                      disabled={loading}
                    />
                  </div>
                  <div className="form-field">
                    <label className="form-label" style={{ fontSize: '0.8rem' }}>Requester / ZDGM Name</label>
                    <input
                      type="text"
                      className="input-text"
                      style={{ width: '100%', fontSize: '0.85rem', padding: '0.45rem' }}
                      value={requesterName}
                      onChange={(e) => setRequesterName(e.target.value)}
                      placeholder="e.g. R.Bhaskar or Madhavareddy"
                      disabled={loading}
                    />
                  </div>
                </div>
              )}
            </div>

            <button
              type="submit"
              className="primary"
              disabled={loading || !tbmSummaryPath || !saveFolderPath || !invoiceNumber || !poNumber}
              style={{ marginTop: '0.5rem', padding: '0.9rem', fontWeight: 'bold' }}
            >
              {loading ? '⏳ Generating / Updating Tax Invoice...' : '🧾 Generate / Update Tax Invoice'}
            </button>
          </form>
        </div>

        {/* Right Status / Result Panel */}
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
              <h3>Processing Invoice...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Extracting PO bills from All-TBMs summary, formatting Sheet 1 (Tax Invoice with {area} Area), Sheet 2 (Activity Expenses), and applying GST formulas.
              </p>
            </div>
          ) : result ? (
            <div>
              <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
                <h4 style={{ color: '#27ae60', margin: 0 }}>
                  {result.isUpdate ? '🔄 Invoice Updated & Appended' : '✓ Invoice Generated Successfully'}
                </h4>
                <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
                  {result.message}
                </p>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Invoice Number</span>
                  <div style={{ fontSize: '1.1rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {result.invoiceNo}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>PO Number</span>
                  <div style={{ fontSize: '1.1rem', fontWeight: 'bold', color: '#27ae60' }}>
                    {result.poNumber}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Area</span>
                  <div style={{ fontSize: '1.05rem', fontWeight: 'bold', color: 'var(--text-color)' }}>
                    📍 {result.area || area}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Grand Total (Inc. GST)</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: '#27ae60' }}>
                    ₹ {result.grandTotalIncGst?.toLocaleString('en-IN')}
                  </div>
                </div>
              </div>

              {result.grandTotalWords && (
                <div style={{
                  background: 'var(--bg-hover)',
                  border: '1px solid var(--border-color)',
                  padding: '0.75rem 1rem',
                  borderRadius: '6px',
                  marginBottom: '1rem',
                  fontSize: '0.84rem'
                }}>
                  <span style={{ color: 'var(--text-muted)', fontSize: '0.75rem', display: 'block' }}>Amount in Words:</span>
                  <strong>{result.grandTotalWords}</strong>
                </div>
              )}

              <ResultPanel result={result} isSummary={true} />
            </div>
          ) : (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
              <h3>Ready to Generate Invoices</h3>
              <p style={{ marginTop: '0.5rem' }}>
                Select your client (Corteva or FMC), choose the operational Area from the dropdown, pick the All-TBMs summary file to reveal detected POs, enter your invoice number, then click <strong>Generate / Update Tax Invoice</strong>.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default InvoiceGeneratorView;
