import React, { useState } from 'react';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';
import InvoiceCompanySelector from '../invoice/InvoiceCompanySelector';
import InvoicePoSelector from '../invoice/InvoicePoSelector';
import InvoiceDetailsForm from '../invoice/InvoiceDetailsForm';
import InvoiceResultSummary from '../invoice/InvoiceResultSummary';

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
  const [detectedPOs, setDetectedPOs] = useState([]);

  const [loading, setLoading] = useState(false);
  const [scanLoading, setScanLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const fileBrowser = useFileBrowser();
  const folderBrowser = useFileBrowser();

  const handleCompanyChange = (newComp) => {
    setCompany(newComp);
    if (newComp === 'Corteva') {
      setServiceChargePercent('5.0');
      if (!['Nellore', 'Kurnool', 'Suryapet'].includes(area)) {
        setArea('Nellore');
      }
    } else {
      setServiceChargePercent('4.5');
      if (!['Nellore', 'Nandyala', 'Kurnool'].includes(area)) {
        setArea('Nellore');
      }
    }
  };

  const scanPOs = async (filePath) => {
    if (!filePath) return;
    setScanLoading(true);
    try {
      const data = await api.scanPosInSummary(filePath.trim());
      if (data.success && data.pos) {
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

  const handleBrowseSummaryFile = () => {
    setErrorMsg(null);
    fileBrowser.browseFile((path) => {
      setTbmSummaryPath(path);
      scanPOs(path);
    });
  };

  const handleBrowseFolder = () => {
    setErrorMsg(null);
    folderBrowser.browseFolder((path) => {
      setSaveFolderPath(path);
    });
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
      const data = await api.generateInvoices({
        company,
        area,
        tbmSummaryPath: tbmSummaryPath.trim(),
        saveFolderPath: saveFolderPath.trim(),
        invoiceNumber: invoiceNumber.trim(),
        poNumber: poNumber.trim(),
        serviceChargePercent: parseFloat(serviceChargePercent) || (company === 'Corteva' ? 5.0 : 4.5),
        invoiceDate: invoiceDate.trim(),
        poValue: parseFloat(poValue) || 250000,
        requesterName: requesterName.trim() || undefined,
      });
      setResult(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to generate/update invoice.');
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
          <InvoiceCompanySelector
            company={company}
            onCompanyChange={handleCompanyChange}
            area={area}
            onAreaChange={setArea}
            disabled={loading}
          />

          <form onSubmit={handleGenerateInvoice} style={{ display: 'flex', flexDirection: 'column', gap: '1.15rem', marginTop: '1rem' }}>
            <InvoicePoSelector
              tbmSummaryPath={tbmSummaryPath}
              onSummaryPathChange={(val) => {
                setTbmSummaryPath(val);
                if (val) scanPOs(val);
              }}
              onBrowseSummary={handleBrowseSummaryFile}
              browseSummaryLoading={fileBrowser.loading || scanLoading}
              poNumber={poNumber}
              onPoNumberChange={setPoNumber}
              detectedPOs={detectedPOs}
              company={company}
              disabled={loading}
            />

            <InvoiceDetailsForm
              invoiceNumber={invoiceNumber}
              onInvoiceNumberChange={setInvoiceNumber}
              saveFolderPath={saveFolderPath}
              onSaveFolderPathChange={setSaveFolderPath}
              onBrowseFolder={handleBrowseFolder}
              browseFolderLoading={folderBrowser.loading}
              serviceChargePercent={serviceChargePercent}
              onServiceChargeChange={setServiceChargePercent}
              invoiceDate={invoiceDate}
              onInvoiceDateChange={setInvoiceDate}
              poValue={poValue}
              onPoValueChange={setPoValue}
              requesterName={requesterName}
              onRequesterNameChange={setRequesterName}
              company={company}
              disabled={loading}
            />

            {(errorMsg || fileBrowser.error || folderBrowser.error) && (
              <div style={{
                padding: '0.85rem',
                background: 'rgba(231, 76, 60, 0.1)',
                border: '1px solid rgba(231, 76, 60, 0.35)',
                borderRadius: '6px',
                color: '#e74c3c',
                fontSize: '0.88rem'
              }}>
                ⚠️ {errorMsg || fileBrowser.error || folderBrowser.error}
              </div>
            )}

            <button
              type="submit"
              className="primary"
              disabled={loading}
              style={{
                marginTop: '0.5rem',
                padding: '0.85rem',
                fontSize: '1rem',
                fontWeight: 'bold',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '0.6rem'
              }}
            >
              {loading ? (
                <><span>⏳</span> Processing Invoice...</>
              ) : (
                <><span>⚡</span> Generate / Update Tax Invoice</>
              )}
            </button>
          </form>

          <InvoiceResultSummary result={result} />
        </div>

        {/* Right Reference Card */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
          <div className="card">
            <h3 style={{ marginTop: 0, fontSize: '1.05rem', color: 'var(--primary-color)' }}>
              📑 Invoice Layout Architecture
            </h3>
            <div style={{ fontSize: '0.83rem', color: 'var(--text-muted)', lineHeight: '1.55' }}>
              <p>
                <strong>Sheet 1 (Tax Invoice):</strong>
                <br />Exact Sri Balaji Traders A4 template with <strong>ORIGINAL</strong> and <strong>DUPLICATE</strong> copies, balanced margins, auto HSN 998596, service charges, 9% CGST + 9% SGST, rounded off grand total, and Indian currency words.
              </p>
              <p>
                <strong>Sheet 2 (Detailed TBM Expense Breakdown):</strong>
                <br />Grouped either by Activity (Corteva) or by Territory (FMC), with exact formulas and full PO numbers.
              </p>
              {company === 'Corteva' && (
                <p>
                  <strong>Sheet 4 (Corteva Summary):</strong>
                  <br />Exact 11-column Corteva corporate summary linking dynamically to Sheet1 Subtotal and Grand Total.
                </p>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default InvoiceGeneratorView;
