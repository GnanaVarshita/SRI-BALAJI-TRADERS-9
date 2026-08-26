import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function FmcSummaryView() {
  const [activeStep, setActiveStep] = useState('step1'); // 'step1' or 'step2'

  // Step 1 states
  const [inputFolderPath, setInputFolderPath] = useState('');
  const [saveFolderPath, setSaveFolderPath] = useState('');

  // Step 2 states
  const [excelPath, setExcelPath] = useState('');

  const [loading, setLoading] = useState(false);
  const [browseInputLoading, setBrowseInputLoading] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  const [browseExcelLoading, setBrowseExcelLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  // Step 1 Browsers
  const handleBrowseInputFolder = async () => {
    setBrowseInputLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-folder', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.folderPath) {
        setInputFolderPath(data.folderPath);
        const parentFolder = data.folderPath.substring(0, data.folderPath.lastIndexOf(data.folderPath.includes('/') ? '/' : '\\'));
        setSaveFolderPath(parentFolder);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open folder browser dialog.');
    } finally {
      setBrowseInputLoading(false);
    }
  };

  const handleBrowseSaveFolder = async () => {
    setBrowseFolderLoading(true);
    setErrorMsg(null);
    setResult(null);
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

  // Step 2 Browser
  const handleBrowseExcel = async () => {
    setBrowseExcelLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setExcelPath(data.filePath);
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseExcelLoading(false);
    }
  };

  // Step 1 Submit
  const handleStep1Submit = async (e) => {
    e.preventDefault();
    if (!inputFolderPath) {
      setErrorMsg('Please select the FMC POs PDF folder first.');
      return;
    }
    if (!saveFolderPath) {
      setErrorMsg('Please select a folder to save the summary file.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/generate-fmc-summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inputFolderPath, saveFolderPath }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
        if (data.outputPath) {
          setExcelPath(data.outputPath); // Auto pre-fill Step 2!
        }
      } else {
        setErrorMsg(data.message || 'Failed to generate FMC Master Budget sheet.');
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Network error: Failed to contact the backend server.');
    } finally {
      setLoading(false);
    }
  };

  // Step 2 Submit
  const handleStep2Submit = async (e) => {
    e.preventDefault();
    if (!excelPath) {
      setErrorMsg('Please select the FMC Budget Excel file first.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/generate-fmc-step2', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ excelPath }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to generate PO Summary Cards.');
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
        <h2>FMC PO Summary Generator</h2>
        <p className="subtitle">Step 1: Build master budget table from PO PDFs. Step 2: Generate 11 PO summary cards per sheet in your workbook.</p>
      </div>

      {/* Step Tabs */}
      <div style={{ display: 'flex', gap: '1rem', marginBottom: '1.5rem' }}>
        <button 
          className={activeStep === 'step1' ? 'primary' : 'secondary'}
          onClick={() => setActiveStep('step1')}
          style={{ padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
           Step 1: Master Budget Table (From PDFs)
        </button>
        <button 
          className={activeStep === 'step2' ? 'primary' : 'secondary'}
          onClick={() => setActiveStep('step2')}
          style={{ padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
          Step 2: PO Summary Cards (11 per Sheet)
        </button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '2rem' }}>
        {/* Settings Form Card */}
        <div className="card">
          {activeStep === 'step1' ? (
            <>
              <h2>Step 1: Extract PDFs & Build Budget Sheet</h2>
              <p style={{ color: 'var(--text-muted)', marginBottom: '1.25rem' }}>
                Inspects all PO PDFs in your selected folder, extracts PO numbers, products, crops, activities, and budget figures without duplicates.
              </p>
              <form onSubmit={handleStep1Submit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                
                <BrowseField 
                  label="Input FMC POs PDF Folder"
                  value={inputFolderPath}
                  onChange={(e) => setInputFolderPath(e.target.value)}
                  onBrowse={handleBrowseInputFolder}
                  browseLoading={browseInputLoading}
                  disabled={loading}
                />

                <BrowseField 
                  label="Save Folder Path"
                  value={saveFolderPath}
                  onChange={(e) => setSaveFolderPath(e.target.value)}
                  onBrowse={handleBrowseSaveFolder}
                  browseLoading={browseFolderLoading}
                  disabled={loading}
                />

                <button type="submit" className="primary" disabled={loading || !inputFolderPath || !saveFolderPath} style={{ marginTop: '0.5rem' }}>
                  {loading ? 'Processing FMC PDFs...' : ' Step 1: Build / Append Master Budget Sheet'}
                </button>
              </form>
            </>
          ) : (
            <>
              <h2>Step 2: Generate PO Summary Cards</h2>
              <p style={{ color: 'var(--text-muted)', marginBottom: '1.25rem' }}>
                Select the FMC Budget Excel file from Step 1. Creates formatted individual PO summary card sheets (11 cards per sheet).
              </p>
              <form onSubmit={handleStep2Submit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                
                <BrowseField 
                  label="Target FMC Budget Excel File (Step 1 Output)"
                  value={excelPath}
                  onChange={(e) => setExcelPath(e.target.value)}
                  onBrowse={handleBrowseExcel}
                  browseLoading={browseExcelLoading}
                  disabled={loading}
                />

                <button type="submit" className="primary" disabled={loading || !excelPath} style={{ marginTop: '0.5rem' }}>
                  {loading ? 'Generating Summary Cards...' : ' Step 2: Generate 11 PO Cards Per Sheet'}
                </button>
              </form>
            </>
          )}
        </div>

        {/* Results Panel Card */}
        <div className="card">
          <h2>Summary File Details</h2>
          
          {errorMsg && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ? {errorMsg}
            </div>
          )}

          {!loading && !result && !errorMsg && (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '5rem 1rem' }}>
              <h3>Awaiting Action</h3>
              <p style={{ marginTop: '0.5rem' }}>
                {activeStep === 'step1' 
                  ? 'Select the FMC PDF folder and click process to generate or append the master budget table.'
                  : 'Select your Step 1 Excel workbook and click process to generate formatted PO summary card sheets.'
                }
              </p>
            </div>
          )}

          {loading && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '5rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>??</div>
              <h3>{activeStep === 'step1' ? 'Processing PO PDFs...' : 'Building Summary Card Sheets...'}</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                {activeStep === 'step1' 
                  ? 'Reading PO numbers, products, crops, activities, and budget figures.'
                  : 'Constructing 11 PO summary blocks per sheet with exact styling, formulas, and headers.'
                }
              </p>
            </div>
          )}

          <ResultPanel result={result} isSummary={true} />

        </div>
      </div>
    </div>
  );
}

export default FmcSummaryView;
