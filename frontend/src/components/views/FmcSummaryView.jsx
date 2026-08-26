import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function FmcSummaryView() {
  const [inputFolderPath, setInputFolderPath] = useState('');
  const [saveFolderPath, setSaveFolderPath] = useState('');

  const [loading, setLoading] = useState(false);
  const [browseInputLoading, setBrowseInputLoading] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseInputFolder = async () => {
    setBrowseInputLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-folder', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.folderPath) {
        setInputFolderPath(data.folderPath);

        // Pre-fill save folder with parent directory
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

  const handleSubmit = async (e) => {
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
        body: JSON.stringify({
          inputFolderPath,
          saveFolderPath
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to generate FMC Summary.');
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
        <p className="subtitle">Inspect FMC PO PDFs from local folders and generate/append to a tracking budget summary sheet.</p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '2rem' }}>
        {/* Settings Form Card */}
        <div className="card">
          <h2>FMC Summary Configuration</h2>
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            
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
              {loading ? '?? Processing FMC PDFs...' : '? Generate / Append FMC PO Summary Workbook'}
            </button>
          </form>
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
              <h3>Awaiting Configuration</h3>
              <p style={{ marginTop: '0.5rem' }}>Select the FMC PDF folder, choose a target save folder on your PC, and click process to generate or append.</p>
            </div>
          )}

          {loading && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '5rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>??</div>
              <h3>Extracting PO PDFs & Updating Summary...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>Reading PO numbers, products, crops, activities, and budget figures without duplicating existing entries.</p>
            </div>
          )}

          <ResultPanel result={result} isSummary={true} />

        </div>
      </div>
    </div>
  );
}

export default FmcSummaryView;
