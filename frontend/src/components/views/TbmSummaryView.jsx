import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function TbmSummaryView() {
  const [tbmFolderPath, setTbmFolderPath] = useState('');
  const [outputPath, setOutputPath] = useState('');
  const [priorityPoList, setPriorityPoList] = useState('');

  const [loading, setLoading] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseTbmFolder = async () => {
    setBrowseFolderLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-folder', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.folderPath) {
        setTbmFolderPath(data.folderPath);
        // Auto derive output filename inside selected folder
        const parts = data.folderPath.split(/[/\\]/);
        let territory = parts[parts.length - 1] || 'All-TBMs';
        if (territory.toLowerCase() === 'tbm s summary' && parts.length > 1) {
          territory = parts[parts.length - 3] || parts[parts.length - 2] || 'All-TBMs';
        }
        const defaultName = `${territory}-All-TBMs-Summary.xlsx`.replace(/\s+/g, '-');
        setOutputPath(`${data.folderPath}\\${defaultName}`);
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
    if (!tbmFolderPath) {
      setErrorMsg('Please select the TBM s Summary folder path first.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/generate-tbm-summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tbmFolderPath,
          outputPath: outputPath.trim(),
          priorityPoList: priorityPoList.trim()
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to consolidate TBM Summary files.');
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
        <h2>TBM Master Summary Consolidation</h2>
        <p className="subtitle">
          Brings all TBM activity Excel sheets into one master workbook. Automatically groups by PO Number, Product &amp; Activity, and TBM Name (Max 9 tables per sheet tab).
        </p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.4fr 1fr', gap: '2rem' }}>
        {/* Input Form Card */}
        <div className="card">
          <h2>Consolidation Settings</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '1.25rem' }}>
            Select the folder containing TBM subfolders (e.g. <code>TBM s Summary</code>). Output master summary file will be created/appended inside that folder.
          </p>

          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <BrowseField 
              label="TBM s Summary Folder Path"
              value={tbmFolderPath}
              onChange={(e) => setTbmFolderPath(e.target.value)}
              onBrowse={handleBrowseTbmFolder}
              browseLoading={browseFolderLoading}
              disabled={loading}
              placeholder="e.g. D:\SRI BALAJI TRADERS\CORTEVA\KURNOOL\2026-2027\TBM s Summary"
            />

            <div className="form-field">
              <label className="form-label">
                Target Output Master Excel File Path
              </label>
              <textarea
                rows={1}
                className="path-input-textarea"
                value={outputPath}
                onChange={(e) => setOutputPath(e.target.value)}
                placeholder="Defaults to [Territory]-All-TBMs-Summary.xlsx inside TBM Summary Folder"
                disabled={loading}
                style={{
                  minHeight: '42px',
                  resize: 'vertical',
                  wordBreak: 'break-all',
                  overflowWrap: 'anywhere',
                  whiteSpace: 'pre-wrap',
                  fontFamily: 'Consolas, "Courier New", monospace, sans-serif',
                  fontSize: '0.88rem',
                  lineHeight: '1.4',
                  padding: '0.65rem 0.75rem'
                }}
              />
              <span className="form-help">
                If the Excel file exists, new activities will be appended automatically.
              </span>
            </div>

            <div className="form-field">
              <label className="form-label">
                Priority PO Numbers List (Optional)
              </label>
              <textarea
                className="input-text"
                style={{ height: '80px', fontFamily: 'monospace', resize: 'vertical' }}
                value={priorityPoList}
                onChange={(e) => setPriorityPoList(e.target.value)}
                placeholder="Enter or paste priority PO numbers (e.g. 500BB20260710377, 500BB20260710177). Unlisted POs will go to a separate sheet tab."
                disabled={loading}
              />
              <span className="form-help">
                POs in this list will be placed in main sheets (Sheet1, Sheet2...). Unlisted POs go to Unlisted POs sheet, missing POs go to No PO sheet.
              </span>
            </div>

            <button
              type="submit"
              className="primary"
              disabled={loading || !tbmFolderPath}
              style={{ marginTop: '0.5rem', padding: '0.85rem' }}
            >
              {loading ? 'Consolidating TBM Summaries...' : '📊 Consolidate & Append TBM Summaries'}
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

          {!loading && !result && !errorMsg && (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
              <h3>Ready for Consolidation</h3>
              <p style={{ marginTop: '0.5rem' }}>
                Select the TBM s Summary folder and click <strong>Consolidate</strong> to aggregate all activity tables into a master Excel file.
              </p>
            </div>
          )}

          {loading && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
              <h3>Processing TBM Excel Files...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Parsing activity rows, formatting tables with TBM headings, and applying max 9 tables per sheet rule.
              </p>
            </div>
          )}

          {result && (
            <div>
              <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1.25rem' }}>
                <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Consolidation Complete</h4>
                <p style={{ fontSize: '0.9rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
                  {result.message}
                </p>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem', marginBottom: '1.25rem' }}>
                <div style={{ background: 'var(--bg-hover)', padding: '0.75rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Total Activities</span>
                  <div style={{ fontSize: '1.25rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{result.totalActivities}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.75rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Total Tables</span>
                  <div style={{ fontSize: '1.25rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{result.totalTables}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.75rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Sheets Created</span>
                  <div style={{ fontSize: '1.25rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{result.sheetsCount}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.75rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Max Tables / Sheet</span>
                  <div style={{ fontSize: '1.25rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>9</div>
                </div>
              </div>

              <ResultPanel result={result} isSummary={true} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default TbmSummaryView;
