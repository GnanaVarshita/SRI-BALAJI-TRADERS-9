import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import ResultPanel from '../common/ResultPanel';

function TbmSummaryView() {
  const [tbmFolderPath, setTbmFolderPath] = useState('');
  const [outputPath, setOutputPath] = useState('');
  const [priorityPoList, setPriorityPoList] = useState('');

  const [loadingStep1, setLoadingStep1] = useState(false);
  const [loadingStep2, setLoadingStep2] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  
  const [step1Result, setStep1Result] = useState(null);
  const [step2Result, setStep2Result] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseTbmFolder = async () => {
    setBrowseFolderLoading(true);
    setErrorMsg(null);
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

  const handleFormatSheets = async () => {
    if (!tbmFolderPath) {
      setErrorMsg('Please select the TBM s Summary folder path first.');
      return;
    }

    setLoadingStep1(true);
    setStep1Result(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/format-tbm-summaries', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tbmFolderPath: tbmFolderPath.trim(),
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setStep1Result(data);
      } else {
        setErrorMsg(data.message || 'Failed to format TBM Summary files.');
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Network error: Failed to contact the backend server.');
    } finally {
      setLoadingStep1(false);
    }
  };

  const handleConsolidate = async (e) => {
    if (e) e.preventDefault();
    if (!tbmFolderPath) {
      setErrorMsg('Please select the TBM s Summary folder path first.');
      return;
    }

    setLoadingStep2(true);
    setStep2Result(null);
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
        setStep2Result(data);
      } else {
        setErrorMsg(data.message || 'Failed to consolidate TBM Summary files.');
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Network error: Failed to contact the backend server.');
    } finally {
      setLoadingStep2(false);
    }
  };

  const isLoading = loadingStep1 || loadingStep2;

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>TBM Summary Formatting &amp; Master Consolidation</h2>
        <p className="subtitle">
          Two-step workflow: Format individual TBM summary sheets on Sheet 2 (with green headers, date normalization, and PO &rarr; Product/Activity grouping), then consolidate all into the Master Summary workbook.
        </p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.35fr 1fr', gap: '2rem' }}>
        {/* Input Form Card */}
        <div className="card">
          <h2>TBM Summary Workflow</h2>
          <p style={{ color: 'var(--text-muted)', marginBottom: '1.25rem' }}>
            Select the folder containing TBM subfolders (e.g. <code>TBM s Summary</code>).
          </p>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <BrowseField 
              label="TBM s Summary Folder Path"
              value={tbmFolderPath}
              onChange={(e) => setTbmFolderPath(e.target.value)}
              onBrowse={handleBrowseTbmFolder}
              browseLoading={browseFolderLoading}
              disabled={isLoading}
              placeholder="e.g. D:\SRI BALAJI TRADERS\CORTEVA\KURNOOL\2026-2027\TBM s Summary"
            />

            {/* STEP 1 CARD */}
            <div style={{
              border: '1px solid var(--border-color)',
              background: 'var(--bg-hover)',
              padding: '1.15rem',
              borderRadius: '8px',
              display: 'flex',
              flexDirection: 'column',
              gap: '0.65rem'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--primary-color)' }}>
                  Step 1: Format TBM Sheets (Second Sheet)
                </h3>
                <span style={{ fontSize: '0.78rem', background: 'var(--primary-color)', color: '#000', padding: '0.15rem 0.5rem', borderRadius: '4px', fontWeight: 'bold' }}>
                  Step 1
                </span>
              </div>
              <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
                Reads raw summaries from Sheet 1, groups by (PO, Product, Activity), formats dates into <code>DD-MM-YYYY</code>, adds Excel formulas, and writes styled tables to <strong>Sheet 2</strong> without touching Sheet 1.
              </p>
              <button
                type="button"
                className="secondary"
                onClick={handleFormatSheets}
                disabled={isLoading || !tbmFolderPath}
                style={{ marginTop: '0.35rem', padding: '0.75rem', fontWeight: '600' }}
              >
                {loadingStep1 ? '⏳ Formatting TBM Summaries...' : '✨ Step 1: Format All TBM Sheets in-place (Sheet 2)'}
              </button>
            </div>

            {/* STEP 2 CARD */}
            <div style={{
              border: '1px solid var(--border-color)',
              background: 'var(--bg-card)',
              padding: '1.15rem',
              borderRadius: '8px',
              display: 'flex',
              flexDirection: 'column',
              gap: '0.85rem'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--primary-color)' }}>
                  Step 2: Consolidate Master TBM Summary
                </h3>
                <span style={{ fontSize: '0.78rem', background: 'var(--primary-color)', color: '#000', padding: '0.15rem 0.5rem', borderRadius: '4px', fontWeight: 'bold' }}>
                  Step 2
                </span>
              </div>

              <div className="form-field">
                <label className="form-label" style={{ fontSize: '0.85rem' }}>
                  Target Output Master Excel File Path
                </label>
                <textarea
                  rows={1}
                  className="path-input-textarea"
                  value={outputPath}
                  onChange={(e) => setOutputPath(e.target.value)}
                  placeholder="Defaults to [Territory]-All-TBMs-Summary.xlsx inside TBM Summary Folder"
                  disabled={isLoading}
                  style={{
                    minHeight: '38px',
                    resize: 'vertical',
                    wordBreak: 'break-all',
                    overflowWrap: 'anywhere',
                    whiteSpace: 'pre-wrap',
                    fontFamily: 'Consolas, "Courier New", monospace, sans-serif',
                    fontSize: '0.82rem',
                    lineHeight: '1.4',
                    padding: '0.55rem 0.65rem'
                  }}
                />
              </div>

              <div className="form-field">
                <label className="form-label" style={{ fontSize: '0.85rem' }}>
                  Priority PO Numbers List (Optional)
                </label>
                <textarea
                  className="input-text"
                  style={{ height: '65px', fontFamily: 'monospace', resize: 'vertical', fontSize: '0.82rem' }}
                  value={priorityPoList}
                  onChange={(e) => setPriorityPoList(e.target.value)}
                  placeholder="Enter priority PO numbers (e.g. 500BB20260710377, 500BB20260710177). Unlisted POs go to Unlisted POs sheet."
                  disabled={isLoading}
                />
              </div>

              <button
                type="button"
                className="primary"
                onClick={handleConsolidate}
                disabled={isLoading || !tbmFolderPath}
                style={{ marginTop: '0.25rem', padding: '0.85rem' }}
              >
                {loadingStep2 ? '⏳ Consolidating TBM Summaries...' : '📊 Step 2: Consolidate & Append Master Summary'}
              </button>
            </div>
          </div>
        </div>

        {/* Results Panel Card */}
        <div className="card">
          <h2>Execution Status</h2>
          
          {errorMsg && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ⚠️ {errorMsg}
            </div>
          )}

          {!isLoading && !step1Result && !step2Result && !errorMsg && (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
              <h3>Ready for Processing</h3>
              <p style={{ marginTop: '0.5rem' }}>
                Select the TBM s Summary folder, run <strong>Step 1</strong> to format the Excel sheets on Sheet 2, then run <strong>Step 2</strong> to build the Master Consolidated Summary.
              </p>
            </div>
          )}

          {loadingStep1 && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
              <h3>Formatting TBM Excel Sheets (Step 1)...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Extracting raw tables, normalizing dates to DD-MM-YYYY, grouping by (PO, Product, Activity), and writing green styled tables to Sheet 2.
              </p>
            </div>
          )}

          {loadingStep2 && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '4rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>⏳</div>
              <h3>Consolidating Master Summary (Step 2)...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Aggregating formatted tables across all TBMs into Master Summary with PO Grand Totals and TBM Amount Summary sheet.
              </p>
            </div>
          )}

          {/* STEP 1 RESULT */}
          {step1Result && !loadingStep1 && (
            <div style={{ marginBottom: '1.5rem' }}>
              <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
                <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Step 1: Formatting Complete</h4>
                <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
                  {step1Result.message}
                </p>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Files Processed</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {step1Result.processedFiles} / {step1Result.totalFiles}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Activities</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {step1Result.totalActivities}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Tables Created</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    {step1Result.totalTables}
                  </div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Destination</span>
                  <div style={{ fontSize: '0.9rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>
                    Sheet 2 (In-Place)
                  </div>
                </div>
              </div>

              {step1Result.details && step1Result.details.length > 0 && (
                <div style={{
                  maxHeight: '160px',
                  overflowY: 'auto',
                  background: 'var(--bg-card)',
                  border: '1px solid var(--border-color)',
                  borderRadius: '6px',
                  padding: '0.5rem',
                  fontSize: '0.78rem'
                }}>
                  {step1Result.details.map((d, idx) => (
                    <div key={idx} style={{ padding: '0.25rem 0', borderBottom: idx < step1Result.details.length - 1 ? '1px solid var(--border-color)' : 'none', color: d.success ? 'var(--text-color)' : '#e74c3c' }}>
                      {d.success ? '✓' : '⚠️'} <strong>{d.file}</strong>: {d.message}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* STEP 2 RESULT */}
          {step2Result && !loadingStep2 && (
            <div>
              <div style={{ backgroundColor: 'rgba(39, 174, 96, 0.15)', border: '1px solid #27ae60', padding: '1rem', borderRadius: '8px', marginBottom: '1rem' }}>
                <h4 style={{ color: '#27ae60', margin: 0 }}>✓ Step 2: Consolidation Complete</h4>
                <p style={{ fontSize: '0.88rem', marginTop: '0.4rem', color: 'var(--text-color)' }}>
                  {step2Result.message}
                </p>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.65rem', marginBottom: '1rem' }}>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Activities</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.totalActivities}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Total Tables</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.totalTables}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Sheets Created</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>{step2Result.sheetsCount}</div>
                </div>
                <div style={{ background: 'var(--bg-hover)', padding: '0.65rem', borderRadius: '6px' }}>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)' }}>Max Tables / Sheet</span>
                  <div style={{ fontSize: '1.15rem', fontWeight: 'bold', color: 'var(--primary-color)' }}>9</div>
                </div>
              </div>

              <ResultPanel result={step2Result} isSummary={true} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default TbmSummaryView;
