import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import SelectField from '../common/SelectField';
import FormField from '../common/FormField';
import ResultPanel from '../common/ResultPanel';

function SummaryView() {
  const [inputPath, setInputPath] = useState('');
  const [saveFolderPath, setSaveFolderPath] = useState('');
  const [outputName, setOutputName] = useState('Nellore PO Summary.xlsx');
  const [poNumber, setPoNumber] = useState('');
  
  const getFormattedDate = () => {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  };
  const [date, setDate] = useState(getFormattedDate());
  
  const [contact, setContact] = useState('K.Subbaramireddy');
  const [territory, setTerritory] = useState('Nellore');
  
  const [loading, setLoading] = useState(false);
  const [browseLoading, setBrowseLoading] = useState(false);
  const [browseFolderLoading, setBrowseFolderLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseInput = async () => {
    setBrowseLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setInputPath(data.filePath);
        
        // Auto-extract territory, save folder, and update output name
        const filename = data.filePath.split(/[\\/]/).pop();
        const nameWithoutExt = filename.replace(/\.[^/.]+$/, "");
        
        // Also pre-fill the save folder with the input file's parent folder!
        const parentFolder = data.filePath.substring(0, data.filePath.lastIndexOf(data.filePath.includes('/') ? '/' : '\\'));
        setSaveFolderPath(parentFolder);
        
        // Try to parse territory
        const match = nameWithoutExt.match(/^([a-zA-Z\s]+)/);
        if (match) {
          const parts = match[1].trim().split(/\s+/);
          if (parts.length > 0) {
            const firstWord = parts[0];
            const capitalized = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();
            
            if (['Nellore', 'Kurnool', 'Suryapet'].includes(capitalized)) {
              setTerritory(capitalized);
            }
            setOutputName(`${capitalized} PO Summary.xlsx`);
          }
        }
      } else if (data.message) {
        setErrorMsg(data.message);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Failed to open file browser dialog.');
    } finally {
      setBrowseLoading(false);
    }
  };

  const handleBrowseFolder = async () => {
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
    if (!inputPath) {
      setErrorMsg('Please select the input Quotation Excel file first.');
      return;
    }
    if (!saveFolderPath) {
      setErrorMsg('Please select a folder to save the file.');
      return;
    }
    if (!outputName) {
      setErrorMsg('Please enter an output file name.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/generate-summary', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          inputPath,
          saveFolderPath,
          outputName,
          poNumber,
          date,
          contact,
          territory
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Failed to generate PO Summary. Verify sheet formats.');
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
        <h2>Corteva PO Summary Generator</h2>
        <p className="subtitle">Convert a quotation excel file into a tracking PO Summary excel sheet.</p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '2rem' }}>
        {/* Settings Form Card */}
        <div className="card">
          <h2>Summary Configuration</h2>
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            
            <BrowseField 
              label="Input Quotation Excel"
              value={inputPath}
              onChange={(e) => setInputPath(e.target.value)}
              onBrowse={handleBrowseInput}
              browseLoading={browseLoading}
              disabled={loading}
            />

            <BrowseField 
              label="Save Folder Path"
              value={saveFolderPath}
              onChange={(e) => setSaveFolderPath(e.target.value)}
              onBrowse={handleBrowseFolder}
              browseLoading={browseFolderLoading}
              disabled={loading}
            />

            <FormField 
              label="Output File Name"
              value={outputName}
              onChange={(e) => setOutputName(e.target.value)}
              placeholder="e.g. Nellore PO Summary.xlsx"
              required
            />

            <FormField 
              label="PO Number (Optional)"
              value={poNumber}
              onChange={(e) => setPoNumber(e.target.value)}
              placeholder="e.g. 4800108504"
            />

            <FormField 
              label="Date (Optional)"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              placeholder="DD-MM-YYYY"
            />

            <SelectField 
              label="Contact Person (To Block)"
              value={contact}
              onChange={(e) => setContact(e.target.value)}
              options={["K.Subbaramireddy", "R.Bhaskar", "Roopsingh K"]}
            />

            <SelectField 
              label="Territory / Area"
              value={territory}
              onChange={(e) => setTerritory(e.target.value)}
              options={["Nellore", "Kurnool", "Suryapet"]}
            />

            <button type="submit" className="primary" disabled={loading || !inputPath || !saveFolderPath} style={{ marginTop: '0.5rem' }}>
              {loading ? '⚙️ Creating PO Summary File...' : '⚡ Generate New PO Summary Workbook'}
            </button>
          </form>
        </div>

        {/* Results Panel Card */}
        <div className="card">
          <h2>Summary File Details</h2>
          
          {errorMsg && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ❌ {errorMsg}
            </div>
          )}

          {!loading && !result && !errorMsg && (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '5rem 1rem' }}>
              <h3>Awaiting Configuration</h3>
              <p style={{ marginTop: '0.5rem' }}>Select the quotation sheet, choose a target save folder on your PC, and click process to generate.</p>
            </div>
          )}

          {loading && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '5rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>🔄</div>
              <h3>Generating brand-new summary...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>Structuring product sheets, rates grids, and balance formulas.</p>
            </div>
          )}

          <ResultPanel result={result} isSummary={true} />

        </div>
      </div>
    </div>
  );
}

export default SummaryView;
