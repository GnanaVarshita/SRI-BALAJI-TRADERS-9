import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import SelectField from '../common/SelectField';
import FormField from '../common/FormField';
import ResultPanel from '../common/ResultPanel';

function QuotationView() {
  const [filePath, setFilePath] = useState('');
  const [company, setCompany] = useState('Corteva Agriscience');
  const [contact, setContact] = useState('K.Subbaramireddy');
  const [designation, setDesignation] = useState('ZDGM');
  const [territory, setTerritory] = useState('Nellore');
  
  const getFormattedDate = () => {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  };
  const [date, setDate] = useState(getFormattedDate());
  
  const [loading, setLoading] = useState(false);
  const [browseLoading, setBrowseLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleBrowseFile = async () => {
    setBrowseLoading(true);
    setErrorMsg(null);
    setResult(null);
    try {
      const res = await fetch('/api/browse-file', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.success && data.filePath) {
        setFilePath(data.filePath);
        
        // Parse territory/area dynamically from the selected filename
        const filename = data.filePath.split(/[\\/]/).pop();
        const nameWithoutExt = filename.replace(/\.[^/.]+$/, "");
        const match = nameWithoutExt.match(/^([a-zA-Z\s]+)/);
        if (match) {
          const parts = match[1].trim().split(/\s+/);
          if (parts.length > 0) {
            const firstWord = parts[0];
            const capitalized = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();
            if (['Nellore', 'Kurnool', 'Suryapet'].includes(capitalized)) {
              setTerritory(capitalized);
            }
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

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!filePath) {
      setErrorMsg('Please select a local Excel file first.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const res = await fetch('/api/process-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filePath,
          company,
          contact,
          designation,
          territory,
          date
        }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        setResult(data);
      } else {
        setErrorMsg(data.message || 'Validation failed. Please verify the Excel sheet structure.');
      }
    } catch (err) {
      console.error(err);
      setErrorMsg('Network error: Failed to process local Excel file.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Local Excel Quotation Generator</h2>
        <p className="subtitle">Select a local spreadsheet, parse details, and append product quotation worksheets directly in-place.</p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '2rem' }}>
        {/* Settings Form Card */}
        <div className="card">
          <h2>Select File & Client Details</h2>
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            
            <BrowseField 
              label="Excel File Path"
              value={filePath}
              onChange={(e) => setFilePath(e.target.value)}
              onBrowse={handleBrowseFile}
              browseLoading={browseLoading}
              disabled={loading}
            />

            <SelectField 
              label="Company / Client Name"
              value={company}
              onChange={(e) => setCompany(e.target.value)}
              options={["Corteva Agriscience", "New Gen (FMC related)"]}
            />

            <SelectField 
              label="Contact Person (To Block)"
              value={contact}
              onChange={(e) => setContact(e.target.value)}
              options={["K.Subbaramireddy", "R.Bhaskar", "Roopsingh K"]}
            />

            <FormField 
              label="Designation"
              value={designation}
              onChange={(e) => setDesignation(e.target.value)}
              placeholder="e.g. ZDGM"
              required
            />

            <SelectField 
              label="Territory / Area"
              value={territory}
              onChange={(e) => setTerritory(e.target.value)}
              options={["Nellore", "Kurnool", "Suryapet"]}
            />

            <FormField 
              label="Quotation Date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              placeholder="DD-MM-YYYY"
              required
            />

            <button type="submit" className="primary" disabled={loading || !filePath} style={{ marginTop: '0.5rem' }}>
              {loading ? '⚙️ Modifying Excel File...' : '⚡ Generate Quotations In-Place'}
            </button>
          </form>
        </div>

        {/* Results Panel Card */}
        <div className="card">
          <h2>Generation Status & Results</h2>
          
          {errorMsg && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ❌ {errorMsg}
            </div>
          )}

          {!loading && !result && !errorMsg && (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '4rem 1rem' }}>
              <h3>Awaiting Excel Selection</h3>
              <p style={{ marginTop: '0.5rem' }}>Browse your local files, verify client details, and click process to update the sheet in-place.</p>
            </div>
          )}

          {loading && (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '5rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '3rem', display: 'inline-block', marginBottom: '1rem' }}>🔄</div>
              <h3>Modifying spreadsheet in-place...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>Creating product quotation tabs directly in your file.</p>
            </div>
          )}

          <ResultPanel result={result} filePath={filePath} />

        </div>
      </div>
    </div>
  );
}

export default QuotationView;
