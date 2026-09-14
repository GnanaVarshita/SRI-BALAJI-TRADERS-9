import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import SelectField from '../common/SelectField';
import FormField from '../common/FormField';
import ResultPanel from '../common/ResultPanel';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';

function Step1PoSummaryCard() {
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
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const fileBrowser = useFileBrowser();
  const folderBrowser = useFileBrowser();

  const handleBrowseInput = () => {
    setErrorMsg(null);
    setResult(null);
    fileBrowser.browseFile((filePath) => {
      setInputPath(filePath);
      
      const filename = filePath.split(/[\\/]/).pop();
      const nameWithoutExt = filename.replace(/\.[^/.]+$/, "");
      const parentFolder = filePath.substring(0, filePath.lastIndexOf(filePath.includes('/') ? '/' : '\\'));
      setSaveFolderPath(parentFolder);
      
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
    });
  };

  const handleBrowseFolder = () => {
    setErrorMsg(null);
    setResult(null);
    folderBrowser.browseFolder((folderPath) => {
      setSaveFolderPath(folderPath);
    });
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
      const data = await api.generateSummary({
        inputPath,
        saveFolderPath,
        outputName,
        poNumber,
        date,
        contact,
        territory
      });
      setResult(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to generate PO Summary. Verify sheet formats.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="card">
      <h3 style={{ marginTop: 0, marginBottom: '1.25rem' }}>Step 1: Generate Individual PO Summary Card</h3>
      <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
        <BrowseField 
          label="Input Quotation Excel File (.xlsx)"
          value={inputPath}
          onChange={(e) => setInputPath(e.target.value)}
          onBrowse={handleBrowseInput}
          browseLoading={fileBrowser.loading}
          disabled={loading}
          placeholder="Select an Excel file containing Quotation 1, Quotation 2 sheets..."
        />

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <BrowseField 
            label="Save Destination Folder"
            value={saveFolderPath}
            onChange={(e) => setSaveFolderPath(e.target.value)}
            onBrowse={handleBrowseFolder}
            browseLoading={folderBrowser.loading}
            disabled={loading}
            placeholder="Select folder to save summary..."
          />
          <FormField 
            label="Output Excel File Name"
            type="text"
            value={outputName}
            onChange={(e) => setOutputName(e.target.value)}
            disabled={loading}
            placeholder="e.g. Nellore PO Summary.xlsx"
            required
          />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <FormField 
            label="PO Number (Optional)"
            type="text"
            value={poNumber}
            onChange={(e) => setPoNumber(e.target.value)}
            disabled={loading}
            placeholder="e.g. 4800108503"
          />
          <FormField 
            label="Date (DD-MM-YYYY)"
            type="text"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            disabled={loading}
            required
          />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <SelectField 
            label="Territory"
            value={territory}
            onChange={(e) => {
              const val = e.target.value;
              setTerritory(val);
              setOutputName(`${val} PO Summary.xlsx`);
            }}
            options={[
              { value: 'Nellore', label: 'Nellore' },
              { value: 'Kurnool', label: 'Kurnool' },
              { value: 'Suryapet', label: 'Suryapet' }
            ]}
            disabled={loading}
          />
          <FormField 
            label="Contact Person Name"
            type="text"
            value={contact}
            onChange={(e) => setContact(e.target.value)}
            disabled={loading}
            placeholder="e.g. K.Subbaramireddy"
            required
          />
        </div>

        {(errorMsg || fileBrowser.error || folderBrowser.error) && (
          <div style={{ padding: '0.75rem', background: 'rgba(231, 76, 60, 0.1)', border: '1px solid #e74c3c', borderRadius: '4px', color: '#e74c3c' }}>
            {errorMsg || fileBrowser.error || folderBrowser.error}
          </div>
        )}

        <button 
          type="submit" 
          className="primary" 
          disabled={loading}
          style={{ alignSelf: 'flex-start', minWidth: '220px', padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
          {loading ? 'Processing...' : 'Generate PO Summary Sheet'}
        </button>
      </form>

      {result && (
        <ResultPanel 
          title="PO Summary Generated Successfully!"
          message={result.message}
          outputPath={result.outputPath}
        />
      )}
    </div>
  );
}

export default Step1PoSummaryCard;
