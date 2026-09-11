import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import SelectField from '../common/SelectField';
import FormField from '../common/FormField';
import ResultPanel from '../common/ResultPanel';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';

function Step2CortevaMasterCard() {
  const [step2FolderPath, setStep2FolderPath] = useState('');
  const [step2OutputName, setStep2OutputName] = useState('Kurnool Master PO Summary.xlsx');
  const [step2Territory, setStep2Territory] = useState('Kurnool');
  const [step2Zdgm, setStep2Zdgm] = useState('Bhaskar');
  const [step2BudgetSeason, setStep2BudgetSeason] = useState('Kharif');
  
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const folderBrowser = useFileBrowser();

  const handleBrowseFolder = () => {
    setErrorMsg(null);
    setResult(null);
    folderBrowser.browseFolder((folderPath) => {
      setStep2FolderPath(folderPath);
      const folderName = folderPath.split(/[\\/]/).pop();
      const match = folderName.match(/^([a-zA-Z\s]+)/);
      if (match) {
        const firstWord = match[1].trim().split(/\s+/)[0];
        const capitalized = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();
        if (['Nellore', 'Kurnool', 'Suryapet'].includes(capitalized)) {
          setStep2Territory(capitalized);
        }
        setStep2OutputName(`${capitalized} Master PO Summary.xlsx`);
      }
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!step2FolderPath) {
      setErrorMsg('Please select the folder containing PO summary card files.');
      return;
    }
    if (!step2OutputName) {
      setErrorMsg('Please enter an output file name.');
      return;
    }

    setLoading(true);
    setResult(null);
    setErrorMsg(null);

    try {
      const data = await api.generateCortevaMasterSummary({
        folderPath: step2FolderPath,
        outputName: step2OutputName,
        territory: step2Territory,
        zdgm: step2Zdgm,
        budgetSeason: step2BudgetSeason
      });
      setResult(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to generate Master PO Summary.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="card">
      <h3 style={{ marginTop: 0, marginBottom: '1.25rem' }}>Step 2: Consolidate Master PO Summary Cards</h3>
      <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
        <BrowseField 
          label="Input Directory Containing PO Summary Files"
          value={step2FolderPath}
          onChange={(e) => setStep2FolderPath(e.target.value)}
          onBrowse={handleBrowseFolder}
          browseLoading={folderBrowser.loading}
          disabled={loading}
          placeholder="Select folder containing individual PO Summary cards..."
        />

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <FormField 
            label="Output Master File Name"
            type="text"
            value={step2OutputName}
            onChange={(e) => setStep2OutputName(e.target.value)}
            disabled={loading}
            placeholder="e.g. Kurnool Master PO Summary.xlsx"
            required
          />
          <SelectField 
            label="Territory"
            value={step2Territory}
            onChange={(e) => {
              const val = e.target.value;
              setStep2Territory(val);
              setStep2OutputName(`${val} Master PO Summary.xlsx`);
            }}
            options={[
              { value: 'Kurnool', label: 'Kurnool' },
              { value: 'Nellore', label: 'Nellore' },
              { value: 'Suryapet', label: 'Suryapet' }
            ]}
            disabled={loading}
          />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
          <FormField 
            label="ZDGM Name"
            type="text"
            value={step2Zdgm}
            onChange={(e) => setStep2Zdgm(e.target.value)}
            disabled={loading}
            placeholder="e.g. Bhaskar"
          />
          <FormField 
            label="Budget Season"
            type="text"
            value={step2BudgetSeason}
            onChange={(e) => setStep2BudgetSeason(e.target.value)}
            disabled={loading}
            placeholder="e.g. Kharif"
          />
        </div>

        {(errorMsg || folderBrowser.error) && (
          <div style={{ padding: '0.75rem', background: 'rgba(231, 76, 60, 0.1)', border: '1px solid #e74c3c', borderRadius: '4px', color: '#e74c3c' }}>
            {errorMsg || folderBrowser.error}
          </div>
        )}

        <button 
          type="submit" 
          className="primary" 
          disabled={loading}
          style={{ alignSelf: 'flex-start', minWidth: '240px', padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
          {loading ? 'Consolidating...' : 'Generate Master PO Summary'}
        </button>
      </form>

      {result && (
        <ResultPanel 
          title="Master PO Summary Generated Successfully!"
          message={result.message}
          outputPath={result.outputPath}
        />
      )}
    </div>
  );
}

export default Step2CortevaMasterCard;
