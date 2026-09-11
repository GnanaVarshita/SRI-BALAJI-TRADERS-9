import React, { useState } from 'react';
import ResultPanel from '../common/ResultPanel';
import FmcBudgetStep1Card from '../fmc/FmcBudgetStep1Card';
import FmcCardsStep2Card from '../fmc/FmcCardsStep2Card';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';

function FmcSummaryView() {
  const [activeStep, setActiveStep] = useState('step1');

  // Step 1 states
  const [inputFolderPath, setInputFolderPath] = useState('');
  const [saveFolderPath, setSaveFolderPath] = useState('');

  // Shared states
  const [amName, setAmName] = useState('Madhavareddy');
  const amOptions = ['Madhavareddy', 'Venkateshwar reddy'];

  // Step 2 states
  const [excelPath, setExcelPath] = useState('');

  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const inputFolderBrowser = useFileBrowser();
  const saveFolderBrowser = useFileBrowser();
  const excelFileBrowser = useFileBrowser();

  // Step 1 Browsers
  const handleBrowseInputFolder = () => {
    setErrorMsg(null);
    setResult(null);
    inputFolderBrowser.browseFolder((folderPath) => {
      setInputFolderPath(folderPath);
      const parentFolder = folderPath.substring(
        0,
        folderPath.lastIndexOf(folderPath.includes('/') ? '/' : '\\')
      );
      setSaveFolderPath(parentFolder);
    });
  };

  const handleBrowseSaveFolder = () => {
    setErrorMsg(null);
    setResult(null);
    saveFolderBrowser.browseFolder((folderPath) => {
      setSaveFolderPath(folderPath);
    });
  };

  // Step 2 Browser
  const handleBrowseExcel = () => {
    setErrorMsg(null);
    setResult(null);
    excelFileBrowser.browseFile((filePath) => {
      setExcelPath(filePath);
    });
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
      const data = await api.generateFmcSummary({ inputFolderPath, saveFolderPath, amName });
      setResult(data);
      if (data.outputPath) {
        setExcelPath(data.outputPath);
      }
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to generate FMC Master Budget sheet.');
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
      const data = await api.generateFmcStep2({ excelPath, amName });
      setResult(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to generate PO Summary Cards.');
    } finally {
      setLoading(false);
    }
  };

  const activeError =
    errorMsg ||
    inputFolderBrowser.error ||
    saveFolderBrowser.error ||
    excelFileBrowser.error;

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>FMC PO Summary Generator</h2>
        <p className="subtitle">
          Step 1: Build master budget table from PO PDFs. Step 2: Generate 11 PO summary cards per sheet in your workbook.
        </p>
      </div>

      {/* Step Tabs */}
      <div style={{ display: 'flex', gap: '1rem', marginBottom: '1.5rem' }}>
        <button
          type="button"
          className={activeStep === 'step1' ? 'primary' : 'secondary'}
          onClick={() => setActiveStep('step1')}
          style={{ padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
          Step 1: Master Budget Table (From PDFs)
        </button>
        <button
          type="button"
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
            <FmcBudgetStep1Card
              amName={amName}
              setAmName={setAmName}
              amOptions={amOptions}
              inputFolderPath={inputFolderPath}
              setInputFolderPath={setInputFolderPath}
              saveFolderPath={saveFolderPath}
              setSaveFolderPath={setSaveFolderPath}
              loading={loading}
              onSubmit={handleStep1Submit}
              onBrowseInputFolder={handleBrowseInputFolder}
              onBrowseSaveFolder={handleBrowseSaveFolder}
              browseInputLoading={inputFolderBrowser.loading}
              browseFolderLoading={saveFolderBrowser.loading}
            />
          ) : (
            <FmcCardsStep2Card
              amName={amName}
              setAmName={setAmName}
              amOptions={amOptions}
              excelPath={excelPath}
              setExcelPath={setExcelPath}
              loading={loading}
              onSubmit={handleStep2Submit}
              onBrowseExcel={handleBrowseExcel}
              browseExcelLoading={excelFileBrowser.loading}
            />
          )}
        </div>

        {/* Execution Output Card */}
        <div className="card">
          <h2>Execution Status</h2>

          {activeError && (
            <div className="toast error" style={{ width: '100%', marginBottom: '1.5rem' }}>
              ⚠️ {activeError}
            </div>
          )}

          {loading ? (
            <div style={{ textAlign: 'center', color: 'var(--primary-color)', padding: '3rem 1rem' }}>
              <div className="spinner" style={{ fontSize: '2.5rem', display: 'inline-block', marginBottom: '1rem' }}>
                ⏳
              </div>
              <h3>Processing Request...</h3>
              <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem' }}>
                Please wait while PDF text extraction and Excel generation is in progress.
              </p>
            </div>
          ) : result ? (
            <ResultPanel result={result} isSummary={true} />
          ) : (
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', padding: '3rem 1rem' }}>
              <h3>Ready to Process</h3>
              <p style={{ marginTop: '0.5rem' }}>
                Fill out the required inputs on the left and click the submit button to execute.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default FmcSummaryView;
