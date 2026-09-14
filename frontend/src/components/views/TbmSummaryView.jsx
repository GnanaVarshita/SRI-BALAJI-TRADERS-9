import React, { useState } from 'react';
import BrowseField from '../common/BrowseField';
import TbmFormatSection from '../tbm/TbmFormatSection';
import TbmConsolidateSection from '../tbm/TbmConsolidateSection';
import TbmStatusPanel from '../tbm/TbmStatusPanel';
import { api } from '../../services/api';
import { useFileBrowser } from '../../hooks/useFileBrowser';

function TbmSummaryView() {
  const [tbmFolderPath, setTbmFolderPath] = useState('');
  const [outputPath, setOutputPath] = useState('');
  const [priorityPoList, setPriorityPoList] = useState('');
  const [forceReformat, setForceReformat] = useState(false);
  const [forceConsolidate, setForceConsolidate] = useState(false);

  const [loadingStep1, setLoadingStep1] = useState(false);
  const [loadingStep2, setLoadingStep2] = useState(false);
  const [step1Result, setStep1Result] = useState(null);
  const [step2Result, setStep2Result] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);

  const folderBrowser = useFileBrowser();
  const fileBrowser = useFileBrowser();

  const handleBrowseTbmFolder = () => {
    setErrorMsg(null);
    folderBrowser.browseFolder((folderPath) => {
      setTbmFolderPath(folderPath);
      // Auto derive output filename inside selected folder
      const parts = folderPath.split(/[/\\]/);
      let territory = parts[parts.length - 1] || 'All-TBMs';
      if (territory.toLowerCase() === 'tbm s summary' && parts.length > 1) {
        territory = parts[parts.length - 3] || parts[parts.length - 2] || 'All-TBMs';
      }
      const defaultName = `${territory}-All-TBMs-Summary.xlsx`.replace(/\s+/g, '-');
      setOutputPath(`${folderPath}\\${defaultName}`);
    });
  };

  const handleBrowseOutputFile = () => {
    setErrorMsg(null);
    fileBrowser.browseFile((filePath) => {
      setOutputPath(filePath);
    });
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
      const data = await api.formatTbmSummaries({
        tbmFolderPath: tbmFolderPath.trim(),
        force: forceReformat,
      });
      setStep1Result(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to format TBM Summary files.');
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
      const data = await api.generateTbmSummary({
        tbmFolderPath: tbmFolderPath.trim(),
        outputPath: outputPath.trim(),
        priorityPoList: priorityPoList.trim(),
        force: forceConsolidate,
      });
      setStep2Result(data);
    } catch (err) {
      console.error(err);
      setErrorMsg(err.message || 'Failed to consolidate TBM Summary files.');
    } finally {
      setLoadingStep2(false);
    }
  };

  const isLoading = loadingStep1 || loadingStep2;
  const activeError = errorMsg || folderBrowser.error || fileBrowser.error;

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
              browseLoading={folderBrowser.loading}
              disabled={isLoading}
              placeholder="e.g. D:\SRI BALAJI TRADERS\CORTEVA\KURNOOL\2026-2027\TBM s Summary"
            />

            <TbmFormatSection
              tbmFolderPath={tbmFolderPath}
              isLoading={isLoading}
              loadingStep1={loadingStep1}
              forceReformat={forceReformat}
              setForceReformat={setForceReformat}
              onFormatSheets={handleFormatSheets}
            />

            <TbmConsolidateSection
              tbmFolderPath={tbmFolderPath}
              outputPath={outputPath}
              setOutputPath={setOutputPath}
              onBrowseOutputFile={handleBrowseOutputFile}
              browseFileLoading={fileBrowser.loading}
              priorityPoList={priorityPoList}
              setPriorityPoList={setPriorityPoList}
              forceConsolidate={forceConsolidate}
              setForceConsolidate={setForceConsolidate}
              isLoading={isLoading}
              loadingStep2={loadingStep2}
              onConsolidate={handleConsolidate}
            />
          </div>
        </div>

        {/* Results Panel Card */}
        <TbmStatusPanel
          isLoading={isLoading}
          loadingStep1={loadingStep1}
          loadingStep2={loadingStep2}
          step1Result={step1Result}
          step2Result={step2Result}
          errorMsg={activeError}
        />
      </div>
    </div>
  );
}

export default TbmSummaryView;
