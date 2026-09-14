import React from 'react';
import BrowseField from '../common/BrowseField';

function TbmConsolidateSection({
  tbmFolderPath,
  outputPath,
  setOutputPath,
  onBrowseOutputFile,
  browseFileLoading,
  priorityPoList,
  setPriorityPoList,
  forceConsolidate,
  setForceConsolidate,
  isLoading,
  loadingStep2,
  onConsolidate
}) {
  return (
    <div
      style={{
        border: '1px solid var(--border-color)',
        background: 'var(--bg-card)',
        padding: '1.15rem',
        borderRadius: '8px',
        display: 'flex',
        flexDirection: 'column',
        gap: '0.85rem',
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <h3 style={{ margin: 0, fontSize: '1.05rem', color: 'var(--primary-color)' }}>
          Step 2: Consolidate Master TBM Summary
        </h3>
        <span
          style={{
            fontSize: '0.78rem',
            background: 'var(--primary-color)',
            color: '#000',
            padding: '0.15rem 0.5rem',
            borderRadius: '4px',
            fontWeight: 'bold',
          }}
        >
          Step 2
        </span>
      </div>

      <BrowseField
        label="Target Output Master Excel File Path"
        value={outputPath}
        onChange={(e) => setOutputPath(e.target.value)}
        onBrowse={onBrowseOutputFile}
        browseLoading={browseFileLoading}
        disabled={isLoading}
        placeholder="Defaults to [Territory]-All-TBMs-Summary.xlsx inside TBM Summary Folder"
      />

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

      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginTop: '-0.2rem' }}>
        <input
          type="checkbox"
          id="forceConsolidateCheck"
          checked={forceConsolidate}
          onChange={(e) => setForceConsolidate(e.target.checked)}
          disabled={isLoading}
          style={{ width: '16px', height: '16px', cursor: 'pointer' }}
        />
        <label
          htmlFor="forceConsolidateCheck"
          style={{ fontSize: '0.82rem', color: 'var(--text-color)', cursor: 'pointer', userSelect: 'none' }}
        >
          Force re-consolidate (include tables already marked as DONE in Sheet 2)
        </label>
      </div>

      <button
        type="button"
        className="primary"
        onClick={onConsolidate}
        disabled={isLoading || !tbmFolderPath}
        style={{ marginTop: '0.25rem', padding: '0.85rem' }}
      >
        {loadingStep2 ? '⏳ Consolidating TBM Summaries...' : '📊 Step 2: Consolidate & Append Master Summary'}
      </button>
    </div>
  );
}

export default TbmConsolidateSection;
