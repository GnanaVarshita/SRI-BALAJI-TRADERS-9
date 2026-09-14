import React, { useState } from 'react';
import Step1PoSummaryCard from '../summary/Step1PoSummaryCard';
import Step2CortevaMasterCard from '../summary/Step2CortevaMasterCard';

function SummaryView() {
  const [activeStep, setActiveStep] = useState('step1');

  return (
    <div className="view-container">
      <div className="view-header">
        <h2>Corteva PO Summary Generator</h2>
        <p className="subtitle">
          {activeStep === 'step1'
            ? 'Convert a quotation excel file into a tracking PO Summary excel workbook.'
            : 'Aggregate all PO summary card workbooks from a folder into a single Master PO Summary sheet.'}
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
          Step 1: PO Summary Cards (From Quotation)
        </button>
        <button
          type="button"
          className={activeStep === 'step2' ? 'primary' : 'secondary'}
          onClick={() => setActiveStep('step2')}
          style={{ padding: '0.75rem 1.5rem', fontWeight: 'bold' }}
        >
          Step 2: Master PO Summary (From Cards Folder)
        </button>
      </div>

      {activeStep === 'step1' ? <Step1PoSummaryCard /> : <Step2CortevaMasterCard />}
    </div>
  );
}

export default SummaryView;
