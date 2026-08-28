import React, { useRef, useEffect } from 'react';

function BrowseField({ label, value, onChange, onBrowse, browseLoading, disabled = false, placeholder = "Click browse or paste absolute path..." }) {
  const textareaRef = useRef(null);

  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
      textareaRef.current.style.height = `${Math.max(42, textareaRef.current.scrollHeight)}px`;
    }
  }, [value]);

  const handleTextChange = (e) => {
    onChange(e);
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
      textareaRef.current.style.height = `${Math.max(42, textareaRef.current.scrollHeight)}px`;
    }
  };

  return (
    <div className="form-group">
      <label>{label}</label>
      <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'flex-start' }}>
        <textarea 
          ref={textareaRef}
          rows={1}
          value={value} 
          onChange={handleTextChange} 
          placeholder={placeholder}
          required
          disabled={disabled}
          className="path-input-textarea"
          style={{ 
            flex: 1, 
            minHeight: '42px',
            resize: 'vertical',
            wordBreak: 'break-all',
            overflowWrap: 'anywhere',
            whiteSpace: 'pre-wrap',
            fontFamily: 'Consolas, "Courier New", monospace, sans-serif',
            fontSize: '0.9rem',
            lineHeight: '1.4',
            padding: '0.65rem 0.75rem'
          }}
        />
        <button 
          type="button" 
          onClick={onBrowse} 
          disabled={browseLoading || disabled}
          style={{ 
            width: 'auto', 
            whiteSpace: 'nowrap', 
            backgroundColor: 'var(--primary-color)', 
            color: '#fff', 
            border: 'none',
            minHeight: '42px',
            padding: '0.65rem 1.25rem',
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '0.35rem',
            flexShrink: 0
          }}
        >
          {browseLoading ? '⌛ Browsing...' : '🔍 Browse PC'}
        </button>
      </div>
    </div>
  );
}

export default BrowseField;
