import React from 'react';

function SelectField({ label, value, onChange, options }) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <select value={value} onChange={onChange} className="filter-select">
        {options.map((opt) => {
          const val = typeof opt === 'object' ? opt.value : opt;
          const labelText = typeof opt === 'object' ? opt.label : opt;
          return (
            <option key={val} value={val}>
              {labelText}
            </option>
          );
        })}
      </select>
    </div>
  );
}

export default SelectField;
