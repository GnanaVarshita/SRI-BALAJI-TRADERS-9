import React from 'react';

function FormField({ label, value, onChange, placeholder, required = false, type = 'text' }) {
  return (
    <div className="form-group">
      <label>{label}</label>
      <input 
        type={type} 
        value={value} 
        onChange={onChange} 
        placeholder={placeholder} 
        required={required} 
      />
    </div>
  );
}

export default FormField;
