import React from 'react';

function Toast({ message, type }) {
  if (!message) return null;
  return (
    <div className="full-width">
      <div className={`toast ${type}`}>
        <span>{message}</span>
      </div>
    </div>
  );
}

export default Toast;
