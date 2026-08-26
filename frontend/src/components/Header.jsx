import React from 'react';

function Header({ isSyncing }) {
  return (
    <header>
      <h1>SRI BALAJI TRADERS - PO AUTOMATION</h1>
      <div className={`status-badge ${isSyncing ? 'syncing' : ''}`}>
        {isSyncing ? '● Sync Active' : '○ Standby'}
      </div>
    </header>
  );
}

export default Header;
