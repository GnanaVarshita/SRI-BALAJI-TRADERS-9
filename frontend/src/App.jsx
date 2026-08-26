import React, { useState, useEffect, useRef } from 'react';
import { HashRouter as Router, Routes, Route } from 'react-router-dom';
import Sidebar from './components/Sidebar';
import Toast from './components/Toast';

// Import Views
import DashboardView from './components/views/DashboardView';
import SyncView from './components/views/SyncView';
import ExplorerView from './components/views/ExplorerView';
import SettingsView from './components/views/SettingsView';

function App() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [hasPassword, setHasPassword] = useState(false);
  
  const [isSyncing, setIsSyncing] = useState(false);
  const [totalSynced, setTotalSynced] = useState(0);
  const [logs, setLogs] = useState([]);
  
  const [statusMessage, setStatusMessage] = useState(null);
  const [messageType, setMessageType] = useState('success');
  
  const consoleRef = useRef(null);

  useEffect(() => {
    fetchConfig();
    fetchStatus();
    const interval = setInterval(fetchStatus, 1000);
    return () => clearInterval(interval);
  }, []);

  const showToast = (msg, type = 'success') => {
    setStatusMessage(msg);
    setMessageType(type);
    setTimeout(() => {
      setStatusMessage(null);
    }, 4000);
  };

  const fetchConfig = async () => {
    try {
      const res = await fetch('/api/config');
      if (res.ok) {
        const data = await res.json();
        setEmail(data.email || '');
        setHasPassword(data.hasPassword || false);
      }
    } catch (err) {
      console.error('Error fetching config:', err);
    }
  };

  const fetchStatus = async () => {
    try {
      const res = await fetch('/api/status');
      if (res.ok) {
        const data = await res.json();
        setIsSyncing(data.isSyncing);
        setLogs(data.logs || []);
        setTotalSynced(data.totalSynced || 0);
      }
    } catch (err) {
      console.error('Error fetching status:', err);
    }
  };

  const handleSaveCredentials = async (e) => {
    e.preventDefault();
    try {
      const res = await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });
      const data = await res.json();
      if (data.success) {
        showToast('Credentials saved successfully!');
        setPassword('');
        fetchConfig();
      } else {
        showToast(data.message || 'Failed to save credentials', 'error');
      }
    } catch (err) {
      showToast('Connection error: Failed to save credentials', 'error');
    }
  };

  const handleStartSync = async () => {
    if (isSyncing) return;
    try {
      const res = await fetch('/api/sync', { method: 'POST' });
      const data = await res.json();
      if (data.success) {
        setIsSyncing(true);
        showToast('Gmail Sync started!');
      } else {
        showToast(data.message || 'Failed to start sync', 'error');
      }
    } catch (err) {
      showToast('Connection error: Failed to trigger sync', 'error');
    }
  };

  const handleOpenFolder = async (folder) => {
    try {
      const res = await fetch('/api/open-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder })
      });
      const data = await res.json();
      if (!data.success) {
        showToast(data.message || 'Failed to open folder', 'error');
      }
    } catch (err) {
      showToast('Connection error: Failed to open folder', 'error');
    }
  };

  const handleResetHistory = async () => {
    const confirm = window.confirm(
      "Are you sure you want to reset your sync history?\n\n" +
      "This clears the database of processed email IDs. The next sync will re-process and check all attachments. No local files will be deleted."
    );
    if (!confirm) return;

    try {
      const res = await fetch('/api/reset', { method: 'POST' });
      const data = await res.json();
      if (data.success) {
        showToast('Sync history has been reset complete.');
        fetchStatus();
      } else {
        showToast(data.message || 'Failed to reset history', 'error');
      }
    } catch (err) {
      showToast('Connection error: Failed to reset history', 'error');
    }
  };

  return (
    <Router>
      <div className="app">
        <Sidebar />
        
        <main className="main-content">
          <Toast message={statusMessage} type={messageType} />

          <Routes>
            <Route 
              path="/" 
              element={
                <DashboardView 
                  isSyncing={isSyncing} 
                  totalSynced={totalSynced} 
                  onOpenFolder={handleOpenFolder} 
                />
              } 
            />
            <Route 
              path="/sync" 
              element={
                <SyncView 
                  isSyncing={isSyncing} 
                  logs={logs} 
                  consoleRef={consoleRef} 
                  onStartSync={handleStartSync} 
                />
              } 
            />
            <Route 
              path="/explorer" 
              element={<ExplorerView />} 
            />
            <Route 
              path="/settings" 
              element={
                <SettingsView 
                  email={email} 
                  setEmail={setEmail} 
                  password={password} 
                  setPassword={setPassword} 
                  hasPassword={hasPassword} 
                  isSyncing={isSyncing} 
                  onSave={handleSaveCredentials} 
                  onResetHistory={handleResetHistory} 
                />
              } 
            />
          </Routes>
        </main>
      </div>
    </Router>
  );
}

export default App;
