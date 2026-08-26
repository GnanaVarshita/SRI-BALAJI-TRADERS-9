import React, { useState } from 'react';

function CredentialsConfig({ email, setEmail, password, setPassword, hasPassword, isSyncing, onSave }) {
  const [showPassword, setShowPassword] = useState(false);

  return (
    <section className="card">
      <h2>Credentials Configuration</h2>
      <form onSubmit={onSave} style={{ display: 'flex', flex: '1', flexDirection: 'column', gap: '1rem' }}>
        <div className="form-group">
          <label htmlFor="email">Gmail Address</label>
          <input 
            type="email" 
            id="email" 
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="your_email@gmail.com"
            required
            disabled={isSyncing}
          />
        </div>
        
        <div className="form-group">
          <label htmlFor="password">Gmail App Password</label>
          <div className="password-input-container">
            <input 
              type={showPassword ? 'text' : 'password'} 
              id="password" 
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder={hasPassword ? '•••••••••••••••• (Saved)' : 'Enter 16-character app password'}
              disabled={isSyncing}
            />
            <button 
              type="button" 
              className="toggle-password-btn"
              onClick={() => setShowPassword(!showPassword)}
            >
              {showPassword ? 'Hide' : 'Show'}
            </button>
          </div>
          <small style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginTop: '0.25rem' }}>
            Create a 16-character App Password under Google Account &gt; Security &gt; 2-Step Verification &gt; App Passwords.
          </small>
        </div>

        <button 
          type="submit" 
          className="primary" 
          style={{ marginTop: 'auto', width: 'auto', alignSelf: 'flex-end' }}
          disabled={isSyncing}
        >
          Save Credentials
        </button>
      </form>
    </section>
  );
}

export default CredentialsConfig;
