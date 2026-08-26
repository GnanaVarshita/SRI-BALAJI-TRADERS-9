import React from 'react';
import { NavLink } from 'react-router-dom';
import logoImg from '../assets/logo.png';

function Sidebar() {
  return (
    <aside className="sidebar">
      <div className="sidebar-logo">
        <img src={logoImg} alt="Sri Balaji Traders Logo" style={{ width: '36px', height: '36px', objectFit: 'contain' }} />
        <h2>Sri Balaji Traders</h2>
      </div>
      <nav className="sidebar-nav">
        <NavLink 
          to="/" 
          end 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">📊</span> Dashboard
        </NavLink>
        <NavLink 
          to="/sync" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">🔄</span> Gmail Sync
        </NavLink>
        <NavLink 
          to="/quotation" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">📝</span> Quotation Gen
        </NavLink>
        <NavLink 
          to="/summary" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">📋</span> Corteva PO Summary
        </NavLink>
        <NavLink 
          to="/fmc-summary" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">📊</span> FMC PO Summary
        </NavLink>
        <NavLink 
          to="/settings" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">⚙️</span> Settings
        </NavLink>
      </nav>
      <div className="sidebar-footer">
        <p>V1.1.0 • Golden Edition</p>
      </div>
    </aside>
  );
}

export default Sidebar;
