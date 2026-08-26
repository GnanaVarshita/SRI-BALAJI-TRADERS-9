import React from 'react';
import { NavLink } from 'react-router-dom';

function Sidebar() {
  return (
    <aside className="sidebar">
      <div className="sidebar-logo">
        <span className="logo-icon">🌳</span>
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
          to="/explorer" 
          className={({ isActive }) => `nav-link ${isActive ? 'active' : ''}`}
        >
          <span className="nav-icon">📁</span> PO Explorer
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
