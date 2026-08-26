import React, { useState, useEffect } from 'react';

function ExplorerView() {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [companyFilter, setCompanyFilter] = useState('All');
  const [yearFilter, setYearFilter] = useState('All');
  const [areaFilter, setAreaFilter] = useState('All');

  useEffect(() => {
    fetchDownloads();
  }, []);

  const fetchDownloads = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/downloads');
      if (res.ok) {
        const data = await res.json();
        setFiles(data);
      }
    } catch (err) {
      console.error('Error fetching downloads:', err);
    } finally {
      setLoading(false);
    }
  };

  const formatBytes = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const formatDate = (isoStr) => {
    if (!isoStr) return '-';
    try {
      const d = new Date(isoStr);
      return d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    } catch (e) {
      return isoStr;
    }
  };

  const companies = ['All', ...new Set(files.map(f => f.company).filter(Boolean))];
  const years = ['All', ...new Set(files.map(f => f.year).filter(y => y && y !== 'Unknown'))];
  const areas = ['All', ...new Set(files.map(f => f.area).filter(a => a && a !== 'Unknown'))];

  const filteredFiles = files.filter(file => {
    const matchesSearch = file.filename.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesCompany = companyFilter === 'All' || file.company === companyFilter;
    const matchesYear = yearFilter === 'All' || file.year === yearFilter;
    const matchesArea = areaFilter === 'All' || file.area === areaFilter;
    return matchesSearch && matchesCompany && matchesYear && matchesArea;
  });

  return (
    <div className="view-container">
      <div className="view-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h2>PO File Explorer</h2>
          <p className="subtitle">Browse, filter, and download all purchase order files organized on your local computer.</p>
        </div>
        <button onClick={fetchDownloads} style={{ width: 'auto', padding: '0.5rem 1rem' }} disabled={loading}>
          🔄 {loading ? 'Loading...' : 'Refresh List'}
        </button>
      </div>

      {/* Filters Card */}
      <div className="card filters-grid">
        <div className="form-group">
          <label>Search Filename</label>
          <input 
            type="text" 
            placeholder="Search by filename..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        <div className="form-group">
          <label>Company</label>
          <select 
            value={companyFilter} 
            onChange={(e) => setCompanyFilter(e.target.value)}
            className="filter-select"
          >
            {companies.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div className="form-group">
          <label>Year</label>
          <select 
            value={yearFilter} 
            onChange={(e) => setYearFilter(e.target.value)}
            className="filter-select"
          >
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>
        <div className="form-group">
          <label>Area</label>
          <select 
            value={areaFilter} 
            onChange={(e) => setAreaFilter(e.target.value)}
            className="filter-select"
          >
            {areas.map(a => <option key={a} value={a}>{a}</option>)}
          </select>
        </div>
      </div>

      {/* File Table Card */}
      <div className="card" style={{ marginTop: '1.5rem', padding: '0', overflowX: 'auto' }}>
        {loading ? (
          <div style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-muted)' }}>
            <h3>Loading files list...</h3>
          </div>
        ) : filteredFiles.length === 0 ? (
          <div style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-muted)' }}>
            <h3>No purchase order files found.</h3>
            <p>Try running a sync or adjusting your filters.</p>
          </div>
        ) : (
          <table className="po-table">
            <thead>
              <tr>
                <th>Company</th>
                <th>Year</th>
                <th>Area</th>
                <th>Filename</th>
                <th>Size</th>
                <th>Synced Date</th>
                <th style={{ textAlign: 'center' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredFiles.map((file, idx) => (
                <tr key={idx}>
                  <td><strong style={{ color: 'var(--primary-color)' }}>{file.company}</strong></td>
                  <td>{file.year}</td>
                  <td>{file.area}</td>
                  <td className="filename-cell" title={file.filename}>{file.filename}</td>
                  <td>{formatBytes(file.sizeBytes)}</td>
                  <td>{formatDate(file.modified)}</td>
                  <td style={{ textAlign: 'center' }}>
                    <a 
                      href={`/api/view-file?path=${encodeURIComponent(file.relativePath)}`}
                      className="download-link"
                      download
                    >
                      📥 Download
                    </a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

export default ExplorerView;
