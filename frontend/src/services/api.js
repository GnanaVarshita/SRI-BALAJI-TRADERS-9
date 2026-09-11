/**
 * Centralized API Service
 * Sri Balaji Traders Automation System
 */

async function postJson(endpoint, data = {}) {
  const res = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  const json = await res.json();
  if (!res.ok) {
    throw new Error(json.message || `Request failed with status ${res.status}`);
  }
  return json;
}

async function getJson(endpoint) {
  const res = await fetch(endpoint);
  const json = await res.json();
  if (!res.ok) {
    throw new Error(json.message || `Request failed with status ${res.status}`);
  }
  return json;
}

export const api = {
  // Dialogs
  browseFile: () => postJson('/api/browse-file'),
  browseFolder: () => postJson('/api/browse-folder'),

  // System & Status
  getConfig: () => getJson('/api/config'),
  saveConfig: (data) => postJson('/api/config', data),
  getStatus: () => getJson('/api/status'),
  startSync: () => postJson('/api/sync'),
  resetSync: () => postJson('/api/reset'),
  openFolder: (folder) => postJson('/api/open-folder', { folder }),
  getDownloads: () => getJson('/api/downloads'),

  // Business Processes
  processExcel: (data) => postJson('/api/process-excel', data),
  generateSummary: (data) => postJson('/api/generate-summary', data),
  generateCortevaMasterSummary: (data) => postJson('/api/generate-corteva-master-summary', data),
  generateFmcSummary: (data) => postJson('/api/generate-fmc-summary', data),
  generateFmcStep2: (data) => postJson('/api/generate-fmc-step2', data),
  formatTbmSummaries: (data) => postJson('/api/format-tbm-summaries', data),
  generateTbmSummary: (data) => postJson('/api/generate-tbm-summary', data),
  syncTbmCards: (data) => postJson('/api/sync-tbm-cards', data),
  generateInvoices: (data) => postJson('/api/generate-invoices', data),
  scanPosInSummary: (tbmSummaryPath) => postJson('/api/scan-pos-in-summary', { tbmSummaryPath }),
  syncDetailsOfBills: (data) => postJson('/api/sync-details-of-bills', data),
};
