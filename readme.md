# Sri Balaji Traders - Gmail PO Automation Dashboard (Golden Edition)

This monorepo project automates the downloading and organization of purchase order (PO) email attachments from Gmail for **Corteva** and **New Gen** activities. It features a luxury **Golden Kalpavruksham** (Royal Gold and Slate Charcoal) theme, a sidebar navigation layout, React routing, and a PO File Explorer.

---
npm run dev 
## Monorepo Directory Structure

```
D:\SRI BALAJI TRADERS\
├── package.json                   # Root package.json (Monorepo scripts)
├── readme.md                      # Project documentation
├── .env                           # Shared credentials
├── processed_emails.json          # Sync history database
├── run_dashboard.ps1              # Unified PowerShell launcher
├── backend/                       # Backend Python Workspace
│   ├── server.py                  # Python API & static file web server
│   └── download_attachments.py    # Gmail downloader script
└── frontend/                      # Frontend React+Vite Workspace
    ├── package.json               # Vite+React configurations
    ├── vite.config.js             # Vite server / proxy settings
    ├── index.html                 # Root html template
    └── src/
        ├── main.jsx               # Entry point
        ├── App.jsx                # Router & State coordinator
        ├── App.css                # Golden theme stylesheet
        ├── components/            # Shared UI components
        │   ├── Sidebar.jsx        # Sidebar navigation panel
        │   ├── ConsoleLogs.jsx    # Real-time sync logs container
        │   ├── CredentialsConfig.jsx # Gmail credentials editor card
        │   └── Toast.jsx          # Success/error notifications
        └── components/views/      # Web pages (React Router)
            ├── DashboardView.jsx  # Home page statistics and folder openers
            ├── SyncView.jsx       # Real-time downloader sync control
            ├── ExplorerView.jsx   # Search, filter, and download PO files
            └── SettingsView.jsx   # Credentials configuration panel
```

---

## Unified Execution Architecture

To launch both applications together:
1. **Production Mode (`npm start` or double-click)**:
   - Builds the frontend React bundle into `frontend/dist/`.
   - Starts the Python server, which serves both the API endpoints *and* the built static React page under a single port (`http://127.0.0.1:5000`).
2. **Development Mode (`npm run dev`)**:
   - Uses `concurrently` (an npm package) to launch:
     - The **Vite Dev Server** (at `http://localhost:5173`) with instant hot-reloading.
     - The **Python Backend Server** (at `http://127.0.0.1:5000`) for the API.
     - A proxy is configured in Vite (`vite.config.js`) so that frontend API requests to `/api/*` are automatically forwarded to the Python server.

---

## Configuration & Credentials Setup

Because Google requires **App Passwords** for third-party scripts to access Gmail, follow these steps to set up your credentials:

1. **Enable 2-Step Verification**: If not already enabled, turn on 2-Step Verification in your Google Account security settings.
2. **Generate an App Password**:
   - Go to your [Google Account Security Settings](https://myaccount.google.com/security).
   - Search for **App passwords** (or go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)).
   - Enter a name (e.g., `Gmail Downloader`) and click **Create**.
   - Copy the 16-character password shown (e.g., `abcd efgh ijkl mnop`).
3. **Configure in Web Dashboard**:
   - Open the web dashboard (see below).
   - Navigate to the **Settings** view in the sidebar.
   - Enter your Gmail address and App Password in the configuration card.
   - Click **Save Credentials**. They will be saved to your local `.env` file securely.

---

## How to Run the Dashboard

### Method A: Use the PowerShell Launcher (Recommended)
1. Navigate to `D:\SRI BALAJI TRADERS` in Windows Explorer.
2. Right-click [`run_dashboard.ps1`](file:///d:/SRI%20BALAJI%20TRADERS/run_dashboard.ps1) and choose **Run with PowerShell**.
3. The script will automatically install dependencies, compile the frontend, start the server, and open your default browser.

### Method B: From the Terminal
Run the following commands in PowerShell:
```powershell
cd "D:\SRI BALAJI TRADERS"
npm run install:all   # Install all dependencies (run once)
npm start             # Compile frontend & start Python server
```
Once started, open your web browser and navigate to:
**[http://127.0.0.1:5000](http://127.0.0.1:5000)**
