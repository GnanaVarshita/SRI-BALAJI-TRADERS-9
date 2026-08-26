# run_dashboard.ps1
# Launcher for the Sri Balaji Traders Web Automation Dashboard.

$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Sri Balaji Traders PO Web Dashboard Launcher" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

# 1. Locate Python
$DefaultUserPath = Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"
$SystemPaths = @(
    "python.exe",
    $DefaultUserPath,
    (Join-Path $env:LocalAppData "Programs\Python\Python313\python.exe"),
    (Join-Path $env:LocalAppData "Programs\Python\Python311\python.exe"),
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files\Python313\python.exe"
)

$PythonExe = $null

foreach ($Path in $SystemPaths) {
    if ($Path -eq "python.exe") {
        Get-Command python.exe -ErrorAction SilentlyContinue | Out-Null
        if ($LastExitCode -eq 0 -or $?) {
            & python.exe --version > $null 2>&1
            if ($?) {
                $PythonExe = "python.exe"
                break
            }
        }
    } else {
        if (Test-Path $Path) {
            $PythonExe = $Path
            break
        }
    }
}

if ($null -eq $PythonExe) {
    Write-Error "Could not find a valid Python installation. Please check your Python installation."
    Write-Host "Press any key to exit..."
    try { $null = [Console]::ReadKey($true) } catch {}
    exit 1
}

# 2. Check Node and npm dependencies
Write-Host "Checking frontend dependencies..." -ForegroundColor Yellow
if (-not (Test-Path "node_modules") -or -not (Test-Path "frontend/node_modules")) {
    Write-Host "Installing npm dependencies (this may take a minute on first run)..." -ForegroundColor Yellow
    npm run install:all
    if ($LastExitCode -ne 0) {
        Write-Error "Failed to install dependencies."
        Write-Host "Press any key to exit..."
        try { $null = [Console]::ReadKey($true) } catch {}
        exit 1
    }
}

# 3. Compile the React frontend
Write-Host "Building React frontend..." -ForegroundColor Yellow
npm run build:frontend
if ($LastExitCode -ne 0) {
    Write-Error "Failed to build the React frontend."
    Write-Host "Press any key to exit..."
    try { $null = [Console]::ReadKey($true) } catch {}
    exit 1
}

# 4. Start the backend Python server in the background
Write-Host "Starting Python API server on http://127.0.0.1:5000 ..." -ForegroundColor Green
Start-Process -FilePath $PythonExe -ArgumentList "backend/server.py" -WorkingDirectory $ScriptDir -NoNewWindow

# 5. Wait a brief moment and open the browser
Start-Sleep -Seconds 2
Write-Host "Opening web browser..." -ForegroundColor Green
Start-Process "http://127.0.0.1:5000"

Write-Host "`nWeb Dashboard is running! You can keep this window open or close it." -ForegroundColor Cyan
Write-Host "Press any key to exit this launcher window..." -ForegroundColor Yellow
try {
    $null = [Console]::ReadKey($true)
} catch {
    # Non-interactive fallback
}
