# run_downloader.ps1
# Helper script to run the Gmail Attachment Downloader with the correct Python interpreter.

$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$PythonScript = Join-Path $ScriptDir "download_attachments.py"
$EnvFile = Join-Path $ScriptDir ".env"

if (-not (Test-Path $EnvFile)) {
    Write-Host "WARNING: .env file is missing! The downloader will fail without credentials." -ForegroundColor Red
    Write-Host "Creating a template .env file..." -ForegroundColor Yellow
    "GMAIL_EMAIL=your_email@gmail.com`nGMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx" | Out-File -FilePath $EnvFile -Encoding utf8
    Write-Host "Please edit $EnvFile and put your actual Gmail credentials." -ForegroundColor Yellow
}

# Define search paths for Python
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
        # Check if python is available in PATH
        Get-Command python.exe -ErrorAction SilentlyContinue | Out-Null
        if ($LastExitCode -eq 0 -or $?) {
            # Ensure it is not the Windows Store execution alias that fails
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

if ($null -ne $PythonExe) {
    Write-Host "Using Python executable: $PythonExe" -ForegroundColor Cyan
    & $PythonExe $PythonScript
} else {
    Write-Error "Could not find a valid Python installation. Please install Python or add it to your PATH."
    Write-Host "We searched standard paths, including: $DefaultUserPath" -ForegroundColor Red
}

Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
try {
    $null = [Console]::ReadKey($true)
} catch {
    # If console input is redirected (e.g. non-interactive), skip waiting
}
