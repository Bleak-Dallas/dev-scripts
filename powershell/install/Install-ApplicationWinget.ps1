<#
.SYNOPSIS
  Installs a list of applications via winget sequentially.
.DESCRIPTION
  This script loops through a predefined array of apps, installs them using winget,
  and waits for each to complete before starting the next install.
#>

# -------------------------------
# 🧩 Configuration Section
# -------------------------------

# Define the list of applications to install
$appList = @(
  "Microsoft.VisualStudioCode",
  "Yubico.Authenticator",
  "Microsoft.PowerAutomateDesktop",
  "DisplayLink.GraphicsDriver"
)

# Define log directory and file path
$logDir = "C:\temp\logs"
$logFile = Join-Path $logDir "winget_install_log.txt"

# Ensure log directory exists (create if it doesn’t)
if (-not (Test-Path -Path $logDir)) {
    Write-Host "📁 Creating log directory at $logDir..." -ForegroundColor Cyan
    New-Item -Path $logDir -ItemType Directory | Out-Null
}

# -------------------------------
# ⚙️ Installation Loop
# -------------------------------

Write-Host "Starting application installations..." -ForegroundColor Cyan

foreach ($app in $appList) {
    Write-Host "--------------------------------------------------"
    Write-Host "🔹 Installing: $app" -ForegroundColor Yellow
    
    # Run winget synchronously; wait until it finishes
    $process = Start-Process -FilePath "winget" `
        -ArgumentList "install --id $app --accept-package-agreements --accept-source-agreements --silent" `
        -Wait -PassThru -NoNewWindow

    # Check exit status
    if ($process.ExitCode -eq 0) {
        Write-Host "✅ Successfully installed: $app" -ForegroundColor Green
        Add-Content -Path $logFile -Value "$(Get-Date): SUCCESS - $app"
    } else {
        Write-Host "❌ Failed to install: $app" -ForegroundColor Red
        Add-Content -Path $logFile -Value "$(Get-Date): FAILED - $app (Exit Code: $($process.ExitCode))"
    }
}

Write-Host "--------------------------------------------------"
Write-Host "🎯 All installs complete. Log saved to: $logFile" -ForegroundColor Cyan