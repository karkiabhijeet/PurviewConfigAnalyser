# Easy Install Script for PurviewConfigAnalyser
# This script provides better feedback during installation

Write-Host "=== PurviewConfigAnalyser Installation Helper ===" -ForegroundColor Green
Write-Host ""

# Check PowerShell version
Write-Host "Checking PowerShell version..." -ForegroundColor Yellow
$psVersion = $PSVersionTable.PSVersion
Write-Host "PowerShell Version: $psVersion" -ForegroundColor Cyan

if ($psVersion.Major -lt 5 -or ($psVersion.Major -eq 5 -and $psVersion.Minor -lt 1)) {
    Write-Host "[ERROR] PowerShell 5.1 or higher is required. Please upgrade PowerShell." -ForegroundColor Red
    exit 1
}
Write-Host "[SUCCESS] PowerShell version is compatible" -ForegroundColor Green
Write-Host ""

# Check execution policy
Write-Host "Checking execution policy..." -ForegroundColor Yellow
$executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
Write-Host "Current execution policy: $executionPolicy" -ForegroundColor Cyan

if ($executionPolicy -eq 'Restricted') {
    Write-Host "Setting execution policy for current user..." -ForegroundColor Yellow
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "[SUCCESS] Execution policy updated" -ForegroundColor Green
} else {
    Write-Host "[SUCCESS] Execution policy is acceptable" -ForegroundColor Green
}
Write-Host ""

# Install ImportExcel first (this is the large dependency)
Write-Host "Step 1: Installing ImportExcel module (this may take 1-2 minutes)..." -ForegroundColor Yellow
Write-Host "[PACKAGE] ImportExcel is ~1MB and provides Excel export capabilities" -ForegroundColor Cyan
Write-Host "[WAIT] Please be patient, especially on slower connections..." -ForegroundColor Cyan

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -Verbose
    } else {
        Write-Host "[SUCCESS] ImportExcel already installed" -ForegroundColor Green
    }
    
    $stopwatch.Stop()
    Write-Host "[SUCCESS] ImportExcel installation completed in $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to install ImportExcel: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[TIP] Try running: Install-Module -Name ImportExcel -Force" -ForegroundColor Yellow
    exit 1
}
Write-Host ""

# Install PurviewConfigAnalyser
Write-Host "Step 2: Installing PurviewConfigAnalyser module..." -ForegroundColor Yellow
Write-Host "[PACKAGE] PurviewConfigAnalyser is ~60KB and should install quickly" -ForegroundColor Cyan

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Install-Module -Name PurviewConfigAnalyser -Force -Scope CurrentUser -Verbose
    $stopwatch.Stop()
    Write-Host "[SUCCESS] PurviewConfigAnalyser installation completed in $($stopwatch.Elapsed.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to install PurviewConfigAnalyser: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Verify installation
Write-Host "Step 3: Verifying installation..." -ForegroundColor Yellow
try {
    Import-Module PurviewConfigAnalyser -Force
    $commands = Get-Command -Module PurviewConfigAnalyser
    Write-Host "[SUCCESS] Module imported successfully" -ForegroundColor Green
    Write-Host "[SUCCESS] Available commands: $($commands.Count)" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "Available Commands:" -ForegroundColor Cyan
    foreach ($cmd in $commands) {
        Write-Host "  - $($cmd.Name)" -ForegroundColor White
    }
    
} catch {
    Write-Host "[ERROR] Failed to import module: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[COMPLETE] Installation completed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Quick Start:" -ForegroundColor Yellow
Write-Host "  1. Run: Invoke-PurviewConfigAnalyser" -ForegroundColor Cyan
Write-Host "  2. Follow the interactive menu" -ForegroundColor Cyan
Write-Host "  3. See CLIENT_INSTALLATION.md for detailed usage" -ForegroundColor Cyan
Write-Host ""
Write-Host "Need help? Check INSTALLATION_TROUBLESHOOTING.md" -ForegroundColor Yellow

# Optional: Launch the main function
$launch = Read-Host "Would you like to launch PurviewConfigAnalyser now? (y/n)"
if ($launch -eq 'y' -or $launch -eq 'yes') {
    Write-Host ""
    Write-Host "Launching PurviewConfigAnalyser..." -ForegroundColor Green
    Invoke-PurviewConfigAnalyser
}
