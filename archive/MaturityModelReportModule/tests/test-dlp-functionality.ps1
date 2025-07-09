# Test script for DLP Policy Evaluation
# This script demonstrates how to use the new DLP evaluation functionality

# Import the module
Import-Module "$PSScriptRoot\..\src\MaturityModelReport.psm1" -Force

# Set paths
$ConfigPath = "$PSScriptRoot\..\examples\Config_sample.json"
$ActualDataPath = "$PSScriptRoot\..\output\SensitivityLabelReport.json"  # Using existing data for demo
$OutputDir = "$PSScriptRoot\..\output"

Write-Host "Testing DLP Policy Evaluation..." -ForegroundColor Cyan

try {
    # Test 1: Generate DLP report only
    Write-Host "`nTest 1: Generating DLP-only report..." -ForegroundColor Yellow
    
    $result = Get-DLPMaturityReport -JsonFilePath $ActualDataPath -ConfigJsonFilePath $ConfigPath -OutputDir $OutputDir
    
    Write-Host "✓ DLP report generated successfully!" -ForegroundColor Green
    Write-Host "  JSON Report: $($result.JsonReport)" -ForegroundColor Gray
    Write-Host "  HTML Report: $($result.HtmlReport)" -ForegroundColor Gray

    # Test 2: Generate combined report with DLP
    Write-Host "`nTest 2: Generating combined report with DLP..." -ForegroundColor Yellow
    
    Get-MaturityModelReport -JsonFilePath $ActualDataPath -ConfigJsonFilePath $ConfigPath -OutputDir $OutputDir -IncludeDLP
    
    Write-Host "✓ Combined report with DLP generated successfully!" -ForegroundColor Green

    # Test 3: Verify output files exist
    Write-Host "`nTest 3: Verifying output files..." -ForegroundColor Yellow
    
    $dlpJsonPath = Join-Path $OutputDir "DLPReport.json"
    $dlpHtmlPath = Join-Path $OutputDir "DLPReport.html"
    
    if (Test-Path $dlpJsonPath) {
        Write-Host "✓ DLP JSON report exists: $dlpJsonPath" -ForegroundColor Green
        $fileSize = (Get-Item $dlpJsonPath).Length
        Write-Host "  File size: $fileSize bytes" -ForegroundColor Gray
    } else {
        Write-Host "✗ DLP JSON report not found!" -ForegroundColor Red
    }
    
    if (Test-Path $dlpHtmlPath) {
        Write-Host "✓ DLP HTML report exists: $dlpHtmlPath" -ForegroundColor Green
        $fileSize = (Get-Item $dlpHtmlPath).Length
        Write-Host "  File size: $fileSize bytes" -ForegroundColor Gray
    } else {
        Write-Host "✗ DLP HTML report not found!" -ForegroundColor Red
    }

    Write-Host "`n🎉 All tests completed successfully!" -ForegroundColor Green
    Write-Host "You can now open the HTML reports to view the DLP evaluation results." -ForegroundColor Cyan

} catch {
    Write-Host "❌ Error occurred during testing:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
}
