# Import the new PSPF control testing function
. "$PSScriptRoot\functions\Test-PSPFControls.ps1"

# Define paths for PSPF configuration
$pspfConfigPath = "$PSScriptRoot\..\config\ControlBook_PSPF_Config.csv"
$pspfPropertyConfigPath = "$PSScriptRoot\..\config\ControlBook_Property_PSPF_Config.csv"
$optimizedReportPath = "$PSScriptRoot\..\output\OptimizedReport.json"
$pspfResultsPath = "$PSScriptRoot\..\output\results_pspf.csv"

Write-Host "Starting PSPF Control-Based Maturity Model Evaluation..."
Write-Host "Configuration files:"
Write-Host "- Controls: $pspfConfigPath"
Write-Host "- Properties: $pspfPropertyConfigPath"
Write-Host "- Data Source: $optimizedReportPath"
Write-Host ""

# Run PSPF control evaluation
Test-PSPFControls -ConfigPath $pspfConfigPath -PropertyConfigPath $pspfPropertyConfigPath -OptimizedReportPath $optimizedReportPath -OutputPath $pspfResultsPath

Write-Host ""
Write-Host "PSPF Control evaluation complete!"
Write-Host "Results saved to: $pspfResultsPath"
Write-Host ""
Write-Host "You can now review the results in Excel or import the CSV to analyze:"
Write-Host "- Pass/Fail status for each control"
Write-Host "- Detailed comments explaining evaluation outcomes"
Write-Host "- Property-level assessment based on your PSPF requirements"
