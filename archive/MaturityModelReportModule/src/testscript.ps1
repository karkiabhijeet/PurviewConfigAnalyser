
# Import functions
. "$PSScriptRoot\functions\Test-SensitivityLabels.ps1"
. "$PSScriptRoot\functions\Test-DLPPolicies.ps1"

# Define paths
$configPath = "$PSScriptRoot\..\examples\Config_sample - v2.json"
$actualPath = "$PSScriptRoot\..\output\OptimizedReport.json"
$sensitivityReportPath = "$PSScriptRoot\..\output\SensitivityLabelReport.json"
$dlpReportPath = "$PSScriptRoot\..\output\DLPReport.json"

# Run tests
Test-SensitivityLabels -ConfigPath $configPath -ActualPath $actualPath -OutputPath $sensitivityReportPath
Test-DLPPolicies -ConfigPath $configPath -ActualPath $actualPath -OutputPath $dlpReportPath

Write-Host "Comprehensive maturity model evaluation complete. Reports generated at:"
Write-Host "- $sensitivityReportPath"
Write-Host "- $dlpReportPath"
