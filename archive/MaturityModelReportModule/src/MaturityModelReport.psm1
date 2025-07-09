function Get-MaturityModelReport {
    param (
        [string]$JsonFilePath,
        [string]$ConfigJsonFilePath,
        [string]$OutputDir = "$PSScriptRoot\..\Output",
        [switch]$IncludeDLP
    )

    . $PSScriptRoot\functions\Test-SensitivityLabels.ps1
    . $PSScriptRoot\functions\Export-SensitivityLabelReportToHtml.ps1
    . $PSScriptRoot\functions\Test-DLPPolicies.ps1
    . $PSScriptRoot\functions\Export-DLPReportToHtml.ps1

    if (-not (Test-Path -Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    # Generate Sensitivity Label Report
    $SensitivityLabelReportPath = Join-Path $OutputDir "SensitivityLabelReport.json"
    $SensitivityLabelHtmlPath = Join-Path $OutputDir "SensitivityLabelReport.html"

    Test-SensitivityLabels -ConfigPath $ConfigJsonFilePath -ActualPath $JsonFilePath -OutputPath $SensitivityLabelReportPath
    Export-SensitivityLabelReportToHtml -JsonReportPath $SensitivityLabelReportPath -HtmlOutputPath $SensitivityLabelHtmlPath

    Write-Host "Sensitivity label evaluation report written to $SensitivityLabelReportPath" -ForegroundColor Green
    Write-Host "HTML sensitivity label report written to $SensitivityLabelHtmlPath" -ForegroundColor Green

    # Generate DLP Report if requested
    if ($IncludeDLP) {
        $DLPReportPath = Join-Path $OutputDir "DLPReport.json"
        $DLPHtmlPath = Join-Path $OutputDir "DLPReport.html"

        Test-DLPPolicies -ConfigPath $ConfigJsonFilePath -ActualPath $JsonFilePath -OutputPath $DLPReportPath
        Export-DLPReportToHtml -ReportFilePath $DLPReportPath -OutputFilePath $DLPHtmlPath

        Write-Host "DLP policy evaluation report written to $DLPReportPath" -ForegroundColor Green
        Write-Host "HTML DLP report written to $DLPHtmlPath" -ForegroundColor Green
    }
}

function Get-DLPMaturityReport {
    param (
        [Parameter(Mandatory)]
        [string]$JsonFilePath,
        [Parameter(Mandatory)]
        [string]$ConfigJsonFilePath,
        [string]$OutputDir = "$PSScriptRoot\..\Output"
    )

    . $PSScriptRoot\functions\Test-DLPPolicies.ps1
    . $PSScriptRoot\functions\Export-DLPReportToHtml.ps1

    if (-not (Test-Path -Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    $DLPReportPath = Join-Path $OutputDir "DLPReport.json"
    $DLPHtmlPath = Join-Path $OutputDir "DLPReport.html"

    Test-DLPPolicies -ConfigPath $ConfigJsonFilePath -ActualPath $JsonFilePath -OutputPath $DLPReportPath
    Export-DLPReportToHtml -ReportFilePath $DLPReportPath -OutputFilePath $DLPHtmlPath

    Write-Host "DLP policy evaluation report written to $DLPReportPath" -ForegroundColor Green
    Write-Host "HTML DLP report written to $DLPHtmlPath" -ForegroundColor Green

    return @{
        JsonReport = $DLPReportPath
        HtmlReport = $DLPHtmlPath
    }
}