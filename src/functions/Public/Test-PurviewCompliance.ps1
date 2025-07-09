function Test-PurviewCompliance {
    <#
    .SYNOPSIS
        Runs compliance assessment against collected Purview configuration data.
    
    .DESCRIPTION
        Evaluates Microsoft Purview configuration against specified control books
        and generates compliance reports.
    
    .PARAMETER OptimizedReportPath
        Path to the OptimizedReport JSON file containing configuration data
    
    .PARAMETER Configuration
        Control book configuration to use for assessment (default: PSPF)
    
    .PARAMETER OutputPath
        Output directory for generated reports
    
    .PARAMETER GenerateExcel
        Generate Excel reports in addition to CSV files
    
    .EXAMPLE
        Test-PurviewCompliance -OptimizedReportPath "C:\Reports\OptimizedReport.json" -Configuration "PSPF" -GenerateExcel
        
        Runs PSPF compliance assessment and generates Excel reports.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OptimizedReportPath,
        
        [Parameter(Mandatory = $false)]
        [string]$Configuration = 'PSPF',
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = (Join-Path $env:USERPROFILE "PurviewConfigAnalyser\Output"),
        
        [Parameter(Mandatory = $false)]
        [switch]$GenerateExcel
    )
    
    # Get module root and configuration paths
    $ModuleRoot = $PSScriptRoot | Split-Path -Parent | Split-Path -Parent | Split-Path -Parent
    $ConfigPath = Join-Path $ModuleRoot "config"
    
    # Configuration file paths
    $ControlConfigPath = Join-Path $ConfigPath "ControlBook_${Configuration}_Config.csv"
    $PropertyConfigPath = Join-Path $ConfigPath "ControlBook_Property_${Configuration}_Config.csv"
    $ResultsPath = Join-Path $OutputPath "results_${Configuration}.csv"
    
    # Validate input files
    if (-not (Test-Path $OptimizedReportPath)) {
        throw "OptimizedReport file not found: $OptimizedReportPath"
    }
    
    if (-not (Test-Path $ControlConfigPath)) {
        throw "Control configuration file not found: $ControlConfigPath"
    }
    
    if (-not (Test-Path $PropertyConfigPath)) {
        throw "Property configuration file not found: $PropertyConfigPath"
    }
    
    Write-Host "Running compliance assessment..." -ForegroundColor Yellow
    Write-Host "  Configuration: $Configuration" -ForegroundColor Gray
    Write-Host "  Control Book: $ControlConfigPath" -ForegroundColor Gray
    Write-Host "  Property Book: $PropertyConfigPath" -ForegroundColor Gray
    Write-Host "  Data Source: $OptimizedReportPath" -ForegroundColor Gray
    
    try {
        # Run the control book assessment
        $Results = Test-ControlBook -ControlConfigPath $ControlConfigPath -PropertyConfigPath $PropertyConfigPath -OptimizedReportPath $OptimizedReportPath -OutputPath $ResultsPath -ConfigurationName $Configuration
        
        # Calculate compliance metrics
        $TotalControls = $Results.Count
        $PassingControls = ($Results | Where-Object { $_.Pass -eq $true }).Count
        $FailingControls = $TotalControls - $PassingControls
        $ComplianceRate = if ($TotalControls -gt 0) { [math]::Round(($PassingControls / $TotalControls) * 100, 2) } else { 0 }
        
        $AssessmentResults = @{
            TotalControls = $TotalControls
            PassingControls = $PassingControls
            FailingControls = $FailingControls
            ComplianceRate = $ComplianceRate
            Results = $Results
            Configuration = $Configuration
            AssessmentDate = Get-Date
        }
        
        Write-Host "✅ Assessment completed successfully" -ForegroundColor Green
        Write-Host "  Total Controls: $TotalControls" -ForegroundColor Gray
        Write-Host "  Passing: $PassingControls" -ForegroundColor Green
        Write-Host "  Failing: $FailingControls" -ForegroundColor Red
        Write-Host "  Compliance Rate: $ComplianceRate%" -ForegroundColor $(
            if ($ComplianceRate -ge 80) { "Green" } 
            elseif ($ComplianceRate -ge 60) { "Yellow" } 
            else { "Red" }
        )
        
        # Generate Excel report if requested
        if ($GenerateExcel) {
            Write-Host "Generating Excel report..." -ForegroundColor Yellow
            
            $ExcelReportPath = Join-Path $OutputPath "MaturityAssessment_${Configuration}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            try {
                # Ensure ImportExcel is available
                if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
                    Install-Module -Name ImportExcel -Force -Scope CurrentUser
                }
                
                Import-Module ImportExcel -Force
                
                # Create Excel workbook with multiple sheets
                $ExcelParams = @{
                    Path = $ExcelReportPath
                    AutoSize = $true
                    AutoFilter = $true
                    BoldTopRow = $true
                    FreezeTopRow = $true
                }
                
                # Summary sheet
                $Summary = [PSCustomObject]@{
                    "Assessment Date" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "Configuration" = $Configuration
                    "Total Controls" = $TotalControls
                    "Passing Controls" = $PassingControls
                    "Failing Controls" = $FailingControls
                    "Compliance Rate %" = $ComplianceRate
                    "Data Source" = $OptimizedReportPath
                }
                
                $Summary | Export-Excel @ExcelParams -WorksheetName "Summary"
                
                # Detailed results sheet
                $Results | Export-Excel @ExcelParams -WorksheetName "Detailed Results"
                
                # Failed controls sheet
                $FailedControls = $Results | Where-Object { $_.Pass -eq $false }
                if ($FailedControls) {
                    $FailedControls | Export-Excel @ExcelParams -WorksheetName "Failed Controls"
                }
                
                # Passed controls sheet
                $PassedControls = $Results | Where-Object { $_.Pass -eq $true }
                if ($PassedControls) {
                    $PassedControls | Export-Excel @ExcelParams -WorksheetName "Passed Controls"
                }
                
                Write-Host "✅ Excel report generated: $ExcelReportPath" -ForegroundColor Green
                $AssessmentResults.ExcelReportPath = $ExcelReportPath
                
            } catch {
                Write-Host "⚠️ Excel generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "  CSV results are available at: $ResultsPath" -ForegroundColor Gray
            }
        }
        
        return $AssessmentResults
        
    } catch {
        Write-Host "❌ Assessment failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
