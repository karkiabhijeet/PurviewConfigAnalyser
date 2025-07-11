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
                
                # Maturity Level Summary sheet - enriched with descriptions
                $MaturityLevelSummary = @()
                $maturityGroups = $Results | Group-Object MaturityLevel
                foreach ($group in $maturityGroups) {
                    $level = if ($group.Name -eq '' -or $null -eq $group.Name) { '(none)' } else { $group.Name }
                    $total = $group.Group.Count
                    $passing = ($group.Group | Where-Object { $_.Pass -eq $true }).Count
                    $failing = $total - $passing
                    $rate = if ($total -gt 0) { [math]::Round(($passing / $total) * 100, 1) } else { 0 }
                    
                    # Enhanced maturity level descriptions
                    $description = switch ($level.ToLower()) {
                        "1" { "Initial Stage - Basic security controls and foundational data protection measures" }
                        "2" { "Intermediate Stage - Enhanced security policies with automated enforcement and monitoring" }
                        "3" { "Advanced Stage - Comprehensive data security with AI-driven protection and full compliance" }
                        "basic" { "Basic Level - Fundamental security controls implementation" }
                        "intermediate" { "Intermediate Level - Enhanced security with policy automation" }
                        "advanced" { "Advanced Level - Sophisticated security with intelligent protection" }
                        "(none)" { "Unclassified - Controls without assigned maturity levels" }
                        default { "Custom Level - Organization-specific maturity classification: $level" }
                    }
                    
                    $status = if ($rate -ge 90) { "Excellent" } 
                             elseif ($rate -ge 80) { "Good" }
                             elseif ($rate -ge 70) { "Acceptable" }
                             elseif ($rate -ge 60) { "Needs Improvement" }
                             else { "Critical" }
                    
                    $MaturityLevelSummary += [PSCustomObject]@{
                        'Maturity Level' = $level
                        'Description' = $description
                        'Total Controls' = $total
                        'Passing Controls' = $passing
                        'Failing Controls' = $failing
                        'Compliance Rate %' = $rate
                        'Status' = $status
                        'Priority' = if ($level -match "^[123]$") { 
                            switch ($level) { "1" { "High - Foundation" }; "2" { "Medium - Enhancement" }; "3" { "Low - Optimization" } }
                        } else { "Medium" }
                    }
                }
                
                $MaturityLevelSummary | Export-Excel @ExcelParams -WorksheetName "Maturity Level Summary"
                
                # Control Summary sheet - aggregated by Control ID
                $ControlSummary = $Results | Group-Object -Property ControlID | ForEach-Object {
                    $controlGroup = $_.Group
                    $controlPassed = $controlGroup | Where-Object { $_.Pass -eq $true }
                    $controlFailed = $controlGroup | Where-Object { $_.Pass -eq $false }
                    
                    # A control is considered failed if ANY of its properties fail
                    $overallResult = if ($controlFailed.Count -gt 0) { "Fail" } else { "Pass" }
                    
                    # Concatenate all comments from properties for this control
                    $allComments = @()
                    foreach ($property in $controlGroup) {
                        if ($property.Comments -and $property.Comments.Trim() -ne "") {
                            $propertyComment = "$($property.Properties): $($property.Comments)"
                            $allComments += $propertyComment
                        }
                    }
                    $combinedComments = if ($allComments.Count -gt 0) { $allComments -join " | " } else { "" }
                    
                    [PSCustomObject]@{
                        "Capability" = $controlGroup[0].Capability
                        "Control ID" = $controlGroup[0].ControlID
                        "Control" = $controlGroup[0].Control
                        "Maturity Level" = $controlGroup[0].MaturityLevel
                        "Total Properties" = $controlGroup.Count
                        "Properties Passed" = $controlPassed.Count
                        "Properties Failed" = $controlFailed.Count
                        "Result" = $overallResult
                        "Comments" = $combinedComments
                        "Configuration" = $controlGroup[0].ConfigurationName
                    }
                }
                
                $ControlSummary | Export-Excel @ExcelParams -WorksheetName "Control Summary"
                
                # Detailed results sheet with improved formatting
                $DetailedResults = $Results | ForEach-Object {
                    $_ | Add-Member -MemberType NoteProperty -Name "Result" -Value $(if ($_.Pass -eq $true) { "Pass" } else { "Fail" }) -Force
                    $_ | Select-Object -Property * -ExcludeProperty Pass
                }
                
                $DetailedResults | Export-Excel @ExcelParams -WorksheetName "Detailed Results"
                
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
