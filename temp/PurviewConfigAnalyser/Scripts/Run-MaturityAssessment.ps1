# Maturity Model Assessment Framework - Main Script
# This script orchestrates the complete assessment workflow:
# 1. Collect configuration data from Microsoft Purview
# 2. Run control book assessments based on specified configuration
# 3. Generate reports in CSV and Excel formats

param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigurationName = "PSPF",
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipDataCollection = $false,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateExcel
)

# Import required functions
. "$PSScriptRoot\..\Private\Test-ControlBook.ps1"

# Define paths based on configuration
$configBasePath = "$PSScriptRoot\..\config"
$outputPath = "$PSScriptRoot\..\output"
$dataCollectionScript = "$PSScriptRoot\..\Collect-PurviewConfiguration.ps1"

# Configuration-specific file paths
$controlConfigPath = "$configBasePath\ControlBook_${ConfigurationName}_Config.csv"
$propertyConfigPath = "$configBasePath\ControlBook_Property_${ConfigurationName}_Config.csv"
$runLogPath = "$outputPath\file_runlog.txt"

# Dynamic file names will be set after getting the OptimizedReport path
$resultsPath = $null
$excelReportPath = $null

# Function to get the latest OptimizedReport JSON file from run log
function Get-LatestOptimizedReport {
    param([string]$RunLogPath, [string]$OutputPath)
    
    if (Test-Path $RunLogPath) {
        $logEntries = Get-Content $RunLogPath | Where-Object { $_ -match "OptimizedReport.*\.json" }
        if ($logEntries) {
            $latestEntry = $logEntries[-1] # Get the last entry
            # Extract filename from log entry: "2025-07-09 12:47:20 - OptimizedReport: OptimizedReport_xxx.json"
            if ($latestEntry -match "OptimizedReport:\s+(OptimizedReport_.*\.json)") {
                $fileName = $matches[1]
                $fullPath = Join-Path $OutputPath $fileName
                if (Test-Path $fullPath) {
                    return $fullPath
                }
            }
        }
    }
    
    # Fallback: search for OptimizedReport*.json files directly
    $jsonFiles = Get-ChildItem -Path $OutputPath -Filter "OptimizedReport*.json" | Sort-Object LastWriteTime -Descending
    if ($jsonFiles) {
        return $jsonFiles[0].FullName
    }
    
    return $null
}

# Function to extract tenant ID from OptimizedReport filename and create dynamic result file names
function Set-DynamicResultPaths {
    param(
        [string]$OptimizedReportPath,
        [string]$ConfigurationName,
        [string]$OutputPath
    )
    
    # Extract tenant ID from OptimizedReport filename
    # Format: OptimizedReport_${TenantId}_$(timestamp).json
    $fileName = Split-Path -Leaf $OptimizedReportPath
    if ($fileName -match "OptimizedReport_([a-f0-9]+)_(\d{14})\.json") {
        $tenantId = $matches[1]
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        
        # Create dynamic file names following the same pattern
        $resultsFileName = "TestResults_${ConfigurationName}_${tenantId}_${timestamp}.csv"
            $excelFileName = "TestResults_${ConfigurationName}_${tenantId}_${timestamp}.xlsx"
        
        return @{
            ResultsPath = Join-Path $OutputPath $resultsFileName
            ExcelPath = Join-Path $OutputPath $excelFileName
            TenantId = $tenantId
        }
    } else {
        # Fallback to original naming if pattern doesn't match
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
            $resultsFileName = "TestResults_${ConfigurationName}_${timestamp}.csv"
            $excelFileName = "TestResults_${ConfigurationName}_${timestamp}.xlsx"
            return @{
                ResultsPath = Join-Path $OutputPath $resultsFileName
                ExcelPath = Join-Path $OutputPath $excelFileName
                TenantId = "unknown"
            }
    }
}

Write-Host "=== Maturity Model Assessment Framework ===" -ForegroundColor Cyan
Write-Host "Configuration: $ConfigurationName" -ForegroundColor White
Write-Host "Start Time: $(Get-Date)" -ForegroundColor White
Write-Host ""

# Step 1: Data Collection (if not skipped)
if (-not $SkipDataCollection) {
    Write-Host "Step 1: Collecting Microsoft Purview Configuration Data..." -ForegroundColor Yellow
    Write-Host "Running data collection script: $dataCollectionScript"
    
    try {
        & $dataCollectionScript
        
        # Get the latest OptimizedReport JSON file
        $optimizedReportPath = Get-LatestOptimizedReport -RunLogPath $runLogPath -OutputPath $outputPath
        
        if ($optimizedReportPath -and (Test-Path $optimizedReportPath)) {
            Write-Host "✅ Data collection completed successfully" -ForegroundColor Green
            $reportSize = (Get-Item $optimizedReportPath).Length / 1MB
            Write-Host "   Using OptimizedReport: $(Split-Path -Leaf $optimizedReportPath) ($([math]::Round($reportSize, 2)) MB)" -ForegroundColor Gray
            
            # Set dynamic result file paths based on OptimizedReport filename
            $filePaths = Set-DynamicResultPaths -OptimizedReportPath $optimizedReportPath -ConfigurationName $ConfigurationName -OutputPath $outputPath
            $resultsPath = $filePaths.ResultsPath
            $excelReportPath = $filePaths.ExcelPath
        } else {
            throw "OptimizedReport JSON file was not found"
        }
    }
    catch {
        Write-Host "❌ Data collection failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please ensure you have proper permissions and connectivity to Microsoft Purview" -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "Step 1: Skipping data collection (using existing OptimizedReport)" -ForegroundColor Yellow
    
    # Get the latest OptimizedReport JSON file
    $optimizedReportPath = Get-LatestOptimizedReport -RunLogPath $runLogPath -OutputPath $outputPath
    
    if ($optimizedReportPath -and (Test-Path $optimizedReportPath)) {
        $reportSize = (Get-Item $optimizedReportPath).Length / 1MB
        Write-Host "   Using OptimizedReport: $(Split-Path -Leaf $optimizedReportPath) ($([math]::Round($reportSize, 2)) MB)" -ForegroundColor Gray
        
        # Set dynamic result file paths based on OptimizedReport filename
        $filePaths = Set-DynamicResultPaths -OptimizedReportPath $optimizedReportPath -ConfigurationName $ConfigurationName -OutputPath $outputPath
        $resultsPath = $filePaths.ResultsPath
        $excelReportPath = $filePaths.ExcelPath
    } else {
        Write-Host "❌ No existing OptimizedReport JSON file found" -ForegroundColor Red
        Write-Host "Please run data collection first or check the run log at: $runLogPath" -ForegroundColor Yellow
        exit 1
    }
}

Write-Host ""

# Step 2: Validate Configuration Files
Write-Host "Step 2: Validating Configuration Files..." -ForegroundColor Yellow

if (-not (Test-Path $controlConfigPath)) {
    Write-Host "❌ Control configuration file not found: $controlConfigPath" -ForegroundColor Red
    Write-Host "Available configurations:" -ForegroundColor Yellow
    Get-ChildItem "$configBasePath\ControlBook_*_Config.csv" | ForEach-Object {
        $configName = $_.Name -replace "ControlBook_|_Config\.csv", ""
        Write-Host "  - $configName" -ForegroundColor Gray
    }
    exit 1
}

if (-not (Test-Path $propertyConfigPath)) {
    Write-Host "❌ Property configuration file not found: $propertyConfigPath" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Configuration files validated" -ForegroundColor Green
Write-Host "   Controls: $controlConfigPath" -ForegroundColor Gray
Write-Host "   Properties: $propertyConfigPath" -ForegroundColor Gray

Write-Host ""

# Step 3: Run Control Book Assessment
Write-Host "Step 3: Running Control Book Assessment..." -ForegroundColor Yellow

try {
    $assessmentResults = Test-ControlBook -ControlConfigPath $controlConfigPath -PropertyConfigPath $propertyConfigPath -OptimizedReportPath $optimizedReportPath -OutputPath $resultsPath -ConfigurationName $ConfigurationName
    
    Write-Host "✅ Assessment completed successfully" -ForegroundColor Green
    Write-Host "   Results saved to: $resultsPath" -ForegroundColor Gray
}
catch {
    Write-Host "❌ Assessment failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 4: Generate Excel Report (if requested)
if ($GenerateExcel) {
        # Always set Excel path to match CSV, just change extension to .xlsx
        $excelReportPath = [IO.Path]::ChangeExtension($resultsPath, '.xlsx')
        $excelDir = Split-Path $excelReportPath -Parent
        if (-not (Test-Path $excelDir)) {
            New-Item -Path $excelDir -ItemType Directory -Force | Out-Null
        }
        # Prepare Maturity Level Summary data - enriched with descriptions
        $maturitySummary = @()
        $maturityGroups = $assessmentResults.Results | Group-Object MaturityLevel
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
            
            $maturitySummary += [PSCustomObject]@{
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
    Write-Host "Step 4: Generating Excel Report..." -ForegroundColor Yellow
    try {
        # Check if ImportExcel module is available
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "⚠️ ImportExcel module not found. Installing..." -ForegroundColor Yellow
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        Import-Module ImportExcel -Force
        # Create Excel workbook with multiple sheets
        $excelParams = @{
            Path = $excelReportPath
            AutoSize = $true
            AutoFilter = $true
            BoldTopRow = $true
            FreezeTopRow = $true
        }
        # Maturity Level Summary sheet (now after $excelParams is set)
        $maturitySummary | Export-Excel @excelParams -WorksheetName 'Maturity Level Summary'
        # Summary sheet
        $uniqueMaturityLevels = ($assessmentResults.Results | Where-Object { $_.MaturityLevel -ne $null -and $_.MaturityLevel -ne '' } | Select-Object -ExpandProperty MaturityLevel -Unique)
        $numMaturityLevels = ($uniqueMaturityLevels | Measure-Object).Count
        $summary = [PSCustomObject]@{
            "Assessment Date" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            "Configuration" = $ConfigurationName
            "Total Controls" = $assessmentResults.TotalControls
            "Passing Controls" = $assessmentResults.PassingControls
            "Failing Controls" = $assessmentResults.FailingControls
            "Compliance Rate %" = $assessmentResults.ComplianceRate
            "Data Source" = $optimizedReportPath
            "Maturity Levels Present" = if ($numMaturityLevels -gt 1) { $uniqueMaturityLevels -join ", " } else { $uniqueMaturityLevels }
        }
        $summary | Export-Excel @excelParams -WorksheetName "Summary"
        
        # Control Summary sheet - aggregated by Control ID
        $controlSummary = $assessmentResults.Results | Group-Object -Property ControlID | ForEach-Object {
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
        $controlSummary | Export-Excel @excelParams -WorksheetName "Control Summary"
        
        # Detailed results sheet with improved formatting (ensure MaturityLevel column is present)
        $detailedResults = $assessmentResults.Results | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "Result" -Value $(if ($_.Pass -eq $true) { "Pass" } else { "Fail" }) -Force
            $_ | Select-Object -Property * -ExcludeProperty Pass
        }
        if (-not ($detailedResults | Get-Member -Name 'MaturityLevel')) {
            $detailedResults = $detailedResults | Select-Object *, @{Name='MaturityLevel';Expression={''}}
        }
        $detailedResults | Export-Excel @excelParams -WorksheetName "Detailed Results"
        Write-Host "✅ Excel report generated successfully" -ForegroundColor Green
        Write-Host "   Report saved to: $excelReportPath" -ForegroundColor Gray
    }
    catch {
        Write-Host "⚠️ Excel generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "   CSV results are still available at: $resultsPath" -ForegroundColor Gray
    }
} else {
    Write-Host "Step 4: Skipping Excel report generation" -ForegroundColor Yellow
}

Write-Host ""

# Summary
Write-Host "=== Assessment Complete ===" -ForegroundColor Cyan
Write-Host "Configuration: $ConfigurationName" -ForegroundColor White
Write-Host "Compliance Rate: $($assessmentResults.ComplianceRate)%" -ForegroundColor $(if ($assessmentResults.ComplianceRate -ge 80) { "Green" } elseif ($assessmentResults.ComplianceRate -ge 60) { "Yellow" } else { "Red" })
Write-Host "Total Controls: $($assessmentResults.TotalControls)" -ForegroundColor White
Write-Host "Passing: $($assessmentResults.PassingControls)" -ForegroundColor Green
Write-Host "Failing: $($assessmentResults.FailingControls)" -ForegroundColor Red
$uniqueMaturityLevels = ($assessmentResults.Results | Where-Object { $_.MaturityLevel -ne $null -and $_.MaturityLevel -ne '' } | Select-Object -ExpandProperty MaturityLevel -Unique)
$numMaturityLevels = ($uniqueMaturityLevels | Measure-Object).Count
if ($numMaturityLevels -gt 1) {
    Write-Host ("Maturity Levels Present: " + ($uniqueMaturityLevels -join ", ")) -ForegroundColor White
} elseif ($numMaturityLevels -eq 1) {
    Write-Host ("Maturity Level: " + $uniqueMaturityLevels) -ForegroundColor White
} else {
    Write-Host "Maturity Level: (none)" -ForegroundColor White
}
Write-Host ""
Write-Host "Output Files:" -ForegroundColor White
Write-Host "- CSV Results: $resultsPath" -ForegroundColor Gray
if ($GenerateExcel -and (Test-Path $excelReportPath)) {
    Write-Host "- Excel Report: $excelReportPath" -ForegroundColor Gray
}
Write-Host ""
Write-Host "End Time: $(Get-Date)" -ForegroundColor White

# Print Maturity Level Summary in terminal
$maturityGroups = $assessmentResults.Results | Group-Object MaturityLevel
Write-Host "" -ForegroundColor White
Write-Host "Maturity Level Summary:" -ForegroundColor Cyan
Write-Host ("{0,-15} {1,15} {2,18} {3,16} {4,18}" -f 'Maturity Level','Total Controls','Passing Controls','Failing Controls','Compliance Rate %') -ForegroundColor White
foreach ($group in $maturityGroups) {
    $level = if ($group.Name -eq '' -or $null -eq $group.Name) { '(none)' } else { $group.Name }
    $total = $group.Group.Count
    $passing = ($group.Group | Where-Object { $_.Pass -eq $true }).Count
    $failing = $total - $passing
    $rate = if ($total -gt 0) { [math]::Round(($passing / $total) * 100, 1) } else { 0 }
    Write-Host ("{0,-15} {1,15} {2,18} {3,16} {4,18}" -f $level, $total, $passing, $failing, $rate) -ForegroundColor Gray
}

# Return results for potential further processing
return $assessmentResults
