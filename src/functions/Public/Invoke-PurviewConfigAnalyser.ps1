function Invoke-PurviewConfigAnalyser {
    <#
    .SYNOPSIS
        Main command for Microsoft Purview configuration analysis and compliance assessment.
    
    .DESCRIPTION
        This command provides a unified interface for collecting Microsoft Purview configuration data,
        running compliance assessments, and generating reports. It supports different modes of operation
        and custom control book configurations.
    
    .PARAMETER Mode
        Specifies the operation mode:
        - CollectAndTest: Collect configuration data and run compliance assessment
        - CollectOnly: Only collect configuration data
        - TestOnly: Only run compliance assessment using existing data
    
    .PARAMETER Configuration
        Specifies the control book configuration to use (default: PSPF)
    
    .PARAMETER OutputPath
        Specifies the output directory for generated files
    
    .PARAMETER GenerateExcel
        Generate Excel reports in addition to CSV files
    
    .PARAMETER TenantId
        Optional tenant ID for multi-tenant scenarios
    
    .EXAMPLE
        Invoke-PurviewConfigAnalyser -Mode CollectAndTest -Configuration PSPF -GenerateExcel
        
        Collects Purview configuration data, runs PSPF compliance assessment, and generates Excel reports.
    
    .EXAMPLE
        Invoke-PurviewConfigAnalyser -Mode CollectOnly -OutputPath "C:\Reports"
        
        Only collects configuration data and saves to the specified output path.
    
    .EXAMPLE
        Invoke-PurviewConfigAnalyser -Mode TestOnly -Configuration PSPF
        
        Runs compliance assessment using the latest collected data.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('CollectAndTest', 'CollectOnly', 'TestOnly')]
        [string]$Mode,
        
        [Parameter(Mandatory = $false)]
        [string]$Configuration = 'PSPF',
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = (Join-Path $env:USERPROFILE "PurviewConfigAnalyser\Output"),
        
        [Parameter(Mandatory = $false)]
        [switch]$GenerateExcel,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    # Initialize module paths
    $ModuleRoot = $PSScriptRoot | Split-Path -Parent
    $ConfigPath = Join-Path $ModuleRoot "config"
    $MasterControlBooksPath = Join-Path $ConfigPath "MasterControlBooks"
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    # Create run log file path
    $RunLogPath = Join-Path $OutputPath "file_runlog.txt"
    
    Write-Host "=== Microsoft Purview Configuration Analyser ===" -ForegroundColor Cyan
    Write-Host "Mode: $Mode" -ForegroundColor White
    Write-Host "Configuration: $Configuration" -ForegroundColor White
    Write-Host "Output Path: $OutputPath" -ForegroundColor White
    Write-Host "Start Time: $(Get-Date)" -ForegroundColor White
    Write-Host ""
    
    $results = @{}
    
    try {
        # Step 1: Data Collection (if required)
        if ($Mode -in @('CollectAndTest', 'CollectOnly')) {
            Write-Host "Step 1: Collecting Microsoft Purview Configuration Data..." -ForegroundColor Yellow
            $configData = Get-PurviewConfig -OutputPath $OutputPath -TenantId $TenantId
            
            if ($configData) {
                $results.DataCollection = $configData
                Write-Host "✅ Data collection completed successfully" -ForegroundColor Green
                
                # Log the generated files
                $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Data Collection Mode: $Mode"
                Add-Content -Path $RunLogPath -Value $logEntry
            } else {
                throw "Data collection failed"
            }
        }
        
        # Step 2: Compliance Assessment (if required)
        if ($Mode -in @('CollectAndTest', 'TestOnly')) {
            Write-Host "Step 2: Running Compliance Assessment..." -ForegroundColor Yellow
            
            # Get the latest OptimizedReport JSON file
            $latestReport = Get-LatestOptimizedReport -OutputPath $OutputPath
            
            if (-not $latestReport) {
                throw "No OptimizedReport JSON file found. Please run data collection first."
            }
            
            $assessmentResults = Test-PurviewCompliance -OptimizedReportPath $latestReport -Configuration $Configuration -OutputPath $OutputPath -GenerateExcel:$GenerateExcel
            
            if ($assessmentResults) {
                $results.Assessment = $assessmentResults
                Write-Host "✅ Compliance assessment completed successfully" -ForegroundColor Green
                
                # Log the assessment results
                $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Assessment completed - Configuration: $Configuration - Compliance: $($assessmentResults.ComplianceRate)%"
                Add-Content -Path $RunLogPath -Value $logEntry
            } else {
                throw "Compliance assessment failed"
            }
        }
        
        # Summary
        Write-Host ""
        Write-Host "=== Operation Complete ===" -ForegroundColor Cyan
        Write-Host "Mode: $Mode" -ForegroundColor White
        Write-Host "Configuration: $Configuration" -ForegroundColor White
        
        if ($results.Assessment) {
            Write-Host "Compliance Rate: $($results.Assessment.ComplianceRate)%" -ForegroundColor $(
                if ($results.Assessment.ComplianceRate -ge 80) { "Green" } 
                elseif ($results.Assessment.ComplianceRate -ge 60) { "Yellow" } 
                else { "Red" }
            )
        }
        
        Write-Host "End Time: $(Get-Date)" -ForegroundColor White
        Write-Host ""
        
        return $results
        
    } catch {
        Write-Host "❌ Operation failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Helper function to get latest OptimizedReport JSON file
function Get-LatestOptimizedReport {
    param([string]$OutputPath)
    
    $runLogPath = Join-Path $OutputPath "file_runlog.txt"
    
    if (Test-Path $runLogPath) {
        $logEntries = Get-Content $runLogPath | Where-Object { $_ -match "OptimizedReport.*\.json" }
        if ($logEntries) {
            $latestEntry = $logEntries[-1]
            # Extract filename using regex - look for .json files
            if ($latestEntry -match "OptimizedReport[^:]*\.json") {
                $fileName = $matches[0]
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
