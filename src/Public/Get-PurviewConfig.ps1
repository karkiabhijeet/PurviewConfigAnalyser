function Get-PurviewConfig {
    <#
    .SYNOPSIS
        Collects Microsoft Purview configuration data.
    
    .DESCRIPTION
        Connects to Microsoft Purview and collects configuration data including
        sensitivity labels, policies, DLP settings, and compliance information.
    
    .PARAMETER OutputPath
        Specifies the output directory for generated files
    
    .PARAMETER TenantId
        Optional tenant ID for multi-tenant scenarios
    
    .EXAMPLE
        Get-PurviewConfig -OutputPath "C:\Reports"
        
        Collects Purview configuration data and saves to the specified path.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = (Join-Path $env:USERPROFILE "PurviewConfigAnalyser\Output"),
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    # Initialize logging
    $LogDirectory = "$env:LOCALAPPDATA\Microsoft\PurviewConfigAnalyser\Logs"
    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $FileName = "PurviewConfigAnalyser-$(Get-Date -Format 'yyyyMMddHHmmss').log"
    $LogFile = "$LogDirectory\$FileName"
    
    Write-Host "Starting Purview Configuration Data Collection..." -ForegroundColor Green
    
    try {
        # Step 1: Establish connection to the Compliance Center
        Connect-ToComplianceCenter
        
        # Step 2: Collect data
        $Collection = @{}
        $Collection = Get-InformationProtectionSettings -Collection $Collection -LogFile $LogFile
        $Collection = Get-RetentionCompliance -Collection $Collection -LogFile $LogFile
        $Collection = Get-DataLossPreventionSettings -Collection $Collection -LogFile $LogFile
        $Collection = Get-InsiderRiskManagementSettings -Collection $Collection -LogFile $LogFile
        $Collection = Get-TenantDetails -Collection $Collection -LogFile $LogFile
        
        # Extract TenantId for filename
        $TenantIdForFile = $Collection["TenantDetails"]["TenantId"] -replace "[^a-zA-Z0-9]", ""
        
        # Step 3: Generate output files
        $OutputFile = Join-Path -Path $OutputPath -ChildPath "OptimizedReport_${TenantIdForFile}_$(Get-Date -Format 'yyyyMMddHHmmss').json"
        $OutputExcelFile = Join-Path -Path $OutputPath -ChildPath "OptimizedReport_${TenantIdForFile}_$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"
        $RunLogFile = Join-Path -Path $OutputPath -ChildPath "file_runlog.txt"
        
        # Convert and save JSON
        Write-Host "Generating OptimizedReport.json..." -ForegroundColor Yellow
        $Collection = Convert-ObjectForJson -InputObject $Collection
        
        # Try to convert to JSON with manageable depth
        try {
            $Collection | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputFile -Encoding UTF8
            Write-Host "[SUCCESS] OptimizedReport.json generated successfully!" -ForegroundColor Green
            
            # Log the generated file
            $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport: $(Split-Path -Leaf $OutputFile)"
            Add-Content -Path $RunLogFile -Value $LogEntry
            
        } catch {
            Write-Host "[ERROR] JSON conversion failed with depth 10: $_" -ForegroundColor Red
            try {
                $Collection | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputFile -Encoding UTF8
                Write-Host "[SUCCESS] OptimizedReport.json generated with reduced depth!" -ForegroundColor Green
                
                $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport: $(Split-Path -Leaf $OutputFile)"
                Add-Content -Path $RunLogFile -Value $LogEntry
                
            } catch {
                Write-Host "[ERROR] JSON conversion failed with depth 5: $_" -ForegroundColor Red
                Write-Host "Creating minimal JSON for Excel processing..." -ForegroundColor Yellow
                
                $MinimalCollection = @{
                    TenantDetails = $Collection["TenantDetails"]
                    GetLabel = $Collection["GetLabel"]
                    GetLabelPolicy = $Collection["GetLabelPolicy"]
                    GetAutoSensitivityLabelPolicy = $Collection["GetAutoSensitivityLabelPolicy"]
                    GetDlpCompliancePolicy = $Collection["GetDlpCompliancePolicy"]
                    GetDlpComplianceRule = $Collection["GetDlpComplianceRule"]
                    GetComplianceTag = $Collection["GetComplianceTag"]
                    InsiderRiskManagement = $Collection["InsiderRiskManagement"]
                }
                $MinimalCollection | ConvertTo-Json -Depth 3 | Out-File -FilePath $OutputFile -Encoding UTF8
                Write-Host "[SUCCESS] Minimal OptimizedReport.json generated!" -ForegroundColor Green
                
                $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport (Minimal): $(Split-Path -Leaf $OutputFile)"
                Add-Content -Path $RunLogFile -Value $LogEntry
            }
        }
        
        # Generate Excel report
        Write-Host "Generating Excel report..." -ForegroundColor Yellow
        $JsonForExcel = Get-Content -Path $OutputFile -Raw | ConvertFrom-Json
        
        if ($JsonForExcel.TenantDetails -and $JsonForExcel.TenantDetails.TenantId -ne "Unknown") {
            $TenantDetails = @{
                TenantId          = $JsonForExcel.TenantDetails.TenantId
                Organization      = $JsonForExcel.TenantDetails.Organization
                UserPrincipalName = $JsonForExcel.TenantDetails.UserPrincipalName
                Timestamp         = $JsonForExcel.TenantDetails.Timestamp
            }
        } else {
            $TenantDetails = @{
                TenantId          = "Unknown"
                Organization      = "Unknown"
                UserPrincipalName = "Unknown"
                Timestamp         = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            }
        }
        
        GenerateExcelFromJSON -JsonFilePath $OutputFile -OutputExcelPath $OutputExcelFile -TenantDetails $TenantDetails
        
        # Log the generated Excel file
        $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Excel Report: $(Split-Path -Leaf $OutputExcelFile)"
        Add-Content -Path $RunLogFile -Value $LogEntry
        
        Write-Host "[SUCCESS] Data collection complete!" -ForegroundColor Green
        Write-Host "   JSON file: $OutputFile" -ForegroundColor Gray
        Write-Host "   Excel file: $OutputExcelFile" -ForegroundColor Gray
        
        return @{
            JsonFile = $OutputFile
            ExcelFile = $OutputExcelFile
            TenantId = $TenantIdForFile
            Collection = $Collection
        }
        
    } catch {
        Write-Host "[ERROR] Data collection failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
