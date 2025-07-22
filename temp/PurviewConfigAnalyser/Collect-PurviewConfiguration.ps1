# Optimized script to collect relevant data and output it into a concise JSON file and then Covert it to Excel.

# Function to establish a connection to the Microsoft Compliance Center
# Ensure required modules are installed and imported
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [Switch]$IsError = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$IsWarn = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$IsInfo = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$MachineInfo = $false,
        [Parameter(Mandatory = $false)]
        [Switch]$StopInfo = $false,
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage,
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$WarnMessage,
        [Parameter(Mandatory = $false)]
        [string]$InfoMessage,
        [Parameter(Mandatory = $false)]
        [string]$StackTraceInfo,
        [String]$LogFile
    )   

    if ($MachineInfo) {
        $ComputerInfoObj = Get-ComputerInfo 
        $CompName = $ComputerInfoObj.CsName
        $OSName = $ComputerInfoObj.OsName
        $OSVersion = $ComputerInfoObj.OsVersion
        $PowerShellVersion = $PSVersionTable.PSVersion
        try {
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Logging Started" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Start time: $(Get-Date)" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Computer Name: $CompName" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Operating System Name: $OSName" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Operating System Version: $OSVersion" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "PowerShell Version: $PowerShellVersion" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
         
        }
        catch {
            Write-Host "$(Get-Date) The local machine information cannot be logged." -ForegroundColor:Yellow
        }

    }
    if ($StopInfo) {
        try {
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "Logging Ended" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "End time: $(Get-Date)" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "********************************************************************************************" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            
            if ($($global:ErrorOccurred) -eq $true) {
                Write-Host "Warning:$(Get-Date) The report generated may have reduced information due to errors in running the tool. These errors may occur due to multiple reasons. Please refer documentation for more details." -ForegroundColor:Yellow
            }
         
        }
        catch {
            Write-Host "$(Get-Date) The finishing time information cannot be logged." -ForegroundColor:Yellow
        }
    }
    #Error
    if ($IsError) {
        if ($($global:ErrorOccurred) -eq $false) {
            $global:ErrorOccurred = $true
        }
        $Log_content = "$(Get-Date) ERROR: $ErrorMessage"
        try {
            $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            "TRACE: $StackTraceInfo" | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
        }
        catch {
            Write-Host "$(Get-Date) An error event cannot be logged." -ForegroundColor:Yellow  
        }           
    }
    #Warning
    if ($IsWarn) {
        foreach ($Warnmsg in $WarnMessage) {
            $Log_content = "$(Get-Date) WARN: $Warnmsg"
            try {
                $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
            }
            catch {
                Write-Host "$(Get-Date) A warning event cannot be logged." -ForegroundColor:Yellow 
            }
        }
    }
    #General
    if ($IsInfo) {
        $Log_content = "$(Get-Date) INFO: $InfoMessage"
        try {
            $Log_content | Out-File $LogFile -Append -ErrorAction:SilentlyContinue
        }
        catch {
            Write-Host "$(Get-Date) A general event cannot be logged." -ForegroundColor:Yellow 
        }
        
    }
}
function EnsureModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module: $ModuleName..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -ErrorAction Stop
    }
    Import-Module -Name $ModuleName -ErrorAction Stop
}

# Import required modules
EnsureModule -ModuleName ImportExcel
EnsureModule -ModuleName ExchangeOnlineManagement

$LogDirectory = "$env:LOCALAPPDATA\Microsoft\PurviewConfigAnalyser\Logs"
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$FileName = "PurviewConfigAnalyser-$(Get-Date -Format 'yyyyMMddHHmmss').log"
$LogFile = "$LogDirectory\$FileName"
Write-Log -IsInfo -InfoMessage "Log File Path : $LogFile" -LogFile $LogFile -ErrorAction:SilentlyContinue
        

function EnsureComplianceCenterConnection {
    param (
        [string]$UserPrincipalName
    )

    try {
        # Check if the session is still valid
        $Session = Get-PSSession | Where-Object { $_.Name -eq "ComplianceCenter" }
        if ($Session -and $Session.State -eq "Opened") {
            Write-Host "Compliance Center session is active." -ForegroundColor Green
        } else {
            Write-Host "Compliance Center session is expired or not established. Reconnecting..." -ForegroundColor Yellow
            Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ErrorAction Stop
            Write-Host "Reconnected to Compliance Center successfully!" -ForegroundColor Green
        }
    } catch {
        Write-Host "Error reconnecting to Compliance Center: $_" -ForegroundColor Red
        exit 1
    }
}

function Connect-ToComplianceCenter {
    Write-Host "Connecting to Microsoft Compliance Center..." -ForegroundColor Yellow
    try {
        # Prompt for credentials
        $userName = Read-Host -Prompt 'Enter your User Principal Name (UPN)'
        # Connect to the Compliance Center using UserPrincipalName
        Connect-IPPSSession -UserPrincipalName $userName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        $InfoMessage = "[SUCCESS] Connection established successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Connection failed: $_" -ForegroundColor Red
        exit 1
    }
}

function Convert-ObjectForJson {
    param (
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.IDictionary]) {
            $newHash = [ordered]@{}
            foreach ($key in $InputObject.Keys) {
                $stringKey = [string]$key
                $newHash[$stringKey] = Convert-ObjectForJson -InputObject $InputObject[$key]
            }
            return [PSCustomObject]$newHash
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $newList = [System.Collections.ArrayList]::new()
            foreach ($item in $InputObject) {
                $null = $newList.Add((Convert-ObjectForJson -InputObject $item))
            }
            return $newList
        }

        if ($InputObject -is [PSCustomObject]) {
            $newObj = [ordered]@{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                $propName = $prop.Name
                $propValue = $prop.Value

                # Targeted fix for arrays that get truncated
                if ($propName -in @("Labels", "ScopedLabels") -and $propValue -is [System.Collections.IEnumerable] -and $propValue -isnot [string]) {
                    $newObj[$propName] = @($propValue | ForEach-Object { "$_" })
                }
                # Preserve LabelActions, which are often pre-formatted JSON strings
                elseif ($propName -in @("LabelActions", "Settings", "LocaleSettings") -and $propValue -is [System.Collections.IEnumerable] -and $propValue -isnot [string]) {
                     $newObj[$propName] = $propValue # Preserve original structure
                }
                else {
                    $newObj[$propName] = Convert-ObjectForJson -InputObject $propValue
                }
            }
            return [PSCustomObject]$newObj
        }

        # For all other types, return as is
        return $InputObject
    }
}

# Function to filter and optimize data
function Optimize-Data {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,
        [array]$Fields
    )
    $OptimizedData = @()
    foreach ($Item in $Data) {
        $FilteredItem = @{}
        foreach ($Field in $Fields) {
            if ($Item.PSObject.Properties[$Field]) {
                $FilteredItem[$Field] = $Item.$Field
            }
        }
        $OptimizedData += [PSCustomObject]$FilteredItem
    }
    return $OptimizedData
}

# Function: Get-InformationProtectionSettings
Function Get-InformationProtectionSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetLabel"] = Get-Label -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetLabelPolicy"] = Get-LabelPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage

        # Add Published attribute to GetLabel data
        if ($Collection["GetLabel"] -ne "Error" -and $Collection["GetLabelPolicy"] -ne "Error") {
            # Extract all ImmutableIds from ScopedLabels in GetLabelPolicy and deduplicate
            $PublishedImmutableIds = @(
                $Collection["GetLabelPolicy"] |
                ForEach-Object { $_.ScopedLabels } |
                Where-Object { $_ -ne $null } |
                ForEach-Object { [string]$_ } |
                Select-Object -Unique
            )

            # Create a mapping of ImmutableId to Policy details
            $ImmutableIdToPolicyMap = @{}
            foreach ($Policy in $Collection["GetLabelPolicy"]) {
                if ($Policy.ScopedLabels) {
                    foreach ($ScopedLabel in $Policy.ScopedLabels) {
                        $ScopedLabelId = [string]$ScopedLabel
                        if (-not $ImmutableIdToPolicyMap.ContainsKey($ScopedLabelId)) {
                            $ImmutableIdToPolicyMap[$ScopedLabelId] = @()
                        }
                        # Add policy identifier (prefer Guid, fallback to Name)
                        $PolicyId = if ($Policy.Guid) { $Policy.Guid } else { $Policy.Name }
                        $ImmutableIdToPolicyMap[$ScopedLabelId] += $PolicyId
                    }
                }
            }

            # Add Published attribute to each label in GetLabel
            foreach ($Label in $Collection["GetLabel"]) {
                # Ensure ImmutableId is a string
                $ImmutableId = [string]$Label.ImmutableId

                # Add the Published property dynamically if it doesn't exist
                if (-not ($Label.PSObject.Properties | Where-Object { $_.Name -eq "Published" })) {
                    $Label | Add-Member -MemberType NoteProperty -Name Published -Value $false -Force
                }

                # Add the PublishedPolicy property dynamically if it doesn't exist
                if (-not ($Label.PSObject.Properties | Where-Object { $_.Name -eq "PublishedPolicy" })) {
                    $Label | Add-Member -MemberType NoteProperty -Name PublishedPolicy -Value $null -Force
                }

                # Check if ImmutableId is in PublishedImmutableIds
                if ($null -ne $ImmutableId -and $ImmutableId -ne "") {
                    if ($PublishedImmutableIds -contains $ImmutableId) {
                        $Label.Published = $true
                        # Set the PublishedPolicy to the policy(ies) that contain this label
                        if ($ImmutableIdToPolicyMap.ContainsKey($ImmutableId)) {
                            $PolicyIds = $ImmutableIdToPolicyMap[$ImmutableId]
                            $Label.PublishedPolicy = if ($PolicyIds.Count -eq 1) { $PolicyIds[0] } else { $PolicyIds -join ", " }
                        }
                    } else {
                        $Label.Published = $false
                        $Label.PublishedPolicy = $null
                    }
                } else {
                    $Label.Published = $false
                    $Label.PublishedPolicy = $null
                }
            }
            
        }

        $InfoMessage = "Get-InformationProtectionSettings - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetLabel"] = "Error"
        $Collection["GetLabelPolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Information Protection information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetAutoSensitivityLabelPolicy"] = Get-AutoSensitivityLabelPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage -ForceValidate
        $InfoMessage = "GetAutoSensitivityLabelPolicy - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetAutoSensitivityLabelPolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching AutoSensitivity Label Policy information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }

    Return $Collection
}

# Function to fetch Tenant ID, Tenant Name, and Current Timestamp
function Get-TenantDetails {
    param (
        $Collection,
        [string]$LogFile
    )
    try {
        # Fetch all tenant details using Get-ConnectionInformation
        [System.Collections.ArrayList]$WarnMessage = @()
        $OrgConfig = Get-ConnectionInformation -ErrorAction Stop 

        # Filter records where TokenStatus is Active
        $ActiveRecords = $OrgConfig | Where-Object { $_.TokenStatus -eq "Active" }

        # Check if there are any active records
        if ($ActiveRecords.Count -eq 0) {
            Write-Host "No active records found in OrgConfig!" -ForegroundColor Red
            $Collection["TenantDetails"] = "No Active Records"
        } else {
            # Select required attributes from the active records
            $SelectedRecord = $ActiveRecords | Select-Object -First 1 TenantID, Organization, State, UserPrincipalName

            # Create a collection with Tenant ID, Tenant Name, and Timestamp
            $Collection["TenantDetails"] = @{
                TenantId          = $SelectedRecord.TenantID
                Organization      = $SelectedRecord.Organization
                Timestamp         = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                State             = $SelectedRecord.State
                UserPrincipalName = $SelectedRecord.UserPrincipalName
            }
            $InfoMessage = "Tenant Details fetched successfully!"
            Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
            Write-Host $InfoMessage
            
        }

        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction SilentlyContinue
    } catch {
        $Collection["TenantDetails"] = "Error"
        Write-Host "Error: $(Get-Date) There was an issue in fetching Tenant Details. Please try running the tool again after some time." -ForegroundColor Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    return $Collection
}
# Get DLP settings
Function Get-DataLossPreventionSettings {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetDlpComplianceRule"] = Get-DlpComplianceRule -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $Collection["GetDLPCustomSIT"] = Get-DlpSensitiveInformationType -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage | Where-Object { $_.Publisher -ne "Microsoft Corporation" } 
        $Collection["GetDlpCompliancePolicy"] = Get-DlpCompliancePolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage -ForceValidate
        $InfoMessage = "GetDlpCompliancePolicy - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {        
        $Collection["GetDlpComplianceRule"] = "Error"
        $Collection["GetDLPCustomSIT"] = "Error"
        $Collection["GetDlpCompliancePolicy"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Data Loss Prevention information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }

    Return $Collection
}
Function Get-RetentionCompliance {
    Param(
        $Collection,
        [string]$LogFile
    )
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["GetRetentionCompliancePolicy"] = Get-RetentionCompliancePolicy -DistributionDetail -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $InfoMessage = "GetRetentionCompliancePolicy - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        $Collection["GetRetentionComplianceRule"] = Get-RetentionComplianceRule -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetComplianceTag"] = Get-ComplianceTag -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $InfoMessage = "GetComplianceTag - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetRetentionCompliancePolicy"] = "Error"
        $Collection["GetRetentionComplianceRule"] = "Error"
        $Collection["GetComplianceTag"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Retention Compliance information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
         
    }
    Return $Collection
}

# Function: Get-InsiderRiskManagementSettings
function Get-InsiderRiskManagementSettings {
    param (
        [hashtable]$Collection
    )
    try {
        $Collection["InsiderRiskManagement"] = Get-InsiderRiskPolicy  -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $InfoMessage = "InsiderRiskManagement - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        #$IRMPolicies = Get-InsiderRiskPolicy  -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        #$Collection["InsiderRiskManagement"] = Optimize-Data -Data $IRMPolicies -Fields @("Name", "Description", "Enabled", "Scope", "Conditions", "Actions")
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    } catch {
        $Collection["InsiderRiskManagement"]= "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Insider Risk Management information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
        
    }
    return $Collection
}

# Main script execution
Write-Host "Starting Purview Configuration Data Collection..." -ForegroundColor Green

# Step 1: Establish connection to the Compliance Center
Connect-ToComplianceCenter

# Step 2: Collect data
$Collection = @{}
$Collection = Get-InformationProtectionSettings  -Collection $Collection -LogFile $LogFile
$Collection = Get-RetentionCompliance -Collection $Collection
$Collection = Get-DataLossPreventionSettings  -Collection $Collection -LogFile $LogFile
$Collection = Get-InsiderRiskManagementSettings -Collection $Collection -LogFile $LogFile
$Collection = Get-TenantDetails -Collection $Collection -LogFile $LogFile
$TenantId = $Collection["TenantDetails"]["TenantId"] -replace "[^a-zA-Z0-9]", "" # Remove special characters

# Step 3: Output directory and file - Version-specific output directory within the module version
$OutputDir = Join-Path -Path $PSScriptRoot -ChildPath "output"
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
$OutputFile = Join-Path -Path $OutputDir -ChildPath "OptimizedReport_${TenantId}_$(Get-Date -Format 'yyyyMMddHHmmss').json"
$RunLogFile = Join-Path -Path $OutputDir -ChildPath "file_runlog.txt"

# Step 4: Write raw data report to JSON file

# First, preprocess the entire collection to ensure it's ready for JSON serialization
Write-Host "Generating OptimizedReport.json..." -ForegroundColor Yellow
$Collection = Convert-ObjectForJson -InputObject $Collection

# Now, try to convert to JSON with maximum depth
try {
    # Try with depth 10 first (much more manageable file size)
    $Collection | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-Host "[SUCCESS] OptimizedReport.json generated successfully!" -ForegroundColor Green
    
    # Log the generated file
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport: $(Split-Path -Leaf $OutputFile)"
    Add-Content -Path $RunLogFile -Value $LogEntry
    
} catch {
    Write-Host "[ERROR] JSON conversion failed with depth 10: $_" -ForegroundColor Red
    try {
        # Try with depth 5 as fallback
        Write-Host "Trying with reduced depth (5)..." -ForegroundColor Yellow
        $Collection | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "[SUCCESS] OptimizedReport.json generated successfully with reduced depth!" -ForegroundColor Green
        
        # Log the generated file
        $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport: $(Split-Path -Leaf $OutputFile)"
        Add-Content -Path $RunLogFile -Value $LogEntry
        
    } catch {
        Write-Host "[ERROR] JSON conversion failed with depth 5: $_" -ForegroundColor Red
        Write-Host "Creating minimal JSON for Excel processing..." -ForegroundColor Yellow
        # Create a minimal JSON for Excel processing
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
        Write-Host "[SUCCESS] Minimal OptimizedReport.json generated for Excel processing!" -ForegroundColor Green
        
        # Log the generated file
        $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - OptimizedReport (Minimal): $(Split-Path -Leaf $OutputFile)"
        Add-Content -Path $RunLogFile -Value $LogEntry
    }
}

Write-Host "[SUCCESS] Data collection complete!" -ForegroundColor Green
Write-Host "   OptimizedReport.json: $OutputFile" -ForegroundColor Gray

# Import the required module for Excel manipulation
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module -Name ImportExcel -ErrorAction Stop

# Function to extract specific columns from JSON data
function ExtractColumns {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,

        [Parameter(Mandatory = $true)]
        [array]$Columns
    )

    $ExtractedData = @()
    foreach ($Item in $Data) {
        $Row = @{}
        foreach ($Column in $Columns) {
            if ($Column -eq "LabelActions_Type" -and $Item.LabelActions) {
                # Special handling for LabelActions_Type
                $LabelActions = $Item.LabelActions | ForEach-Object {
                    ($_ | ConvertFrom-Json).Type
                }
                $Row[$Column] = ($LabelActions -join ", ")
            } else {
                $Row[$Column] = $Item.$Column
            }
        }
        $ExtractedData += [PSCustomObject]$Row
    }
    return $ExtractedData
}

# Function to extract specific columns from JSON data and add TenantDetails
function ExtractColumns {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Data,

        [Parameter(Mandatory = $true)]
        [array]$Columns,

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantDetails
    )

    $ExtractedData = @()
    foreach ($Item in $Data) {
        $Row = @{}

        # Add TenantDetails to each row
        foreach ($TenantKey in $TenantDetails.Keys) {
            $Row[$TenantKey] = $TenantDetails[$TenantKey]
        }

        # Add specified columns
        foreach ($Column in $Columns) {
            if ($Column -eq "LabelActions_Type" -and $Item.LabelActions) {
                # Special handling for LabelActions_Type
                $LabelActions = $Item.LabelActions | ForEach-Object {
                    ($_ | ConvertFrom-Json).Type
                }
                $Row[$Column] = ($LabelActions -join ", ")
            } else {
                $Row[$Column] = $Item.$Column
            }
        }
        $ExtractedData += [PSCustomObject]$Row
    }
    return $ExtractedData
}

# Function to dynamically extract LabelActions and create a new tab
function ExtractLabelActions {
    param (
        [Parameter(Mandatory = $true)]
        [array]$LabelActions,  # The LabelActions array to process

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantDetails,  # Tenant details to add to each row

        [Parameter(Mandatory = $true)]
        [string]$LabelName,  # Label Name to add to each row

        [Parameter(Mandatory = $false)]
        [string]$ParentLabelName = "N/A",  # Parent Label Name to add to each row

        [Parameter(Mandatory = $true)]
        [string]$ImmutableId  # Immutable ID to add to each row
    )

    $ExtractedData = @()

    foreach ($Action in $LabelActions) {
        try {
            # Parse the LabelAction JSON string
            $ParsedAction = $Action | ConvertFrom-Json

            # Extract Type and SubType
            $Type = $ParsedAction.Type
            $SubType = $ParsedAction.SubType

            # Process Settings array
            foreach ($Setting in $ParsedAction.Settings) {
                $Row = @{
                    # Add Tenant Details
                    TenantId          = $TenantDetails["TenantId"]
                    Organization      = $TenantDetails["Organization"]
                    UserPrincipalName = $TenantDetails["UserPrincipalName"]
                    Timestamp         = $TenantDetails["Timestamp"]

                    # Add Label Details
                    LabelName         = $LabelName
                    ImmutableId       = $ImmutableId
                    ParentLabelName  = $ParentLabelName

                    # Add LabelActions Details
                    Type              = $Type
                    SubType           = $SubType
                    SettingsKey       = $Setting.Key
                    SettingsValue     = $Setting.Value
                    SettingsValueIdentity = $null
                    SettingsValueRights   = $null
                }

                # If the Value contains JSON (e.g., rightsdefinitions), parse it further
                if ($Setting.Value -is [string] -and $Setting.Value.StartsWith("[") -and $Setting.Value.EndsWith("]")) {
                    $ParsedValue = $Setting.Value | ConvertFrom-Json
                    foreach ($Item in $ParsedValue) {
                        $Row.SettingsValueIdentity = $Item.Identity
                        $Row.SettingsValueRights = $Item.Rights
                        $ExtractedData += [PSCustomObject]$Row
                    }
                } else {
                    # Add the row as-is if no further parsing is needed
                    $ExtractedData += [PSCustomObject]$Row
                }
            }
        } catch {
            Write-Host "Error processing LabelAction: $_" -ForegroundColor Red
        }
    }

    return $ExtractedData
}

# Function to generate an Excel file from a JSON file
function GenerateExcelFromJSON {
    param (
        [Parameter(Mandatory = $true)]
        [string]$JsonFilePath,  # Path to the JSON file

        [Parameter(Mandatory = $true)]
        [string]$OutputExcelPath, # Path to save the generated Excel file

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantDetails # Tenant details to add to all tabs
    )

    # Read and parse the JSON file
    try {
        $JsonData = Get-Content -Path $JsonFilePath -Raw | ConvertFrom-Json
    } catch {
        Write-Host "Error: Unable to read or parse the JSON file. $_" -ForegroundColor Red
        return
    }

    # Initialize an array to store Excel sheet data
    $ExcelSheets = @()

    # Process each section based on the specified columns
    if ($JsonData.TenantDetails) {
        $Columns = @("TenantId", "Organization", "UserPrincipalName", "Timestamp")
        $ExcelSheets += @{
            Name = "TenantDetails"
            Data = ExtractColumns -Data @($JsonData.TenantDetails) -Columns $Columns -TenantDetails $TenantDetails
        }
    }

    if ($JsonData.GetLabelPolicy) {
        # Flatten the GetLabelPolicy collection
        $FlattenedData = $JsonData.GetLabelPolicy | ForEach-Object {
            $FlattenedRow = @{}
            $_.PSObject.Properties | ForEach-Object {
                $FlattenedRow[$_.Name] = $_.Value -join ", "  # Flatten arrays into comma-separated strings
            }
            [PSCustomObject]$FlattenedRow
        }

        $ExcelSheets += @{
            Name = "GetLabelPolicy"
            Data = $FlattenedData
        }
    }
    if ($JsonData.GetLabel) {
        $Columns = @("ImmutableId", "Name", "DisplayName", "Priority", "ParentId", "ParentLabelDisplayName", "IsParent", "Tooltip", "ContentType", "Workload", "IsValid", "CreatedBy", "LastModifiedBy", "WhenCreated", "WhenCreatedUTC", "WhenChangedUTC", "OrganizationId", "LabelActions_Type", "Policy", "Published", "PublishedPolicy")
        $ExcelSheets += @{
            Name = "GetLabel"
            Data = ExtractColumns -Data $JsonData.GetLabel -Columns $Columns -TenantDetails $TenantDetails
        }
    
        # Extract LabelActions and create a new tab
        $LabelActionsData = @()
        foreach ($Label in $JsonData.GetLabel) {
            if ($Label.LabelActions) {
                $ParentLabelName = if ([string]::IsNullOrWhiteSpace($Label.ParentLabelDisplayName)) { "N/A" } else { $Label.ParentLabelDisplayName }
                $LabelActionsData += ExtractLabelActions -LabelActions $Label.LabelActions `
                    -TenantDetails $TenantDetails `
                    -LabelName $Label.DisplayName `
                    -ParentLabelName $ParentLabelName  `
                    -ImmutableId $Label.ImmutableId
            }
        }
    
        if ($LabelActionsData.Count -gt 0) {
            $ExcelSheets += @{
                Name = "LabelActions"
                Data = $LabelActionsData
            }
        }
    }


    if ($JsonData.GetAutoSensitivityLabelPolicy) {
        $Columns = @("Type", "Name", "Guid", "LabelDisplayName", "ApplySensitivityLabel", "OverwriteLabel", "Mode", "Comment", "Workload", "CreatedBy", "LastModifiedBy", "ModificationTimeUtc", "CreationTimeUtc")
        $ExcelSheets += @{
            Name = "GetAutoSensitivityLabelPolicy"
            Data = ExtractColumns -Data $JsonData.GetAutoSensitivityLabelPolicy -Columns $Columns -TenantDetails $TenantDetails
        }
    }    if ($JsonData.GetDlpCompliancePolicy) {
        $Columns = @("Name", "Guid", "DisplayName", "Type", "PolicyCategory", "IsSimulationPolicy", "SimulationStatus", "AutoEnableAfter", "Workload", "Comment", "Enabled", "CreationTimeUtc", "ModificationTimeUtc", "Mode")
        $ExcelSheets += @{
            Name = "GetDlpCompliancePolicy"
            Data = ExtractColumns -Data $JsonData.GetDlpCompliancePolicy -Columns $Columns -TenantDetails $TenantDetails
        }
    }

    if ($JsonData.GetDlpComplianceRule) {
        $Columns = @("Name", "Guid", "DisplayName", "RulePriority", "Comment", "Enabled", "SensitiveInformationType", "AccessScopeIs", "BlockAccess", "BlockAccessScope", "NotifyUser", "NotifyPolicyTipDisplayOption", "EncryptionRMSTemplate", "CreationTimeUtc", "ModificationTimeUtc")
        $ExcelSheets += @{
            Name = "GetDlpComplianceRule"
            Data = ExtractColumns -Data $JsonData.GetDlpComplianceRule -Columns $Columns -TenantDetails $TenantDetails
        }
    }

    if ($JsonData.GetComplianceTag) {
        $Columns = @("Guid", "Name", "RetentionAction", "RetentionType", "AutoApprovalPeriod", "IsRecordLabel", "HasRetentionAction", "ComplianceTagType", "Workload", "Policy", "CreatedBy", "LastModifiedBy")
        $ExcelSheets += @{
            Name = "GetComplianceTag"
            Data = ExtractColumns -Data $JsonData.GetComplianceTag -Columns $Columns -TenantDetails $TenantDetails
        }
    }

    if ($JsonData.InsiderRiskManagement) {
        $Columns = @("Name", "Enabled", "Description", "Actions", "Scope", "Conditions")
        $ExcelSheets += @{
            Name = "InsiderRiskManagement"
            Data = ExtractColumns -Data $JsonData.InsiderRiskManagement -Columns $Columns -TenantDetails $TenantDetails
        }
    }

    # Generate the Excel file
    try {
        $ExcelSheets | ForEach-Object {
            Export-Excel -Path $OutputExcelPath -WorksheetName $_.Name -AutoSize -ClearSheet -InputObject $_.Data
        }
        $InfoMessage = "[SUCCESS] Excel file generated successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        
    } catch {
        Write-Host "Error: Unable to generate the Excel file. $_" -ForegroundColor Red
    }
}

# Example usage
$OriginalJsonFile = $OutputFile
$OutputExcelDir = Join-Path -Path $PSScriptRoot -ChildPath "..\output"

if (-not (Test-Path -Path $OutputExcelDir)) {
    New-Item -ItemType Directory -Path $OutputExcelDir | Out-Null
}


$OutputExcelFile = Join-Path -Path $OutputExcelDir -ChildPath "OptimizedReport_${TenantId}_$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"

# Extract TenantDetails for adding to all tabs
$JsonForExcel = Get-Content -Path $OriginalJsonFile -Raw | ConvertFrom-Json
if ($JsonForExcel.TenantDetails -and $JsonForExcel.TenantDetails.TenantId -ne "Unknown") {
    $TenantDetails = @{
        TenantId          = $JsonForExcel.TenantDetails.TenantId
        Organization      = $JsonForExcel.TenantDetails.Organization
        UserPrincipalName = $JsonForExcel.TenantDetails.UserPrincipalName
        Timestamp         = $JsonForExcel.TenantDetails.Timestamp
    }
} else {
    # Fallback tenant details
    $TenantDetails = @{
        TenantId          = "Unknown"
        Organization      = "Unknown"
        UserPrincipalName = "Unknown"
        Timestamp         = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    }
}

# Generate the Excel file from the processed JSON
if (Test-Path -Path $OriginalJsonFile) {
    GenerateExcelFromJSON -JsonFilePath $OriginalJsonFile -OutputExcelPath $OutputExcelFile -TenantDetails $TenantDetails
    
    Write-Host "[SUCCESS] Excel report generated!" -ForegroundColor Green
    Write-Host "   Excel file: $OutputExcelFile" -ForegroundColor Gray
    
    # Log the generated Excel file
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Excel Report: $(Split-Path -Leaf $OutputExcelFile)"
    Add-Content -Path $RunLogFile -Value $LogEntry
    
} else {
    Write-Host "[ERROR] JSON file not found. Cannot generate Excel report." -ForegroundColor Red
    Write-Host "   Expected JSON file: $OriginalJsonFile" -ForegroundColor Gray
}