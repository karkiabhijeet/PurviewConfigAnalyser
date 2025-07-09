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
        $InfoMessage = "Module $ModuleName is not installed. Installing..."
        Write-Host "$(Get-Date) $InfoMessage"
     
        Install-Module -Name $ModuleName -Force -ErrorAction Stop
    }
    Import-Module -Name $ModuleName -ErrorAction Stop
    $InfoMessage = "Module $ModuleName is imported successfully."
    Write-Host "$(Get-Date) $InfoMessage"
    
        
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
$InfoMessage = "Log File Path : $LogFile"
Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
Write-Host $InfoMessage
$InfoMessage = "Main Directory Path : $env:LOCALAPPDATA\Microsoft\PurviewConfigAnalyser"
Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
Write-Host $InfoMessage
        

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
    Write-Host "Establishing connection to Microsoft Compliance Center..."
    try {
        # Prompt for credentials
        $userName = Read-Host -Prompt 'Enter your User Principal Name (UPN)'
        #EnsureComplianceCenterConnection -UserPrincipalName $userName
        # Connect to the Compliance Center using UserPrincipalName
        Connect-IPPSSession -UserPrincipalName $userName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        $InfoMessage = "Connection established successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
    } catch {
        Write-Host "Error establishing connection: $_"
        
        exit 1
    }
}

function Convert-HashtableKeysToString {
    param (
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )
    process {
        if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.IDictionary]) {
            $newHash = @{}
            foreach ($key in $InputObject.Keys) {
                $stringKey = [string]$key
                $newHash[$stringKey] = Convert-HashtableKeysToString $InputObject[$key]
            }
            return [PSCustomObject]$newHash
        } elseif ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
            return @($InputObject | ForEach-Object { Convert-HashtableKeysToString $_ })
        } elseif ($InputObject -is [PSCustomObject]) {
            $props = @{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                $props[$prop.Name] = Convert-HashtableKeysToString $prop.Value
            }
            return [PSCustomObject]$props
        } else {
            return $InputObject
        }
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
            # Extract all ImmutableIds from ScopedLabels in GetLabelPolicy
        # Extract all ImmutableIds from ScopedLabels in GetLabelPolicy and deduplicate
        $PublishedImmutableIds = @(
            $Collection["GetLabelPolicy"] |
            ForEach-Object { $_.ScopedLabels } |
            Where-Object { $_ -ne $null } |
            ForEach-Object { [string]$_ } |
            Select-Object -Unique
        )

# Debug: Log the deduplicated ImmutableIds
Write-Host "Published ImmutableIds: $($PublishedImmutableIds -join ', ')" -ForegroundColor Cyan

# Add Published attribute to each label in GetLabel
foreach ($Label in $Collection["GetLabel"]) {
    Write-Host "Processing Label: $($Label.DisplayName)" -ForegroundColor Yellow

    # Ensure ImmutableId is a string
    $ImmutableId = [string]$Label.ImmutableId
    Write-Host "ImmutableId (as string): $ImmutableId" -ForegroundColor Cyan

    # Add the Published property dynamically if it doesn't exist
    if (-not ($Label.PSObject.Properties | Where-Object { $_.Name -eq "Published" })) {
        $Label | Add-Member -MemberType NoteProperty -Name Published -Value $false -Force
    }

    # Check if ImmutableId is in PublishedImmutableIds
    if ($null -ne $ImmutableId -and $ImmutableId -ne "") {
        if ($PublishedImmutableIds -contains $ImmutableId) {
            Write-Host "Match Found: $ImmutableId is in PublishedImmutableIds" -ForegroundColor Green
            $Label.Published = $true
        } else {
            Write-Host "No Match: $ImmutableId is NOT in PublishedImmutableIds" -ForegroundColor Red
            $Label.Published = $false
        }
    } else {
        Write-Host "Warning: Label $($Label.DisplayName) has a null or empty ImmutableId." -ForegroundColor Red
        $Label.Published = $false
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

        # Debug: Check the output of Get-ConnectionInformation
        #Write-Host "OrgConfig Output:" -ForegroundColor Yellow
        #Write-Host ($OrgConfig | Out-String)

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
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction SilentlyContinue
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
        $Collection["GetDlpCompliancePolicy"] = Get-DlpCompliancePolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
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
        #$Collection["GetRetentionComplianceRule"] = Get-RetentionComplianceRule -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage
        $Collection["GetComplianceTag"] = Get-ComplianceTag -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $InfoMessage = "GetComplianceTag - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    catch {
        $Collection["GetRetentionCompliancePolicy"] = "Error"
        #$Collection["GetRetentionComplianceRule"] = "Error"
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

# Function to evaluate sensitivity labels
function EvaluateSensitivityLabels {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Labels
    )

    # Initialize the result collection
    $LabelTests = @()

    foreach ($Label in $Labels) {
        # Skip labels that are not published
        if (-not $Label.Published) {
            Write-Host "Skipping Label: $($Label.DisplayName) as it is not published." -ForegroundColor Yellow
            continue
        }

        $LabelName = $Label.DisplayName
        $IsPublished = $Label.Published -eq $true
        $TestResult = @{
            LabelName = $LabelName
            Published = $IsPublished
            TestUnofficial = "Fail"
            TestOfficial = "Fail"
            TestOfficialSensitive = "Fail"
        }

        # Check for exact matches
        if ($LabelName -eq "Unofficial") {
            $TestResult.TestUnofficial = "Pass"
        }
        if ($LabelName -eq "Official") {
            $TestResult.TestOfficial = "Pass"
        }
        if ($LabelName -eq "Official Sensitive") {
            $TestResult.TestOfficialSensitive = "Pass"
        }

        # Check for partial matches
        if ($LabelName -like "*Unofficial*" -and $TestResult.TestUnofficial -ne "Pass") {
            $TestResult.TestUnofficial = "Partial Pass"
        }
        if ($LabelName -like "*Official*" -and $LabelName -notlike "*Sensitive*" -and $LabelName -notlike "*Unofficial*" -and $TestResult.TestOfficial -ne "Pass") {
            $TestResult.TestOfficial = "Partial Pass"
        }
        if ($LabelName -like "*Official*Sensitive*" -and $TestResult.TestOfficialSensitive -ne "Pass") {
            $TestResult.TestOfficialSensitive = "Partial Pass"
        }

        # Add the test result to the collection
        $LabelTests += $TestResult
    }

    $InfoMessage = "Evaluation of Sensitivity Labels - Completed successfully!"
    Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    Write-Host $InfoMessage
    return $LabelTests
}

# Main script execution
Write-Host "Starting Simplified Purview Configuration Report Generation..."

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

# Step 3: Evaluate sensitivity labels
if ($Collection["GetLabel"] -ne "Error") {
    $SensitivityLabelTests = EvaluateSensitivityLabels -Labels $Collection["GetLabel"]
    $Collection["SensitivityLabelTests"] = $SensitivityLabelTests
}

# Step 3: Output directory and file
$OutputDir = "$env:LOCALAPPDATA\Microsoft\PurviewConfigAnalyser\RawData"
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
$OutputFile = Join-Path -Path $OutputDir -ChildPath "OptimizedReport_${TenantId}_$(Get-Date -Format 'yyyyMMddHHmmss').json"

# Step 4: Write report to JSON file
$Collection = Convert-HashtableKeysToString $Collection
$Collection | ConvertTo-Json -Depth 15 | Out-File -FilePath $OutputFile -Encoding UTF8


Write-Host "Optimized Purview Configuration Report generated successfully at: $OutputFile"

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

# Function to generate an Excel file from a JSON file
# Function to preprocess JSON and handle duplicate keys
function PreprocessJsonFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$JsonFilePath,  # Path to the JSON file

        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath # Path to save the processed JSON file
    )

    try {
        # Read the JSON file as a raw string
        $RawJson = Get-Content -Path $JsonFilePath -Raw

        # Handle duplicate keys by renaming 'Value' to 'Value_duplicate'
        $ProcessedJson = $RawJson -replace '"Value":', '"Value_duplicate":'

        # Save the processed JSON to a new file
        Set-Content -Path $OutputFilePath -Value $ProcessedJson -Encoding UTF8

        Write-Host "JSON file preprocessed successfully and saved to: $OutputFilePath" -ForegroundColor Green
    } catch {
        Write-Host "Error preprocessing JSON file: $_" -ForegroundColor Red
    }
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

    if ($JsonData.SensitivityLabelTests) {
        $Columns = @("LabelName", "TestUnofficial", "TestOfficialSensitive", "TestOfficial")
        $ExcelSheets += @{
            Name = "SensitivityLabelTests"
            Data = ExtractColumns -Data $JsonData.SensitivityLabelTests -Columns $Columns -TenantDetails $TenantDetails
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
        $Columns = @("ImmutableId", "Name", "DisplayName", "Priority", "ParentId", "ParentLabelDisplayName", "IsParent", "Tooltip", "ContentType", "Workload", "IsValid", "CreatedBy", "LastModifiedBy", "WhenCreated", "WhenCreatedUTC", "WhenChangedUTC", "OrganizationId", "LabelActions_Type", "Policy", "Published")
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
    }

    if ($JsonData.GetDlpCompliancePolicy) {
        $Columns = @("Name", "Guid", "DisplayName", "Type", "PolicyCategory", "IsSimulationPolicy", "SimulationStatus", "AutoEnableAfter", "Workload", "Comment", "Enabled", "CreationTimeUtc", "ModificationTimeUtc", "Mode")
        $ExcelSheets += @{
            Name = "GetDlpCompliancePolicy"
            Data = ExtractColumns -Data $JsonData.GetDlpCompliancePolicy -Columns $Columns -TenantDetails $TenantDetails
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
        $InfoMessage = "Excel file generated successfully at: $OutputExcelPath"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        
    } catch {
        Write-Host "Error: Unable to generate the Excel file. $_" -ForegroundColor Red
    }
}

# Example usage
$OriginalJsonFile = $OutputFile
$OutputExcelDir = "$env:LOCALAPPDATA\Microsoft\PurviewConfigAnalyser\ReportData"

if (-not (Test-Path -Path $OutputExcelDir)) {
    New-Item -ItemType Directory -Path $OutputExcelDir | Out-Null
}


$ProcessedJsonFile = Join-Path -Path $OutputDir -ChildPath "Preprocessed_OptimizedReport_${TenantId}_$(Get-Date -Format 'yyyyMMddHHmmss').JSON"
$OutputExcelFile = Join-Path -Path $OutputExcelDir -ChildPath "OptimizedReport_${TenantId}_$(Get-Date -Format 'yyyyMMddHHmmss').xlsx"

# Preprocess the JSON file to handle duplicate keys
PreprocessJsonFile -JsonFilePath $OriginalJsonFile -OutputFilePath $ProcessedJsonFile

# Extract TenantDetails for adding to all tabs
$TenantDetails = @{
    TenantId          = $Collection["TenantDetails"]["TenantId"]
    Organization      = $Collection["TenantDetails"]["Organization"]
    UserPrincipalName = $Collection["TenantDetails"]["UserPrincipalName"]
    Timestamp         = $Collection["TenantDetails"]["Timestamp"]
}

# Generate the Excel file from the processed JSON
GenerateExcelFromJSON -JsonFilePath $ProcessedJsonFile -OutputExcelPath $OutputExcelFile -TenantDetails $TenantDetails

# Delete the preprocessed JSON file
Remove-Item -Path $ProcessedJsonFile -Force
Write-Host "Preprocessed JSON file deleted: $ProcessedJsonFile" -ForegroundColor Green