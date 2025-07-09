function ExtractColumns {
    <#
    .SYNOPSIS
        Extracts specific columns from JSON data and adds tenant details.
    
    .DESCRIPTION
        This function processes JSON data to extract specific columns and adds
        tenant details to each row for Excel generation.
    
    .PARAMETER Data
        The array of data to process.
    
    .PARAMETER Columns
        The columns to extract from the data.
    
    .PARAMETER TenantDetails
        The tenant details to add to each row.
    
    .EXAMPLE
        ExtractColumns -Data $JsonData.GetLabel -Columns $Columns -TenantDetails $TenantDetails
    #>
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

function ExtractLabelActions {
    <#
    .SYNOPSIS
        Extracts LabelActions data for Excel generation.
    
    .DESCRIPTION
        This function processes LabelActions data to create structured
        output for Excel worksheets.
    
    .PARAMETER LabelActions
        The LabelActions array to process.
    
    .PARAMETER TenantDetails
        Tenant details to add to each row.
    
    .PARAMETER LabelName
        Label Name to add to each row.
    
    .PARAMETER ParentLabelName
        Parent Label Name to add to each row.
    
    .PARAMETER ImmutableId
        Immutable ID to add to each row.
    
    .EXAMPLE
        ExtractLabelActions -LabelActions $Label.LabelActions -TenantDetails $TenantDetails -LabelName $Label.DisplayName -ParentLabelName $ParentLabelName -ImmutableId $Label.ImmutableId
    #>
    param (
        [Parameter(Mandatory = $true)]
        [array]$LabelActions,

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantDetails,

        [Parameter(Mandatory = $true)]
        [string]$LabelName,

        [Parameter(Mandatory = $false)]
        [string]$ParentLabelName = "N/A",

        [Parameter(Mandatory = $true)]
        [string]$ImmutableId
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

function GenerateExcelFromJSON {
    <#
    .SYNOPSIS
        Generates an Excel file from JSON data.
    
    .DESCRIPTION
        This function reads JSON data and creates an Excel file with multiple
        worksheets containing the structured data.
    
    .PARAMETER JsonFilePath
        Path to the JSON file to process.
    
    .PARAMETER OutputExcelPath
        Path to save the generated Excel file.
    
    .PARAMETER TenantDetails
        Tenant details to add to all tabs.
    
    .EXAMPLE
        GenerateExcelFromJSON -JsonFilePath $JsonFile -OutputExcelPath $ExcelFile -TenantDetails $TenantDetails
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$JsonFilePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputExcelPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$TenantDetails
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
                $FlattenedRow[$_.Name] = $_.Value -join ", "
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
                $LabelActionsData += ExtractLabelActions -LabelActions $Label.LabelActions -TenantDetails $TenantDetails -LabelName $Label.DisplayName -ParentLabelName $ParentLabelName -ImmutableId $Label.ImmutableId
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
        Write-Host "Excel file generated successfully at: $OutputExcelPath" -ForegroundColor Green
        
        # Log the success if LogFile is available
        if (Get-Variable -Name "LogFile" -ErrorAction SilentlyContinue) {
            $InfoMessage = "Excel file generated successfully at: $OutputExcelPath"
            Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        }
        
    } catch {
        Write-Host "Error: Unable to generate the Excel file. $_" -ForegroundColor Red
    }
}
