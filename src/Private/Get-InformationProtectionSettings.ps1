function Get-InformationProtectionSettings {
    <#
    .SYNOPSIS
        Retrieves Information Protection settings including labels and policies.
    
    .DESCRIPTION
        This function collects sensitivity labels and label policies from the compliance center,
        and adds a Published attribute to labels based on their presence in active policies.
    
    .PARAMETER Collection
        The collection hashtable to store the results.
    
    .PARAMETER LogFile
        The path to the log file for writing log entries.
    
    .EXAMPLE
        Get-InformationProtectionSettings -Collection $Collection -LogFile $LogFile
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Collection,
        
        [Parameter(Mandatory = $true)]
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
