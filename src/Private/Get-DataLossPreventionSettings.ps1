function Get-DataLossPreventionSettings {
    <#
    .SYNOPSIS
        Retrieves Data Loss Prevention (DLP) settings and policies.
    
    .DESCRIPTION
        This function collects DLP compliance policies, rules, and custom sensitive information types
        from the compliance center for analysis and reporting.
    
    .PARAMETER Collection
        The collection hashtable to store the results.
    
    .PARAMETER LogFile
        The path to the log file for writing log entries.
    
    .EXAMPLE
        Get-DataLossPreventionSettings -Collection $Collection -LogFile $LogFile
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
