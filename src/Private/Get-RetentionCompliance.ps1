function Get-RetentionCompliance {
    <#
    .SYNOPSIS
        Retrieves retention compliance policies and tags.
    
    .DESCRIPTION
        This function collects retention compliance policies, rules, and compliance tags
        from the compliance center for analysis and reporting.
    
    .PARAMETER Collection
        The collection hashtable to store the results.
    
    .PARAMETER LogFile
        The path to the log file for writing log entries.
    
    .EXAMPLE
        Get-RetentionCompliance -Collection $Collection -LogFile $LogFile
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
