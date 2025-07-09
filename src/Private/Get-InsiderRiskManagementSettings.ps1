function Get-InsiderRiskManagementSettings {
    <#
    .SYNOPSIS
        Retrieves Insider Risk Management settings and policies.
    
    .DESCRIPTION
        This function collects Insider Risk Management policies from the compliance center
        for analysis and reporting.
    
    .PARAMETER Collection
        The collection hashtable to store the results.
    
    .PARAMETER LogFile
        The path to the log file for writing log entries.
    
    .EXAMPLE
        Get-InsiderRiskManagementSettings -Collection $Collection -LogFile $LogFile
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Collection,
        
        [Parameter(Mandatory = $true)]
        [string]$LogFile
    )
    
    try {
        [System.Collections.ArrayList]$WarnMessage = @()
        $Collection["InsiderRiskManagement"] = Get-InsiderRiskPolicy -ErrorAction:SilentlyContinue -WarningVariable +WarnMessage 
        $InfoMessage = "InsiderRiskManagement - Completed successfully!"
        Write-Log -IsInfo -InfoMessage $InfoMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
        Write-Host $InfoMessage
        Write-Log -IsWarn -WarnMessage $WarnMessage -LogFile $LogFile -ErrorAction:SilentlyContinue
    } catch {
        $Collection["InsiderRiskManagement"] = "Error"
        Write-Host "Error:$(Get-Date) There was an issue in fetching Insider Risk Management information. Please try running the tool again after some time." -ForegroundColor:Red
        $ErrorMessage = $_.ToString()
        $StackTraceInfo = $_.ScriptStackTrace
        Write-Log -IsError -ErrorMessage $ErrorMessage -StackTraceInfo $StackTraceInfo -LogFile $LogFile -ErrorAction:SilentlyContinue
    }
    
    return $Collection
}
