function Get-TenantDetails {
    <#
    .SYNOPSIS
        Retrieves tenant details including TenantId, Organization, and current timestamp.
    
    .DESCRIPTION
        This function fetches tenant information from the compliance center connection
        and extracts relevant details for reporting purposes.
    
    .PARAMETER Collection
        The collection hashtable to store the results.
    
    .PARAMETER LogFile
        The path to the log file for writing log entries.
    
    .EXAMPLE
        Get-TenantDetails -Collection $Collection -LogFile $LogFile
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Collection,
        
        [Parameter(Mandatory = $true)]
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
