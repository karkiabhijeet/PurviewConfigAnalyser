function Connect-ToComplianceCenter {
    <#
    .SYNOPSIS
        Establishes connection to Microsoft Compliance Center.
    
    .DESCRIPTION
        Prompts for user credentials and establishes a connection to Microsoft Compliance Center
        using the ExchangeOnlineManagement module.
    #>
    
    Write-Host "Connecting to Microsoft Compliance Center..." -ForegroundColor Yellow
    try {
        # Prompt for credentials
        $userName = Read-Host -Prompt 'Enter your User Principal Name (UPN)'
        # Connect to the Compliance Center using UserPrincipalName
        Connect-IPPSSession -UserPrincipalName $userName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

        $InfoMessage = "[SUCCESS] Connection established successfully!"
        Write-Host $InfoMessage -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Connection failed: $_" -ForegroundColor Red
        throw
    }
}

function EnsureComplianceCenterConnection {
    <#
    .SYNOPSIS
        Ensures Microsoft Compliance Center connection is active.
    
    .DESCRIPTION
        Checks if the compliance center session is still active and reconnects if necessary.
    #>
    
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
        throw
    }
}
