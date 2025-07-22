# PurviewConfigAnalyser PowerShell Module
# Main module file that imports all functions and sets up the module environment

# Get the module root path
$ModuleRoot = $PSScriptRoot

# Import all private functions first
$PrivateFunctions = Get-ChildItem -Path "$ModuleRoot\Private\*.ps1" -ErrorAction SilentlyContinue
foreach ($Function in $PrivateFunctions) {
    . $Function.FullName
}

# Import all public functions
$PublicFunctions = Get-ChildItem -Path "$ModuleRoot\Public\*.ps1" -ErrorAction SilentlyContinue
foreach ($Function in $PublicFunctions) {
    . $Function.FullName
}

# Auto-install required dependencies if not present
function Initialize-Dependencies {
    $RequiredModules = @(
        @{Name = 'ImportExcel'; MinVersion = '7.0.0'},
        @{Name = 'ExchangeOnlineManagement'; MinVersion = '3.0.0'}
    )
    
    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module.Name)) {
            Write-Host "Installing required module: $($Module.Name)..." -ForegroundColor Yellow
            try {
                Install-Module -Name $Module.Name -MinimumVersion $Module.MinVersion -Force -Scope CurrentUser -ErrorAction Stop
                Write-Host "✅ Successfully installed $($Module.Name)" -ForegroundColor Green
            } catch {
                Write-Warning "Failed to install $($Module.Name): $($_.Exception.Message)"
            }
        }
    }
}

# Initialize dependencies on module import
Initialize-Dependencies

# Export module members
Export-ModuleMember -Function @(
    'Invoke-PurviewConfigAnalyser',
    'Get-PurviewConfig', 
    'Test-PurviewCompliance',
    'New-CustomControlBook'
)

# Module cleanup
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    # Cleanup code here if needed
}
