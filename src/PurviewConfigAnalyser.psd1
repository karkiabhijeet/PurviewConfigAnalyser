@{
    # Module manifest for PurviewConfigAnalyser
    RootModule = 'PurviewConfigAnalyser.psm1'
    ModuleVersion = '1.0.0'
    GUID = '12345678-1234-1234-1234-123456789abc'
    Author = 'Microsoft'
    CompanyName = 'Microsoft'
    Copyright = '© Microsoft Corporation. All rights reserved.'
    Description = 'A PowerShell module for Microsoft Purview configuration analysis and compliance assessment'
    
    # Minimum version of PowerShell required
    PowerShellVersion = '5.1'
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{ModuleName = 'ImportExcel'; ModuleVersion = '7.0.0'; },
        @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.0.0'; }
    )
    
    # Functions to export from this module
    FunctionsToExport = @(
        'Invoke-PurviewConfigAnalyser',
        'Get-PurviewConfig',
        'Test-PurviewCompliance',
        'New-CustomControlBook'
    )
    
    # Cmdlets to export from this module
    CmdletsToExport = @()
    
    # Variables to export from this module
    VariablesToExport = @()
    
    # Aliases to export from this module
    AliasesToExport = @()
    
    # List of all files packaged with this module
    FileList = @(
        'PurviewConfigAnalyser.psd1',
        'PurviewConfigAnalyser.psm1',
        'functions\Private\Convert-ObjectForJson.ps1',
        'functions\Private\EnsureModule.ps1',
        'functions\Private\Write-Log.ps1',
        'functions\Public\Get-PurviewConfig.ps1',
        'functions\Public\Invoke-PurviewConfigAnalyser.ps1',
        'functions\Public\New-CustomControlBook.ps1',
        'functions\Public\Test-PurviewCompliance.ps1'
    )
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            Tags = @('Purview', 'Compliance', 'Microsoft365', 'Security', 'Assessment')
            LicenseUri = ''
            ProjectUri = ''
            IconUri = ''
            ReleaseNotes = 'Initial release of PurviewConfigAnalyser module'
        }
    }
}
