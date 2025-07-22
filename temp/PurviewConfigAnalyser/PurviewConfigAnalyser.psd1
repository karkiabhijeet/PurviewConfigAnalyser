@{
    # Module manifest for PurviewConfigAnalyser
    RootModule = 'PurviewConfigAnalyser.psm1'
    ModuleVersion = '1.0.4'
    GUID = '7922a05c-1dac-422d-9720-06bf4421e59b'
    Author = 'Abhijeet Karki'
    CompanyName = 'Individual'
    Copyright = '© 2025 Abhijeet Karki. All rights reserved.'
    Description = 'Microsoft Purview Configuration Analyser - Automated compliance assessment for Sensitivity Labels, Auto-labeling, and Data Loss Prevention policies with comprehensive reporting capabilities.'
    
    # Minimum version of PowerShell required
    PowerShellVersion = '5.1'
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{ModuleName = 'ImportExcel'; ModuleVersion = '7.0.0'; }
    )
    
    # Functions to export from this module
    FunctionsToExport = @(
        'Test-PurviewCompliance',
        'Invoke-PurviewConfigAnalyser',
        'Get-PurviewConfig',
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
        'Public\Test-PurviewCompliance.ps1',
        'Public\Invoke-PurviewConfigAnalyser.ps1',
        'Public\Get-PurviewConfig.ps1',
        'Public\New-CustomControlBook.ps1',
        'Private\Test-ControlBook.ps1',
        'Private\DlpAdvancedParser.ps1',
        'Private\GenerateExcelFromJSON.ps1',
        'Collect-PurviewConfiguration.ps1',
        'config\ControlBook_AUGov_Config.csv',
        'config\ControlBook_Property_AUGov_Config.csv',
        'config\MasterControlBooks\ControlBook_Reference.csv',
        'config\MasterControlBooks\ControlBook_Property_Reference.csv'
    )
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            Tags = @('Purview', 'Compliance', 'Microsoft365', 'Security', 'Assessment', 'DLP', 'SensitivityLabels', 'Governance')
            LicenseUri = 'https://github.com/karkiabhijeet/PurviewConfigAnalyser/blob/main/LICENSE'
            ProjectUri = 'https://github.com/karkiabhijeet/PurviewConfigAnalyser'
            IconUri = ''
            ReleaseNotes = @'
# PurviewConfigAnalyser v1.0.4

## New in v1.0.4 - CONFIG FILES INCLUDED
- [CRITICAL FIX] Added missing config files to PowerShell Gallery package
- [INCLUDE] ControlBook_AUGov_Config.csv - Australian Government compliance framework
- [INCLUDE] ControlBook_Property_AUGov_Config.csv - Property mappings for AUGov framework  
- [INCLUDE] Master control books for reference configurations
- [FIX] Resolves "Cannot find path config" error during validation step
- Module now fully functional with all required configuration files

## New in v1.0.3
- [CRITICAL FIX] Fixed all remaining Unicode quote character issues that caused parser errors
- [FIX] Replaced smart quotes (U+2018, U+2019, U+201C, U+201D) with standard ASCII quotes
- [FIX] Fixed "Array index expression missing" errors in string literals
- [FIX] Fixed "Missing argument in parameter list" errors from malformed quotes
- Module now imports successfully without any syntax errors - FINAL FIX

## New in v1.0.2
- [CRITICAL FIX] Fixed syntax errors that prevented module import in v1.0.1
- [FIX] Removed remaining Unicode character causing parser errors
- [FIX] Fixed curly quotes in ReadKey calls that blocked module loading

## New in v1.0.1
- [COMPATIBILITY] Replaced Unicode emoji icons with text equivalents for better PowerShell compatibility
- [FIX] Resolves installation hanging issues in various PowerShell environments
- [IMPROVEMENT] Added comprehensive installation troubleshooting documentation

## Features
- Comprehensive compliance assessment for Microsoft Purview configurations
- Support for Sensitivity Labels, Auto-labeling, and DLP policies
- Advanced parsing for complex nested policy structures
- Excel and CSV reporting capabilities
- 96.3% compliance rate achieved on reference implementation

## Technical Capabilities
- Deep recursive JSON parsing for complex policy conditions
- Case-insensitive property matching
- Support for compound property paths with >> operator
- Enhanced DLP rule parsing for nested SubConditions
- Automated control book validation
- Universal PowerShell environment compatibility

## Requirements
- PowerShell 5.1 or higher
- ImportExcel module 7.0.0 or higher
- Microsoft Purview OptimizedReport JSON export

## Usage
Import-Module PurviewConfigAnalyser
Test-PurviewCompliance -OptimizedReportPath "report.json" -Configuration "AUGov" -OutputPath ".\results"
'@
        }
    }
}
