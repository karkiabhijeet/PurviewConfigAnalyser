@{
    # Module manifest for PurviewConfigAnalyser
    RootModule = 'PurviewConfigAnalyser.psm1'
    ModuleVersion = '1.0.8'
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
# PurviewConfigAnalyser v1.0.8

## New in v1.0.8 - VERSION-SPECIFIC OUTPUT DIRECTORIES
- [CRITICAL FIX] Fixed output directory structure for multiple module versions
- [FIX] Output directories are now version-specific within each installed module version
- [FIX] Data collection and file lookups now use consistent ../output paths within each version
- [FIX] Prevents conflicts between multiple installed versions (1.0.0, 1.0.6, 1.0.7, etc.)
- [FIX] Each module version maintains its own output folder structure
- [FIX] Resolves issues where older versions were being referenced due to module loading order
- [IMPROVEMENT] Users no longer need to uninstall previous versions before installing new ones
- Module versions now properly isolated and self-contained

## New in v1.0.7 - OUTPUT PATH CONSISTENCY FIX
- [CRITICAL FIX] Fixed output directory path inconsistency between data collection and file lookup
- [FIX] Data collection creates files at module root level (..\output) but lookups were using version-specific path
- [FIX] Aligned all output path references to use module root level consistently (..\..\output from Public folder)  
- [FIX] Resolves "Cannot find path output" error after successful data collection
- [FIX] Files are now correctly found after being created by data collection process
- Module now works correctly for both data collection and subsequent file lookups

## New in v1.0.6 - USER-FRIENDLY VALIDATION
- [CRITICAL FIX] Fixed validation workflow for installed modules without output directory
- [IMPROVEMENT] Interactive prompt for OptimizedReport JSON file when validation tests are run
- [UX] Clear instructions on how to obtain OptimizedReport JSON files
- [FIX] Graceful handling of missing output directory in Run-MaturityAssessment script
- [FIX] Added automatic output directory creation when needed
- [FIX] Better error handling for file path validation
- Validation tests now work seamlessly for installed modules

## New in v1.0.5 - CRITICAL PATH FIX
- [CRITICAL FIX] Fixed config file path resolution for installed modules
- [FIX] Corrected PSScriptRoot path calculations from "../../config" to "../config"
- [FIX] Fixed config lookup in Invoke-PurviewConfigAnalyser, Test-PurviewCompliance, and all scripts
- [FIX] Module now properly finds config files at correct installed location
- [FIX] Resolves "Cannot find path config" error in PowerShell Gallery installations
- Module paths now work correctly for both development and installed environments

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
