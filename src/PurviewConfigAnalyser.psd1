@{
    # Module manifest for PurviewConfigAnalyser
    RootModule = 'PurviewConfigAnalyser.psm1'
    ModuleVersion = '2.0.1'
    GUID = '7922a05c-1dac-422d-9720-06bf4421e59b'
    Author = 'Abhijeet Karki'
    CompanyName = 'Individual'
    Copyright = '(c) 2025 Abhijeet Karki. All rights reserved.'
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
        'Invoke-PurviewConfigAnalyser'
    )
    
    # Cmdlets to export from this module
    CmdletsToExport = @()
    
    # Variables to export from this module
    VariablesToExport = @()
    
    # Aliases to export from this module
    AliasesToExport = @()
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
    # ModuleList = @()
    
    # List of all files packaged with this module
    # FileList = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @(
                'Microsoft',
                'Purview',
                'Compliance',
                'Security',
                'DLP',
                'SensitivityLabels',
                'AutoLabeling',
                'Assessment',
                'Reporting',
                'Governance',
                'InformationProtection',
                'DataClassification'
            )
            
            # A URL to the license for this module.
            LicenseUri = 'https://github.com/karkiabhijeet/PurviewConfigAnalyser/blob/main/LICENSE'
            
            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/karkiabhijeet/PurviewConfigAnalyser'
            
            # A URL to an icon representing this module.
            # IconUri = ''
            
            # ReleaseNotes of this module
            ReleaseNotes = @'
# PurviewConfigAnalyser v2.0.1 - CRITICAL TAXONOMY FIX

## CRITICAL HOTFIX - Auto-labeling Policy Compatibility 
- [CRITICAL] Fixed SAL_2.4 test failure caused by hardcoded PSPF taxonomy 
- [UNIVERSAL] Get-TaxonomyLabels now dynamically reads SL_1.3 from configuration
- [COMPATIBILITY] Auto-labeling tests now work with any tenant taxonomy
- [INTELLIGENT] Falls back to SL_1.3 defined labels instead of hardcoded values

## What This Fixes
- SAL_2.4 "No taxonomy auto-labeling policies found" error resolved
- Dynamic taxonomy detection from SL_1.3 configuration
- Perfect compatibility with Australian Government PSPF taxonomy
- Support for custom taxonomies defined in control books
- Eliminates hardcoded "UNOFFICIAL, OFFICIAL, OFFICIAL SENSITIVE" assumptions

## Technical Details
The Get-TaxonomyLabels function now:
1. Reads SL_1.3 control from ControlBook_Property_AUGov_Config.csv
2. Extracts the actual taxonomy labels from the "GetLabel > DisplayName" property
3. Uses those labels for sensitivity auto-labeling policy validation
4. Falls back to AUGov taxonomy only if SL_1.3 cannot be found

## Upgrade Immediately
This is a critical fix for auto-labeling policy compliance testing:
```powershell
Install-Module PurviewConfigAnalyser -RequiredVersion 2.0.1 -Force
```

## Previous Stable Release (v2.0.0)
v2.0.0 was the production-ready release with all path issues resolved.
v2.0.1 adds critical auto-labeling taxonomy compatibility.
'@
        }
    }
}
