@{
    ModuleVersion = '2.0.0'
    GUID = 'd1e7c3b5-8b5f-4c8b-9b5e-1f5c5e7c3b5e'
    Author = 'Your Name'
    CompanyName = 'Your Company'
    Copyright = '2024 Your Company'
    Description = 'A module for evaluating maturity models including Sensitivity Labels and DLP policies based on JSON configurations.'
    FunctionsToExport = @(
        'Get-MaturityModelReport',
        'Get-DLPMaturityReport',
        'Test-SensitivityLabels',
        'Test-DLPPolicies',
        'Export-SensitivityLabelReportToHtml',
        'Export-DLPReportToHtml'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    NestedModules = @()
    RequiredModules = @('ExchangeOnlineManagement')
    RequiredAssemblies = @()
    FileList = @(
        'MaturityModelReport.psm1',
        'functions\Test-SensitivityLabels.ps1',
        'functions\Test-DLPPolicies.ps1',
        'functions\Export-SensitivityLabelReportToHtml.ps1',
        'functions\Export-DLPReportToHtml.ps1',
        'types\MaturityModel.types.ps1xml'
    )
    PrivateData = @{
        PSData = @{
            LicenseUri = 'https://opensource.org/licenses/MIT'
            ProjectUri = 'https://github.com/YourUsername/MaturityModelReportModule'
            BugTrackerUri = 'https://github.com/YourUsername/MaturityModelReportModule/issues'
            ReleaseNotes = 'Version 2.0.0: Added DLP policy evaluation functionality with comprehensive maturity model assessment for both Sensitivity Labels and Data Loss Prevention policies.'
        }
    }
}