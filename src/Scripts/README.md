# Utility Scripts

This folder contains standalone utility scripts that are not part of the main PowerShell module but provide additional functionality.

## Scripts

### `DlpPolicyEvaluator.ps1`
**Purpose**: Standalone DLP policy analysis tool  
**Usage**: `.\DlpPolicyEvaluator.ps1 -ReportPath .\output\OptimizedReport_*.json`  
**Description**: Analyzes DLP policies and rules, extracts sensitive types, and flags simulation/test policies.

### `SimplifiedPurviewConfigAnalyserReport.ps1`  
**Purpose**: Alternative simplified reporting interface  
**Usage**: Standalone script for basic compliance assessment  
**Description**: Provides a simplified interface for running compliance assessments with basic reporting.

### `Run-MaturityAssessment.ps1`
**Purpose**: Batch assessment runner  
**Usage**: Automated assessment execution  
**Description**: Wrapper script for running multiple assessments or batch processing.

## Note

These scripts are **not part of the main PowerShell module** and are provided as additional utilities. The main module functionality is accessed through:

```powershell
Import-Module PurviewConfigAnalyser
Test-PurviewCompliance -OptimizedReportPath "report.json" -Configuration "AUGov"
```

For PowerShell Gallery installation, these utility scripts are not included in the published module package.
