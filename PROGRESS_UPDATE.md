# AI Development Context - Microsoft Purview Configuration Analyser

**Last Updated:** July 9, 2025  
**Current Version:** 1.0 (Pre-Release)  
**Status:** Ready for initial GitHub publication  
**Session Context:** Preserved for AI agent continuation

---

## 🤖 AI Agent Instructions

**PURPOSE**: This document provides technical implementation context for AI agents working on the Microsoft Purview Configuration Analyser project. Contains code references, architectural decisions, and implementation details.

**TARGET AUDIENCE**: AI development agents, not end users (see README.md for user documentation)

---

## 🎯 Technical Project Overview

**Core Architecture**: PowerShell module with CLI primary interface + GUI for custom configuration creation only  
**Framework**: Windows Forms for GUI, PowerShell 5.1+ compatibility  
**Data Flow**: Microsoft Graph API → JSON collection → CSV control books → Assessment engine → Excel/CSV reports  
**Key Design Decision**: CLI-first approach with GUI only for Option 4 (Create Custom Configuration)

---

## 🔧 Implementation Status & Code References

### Core Module Files (COMPLETED)
```
src/PurviewConfigAnalyser.psm1     # Main module - imports all functions
src/PurviewConfigAnalyser.psd1     # Module manifest with metadata
src/functions/Public/              # Exported functions
src/functions/Private/             # Internal helper functions
```

### Key Function Implementations
- **`Invoke-PurviewConfigAnalyser.ps1`** - Main entry point with 4-option CLI menu
  - Location: `src/functions/Public/Invoke-PurviewConfigAnalyser.ps1`
  - Status: ✅ IMPLEMENTED - Interactive menu with validation
  - Integration Point: Option 4 calls GUI function (TO BE IMPLEMENTED)

- **`Collect-PurviewConfiguration.ps1`** - Data collection from Microsoft Graph
  - Location: `src/Collect-PurviewConfiguration.ps1`
  - Status: ✅ IMPLEMENTED - Full Graph API integration
  - Output: JSON files in `output/` directory

- **`Run-MaturityAssessment.ps1`** - Assessment engine
  - Location: `src/Run-MaturityAssessment.ps1`
  - Status: ✅ IMPLEMENTED - CSV and Excel report generation
  - Dependencies: ImportExcel module, control books

- **`Show-PurviewConfigAnalyserGUI.ps1`** - GUI for custom configuration
  - Location: `src/Show-PurviewConfigAnalyserGUI.ps1`
  - Status: ❌ NOT IMPLEMENTED - Priority 1 for next session
  - Requirements: Windows Forms, TreeView controls, property editors

### Configuration System Architecture
```
config/
├── ControlBook_PSPF_Config.csv              # PSPF framework controls
├── ControlBook_Property_PSPF_Config.csv     # PSPF validation criteria
└── MasterControlBooks/                      # Master templates
    ├── ControlBook_Reference.csv            # All available controls
    └── ControlBook_Property_Reference.csv   # All validation properties
```

**Control Book Logic**:
1. Master reference books define ALL available controls and properties
2. Framework-specific books (e.g., PSPF) select subsets and customize criteria
3. Custom configurations created via GUI use master books as templates
4. Assessment engine reads control books to determine what to test and validation criteria

### Data Flow Implementation
```
Microsoft Graph API → Collect-PurviewConfiguration.ps1 → OptimizedReport_[GUID].json
↓
Control Books (CSV) → Run-MaturityAssessment.ps1 → TestResults_[Framework].csv
↓
ImportExcel → MaturityAssessment_[Framework].xlsx
```

## � Critical Implementation Details for AI Agents

### User Requirements (FROM PREVIOUS SESSION)
- **CLI Interface**: MUST remain primary interaction method
- **GUI Integration**: ONLY for Option 4 (Create Custom Configuration)
- **No GUI for Options 1-3**: User explicitly rejected full GUI implementation
- **Contact Email**: karkiabhijeet@gmail.com (updated in README.md)

### Code Integration Points
1. **CLI Menu Option 4**: Currently shows placeholder - needs GUI function call
   ```powershell
   # Location: src/functions/Public/Invoke-PurviewConfigAnalyser.ps1
   # Around line 150-200 (estimate)
   4 { 
       # TODO: Implement GUI call
       # Should call: Show-PurviewConfigAnalyserGUI
   }
   ```

2. **GUI Function Signature** (TO BE IMPLEMENTED):
   ```powershell
   function Show-PurviewConfigAnalyserGUI {
       param(
           [string]$MasterControlPath = "$PSScriptRoot\..\..\config\MasterControlBooks\ControlBook_Reference.csv",
           [string]$MasterPropertyPath = "$PSScriptRoot\..\..\config\MasterControlBooks\ControlBook_Property_Reference.csv",
           [string]$OutputPath = "$PSScriptRoot\..\..\config\"
       )
       # Windows Forms implementation needed
   }
   ```

### Control Book CSV Structure (FOR GUI IMPLEMENTATION)
**ControlBook_Reference.csv columns**:
- `Capability` - Grouping (e.g., "Sensitivity Labels", "DLP")
- `ControlID` - Unique identifier (e.g., "SL-001", "DLP-003")
- `Control` - Description of what is being tested
- `IsActive` - Boolean for availability

**ControlBook_Property_Reference.csv columns**:
- `ControlID` - Links to control
- `Properties` - Property being validated
- `DefaultValue` - Default validation criteria
- `MustConfigure` - Boolean for required customization

### Assessment Engine Integration
**File**: `src/Run-MaturityAssessment.ps1`
**Function**: `Test-ControlBook`
**Parameters**: 
- `-ControlConfigPath` - Path to control book CSV
- `-PropertyConfigPath` - Path to property book CSV
- `-OptimizedReportPath` - Path to collected JSON data
- `-OutputPath` - Where to save results

### Missing Dependencies for GUI
- Windows Forms assembly loading
- TreeView control implementation
- Property grid for configuration editing
- CSV export functionality for custom control books

### Error Handling Patterns (EXISTING)
```powershell
try {
    # Implementation
    Write-Host "✅ Success message" -ForegroundColor Green
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
```

## � Next Session Implementation Tasks

### PRIORITY 1: GUI Implementation for Custom Configuration Creation
**File to create**: `src/Show-PurviewConfigAnalyserGUI.ps1`

**Required Windows Forms Components**:
```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Custom Configuration Creator"
$form.Size = New-Object System.Drawing.Size(800, 600)

# TreeView for control selection (by Capability)
$treeView = New-Object System.Windows.Forms.TreeView
# Property Panel for configuration
$propertyPanel = New-Object System.Windows.Forms.Panel
# Save/Export buttons
$saveButton = New-Object System.Windows.Forms.Button
```

**GUI Workflow Logic**:
1. Load master reference books from CSV
2. Parse controls by Capability (create tree nodes)
3. Display property configuration panel when control selected
4. Validate required properties are configured
5. Export to custom control book CSV files

**Integration with CLI**:
- Called from Option 4 in `Invoke-PurviewConfigAnalyser.ps1`
- Should return to CLI menu after completion
- Must handle cancellation gracefully

### PRIORITY 2: Enhanced Test Coverage Implementation
**Files to examine**:
- `config/MasterControlBooks/ControlBook_Reference.csv` - Identify missing test implementations
- `src/functions/Private/` - Look for existing test functions

**Test Implementation Pattern**:
```powershell
function Test-ControlID {
    param(
        [Parameter(Mandatory = $true)]$CollectedData,
        [Parameter(Mandatory = $true)]$ExpectedValue,
        [Parameter(Mandatory = $true)]$ControlID
    )
    
    # Implementation logic
    # Return: [PSCustomObject]@{
    #     ControlID = $ControlID
    #     Pass = $true/$false
    #     CurrentValue = $actual
    #     ExpectedValue = $expected
    #     Recommendation = $message
    # }
}
```

### PRIORITY 3: Code Quality and Performance
**Areas to examine**:
- Memory usage during large JSON processing
- Error handling in data collection
- PowerShell compatibility across versions
- Excel generation performance optimization

---

## 📊 Technical State Assessment

### Working Components (TESTED)
- ✅ CLI menu system with input validation
- ✅ Microsoft Graph API data collection
- ✅ PSPF framework control book processing
- ✅ CSV and Excel report generation
- ✅ Module import/export functionality

### Missing Components (IMPLEMENTATION NEEDED)
- ❌ Windows Forms GUI for custom configuration
- ❌ TreeView control binding to master reference books
- ❌ Property configuration panels
- ❌ Custom control book CSV export
- ❌ Additional test implementations for all master controls

### Integration Points (REQUIRES ATTENTION)
- 🔄 CLI Option 4 → GUI function call
- 🔄 GUI → Master reference book loading
- 🔄 GUI → Custom control book generation
- 🔄 New tests → Assessment engine integration

---

## � Development Environment Context

### Required Tools
- **PowerShell ISE** or **VS Code** with PowerShell extension
- **Windows Forms Designer** (optional, code-first approach preferred)
- **Microsoft Graph PowerShell SDK** (auto-installed)
- **ImportExcel module** (auto-installed)

### Testing Requirements
- **Active Microsoft 365 tenant** for data collection testing
- **Global Admin or Compliance Admin** permissions
- **Test control books** for validation

### Code Standards
- **Error handling**: Try-catch with colored output
- **Logging**: Use existing Write-Log function pattern
- **Parameters**: CmdletBinding with proper validation
- **Output**: Consistent formatting with existing functions

---

## � Debugging and Troubleshooting Context

### Common Issues Encountered
1. **Module Import**: Functions not recognized - check module manifest
2. **CSV Processing**: Encoding issues - use UTF-8 explicitly
3. **GUI Threading**: Form not responsive - use proper event handling
4. **Memory**: Large JSON processing - implement streaming if needed

### Validation Commands
```powershell
# Test module import
Import-Module .\src\PurviewConfigAnalyser.psm1 -Force
Get-Command -Module PurviewConfigAnalyser

# Test CLI menu
Invoke-PurviewConfigAnalyser

# Test control book loading
Import-Csv ".\config\MasterControlBooks\ControlBook_Reference.csv"
```

### Log File Locations
- **Execution logs**: `output/file_runlog.txt`
- **Collection data**: `output/OptimizedReport_[GUID]_[Timestamp].json`
- **Test results**: `output/TestResults_[Framework]_[GUID]_[Timestamp].csv`

---

## � Session Continuation Instructions for AI Agents

### IMMEDIATE ACTIONS FOR NEXT SESSION:
1. **Read this entire document** - Contains all technical context
2. **Examine master reference books** - `config/MasterControlBooks/` to understand control structure
3. **Review CLI integration point** - `src/functions/Public/Invoke-PurviewConfigAnalyser.ps1` Option 4
4. **Begin GUI implementation** - Create `src/Show-PurviewConfigAnalyserGUI.ps1`

### DECISION CONTEXT:
- **User explicitly wants CLI-first approach**
- **GUI ONLY for Option 4 (Create Custom Configuration)**
- **No changes to Options 1-3 workflows**
- **Must integrate seamlessly with existing CLI menu**

### SUCCESS CRITERIA:
- [ ] GUI launches from CLI Option 4
- [ ] GUI loads master reference books
- [ ] GUI allows control selection by capability
- [ ] GUI generates custom control book CSV files
- [ ] GUI returns to CLI menu after completion
- [ ] Enhanced test coverage for all master controls

---

**AI Agent Context Preservation**: This document contains all technical implementation details needed to continue development without losing context. The user's preferences and architectural decisions are preserved above.

**Contact**: karkiabhijeet@gmail.com  
**Last Session**: July 9, 2025  
**Next Focus**: GUI implementation for custom configuration creation

---

*AI Development Context Document - Technical Implementation Guide*
