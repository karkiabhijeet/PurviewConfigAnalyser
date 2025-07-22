# 📦 Microsoft Purview Configuration Analyser - Client Installation

## 🚀 Simple Installation (Recommended)

### Option 1: PowerShell Gallery (AVAILABLE NOW!)
```powershell
# Install from PowerShell Gallery
Install-Module -Name PurviewConfigAnalyser -Force
Import-Module PurviewConfigAnalyser

# Start using immediately
Invoke-PurviewConfigAnalyser
```

> **Installation Note**: The module depends on ImportExcel (~1MB). The installation may take 1-2 minutes to download dependencies, especially on slower connections. If it appears to "hang", it's likely downloading the ImportExcel module. For troubleshooting, see [INSTALLATION_TROUBLESHOOTING.md](./INSTALLATION_TROUBLESHOOTING.md).

**Quick Fix for Slow Installations:**
```powershell
# If installation seems stuck, try with verbose output to see progress
Install-Module -Name PurviewConfigAnalyser -Force -Verbose

# Or install dependencies first
Install-Module -Name ImportExcel -Force
Install-Module -Name PurviewConfigAnalyser -Force
```

### Option 2: Manual Installation (Alternative)
```powershell
# Download and extract the module
# Place in your PowerShell modules directory, then:
Import-Module .\PurviewConfigAnalyser\src\PurviewConfigAnalyser.psd1 -Force
```

## 📋 Prerequisites
- **PowerShell 5.1** or higher (check with `$PSVersionTable.PSVersion`)
- **ImportExcel module** (automatically installed with the module)

## 🎯 Quick Start - 3 Simple Steps

### Step 1: Get Your Purview Report
Export your Microsoft Purview configuration as an OptimizedReport JSON file from the Purview portal.

### Step 2: Run Assessment
```powershell
Test-PurviewCompliance -OptimizedReportPath ".\OptimizedReport_*.json" -Configuration "AUGov" -OutputPath ".\results"
```

### Step 3: Review Results
- **Excel Report**: `.\results\results_AUGov.xlsx` (detailed with multiple tabs)
- **CSV Report**: `.\results\results_AUGov.csv` (simple data format)

## Compliance Assessment For:
- [YES] **Sensitivity Labels** (11 controls)
- [YES] **Sensitivity Auto-labeling** (2 controls) 
- [YES] **Data Loss Prevention** (8 controls)

### Reports Include:
- **Overall Compliance Rate** (currently achieving 96.3% on reference data)
- **Control-by-Control Analysis** with pass/fail status
- **Detailed Comments** explaining why controls pass or fail
- **Maturity Level Summary** showing progression across capability areas

### Sample Output:
```
=== AUGov Summary ===
Total Controls Evaluated: 27
Controls Passing: 26
Controls Failing: 1
Compliance Rate: 96.3%
```

## 🛠️ Troubleshooting

### If ImportExcel module is missing:
```powershell
Install-Module -Name ImportExcel -Force
```

### If you get execution policy errors:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### For module path issues:
```powershell
# Check your module paths
$env:PSModulePath -split ';'

# Install in user scope if needed
Install-Module -Name PurviewConfigAnalyser -Scope CurrentUser
```

## 🔗 Support & Documentation
- **GitHub Repository**: https://github.com/karkiabhijeet/PurviewConfigAnalyser
- **Detailed Documentation**: See README.md and PROGRESS_UPDATE.md
- **Issues & Questions**: Create an issue on GitHub

## 📈 Advanced Usage

### Custom Configuration:
```powershell
# Use different control frameworks
Test-PurviewCompliance -Configuration "Custom" -ControlBookPath ".\custom-controls.csv"

# Multiple report formats
Test-PurviewCompliance -ExportExcel -ExportCSV -OutputPath ".\detailed-results"
```

---
**Ready to assess your Microsoft Purview compliance in minutes!** 🚀
