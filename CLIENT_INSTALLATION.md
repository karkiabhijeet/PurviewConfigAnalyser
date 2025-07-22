# 📦 Microsoft Purview Configuration Analyser - Client Installation

## 🚀 Simple Installation (Recommended)

### Option 1: PowerShell Gallery (Coming Soon)
```powershell
# Install from PowerShell Gallery (when published)
Install-Module -Name PurviewConfigAnalyser -Force
Import-Module PurviewConfigAnalyser
```

### Option 2: Manual Installation (Current)
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

## 📊 What You Get

### Compliance Assessment For:
- ✅ **Sensitivity Labels** (11 controls)
- ✅ **Sensitivity Auto-labeling** (2 controls) 
- ✅ **Data Loss Prevention** (8 controls)

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
