# 🔧 Installation Troubleshooting Guide

## Common Installation Issues & Solutions

### Issue 1: Installation Appears to "Hang" or Get Stuck

**Symptoms:**
- `Install-Module -Name PurviewConfigAnalyser` seems to freeze
- Progress bar stops or appears stuck on "Installing dependent package 'ImportExcel'"
- No error messages, but installation doesn't complete

**Why This Happens:**
The PurviewConfigAnalyser module depends on ImportExcel (~1MB), which can take time to download on slower connections.

**Solutions:**

#### Quick Fix - Install with Progress Indication:
```powershell
# Install with detailed progress information
Install-Module -Name PurviewConfigAnalyser -Force -Verbose
```

#### Alternative - Install Dependencies First:
```powershell
# Pre-install the large dependency
Install-Module -Name ImportExcel -Force -Verbose
# Then install our module
Install-Module -Name PurviewConfigAnalyser -Force
```

#### Network Timeout Fix:
```powershell
# Increase timeout for slow connections
$PSDefaultParameterValues['*:TimeoutSec'] = 300
Install-Module -Name PurviewConfigAnalyser -Force
```

### Issue 2: Execution Policy Restrictions

**Symptoms:**
- Error: "Execution of scripts is disabled on this system"
- Module installs but won't import

**Solution:**
```powershell
# Allow scripts for current user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Issue 3: Module Path Issues

**Symptoms:**
- Module installs but can't be found
- `Import-Module PurviewConfigAnalyser` fails

**Solution:**
```powershell
# Check your module paths
$env:PSModulePath -split ';'

# Install specifically for current user
Install-Module -Name PurviewConfigAnalyser -Scope CurrentUser -Force
```

### Issue 4: Multiple PowerShell Versions

**Symptoms:**
- Module works in one PowerShell but not another
- Commands not recognized after installation

**Solution:**
```powershell
# Check your PowerShell version
$PSVersionTable.PSVersion

# Install for all users if you use multiple PS versions
Install-Module -Name PurviewConfigAnalyser -Scope AllUsers -Force
```

## Installation Progress Expectations

### Normal Installation Timeline:
1. **0-10 seconds**: Finding module in PowerShell Gallery
2. **10-60 seconds**: Downloading ImportExcel dependency (~1MB)
3. **60-70 seconds**: Downloading PurviewConfigAnalyser (~60KB)
4. **70-80 seconds**: Installing and validating modules
5. **Complete**: Ready to use!

### What You Should See:
```
Installing package 'PurviewConfigAnalyser' [Installing dependent package 'ImportExcel']
  Installing package 'ImportExcel' [Downloaded 0.67 MB out of 1.09 MB.]
  Installing package 'ImportExcel' [Downloaded 1.01 MB out of 1.09 MB.]
Installing package 'PurviewConfigAnalyser' [Downloaded 0.00 MB out of 0.06 MB.]
```

## Quick Verification

After installation, verify everything works:

```powershell
# Import the module
Import-Module PurviewConfigAnalyser

# Check available commands
Get-Command -Module PurviewConfigAnalyser

# Test the main function
Invoke-PurviewConfigAnalyser
```

Expected output:
```
CommandType     Name                                Version    Source
-----------     ----                                -------    ------
Function        Get-PurviewConfig                   1.0.0      PurviewConfigAnalyser
Function        Invoke-PurviewConfigAnalyser        1.0.0      PurviewConfigAnalyser
Function        New-CustomControlBook               1.0.0      PurviewConfigAnalyser
Function        Test-PurviewCompliance              1.0.0      PurviewConfigAnalyser
```

## Still Having Issues?

### Manual Installation (Fallback):
1. Download the module from [GitHub Releases](https://github.com/karkiabhijeet/PurviewConfigAnalyser/releases)
2. Extract to your PowerShell modules folder
3. Install ImportExcel separately: `Install-Module ImportExcel -Force`

### Get Help:
- **GitHub Issues**: [Report problems here](https://github.com/karkiabhijeet/PurviewConfigAnalyser/issues)
- **Check Module Status**: `Get-Module PurviewConfigAnalyser -ListAvailable`
- **Verbose Logs**: Add `-Verbose` to any command for detailed information

## Pro Tips for Network Issues

```powershell
# For corporate networks with proxies
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials

# For slow connections - be patient!
# The ImportExcel dependency is 1MB+ and may take time
```

---
**Remember**: Most "hanging" installations are just slow downloads. Be patient, especially on the ImportExcel dependency! ⏳
