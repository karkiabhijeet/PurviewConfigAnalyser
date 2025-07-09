# Microsoft Purview Configuration Analyser

A comprehensive PowerShell module for assessing Microsoft Purview compliance posture against industry standards and custom organizational requirements.

## 🎯 Overview

The Microsoft Purview Configuration Analyser helps organizations evaluate their Microsoft Purview implementation against established compliance frameworks such as:

- **PSPF (Protective Security Policy Framework)** - Australian Government security standards
- **Custom Organizational Frameworks** - Tailored to your specific requirements
- **Industry Standards** - Extensible to support various compliance frameworks

## 📋 Prerequisites

### Software Requirements
- **PowerShell 5.1** or higher
- **Microsoft 365 Admin Access** - Required for data collection
- **Microsoft Graph PowerShell SDK** - Will be installed automatically if missing
- **Windows PowerShell ISE** or **VS Code** (recommended for development)

### Permissions Required
- **Global Administrator** or **Compliance Administrator** role in Microsoft 365
- **Microsoft Graph API Permissions**:
  - `InformationProtectionPolicy.Read.All`
  - `Policy.Read.All`
  - `Directory.Read.All`

## 🚀 Installation


### Option 1: Direct Download from GitHub
```powershell
# Download the repository
git clone https://github.com/your-org/PurviewConfigAnalyser.git
cd PurviewConfigAnalyser

# Import the module (seamless experience)
Import-Module .\src\PurviewConfigAnalyser.psm1 -Force
```

### Option 2: PowerShell Gallery (Coming Soon)
```powershell
# Install from PowerShell Gallery
Install-Module -Name PurviewConfigAnalyser

# Import the module
Import-Module PurviewConfigAnalyser
```

---

## 🚀 Seamless Module-Based Usage for Third Parties

The module is designed for a frictionless experience:

1. **Install or import the module** (see above)
2. **Run a single command:**
  ```powershell
  Invoke-PurviewConfigAnalyser
  ```
  - This launches the interactive menu for all assessment, reporting, and custom configuration tasks.
  - No manual file editing or script modification required.
3. **All outputs** (CSV, Excel, logs) are generated in the `output/` folder.

**Advanced:** You can also use `Invoke-PurviewConfigAnalyser -Mode ...` for automation or CI/CD.

---

## 🎮 Getting Started

### Quick Start - Interactive Menu
The easiest way to get started is with the interactive menu:

```powershell
# Launch the interactive menu
Invoke-PurviewConfigAnalyser
```

This will present you with a user-friendly menu system with 4 main options:

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           MAIN MENU - CHOOSE YOUR ACTION                           │
├─────────────────────────────────────────────────────────────────────────────────┤
│  1. 🔄 Extract Configuration & Run Tests                                          │
│     → Connect to your tenant, collect data, then run compliance tests             │
│     → Best for: Complete assessment from start to finish                          │
│                                                                                     │
│  2. 📊 Extract Configuration Only                                                 │
│     → Connect to your tenant and collect configuration data                        │
│     → Best for: Data collection without immediate testing                          │
│                                                                                     │
│  3. ✅ Run Validation Tests Only                                                  │
│     → Use existing data to run compliance tests                                    │
│     → Best for: Testing against previously collected data                          │
│                                                                                     │
│  4. ⚙️ Create Custom Configuration                                                 │
│     → Build your own control book for organization-specific requirements          │
│     → Best for: Custom compliance frameworks                                       │
│                                                                                     │
│  5. 🚪 Exit                                                                        │
│     → Close the application                                                         │
└─────────────────────────────────────────────────────────────────────────────────┘
```

## 📖 Detailed Usage Guide

### Option 1: Extract Configuration & Run Tests
**Perfect for first-time users or complete assessments**

This option performs a full end-to-end assessment:

1. **Authentication**: Connects to your Microsoft 365 tenant
2. **Data Collection**: Extracts Purview configuration data
3. **Framework Selection**: Choose from available compliance frameworks
4. **Testing**: Runs validation tests against your configuration
5. **Reporting**: Generates comprehensive reports (CSV and Excel)

**Example workflow:**
```powershell
Invoke-PurviewConfigAnalyser
# Select option 1
# Follow prompts for authentication
# Select compliance framework (e.g., PSPF)
# Review generated reports in output folder
```

### Option 2: Extract Configuration Only
**Ideal for data collection without immediate testing**

Use this when you want to:
- Collect configuration data for later analysis
- Set up a baseline for regular monitoring
- Separate data collection from testing phases

**Generated files:**
- `OptimizedReport_[GUID]_[Timestamp].json` - Complete configuration data
- `CollectionRaw_AfterPreProcessing.xml` - Raw collection data
- `file_runlog.txt` - Detailed collection log

### Option 3: Run Validation Tests Only
**Best for testing against previously collected data**

This option assumes you already have configuration data and want to:
- Test against a different compliance framework
- Re-run tests after configuration changes
- Generate reports in different formats

**Available frameworks:**
- **PSPF**: Australian Government Protective Security Policy Framework
- **Custom**: Any custom configurations you've created

### Option 4: Create Custom Configuration
**Advanced users: Build organization-specific compliance frameworks**

This interactive wizard helps you create custom control books tailored to your organization's requirements.

#### Understanding Control Books

A **Control Book** consists of two main components:

1. **ControlBook_[Name]_Config.csv** - Defines which controls are active
2. **ControlBook_Property_[Name]_Config.csv** - Defines the validation criteria

#### Master Reference Books

The module uses master reference books as templates:

- **`config/MasterControlBooks/ControlBook_Reference.csv`** - Master list of all available controls
- **`config/MasterControlBooks/ControlBook_Property_Reference.csv`** - Master list of all properties and criteria

#### Creating Custom Configurations

When you select Option 4, the wizard will:

1. **Prompt for Configuration Name**: Enter a unique name (e.g., "ACME_Corp", "Healthcare_Custom")
2. **Control Selection**: Choose which controls to include from the master reference
3. **Property Configuration**: Set specific criteria for each selected control
4. **Validation**: Configure required vs. optional properties
5. **Generation**: Create your custom control book files

**Example custom configuration process:**
```powershell
Invoke-PurviewConfigAnalyser
# Select option 4
# Enter configuration name: "ACME_Corp"
# Select controls by capability:
#   - Sensitivity Labels: Enable
#   - DLP Policies: Enable  
#   - Retention: Customize
# Configure properties for each control
# Save configuration
```

**Generated files:**
- `config/ControlBook_ACME_Corp_Config.csv`
- `config/ControlBook_Property_ACME_Corp_Config.csv`

## 📁 Project Structure

```
PurviewConfigAnalyser/
├── README.md                          # This file
├── src/
│   ├── PurviewConfigAnalyser.psm1    # Main module file
│   ├── PurviewConfigAnalyser.psd1    # Module manifest
│   ├── Public/                       # Public functions
│   │   ├── Invoke-PurviewConfigAnalyser.ps1
│   │   ├── Get-PurviewConfig.ps1
│   │   ├── Test-PurviewCompliance.ps1
│   │   └── New-CustomControlBook.ps1
│   ├── Private/                      # Private helper functions
│   └── functions/                    # Additional supporting functions
├── config/                           # Configuration files
│   ├── ControlBook_PSPF_Config.csv          # PSPF control definitions
│   ├── ControlBook_Property_PSPF_Config.csv # PSPF property criteria
│   └── MasterControlBooks/                  # Master reference files
│       ├── ControlBook_Reference.csv        # Master control list
│       └── ControlBook_Property_Reference.csv # Master property list
├── output/                           # Generated reports and logs
└── archive/                          # Archived components
```

## 🔧 Configuration Framework Deep Dive

### How Control Books Work

#### 1. Master Reference Books
Located in `config/MasterControlBooks/`, these are the source of truth:

**ControlBook_Reference.csv** contains:
- `Capability`: Grouping (e.g., "Sensitivity Labels", "DLP")
- `ControlID`: Unique identifier (e.g., "SL-001", "DLP-003")
- `Control`: Description of what is being tested
- `IsActive`: Whether the control is available for use

**ControlBook_Property_Reference.csv** contains:
- `ControlID`: Links to the control
- `Properties`: What property is being validated
- `DefaultValue`: Default validation criteria
- `MustConfigure`: Whether this property requires customization

#### 2. Framework-Specific Control Books
Each compliance framework has its own pair of files:

**ControlBook_[Framework]_Config.csv**:
- Defines which controls from the master list are active for this framework
- Can modify control descriptions for framework-specific language
- Controls the scope of testing

**ControlBook_Property_[Framework]_Config.csv**:
- Defines the specific validation criteria for each control
- Customizes thresholds, requirements, and expected values
- Supports organization-specific requirements

#### 3. Custom Configuration Creation Process

When creating a custom configuration:

1. **Template Loading**: Master reference books are loaded as templates
2. **Control Selection**: Choose which controls to include (by capability)
3. **Property Customization**: Set specific criteria for each control
4. **Validation**: Ensure required properties are configured
5. **Export**: Generate the two CSV files for your custom framework

### Example: PSPF Framework

The included PSPF (Protective Security Policy Framework) configuration demonstrates:

**Controls included:**
- Sensitivity Labels configuration and usage
- Data Loss Prevention policies
- Retention policies and settings
- Information barriers
- Trainable classifiers

**Validation criteria:**
- Minimum number of sensitivity labels required
- DLP policy coverage requirements
- Retention period specifications
- Compliance with Australian Government requirements


## 📊 Report Generation

### Excel Output Troubleshooting

If you encounter errors saving Excel files (e.g., "Error saving file ..."), ensure:
- The `output/` directory exists and is writable
- The output path is a file, not a directory
- The `ImportExcel` module is installed (the module will auto-install if missing)

If issues persist, run PowerShell as Administrator or check OneDrive sync status.

### Output Files

All reports are generated in the `output/` directory:

#### Data Collection Files
- **OptimizedReport_[GUID]_[Timestamp].json** - Complete configuration data
- **CollectionRaw_AfterPreProcessing.xml** - Raw XML data from Microsoft Graph
- **file_runlog.txt** - Detailed execution log

#### Assessment Reports
- **TestResults_[Framework]_[GUID]_[Timestamp].csv** - Detailed test results
- **MaturityAssessment_[Framework]_[GUID]_[Timestamp].xlsx** - Executive summary report
- **results_[Framework].csv** - Summary results file

#### Report Contents

**CSV Reports** contain:
- Control ID and description
- Test result (Pass/Fail/Warning)
- Current configuration value
- Expected value
- Recommendation for remediation

**Excel Reports** include:
- Executive summary dashboard
- Detailed findings by capability
- Remediation recommendations
- Compliance scoring

## 🔍 Troubleshooting

### Common Issues

#### Authentication Problems
```powershell
# Clear cached credentials
Disconnect-MgGraph
Connect-MgGraph -Scopes "InformationProtectionPolicy.Read.All", "Policy.Read.All"
```

#### Module Import Issues
```powershell
# Force reimport
Remove-Module PurviewConfigAnalyser -Force -ErrorAction SilentlyContinue
Import-Module .\src\PurviewConfigAnalyser.psm1 -Force
```

#### Missing Dependencies
```powershell
# Install required modules
Install-Module Microsoft.Graph -Force
Install-Module ImportExcel -Force
```

### Debug Mode
Enable verbose logging for troubleshooting:

```powershell
# Enable verbose output
$VerbosePreference = "Continue"
Invoke-PurviewConfigAnalyser -Verbose
```

## 🔄 Advanced Usage

### Direct Function Calls
For automation scenarios, you can call functions directly:

```powershell
# Data collection only
$configData = Get-PurviewConfig -OutputPath "C:\Reports"

# Test specific framework
$results = Test-PurviewCompliance -ConfigurationName "PSPF" -SkipDataCollection

# Create custom configuration programmatically
New-CustomControlBook -ConfigurationName "MyOrg" -Controls $selectedControls
```

### Scheduled Assessments
Create scheduled tasks for regular compliance monitoring:

```powershell
# Create a scheduled task
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-Command `"Import-Module PurviewConfigAnalyser; Invoke-PurviewConfigAnalyser -Mode CollectAndTest -ConfigurationName PSPF`""
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6:00AM
Register-ScheduledTask -TaskName "PurviewCompliance" -Action $Action -Trigger $Trigger
```

## 🤝 Contributing

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

### Adding New Frameworks
To add a new compliance framework:

1. **Create control definitions** in `config/ControlBook_[Framework]_Config.csv`
2. **Define validation criteria** in `config/ControlBook_Property_[Framework]_Config.csv`
3. **Test the configuration** using the module
4. **Document the framework** in this README

### Extending Master Reference Books
To add new controls:

1. **Add to master reference** in `config/MasterControlBooks/ControlBook_Reference.csv`
2. **Define properties** in `config/MasterControlBooks/ControlBook_Property_Reference.csv`
3. **Update existing frameworks** if relevant
4. **Test thoroughly** with multiple configurations

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

### Documentation
- **README.md** - This comprehensive guide
- **Function Help** - Use `Get-Help Invoke-PurviewConfigAnalyser -Detailed`
- **Example Scripts** - Located in `/examples` folder

### Community Support
- **GitHub Issues** - Report bugs and request features
- **Discussions** - Ask questions and share experiences
- **Wiki** - Additional documentation and examples

### Enterprise Support
For enterprise deployments and custom development:
- Contact: [karkiabhijeet@gmail.com]
- Professional services available for:
  - Custom framework development
  - Integration with existing tools
  - Training and workshops

## 🚀 Roadmap

### Version 2.0 (Planned)
- [ ] Additional compliance frameworks (NIST, ISO 27001)
- [ ] PowerBI dashboard integration
- [ ] REST API for integration scenarios
- [ ] Enhanced reporting with trend analysis

### Version 1.5 (In Progress)
- [ ] GUI for custom configuration creation
- [ ] Automated remediation suggestions
- [ ] Integration with Microsoft Sentinel
- [ ] PowerShell Gallery publication

---

## 📈 Quick Reference


### Common Commands
```powershell
# Interactive menu (recommended for all users)
Invoke-PurviewConfigAnalyser

# Quick PSPF assessment
Invoke-PurviewConfigAnalyser -Mode CollectAndTest -ConfigurationName "PSPF"

# Data collection only
Invoke-PurviewConfigAnalyser -Mode CollectOnly

# Test existing data
Invoke-PurviewConfigAnalyser -Mode TestOnly -ConfigurationName "PSPF"
```

---

## ✅ Seamless Experience for Third Parties

- **No manual setup required**: Just import the module and run `Invoke-PurviewConfigAnalyser`.
- **All dependencies auto-installed** (ImportExcel, ExchangeOnlineManagement).
- **All outputs in `output/`**: CSV, Excel, logs.
- **Custom frameworks supported**: Use the interactive menu to create and test custom compliance frameworks.

For any issues, see the Troubleshooting section above.

### File Locations
- **Module**: `src/PurviewConfigAnalyser.psm1`
- **Configuration**: `config/`
- **Reports**: `output/`
- **Logs**: `output/file_runlog.txt`

### Key Concepts
- **Control Book**: Defines what to test
- **Property Book**: Defines how to test
- **Master Reference**: Template for custom configurations
- **Framework**: Complete set of controls and criteria

---

*Happy compliance testing! 🎉*
