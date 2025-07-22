function Invoke-PurviewConfigAnalyser {
    <#
    .SYNOPSIS
        Interactive Microsoft Purview Configuration Analyser - Your gateway to Purview compliance assessment.
    .NOTES
        First-time users: Simply run 'Invoke-PurviewConfigAnalyser' with no parameters!
        The interactive menu will guide you through everything you need to know.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, HelpMessage = "Operation mode - leave blank for interactive menu")]
        [ValidateSet('CollectAndTest', 'CollectOnly', 'TestOnly')]
        [string]$Mode,
        
        [Parameter(Mandatory = $false, HelpMessage = "Configuration name like 'PSPF' - leave blank to choose from menu")]
        [string]$ConfigurationName,
        
        [Parameter(Mandatory = $false, HelpMessage = "Output directory - leave blank for default location")]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false, HelpMessage = "Your Microsoft 365 username - leave blank to be prompted")]
        [string]$UserPrincipalName
    )
    
    Write-Host "=== Microsoft Purview Configuration Analyser ===" -ForegroundColor Cyan
    Write-Host "Version: 1.0.0" -ForegroundColor Gray
    Write-Host "Start Time: $(Get-Date)" -ForegroundColor White
    Write-Host ""
    
    # If Mode is not specified, show interactive menu
    if (-not $Mode) {
        Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
        return
    }
    
    # Direct mode execution (for backward compatibility)
    Write-Host "Direct Mode: $Mode" -ForegroundColor White
    Write-Host "Configuration: $ConfigurationName" -ForegroundColor White
    Write-Host ""
    
    try {
        switch ($Mode) {
            'CollectAndTest' {
                Write-Host "Step 1: Collecting Purview Configuration Data..." -ForegroundColor Yellow
                $configPath = Get-PurviewConfig -OutputPath $OutputPath
                
                Write-Host "Step 2: Running Compliance Tests..." -ForegroundColor Yellow
                $results = Test-PurviewCompliance -ConfigurationName $ConfigurationName -OptimizedReportPath $configPath
                
                Write-Host "[SUCCESS] Collection and testing completed successfully!" -ForegroundColor Green
                return $results
            }
            
            'CollectOnly' {
                Write-Host "Collecting Purview Configuration Data..." -ForegroundColor Yellow
                $configPath = Get-PurviewConfig -OutputPath $OutputPath
                
                Write-Host "[SUCCESS] Configuration collection completed successfully!" -ForegroundColor Green
                Write-Host "Configuration saved to: $configPath" -ForegroundColor Gray
                return $configPath
            }
            
            'TestOnly' {
                Write-Host "Running Compliance Tests..." -ForegroundColor Yellow
                $results = Test-PurviewCompliance -ConfigurationName $ConfigurationName
                
                Write-Host "[SUCCESS] Compliance testing completed successfully!" -ForegroundColor Green
                return $results
            }
        }
    }
    catch {
        Write-Host "[ERROR] Operation failed: $($_.Exception.Message)" -ForegroundColor Red
        throw $_
    }
}

# Helper function to get the latest OptimizedReport JSON file from run log
function Get-LatestOptimizedReport {
    param([string]$RunLogPath, [string]$OutputPath)
    
    if (Test-Path $RunLogPath) {
        $logEntries = Get-Content $RunLogPath | Where-Object { $_ -match "OptimizedReport" }
        if ($logEntries) {
            $latestEntry = $logEntries[-1] # Get the last entry
            $fileName = ($latestEntry -split " - OptimizedReport.*?: ")[1]
            $fullPath = Join-Path $OutputPath $fileName
            if (Test-Path $fullPath) {
                return $fullPath
            }
        }
    }
    
    # Fallback: search for OptimizedReport*.json files directly
    $jsonFiles = Get-ChildItem -Path $OutputPath -Filter "OptimizedReport*.json" | Sort-Object LastWriteTime -Descending
    if ($jsonFiles) {
        return $jsonFiles[0].FullName
    }
    
    return $null
}

function Show-MainMenu {
    <#
    .SYNOPSIS
        Displays an intuitive main menu with detailed explanations for each option.
    
    .DESCRIPTION
        This function provides a comprehensive, user-friendly menu system that explains
        each option clearly to help users make informed decisions.
    
    .PARAMETER OutputPath
        The output path for generated reports
    
    .PARAMETER UserPrincipalName
        The User Principal Name for authentication
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName
    )
    
    # Show welcome message and instructions
    Write-Host "Welcome to the Microsoft Purview Configuration Analyser!" -ForegroundColor Green
    Write-Host "This tool helps you assess your Microsoft Purview compliance posture." -ForegroundColor Gray
    Write-Host ""
    Write-Host "What you can do:" -ForegroundColor Cyan
    Write-Host "  - Extract configuration data from your Microsoft Purview tenant" -ForegroundColor Gray
    Write-Host "  - Run compliance assessments against industry frameworks (like PSPF)" -ForegroundColor Gray
    Write-Host "  - Create custom control books tailored to your organization" -ForegroundColor Gray
    Write-Host "  - Generate detailed reports in CSV and Excel formats" -ForegroundColor Gray
    Write-Host ""
    
    while ($true) {
        Write-Host "+-----------------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "                         MAIN MENU - CHOOSE YOUR ACTION                     " -ForegroundColor Cyan
        Write-Host "+-----------------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  1. Extract Configuration and Run Tests                                    " -ForegroundColor White
        Write-Host "     -> Connect to your tenant, collect data, then run compliance tests   " -ForegroundColor Gray
        Write-Host "     -> Best for: Complete assessment from start to finish                " -ForegroundColor Gray
        Write-Host "                                                                           " -ForegroundColor Gray
        Write-Host "  2. Extract Configuration Only                                            " -ForegroundColor White
        Write-Host "     -> Connect to your tenant and collect configuration data             " -ForegroundColor Gray
        Write-Host "     -> Best for: Data collection without immediate testing               " -ForegroundColor Gray
        Write-Host "                                                                           " -ForegroundColor Gray
        Write-Host "  3. Run Validation Tests Only                                             " -ForegroundColor White
        Write-Host "     -> Use existing data to run compliance tests                         " -ForegroundColor Gray
        Write-Host "     -> Best for: Testing against previously collected data               " -ForegroundColor Gray
        Write-Host "                                                                           " -ForegroundColor Gray
        Write-Host "  4. Exit                                                                  " -ForegroundColor White
        Write-Host "     -> Close the application                                              " -ForegroundColor Gray
        Write-Host "+-----------------------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "[INFO] CUSTOM CONFIGURATION GUIDE" -ForegroundColor Yellow
        Write-Host "Need a custom compliance framework? Create your own control books:" -ForegroundColor Gray
        Write-Host "• Step 1: Create ControlBook_[YourName]_Config.csv with your controls" -ForegroundColor Gray
        Write-Host "• Step 2: Create ControlBook_Property_[YourName]_Config.csv for properties" -ForegroundColor Gray
        Write-Host "• Step 3: Place both files in the config/ folder" -ForegroundColor Gray
        Write-Host "• Step 4: Run Test-PurviewCompliance -Configuration `"[YourName]`"" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Tip: New users should start with option 1. See custom guide above for frameworks!" -ForegroundColor Yellow
        Write-Host ""
        $choice = Read-Host "Please select an option (1-4)"
        if ($choice -match '^[1-4]$') {
            switch ($choice) {
            '1' {
                Write-Host ""
            Write-Host "EXTRACT CONFIGURATION AND RUN TESTS" -ForegroundColor Yellow
            Write-Host "====================================" -ForegroundColor Yellow
            Write-Host "This will:" -ForegroundColor White
            Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
            Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
            Write-Host "  3. Present available compliance frameworks for testing" -ForegroundColor Gray
            Write-Host "  4. Generate comprehensive compliance reports" -ForegroundColor Gray
            Write-Host ""
            Execute-CollectAndTest -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
            }
            '2' {
                Write-Host ""
                Write-Host "EXTRACT CONFIGURATION ONLY" -ForegroundColor Yellow
                Write-Host "==========================" -ForegroundColor Yellow
                Write-Host "This will:" -ForegroundColor White
                Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
                Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
                Write-Host "  3. Save data for later analysis" -ForegroundColor Gray
                Write-Host ""
                Execute-CollectOnly -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
            }
            '3' {
                Write-Host ""
                Write-Host "RUN VALIDATION TESTS ONLY" -ForegroundColor Yellow
                Write-Host "=========================" -ForegroundColor Yellow
                Write-Host "This will:" -ForegroundColor White
                Write-Host "  1. Use existing configuration data" -ForegroundColor Gray
                Write-Host "  2. Present available compliance frameworks" -ForegroundColor Gray
                Write-Host "  3. Generate compliance assessment reports" -ForegroundColor Gray
                Write-Host ""
                Execute-TestOnly -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
            }
            '4' {
                Write-Host ""
                Write-Host "CUSTOM CONFIGURATION GUIDE" -ForegroundColor Yellow
                Write-Host "==========================" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "[INFO] How to Create Your Own Compliance Framework:" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "Step 1: Navigate to Master Control Books" -ForegroundColor White
                Write-Host "   Location: .\config\MasterControlBooks\" -ForegroundColor Gray
                Write-Host ""
                Write-Host "Step 2: Create Your Control Book Files" -ForegroundColor White
                Write-Host "   Create two new CSV files with your chosen name:" -ForegroundColor Gray
                Write-Host "   • ControlBook_[YourName]_Config.csv" -ForegroundColor Yellow
                Write-Host "   • ControlBook_Property_[YourName]_Config.csv" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "[EXAMPLE] Examples:" -ForegroundColor Cyan
                Write-Host "   • ControlBook_Healthcare_Config.csv" -ForegroundColor Gray
                Write-Host "   • ControlBook_Property_Healthcare_Config.csv" -ForegroundColor Gray
                Write-Host "   • ControlBook_Banking_Config.csv" -ForegroundColor Gray
                Write-Host "   • ControlBook_Property_Banking_Config.csv" -ForegroundColor Gray
                Write-Host ""
                Write-Host "Step 3: Configure Your Controls" -ForegroundColor White
                Write-Host "   In the main control book (ControlBook_[Name]_Config.csv):" -ForegroundColor Gray
                Write-Host "   • Define your control IDs, capabilities, and descriptions" -ForegroundColor Gray
                Write-Host ""
                Write-Host "   In the property book (ControlBook_Property_[Name]_Config.csv):" -ForegroundColor Gray
                Write-Host "   • Set properties to test for each control ID" -ForegroundColor Gray
                Write-Host "   • Define expected values and requirements" -ForegroundColor Gray
                Write-Host ""
                Write-Host "[WARNING]  Important: Control IDs must match between both files" -ForegroundColor Red
                Write-Host ""
                Write-Host "Step 4: Run Your Custom Assessment" -ForegroundColor White
                Write-Host "   Test-PurviewCompliance -Configuration `"[YourName]`"" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "[TIP] Pro Tip: Copy existing AUGov files as templates!" -ForegroundColor Green
                Write-Host ""
            }
            '4' {
                Write-Host ""
                Write-Host "EXITING APPLICATION" -ForegroundColor Green
                Write-Host "===================" -ForegroundColor Green
                Write-Host "Thank you for using Microsoft Purview Configuration Analyser!" -ForegroundColor Cyan
                Write-Host ""
                return
            }
        }
        
            if ($choice -ne '4') {
                Write-Host ""
                Write-Host "---------------------------------------------------------------" -ForegroundColor DarkGray
                Write-Host "Press any key to return to the main menu..." -ForegroundColor Cyan
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                Write-Host ""
            } else {
                break
            }
        } else {
            Write-Host "Invalid input. Please enter a number between 1 and 4." -ForegroundColor Red
        }
    }
}

function Execute-CollectAndTest {
    <#
    .SYNOPSIS
        Executes the collect and test workflow.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "EXTRACT CONFIGURATION and RUN TESTS" -ForegroundColor Cyan
        Write-Host "==================================" -ForegroundColor Cyan
        Write-Host "This will:" -ForegroundColor White
        Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
        Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
        Write-Host "  3. Present available compliance frameworks for testing" -ForegroundColor Gray
        Write-Host "  4. Generate comprehensive compliance reports" -ForegroundColor Gray
        Write-Host ""

        Write-Host "Step 1: Collecting Purview Configuration Data..." -ForegroundColor Yellow

        # Use the standalone script which handles authentication and dependencies
        $dataCollectionScript = "$PSScriptRoot\..\Collect-PurviewConfiguration.ps1"
        if (-not (Test-Path $dataCollectionScript)) {
            throw "Data collection script not found at: $dataCollectionScript"
        }

        & $dataCollectionScript

        # Get the latest OptimizedReport JSON file  
        $configBasePath = "$PSScriptRoot\..\config"
        $outputBasePath = "$PSScriptRoot\..\output"

        $optimizedReportPath = Get-LatestOptimizedReport -RunLogPath "$outputBasePath\file_runlog.txt" -OutputPath $outputBasePath

        if (-not $optimizedReportPath -or -not (Test-Path $optimizedReportPath)) {
            throw "OptimizedReport JSON file was not found after data collection"
        }

        Write-Host "Configuration collection completed successfully!" -ForegroundColor Green
        $reportSize = (Get-Item $optimizedReportPath).Length / 1MB
        $sizeMB = [math]::Round($reportSize, 2)
    Write-Host ("   Using OptimizedReport: {0} ({1}) MB" -f (Split-Path -Leaf $optimizedReportPath), $sizeMB) -ForegroundColor Gray
        Write-Host ""

        Write-Host "Step 2: Select Validation Test Configuration..." -ForegroundColor Yellow
        $selectedConfig = Show-ValidationConfigurationMenu

        if ($selectedConfig -eq 'BackToMainMenu') {
            Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
            return
        }

        if ($selectedConfig) {
            Write-Host "Step 3: Running Validation Tests..." -ForegroundColor Yellow

            # Use the Run-MaturityAssessment.ps1 script (it's in the src/Scripts folder)
            $assessmentScript = "$PSScriptRoot\..\Scripts\Run-MaturityAssessment.ps1"
            if (-not (Test-Path $assessmentScript)) {
                throw "Assessment script not found at: $assessmentScript"
            }

            & $assessmentScript -ConfigurationName $selectedConfig -SkipDataCollection -GenerateExcel

            Write-Host "Collection and testing completed successfully!" -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("Operation failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host "Press any key to return to the main menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
    }
}

function Execute-CollectOnly {
    <#
    .SYNOPSIS
        Executes the collect only workflow.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "EXTRACT CONFIGURATION ONLY" -ForegroundColor Yellow
        Write-Host "==========================" -ForegroundColor Yellow
        Write-Host "This will:" -ForegroundColor White
        Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
        Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
        Write-Host "  3. Save the data for later testing" -ForegroundColor Gray
        Write-Host ""

        Write-Host "Collecting Purview Configuration Data..." -ForegroundColor Yellow

        # Use the standalone script which handles authentication and dependencies
        $dataCollectionScript = "$PSScriptRoot\..\Collect-PurviewConfiguration.ps1"
        if (-not (Test-Path $dataCollectionScript)) {
            throw "Data collection script not found at: $dataCollectionScript"
        }

        & $dataCollectionScript

        # Get the latest OptimizedReport JSON file to confirm success
        $outputBasePath = "$PSScriptRoot\..\..\output"
        $optimizedReportPath = Get-LatestOptimizedReport -RunLogPath "$outputBasePath\file_runlog.txt" -OutputPath $outputBasePath

        if ($optimizedReportPath -and (Test-Path $optimizedReportPath)) {
            Write-Host "Configuration collection completed successfully!" -ForegroundColor Green
            $reportSize = (Get-Item $optimizedReportPath).Length / 1MB
            $sizeMB = [math]::Round($reportSize, 2)
            Write-Host ("   Configuration saved to: {0} ({1}) MB" -f (Split-Path -Leaf $optimizedReportPath), $sizeMB) -ForegroundColor Gray
        } else {
            Write-Host "Configuration collection completed, but OptimizedReport file not found" -ForegroundColor Yellow
        }
    }
    catch {
    Write-Host "Operation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "----------------------------------------------------------------" -ForegroundColor Gray
    Write-Host "Press any key to return to the main menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
    }
}

function Execute-TestOnly {
    <#
    .SYNOPSIS
        Executes the test only workflow.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-Host "[SUCCESS] RUN VALIDATION TESTS ONLY" -ForegroundColor Green
        Write-Host "══════════════════════════════" -ForegroundColor Green
        Write-Host "This will:" -ForegroundColor White
        Write-Host "  1. Use existing Purview configuration data" -ForegroundColor Gray
        Write-Host "  2. Run compliance tests against selected framework" -ForegroundColor Gray
        Write-Host "  3. Generate detailed compliance reports" -ForegroundColor Gray
        Write-Host ""
        
        Write-Host "Select Validation Test Configuration..." -ForegroundColor Yellow
        $selectedConfig = Show-ValidationConfigurationMenu
        
        if ($selectedConfig -eq 'BackToMainMenu') {
            Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
            return
        }
        
        if ($selectedConfig) {
            Write-Host "Running Validation Tests..." -ForegroundColor Yellow
            
            # Use the Run-MaturityAssessment.ps1 script with SkipDataCollection flag
            $assessmentScript = "$PSScriptRoot\..\Scripts\Run-MaturityAssessment.ps1"
            if (-not (Test-Path $assessmentScript)) {
                throw "Assessment script not found at: $assessmentScript"
            }
            
            & $assessmentScript -ConfigurationName $selectedConfig -SkipDataCollection -GenerateExcel
            
            Write-Host "[SUCCESS] Validation tests completed successfully!" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "[ERROR] Operation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host "🔄 Press any key to return to the main menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Show-MainMenu -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
    }
}

function Show-ConfigurationMenu {
    <#
    .SYNOPSIS
        Displays an intuitive configuration selection menu with detailed explanations.
    
    .DESCRIPTION
        This function scans for available configuration files and presents them
        in a user-friendly format with explanations of what each configuration does.
    
    .OUTPUTS
        Returns the selected configuration name or 'CreateCustom' if user chooses to create a custom configuration.
    #>
    [CmdletBinding()]
    param()
    
    # Get module root path
    $moduleRoot = Split-Path -Parent $PSScriptRoot
    $configPath = Join-Path -Path $moduleRoot -ChildPath "config"
    
    # Scan for available configuration files
    $configFiles = Get-ChildItem -Path $configPath -Filter "ControlBook_*_Config.csv" | Where-Object { $_.Name -notmatch "Property" }
    
    if ($configFiles.Count -eq 0) {
        Write-Host "[ERROR] No configuration files found in $configPath" -ForegroundColor Red
        Write-Host "Please ensure configuration files are properly installed." -ForegroundColor Yellow
        return $null
    }
    
    # Extract configuration names and create descriptions
    $configurations = @()
    $configDescriptions = @{
        'PSPF' = 'Protective Security Policy Framework - Australian Government security standard'
        'NIST' = 'National Institute of Standards and Technology Cybersecurity Framework'
        'ISO27001' = 'International Organization for Standardization 27001 Information Security'
        'Custom' = 'Organization-specific compliance framework'
    }
    
    foreach ($file in $configFiles) {
        $configName = $file.Name -replace "ControlBook_", "" -replace "_Config\.csv", ""
        $configurations += $configName
    }
    
    Write-Host "SELECT COMPLIANCE FRAMEWORK" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════" -ForegroundColor Yellow
    Write-Host "Choose the compliance framework you want to assess against:" -ForegroundColor White
    Write-Host ""

    Write-Host "┌─────────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
    Write-Host "│                        AVAILABLE COMPLIANCE FRAMEWORKS                          │" -ForegroundColor Cyan
    Write-Host "├─────────────────────────────────────────────────────────────────────────────────┤" -ForegroundColor Cyan
    
    $optionNumber = 1
    foreach ($config in $configurations) {
        $description = $configDescriptions[$config]
        if (-not $description) {
            $description = "Custom configuration for $config requirements"
        }
        
    Write-Host "│  $optionNumber. $config" -ForegroundColor White
    Write-Host "│     - $description" -ForegroundColor Gray
    Write-Host "│" -ForegroundColor Gray
        $optionNumber++
    }
    
    Write-Host "│  $optionNumber. Create New Custom Configuration" -ForegroundColor White
    Write-Host "│     - Build your own control book for specific requirements" -ForegroundColor Gray
    Write-Host "│" -ForegroundColor Gray
    Write-Host "│  $($optionNumber + 1). Cancel and Return to Main Menu" -ForegroundColor White
    Write-Host "│     - Go back without making a selection" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Tip: Each framework has different controls and requirements. Choose the one that matches your organization's needs." -ForegroundColor Yellow
    Write-Host ""
    
    do {
        $choice = Read-Host "Please select a configuration (1-$($optionNumber + 1))"
        
        if ($choice -match '^\d+$') {
            $choiceNum = [int]$choice
            if ($choiceNum -ge 1 -and $choiceNum -le $configurations.Count) {
                $selectedConfig = $configurations[$choiceNum - 1]
                Write-Host ""
                Write-Host "[SUCCESS] Selected: $selectedConfig" -ForegroundColor Green
                $description = $configDescriptions[$selectedConfig]
                if ($description) {
                    Write-Host "   Framework: $description" -ForegroundColor Gray
                }
                return $selectedConfig
            }
            elseif ($choiceNum -eq ($configurations.Count + 1)) {
                Write-Host ""
                Write-Host "Redirecting to Custom Configuration Creator..." -ForegroundColor Yellow
                return 'CreateCustom'
            }
            elseif ($choiceNum -eq ($configurations.Count + 2)) {
                Write-Host ""
                Write-Host "Returning to main menu..." -ForegroundColor Gray
                return $null
            }
        }
        
    Write-Host "Invalid input. Please enter a number between 1 and $($optionNumber + 1)." -ForegroundColor Red
        
    } while ($true)
}

function Show-ValidationConfigurationMenu {
    <#
    .SYNOPSIS
        Shows available configurations for validation testing after data collection.
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "[INFO] SELECT VALIDATION CONFIGURATION" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════" -ForegroundColor Yellow
    Write-Host "Choose which compliance framework you want to validate against:" -ForegroundColor White
    Write-Host ""
    
    # Get available configurations
    $configBasePath = "$PSScriptRoot\..\config"
    $availableConfigs = @()
    
    # Look for available control book configurations (ignoring MasterControlBooks folder)
    $configFiles = Get-ChildItem "$configBasePath\ControlBook_*_Config.csv" | Where-Object { $_.Name -notmatch "Property" }
    
    foreach ($file in $configFiles) {
        $configName = $file.Name -replace "ControlBook_|_Config\.csv", ""
        
        # Check if corresponding property config exists
        $propertyConfig = "$configBasePath\ControlBook_Property_${configName}_Config.csv"
        if (Test-Path $propertyConfig) {
            $availableConfigs += $configName
        }
    }
    
    if ($availableConfigs.Count -eq 0) {
        Write-Host "[ERROR] No validation configurations found in the config directory." -ForegroundColor Red
        Write-Host "   Please ensure configuration files are properly set up." -ForegroundColor Yellow
        return $null
    }
    
    # Display menu
    Write-Host "┌─────────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor Gray
    Write-Host "│                        AVAILABLE VALIDATION CONFIGURATIONS                        │" -ForegroundColor Gray
    Write-Host "├─────────────────────────────────────────────────────────────────────────────────┤" -ForegroundColor Gray
    
    $optionNumber = 1
    foreach ($config in $availableConfigs) {
        $description = switch ($config) {
            "PSPF" { "Australian Government Protective Security Policy Framework" }
            default { "Custom compliance framework: $config" }
        }
        
    Write-Host "│  $optionNumber. $config" -ForegroundColor Gray
    Write-Host "│     - $description" -ForegroundColor Gray
    Write-Host "│" -ForegroundColor Gray
        $optionNumber++
    }
    
    Write-Host "│  $optionNumber. Back to Main Menu" -ForegroundColor Gray
    Write-Host "│     - Return to main menu to create custom configuration (Option 4)" -ForegroundColor Gray
    Write-Host "└─────────────────────────────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Note: To create a new custom configuration, go back to the main menu and select Option 4." -ForegroundColor Cyan
    Write-Host ""
    
    do {
        $choice = Read-Host "Please select an option (1-$optionNumber)"
        
        if ($choice -match '^\d+$') {
            $choiceNum = [int]$choice
            if ($choiceNum -ge 1 -and $choiceNum -le $availableConfigs.Count) {
                $selectedConfig = $availableConfigs[$choiceNum - 1]
                Write-Host "[SUCCESS] Selected configuration: $selectedConfig" -ForegroundColor Green
                return $selectedConfig
            } elseif ($choiceNum -eq $optionNumber) {
                Write-Host "[BACK] Returning to main menu..." -ForegroundColor Gray
                return 'BackToMainMenu'
            }
        }
        
        Write-Host "[ERROR] Invalid input. Please enter a number between 1 and $optionNumber." -ForegroundColor Red
        
    } while ($true)
}

function Execute-CreateCustomConfig {
    <#
    .SYNOPSIS
        Interactive form-based interface for creating custom control book configurations.
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "CREATE CUSTOM CONFIGURATION" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════" -ForegroundColor Yellow
    Write-Host "This interactive wizard will help you create a custom control book tailored to your organization's specific requirements." -ForegroundColor White
    Write-Host ""
    Write-Host "What is a Control Book?" -ForegroundColor Cyan
    Write-Host "  - A control book defines the compliance rules and checks for your organization" -ForegroundColor Gray
    Write-Host "  - It contains specific controls that will be tested against your Purview configuration" -ForegroundColor Gray
    Write-Host "  - Each control has criteria that determine if your setup is compliant" -ForegroundColor Gray
    Write-Host ""
    Write-Host "How this works:" -ForegroundColor Cyan
    Write-Host "  - You'll be shown controls grouped by capability (Sensitivity Labels, DLP, etc.)" -ForegroundColor Gray
    Write-Host "  - For each control, you can accept the default value or provide your own" -ForegroundColor Gray
    Write-Host "  - Required fields are marked with [REQUIRED] and must be filled" -ForegroundColor Gray
    Write-Host "  - A new configuration will be created with your custom settings" -ForegroundColor Gray
    Write-Host ""
    
    # Confirm user wants to proceed
    do {
    $proceed = Read-Host "Do you want to proceed with creating a custom configuration? (Y/N)"
        if ($proceed -match '^[Yy]') {
            break
        } elseif ($proceed -match '^[Nn]') {
            Write-Host "[ERROR] Custom configuration creation cancelled." -ForegroundColor Yellow
            return
        } else {
            Write-Host "[ERROR] Please enter Y for Yes or N for No." -ForegroundColor Red
        }
    } while ($true)
    
    try {
        # Load reference files
        $referenceBasePath = "$PSScriptRoot\..\config\MasterControlBooks"
        $controlBookReference = Import-Csv "$referenceBasePath\ControlBook_Reference.csv"
        $propertyReference = Import-Csv "$referenceBasePath\ControlBook_Property_Reference.csv"
        
        Write-Host ""
    Write-Host "CONFIGURATION DETAILS" -ForegroundColor Yellow
    Write-Host "──────────────────────────" -ForegroundColor Yellow
        
        # Get configuration name
        do {
            $configName = Read-Host "Enter a name for your custom configuration (e.g., 'CustomOrg', 'ACME_Corp')"
            if ([string]::IsNullOrWhiteSpace($configName)) {
                Write-Host "[ERROR] Configuration name cannot be empty. Please try again." -ForegroundColor Red
            } elseif ($configName -match '[^a-zA-Z0-9_]') {
                Write-Host "[ERROR] Configuration name can only contain letters, numbers, and underscores. Please try again." -ForegroundColor Red
            } else {
                # Check if configuration already exists
                $configPath = "$PSScriptRoot\..\config\ControlBook_${configName}_Config.csv"
                if (Test-Path $configPath) {
                    $overwrite = Read-Host ("Configuration '{0}' already exists. Overwrite? (Y/N)" -f $configName)
                    if ($overwrite -match '^[Yy]') {
                        break
                    }
                } else {
                    break
                }
            }
        } while ($true)
        
        Write-Host ""
    Write-Host "INTERACTIVE CONFIGURATION BUILDER" -ForegroundColor Cyan
    Write-Host "════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "For each control, you can:" -ForegroundColor White
    Write-Host "  - Press ENTER to accept the default value" -ForegroundColor Gray
    Write-Host "  - Type a new value to override the default" -ForegroundColor Gray
    Write-Host "  - [REQUIRED] fields must be filled with appropriate values" -ForegroundColor Gray
    Write-Host ""
        
        # Group controls by capability
        $capabilities = $controlBookReference | Group-Object -Property Capability
        
        $customControls = @()
        $customProperties = @()
        
        while ($true) {
            Write-Host "┌─────────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
            Write-Host "│                           MAIN MENU - CHOOSE YOUR ACTION                           │" -ForegroundColor Cyan
            Write-Host "├─────────────────────────────────────────────────────────────────────────────────┤" -ForegroundColor Cyan
            Write-Host "│  1. Extract Configuration and Run Tests                                          │" -ForegroundColor White
            Write-Host "│     -> Connect to your tenant, collect data, then run compliance tests           │" -ForegroundColor Gray
            Write-Host "│     -> Best for: Complete assessment from start to finish                        │" -ForegroundColor Gray
            Write-Host "│                                                                                 │" -ForegroundColor Gray
            Write-Host "│  2. Extract Configuration Only                                                   │" -ForegroundColor White
            Write-Host "│     -> Connect to your tenant and collect configuration data                     │" -ForegroundColor Gray
            Write-Host "│     -> Best for: Data collection without immediate testing                       │" -ForegroundColor Gray
            Write-Host "│                                                                                 │" -ForegroundColor Gray
            Write-Host "│  3. Run Validation Tests Only                                                    │" -ForegroundColor White
            Write-Host "│     -> Use existing data to run compliance tests                                 │" -ForegroundColor Gray
            Write-Host "│     -> Best for: Testing against previously collected data                       │" -ForegroundColor Gray
            Write-Host "│                                                                                 │" -ForegroundColor Gray
            Write-Host "│  4. Create Custom Configuration                                                  │" -ForegroundColor White
            Write-Host "│     -> Build your own control book for organization-specific requirements        │" -ForegroundColor Gray
            Write-Host "│     -> Best for: Custom compliance frameworks                                    │" -ForegroundColor Gray
            Write-Host "│                                                                                 │" -ForegroundColor Gray
            Write-Host "│  5. Exit                                                                        │" -ForegroundColor White
            Write-Host "│     -> Close the application                                                    │" -ForegroundColor Gray
            Write-Host "└─────────────────────────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Tip: If you're new to this tool, start with option 1 for a complete assessment!" -ForegroundColor Yellow
            Write-Host ""
            $choice = Read-Host "Please select an option (1-5)"
            if ($choice -match '^[1-5]$') {
                if ($choice -eq '1') {
                    Write-Host ""
                    Write-Host "EXTRACT CONFIGURATION AND RUN TESTS" -ForegroundColor Yellow
                    Write-Host "====================================" -ForegroundColor Yellow
                    Write-Host "This will:" -ForegroundColor White
                    Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
                    Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
                    Write-Host "  3. Present available compliance frameworks for testing" -ForegroundColor Gray
                    Write-Host "  4. Generate comprehensive compliance reports" -ForegroundColor Gray
                    Write-Host ""
                    Execute-CollectAndTest -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
                } elseif ($choice -eq '2') {
                    Write-Host ""
                    Write-Host "EXTRACT CONFIGURATION ONLY" -ForegroundColor Yellow
                    Write-Host "==========================" -ForegroundColor Yellow
                    Write-Host "This will:" -ForegroundColor White
                    Write-Host "  1. Connect to your Microsoft 365 tenant" -ForegroundColor Gray
                    Write-Host "  2. Extract Purview configuration data" -ForegroundColor Gray
                    Write-Host "  3. Save data for later analysis" -ForegroundColor Gray
                    Write-Host ""
                    Execute-CollectOnly -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
                } elseif ($choice -eq '3') {
                    Write-Host ""
                    Write-Host "RUN VALIDATION TESTS ONLY" -ForegroundColor Yellow
                    Write-Host "=========================" -ForegroundColor Yellow
                    Write-Host "This will:" -ForegroundColor White
                    Write-Host "  1. Use existing configuration data" -ForegroundColor Gray
                    Write-Host "  2. Present available compliance frameworks" -ForegroundColor Gray
                    Write-Host "  3. Generate compliance assessment reports" -ForegroundColor Gray
                    Write-Host ""
                    Execute-TestOnly -OutputPath $OutputPath -UserPrincipalName $UserPrincipalName
                } elseif ($choice -eq '4') {
                    Write-Host ""
                    Write-Host "CREATE CUSTOM CONFIGURATION" -ForegroundColor Yellow
                    Write-Host "===========================" -ForegroundColor Yellow
                    Write-Host "This will:" -ForegroundColor White
                    Write-Host "  1. Launch the Windows Forms GUI to create a custom control book" -ForegroundColor Gray
                    Write-Host "  2. Allow you to define organization-specific controls" -ForegroundColor Gray
                    Write-Host "  3. Save your configuration for future use" -ForegroundColor Gray
                    Write-Host ""
                    $guiScript = Join-Path $PSScriptRoot 'Show-PurviewConfigAnalyserGUI.ps1'
                    if (Test-Path $guiScript) {
                        . $guiScript
                        Show-PurviewConfigAnalyserGUI
                    } else {
                        Write-Host "GUI script not found at: $guiScript" -ForegroundColor Red
                    }
                } elseif ($choice -eq '5') {
                    Write-Host ""
                    Write-Host "EXITING APPLICATION" -ForegroundColor Green
                    Write-Host "===================" -ForegroundColor Green
                    Write-Host "Thank you for using the Microsoft Purview Configuration Analyser!" -ForegroundColor Green
                    Write-Host "Your compliance journey continues..." -ForegroundColor Gray
                    Write-Host ""
                    break
                }
                if ($choice -ne '5') {
                    Write-Host ""
                    Write-Host "---------------------------------------------------------------" -ForegroundColor DarkGray
                    Write-Host "Press any key to return to the main menu..." -ForegroundColor Cyan
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    Write-Host ""
                }
            } else {
                Write-Host "Invalid input. Please enter a number between 1 and 5." -ForegroundColor Red
            }
        }
        Write-Host "Total Controls: $($customControls.Count)" -ForegroundColor White
        Write-Host "Active Controls: $($customControls.Count)" -ForegroundColor White
        Write-Host ""
    Write-Host "Your custom configuration '$configName' is now available for testing!" -ForegroundColor Green
    Write-Host "You can now run validation tests against this configuration." -ForegroundColor Gray
        
        # Ask if user wants to test the configuration
        Write-Host ""
        do {
            $testNow = Read-Host "Would you like to test this configuration now? (Y/N)"
            if ($testNow -match '^[Yy]') {
                Write-Host ""
                Write-Host "TESTING CUSTOM CONFIGURATION" -ForegroundColor Cyan
                Write-Host "═══════════════════════════════════" -ForegroundColor Cyan
                
                # Run the maturity assessment with the new configuration
                $assessmentScript = "$PSScriptRoot\..\Scripts\Run-MaturityAssessment.ps1"
                if (Test-Path $assessmentScript) {
                    & $assessmentScript -ConfigurationName $configName -SkipDataCollection -GenerateExcel
                } else {
                    Write-Host "[ERROR] Assessment script not found. Please run tests manually." -ForegroundColor Red
                }
                break
            } elseif ($testNow -match '^[Nn]') {
                Write-Host "Configuration saved. You can test it later by selecting 'Run Validation Tests Only' from the main menu." -ForegroundColor Gray
                break
            } else {
                Write-Host "[ERROR] Please enter Y for Yes or N for No." -ForegroundColor Red
            }
        } while ($true)
        
    } catch {
        Write-Host "[ERROR] Error creating custom configuration: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please check that the reference files exist and try again." -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Press any key to return to the main menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ...existing code...
