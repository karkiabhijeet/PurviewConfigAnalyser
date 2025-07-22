function New-CustomControlBook {
    <#
    .SYNOPSIS
        Creates a new custom control book configuration based on master reference files.
    
    .DESCRIPTION
        Copies the master control book reference files and creates customized versions
        for specific assessment scenarios.
    
    .PARAMETER ConfigurationName
        Name for the new custom configuration
    
    .PARAMETER OutputPath
        Output directory for the new control book files (defaults to module config directory)
    
    .PARAMETER BasedOn
        Base configuration to copy from (default: Reference)
    
    .EXAMPLE
        New-CustomControlBook -ConfigurationName "CustomAssessment"
        
        Creates new control book files based on the master reference files.
    
    .EXAMPLE
        New-CustomControlBook -ConfigurationName "SOX" -BasedOn "PSPF"
        
        Creates new control book files based on the PSPF configuration.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigurationName,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$BasedOn = "Reference"
    )
    
    # Get module root and configuration paths
    $ModuleRoot = $PSScriptRoot | Split-Path -Parent | Split-Path -Parent | Split-Path -Parent
    $ConfigPath = Join-Path $ModuleRoot "config"
    $MasterControlBooksPath = Join-Path $ConfigPath "MasterControlBooks"
    
    # Set output path if not provided
    if (-not $OutputPath) {
        $OutputPath = $ConfigPath
    }
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    Write-Host "Creating custom control book configuration: $ConfigurationName" -ForegroundColor Yellow
    
    try {
        # Determine source files based on BasedOn parameter
        if ($BasedOn -eq "Reference") {
            $SourceControlBook = Join-Path $MasterControlBooksPath "ControlBook_Reference.csv"
            $SourcePropertyBook = Join-Path $MasterControlBooksPath "ControlBook_Property_Reference.csv"
        } else {
            $SourceControlBook = Join-Path $ConfigPath "ControlBook_${BasedOn}_Config.csv"
            $SourcePropertyBook = Join-Path $ConfigPath "ControlBook_Property_${BasedOn}_Config.csv"
        }
        
        # Validate source files exist
        if (-not (Test-Path $SourceControlBook)) {
            throw "Source control book file not found: $SourceControlBook"
        }
        
        if (-not (Test-Path $SourcePropertyBook)) {
            throw "Source property book file not found: $SourcePropertyBook"
        }
        
        # Define target files
        $TargetControlBook = Join-Path $OutputPath "ControlBook_${ConfigurationName}_Config.csv"
        $TargetPropertyBook = Join-Path $OutputPath "ControlBook_Property_${ConfigurationName}_Config.csv"
        
        # Check if target files already exist
        if ((Test-Path $TargetControlBook) -or (Test-Path $TargetPropertyBook)) {
            $Response = Read-Host "Configuration '$ConfigurationName' already exists. Overwrite? (y/N)"
            if ($Response -notmatch "^[Yy]") {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                return
            }
        }
        
        # Copy the files
        Copy-Item -Path $SourceControlBook -Destination $TargetControlBook -Force
        Copy-Item -Path $SourcePropertyBook -Destination $TargetPropertyBook -Force
        
        Write-Host "✅ Custom control book configuration created successfully!" -ForegroundColor Green
        Write-Host "  Control Book: $TargetControlBook" -ForegroundColor Gray
        Write-Host "  Property Book: $TargetPropertyBook" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor White
        Write-Host "1. Edit the CSV files to customize controls and properties" -ForegroundColor Gray
        Write-Host "2. Run assessment using: Invoke-PurviewConfigAnalyser -Mode TestOnly -Configuration $ConfigurationName" -ForegroundColor Gray
        
        return @{
            ConfigurationName = $ConfigurationName
            ControlBookPath = $TargetControlBook
            PropertyBookPath = $TargetPropertyBook
            BasedOn = $BasedOn
        }
        
    } catch {
        Write-Host "❌ Failed to create custom control book: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
