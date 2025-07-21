# DLP: Evaluate DLP Compliance Policy property
function Test-GetDlpCompliancePolicyProperty {
    param($Result, $PropertyParts, $ExpectedValue, $OptimizedReport)
    $policies = $OptimizedReport.GetDlpCompliancePolicy
    if (-not $policies) {
        $Result.Comments = "No DLP policies found in report"
        return $Result
    }
    $propertyName = $PropertyParts[1].Trim()
    switch ($propertyName) {
        "Mode" {
            $found = $policies | Where-Object { $_.Mode -eq $ExpectedValue }
            if ($found) {
                $Result.Pass = $true
                $Result.Comments = "Found DLP policy with Mode: $ExpectedValue"
            } else {
                $Result.Comments = "No DLP policy found with Mode: $ExpectedValue"
            }
        }
        default {
            $Result.Comments = "Unknown DLP policy property: $propertyName"
        }
    }
    return $Result
}

# DLP: Evaluate DLP Compliance Rule property
function Test-GetDlpComplianceRuleProperty {
    param($Result, $PropertyParts, $ExpectedValue, $OptimizedReport)
    $rules = $OptimizedReport.GetDlpComplianceRule
    $policies = $OptimizedReport.GetDlpCompliancePolicy
    
    if (-not $rules) {
        $Result.Comments = "No DLP rules found in report"
        return $Result
    }

    # Add state to track which rule we're working with for this control
    if (-not $script:CurrentDlpControlRules) {
        $script:CurrentDlpControlRules = @{}
    }
    $controlId = $Result.ControlID

    # Filter rules to only those from enabled policies
    $enabledRules = @()
    $simulationRules = @()
    
    foreach ($rule in $rules) {
        if ($rule.PSObject.Properties.Name -contains "Policy") {
            $parentPolicy = $policies | Where-Object { $_.Guid -eq $rule.Policy }
            if ($parentPolicy) {
                if ($parentPolicy.Mode -eq "Enable") {
                    $enabledRules += $rule
                } elseif ($parentPolicy.Mode -like "Test*") {
                    $simulationRules += $rule
                }
            }
        }
    }

    $propertyName = $PropertyParts[1].Trim()
    
    # Reconstruct the full property path for compound DLP properties
    # PropertyParts[1] might be "AdvancedRule >> Sensitivetypes" and PropertyParts[2] might be "Mincount"
    if ($PropertyParts.Count -gt 2 -and $propertyName -like "AdvancedRule*") {
        # Rebuild the compound property path
        $additionalParts = @()
        for ($i = 2; $i -lt $PropertyParts.Count; $i++) {
            $additionalParts += $PropertyParts[$i].Trim()
        }
        $propertyName = $propertyName + " > " + ($additionalParts -join " > ")
    }
    
    # Handle AdvancedRule properties with >> syntax
    if ($propertyName -like "AdvancedRule*") {
        # Parse AdvancedRule properties - need to handle compound paths correctly
        # Example: "AdvancedRule >> Sensitivetypes > Mincount" should extract "Sensitivetypes > Mincount"
        if ($propertyName -like "AdvancedRule >> *") {
            # For compound paths, extract everything after "AdvancedRule >> "
            $deepProp = $propertyName -replace '^AdvancedRule >> ', ''
            $deepVal = $ExpectedValue
            $found = $false
            $foundRule = $null
            
            # Check if we already have a matched rule for this control
            if ($script:CurrentDlpControlRules.ContainsKey($controlId)) {
                $previousRule = $script:CurrentDlpControlRules[$controlId]
                if ($previousRule.PSObject.Properties.Name -contains "AdvancedRule" -and $previousRule.AdvancedRule) {
                    try {
                        $adv = $previousRule.AdvancedRule | ConvertFrom-Json -ErrorAction Stop
                        $names = Get-DlpDeepProperty -Node $adv -Target $deepProp
                        if ($deepVal -eq 'Not Null') {
                            if ($names.Count -gt 0) { 
                                $Result.Pass = $true
                                $Result.Comments = "Previously matched DLP rule has $deepProp configured in AdvancedRule"
                                return $Result
                            }
                        } else {
                            if ($names -contains $deepVal) { 
                                $Result.Pass = $true
                                $Result.Comments = "Previously matched DLP rule has $deepProp = $deepVal in AdvancedRule"
                                return $Result
                            }
                        }
                    } catch {
                        # Continue to search other rules
                    }
                }
            }
            
            # Search enabled rules for AdvancedRule content (and disabled rules for specific controls)
            $rulesToSearch = $enabledRules
            if ($controlId -eq "DLP_4.6" -or $controlId -eq "DLP_4.7" -or $controlId -eq "DLP_4.8") {
                # For these specific controls, also search disabled rules since target rule might be disabled
                $rulesToSearch = $rules
            }
            
            foreach ($rule in $rulesToSearch) {
                if ($rule.PSObject.Properties.Name -contains "AdvancedRule" -and $rule.AdvancedRule) {
                    try {
                        $adv = $rule.AdvancedRule | ConvertFrom-Json -ErrorAction Stop
                        
                        # Use enhanced DLP parsing for compound properties (DLP_4.6, 4.7, 4.8)
                        $isCompoundProp = $deepProp -like "*>*"
                        $isTargetControl = ($controlId -eq "DLP_4.6" -or $controlId -eq "DLP_4.7" -or $controlId -eq "DLP_4.8")
                        
                        if ($isCompoundProp -and $isTargetControl) {
                            try {
                                # Load enhanced parser functions
                                . "$PSScriptRoot\DlpAdvancedParser.ps1"
                                $parseResult = Test-DlpAdvancedRuleProperty -AdvRule $adv -DeepProp $deepProp -DeepVal $deepVal -ControlId $controlId
                                if ($parseResult.Found) {
                                    $found = $true
                                    $foundRule = $rule
                                    break
                                }
                            } catch {
                                # Continue to next rule
                            }
                        } else {
                            # Use existing logic for simple properties
                            $names = Get-DlpDeepProperty -Node $adv -Target $deepProp
                            
                            if ($deepVal -eq 'Not Null') {
                                if ($names.Count -gt 0) { 
                                    $found = $true
                                    $foundRule = $rule
                                    break
                                }
                            } elseif ($deepProp -like "*MinCount" -and $deepVal -match '^\d+$') {
                                # Handle numeric MinCount comparison (>= expected value)
                                $targetMinCount = [int]$deepVal
                                $highestMinCount = if ($names.Count -gt 0) { ($names | Measure-Object -Maximum).Maximum } else { 0 }
                                if ($highestMinCount -ge $targetMinCount) {
                                    $found = $true
                                    $foundRule = $rule
                                    break
                                }
                            } elseif ($deepVal -match "," -and $deepProp -like "*Name") {
                                # Handle comma-separated list of required names
                                $requiredNames = $deepVal -split ',' | ForEach-Object { $_.Trim() }
                                $allFound = $true
                                foreach ($reqName in $requiredNames) {
                                    if ($names -notcontains $reqName) {
                                        $allFound = $false
                                        break
                                    }
                                }
                                if ($allFound) {
                                    $found = $true
                                    $foundRule = $rule
                                    break
                                }
                            } else {
                                # Standard string matching
                                if ($names -contains $deepVal) { 
                                    $found = $true
                                    $foundRule = $rule
                                    break
                                }
                            }
                        }
                    } catch {
                        # Continue to next rule
                    }
                }
            }
            
            if ($found) {
                $script:CurrentDlpControlRules[$controlId] = $foundRule
                $Result.Pass = $true
                if ($deepVal -eq 'Not Null') {
                    $Result.Comments = "Found DLP rule with $deepProp configured in AdvancedRule"
                } else {
                    $Result.Comments = "Found DLP rule with $deepProp = $deepVal in AdvancedRule"
                }
            } else {
                # Check simulation rules for informational comments
                $simulationFound = $false
                foreach ($rule in $simulationRules) {
                    if ($rule.PSObject.Properties.Name -contains "AdvancedRule" -and $rule.AdvancedRule) {
                        try {
                            $adv = $rule.AdvancedRule | ConvertFrom-Json -ErrorAction Stop
                            
                            # Use enhanced DLP parsing for compound properties (DLP_4.6, 4.7, 4.8)
                            if ($deepProp -like "*>*" -and ($controlId -eq "DLP_4.6" -or $controlId -eq "DLP_4.7" -or $controlId -eq "DLP_4.8")) {
                                # Load enhanced parser functions (already loaded above)
                                if (Get-Command Test-DlpAdvancedRuleProperty -ErrorAction SilentlyContinue) {
                                    $parseResult = Test-DlpAdvancedRuleProperty -AdvRule $adv -DeepProp $deepProp -DeepVal $deepVal -ControlId $controlId
                                    if ($parseResult.Found) { $simulationFound = $true; break }
                                }
                            } else {
                                # Use existing logic
                                $names = Get-DlpDeepProperty -Node $adv -Target $deepProp
                                
                                if ($deepVal -eq 'Not Null') {
                                    if ($names.Count -gt 0) { $simulationFound = $true; break }
                                } elseif ($deepProp -like "*MinCount" -and $deepVal -match '^\d+$') {
                                    # Handle numeric MinCount comparison (>= expected value)
                                    $targetMinCount = [int]$deepVal
                                    $highestMinCount = if ($names.Count -gt 0) { ($names | Measure-Object -Maximum).Maximum } else { 0 }
                                    if ($highestMinCount -ge $targetMinCount) { $simulationFound = $true; break }
                                } elseif ($deepVal -match "," -and $deepProp -like "*Name") {
                                    # Handle comma-separated list of required names
                                    $requiredNames = $deepVal -split ',' | ForEach-Object { $_.Trim() }
                                    $allFound = $true
                                    foreach ($reqName in $requiredNames) {
                                        if ($names -notcontains $reqName) {
                                            $allFound = $false
                                            break
                                        }
                                    }
                                    if ($allFound) { $simulationFound = $true; break }
                                } else {
                                    # Standard string matching
                                    if ($names -contains $deepVal) { $simulationFound = $true; break }
                                }
                            }
                        } catch {}
                    }
                }
                
                if ($simulationFound) {
                    $Result.Comments = "No enabled DLP rule found with $deepProp in AdvancedRule. Found matching rule(s) in simulation/test mode"
                } else {
                    $Result.Comments = "No DLP rule found with $deepProp in AdvancedRule"
                }
            }
            return $Result
        } else {
            $Result.Comments = "Invalid AdvancedRule property path. Expected format: AdvancedRule >> PropertyName"
            return $Result
        }
    }
    
    switch ($propertyName) {
        "Workload" {
            if ($script:CurrentDlpControlRules.ContainsKey($controlId)) {
                # Check if the previously matched rule has this workload
                $previousRule = $script:CurrentDlpControlRules[$controlId]
                if ($previousRule.PSObject.Properties.Name -contains "Workload") {
                    # Handle workload as either array or comma-separated string
                    $workloads = @()
                    if ($previousRule.Workload -is [array]) {
                        $workloads = $previousRule.Workload
                    } else {
                        $workloads = $previousRule.Workload -split ',' | ForEach-Object { $_.Trim() }
                    }
                    
                    if ($workloads -contains $ExpectedValue) {
                        $Result.Pass = $true
                        $Result.Comments = "Previously matched DLP rule has workload: $ExpectedValue"
                        return $Result
                    }
                }
            }

            # Look for a rule in enabled policies that has the required workload
            foreach ($rule in $enabledRules) {
                if ($rule.PSObject.Properties.Name -contains "Workload") {
                    # Handle workload as either array or comma-separated string
                    $workloads = @()
                    if ($rule.Workload -is [array]) {
                        $workloads = $rule.Workload
                    } else {
                        $workloads = $rule.Workload -split ',' | ForEach-Object { $_.Trim() }
                    }
                    
                    if ($workloads -contains $ExpectedValue) {
                        # Store this rule for the control ID
                        $script:CurrentDlpControlRules[$controlId] = $rule
                        $Result.Pass = $true
                        $Result.Comments = "Found DLP rule with workload: $ExpectedValue in enabled policy"
                        return $Result
                    }
                }
            }
            
            # Check simulation rules for informational comments
            $simulationMatches = @()
            foreach ($rule in $simulationRules) {
                if ($rule.PSObject.Properties.Name -contains "Workload") {
                    # Handle workload as either array or comma-separated string
                    $workloads = @()
                    if ($rule.Workload -is [array]) {
                        $workloads = $rule.Workload
                    } else {
                        $workloads = $rule.Workload -split ',' | ForEach-Object { $_.Trim() }
                    }
                    
                    if ($workloads -contains $ExpectedValue) {
                        $simulationMatches += $rule
                    }
                }
            }
            
            if ($simulationMatches.Count -gt 0) {
                $Result.Comments = "No enabled DLP rule found with workload: $ExpectedValue. Found $($simulationMatches.Count) rule(s) in simulation/test mode"
            } else {
                $Result.Comments = "No DLP rule found with workload: $ExpectedValue"
            }
        }
        "PrependSubject" {
            if ($script:CurrentDlpControlRules.ContainsKey($controlId)) {
                # Check if the previously matched rule has PrependSubject
                $previousRule = $script:CurrentDlpControlRules[$controlId]
                if ($previousRule.PSObject.Properties.Name -contains "PrependSubject" -and $previousRule.PrependSubject) {
                    $Result.Pass = $true
                    $Result.Comments = "Previously matched DLP rule has PrependSubject configured"
                    return $Result
                }
            }

            # Look for a rule in enabled policies that has PrependSubject configured
            foreach ($rule in $enabledRules) {
                if ($rule.PSObject.Properties.Name -contains "PrependSubject" -and $rule.PrependSubject) {
                    # Store this rule for the control ID
                    $script:CurrentDlpControlRules[$controlId] = $rule
                    $Result.Pass = $true
                    $Result.Comments = "Found DLP rule with PrependSubject configured in enabled policy"
                    return $Result
                }
            }
            
            # Check simulation rules for informational comments
            $simulationMatches = $simulationRules | Where-Object { 
                $_.PSObject.Properties.Name -contains "PrependSubject" -and $_.PrependSubject 
            }
            
            if ($simulationMatches.Count -gt 0) {
                $Result.Comments = "No enabled DLP rule found with PrependSubject configured. Found $($simulationMatches.Count) rule(s) in simulation/test mode"
            } else {
                $Result.Comments = "No DLP rule found with PrependSubject configured"
            }
        }
        "SetHeader" {
            if ($script:CurrentDlpControlRules.ContainsKey($controlId)) {
                # Check if the previously matched rule has SetHeader
                $previousRule = $script:CurrentDlpControlRules[$controlId]
                if ($previousRule.PSObject.Properties.Name -contains "SetHeader" -and $previousRule.SetHeader) {
                    $Result.Pass = $true
                    $Result.Comments = "Previously matched DLP rule has SetHeader configured"
                    return $Result
                }
            }

            # Look for a rule in enabled policies that has SetHeader configured
            foreach ($rule in $enabledRules) {
                if ($rule.PSObject.Properties.Name -contains "SetHeader" -and $rule.SetHeader) {
                    # Store this rule for the control ID
                    $script:CurrentDlpControlRules[$controlId] = $rule
                    $Result.Pass = $true
                    $Result.Comments = "Found DLP rule with SetHeader configured in enabled policy"
                    return $Result
                }
            }
            
            # Check simulation rules for informational comments
            $simulationMatches = $simulationRules | Where-Object { 
                $_.PSObject.Properties.Name -contains "SetHeader" -and $_.SetHeader 
            }
            
            if ($simulationMatches.Count -gt 0) {
                $Result.Comments = "No enabled DLP rule found with SetHeader configured. Found $($simulationMatches.Count) rule(s) in simulation/test mode"
            } else {
                $Result.Comments = "No DLP rule found with SetHeader configured"
            }
        }
        default {
            $Result.Comments = "Unknown DLP rule property: $propertyName"
        }
    }
    return $Result
}

# Helper: Recursively search for a property (e.g., Sensitivetypes > Name) in AdvancedRule JSON
function Get-DlpDeepProperty {
    param([object]$Node, [string]$Target)
    $results = @()
    if ($null -eq $Node) { return $results }
    
    # Handle compound property paths like "Sensitivetypes > Name"
    if ($Target -like "*>*") {
        $parts = $Target -split '\s*>\s*'
        $containerProp = $parts[0]
        $targetProp = $parts[1]
        
        # First find the container (e.g., "Sensitivetypes")
        $containers = Get-DlpDeepProperty -Node $Node -Target $containerProp
        
        # Then extract the target property from each container item
        foreach ($container in $containers) {
            if ($container) {
                # Try case-insensitive property matching
                $matchingProp = $container.PSObject.Properties | Where-Object { $_.Name -ieq $targetProp }
                if ($matchingProp) {
                    $value = $matchingProp.Value
                    if ($value) { $results += $value }
                }
            }
        }
        return $results
    }
    
    # Original logic for simple properties with case-insensitive matching
    if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
        foreach ($item in $Node) {
            $results += Get-DlpDeepProperty -Node $item -Target $Target
        }
    } elseif ($Node -is [hashtable] -or $Node -is [PSCustomObject]) {
        foreach ($key in $Node.PSObject.Properties.Name) {
            $propValue = $Node.PSObject.Properties[$key].Value
            # Use case-insensitive comparison for property names
            if ($key -ieq $Target -and $propValue) {
                if ($propValue -is [System.Collections.IEnumerable] -and -not ($propValue -is [string])) {
                    foreach ($item in $propValue) {
                        $results += $item
                    }
                } else {
                    $results += $propValue
                }
            } elseif ($propValue -ne $null) {
                $results += Get-DlpDeepProperty -Node $propValue -Target $Target
            }
        }
    }
    return $results
}
function Test-ControlBook {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ControlConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$PropertyConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$OptimizedReportPath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$ConfigurationName = "Control Book Assessment"
    )
    
    Write-Host "Starting $ConfigurationName Control Evaluation..."
    
    # Load configuration files
    $controls = Import-Csv -Path $ControlConfigPath
    $properties = Import-Csv -Path $PropertyConfigPath
    
    # Load optimized report
    Write-Host "Loading OptimizedReport.json..."
    $optimizedReport = Get-Content -Path $OptimizedReportPath -Raw | ConvertFrom-Json
    
    # Extract published labels for reference
    $publishedLabels = Get-PublishedLabels -OptimizedReport $optimizedReport
    $allLabels = $optimizedReport.GetLabel
    Write-Host "Found $($publishedLabels.Count) published labels out of $($allLabels.Count) total labels"
    
    # Initialize results array
    $results = @()
    
    # Process each control
    foreach ($control in $controls) {
        $controlProperties = $properties | Where-Object { $_.ControlID -eq $control.ControlID }
        $controlMaturity = $null
        if ($control.PSObject.Properties["MaturityLevel"]) {
            $controlMaturity = $control.MaturityLevel
        } elseif ($control.PSObject.Properties["Maturity Level"]) {
            $controlMaturity = $control."Maturity Level"
        }
        foreach ($property in $controlProperties) {
            Write-Host "Evaluating $($control.ControlID) - $($property.Properties)"
            $propertyMaturity = $null
            if ($property.PSObject.Properties["MaturityLevel"]) { $propertyMaturity = $property.MaturityLevel }
            $result = Test-ControlProperty -Control $control -Property $property -OptimizedReport $optimizedReport -PublishedLabels $publishedLabels -AllLabels $allLabels -ConfigurationName $ConfigurationName
            # Always set MaturityLevel in result, prefer property, then control
            if ($propertyMaturity) {
                $result.MaturityLevel = $propertyMaturity
            } elseif ($controlMaturity) {
                $result.MaturityLevel = $controlMaturity
            } else {
                $result.MaturityLevel = ''
            }
            $results += $result
        }
    }
    
    # Export results to CSV
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "$ConfigurationName control evaluation complete. Results saved to: $OutputPath"
    
    # Return summary statistics
    $passCount = ($results | Where-Object { $_.Pass -eq $true }).Count
    $totalCount = $results.Count
    $passPercentage = if ($totalCount -gt 0) { [math]::Round(($passCount / $totalCount) * 100, 1) } else { 0 }
    
    Write-Host ""
    Write-Host "=== $ConfigurationName Summary ===" -ForegroundColor Cyan
    Write-Host "Total Controls Evaluated: $totalCount" -ForegroundColor White
    Write-Host "Controls Passing: $passCount" -ForegroundColor Green
    Write-Host "Controls Failing: $($totalCount - $passCount)" -ForegroundColor Red
    Write-Host "Compliance Rate: $passPercentage%" -ForegroundColor $(if ($passPercentage -ge 80) { "Green" } elseif ($passPercentage -ge 60) { "Yellow" } else { "Red" })
    
    return @{
        TotalControls = $totalCount
        PassingControls = $passCount
        FailingControls = ($totalCount - $passCount)
        ComplianceRate = $passPercentage
        Results = $results
    }
}

function Get-PublishedLabels {
    param($OptimizedReport)
    
    $publishedLabels = @()
    
    # Get all labels that are marked as Published = true
    $allLabels = $optimizedReport.GetLabel | Where-Object { $_.Published -eq $true }
    
    # Also cross-reference with LabelPolicy to ensure they're actually scoped
    $labelPolicies = $optimizedReport.GetLabelPolicy
    $scopedLabelIds = @()
    
    foreach ($policy in $labelPolicies) {
        if ($policy.ScopedLabels) {
            $scopedLabelIds += $policy.ScopedLabels
        }
    }
    
    # Filter to only include labels that are both Published=true and in policy scope
    foreach ($label in $allLabels) {
        if ($label.ImmutableId -in $scopedLabelIds) {
            $publishedLabels += $label
        }
    }
    
    return $publishedLabels
}

function Test-ControlProperty {
    param($Control, $Property, $OptimizedReport, $PublishedLabels, $AllLabels, $ConfigurationName)
    
    $maturityLevel = $null
    if ($Control.PSObject.Properties["MaturityLevel"]) {
        $maturityLevel = $Control.MaturityLevel
    } elseif ($Control.PSObject.Properties["Maturity Level"]) {
        $maturityLevel = $Control."Maturity Level"
    } elseif ($Property.PSObject.Properties["MaturityLevel"]) {
        $maturityLevel = $Property.MaturityLevel
    } elseif ($Property.PSObject.Properties["Maturity Level"]) {
        $maturityLevel = $Property."Maturity Level"
    }
    $result = [PSCustomObject]@{
        Capability = $Control.Capability
        ControlID = $Control.ControlID
        Control = $Control.Control
        Properties = $Property.Properties
        DefaultValue = $Property.DefaultValue
        MustConfigure = $Property.MustConfigure
        MaturityLevel = $maturityLevel
        Pass = $false
        Comments = ""
        ConfigurationName = $ConfigurationName
    }
    
    try {
        # Parse the property path (e.g., "GetLabel > Published")
        $propertyParts = $Property.Properties -split ' > '
        $dataSource = $propertyParts[0].Trim()
        
        # Get expected value
        $expectedValue = if ($Property.MustConfigure -eq $true -and $Property.DefaultValue) {
            $Property.DefaultValue
        } else {
            $Property.DefaultValue
        }
        
        switch ($dataSource) {
            "GetLabel" {
                $result = Test-GetLabelProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -PublishedLabels $PublishedLabels -AllLabels $AllLabels
            }
            "GetLabelPolicy" {
                $result = Test-GetLabelPolicyProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -OptimizedReport $OptimizedReport -AllLabels $AllLabels
            }
            "GetAutoSensitivityLabelPolicy" {
                $result = Test-GetAutoSensitivityLabelPolicyProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -OptimizedReport $OptimizedReport -AllLabels $AllLabels
            }
            "GetDlpComplianceRule" {
                $result = Test-GetDlpComplianceRuleProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -OptimizedReport $OptimizedReport
            }
            "GetDlpCompliancePolicy" {
                $result = Test-GetDlpCompliancePolicyProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -OptimizedReport $OptimizedReport
            }
            default {
                $result.Comments = "Unknown data source: $dataSource"
            }
        }
    }
    catch {
        $result.Comments = "Error evaluating property: $($_.Exception.Message)"
    }
    
    return $result
}

function Test-GetLabelProperty {
    param($Result, $PropertyParts, $ExpectedValue, $PublishedLabels, $AllLabels)
    
    if ($PropertyParts.Count -lt 2) {
        $Result.Comments = "Invalid property path"
        return $Result
    }
    
    $propertyName = $PropertyParts[1].Trim()
    
    # Handle >> operator for deep property access
    $deepProperty = $null
    if ($propertyName -like "*>>*") {
        $deepParts = $propertyName -split '\s*>>\s*'
        $propertyName = $deepParts[0].Trim()
        $deepProperty = $deepParts[1].Trim()
    }
    
    switch ($propertyName) {
        "Published" {
            # Check if any labels are published
            if ($PublishedLabels.Count -gt 0) {
                $Result.Pass = $true
                $Result.Comments = "Found $($PublishedLabels.Count) published labels"
            } else {
                $Result.Comments = "No published labels found"
            }
        }
        "DisplayName" {
            # Check for comma-separated required labels
            $requiredLabels = $ExpectedValue -split ',' | ForEach-Object { $_.Trim() }
            $foundPublishedLabels = @()
            $foundUnpublishedLabels = @()
            $missingLabels = @()
            
            foreach ($requiredLabel in $requiredLabels) {
                $publishedMatch = $PublishedLabels | Where-Object { $_.DisplayName -ieq $requiredLabel }
                $allMatch = $AllLabels | Where-Object { $_.DisplayName -ieq $requiredLabel }
                
                if ($publishedMatch) {
                    $foundPublishedLabels += $requiredLabel
                } elseif ($allMatch) {
                    $foundUnpublishedLabels += $requiredLabel
                } else {
                    $missingLabels += $requiredLabel
                }
            }
            
            if ($foundPublishedLabels.Count -eq $requiredLabels.Count) {
                $Result.Pass = $true
                $Result.Comments = "All required labels found and published: $($foundPublishedLabels -join ', ')"
            } else {
                $comments = @()
                if ($foundPublishedLabels.Count -gt 0) {
                    $comments += "Published: $($foundPublishedLabels -join ', ')"
                }
                if ($foundUnpublishedLabels.Count -gt 0) {
                    $comments += "Exists but not published: $($foundUnpublishedLabels -join ', ')"
                }
                if ($missingLabels.Count -gt 0) {
                    $comments += "Not found: $($missingLabels -join ', ')"
                }
                $Result.Comments = $comments -join '. '
            }
        }
        "LabelActions" {
            if ($PropertyParts.Count -lt 3) {
                $Result.Comments = "Invalid LabelActions property path"
                return $Result
            }
            
            $actionProperty = $PropertyParts[2].Trim()
            if ($actionProperty -eq "Type") {
                $labelsWithAction = @()
                foreach ($label in $PublishedLabels) {
                    $labelActions = Parse-LabelActions -Label $label
                    $hasAction = $labelActions | Where-Object { $_.Type -eq $ExpectedValue }
                    if ($hasAction) {
                        $labelsWithAction += $label.DisplayName
                    }
                }
                
                if ($labelsWithAction.Count -gt 0) {
                    $Result.Pass = $true
                    $Result.Comments = "Labels with ${ExpectedValue} action: $($labelsWithAction -join ', ')"
                } else {
                    $Result.Comments = "No published labels found with ${ExpectedValue} action"
                }
            }
        }
        "ContentType" {
            $expectedTypes = $ExpectedValue -split ',' | ForEach-Object { $_.Trim() }
            $labelsWithContentType = @()
            
            foreach ($label in $PublishedLabels) {
                $contentTypes = Parse-ContentType -Label $label
                $hasAllTypes = $true
                foreach ($expectedType in $expectedTypes) {
                    if ($expectedType -notin $contentTypes) {
                        $hasAllTypes = $false
                        break
                    }
                }
                if ($hasAllTypes) {
                    $labelsWithContentType += $label.DisplayName
                }
            }
            
            if ($labelsWithContentType.Count -gt 0) {
                $Result.Pass = $true
                $Result.Comments = "Labels with required content types: $($labelsWithContentType -join ', ')"
            } else {
                $Result.Comments = "No published labels found with required content types: $($expectedTypes -join ', ')"
            }
        }
        "Conditions" {
            # Handle both old format (GetLabel > Conditions > Key) and new format (GetLabel > Conditions >> Key)
            $conditionProperty = $null
            if ($deepProperty) {
                # New format with >> operator
                $conditionProperty = $deepProperty
            } elseif ($PropertyParts.Count -ge 3) {
                # Old format with single > operator
                $conditionProperty = $PropertyParts[2].Trim()
            } else {
                $Result.Comments = "Invalid Conditions property path"
                return $Result
            }
            
            # For Sensitivity Auto-labelling controls, only consider labels from the taxonomy defined in DisplayName control
            $labelsToCheck = $PublishedLabels
            if ($Result.Capability -eq "Sensitivity Auto-labelling") {
                # Get the required labels from the DisplayName control in the same configuration
                $taxonomyControl = Get-TaxonomyLabels -AllLabels $AllLabels -Result $Result
                $taxonomyLabelNames = $taxonomyControl.RequiredLabels
                $taxonomyLabels = $PublishedLabels | Where-Object { $_.DisplayName -in $taxonomyLabelNames }
                $outOfScopeLabels = @()
                
                # Check all published labels for the condition, but separate taxonomy from others
                foreach ($label in $PublishedLabels) {
                    $conditions = Parse-LabelConditions -Label $label
                    $hasCondition = $false
                    
                    if ($conditionProperty -eq "Key" -and $conditions.ContainsKey($ExpectedValue)) {
                        $hasCondition = $true
                    } elseif ($conditionProperty -eq "Value" -and $conditions.ContainsValue($ExpectedValue)) {
                        $hasCondition = $true
                    }
                    
                    if ($hasCondition -and $label.DisplayName -notin $taxonomyLabelNames) {
                        $outOfScopeLabels += $label.DisplayName
                    }
                }
                
                $labelsToCheck = $taxonomyLabels
                $labelsWithCondition = @()
                foreach ($label in $taxonomyLabels) {
                    $conditions = Parse-LabelConditions -Label $label
                    if ($conditionProperty -eq "Key" -and $conditions.ContainsKey($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    } elseif ($conditionProperty -eq "Value" -and $conditions.ContainsValue($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    }
                }
                
                if ($labelsWithCondition.Count -gt 0) {
                    $Result.Pass = $true
                    $comments = "Taxonomy labels with condition ${conditionProperty} = ${ExpectedValue}: $($labelsWithCondition -join ', ')"
                    if ($outOfScopeLabels.Count -gt 0) {
                        $comments += ". Out-of-scope labels also have this condition: $($outOfScopeLabels -join ', ')"
                    }
                    $Result.Comments = $comments
                } else {
                    $comments = "No taxonomy labels found with condition ${conditionProperty} = ${ExpectedValue}"
                    if ($outOfScopeLabels.Count -gt 0) {
                        $comments += ". Out-of-scope labels have this condition: $($outOfScopeLabels -join ', ')"
                    }
                    $Result.Comments = $comments
                }
            } else {
                # For non-auto-labeling controls, check all published labels
                $labelsWithCondition = @()
                foreach ($label in $PublishedLabels) {
                    $conditions = Parse-LabelConditions -Label $label
                    if ($conditionProperty -eq "Key" -and $conditions.ContainsKey($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    } elseif ($conditionProperty -eq "Value" -and $conditions.ContainsValue($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    }
                }
                
                if ($labelsWithCondition.Count -gt 0) {
                    $Result.Pass = $true
                    $Result.Comments = "Labels with condition ${conditionProperty} = ${ExpectedValue}: $($labelsWithCondition -join ', ')"
                } else {
                    $Result.Comments = "No published labels found with condition ${conditionProperty} = ${ExpectedValue}"
                }
            }
        }
        default {
            $Result.Comments = "Unknown GetLabel property: $propertyName"
        }
    }
    
    return $Result
}

function Test-GetLabelPolicyProperty {
    param($Result, $PropertyParts, $ExpectedValue, $OptimizedReport, $AllLabels)
    
    if ($PropertyParts.Count -lt 3) {
        $Result.Comments = "Invalid GetLabelPolicy property path"
        return $Result
    }
    
    $settingName = $PropertyParts[2].Trim()
    $labelPolicies = $OptimizedReport.GetLabelPolicy
    
    # Get the required labels from the taxonomy (dynamically determined from configuration)
    $taxonomyControl = Get-TaxonomyLabels -AllLabels $AllLabels -Result $Result
    $requiredLabelNames = $taxonomyControl.RequiredLabels
    $requiredLabelIds = @()
    
    # Find the ImmutableIds for our required labels from GetLabel
    foreach ($requiredLabel in $requiredLabelNames) {
        $matchingLabel = $AllLabels | Where-Object { $_.DisplayName -eq $requiredLabel }
        if ($matchingLabel) {
            $requiredLabelIds += $matchingLabel.ImmutableId
        }
    }
    
    if ($requiredLabelIds.Count -eq 0) {
        $Result.Comments = "None of the required taxonomy labels ($($requiredLabelNames -join ', ')) found in tenant"
        return $Result
    }
    
    # Find policies that contain any of our required labels
    $relevantPolicies = @()
    foreach ($policy in $labelPolicies) {
        if ($policy.ScopedLabels) {
            $hasRequiredLabel = $false
            foreach ($labelId in $requiredLabelIds) {
                # Check if the ImmutableId is in the policy's ScopedLabels array
                if ($labelId -in $policy.ScopedLabels) {
                    $hasRequiredLabel = $true
                    break
                }
            }
            if ($hasRequiredLabel) {
                $relevantPolicies += $policy
            }
        }
    }
    
    if ($relevantPolicies.Count -eq 0) {
        $Result.Comments = "No label policies found containing the required taxonomy labels"
        return $Result
    }
    
    # Check if ANY of the relevant policies has the required setting
    $policiesWithSetting = @()
    foreach ($policy in $relevantPolicies) {
        $settings = Parse-PolicySettings -Policy $policy
        if ($settings.ContainsKey($settingName)) {
            $actualValue = $settings[$settingName]
            
            # Handle different validation types based on expected value
            $isMatch = $false
            if ($ExpectedValue -eq "Not Null") {
                # For settings that should have any non-null value
                if ($actualValue -and $actualValue -ne "None" -and $actualValue -ne "") {
                    $isMatch = $true
                }
            } else {
                # For exact value matching
                if ($actualValue -eq $ExpectedValue) {
                    $isMatch = $true
                }
            }
            
            $policiesWithSetting += @{
                Policy = $policy
                Value = $actualValue
                Match = $isMatch
            }
        }
    }
    
    $matchingPolicies = $policiesWithSetting | Where-Object { $_.Match -eq $true }
    if ($matchingPolicies.Count -gt 0) {
        $Result.Pass = $true
        if ($ExpectedValue -eq "Not Null") {
            $Result.Comments = "Found $($matchingPolicies.Count) policies containing taxonomy labels with ${settingName} configured (not null)"
        } else {
            $Result.Comments = "Found $($matchingPolicies.Count) policies containing taxonomy labels with ${settingName} = ${ExpectedValue}"
        }
    } else {
        $allValues = ($policiesWithSetting | ForEach-Object { $_.Value }) -join ', '
        if ($ExpectedValue -eq "Not Null") {
            $Result.Comments = "No policies containing taxonomy labels found with ${settingName} configured. Found values: $allValues"
        } else {
            $Result.Comments = "No policies containing taxonomy labels found with ${settingName} = ${ExpectedValue}. Found values: $allValues"
        }
    }
    
    return $Result
}

function Test-GetAutoSensitivityLabelPolicyProperty {
    param($Result, $PropertyParts, $ExpectedValue, $OptimizedReport, $AllLabels)
    
    if ($PropertyParts.Count -lt 2) {
        $Result.Comments = "Invalid GetAutoSensitivityLabelPolicy property path"
        return $Result
    }
    
    $propertyName = $PropertyParts[1].Trim()
    $autoPolicies = $OptimizedReport.GetAutoSensitivityLabelPolicy
    
    # For Sensitivity Auto-labelling controls, filter to only taxonomy-related policies
    $taxonomyControl = Get-TaxonomyLabels -AllLabels $AllLabels -Result $Result
    $taxonomyLabelNames = $taxonomyControl.RequiredLabels
    $taxonomyRelatedPolicies = @()
    $outOfScopePolicies = @()
    
    foreach ($policy in $autoPolicies) {
        $isTaxonomyRelated = $false
        if ($policy.LabelDisplayName) {
            foreach ($taxonomyLabel in $taxonomyLabelNames) {
                if ($policy.LabelDisplayName -like "*$taxonomyLabel*") {
                    $isTaxonomyRelated = $true
                    break
                }
            }
        }
        
        if ($isTaxonomyRelated) {
            $taxonomyRelatedPolicies += $policy
        } else {
            $outOfScopePolicies += $policy
        }
    }
    
    switch ($propertyName) {
        "Mode" {
            $enabledPolicies = $taxonomyRelatedPolicies | Where-Object { $_.Mode -eq $ExpectedValue }
            $outOfScopeEnabledPolicies = $outOfScopePolicies | Where-Object { $_.Mode -eq $ExpectedValue }
            
            if ($enabledPolicies.Count -gt 0) {
                $Result.Pass = $true
                $comments = "Found $($enabledPolicies.Count) taxonomy auto-labeling policies with Mode = ${ExpectedValue}"
                if ($outOfScopeEnabledPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also enabled: $($outOfScopeEnabledPolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No taxonomy auto-labeling policies found with Mode = ${ExpectedValue}"
                if ($outOfScopeEnabledPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies with this mode: $($outOfScopeEnabledPolicies.Count)"
                }
                $Result.Comments = $comments
            }
        }
        "Type" {
            $typePolicies = $taxonomyRelatedPolicies | Where-Object { $_.Type -eq $ExpectedValue }
            $outOfScopeTypePolicies = $outOfScopePolicies | Where-Object { $_.Type -eq $ExpectedValue }
            
            if ($typePolicies.Count -gt 0) {
                $Result.Pass = $true
                $comments = "Found $($typePolicies.Count) taxonomy policies with Type = ${ExpectedValue}"
                if ($outOfScopeTypePolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also have this type: $($outOfScopeTypePolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No taxonomy policies found with Type = ${ExpectedValue}"
                if ($outOfScopeTypePolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies with this type: $($outOfScopeTypePolicies.Count)"
                }
                $Result.Comments = $comments
            }
        }
        "Workload" {
            $expectedWorkloads = $ExpectedValue -split ',' | ForEach-Object { $_.Trim() }
            $matchingPolicies = @()
            $outOfScopeMatchingPolicies = @()
            
            foreach ($policy in $taxonomyRelatedPolicies) {
                $hasAllWorkloads = $true
                foreach ($expectedWorkload in $expectedWorkloads) {
                    $workloadFound = $false
                    # Check various workload properties
                    if ($expectedWorkload -eq "SharePoint" -and $policy.SharePointLocation) {
                        $workloadFound = $true
                    } elseif ($expectedWorkload -eq "OneDriveForBusiness" -and $policy.OneDriveLocation) {
                        $workloadFound = $true
                    } elseif ($expectedWorkload -eq "Exchange" -and $policy.ExchangeLocation) {
                        $workloadFound = $true
                    }
                    
                    if (-not $workloadFound) {
                        $hasAllWorkloads = $false
                        break
                    }
                }
                
                if ($hasAllWorkloads) {
                    $matchingPolicies += $policy
                }
            }
            
            # Check out-of-scope policies too for reporting
            foreach ($policy in $outOfScopePolicies) {
                $hasAllWorkloads = $true
                foreach ($expectedWorkload in $expectedWorkloads) {
                    $workloadFound = $false
                    if ($expectedWorkload -eq "SharePoint" -and $policy.SharePointLocation) {
                        $workloadFound = $true
                    } elseif ($expectedWorkload -eq "OneDriveForBusiness" -and $policy.OneDriveLocation) {
                        $workloadFound = $true
                    } elseif ($expectedWorkload -eq "Exchange" -and $policy.ExchangeLocation) {
                        $workloadFound = $true
                    }
                    
                    if (-not $workloadFound) {
                        $hasAllWorkloads = $false
                        break
                    }
                }
                
                if ($hasAllWorkloads) {
                    $outOfScopeMatchingPolicies += $policy
                }
            }
            
            if ($matchingPolicies.Count -gt 0) {
                $Result.Pass = $true
                $comments = "Found $($matchingPolicies.Count) taxonomy policies with required workloads: $($expectedWorkloads -join ', ')"
                if ($outOfScopeMatchingPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also have these workloads: $($outOfScopeMatchingPolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No taxonomy policies found with required workloads: $($expectedWorkloads -join ', ')"
                if ($outOfScopeMatchingPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies have these workloads: $($outOfScopeMatchingPolicies.Count)"
                }
                $Result.Comments = $comments
            }
        }
        default {
            $Result.Comments = "Unknown GetAutoSensitivityLabelPolicy property: $propertyName"
        }
    }
    
    return $Result
}

function Get-TaxonomyLabels {
    param($AllLabels, $Result)
    
    # This function dynamically determines the required taxonomy labels from the configuration
    # It looks for a DisplayName control in the same capability that defines the taxonomy
    
    # For now, we'll use the PSPF taxonomy as default, but this could be made more dynamic
    # by reading from the configuration files to find taxonomy-defining controls
    $defaultTaxonomy = @("UNOFFICIAL", "OFFICIAL", "OFFICIAL SENSITIVE")
    
    return @{
        RequiredLabels = $defaultTaxonomy
        Source = "Default PSPF Taxonomy"
    }
}

# Helper functions for parsing complex data structures
function Parse-LabelActions {
    param($Label)
    
    $actions = @()
    if ($Label.LabelActions) {
        foreach ($actionJson in $Label.LabelActions) {
            try {
                $action = $actionJson | ConvertFrom-Json
                $actions += $action
            } catch {
                # Handle malformed JSON
            }
        }
    }
    return $actions
}

function Parse-ContentType {
    param($Label)
    
    $contentTypes = @()
    if ($Label.Settings) {
        foreach ($setting in $Label.Settings) {
            if ($setting -match '\[contenttype,\s*(.+?)\]') {
                $types = $matches[1] -split ',' | ForEach-Object { $_.Trim() }
                $contentTypes += $types
            }
        }
    }
    return $contentTypes
}

function Parse-LabelConditions {
    param($Label)
    
    $conditions = @{}
    if ($Label.Conditions) {
        try {
            $conditionsObj = $Label.Conditions | ConvertFrom-Json
            # Parse the complex conditions structure recursively
            Parse-ConditionsRecursively -Node $conditionsObj -Conditions ([ref]$conditions)
        } catch {
            # Handle malformed JSON
        }
    }
    return $conditions
}

function Parse-ConditionsRecursively {
    param($Node, [ref]$Conditions)
    
    if ($null -eq $Node) { return }
    
    # Handle arrays
    if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
        foreach ($item in $Node) {
            Parse-ConditionsRecursively -Node $item -Conditions $Conditions
        }
        return
    }
    
    # Handle objects
    if ($Node -is [hashtable] -or $Node -is [PSCustomObject]) {
        # Check for Settings array (this is where Key/Value pairs are stored)
        if ($Node.Settings) {
            foreach ($setting in $Node.Settings) {
                if ($setting.Key -and $setting.Value) {
                    $Conditions.Value[$setting.Key] = $setting.Value
                }
            }
        }
        
        # Recursively parse all properties (And, Or, etc.)
        foreach ($prop in $Node.PSObject.Properties) {
            if ($prop.Value) {
                Parse-ConditionsRecursively -Node $prop.Value -Conditions $Conditions
            }
        }
    }
}

function Parse-PolicySettings {
    param($Policy)
    
    $settings = @{}
    if ($Policy.Settings) {
        foreach ($setting in $Policy.Settings) {
            if ($setting -match '\[(.+?),\s*(.+?)\]') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $settings[$key] = $value
            }
        }
    }
    return $settings
}
