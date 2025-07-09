function Test-PSPFControls {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$PropertyConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$OptimizedReportPath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    Write-Host "Starting PSPF Control Evaluation..."
    
    # Load configuration files
    $controls = Import-Csv -Path $ConfigPath
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
        
        foreach ($property in $controlProperties) {
            Write-Host "Evaluating $($control.ControlID) - $($property.Properties)"
            
            $result = Test-ControlProperty -Control $control -Property $property -OptimizedReport $optimizedReport -PublishedLabels $publishedLabels -AllLabels $allLabels
            $results += $result
        }
    }
    
    # Export results to CSV
    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "PSPF Control evaluation complete. Results saved to: $OutputPath"
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
        if ($policy.Labels) {
            $scopedLabelIds += $policy.Labels
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
    param($Control, $Property, $OptimizedReport, $PublishedLabels, $AllLabels)
    
    $result = [PSCustomObject]@{
        Capability = $Control.Capability
        ControlID = $Control.ControlID
        Control = $Control.Control
        Properties = $Property.Properties
        DefaultValue = $Property.DefaultValue
        MustConfigure = $Property.MustConfigure
        Pass = $false
        Comments = ""
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
                $result = Test-GetAutoSensitivityLabelPolicyProperty -Result $result -PropertyParts $propertyParts -ExpectedValue $expectedValue -OptimizedReport $OptimizedReport
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
            # Special handling for SL_1.3 - check for comma-separated required labels
            if ($Result.ControlID -eq "SL_1.3") {
                $requiredLabels = $ExpectedValue -split ',' | ForEach-Object { $_.Trim() }
                $foundPublishedLabels = @()
                $foundUnpublishedLabels = @()
                $missingLabels = @()
                
                foreach ($requiredLabel in $requiredLabels) {
                    $publishedMatch = $PublishedLabels | Where-Object { $_.DisplayName -eq $requiredLabel }
                    $allMatch = $AllLabels | Where-Object { $_.DisplayName -eq $requiredLabel }
                    
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
            } else {
                # General DisplayName check
                $matchingLabels = $PublishedLabels | Where-Object { $_.DisplayName -match $ExpectedValue }
                if ($matchingLabels) {
                    $Result.Pass = $true
                    $Result.Comments = "Found matching labels: $($matchingLabels.DisplayName -join ', ')"
                } else {
                    $Result.Comments = "No labels found matching: $ExpectedValue"
                }
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
            if ($PropertyParts.Count -lt 3) {
                $Result.Comments = "Invalid Conditions property path"
                return $Result
            }
            
            $conditionProperty = $PropertyParts[2].Trim()
            
            # For Sensitivity Auto-labelling controls, only consider PSPF labels
            $labelsToCheck = $PublishedLabels
            if ($Result.Capability -eq "Sensitivity Auto-labelling") {
                $pspfLabelNames = @("UNOFFICIAL", "OFFICIAL", "OFFICIAL SENSITIVE")
                $pspfLabels = $PublishedLabels | Where-Object { $_.DisplayName -in $pspfLabelNames }
                $outOfScopeLabels = @()
                
                # Check all published labels for the condition, but separate PSPF from others
                $allLabelsWithCondition = @()
                foreach ($label in $PublishedLabels) {
                    $conditions = Parse-LabelConditions -Label $label
                    $hasCondition = $false
                    
                    if ($conditionProperty -eq "Key" -and $conditions.ContainsKey($ExpectedValue)) {
                        $hasCondition = $true
                    } elseif ($conditionProperty -eq "Value" -and $conditions.ContainsValue($ExpectedValue)) {
                        $hasCondition = $true
                    }
                    
                    if ($hasCondition) {
                        $allLabelsWithCondition += $label.DisplayName
                        if ($label.DisplayName -notin $pspfLabelNames) {
                            $outOfScopeLabels += $label.DisplayName
                        }
                    }
                }
                
                $labelsToCheck = $pspfLabels
                $labelsWithCondition = @()
                foreach ($label in $pspfLabels) {
                    $conditions = Parse-LabelConditions -Label $label
                    if ($conditionProperty -eq "Key" -and $conditions.ContainsKey($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    } elseif ($conditionProperty -eq "Value" -and $conditions.ContainsValue($ExpectedValue)) {
                        $labelsWithCondition += $label.DisplayName
                    }
                }
                
                if ($labelsWithCondition.Count -gt 0) {
                    $Result.Pass = $true
                    $comments = "PSPF labels with condition ${conditionProperty} = ${ExpectedValue}: $($labelsWithCondition -join ', ')"
                    if ($outOfScopeLabels.Count -gt 0) {
                        $comments += ". Out-of-scope labels also have this condition: $($outOfScopeLabels -join ', ')"
                    }
                    $Result.Comments = $comments
                } else {
                    $comments = "No PSPF labels found with condition ${conditionProperty} = ${ExpectedValue}"
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
    
    # Get the required labels for PSPF (UNOFFICIAL, OFFICIAL, OFFICIAL SENSITIVE)
    $requiredLabels = @("UNOFFICIAL", "OFFICIAL", "OFFICIAL SENSITIVE")
    $requiredLabelIds = @()
    
    # Find the ImmutableIds for our required labels from GetLabel
    foreach ($requiredLabel in $requiredLabels) {
        $matchingLabel = $AllLabels | Where-Object { $_.DisplayName -eq $requiredLabel }
        if ($matchingLabel) {
            $requiredLabelIds += $matchingLabel.ImmutableId
        }
    }
    
    if ($requiredLabelIds.Count -eq 0) {
        $Result.Comments = "None of the required PSPF labels (UNOFFICIAL, OFFICIAL, OFFICIAL SENSITIVE) found in tenant"
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
        $Result.Comments = "No label policies found containing the required PSPF labels"
        return $Result
    }
    
    # Check if ANY of the relevant policies has the required setting
    $policiesWithSetting = @()
    foreach ($policy in $relevantPolicies) {
        $settings = Parse-PolicySettings -Policy $policy
        if ($settings.ContainsKey($settingName)) {
            $actualValue = $settings[$settingName]
            
            # Handle different validation types based on control ID
            $isMatch = $false
            if ($Result.ControlID -in @("SL_1.10", "SL_1.11", "SL_1.12")) {
                # For custom URL and default label settings, check for "Not Null"
                if ($ExpectedValue -eq "Not Null" -and $actualValue -and $actualValue -ne "None" -and $actualValue -ne "") {
                    $isMatch = $true
                }
            } else {
                # For boolean settings (SL_1.4 to SL_1.9), check exact match
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
        if ($Result.ControlID -in @("SL_1.10", "SL_1.11", "SL_1.12")) {
            $Result.Comments = "Found $($matchingPolicies.Count) policies containing PSPF labels with ${settingName} configured (not null)"
        } else {
            $Result.Comments = "Found $($matchingPolicies.Count) policies containing PSPF labels with ${settingName} = ${ExpectedValue}"
        }
    } else {
        $allValues = ($policiesWithSetting | ForEach-Object { $_.Value }) -join ', '
        if ($Result.ControlID -in @("SL_1.10", "SL_1.11", "SL_1.12")) {
            $Result.Comments = "No policies containing PSPF labels found with ${settingName} configured. Found values: $allValues"
        } else {
            $Result.Comments = "No policies containing PSPF labels found with ${settingName} = ${ExpectedValue}. Found values: $allValues"
        }
    }
    
    return $Result
}

function Test-GetAutoSensitivityLabelPolicyProperty {
    param($Result, $PropertyParts, $ExpectedValue, $OptimizedReport)
    
    if ($PropertyParts.Count -lt 2) {
        $Result.Comments = "Invalid GetAutoSensitivityLabelPolicy property path"
        return $Result
    }
    
    $propertyName = $PropertyParts[1].Trim()
    $autoPolicies = $OptimizedReport.GetAutoSensitivityLabelPolicy
    
    # For Sensitivity Auto-labelling controls, filter to only PSPF-related policies
    $pspfLabelNames = @("UNOFFICIAL", "OFFICIAL", "OFFICIAL SENSITIVE")
    $pspfRelatedPolicies = @()
    $outOfScopePolicies = @()
    
    foreach ($policy in $autoPolicies) {
        $isPspfRelated = $false
        if ($policy.LabelDisplayName) {
            foreach ($pspfLabel in $pspfLabelNames) {
                if ($policy.LabelDisplayName -like "*$pspfLabel*") {
                    $isPspfRelated = $true
                    break
                }
            }
        }
        
        if ($isPspfRelated) {
            $pspfRelatedPolicies += $policy
        } else {
            $outOfScopePolicies += $policy
        }
    }
    
    switch ($propertyName) {
        "Mode" {
            $enabledPolicies = $pspfRelatedPolicies | Where-Object { $_.Mode -eq $ExpectedValue }
            $outOfScopeEnabledPolicies = $outOfScopePolicies | Where-Object { $_.Mode -eq $ExpectedValue }
            
            if ($enabledPolicies.Count -gt 0) {
                $Result.Pass = $true
                $comments = "Found $($enabledPolicies.Count) PSPF auto-labeling policies with Mode = ${ExpectedValue}"
                if ($outOfScopeEnabledPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also enabled: $($outOfScopeEnabledPolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No PSPF auto-labeling policies found with Mode = ${ExpectedValue}"
                if ($outOfScopeEnabledPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies with this mode: $($outOfScopeEnabledPolicies.Count)"
                }
                $Result.Comments = $comments
            }
        }
        "Type" {
            $typePolicies = $pspfRelatedPolicies | Where-Object { $_.Type -eq $ExpectedValue }
            $outOfScopeTypePolicies = $outOfScopePolicies | Where-Object { $_.Type -eq $ExpectedValue }
            
            if ($typePolicies.Count -gt 0) {
                $Result.Pass = $true
                $comments = "Found $($typePolicies.Count) PSPF policies with Type = ${ExpectedValue}"
                if ($outOfScopeTypePolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also have this type: $($outOfScopeTypePolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No PSPF policies found with Type = ${ExpectedValue}"
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
            
            foreach ($policy in $pspfRelatedPolicies) {
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
                $comments = "Found $($matchingPolicies.Count) PSPF policies with required workloads: $($expectedWorkloads -join ', ')"
                if ($outOfScopeMatchingPolicies.Count -gt 0) {
                    $comments += ". Out-of-scope policies also have these workloads: $($outOfScopeMatchingPolicies.Count)"
                }
                $Result.Comments = $comments
            } else {
                $comments = "No PSPF policies found with required workloads: $($expectedWorkloads -join ', ')"
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
            # Parse the complex conditions structure
            if ($conditionsObj.And) {
                foreach ($condition in $conditionsObj.And) {
                    if ($condition.And) {
                        foreach ($subCondition in $condition.And) {
                            if ($subCondition.Settings) {
                                foreach ($setting in $subCondition.Settings) {
                                    $conditions[$setting.Key] = $setting.Value
                                }
                            }
                        }
                    }
                }
            }
        } catch {
            # Handle malformed JSON
        }
    }
    return $conditions
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
