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
        
        foreach ($property in $controlProperties) {
            Write-Host "Evaluating $($control.ControlID) - $($property.Properties)"
            
            $result = Test-ControlProperty -Control $control -Property $property -OptimizedReport $optimizedReport -PublishedLabels $publishedLabels -AllLabels $allLabels -ConfigurationName $ConfigurationName
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
    
    $result = [PSCustomObject]@{
        Capability = $Control.Capability
        ControlID = $Control.ControlID
        Control = $Control.Control
        Properties = $Property.Properties
        DefaultValue = $Property.DefaultValue
        MustConfigure = $Property.MustConfigure
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
            # Check for comma-separated required labels
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
