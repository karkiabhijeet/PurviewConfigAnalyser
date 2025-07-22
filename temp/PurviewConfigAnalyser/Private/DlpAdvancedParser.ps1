# Enhanced DLP parsing functions to handle nested SubConditions and compound property paths
# These functions extend the basic DLP parsing to support:
# - Nested SubConditions (for DLP_4.6, DLP_4.7, DLP_4.8)
# - Compound property paths like "Sensitivetypes > Name", "Sensitivetypes > MinCount", etc.

function Test-DlpAdvancedRuleProperty {
    param(
        [PSCustomObject]$AdvRule,
        [string]$DeepProp,
        [string]$DeepVal,
        [string]$ControlId
    )
    
    $result = @{
        Found = $false
        AnyTypesFound = $false
        HighestMinCount = 0
        ClassifierTypeFound = $false
        AllNames = @()
    }
    
    # Handle both simple "Sensitivetypes" and compound "Sensitivetypes > Property" paths
    if ($DeepProp -like "Sensitivetypes*" -and $AdvRule.Condition.SubConditions) {
        # Parse the deeper property if it exists (e.g., "Sensitivetypes > Name")
        $sensitivetypesParts = $DeepProp -split ' > '
        $targetProperty = if ($sensitivetypesParts.Count -ge 2) { $sensitivetypesParts[1].Trim() } else { $null }
        
        # Process top-level SubConditions
        foreach ($subCond in $AdvRule.Condition.SubConditions) {
            if ($subCond.ConditionName -eq "ContentContainsSensitiveInformation" -and $subCond.Value) {
                foreach ($valueItem in $subCond.Value) {
                    if ($valueItem.Groups) {
                        foreach ($group in $valueItem.Groups) {
                            if ($group.Sensitivetypes -and $group.Sensitivetypes.Count -gt 0) {
                                $sensTypes = $group.Sensitivetypes
                                $result.AnyTypesFound = $true
                                
                                # Collect all Names
                                foreach ($sensType in $sensTypes) {
                                    if ($sensType.Name) {
                                        $result.AllNames += $sensType.Name
                                    }
                                    
                                    # Track highest MinCount (handle both Mincount and MinCount)
                                    $minCountValue = if ($sensType.Mincount) { $sensType.Mincount } elseif ($sensType.MinCount) { $sensType.MinCount } else { 0 }
                                    if ($minCountValue -gt $result.HighestMinCount) {
                                        $result.HighestMinCount = $minCountValue
                                    }
                                    
                                    # Check for ClassifierType (handle both variants)
                                    $classifierTypeValue = if ($sensType.Classifiertype) { $sensType.Classifiertype } elseif ($sensType.ClassifierType) { $sensType.ClassifierType } else { $null }
                                    if ($classifierTypeValue -eq "MLModel") {
                                        $result.ClassifierTypeFound = $true
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            # For specific failing controls (DLP_4.6, DLP_4.7, DLP_4.8), also check one level deeper
            if (($ControlId -eq "DLP_4.6" -or $ControlId -eq "DLP_4.7" -or $ControlId -eq "DLP_4.8") -and 
                $subCond.SubConditions -and $subCond.SubConditions.Count -gt 0) {
                foreach ($nestedSubCond in $subCond.SubConditions) {
                    if ($nestedSubCond.ConditionName -eq "ContentContainsSensitiveInformation" -and $nestedSubCond.Value) {
                        foreach ($valueItem in $nestedSubCond.Value) {
                            if ($valueItem.Groups) {
                                foreach ($group in $valueItem.Groups) {
                                    if ($group.Sensitivetypes -and $group.Sensitivetypes.Count -gt 0) {
                                        $sensTypes = $group.Sensitivetypes
                                        $result.AnyTypesFound = $true
                                        
                                        # Collect all Names
                                        foreach ($sensType in $sensTypes) {
                                            if ($sensType.Name) {
                                                $result.AllNames += $sensType.Name
                                            }
                                            
                                            # Track highest MinCount (handle both variants)
                                            $minCountValue = if ($sensType.Mincount) { $sensType.Mincount } elseif ($sensType.MinCount) { $sensType.MinCount } else { 0 }
                                            if ($minCountValue -gt $result.HighestMinCount) {
                                                $result.HighestMinCount = $minCountValue
                                            }
                                            
                                            # Check for ClassifierType (handle both variants)
                                            $classifierTypeValue = if ($sensType.Classifiertype) { $sensType.Classifiertype } elseif ($sensType.ClassifierType) { $sensType.ClassifierType } else { $null }
                                            if ($classifierTypeValue -eq "MLModel") {
                                                $result.ClassifierTypeFound = $true
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        # Check if the specific property requirements are met
        if ($targetProperty -eq "Name" -and $DeepVal -eq 'Not Null') {
            $result.Found = $result.AnyTypesFound
        }
        elseif ($targetProperty -eq "MinCount") {
            $result.Found = ($result.HighestMinCount -ge [int]$DeepVal)
        }
        elseif ($targetProperty -eq "Name" -and $DeepVal -match ",") {
            # For comma-separated names check
            $requiredNames = $DeepVal -split ',' | ForEach-Object { $_.Trim() }
            $allFound = $true
            foreach ($reqName in $requiredNames) {
                if ($result.AllNames -notcontains $reqName) {
                    $allFound = $false
                    break
                }
            }
            $result.Found = $allFound
        }
        elseif ($targetProperty -eq "ClassifierType") {
            $result.Found = $result.ClassifierTypeFound
        }
        elseif ($targetProperty -eq "Name") {
            # For specific Name match
            $result.Found = ($result.AllNames -contains $DeepVal)
        }
        elseif ($targetProperty -eq $null -and $DeepVal -eq 'Not Null') {
            # For legacy format compatibility - old format without target property
            $result.Found = $result.AnyTypesFound
        }
    }
    
    return $result
}

function Get-DlpPropertyComment {
    param(
        [string]$DeepProp,
        [string]$DeepVal,
        [string]$CommentType = "Found" # "Found", "NotFoundWithSimulation", "NotFound"
    )
    
    # Parse deeper properties if they exist (e.g., "Sensitivetypes > MinCount")
    $targetProperty = $null
    if ($DeepProp -like "*>*") {
        $deepPropParts = $DeepProp -split ' > '
        if ($deepPropParts.Count -ge 2) {
            $targetProperty = $deepPropParts[1].Trim()
        }
    }
    
    switch ($CommentType) {
        "Found" {
            if ($targetProperty -eq "MinCount") {
                return "Found DLP rule with $DeepProp >= $DeepVal"
            } elseif ($targetProperty -eq "Name" -and $DeepVal -match ",") {
                return "Found DLP rule with all required sensitive type names: $DeepVal"
            } elseif ($targetProperty -eq "ClassifierType") {
                return "Found DLP rule with $DeepProp = $DeepVal"
            } elseif ($DeepVal -eq 'Not Null') {
                return "Found DLP rule with $DeepProp configured"
            } else {
                return "Found DLP rule with $DeepProp = $DeepVal"
            }
        }
        "NotFoundWithSimulation" {
            if ($targetProperty -eq "MinCount") {
                return "No enabled DLP rule found with $DeepProp >= $DeepVal. Found matching rule(s) in simulation/test mode"
            } elseif ($targetProperty -eq "Name" -and $DeepVal -match ",") {
                return "No enabled DLP rule found with all required sensitive type names: $DeepVal. Found matching rule(s) in simulation/test mode"
            } elseif ($targetProperty -eq "ClassifierType") {
                return "No enabled DLP rule found with $DeepProp = $DeepVal. Found matching rule(s) in simulation/test mode"
            } else {
                return "No enabled DLP rule found with $DeepProp in AdvancedRule. Found matching rule(s) in simulation/test mode"
            }
        }
        "NotFound" {
            if ($targetProperty -eq "MinCount") {
                return "No DLP rule found with $DeepProp >= $DeepVal"
            } elseif ($targetProperty -eq "Name" -and $DeepVal -match ",") {
                return "No DLP rule found with all required sensitive type names: $DeepVal"
            } elseif ($targetProperty -eq "ClassifierType") {
                return "No DLP rule found with $DeepProp = $DeepVal"
            } else {
                return "No DLP rule found with $DeepProp in AdvancedRule"
            }
        }
    }
}
