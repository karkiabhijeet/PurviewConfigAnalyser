function Test-SensitivityLabels {
    param (
        [Parameter(Mandatory)]
        [string]$ConfigPath,
        [Parameter(Mandatory)]
        [string]$ActualPath,
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    # Load config and actual data
    $configModels = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $actual = Get-Content $ActualPath -Raw | ConvertFrom-Json

    $finalReport = @()

    foreach ($config in $configModels) {
        $report = @{
            MaturityModel = $config.MaturityModel
            SensitivityLabels = @()
        }

        foreach ($expected in $config.SensitivityLabels) {
            # Find matching label in actual config
            $actualLabel = $actual.GetLabel | Where-Object {
                $_.DisplayName -eq $expected.DisplayName -and
                $_.ParentLabelDisplayName -eq $expected.ParentLabelDisplayName
            }

            $labelReport = @{
                DisplayName = $expected.DisplayName
                ParentLabelDisplayName = $expected.ParentLabelDisplayName
                Workload = $null
                ContentType = $null
                Autolabel = $null
                LabelActions = @()
                AllActionsMet = $true
            }

            # Workload: check if all expected workloads are present in actual
            if ($null -ne $expected.Workload) {
                $actualValue = $actualLabel.Workload
                $expectedValue = $expected.Workload
                $actualArr = @()
                $expectedArr = @()
                if ($actualValue -is [string]) {
                    $actualArr = $actualValue -split ',' | ForEach-Object { $_.Trim() }
                } elseif ($actualValue -is [array]) {
                    $actualArr = $actualValue | ForEach-Object { $_.Trim() }
                }
                if ($expectedValue -is [string]) {
                    $expectedArr = $expectedValue -split ',' | ForEach-Object { $_.Trim() }
                } elseif ($expectedValue -is [array]) {
                    $expectedArr = $expectedValue | ForEach-Object { $_.Trim() }
                }
                $met = $null
                if ($null -ne $actualValue) {
                    $met = @($expectedArr | Where-Object { $actualArr -contains $_ }).Count -eq $expectedArr.Count
                }
                $labelReport.Workload = @{
                    Expected = $expectedArr -join ", "
                    Actual   = $actualArr -join ", "
                    Met      = $met
                }
            }

            # ContentType: check if all expected types are present in actual
            if ($null -ne $expected.ContentType) {
                $actualValue = $actualLabel.ContentType
                $expectedValue = $expected.ContentType
                $actualArr = @()
                $expectedArr = @()
                if ($actualValue -is [string]) {
                    $actualArr = $actualValue -split ',' | ForEach-Object { $_.Trim() }
                } elseif ($actualValue -is [array]) {
                    $actualArr = $actualValue | ForEach-Object { $_.Trim() }
                }
                if ($expectedValue -is [string]) {
                    $expectedArr = $expectedValue -split ',' | ForEach-Object { $_.Trim() }
                } elseif ($expectedValue -is [array]) {
                    $expectedArr = $expectedValue | ForEach-Object { $_.Trim() }
                }
                $met = $null
                if ($null -ne $actualValue) {
                    $met = @($expectedArr | Where-Object { $actualArr -contains $_ }).Count -eq $expectedArr.Count
                }
                $labelReport.ContentType = @{
                    Expected = $expectedArr -join ", "
                    Actual   = $actualArr -join ", "
                    Met      = $met
                }
            }

            # Autolabel
            if ($null -ne $expected.Autolabel) {
                $actualValue = ($actualLabel.Capabilities -contains "AutoLabel")
                $expectedValue = $expected.Autolabel
                $met = $null
                if ($null -ne $actualValue) {
                    $met = ($expectedValue -eq $actualValue)
                }
                $labelReport.Autolabel = @{
                    Expected = $expectedValue
                    Actual   = $actualValue
                    Met      = $met
                }
            }
            # Tooltip
            if ($expected.Tooltip -and $expected.Tooltip.Count -gt 0) {
                $tooltipTest = $expected.Tooltip[0]
                $actualTooltip = $actualLabel.Tooltip
                $tooltipResults = @{}

                # TextOnly
                if ($null -ne $tooltipTest.TextOnly) {
                    $met = $null
                    $actual = $null
                    if ($tooltipTest.TextOnly -eq "true") {
                        if ($null -ne $actualTooltip) {
                            $met = ($actualTooltip -ne "")
                            $actual = if ($actualTooltip -ne "") { "true" } else { "false" }
                        } else {
                            $actual = "false"
                        }
                    }
                    $tooltipResults.TextOnly = @{
                        Expected = $tooltipTest.TextOnly
                        Actual   = $actual
                        Met      = $met
                    }
                } else {
                    $tooltipResults.TextOnly = @{
                        Expected = $null
                        Actual   = $null
                        Met      = $null
                    }
                }

                # TextMatch
                if ($null -ne $tooltipTest.TextMatch) {
                    $met = $null
                    $actual = $actualTooltip
                    if ($null -ne $actualTooltip) {
                        $met = ($actualTooltip -eq $tooltipTest.TextMatch)
                    }
                    $tooltipResults.TextMatch = @{
                        Expected = $tooltipTest.TextMatch
                        Actual   = $actual
                        Met      = $met
                    }
                } else {
                    $tooltipResults.TextMatch = @{
                        Expected = $null
                        Actual   = $null
                        Met      = $null
                    }
                }

                # TextwithLink
                if ($null -ne $tooltipTest.TextwithLink) {
                    $met = $null
                    $actual = $null
                    if ($tooltipTest.TextwithLink -eq "true") {
                        if ($null -ne $actualTooltip) {
                            $met = ($actualTooltip -match 'https?://')
                            $actual = if ($actualTooltip -match 'https?://') { "true" } else { "false" }
                        } else {
                            $actual = "false"
                        }
                    }
                    $tooltipResults.TextwithLink = @{
                        Expected = $tooltipTest.TextwithLink
                        Actual   = $actual
                        Met      = $met
                    }
                } else {
                    $tooltipResults.TextwithLink = @{
                        Expected = $null
                        Actual   = $null
                        Met      = $null
                    }
                }

                $labelReport.Tooltip = $tooltipResults
            }
            # LabelActions
            if ($expected.LabelActions -and $expected.LabelActions.Count -gt 0) {
                foreach ($expAction in $expected.LabelActions) {
                    if ($null -eq $expAction.Type) { continue }
                    $matched = $false
                    foreach ($actActionStr in $actualLabel.LabelActions) {
                        $actAction = $null
                        try { $actAction = $actActionStr | ConvertFrom-Json } catch {}
                        if ($null -eq $actAction) { continue }

                        if ($expAction.Type -eq $actAction.Type) {
                            if ($expAction.Type -eq "applycontentmarking") {
                                $expText = $expAction.Text
                                $actText = ($actAction.Settings | Where-Object { $_.Key -eq "text" }).Value
                                $met = $null
                                if ($null -ne $actText) {
                                    $met = ($expText -eq $actText)
                                }
                                if ($met) {
                                    $matched = $true
                                }
                                $labelReport.LabelActions += @{
                                    Type = "applycontentmarking"
                                    Expected = $expText
                                    Actual = $actText
                                    Met = $met
                                }
                                if ($met) { break }
                            }
                            elseif ($expAction.Type -eq "encrypt") {
                                $expProt = $expAction.ProtectionType
                                $actProt = ($actAction.Settings | Where-Object { $_.Key -eq "protectiontype" }).Value
                                $expRights = $expAction.RightsDefinitions
                                $actRightsRaw = ($actAction.Settings | Where-Object { $_.Key -eq "rightsdefinitions" }).Value
                                $actRights = @()
                                if ($actRightsRaw) {
                                    try { $actRights = $actRightsRaw | ConvertFrom-Json } catch {}
                                }
                                $protMet = $null
                                if ($null -ne $actProt) {
                                    $protMet = ($expProt -eq $actProt)
                                }
                                $rightsMet = $null
                                if ($actRights.Count -gt 0) {
                                    $rightsMet = $false
                                    foreach ($expRD in $expRights) {
                                        foreach ($actRD in $actRights) {
                                            if ($expRD.Identity -eq $actRD.Identity -and
                                                ($expRD.Rights -join ",") -eq $actRD.Rights) {
                                                $rightsMet = $true
                                                break
                                            }
                                        }
                                    }
                                }
                                $met = $null
                                if ($null -ne $protMet -and $null -ne $rightsMet) {
                                    $met = ($protMet -and $rightsMet)
                                }
                                if ($met) {
                                    $matched = $true
                                }
                                $labelReport.LabelActions += @{
                                    Type = "encrypt"
                                    Expected = $expAction
                                    Actual = @{
                                        ProtectionType = $actProt
                                        RightsDefinitions = $actRights
                                    }
                                    Met = $met
                                }
                                if ($met) { break }
                            }
                        }
                    }
                    if (-not $matched) {
                        $labelReport.LabelActions += @{
                            Type = $expAction.Type
                            Expected = $expAction
                            Actual = $null
                            Met = $null
                        }
                        $labelReport.AllActionsMet = $false
                    }
                }
            }

            # Label Policy Settings
            $labelReport.LabelPolicySettings = @()
            $allPolicySettingsMet = $true

            if ($expected.LabelPolicySettings) {
                $matchingPolicy = $null
                foreach ($policy in $actual.GetLabelPolicy) {
 
                    if ($null -ne $actualLabel -and $policy.ScopedLabels -contains $actualLabel.ImmutableId) {
                        $matchingPolicy = $policy
                        break
                    }
                }
                if ($matchingPolicy) {
                    $actualSettings = @{}
                    foreach ($setting in $matchingPolicy.Settings) {
                        if ($setting -match "^\[(.+?),\s*(.+?)\]$") {
                            $actualSettings[$matches[1].Trim()] = $matches[2].Trim()
                        }
                    }
                    foreach ($property in $expected.LabelPolicySettings[0].psobject.Properties) {
                        if ($null -eq $property.Value) { continue }
                        $key = $property.Name
                        $expectedValue = $property.Value
                        $actualValue = $null
                        if ($actualSettings.ContainsKey($key)) {
                            $actualValue = $actualSettings[$key]
                        }
                        $met = $null
                        if ($null -ne $actualValue) {
                            # Normalize both to lower-case strings for comparison
                            $expectedStr = $expectedValue.ToString().ToLower()
                            $actualStr = $actualValue.ToString().ToLower()
                            # Handle boolean normalization
                            if (($expectedStr -eq "true" -or $expectedStr -eq "false") -and
                                ($actualStr -eq "true" -or $actualStr -eq "false")) {
                                $met = ([bool]::Parse($expectedStr) -eq [bool]::Parse($actualStr))
                            }
                            # Handle "none" as a special string (case-insensitive)
                            elseif ($expectedStr -eq "none" -or $actualStr -eq "none") {
                                $met = ($expectedStr -eq $actualStr)
                            }
                            # General string comparison (case-insensitive, trimmed)
                            else {
                                $met = ($expectedStr -eq $actualStr)
                            }
                        }
                        $labelReport.LabelPolicySettings += @{
                            Key = $key
                            Expected = $expectedValue
                            Actual = $actualValue
                            Met = $met
                        }
                        if ($met -eq $false) { $allPolicySettingsMet = $false }
                    }
                } else {
                    foreach ($property in $expected.LabelPolicySettings[0].psobject.Properties) {
                        if ($null -eq $property.Value) { continue }
                        $key = $property.Name
                        $expectedValue = $property.Value
                        $labelReport.LabelPolicySettings += @{
                            Key = $key
                            Expected = $expectedValue
                            Actual = $null
                            Met = $null
                        }
                    }
                    $allPolicySettingsMet = $false
                }
            }

            # Overall label met if all properties, all actions, and all policy settings are met
            $labelReport.OverallMet = $true
            if ($labelReport.Workload -and $labelReport.Workload.Met -eq $false) { $labelReport.OverallMet = $false }
            if ($labelReport.ContentType -and $labelReport.ContentType.Met -eq $false) { $labelReport.OverallMet = $false }
            if ($labelReport.Autolabel -and $labelReport.Autolabel.Met -eq $false) { $labelReport.OverallMet = $false }
            if (-not $labelReport.AllActionsMet) { $labelReport.OverallMet = $false }
            if (-not $allPolicySettingsMet) { $labelReport.OverallMet = $false }

            $report.SensitivityLabels += $labelReport
        }

        $finalReport += $report
    }

    # Output report as JSON
    $finalReport | ConvertTo-Json -Depth 6 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "Sensitivity label evaluation report written to $OutputPath"
}