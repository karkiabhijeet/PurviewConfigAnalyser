function Test-DLPPolicies {
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
            DLPPolicies = @()
            ComplianceScore = @{
                TotalExpected = 0
                TotalMet = 0
                Percentage = 0
            }
        }

        foreach ($expected in $config.DLPPolicies) {
            # Find matching DLP policy in actual config
            $actualPolicy = $actual.GetDlpCompliancePolicy | Where-Object {
                $_.Name -eq $expected.DisplayName -or
                $_.DisplayName -eq $expected.DisplayName
            }

            $policyReport = @{
                DisplayName = $expected.DisplayName
                Description = $expected.Description
                Enabled = $null
                Rules = @()
                Analytics = $null
                OverallMet = $true
            }

            $report.ComplianceScore.TotalExpected++

            # Check if policy is enabled
            if ($null -ne $expected.Enabled) {
                $actualEnabled = if ($actualPolicy) { $actualPolicy.Enabled } else { $false }
                $expectedEnabled = $expected.Enabled
                $met = $actualEnabled -eq $expectedEnabled

                $policyReport.Enabled = @{
                    Expected = $expectedEnabled
                    Actual = $actualEnabled
                    Met = $met
                }

                if (-not $met) {
                    $policyReport.OverallMet = $false
                }
            }

            # Check Rules if specified
            if ($expected.Rules -and $expected.Rules.Count -gt 0) {
                foreach ($expectedRule in $expected.Rules) {
                    $ruleReport = @{
                        Name = $expectedRule.Name
                        Conditions = $null
                        Actions = $null
                        Met = $true
                    }

                    # Find matching rule in actual policy
                    $actualRule = $null
                    if ($actualPolicy) {
                        $policyGuid = $actualPolicy.Guid
                        $actualRule = $actual.GetDlpComplianceRule | Where-Object {
                            $_.Policy -eq $policyGuid -and
                            ($_.Name -eq $expectedRule.Name -or $_.DisplayName -eq $expectedRule.Name)
                        }
                    }

                    # Check Conditions
                    if ($expectedRule.Conditions) {
                        $conditionsReport = @{
                            Expected = $expectedRule.Conditions
                            Actual = if ($actualRule) { $actualRule.Conditions } else { @() }
                            Met = $false
                        }

                        # Simple check if conditions exist
                        if ($actualRule -and $actualRule.Conditions) {
                            $conditionsReport.Met = $true
                        }

                        $ruleReport.Conditions = $conditionsReport
                        if (-not $conditionsReport.Met) {
                            $ruleReport.Met = $false
                        }
                    }

                    # Check Actions
                    if ($expectedRule.Actions) {
                        $actionsReport = @{
                            Expected = $expectedRule.Actions
                            Actual = if ($actualRule) { $actualRule.Actions } else { @() }
                            Met = $false
                        }

                        # Simple check if actions exist
                        if ($actualRule -and $actualRule.Actions) {
                            $actionsReport.Met = $true
                        }

                        $ruleReport.Actions = $actionsReport
                        if (-not $actionsReport.Met) {
                            $ruleReport.Met = $false
                        }
                    }

                    if (-not $ruleReport.Met) {
                        $policyReport.OverallMet = $false
                    }

                    $policyReport.Rules += $ruleReport
                }
            }

            # Check Analytics if specified
            if ($expected.Analytics) {
                $analyticsEnabled = $false
                
                # Check if DLP Analytics is enabled in the tenant
                if ($actual.TenantDetails -and $actual.TenantDetails.DLPAnalytics) {
                    $analyticsEnabled = $actual.TenantDetails.DLPAnalytics.Enabled
                }

                $policyReport.Analytics = @{
                    Expected = $expected.Analytics
                    Actual = $analyticsEnabled
                    Met = $analyticsEnabled
                }

                if (-not $analyticsEnabled) {
                    $policyReport.OverallMet = $false
                }
            }

            # Update compliance score
            if ($policyReport.OverallMet) {
                $report.ComplianceScore.TotalMet++
            }

            $report.DLPPolicies += $policyReport
        }

        # Calculate percentage
        if ($report.ComplianceScore.TotalExpected -gt 0) {
            $report.ComplianceScore.Percentage = [math]::Round(($report.ComplianceScore.TotalMet / $report.ComplianceScore.TotalExpected) * 100, 2)
        }

        # Overall compliance determination
        $report.OverallCompliance = @{
            DLPCompliant = $report.ComplianceScore.Percentage -eq 100
            ComplianceLevel = if ($report.ComplianceScore.Percentage -eq 100) { "Fully Compliant" }
                             elseif ($report.ComplianceScore.Percentage -ge 75) { "Mostly Compliant" }
                             elseif ($report.ComplianceScore.Percentage -ge 50) { "Partially Compliant" }
                             else { "Non-Compliant" }
        }

        $finalReport += $report
    }

    # Output report as JSON
    $finalReport | ConvertTo-Json -Depth 100 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "DLP policy evaluation report written to $OutputPath"
}
