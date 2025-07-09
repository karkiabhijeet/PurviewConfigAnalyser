function Export-DLPReportToHtml {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ReportFilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath
    )

    # Read the DLP report JSON
    if (!(Test-Path $ReportFilePath)) {
        throw "DLP report file not found: $ReportFilePath"
    }
    
    $dlpReport = Get-Content $ReportFilePath -Raw | ConvertFrom-Json

    # Start building HTML content
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DLP Policy Compliance Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0078d4;
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        h2 {
            color: #323130;
            margin-top: 30px;
            margin-bottom: 15px;
            border-left: 4px solid #0078d4;
            padding-left: 15px;
        }
        h3 {
            color: #605e5c;
            margin-top: 20px;
            margin-bottom: 10px;
        }
        .summary-cards {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-bottom: 30px;
        }
        .summary-card {
            flex: 1;
            min-width: 250px;
            background: linear-gradient(135deg, #0078d4, #106ebe);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        .summary-card h3 {
            color: white;
            margin: 0 0 10px 0;
        }
        .summary-number {
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }
        .policy-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .policy-table th,
        .policy-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e1dfdd;
        }
        .policy-table th {
            background-color: #f3f2f1;
            font-weight: 600;
            color: #323130;
        }
        .policy-table tr:hover {
            background-color: #f8f8f8;
        }
        .status-met {
            color: #107c10;
            font-weight: bold;
        }
        .status-not-met {
            color: #d13438;
            font-weight: bold;
        }
        .compliance-bar {
            width: 100%;
            height: 20px;
            background-color: #e1dfdd;
            border-radius: 10px;
            overflow: hidden;
            margin: 10px 0;
        }
        .compliance-progress {
            height: 100%;
            background: linear-gradient(90deg, #107c10, #13a10e);
            border-radius: 10px;
            transition: width 0.3s ease;
        }
        .rule-details {
            background-color: #f8f8f8;
            margin: 10px 0;
            padding: 15px;
            border-radius: 5px;
            border-left: 4px solid #0078d4;
        }
        .rule-item {
            margin: 8px 0;
            padding: 8px;
            background: white;
            border-radius: 4px;
            border: 1px solid #e1dfdd;
        }
        .rule-name {
            font-weight: bold;
            color: #323130;
        }
        .rule-status {
            margin-top: 5px;
        }
        .analytics-section {
            background-color: #fff4ce;
            padding: 15px;
            border-radius: 5px;
            border-left: 4px solid #ffb900;
            margin: 10px 0;
        }
        .expandable {
            cursor: pointer;
            user-select: none;
        }
        .expandable:hover {
            background-color: #f3f2f1;
        }
        .collapsible-content {
            display: none;
            margin-top: 10px;
        }
        .timestamp {
            text-align: center;
            color: #605e5c;
            font-style: italic;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e1dfdd;
        }
    </style>
    <script>
        function toggleSection(element) {
            const content = element.nextElementSibling;
            if (content.style.display === 'none' || content.style.display === '') {
                content.style.display = 'block';
                element.innerHTML = element.innerHTML.replace('▶', '▼');
            } else {
                content.style.display = 'none';
                element.innerHTML = element.innerHTML.replace('▼', '▶');
            }
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>🔒 Data Loss Prevention (DLP) Policy Compliance Report</h1>
"@

    # Add summary section
    $totalPolicies = 0
    $totalCompliantPolicies = 0
    $overallScore = 0
    
    foreach ($maturityModel in $dlpReport) {
        $totalPolicies += $maturityModel.DLPPolicies.Count
        $totalCompliantPolicies += ($maturityModel.DLPPolicies | Where-Object { $_.OverallMet }).Count
        $overallScore += $maturityModel.ComplianceScore.Percentage
    }
    
    if ($dlpReport.Count -gt 0) {
        $averageScore = [math]::Round($overallScore / $dlpReport.Count, 1)
    } else {
        $averageScore = 0
    }

    $htmlContent += @"
        <div class="summary-cards">
            <div class="summary-card">
                <h3>Maturity Models</h3>
                <div class="summary-number">$($dlpReport.Count)</div>
                <p>Total evaluated</p>
            </div>
            <div class="summary-card">
                <h3>DLP Policies</h3>
                <div class="summary-number">$totalPolicies</div>
                <p>Total policies evaluated</p>
            </div>
            <div class="summary-card">
                <h3>Compliant Policies</h3>
                <div class="summary-number">$totalCompliantPolicies</div>
                <p>Fully compliant policies</p>
            </div>
            <div class="summary-card">
                <h3>Average Score</h3>
                <div class="summary-number">$averageScore%</div>
                <p>Overall compliance score</p>
            </div>
        </div>
"@

    # Add detailed report for each maturity model
    foreach ($maturityModel in $dlpReport) {
        $htmlContent += @"
        <h2>📊 $($maturityModel.MaturityModel)</h2>
        <div style="margin-bottom: 20px;">
            <strong>Compliance Score:</strong> $($maturityModel.ComplianceScore.TotalMet) / $($maturityModel.ComplianceScore.TotalExpected) ($($maturityModel.ComplianceScore.Percentage)%)
            <div class="compliance-bar">
                <div class="compliance-progress" style="width: $($maturityModel.ComplianceScore.Percentage)%;"></div>
            </div>
        </div>
"@

        if ($maturityModel.DLPPolicies.Count -gt 0) {
            $htmlContent += @"
        <table class="policy-table">
            <thead>
                <tr>
                    <th>Policy Name</th>
                    <th>Description</th>
                    <th>Enabled Status</th>
                    <th>Rules</th>
                    <th>Analytics</th>
                    <th>Overall Status</th>
                </tr>
            </thead>
            <tbody>
"@

            foreach ($policy in $maturityModel.DLPPolicies) {
                $enabledStatus = if ($policy.Enabled.Met) { 
                    "<span class='status-met'>✓ Met</span>" 
                } else { 
                    "<span class='status-not-met'>✗ Not Met</span>" 
                }
                
                $overallStatus = if ($policy.OverallMet) { 
                    "<span class='status-met'>✓ Compliant</span>" 
                } else { 
                    "<span class='status-not-met'>✗ Non-Compliant</span>" 
                }

                $rulesCount = $policy.Rules.Count
                $analyticsStatus = if ($policy.Analytics) {
                    if ($policy.Analytics.Met) { "✓ Enabled" } else { "✗ Not Enabled" }
                } else { "N/A" }

                $htmlContent += @"
                <tr>
                    <td><strong>$($policy.DisplayName)</strong></td>
                    <td>$($policy.Description)</td>
                    <td>$enabledStatus<br><small>Expected: $($policy.Enabled.Expected) | Actual: $($policy.Enabled.Actual)</small></td>
                    <td>$rulesCount rule(s)</td>
                    <td>$analyticsStatus</td>
                    <td>$overallStatus</td>
                </tr>
"@

                # Add detailed rules section if there are rules
                if ($policy.Rules.Count -gt 0) {
                    $htmlContent += @"
                <tr>
                    <td colspan="6">
                        <div class="expandable" onclick="toggleSection(this)">
                            ▶ <strong>Rule Details for $($policy.DisplayName)</strong>
                        </div>
                        <div class="collapsible-content rule-details">
"@

                    foreach ($rule in $policy.Rules) {
                        $ruleEnabledStatus = if ($rule.Enabled.Met) { "✓" } else { "✗" }
                        $conditionsStatus = if ($rule.Conditions.Expected) {
                            if ($rule.Conditions.Met) { "✓" } else { "✗" }
                        } else { "N/A" }
                        $actionsStatus = if ($rule.Actions.Expected) {
                            if ($rule.Actions.Met) { "✓" } else { "✗" }
                        } else { "N/A" }

                        $htmlContent += @"
                            <div class="rule-item">
                                <div class="rule-name">$($rule.Name)</div>
                                <div class="rule-status">
                                    <strong>Enabled:</strong> $ruleEnabledStatus (Expected: $($rule.Enabled.Expected) | Actual: $($rule.Enabled.Actual))<br>
                                    <strong>Conditions:</strong> $conditionsStatus<br>
                                    <strong>Actions:</strong> $actionsStatus
                                </div>
                            </div>
"@
                    }

                    $htmlContent += @"
                        </div>
                    </td>
                </tr>
"@
                }

                # Add analytics details if present
                if ($policy.Analytics) {
                    $htmlContent += @"
                <tr>
                    <td colspan="6">
                        <div class="analytics-section">
                            <strong>📈 DLP Analytics Configuration</strong><br>
                            Expected: $($policy.Analytics.Expected -join ', ')<br>
                            Status: $(if ($policy.Analytics.Met) { '<span class="status-met">✓ Enabled</span>' } else { '<span class="status-not-met">✗ Not Enabled</span>' })
                        </div>
                    </td>
                </tr>
"@
                }
            }

            $htmlContent += @"
            </tbody>
        </table>
"@
        } else {
            $htmlContent += "<p><em>No DLP policies defined for this maturity model.</em></p>"
        }
    }

    # Add footer
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $htmlContent += @"
        <div class="timestamp">
            Report generated on $timestamp
        </div>
    </div>
</body>
</html>
"@

    # Write HTML to file
    $htmlContent | Set-Content -Path $OutputFilePath -Encoding UTF8
    Write-Host "DLP HTML report exported to: $OutputFilePath" -ForegroundColor Green
}
