function Export-SensitivityLabelReportToHtml {
    param (
        [Parameter(Mandatory)]
        [string]$JsonReportPath,
        [Parameter(Mandatory)]
        [string]$HtmlOutputPath
    )

    $report = Get-Content $JsonReportPath -Raw | ConvertFrom-Json

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sensitivity Label Compliance Report</title>
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
        h4 {
            color: #605e5c;
            margin-top: 15px;
            margin-bottom: 8px;
            font-size: 1.1em;
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
        .label-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .label-table th,
        .label-table td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e1dfdd;
        }
        .label-table th {
            background-color: #f3f2f1;
            font-weight: 600;
            color: #323130;
        }
        .label-table tr:hover {
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
        .label-section {
            background-color: #f8f8f8;
            margin: 20px 0;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #0078d4;
        }
        .label-header {
            background: white;
            margin: -10px -10px 15px -10px;
            padding: 15px;
            border-radius: 6px;
            border-bottom: 1px solid #e1dfdd;
        }
        .label-name {
            font-size: 1.3em;
            font-weight: bold;
            color: #323130;
            margin-bottom: 5px;
        }
        .label-status {
            font-size: 1.1em;
            margin-top: 10px;
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
        .model-summary {
            background: linear-gradient(135deg, #f3f2f1, #e1dfdd);
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
            border-left: 4px solid #605e5c;
        }
        .summary-stats {
            display: flex;
            gap: 20px;
            justify-content: space-around;
            flex-wrap: wrap;
        }
        .stat-item {
            text-align: center;
        }
        .stat-number {
            font-size: 1.8em;
            font-weight: bold;
            color: #323130;
        }
        .stat-label {
            color: #605e5c;
            font-size: 0.9em;
        }
        @media (max-width: 768px) {
            .container {
                padding: 15px;
            }
            .summary-cards {
                flex-direction: column;
            }
            .label-table {
                font-size: 0.9em;
            }
            .label-table th,
            .label-table td {
                padding: 8px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Sensitivity Label Compliance Report</h1>
        <div class="timestamp">Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
"@    # Handle both array and single object
    if ($report -isnot [System.Collections.IEnumerable] -or $report -is [string]) {
        $report = @($report)
    }

    # Calculate overall statistics
    $overallExpected = 0
    $overallMet = 0
    $modelCount = $report.Count

    foreach ($model in $report) {
        $totalExpected = 0
        $totalMet = 0

        $html += "<h2>Maturity Model: $($model.MaturityModel)</h2>"

        # Summary cards for this model
        $labelCount = if ($model.SensitivityLabels) { $model.SensitivityLabels.Count } else { 0 }
        
        foreach ($label in $model.SensitivityLabels) {
            $labelExpected = 0
            $labelMet = 0

            $html += "<div class='label-section'>"
            $html += "<div class='label-header'>"
            $html += "<div class='label-name'>$($label.DisplayName)</div>"
            
            # Calculate label compliance
            $labelProperties = @()
            
            if ($label.Workload) { $labelProperties += $label.Workload }
            if ($label.ContentType) { $labelProperties += $label.ContentType }
            if ($label.Autolabel) { $labelProperties += $label.Autolabel }
            if ($label.LabelActions) { $labelProperties += $label.LabelActions }
            if ($label.LabelPolicySettings) { $labelProperties += $label.LabelPolicySettings }
            
            $labelCompliance = if ($labelProperties.Count -gt 0) { 
                $metCount = ($labelProperties | Where-Object { $_.Met -eq $true }).Count
                [math]::Round(($metCount / $labelProperties.Count) * 100, 1)
            } else { 0 }
            
            $html += "<div class='compliance-bar'>"
            $html += "<div class='compliance-progress' style='width: $labelCompliance%'></div>"
            $html += "</div>"
            $html += "<div class='label-status'><b>Overall Status:</b> <span class='$(if ($label.OverallMet) { "status-met" } else { "status-not-met" })'>$(if ($label.OverallMet) { "Compliant" } else { "Non-Compliant" })</span> ($labelCompliance% Complete)</div>"
            $html += "</div>"

            $html += "<table class='label-table'>"
            $html += "<tr><th>Property</th><th>Expected</th><th>Actual</th><th>Status</th></tr>"

            # Workload
            if ($label.Workload) {
                $totalExpected++; $labelExpected++
                $statusClass = if ($label.Workload.Met) { $totalMet++; $labelMet++; "status-met" } else { "status-not-met" }
                $statusText = if ($label.Workload.Met) { "✓ Met" } else { "✗ Not Met" }
                $html += "<tr><td>Workload</td><td>$($label.Workload.Expected)</td><td>$($label.Workload.Actual)</td><td class='$statusClass'>$statusText</td></tr>"
            }

            # ContentType
            if ($label.ContentType) {
                $totalExpected++; $labelExpected++
                $statusClass = if ($label.ContentType.Met) { $totalMet++; $labelMet++; "status-met" } else { "status-not-met" }
                $statusText = if ($label.ContentType.Met) { "✓ Met" } else { "✗ Not Met" }
                $html += "<tr><td>ContentType</td><td>$($label.ContentType.Expected)</td><td>$($label.ContentType.Actual)</td><td class='$statusClass'>$statusText</td></tr>"
            }

            # Autolabel
            if ($label.Autolabel) {
                $totalExpected++; $labelExpected++
                $statusClass = if ($label.Autolabel.Met) { $totalMet++; $labelMet++; "status-met" } else { "status-not-met" }
                $statusText = if ($label.Autolabel.Met) { "✓ Met" } else { "✗ Not Met" }
                $html += "<tr><td>Autolabel</td><td>$($label.Autolabel.Expected)</td><td>$($label.Autolabel.Actual)</td><td class='$statusClass'>$statusText</td></tr>"
            }

            # LabelActions
            foreach ($action in $label.LabelActions) {
                $totalExpected++; $labelExpected++
                $statusClass = if ($action.Met) { $totalMet++; $labelMet++; "status-met" } else { "status-not-met" }
                $statusText = if ($action.Met) { "✓ Met" } else { "✗ Not Met" }
                $expected = if ($action.Expected -is [string]) { $action.Expected } else { ($action.Expected | ConvertTo-Json -Compress) }
                $actual = if ($action.Actual -is [string]) { $action.Actual } elseif ($null -ne $action.Actual) { ($action.Actual | ConvertTo-Json -Compress) } else { "Not Configured" }
                $html += "<tr><td>Action: $($action.Type)</td><td>$expected</td><td>$actual</td><td class='$statusClass'>$statusText</td></tr>"
            }

            $html += "</table>"

            # Label Policy Settings Table
            if ($label.LabelPolicySettings -and $label.LabelPolicySettings.Count -gt 0) {
                $html += "<h4>Label Policy Settings</h4>"
                $html += "<table class='label-table'>"
                $html += "<tr><th>Setting</th><th>Expected</th><th>Actual</th><th>Status</th></tr>"
                foreach ($setting in $label.LabelPolicySettings) {
                    $totalExpected++; $labelExpected++
                    $statusClass = if ($setting.Met) { $totalMet++; $labelMet++; "status-met" } else { "status-not-met" }
                    $statusText = if ($setting.Met) { "✓ Met" } else { "✗ Not Met" }
                    $actualValue = if ($setting.Actual) { $setting.Actual } else { "Not Configured" }
                    $html += "<tr><td>$($setting.Key)</td><td>$($setting.Expected)</td><td>$actualValue</td><td class='$statusClass'>$statusText</td></tr>"
                }
                $html += "</table>"
            }

            $html += "</div>"
        }

        # Model Summary
        $modelCompliance = if ($totalExpected -gt 0) { [math]::Round(($totalMet / $totalExpected) * 100, 1) } else { 0 }
        $overallExpected += $totalExpected
        $overallMet += $totalMet

        $html += "<div class='model-summary'>"
        $html += "<h3>Summary for $($model.MaturityModel)</h3>"
        $html += "<div class='compliance-bar'>"
        $html += "<div class='compliance-progress' style='width: $modelCompliance%'></div>"
        $html += "</div>"
        $html += "<div class='summary-stats'>"
        $html += "<div class='stat-item'><div class='stat-number'>$labelCount</div><div class='stat-label'>Labels Evaluated</div></div>"
        $html += "<div class='stat-item'><div class='stat-number'>$totalExpected</div><div class='stat-label'>Total Requirements</div></div>"
        $html += "<div class='stat-item'><div class='stat-number'>$totalMet</div><div class='stat-label'>Requirements Met</div></div>"
        $html += "<div class='stat-item'><div class='stat-number'>$modelCompliance%</div><div class='stat-label'>Compliance Score</div></div>"
        $html += "</div>"
        $html += "</div>"
    }

    # Overall Summary Cards
    $overallCompliance = if ($overallExpected -gt 0) { [math]::Round(($overallMet / $overallExpected) * 100, 1) } else { 0 }
    
    $html += "<h2>Overall Summary</h2>"
    $html += "<div class='summary-cards'>"
    $html += "<div class='summary-card'>"
    $html += "<h3>Maturity Models</h3>"
    $html += "<div class='summary-number'>$modelCount</div>"
    $html += "</div>"
    $html += "<div class='summary-card'>"
    $html += "<h3>Total Requirements</h3>"
    $html += "<div class='summary-number'>$overallExpected</div>"
    $html += "</div>"
    $html += "<div class='summary-card'>"
    $html += "<h3>Requirements Met</h3>"
    $html += "<div class='summary-number'>$overallMet</div>"
    $html += "</div>"
    $html += "<div class='summary-card'>"
    $html += "<h3>Overall Compliance</h3>"
    $html += "<div class='summary-number'>$overallCompliance%</div>"
    $html += "</div>"
    $html += "</div>"

    $html += "</div></body></html>"

    Set-Content -Path $HtmlOutputPath -Value $html -Encoding UTF8
    Write-Host "HTML sensitivity label report written to $HtmlOutputPath"
}