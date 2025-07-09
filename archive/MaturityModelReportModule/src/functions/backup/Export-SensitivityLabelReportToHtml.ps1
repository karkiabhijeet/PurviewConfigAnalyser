function Export-SensitivityLabelReportToHtml {
    param (
        [Parameter(Mandatory)]
        [string]$JsonReportPath,
        [Parameter(Mandatory)]
        [string]$HtmlOutputPath
    )

    $report = Get-Content $JsonReportPath -Raw | ConvertFrom-Json

    $html = @"
<html>
<head>
    <title>Sensitivity Label Evaluation Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #2e6c80; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 30px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .met { background-color: #d4edda; }
        .notmet { background-color: #f8d7da; }
    </style>
</head>
<body>
    <h1>Sensitivity Label Evaluation Report</h1>
"@

    # Handle both array and single object
    if ($report -isnot [System.Collections.IEnumerable] -or $report -is [string]) {
        $report = @($report)
    }

    foreach ($model in $report) {
        $totalExpected = 0
        $totalMet = 0

        $html += "<h2>Maturity Model: $($model.MaturityModel)</h2>"

        foreach ($label in $model.SensitivityLabels) {
            $html += "<h3>Label: <b>$($label.DisplayName)</b></h3>"
            $html += "<table>"
            $html += "<tr><th>Property</th><th>Expected</th><th>Actual</th><th>Met?</th></tr>"

            # Workload
            if ($label.Workload) {
                $totalExpected++
                $metClass = if ($label.Workload.Met) { $totalMet++; "met" } else { "notmet" }
                $html += "<tr class='$metClass'><td>Workload</td><td>$($label.Workload.Expected)</td><td>$($label.Workload.Actual)</td><td>$($label.Workload.Met)</td></tr>"
            }

            # ContentType
            if ($label.ContentType) {
                $totalExpected++
                $metClass = if ($label.ContentType.Met) { $totalMet++; "met" } else { "notmet" }
                $html += "<tr class='$metClass'><td>ContentType</td><td>$($label.ContentType.Expected)</td><td>$($label.ContentType.Actual)</td><td>$($label.ContentType.Met)</td></tr>"
            }

            # Autolabel
            if ($label.Autolabel) {
                $totalExpected++
                $metClass = if ($label.Autolabel.Met) { $totalMet++; "met" } else { "notmet" }
                $html += "<tr class='$metClass'><td>Autolabel</td><td>$($label.Autolabel.Expected)</td><td>$($label.Autolabel.Actual)</td><td>$($label.Autolabel.Met)</td></tr>"
            }

            # LabelActions
            foreach ($action in $label.LabelActions) {
                $totalExpected++
                $metClass = if ($action.Met) { $totalMet++; "met" } else { "notmet" }
                $expected = if ($action.Expected -is [string]) { $action.Expected } else { ($action.Expected | ConvertTo-Json -Compress) }
                $actual = if ($action.Actual -is [string]) { $action.Actual } elseif ($null -ne $action.Actual) { ($action.Actual | ConvertTo-Json -Compress) } else { "" }
                $html += "<tr class='$metClass'><td>Action: $($action.Type)</td><td>$expected</td><td>$actual</td><td>$($action.Met)</td></tr>"
            }

            $html += "</table>"

            # --- Label Policy Settings Table ---
            if ($label.LabelPolicySettings -and $label.LabelPolicySettings.Count -gt 0) {
                $html += "<h4>Label Policy Settings</h4>"
                $html += "<table>"
                $html += "<tr><th>Setting</th><th>Expected</th><th>Actual</th><th>Met?</th></tr>"
                foreach ($setting in $label.LabelPolicySettings) {
                    $totalExpected++
                    $metClass = if ($setting.Met) { $totalMet++; "met" } else { "notmet" }
                    $html += "<tr class='$metClass'><td>$($setting.Key)</td><td>$($setting.Expected)</td><td>$($setting.Actual)</td><td>$($setting.Met)</td></tr>"
                }
                $html += "</table>"
            }

            $html += "<p><b>Overall Met:</b> <span class='$(if ($label.OverallMet) { "met" } else { "notmet" })'>$($label.OverallMet)</span></p>"
        }

        $html += "<h3>Summary for $($model.MaturityModel):</h3>"
        $html += "<p><b>Total Expected:</b> $totalExpected<br/><b>Total Achieved:</b> $totalMet</p>"
        $html += "<hr/>"
    }

    $html += "</body></html>"

    Set-Content -Path $HtmlOutputPath -Value $html -Encoding UTF8
    Write-Host "HTML sensitivity label report written to $HtmlOutputPath"
}