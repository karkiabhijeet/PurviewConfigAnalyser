<#
.SYNOPSIS
    Evaluates DLP (Data Loss Prevention) policies and rules from an OptimizedReport JSON file.
.DESCRIPTION
    - Only considers DLP policies with Mode=Enable as Active for compliance evaluation.
    - Maps each GetDlpCompliancePolicy (parent) to its GetDlpComplianceRule (child) by Guid/Policy.
    - Unpacks AdvancedRule JSON in each rule and recursively extracts all Sensitivetypes > Name.
    - Flags simulation/test policies (Mode=TestWithNotifications, Mode=TestWithoutNotifications, IsSimulationPolicy=true) as informational.
    - Outputs a summary of active DLP policies, their rules, workloads, and detected sensitive types.
.PARAMETER ReportPath
    Path to the OptimizedReport JSON file.
.EXAMPLE
    .\DlpPolicyEvaluator.ps1 -ReportPath .\output\OptimizedReport_*.json
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ReportPath
)

function Get-SensitiveTypesRecursive {
    param([object]$Node)
    $results = @()
    if ($null -eq $Node) { return $results }
    if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
        foreach ($item in $Node) {
            $results += Get-SensitiveTypesRecursive -Node $item
        }
    } elseif ($Node -is [hashtable] -or $Node -is [PSCustomObject]) {
        foreach ($key in $Node.PSObject.Properties.Name) {
            if ($key -eq 'Sensitivetypes' -and $Node[$key]) {
                foreach ($stype in $Node[$key]) {
                    if ($stype.Name) { $results += $stype.Name }
                }
            } else {
                $results += Get-SensitiveTypesRecursive -Node $Node[$key]
            }
        }
    }
    return $results
}

if (-not (Test-Path $ReportPath)) {
    Write-Host "File not found: $ReportPath" -ForegroundColor Red
    exit 1
}

$json = Get-Content $ReportPath -Raw | ConvertFrom-Json
$policies = $json.GetDlpCompliancePolicy
$rules = $json.GetDlpComplianceRule

if (-not $policies -or -not $rules) {
    Write-Host "No DLP policies or rules found in report." -ForegroundColor Yellow
    exit 0
}

# Build a lookup for rules by Policy GUID
$rulesByPolicy = @{}
foreach ($rule in $rules) {
    if ($rule.Policy) {
        if (-not $rulesByPolicy.ContainsKey($rule.Policy)) {
            $rulesByPolicy[$rule.Policy] = @()
        }
        $rulesByPolicy[$rule.Policy] += $rule
    }
}

foreach ($policy in $policies) {
    $isSimulation = $false
    if ($policy.Mode -eq 'Enable') {
        $status = 'Active'
    } elseif ($policy.Mode -like 'Test*' -or ($policy.PSObject.Properties.Name -contains 'IsSimulationPolicy' -and $policy.IsSimulationPolicy)) {
        $status = 'Simulation'
        $isSimulation = $true
    } else {
        $status = 'Disabled'
    }
    Write-Host "Policy: $($policy.Name) [$status]" -ForegroundColor Cyan
    Write-Host "  Mode: $($policy.Mode)  Guid: $($policy.Guid)"
    if ($isSimulation) {
        Write-Host "  [INFO] This policy is in simulation/test mode and will not be used for compliance pass/fail." -ForegroundColor Yellow
    }
    if ($status -ne 'Active') { continue }
    if ($rulesByPolicy.ContainsKey($policy.Guid)) {
        foreach ($rule in $rulesByPolicy[$policy.Guid]) {
            Write-Host "    Rule: $($rule.Name)"
            Write-Host "      Workload: $($rule.Workload)"
            if ($rule.PSObject.Properties.Name -contains 'AdvancedRule' -and $rule.AdvancedRule) {
                try {
                    $adv = $rule.AdvancedRule | ConvertFrom-Json -ErrorAction Stop
                    $stypeNames = Get-SensitiveTypesRecursive -Node $adv | Select-Object -Unique
                    if ($stypeNames) {
                        Write-Host "      Sensitive Types: $($stypeNames -join ', ')" -ForegroundColor Green
                    } else {
                        Write-Host "      Sensitive Types: None found" -ForegroundColor DarkGray
                    }
                } catch {
                    Write-Host "      [ERROR] Failed to parse AdvancedRule JSON." -ForegroundColor Red
                }
            } else {
                Write-Host "      No AdvancedRule present." -ForegroundColor DarkGray
            }
        }
    } else {
        Write-Host "    No rules found for this policy." -ForegroundColor DarkGray
    }
}
