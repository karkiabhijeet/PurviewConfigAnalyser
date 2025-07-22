# Script to replace Unicode icons with text equivalents for better PowerShell compatibility

$replacements = @{
    '✅' = '[SUCCESS]'
    '❌' = '[ERROR]'
    '⚠️' = '[WARNING]'
    '📋' = '[INFO]'
    '📝' = '[EXAMPLE]'
    '💡' = '[TIP]'
    '🎉' = '[COMPLETE]'
    '📦' = '[PACKAGE]'
    '⏳' = '[WAIT]'
    '🚀' = '[LAUNCH]'
    '🛠️' = '[TOOLS]'
    '🔧' = '[CONFIG]'
    '📊' = '[REPORT]'
    '📈' = '[ANALYTICS]'
    '🔗' = '[LINK]'
}

$filesToProcess = @(
    "src\Public\Test-PurviewCompliance.ps1",
    "src\Public\Invoke-PurviewConfigAnalyser.ps1", 
    "src\Public\Get-PurviewConfig.ps1",
    "src\Collect-PurviewConfiguration.ps1",
    "src\Private\Connect-ToComplianceCenter.ps1"
)

foreach ($file in $filesToProcess) {
    $filePath = Join-Path $PSScriptRoot $file
    if (Test-Path $filePath) {
        Write-Host "Processing: $file" -ForegroundColor Yellow
        
        $content = Get-Content $filePath -Raw -Encoding UTF8
        $originalContent = $content
        
        foreach ($unicode in $replacements.Keys) {
            $replacement = $replacements[$unicode]
            $content = $content -replace [regex]::Escape($unicode), $replacement
        }
        
        if ($content -ne $originalContent) {
            Set-Content $filePath -Value $content -Encoding UTF8 -NoNewline
            Write-Host "  -> Updated with text equivalents" -ForegroundColor Green
        } else {
            Write-Host "  -> No changes needed" -ForegroundColor Gray
        }
    } else {
        Write-Host "File not found: $filePath" -ForegroundColor Red
    }
}

Write-Host "`nReplacement complete! All Unicode icons have been replaced with text equivalents." -ForegroundColor Green
Write-Host "This should improve compatibility across different PowerShell environments." -ForegroundColor Cyan
