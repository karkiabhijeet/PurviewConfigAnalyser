Import-Module "$PSScriptRoot\..\src\MaturityModelReport.psm1" -Force

Get-MaturityModelReport `
    -JsonFilePath "$PSScriptRoot\..\examples\OptimizedReport_7922a05c1dac422d972006bf4421e59b_20250505141615.json" `
    -ConfigJsonFilePath "$PSScriptRoot\..\examples\Config_sample.json"