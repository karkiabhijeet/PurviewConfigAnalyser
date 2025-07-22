function Convert-ObjectForJson {
    <#
    .SYNOPSIS
        Converts PowerShell objects to JSON-serializable format.
    
    .DESCRIPTION
        Preprocesses PowerShell objects to ensure proper JSON serialization,
        handling arrays, hashtables, and custom objects.
    #>
    
    param (
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [hashtable] -or $InputObject -is [System.Collections.IDictionary]) {
            $newHash = [ordered]@{}
            foreach ($key in $InputObject.Keys) {
                $stringKey = [string]$key
                $newHash[$stringKey] = Convert-ObjectForJson -InputObject $InputObject[$key]
            }
            return [PSCustomObject]$newHash
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $newList = [System.Collections.ArrayList]::new()
            foreach ($item in $InputObject) {
                $null = $newList.Add((Convert-ObjectForJson -InputObject $item))
            }
            return $newList
        }

        if ($InputObject -is [PSCustomObject]) {
            $newObj = [ordered]@{}
            foreach ($prop in $InputObject.PSObject.Properties) {
                $propName = $prop.Name
                $propValue = $prop.Value

                # Targeted fix for arrays that get truncated
                if ($propName -in @("Labels", "ScopedLabels") -and $propValue -is [System.Collections.IEnumerable] -and $propValue -isnot [string]) {
                    $newObj[$propName] = @($propValue | ForEach-Object { "$_" })
                }
                # Preserve LabelActions, which are often pre-formatted JSON strings
                elseif ($propName -in @("LabelActions", "Settings", "LocaleSettings") -and $propValue -is [System.Collections.IEnumerable] -and $propValue -isnot [string]) {
                     $newObj[$propName] = $propValue # Preserve original structure
                }
                else {
                    $newObj[$propName] = Convert-ObjectForJson -InputObject $propValue
                }
            }
            return [PSCustomObject]$newObj
        }

        # For all other types, return as is
        return $InputObject
    }
}
