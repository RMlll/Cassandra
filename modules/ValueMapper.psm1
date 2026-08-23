<#
    Excel表示値をAPI値へ変換する内部モジュール。
    JSON構造の組み立てとApplyWhenAny判定が完了した後、出力直前に適用する。
#>

function Import-ValueMappingConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RulesPath,
        [Parameter(Mandatory = $true)]
        [string]$MappingsPath
    )

    if (-not (Test-Path -LiteralPath $RulesPath -PathType Leaf)) {
        throw "ルールJSONがありません。Export-JsonRules.ps1を実行してください: $RulesPath"
    }
    if (-not (Test-Path -LiteralPath $MappingsPath -PathType Leaf)) {
        throw "マッピングJSONがありません。Export-JsonRules.ps1を実行してください: $MappingsPath"
    }

    $rulesDocument = Get-Content -LiteralPath $RulesPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $mappingsDocument = Get-Content -LiteralPath $MappingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $typesByPath = @{}
    foreach ($rule in @($rulesDocument.rules)) {
        $path = ([string]$rule.path).Trim()
        if ($path -ne "") { $typesByPath[$path] = ([string]$rule.type).Trim().ToLower() }
    }

    $valuesByPath = @{}
    $mappingCount = 0
    foreach ($mapping in @($mappingsDocument.mappings)) {
        $path = ([string]$mapping.path).Trim()
        $inputValue = ([string]$mapping.inputValue).Trim()
        $apiValue = ([string]$mapping.apiValue).Trim()
        $action = ([string]$mapping.action).Trim().ToLower()
        if ($action -eq "") { $action = "set" }
        if ($path -eq "" -or $inputValue -eq "") { continue }
        if (-not $typesByPath.ContainsKey($path)) {
            throw "マッピングのJSONPathがルールにありません: $path"
        }
        if (-not $valuesByPath.ContainsKey($path)) { $valuesByPath[$path] = @{} }
        if ($valuesByPath[$path].ContainsKey($inputValue)) {
            throw "マッピングが重複しています: $path / $inputValue"
        }
        if ($action -notin @("set", "omit")) {
            throw "マッピングのActionが不正です: $path / $inputValue / $action"
        }
        $valuesByPath[$path][$inputValue] = [pscustomobject]@{
            Action = $action
            ApiValue = $apiValue
        }
        $mappingCount++
    }

    return [pscustomobject]@{
        TypesByPath = $typesByPath
        ValuesByPath = $valuesByPath
        MappingCount = $mappingCount
    }
}

function ConvertTo-ApiTypedValue {
    param(
        [object]$Value,
        [string]$Type,
        [string]$Path
    )

    $text = ([string]$Value).Trim()
    switch ($Type) {
        "boolean" {
            if ($text -eq "true") { return $true }
            if ($text -eq "false") { return $false }
            throw "$Path のマッピング結果をbooleanへ変換できません: $Value"
        }
        "integer" {
            [long]$number = 0
            if (-not [long]::TryParse($text, [ref]$number)) {
                throw "$Path のマッピング結果をintegerへ変換できません: $Value"
            }
            return $number
        }
        "number" {
            [double]$number = 0
            if (-not [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
                throw "$Path のマッピング結果をnumberへ変換できません: $Value"
            }
            return $number
        }
        "string" { return $text }
        default { return $Value }
    }
}

function Get-MappingResult {
    param(
        [object]$Value,
        [hashtable]$PathMappings,
        [string]$Type,
        [string]$Path
    )

    if ($null -eq $Value) {
        return [pscustomobject]@{ Matched=$false; Omit=$false; Value=$Value }
    }
    $lookupValue = ([string]$Value).Trim()
    if (-not $PathMappings.ContainsKey($lookupValue)) {
        return [pscustomobject]@{ Matched=$false; Omit=$false; Value=$Value }
    }
    $mapping = $PathMappings[$lookupValue]
    if ([string]$mapping.Action -eq "omit") {
        return [pscustomobject]@{ Matched=$true; Omit=$true; Value=$null }
    }
    $converted = ConvertTo-ApiTypedValue $mapping.ApiValue $Type $Path
    return [pscustomobject]@{ Matched=$true; Omit=$false; Value=$converted }
}

function Set-MappedPathValue {
    param(
        [object]$Node,
        [string[]]$Segments,
        [int]$SegmentIndex,
        [hashtable]$PathMappings,
        [string]$Type,
        [string]$Path
    )

    if ($Node -isnot [System.Collections.IDictionary]) { return }

    $segment = $Segments[$SegmentIndex]
    $isArray = $segment.EndsWith("[]")
    $name = if ($isArray) { $segment.Substring(0, $segment.Length - 2) } else { $segment }
    if (-not $Node.Contains($name)) { return }
    $isLast = $SegmentIndex -eq ($Segments.Count - 1)

    if ($isArray) {
        $items = $Node[$name]
        if ($items -isnot [System.Collections.IList] -or $items -is [string]) { return }
        for ($index = 0; $index -lt $items.Count; $index++) {
            if ($isLast) {
                $result = Get-MappingResult $items[$index] $PathMappings $Type $Path
                if ($result.Omit) {
                    throw "$Path は配列要素自体のOmitに対応していません。プロパティを対象にしてください。"
                }
                if ($result.Matched) { $items[$index] = $result.Value }
            }
            else {
                Set-MappedPathValue $items[$index] $Segments ($SegmentIndex + 1) $PathMappings $Type $Path
            }
        }
        return
    }

    if ($isLast) {
        $result = Get-MappingResult $Node[$name] $PathMappings $Type $Path
        if ($result.Matched -and $result.Omit) {
            [void]$Node.Remove($name)
        }
        elseif ($result.Matched) {
            $Node[$name] = $result.Value
        }
    }
    else {
        Set-MappedPathValue $Node[$name] $Segments ($SegmentIndex + 1) $PathMappings $Type $Path
    }
}

function ConvertTo-ApiValues {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Body,
        [Parameter(Mandatory = $true)]
        [object]$Configuration
    )

    foreach ($path in $Configuration.ValuesByPath.Keys) {
        $segments = @($path.Split('.'))
        if ($segments.Count -eq 0) { continue }
        Set-MappedPathValue `
            -Node $Body `
            -Segments $segments `
            -SegmentIndex 0 `
            -PathMappings $Configuration.ValuesByPath[$path] `
            -Type ([string]$Configuration.TypesByPath[$path]) `
            -Path $path
    }

    return $Body
}

Export-ModuleMember -Function @(
    "Import-ValueMappingConfiguration",
    "ConvertTo-ApiValues"
)
