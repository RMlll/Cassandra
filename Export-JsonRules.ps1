<#
    JsonCreateRules.xlsxを管理用の正本として読み込み、
    Validate-Json.ps1が使用するTerminalCreateRules.jsonを書き出す。
#>
[CmdletBinding()]
param(
    [string]$ExcelPath = "",
    [string]$OutputPath = "",
    [string]$MappingsOutputPath = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $ExcelPath = Join-Path $PSScriptRoot "rules\JsonCreateRules.xlsx"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot "rules\TerminalCreateRules.json"
}
if ([string]::IsNullOrWhiteSpace($MappingsOutputPath)) {
    $MappingsOutputPath = Join-Path $PSScriptRoot "rules\TerminalValueMappings.json"
}

function Remove-ComReference([object]$ComObject) {
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Get-WorksheetMatrix([object]$Worksheets, [string]$Name) {
    $sheet = $null
    $range = $null
    try {
        $sheet = $Worksheets.Item($Name)
        $range = $sheet.UsedRange
        return [pscustomobject]@{
            Values = $range.Value2
            Rows = [int]$range.Rows.Count
            Columns = [int]$range.Columns.Count
        }
    }
    catch {
        throw "ルール管理Excelに '$Name' シートがありません。"
    }
    finally {
        Remove-ComReference $range
        Remove-ComReference $sheet
    }
}

function Get-MatrixValue([object]$Matrix, [int]$Row, [int]$Column) {
    if ($null -eq $Matrix) { return $null }
    if ($Matrix -is [object[,]]) { return $Matrix[$Row, $Column] }
    if ($Row -eq 1 -and $Column -eq 1) { return $Matrix }
    return $null
}

function Find-HeaderRow([object]$SheetData, [string]$FirstHeader) {
    for ($row = 1; $row -le [Math]::Min(20, $SheetData.Rows); $row++) {
        if (([string](Get-MatrixValue $SheetData.Values $row 1)).Trim() -eq $FirstHeader) {
            return $row
        }
    }
    throw "ヘッダー '$FirstHeader' が見つかりません。"
}

function Get-HeaderColumns([object]$SheetData, [int]$HeaderRow) {
    $columns = @{}
    for ($column = 1; $column -le $SheetData.Columns; $column++) {
        $name = ([string](Get-MatrixValue $SheetData.Values $HeaderRow $column)).Trim()
        if ($name -ne "") { $columns[$name] = $column }
    }
    return $columns
}

function Get-OptionalInteger([object]$Value, [string]$ColumnName, [string]$Path) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $number = 0
    if (-not [int]::TryParse(([string]$Value).Trim(), [ref]$number)) {
        throw "$Path の $ColumnName は整数で指定してください: $Value"
    }
    return $number
}

function ConvertTo-Boolean([object]$Value) {
    $text = ([string]$Value).Trim()
    return $text -in @("TRUE", "True", "true", "1")
}

function Write-JsonAtomically([object]$Value, [string]$Path) {
    $json = $Value | ConvertTo-Json -Depth 12
    $outputDirectory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    $temporaryPath = Join-Path $outputDirectory (".{0}.{1}.tmp" -f [IO.Path]::GetFileName($Path), [guid]::NewGuid().ToString("N"))
    try {
        [IO.File]::WriteAllText($temporaryPath, $json, (New-Object Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
    }
}

if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) {
    throw "ルール管理Excelがありません: $ExcelPath"
}

$excel = $null
$workbooks = $null
$book = $null
$worksheets = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    $book = $workbooks.Open((Resolve-Path -LiteralPath $ExcelPath).Path, 0, $true)
    $worksheets = $book.Worksheets

    $guide = Get-WorksheetMatrix $worksheets "Guide"
    $rulesData = Get-WorksheetMatrix $worksheets "Rules"
    $enumData = Get-WorksheetMatrix $worksheets "Enum Values"
    $mappingData = Get-WorksheetMatrix $worksheets "Value Mappings"

    $metadata = @{}
    for ($row = 1; $row -le $guide.Rows; $row++) {
        $label = ([string](Get-MatrixValue $guide.Values $row 1)).Trim()
        if ($label -ne "") {
            $metadata[$label] = Get-MatrixValue $guide.Values $row 2
        }
    }

    $enumByPath = @{}
    $enumHeaderRow = Find-HeaderRow $enumData "JSONPath"
    $enumColumns = Get-HeaderColumns $enumData $enumHeaderRow
    foreach ($requiredHeader in @("JSONPath", "Enum Value")) {
        if (-not $enumColumns.ContainsKey($requiredHeader)) {
            throw "Enum Valuesシートに '$requiredHeader' 列がありません。"
        }
    }
    for ($row = $enumHeaderRow + 1; $row -le $enumData.Rows; $row++) {
        $path = ([string](Get-MatrixValue $enumData.Values $row $enumColumns["JSONPath"])).Trim()
        $value = ([string](Get-MatrixValue $enumData.Values $row $enumColumns["Enum Value"])).Trim()
        if ($path -eq "" -and $value -eq "") { continue }
        if ($path -eq "" -or $value -eq "") {
            throw "Enum Valuesシートの$row行目はJSONPathまたはEnum Valueが空です。"
        }
        if (-not $enumByPath.ContainsKey($path)) {
            $enumByPath[$path] = New-Object System.Collections.ArrayList
        }
        if ($value -in $enumByPath[$path]) {
            throw "Enum Valuesに重複があります: $path = $value"
        }
        [void]$enumByPath[$path].Add($value)
    }

    $rulesHeaderRow = Find-HeaderRow $rulesData "No"
    $ruleColumns = Get-HeaderColumns $rulesData $rulesHeaderRow
    $requiredColumns = @(
        "Enabled", "JSONPath", "Type", "Required", "Format", "Pattern",
        "MinLength", "MaxLength", "MinItems", "MaxItems"
    )
    foreach ($requiredHeader in $requiredColumns) {
        if (-not $ruleColumns.ContainsKey($requiredHeader)) {
            throw "Rulesシートに '$requiredHeader' 列がありません。"
        }
    }

    $rules = New-Object System.Collections.ArrayList
    $seenPaths = @{}
    $ruleByPath = @{}
    for ($row = $rulesHeaderRow + 1; $row -le $rulesData.Rows; $row++) {
        $path = ([string](Get-MatrixValue $rulesData.Values $row $ruleColumns["JSONPath"])).Trim()
        if ($path -eq "") { continue }
        $enabledValue = Get-MatrixValue $rulesData.Values $row $ruleColumns["Enabled"]
        if (-not (ConvertTo-Boolean $enabledValue)) { continue }
        if ($seenPaths.ContainsKey($path)) { throw "RulesにJSONPathの重複があります: $path" }
        $seenPaths[$path] = $true

        $enumValues = @()
        if ($enumByPath.ContainsKey($path)) { $enumValues = @($enumByPath[$path]) }

        $rule = [ordered]@{
            path = $path
            type = ([string](Get-MatrixValue $rulesData.Values $row $ruleColumns["Type"])).Trim()
            required = ConvertTo-Boolean (Get-MatrixValue $rulesData.Values $row $ruleColumns["Required"])
            format = ([string](Get-MatrixValue $rulesData.Values $row $ruleColumns["Format"])).Trim()
            pattern = ([string](Get-MatrixValue $rulesData.Values $row $ruleColumns["Pattern"])).Trim()
            enum = $enumValues
            minLength = Get-OptionalInteger (Get-MatrixValue $rulesData.Values $row $ruleColumns["MinLength"]) "MinLength" $path
            maxLength = Get-OptionalInteger (Get-MatrixValue $rulesData.Values $row $ruleColumns["MaxLength"]) "MaxLength" $path
            minItems = Get-OptionalInteger (Get-MatrixValue $rulesData.Values $row $ruleColumns["MinItems"]) "MinItems" $path
            maxItems = Get-OptionalInteger (Get-MatrixValue $rulesData.Values $row $ruleColumns["MaxItems"]) "MaxItems" $path
        }
        foreach ($key in @("type", "format", "pattern")) {
            if ([string]::IsNullOrWhiteSpace([string]$rule[$key])) { $rule[$key] = $null }
        }
        $ruleObject = [pscustomobject]$rule
        [void]$rules.Add($ruleObject)
        $ruleByPath[$path] = $ruleObject
    }

    foreach ($enumPath in $enumByPath.Keys) {
        if (-not $seenPaths.ContainsKey($enumPath)) {
            throw "Enum ValuesのJSONPathがRulesに存在しないか、無効化されています: $enumPath"
        }
    }

    $mappingHeaderRow = Find-HeaderRow $mappingData "Enabled"
    $mappingColumns = Get-HeaderColumns $mappingData $mappingHeaderRow
    foreach ($requiredHeader in @("Enabled", "JSONPath", "Excel Value", "API Value", "Action")) {
        if (-not $mappingColumns.ContainsKey($requiredHeader)) {
            throw "Value Mappingsシートに '$requiredHeader' 列がありません。"
        }
    }

    $mappings = New-Object System.Collections.ArrayList
    $seenMappings = @{}
    for ($row = $mappingHeaderRow + 1; $row -le $mappingData.Rows; $row++) {
        $path = ([string](Get-MatrixValue $mappingData.Values $row $mappingColumns["JSONPath"])).Trim()
        $inputValue = ([string](Get-MatrixValue $mappingData.Values $row $mappingColumns["Excel Value"])).Trim()
        $apiValue = ([string](Get-MatrixValue $mappingData.Values $row $mappingColumns["API Value"])).Trim()
        $action = ([string](Get-MatrixValue $mappingData.Values $row $mappingColumns["Action"])).Trim().ToLower()
        if ($path -eq "" -and $inputValue -eq "" -and $apiValue -eq "" -and $action -eq "") { continue }
        if (-not (ConvertTo-Boolean (Get-MatrixValue $mappingData.Values $row $mappingColumns["Enabled"]))) { continue }
        if ($action -notin @("set", "omit")) {
            throw "Value Mappingsシートの${row}行目のActionはSetまたはOmitで指定してください: $action"
        }
        if ($path -eq "" -or $inputValue -eq "" -or ($action -eq "set" -and $apiValue -eq "")) {
            throw "Value Mappingsシートの$row行目はJSONPath、Excel Value、またはSet時のAPI Valueが空です。"
        }
        if (-not $ruleByPath.ContainsKey($path)) {
            throw "Value MappingsのJSONPathがRulesに存在しないか、無効化されています: $path"
        }

        $mappingKey = $path + [char]0 + $inputValue
        if ($seenMappings.ContainsKey($mappingKey)) {
            throw "Value Mappingsに重複があります: $path / $inputValue"
        }
        $seenMappings[$mappingKey] = $true

        $rule = $ruleByPath[$path]
        if ($action -eq "set" -and [string]$rule.type -eq "boolean" -and $apiValue -notin @("true", "false")) {
            throw "$path のAPI Valueはbooleanのためtrueまたはfalseで指定してください: $apiValue"
        }
        if ($action -eq "set" -and [string]$rule.type -eq "integer") {
            [long]$integerValue = 0
            if (-not [int64]::TryParse($apiValue, [ref]$integerValue)) {
                throw "$path のAPI Valueはintegerとして解釈できません: $apiValue"
            }
        }
        if ($action -eq "set" -and [string]$rule.type -eq "number") {
            $numberValue = 0.0
            if (-not [double]::TryParse($apiValue, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$numberValue)) {
                throw "$path のAPI Valueはnumberとして解釈できません: $apiValue"
            }
        }
        $enumValues = @($rule.enum | Where-Object { $null -ne $_ -and [string]$_ -ne "" })
        if ($action -eq "set" -and $enumValues.Count -gt 0 -and $apiValue -notin $enumValues) {
            throw "$path のAPI Valueがenum候補にありません: $apiValue / 候補: $($enumValues -join ', ')"
        }

        [void]$mappings.Add([pscustomobject][ordered]@{
            path = $path
            inputValue = $inputValue
            apiValue = $apiValue
            action = $action
        })
    }

    $result = [ordered]@{
        sourceUrl = [string]$metadata["参照URL"]
        retrievedDate = [string]$metadata["取得日"]
        operationId = [string]$metadata["Operation ID"]
        rootSchema = [string]$metadata["Root Schema"]
        rules = @($rules)
    }

    $mappingResult = [ordered]@{
        operationId = [string]$metadata["Operation ID"]
        rootSchema = [string]$metadata["Root Schema"]
        mappings = @($mappings)
    }

    Write-JsonAtomically $result $OutputPath
    Write-JsonAtomically $mappingResult $MappingsOutputPath

    Write-Host "[OK] ルールJSON: $OutputPath"
    Write-Host "     Rules: $($rules.Count) / Enum values: $($enumByPath.Values | ForEach-Object Count | Measure-Object -Sum | Select-Object -ExpandProperty Sum)"
    Write-Host "[OK] マッピングJSON: $MappingsOutputPath"
    Write-Host "     Mappings: $($mappings.Count)"
}
finally {
    if ($book) { try { $book.Close($false) } catch { Write-Verbose $_.Exception.Message } }
    Remove-ComReference $worksheets
    Remove-ComReference $book
    Remove-ComReference $workbooks
    if ($excel) { try { $excel.Quit() } catch { Write-Verbose $_.Exception.Message } }
    Remove-ComReference $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
