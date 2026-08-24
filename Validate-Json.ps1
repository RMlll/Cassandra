<# 生成済みJSONを検証し、HTMLレポートを出力する。 #>
[CmdletBinding()]
param(
    [string]$InputDir = "",
    [string]$ReportPath = "",
    [string]$RulesPath = "",
    [string]$SiteSelection = "",
    [switch]$FailOnError
)

$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "modules\SiteSelection.psm1") -Force
if ([string]::IsNullOrWhiteSpace($InputDir)) {
    $InputDir = Join-Path $PSScriptRoot "output"
}
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path $PSScriptRoot "validation-report.html"
}
if ([string]::IsNullOrWhiteSpace($RulesPath)) {
    $RulesPath = Join-Path $PSScriptRoot "rules\TerminalCreateRules.json"
}
$issues = New-Object System.Collections.ArrayList
$fileResults = New-Object System.Collections.ArrayList

function ConvertTo-HtmlEncoded([object]$value) {
    return [Net.WebUtility]::HtmlEncode([string]$value)
}

function Add-Issue(
    [string]$file,
    [string]$severity,
    [string]$path,
    [object]$value,
    [string]$message,
    [string]$rule
) {
    $text = [string]$value
    if ($text.Length -gt 160) { $text = $text.Substring(0, 157) + "..." }
    [void]$issues.Add([pscustomobject]@{
        File=$file; Severity=$severity; Path=$path; Value=$text
        Message=$message; Rule=$rule
    })
}

function Get-Property([object]$object, [string]$name) {
    if ($null -eq $object) {
        return [pscustomobject]@{ Exists=$false; Value=$null }
    }
    if ($object -is [System.Collections.IDictionary]) {
        $exists = $object.Contains($name)
        $value = $null
        if ($exists) { $value = $object[$name] }
        return [pscustomobject]@{
            Exists=$exists
            Value=$value
        }
    }
    $property = $object.PSObject.Properties[$name]
    $value = $null
    if ($null -ne $property) { $value = $property.Value }
    return [pscustomobject]@{
        Exists=($null -ne $property)
        Value=$value
    }
}

function Get-Nodes([object]$root, [string]$path) {
    $nodes = @([pscustomobject]@{ Value=$root; ActualPath='$' })
    if ([string]::IsNullOrWhiteSpace($path)) { return $nodes }

    foreach ($segmentText in $path.Split('.')) {
        $isArray = $segmentText.EndsWith('[]')
        $name = if ($isArray) {
            $segmentText.Substring(0, $segmentText.Length - 2)
        } else {
            $segmentText
        }

        $next = New-Object System.Collections.ArrayList
        foreach ($node in $nodes) {
            $property = Get-Property $node.Value $name
            if (-not $property.Exists) { continue }
            $actual = $node.ActualPath + "." + $name

            if ($isArray) {
                if ($property.Value -is [System.Collections.IList] -and $property.Value -isnot [string]) {
                    for ($index=0; $index -lt $property.Value.Count; $index++) {
                        [void]$next.Add([pscustomobject]@{
                            Value=$property.Value[$index]
                            ActualPath=$actual + "[$index]"
                        })
                    }
                }
            } else {
                [void]$next.Add([pscustomobject]@{
                    Value=$property.Value; ActualPath=$actual
                })
            }
        }
        $nodes = @($next)
    }
    return $nodes
}

function Get-ValueType([object]$value) {
    if ($null -eq $value) { return "null" }
    if ($value -is [bool]) { return "boolean" }
    if ($value -is [string]) { return "string" }
    if ($value -is [System.Collections.IList]) { return "array" }
    if ($value -is [System.Collections.IDictionary] -or $value -is [pscustomobject]) { return "object" }
    if ($value -is [byte] -or $value -is [int16] -or $value -is [int32] -or $value -is [int64]) { return "integer" }
    if ($value -is [single] -or $value -is [double] -or $value -is [decimal]) { return "number" }
    return $value.GetType().Name
}

function Test-Value([object]$rule, [object]$value, [string]$file, [string]$path) {
    $actualType = Get-ValueType $value
    $expectedType = [string]$rule.type

    if ($actualType -eq "null") {
        Add-Issue $file "Error" $path $value "nullは許可されていません。" ("type=" + $expectedType)
        return
    }

    $matches = if ($expectedType -eq "number") {
        $actualType -in @("number","integer")
    } elseif ([string]::IsNullOrEmpty($expectedType)) {
        $true
    } else {
        $actualType -eq $expectedType
    }

    if (-not $matches) {
        Add-Issue $file "Error" $path $value ("型が一致しません。期待: " + $expectedType + " / 実際: " + $actualType) "OpenAPI type"
        return
    }

    if ($expectedType -eq "string") {
        $text = [string]$value
        if ($null -ne $rule.minLength -and $text.Length -lt [int]$rule.minLength) {
            Add-Issue $file "Error" $path $text "文字数が最小値を下回っています。" ("minLength=" + $rule.minLength)
        }
        if ($null -ne $rule.maxLength -and $text.Length -gt [int]$rule.maxLength) {
            Add-Issue $file "Error" $path $text "文字数が最大値を超えています。" ("maxLength=" + $rule.maxLength)
        }
        if (-not [string]::IsNullOrEmpty([string]$rule.pattern)) {
            try {
                if (-not [regex]::IsMatch($text, [string]$rule.pattern)) {
                    Add-Issue $file "Error" $path $text "正規表現に一致しません。" ([string]$rule.pattern)
                }
            } catch {
                Add-Issue $file "Warning" $path $text "検証用正規表現を評価できませんでした。" ([string]$rule.pattern)
            }
        }
        if ([string]$rule.format -eq "date") {
            $parsed = [DateTime]::MinValue
            $valid = [DateTime]::TryParseExact($text, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)
            if (-not $valid) {
                Add-Issue $file "Error" $path $text "日付はyyyy-MM-dd形式で指定してください。" "OpenAPI format=date"
            }
        }
        # OpenAPIにenum指定がないルールは、スナップショット上で[null]になる場合がある。
        # nullや空文字を候補として扱うと、すべての通常文字列がenum違反になるため除外する。
        $enum = @($rule.enum | Where-Object { $null -ne $_ -and -not [string]::IsNullOrEmpty([string]$_) })
        if ($enum.Count -gt 0 -and $text -notin $enum) {
            Add-Issue $file "Error" $path $text "列挙値に含まれていません。" ("enum=" + ($enum -join ", "))
        }
    }

    if ($expectedType -eq "array") {
        if ($null -ne $rule.minItems -and $value.Count -lt [int]$rule.minItems) {
            Add-Issue $file "Error" $path $value.Count "配列要素数が最小値を下回っています。" ("minItems=" + $rule.minItems)
        }
        if ($null -ne $rule.maxItems -and $value.Count -gt [int]$rule.maxItems) {
            Add-Issue $file "Error" $path $value.Count "配列要素数が最大値を超えています。" ("maxItems=" + $rule.maxItems)
        }
    }
}

function Test-Schema([object]$root, [object[]]$rules, [string]$file) {
    foreach ($rule in $rules) {
        $dot = $rule.path.LastIndexOf('.')
        $parentPath = if ($dot -ge 0) { $rule.path.Substring(0,$dot) } else { "" }
        $propertyName = if ($dot -ge 0) { $rule.path.Substring($dot+1) } else { $rule.path }

        foreach ($parent in @(Get-Nodes $root $parentPath)) {
            $property = Get-Property $parent.Value $propertyName
            $actualPath = $parent.ActualPath + "." + $propertyName
            if (-not $property.Exists) {
                if ([bool]$rule.required) {
                    Add-Issue $file "Error" $actualPath $null "必須項目がありません。" "OpenAPI required"
                }
                continue
            }
            Test-Value $rule $property.Value $file $actualPath
        }
    }
}

function Inspect-Properties(
    [object]$value,
    [string]$normalized,
    [string]$actual,
    [hashtable]$allowed,
    [string]$file
) {
    if ($null -eq $value) { return }

    if ($value -is [System.Collections.IList] -and $value -isnot [string]) {
        for ($index=0; $index -lt $value.Count; $index++) {
            Inspect-Properties $value[$index] ($normalized+"[]") ($actual+"[$index]") $allowed $file
        }
        return
    }

    if ($value -isnot [pscustomobject] -and $value -isnot [System.Collections.IDictionary]) { return }

    $properties = if ($value -is [System.Collections.IDictionary]) {
        @($value.Keys | ForEach-Object { [pscustomobject]@{ Name=[string]$_; Value=$value[$_] } })
    } else {
        @($value.PSObject.Properties)
    }

    foreach ($property in $properties) {
        $normalPath = if ($normalized) { $normalized+"."+$property.Name } else { $property.Name }
        $actualPath = $actual+"."+$property.Name

        if (-not $allowed.ContainsKey($normalPath)) {
            Add-Issue $file "Warning" $actualPath $property.Value "APIスキーマにない項目です。" "TerminalCreate"
        }

        if ($property.Name -match 'Address' -and $property.Value -is [string] -and $property.Value.Contains(' ')) {
            Add-Issue $file "Warning" $actualPath $property.Value "住所に半角スペースが含まれています。" "運用ルール"
        }

        if ($normalPath -in @("mobile.picName","mobile.picNameKana") -and $property.Value -is [string]) {
            if ($property.Value.Contains(' ')) {
                Add-Issue $file "Error" $actualPath $property.Value "姓と名の区切りに半角スペースが使われています。" "全角スペースを使用"
            }
            if (-not $property.Value.Contains([string][char]0x3000)) {
                Add-Issue $file "Error" $actualPath $property.Value "姓と名の間に全角スペースが必要です。" "OpenAPI API項目説明"
            }
        }

        Inspect-Properties $property.Value $normalPath $actualPath $allowed $file
    }
}

if (-not (Test-Path -LiteralPath $RulesPath -PathType Leaf)) { throw "検証ルールがありません: $RulesPath" }
if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) { throw "JSON入力フォルダがありません: $InputDir" }

$ruleDocument = Get-Content -LiteralPath $RulesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rules = @($ruleDocument.rules)
$allowed = @{}
foreach ($rule in $rules) { $allowed[[string]$rule.path] = $true }
$targetSiteNumbers = @(ConvertFrom-SiteSelection -Selection $SiteSelection)
if ($targetSiteNumbers.Count -gt 0) {
    $jsonFiles = @(
        foreach ($siteNo in $targetSiteNumbers) {
            $filePath = Join-Path $InputDir ("terminal_{0}.json" -f $siteNo)
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                throw "対象拠点のJSONがありません: $filePath"
            }
            Get-Item -LiteralPath $filePath
        }
    )
}
else {
    $jsonFiles = @(Get-ChildItem -LiteralPath $InputDir -Filter "*.json" -File | Sort-Object Name)
}

foreach ($jsonFile in $jsonFiles) {
    $before = $issues.Count
    try {
        $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        Test-Schema $json $rules $jsonFile.Name
        Inspect-Properties $json "" '$' $allowed $jsonFile.Name
    } catch {
        Add-Issue $jsonFile.Name "Error" '$' $null ("JSONを読み込めません: "+$_.Exception.Message) "JSON syntax"
    }

    $added = @($issues | Select-Object -Skip $before)
    [void]$fileResults.Add([pscustomobject]@{
        File=$jsonFile.Name
        Errors=@($added | Where-Object Severity -eq "Error").Count
        Warnings=@($added | Where-Object Severity -eq "Warning").Count
    })
}

$errorCount = @($issues | Where-Object Severity -eq "Error").Count
$warningCount = @($issues | Where-Object Severity -eq "Warning").Count
$fileRows = foreach ($result in $fileResults) {
    $class = if ($result.Errors) { "error" } elseif ($result.Warnings) { "warning" } else { "ok" }
    $label = if ($result.Errors) { "エラー" } elseif ($result.Warnings) { "警告" } else { "OK" }
    "<tr><td>{0}</td><td><span class='badge {1}'>{2}</span></td><td>{3}</td><td>{4}</td></tr>" -f (ConvertTo-HtmlEncoded $result.File),$class,$label,$result.Errors,$result.Warnings
}
$issueRows = foreach ($issue in $issues) {
    "<tr><td><span class='badge {0}'>{1}</span></td><td>{2}</td><td><code>{3}</code></td><td>{4}</td><td>{5}</td><td><code>{6}</code></td></tr>" -f $issue.Severity.ToLowerInvariant(),(ConvertTo-HtmlEncoded $issue.Severity),(ConvertTo-HtmlEncoded $issue.File),(ConvertTo-HtmlEncoded $issue.Path),(ConvertTo-HtmlEncoded $issue.Value),(ConvertTo-HtmlEncoded $issue.Message),(ConvertTo-HtmlEncoded $issue.Rule)
}
if (@($fileRows).Count -eq 0) { $fileRows = "<tr><td colspan='4'>JSONファイルがありません。</td></tr>" }
if (@($issueRows).Count -eq 0) { $issueRows = "<tr><td colspan='6'>問題は見つかりませんでした。</td></tr>" }


$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$html = @"
<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cassandra JSON検証レポート</title>
<style>
:root{--bg:#f4f7fb;--card:#fff;--text:#172033;--line:#dce3ee;--red:#b42318;--redbg:#fee4e2;--amber:#b54708;--amberbg:#fef0c7;--green:#067647;--greenbg:#d1fadf}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:"Segoe UI","Yu Gothic UI",sans-serif}
main{max-width:1400px;margin:auto;padding:32px 24px}h1{margin:0}.meta{color:#657087;margin:8px 0 24px}
.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.card,section{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin-bottom:20px}
.number{font-size:30px;font-weight:700}.label{color:#657087}table{width:100%;border-collapse:collapse;font-size:14px}th,td{padding:11px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}th{background:#f8fafc}code{overflow-wrap:anywhere}
.badge{padding:3px 9px;border-radius:99px;font-weight:700;font-size:12px}.error{color:var(--red);background:var(--redbg)}.warning{color:var(--amber);background:var(--amberbg)}.ok{color:var(--green);background:var(--greenbg)}
.note{padding:14px;background:#eef4ff;border-left:4px solid #3b82f6}@media(max-width:700px){.cards{grid-template-columns:1fr}}
</style></head><body><main>
<h1>Cassandra JSON検証レポート</h1><div class="meta">生成日時: $generatedAt</div>
<div class="cards"><div class="card"><div class="number">$($jsonFiles.Count)</div><div class="label">JSONファイル</div></div><div class="card"><div class="number">$errorCount</div><div class="label">エラー</div></div><div class="card"><div class="number">$warningCount</div><div class="label">警告</div></div></div>
<section><h2>ファイル別結果</h2><table><thead><tr><th>ファイル</th><th>結果</th><th>エラー</th><th>警告</th></tr></thead><tbody>$(@($fileRows)-join[Environment]::NewLine)</tbody></table></section>
<section><h2>検出内容</h2><table><thead><tr><th>重要度</th><th>ファイル</th><th>JSONPath</th><th>値</th><th>内容</th><th>ルール</th></tr></thead><tbody>$(@($issueRows)-join[Environment]::NewLine)</tbody></table></section>
</main></body></html>
"@

$fullReportPath = [IO.Path]::GetFullPath($ReportPath)
$reportDir = Split-Path -Parent $fullReportPath
if (-not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Path $reportDir | Out-Null }
[IO.File]::WriteAllText($fullReportPath,$html,(New-Object Text.UTF8Encoding($false)))
Write-Host ("[OK] 検証レポート: "+$fullReportPath)
Write-Host ("     JSON: {0} / エラー: {1} / 警告: {2}" -f $jsonFiles.Count,$errorCount,$warningCount)
if ($FailOnError -and $errorCount -gt 0) { exit 1 }
