<# JSON生成と検証を同じ引数体系で呼び出すCLIエントリーポイント。 #>
[CmdletBinding()]
param(
    [ValidateSet("Generate", "Validate", "All")]
    [string]$Mode = "All",
    [string]$ExcelPath = "",
    [string]$OutputDir = "",
    [string]$ReportPath = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $ExcelPath = Join-Path $PSScriptRoot "param_v7.0.xlsx"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $PSScriptRoot "output"
}
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path $PSScriptRoot "validation-report.html"
}

$mainScript = Join-Path $PSScriptRoot "main.ps1"
$validatorScript = Join-Path $PSScriptRoot "Validate-Json.ps1"

if ($Mode -in @("Generate", "All")) {
    if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) {
        throw "Excelファイルがありません: $ExcelPath"
    }

    $excelItem = Get-Item -LiteralPath $ExcelPath
    $lockFile = Join-Path $excelItem.DirectoryName ("~$" + $excelItem.Name)
    if (Test-Path -LiteralPath $lockFile) {
        Write-Warning "Excelが開かれています。未保存の変更はJSONへ反映されません: $($excelItem.Name)"
    }

    Write-Host "[開始] JSON生成"
    & $mainScript -ExcelPath $excelItem.FullName -OutputDir $OutputDir
    Write-Host "[完了] JSON生成"
}

if ($Mode -in @("Validate", "All")) {
    if (-not (Test-Path -LiteralPath $OutputDir -PathType Container)) {
        throw "JSON出力フォルダがありません: $OutputDir"
    }
    if (@(Get-ChildItem -LiteralPath $OutputDir -Filter "*.json" -File).Count -eq 0) {
        throw "検証対象のJSONがありません: $OutputDir"
    }

    Write-Host "[開始] JSON検証"
    & $validatorScript -InputDir $OutputDir -ReportPath $ReportPath
    Write-Host "[完了] JSON検証"
}

