<#
    Excelファイル・出力先・読み込むシートを指定する実行ファイル。

    シートを追加するときは、$sheetDefinitionsへ設定を1件追加する。
    Excelの読込・JSON化・競合検知はExcelJson.psm1が担当する。
#>
param(
    [string]$ExcelPath = (Join-Path $PSScriptRoot "param_v7.0.xlsx"),
    [string]$OutputDir = (Join-Path $PSScriptRoot "output")
)

$ErrorActionPreference = "Stop"

# Layout:
#   Horizontal = 1行を1拠点として読む
#   Vertical   = ManagementNoごとの複数行を読む
#   KeyValue   = JSONPath列とValue列の固定値を縦方向に読む
#
# Verticalでは、OrderColumnの1始まりの番号を最初の [] の配列番号に使う。
#
# 横型シートを追加する例:
#   @{
#       Name = "拠点情報"
#       Layout = "Horizontal"
#       DisplayedTextPaths = @(
#           "deliveryDate",
#           "mobile.picDateOfBirth"
#       )
#   },
#
# 縦型シートを追加する例:
#   @{
#       Name = "パケットフィルタ"
#       Layout = "Vertical"
#       OrderColumn = "Inner_ManageNo"
#       DisplayedTextPaths = @(
#           "contracts[].startDate"
#       )
#   },
#
# Key-Value型シートを追加する例:
#   @{
#       Name = "KeyValueシート名"
#       Layout = "KeyValue"
#       JsonPathColumn = "JSONPath"
#       ValueColumn = "Value"
#       ※ApplyWhenAny を指定すると、条件に合致する場合のみ、JSONPathとValueを本文へ追加する
#       ApplyWhenAny = @(
#           @{ Path = "Key名"; Equals = "Value" },
#           @{ Path = "Key名2"; Equals = "Value" }
#       )
#   },
#
# ApplyWhenAnyを省略したKeyValueシートは、すべての拠点へ適用する。
# 条件判定に使うシートより、KeyValueシートを後ろへ記載する。
#
# 追加した各設定は、末尾の設定を除いて } の後ろにカンマが必要。
$sheetDefinitions = @(
    @{
        Name = "API"
        Layout = "Horizontal"
        DisplayedTextPaths = @(
            "deliveryDate"
        )
    },
    @{
        Name = "盾形"
        Layout = "Vertical"
        OrderColumn = "Inner_ManageNo"
    },
    @{
        Name = "モバイル情報"
        Layout = "KeyValue"
        JsonPathColumn = "JSONPath"
        ValueColumn = "Value"
        DisplayedTextPaths = @(
            "mobile.picDateOfBirth"
        )
        ApplyWhenAny = @(
            @{ Path = "primaryCircuitType"; Equals = "ワイヤレス" },
            @{ Path = "secondaryCircuitType"; Equals = "ワイヤレス" }
        )
    }
)

Import-Module (Join-Path $PSScriptRoot "modules\ExcelJson.psm1") -Force

$rulesPath = Join-Path $PSScriptRoot "rules\TerminalCreateRules.json"
$mappingsPath = Join-Path $PSScriptRoot "rules\TerminalValueMappings.json"

Convert-ExcelJsonWorkbook `
    -ExcelPath $ExcelPath `
    -OutputDir $OutputDir `
    -SheetDefinitions $sheetDefinitions `
    -RulesPath $rulesPath `
    -MappingsPath $mappingsPath
