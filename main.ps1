<#
    Excelファイル・出力先・読み込むシートを指定する実行ファイル。

    シートを追加するときは、$sheetDefinitionsへ設定を1件追加する。
    Excelの読込・JSON化・競合検知はExcelJson.psm1が担当する。
#>
param(
    [string]$ExcelPath = (Join-Path $PSScriptRoot "param_v7.0.xlsx"),
    [string]$OutputDir = (Join-Path $PSScriptRoot "output"),
    [string]$SiteSelection = ""
)

$ErrorActionPreference = "Stop"

# SiteSelection:
#   空欄またはALL = 全拠点
#   1             = ManagementNo 1のみ
#   1-10, 50-51   = 範囲と複数指定（重複は自動で除外）
#
# Layout:
#   Horizontal = 1行を1拠点として読む
#   Vertical   = ManagementNoごとの複数行を読む
#   KeyValue   = JSONPath列とValue列の固定値を縦方向に読む
#
# Verticalでは、OrderColumnの1始まりの番号を最初の [] の配列番号に使う。
# ManagementNoへALLと記載した行は、生成対象の全ManagementNoへ適用する。
# ALLを使用するVerticalシートは、ManagementNoを作るシートより後ろへ記載する。
# 同じVerticalシート内に個別ManagementNoがある拠点は、ALLを使わず個別行だけを使う。
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
        Name = ""
        Layout = "Horizontal"
        DisplayedTextPaths = @(
            "deliveryDate"
        )
    },
    @{
        Name = ""
        Layout = "Vertical"
        OrderColumn = "Inner_ManageNo"
    },
    @{
        Name = ""
        Layout = "Vertical"
        OrderColumn = "Inner_ManageNo"
    },
    @{
        Name = ""
        Layout = "KeyValue"
        JsonPathColumn = "JSONPath"
        ValueColumn = "Value"
        DisplayedTextPaths = @(
            "mobile.picDateOfBirth"
        )
        ApplyWhenAny = @(
            @{ Path = ""; Equals = "" },
            @{ Path = ""; Equals = "" }
        )
    }
)

Import-Module (Join-Path $PSScriptRoot "modules\ExcelJson.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "modules\SiteSelection.psm1") -Force

$rulesPath = Join-Path $PSScriptRoot "rules\TerminalCreateRules.json"
$mappingsPath = Join-Path $PSScriptRoot "rules\TerminalValueMappings.json"
$targetSiteNumbers = @(ConvertFrom-SiteSelection -Selection $SiteSelection)

Convert-ExcelJsonWorkbook `
    -ExcelPath $ExcelPath `
    -OutputDir $OutputDir `
    -SheetDefinitions $sheetDefinitions `
    -TargetSiteNumbers $targetSiteNumbers `
    -RulesPath $rulesPath `
    -MappingsPath $mappingsPath
