<#
    ExcelのAPI用一覧シートを読み込み、拠点別JSONを出力する公開モジュール。
    Excel読取・JSON構築・値変換の詳細は同じmodules配下へ委譲する。
#>

Import-Module (Join-Path $PSScriptRoot "ExcelReader.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "JsonBuilder.psm1") -Force
Import-Module (Join-Path $PSScriptRoot "ValueMapper.psm1") -Force

# 全縦型シート共通の除外列
$script:ControlPaths = @(
    "ManagementNo", "管理No", "内部管理No", "Inner_ManageNo"
)

function Convert-ExcelJsonWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExcelPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDir,

        [Parameter(Mandatory = $true)]
        [object[]]$SheetDefinitions,

        [string]$RulesPath = "",

        [string]$MappingsPath = ""
    )

    if ([string]::IsNullOrWhiteSpace($RulesPath)) {
        $RulesPath = Join-Path (Split-Path -Parent $PSScriptRoot) "rules\TerminalCreateRules.json"
    }
    if ([string]::IsNullOrWhiteSpace($MappingsPath)) {
        $MappingsPath = Join-Path (Split-Path -Parent $PSScriptRoot) "rules\TerminalValueMappings.json"
    }

    $mappingConfiguration = Import-ValueMappingConfiguration `
        -RulesPath $RulesPath `
        -MappingsPath $MappingsPath

    if (-not (Test-Path -LiteralPath $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir | Out-Null
    }

    $excel = $null
    $workbooks = $null
    $book = $null
    $worksheets = $null
    $worksheet = $null
    $bodiesBySite = @{}
    $ownersBySite = @{}

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbooks = $excel.Workbooks
        $book = $workbooks.Open((Resolve-Path -LiteralPath $ExcelPath).Path, 0, $true)
        $worksheets = $book.Worksheets

        $availableSheetNames = @()
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $sheetForName = $null
            try {
                $sheetForName = $worksheets.Item($sheetIndex)
                $availableSheetNames += [string]$sheetForName.Name
            }
            finally {
                Remove-ExcelComReference $sheetForName
            }
        }

        foreach ($definition in $SheetDefinitions) {
            $sheetName = [string]$definition.Name
            $layout = ([string]$definition.Layout).Trim().ToLower()

            if ($sheetName -eq "") { throw "SheetDefinitionsのNameが空です。" }
            if ($layout -notin @("horizontal", "vertical", "keyvalue")) {
                throw "$sheetName のLayoutはHorizontal、Vertical、KeyValueのいずれかで指定してください。"
            }
            if ($sheetName -notin $availableSheetNames) {
                throw "Excelにシート '$sheetName' がありません。存在するシート: $($availableSheetNames -join ', ')"
            }

            try {
                $worksheet = $worksheets.Item($sheetName)
                $displayedTextPaths = @($definition.DisplayedTextPaths)

                if ($layout -eq "horizontal") {
                    $sheetAssignments = Read-HorizontalSheet `
                        -Worksheet $worksheet `
                        -WorksheetName $sheetName `
                        -DisplayedTextPaths $displayedTextPaths
                }
                elseif ($layout -eq "vertical") {
                    $orderColumn = [string]$definition.OrderColumn
                    if ($orderColumn -eq "") { $orderColumn = "Inner_ManageNo" }

                    $excludedPaths = @($script:ControlPaths)
                    if ($null -ne $definition.ExcludedPaths) {
                        $excludedPaths += @($definition.ExcludedPaths)
                    }

                    $sheetAssignments = Read-VerticalSheet `
                        -Worksheet $worksheet `
                        -WorksheetName $sheetName `
                        -OrderColumn $orderColumn `
                        -ExcludedPaths $excludedPaths `
                        -DisplayedTextPaths $displayedTextPaths

                    $sheetAssignments = Expand-AllSiteAssignments `
                        -SheetAssignments $sheetAssignments `
                        -BodiesBySite $bodiesBySite `
                        -WorksheetName $sheetName
                }
                else {
                    $jsonPathColumn = [string]$definition.JsonPathColumn
                    if ($jsonPathColumn -eq "") { $jsonPathColumn = "JSONPath" }

                    $valueColumn = [string]$definition.ValueColumn
                    if ($valueColumn -eq "") { $valueColumn = "Value" }

                    $keyValueAssignments = Read-KeyValueSheet `
                        -Worksheet $worksheet `
                        -WorksheetName $sheetName `
                        -JsonPathColumnName $jsonPathColumn `
                        -ValueColumnName $valueColumn `
                        -DisplayedTextPaths $displayedTextPaths

                    $sheetAssignments = Convert-KeyValueAssignmentsToSites `
                        -Assignments $keyValueAssignments `
                        -BodiesBySite $bodiesBySite `
                        -ApplyWhenAny @($definition.ApplyWhenAny) `
                        -WorksheetName $sheetName
                }

                Add-SheetAssignments `
                    -BodiesBySite $bodiesBySite `
                    -OwnersBySite $ownersBySite `
                    -SheetAssignments $sheetAssignments
            }
            finally {
                Remove-ExcelComReference $worksheet
                $worksheet = $null
            }
        }

        $siteNumbers = @($bodiesBySite.Keys | Sort-Object { [int]$_ })
        foreach ($siteNo in $siteNumbers) {
            [void](ConvertTo-ApiValues `
                -Body $bodiesBySite[$siteNo] `
                -Configuration $mappingConfiguration)
            Assert-NoArrayHoles $bodiesBySite[$siteNo]

            $json = ConvertTo-IndentedJson $bodiesBySite[$siteNo]
            $outputPath = Join-Path $OutputDir ("terminal_{0}.json" -f $siteNo)
            [IO.File]::WriteAllText(
                $outputPath,
                $json,
                (New-Object Text.UTF8Encoding($false))
            )
            Write-Host "[OK] ManagementNo $siteNo -> $outputPath"
        }

        if ($siteNumbers.Count -eq 0) {
            Write-Warning "JSON出力対象のManagementNoがありません。"
        }
    }
    finally {
        Remove-ExcelComReference $worksheet

        if ($book) {
            try { $book.Close($false) } catch { Write-Verbose $_.Exception.Message }
        }

        Remove-ExcelComReference $worksheets
        Remove-ExcelComReference $book
        Remove-ExcelComReference $workbooks

        if ($excel) {
            try { $excel.Quit() } catch { Write-Verbose $_.Exception.Message }
        }
        Remove-ExcelComReference $excel

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Export-ModuleMember -Function Convert-ExcelJsonWorkbook
