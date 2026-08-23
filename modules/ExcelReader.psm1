<#
    Excel COMからAPI用一覧シートを読み取り、Assignmentへ変換する内部モジュール。
    シート全体はValue2で一括取得し、DisplayedTextPathsだけ.Textで取得する。
#>

function Remove-ExcelComReference([object]$ComObject) {
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-ExcelCellText(
    [object]$Cells,
    [int]$Row,
    [int]$Column,
    [string]$WorksheetName
) {
    $cell = $null
    try {
        $cell = $Cells.Item($Row,$Column)
        $text = [string]$cell.Text
        if ($text -match '^#+$') {
            throw ("{0}シート {1}行 {2}列: 列幅不足で表示値が'{3}'です。" -f $WorksheetName,$Row,$Column,$text)
        }
        return $text
    } finally {
        Remove-ExcelComReference $cell
    }
}

function Get-WorksheetData([object]$Worksheet) {
    $usedRange=$null; $rows=$null; $columns=$null
    try {
        $usedRange=$Worksheet.UsedRange
        $rows=$usedRange.Rows
        $columns=$usedRange.Columns
        $firstRow=[int]$usedRange.Row
        $firstColumn=[int]$usedRange.Column
        $rowCount=[int]$rows.Count
        $columnCount=[int]$columns.Count
        return [pscustomobject]@{
            Matrix=$usedRange.Value2
            FirstRow=$firstRow
            FirstColumn=$firstColumn
            LastRow=$firstRow+$rowCount-1
            LastColumn=$firstColumn+$columnCount-1
            RowCount=$rowCount
            ColumnCount=$columnCount
        }
    } finally {
        Remove-ExcelComReference $columns
        Remove-ExcelComReference $rows
        Remove-ExcelComReference $usedRange
    }
}

function Get-WorksheetDataValue(
    [object]$WorksheetData,
    [int]$Row,
    [int]$Column
) {
    if ($WorksheetData.RowCount -eq 1 -and $WorksheetData.ColumnCount -eq 1) {
        return $WorksheetData.Matrix
    }
    $matrixRow=$Row-$WorksheetData.FirstRow+1
    $matrixColumn=$Column-$WorksheetData.FirstColumn+1
    return $WorksheetData.Matrix[$matrixRow,$matrixColumn]
}

function Get-ExcelOutputValue(
    [object]$WorksheetData,
    [object]$Cells,
    [int]$Row,
    [int]$Column,
    [string]$Path,
    [string[]]$DisplayedTextPaths,
    [string]$WorksheetName
) {
    if ($Path -in $DisplayedTextPaths) {
        return Get-ExcelCellText $Cells $Row $Column $WorksheetName
    }
    return [string](Get-WorksheetDataValue $WorksheetData $Row $Column)
}

function Get-WorksheetJsonLayout(
    [object]$WorksheetData,
    [string]$WorksheetName
) {
    $jsonPathRow=0
    $managementColumn=0
    $searchLastRow=[Math]::Min($WorksheetData.LastRow,$WorksheetData.FirstRow+9)

    for ($row=$WorksheetData.FirstRow; $row -le $searchLastRow; $row++) {
        for ($column=$WorksheetData.FirstColumn; $column -le $WorksheetData.LastColumn; $column++) {
            $heading=[string](Get-WorksheetDataValue $WorksheetData $row $column)
            if ($heading.Trim() -eq "ManagementNo") {
                $jsonPathRow=$row
                $managementColumn=$column
                break
            }
        }
        if ($managementColumn -gt 0) { break }
    }

    if ($managementColumn -eq 0) {
        throw ("{0}シートにManagementNoがありません。" -f $WorksheetName)
    }

    return [pscustomobject]@{
        JsonPathRow=$jsonPathRow
        DataStartRow=$jsonPathRow+1
        ManagementColumn=$managementColumn
        LastRow=$WorksheetData.LastRow
        LastColumn=$WorksheetData.LastColumn
        FirstColumn=$WorksheetData.FirstColumn
    }
}

function Get-JsonPathColumns([object]$WorksheetData,[object]$Layout) {
    $result=New-Object System.Collections.ArrayList
    for ($column=$Layout.FirstColumn; $column -le $Layout.LastColumn; $column++) {
        $path=[string](Get-WorksheetDataValue $WorksheetData $Layout.JsonPathRow $column)
        $path=$path.Trim()
        if ($path -eq "") { continue }
        [void]$result.Add([pscustomobject]@{ Column=$column; Path=$path })
    }
    return $result
}

function Add-Assignment(
    [hashtable]$AssignmentsBySite,
    [string]$SiteNo,
    [string]$Path,
    [string]$Value,
    [string]$SheetName,
    [int]$ExcelRow,
    [int]$ExcelColumn
) {
    if (-not $AssignmentsBySite.ContainsKey($SiteNo)) {
        $AssignmentsBySite[$SiteNo]=New-Object System.Collections.ArrayList
    }
    [void]$AssignmentsBySite[$SiteNo].Add([pscustomobject]@{
        Path=$Path; Value=$Value; SheetName=$SheetName
        ExcelRow=$ExcelRow; ExcelColumn=$ExcelColumn
    })
}

function Read-HorizontalSheet(
    [object]$Worksheet,
    [string]$WorksheetName,
    [string[]]$DisplayedTextPaths
) {
    $data=Get-WorksheetData $Worksheet
    $layout=Get-WorksheetJsonLayout $data $WorksheetName
    $assignments=@{}
    $seen=@{}
    $cells=$null
    try {
        $cells=$Worksheet.Cells
        $pathColumns=Get-JsonPathColumns $data $layout
        for ($row=$layout.DataStartRow; $row -le $layout.LastRow; $row++) {
            $siteNo=([string](Get-WorksheetDataValue $data $row $layout.ManagementColumn)).Trim()
            if ($siteNo -notmatch '^[1-9]\d*$') { continue }
            if ($seen.ContainsKey($siteNo)) {
                throw ("{0}シートでManagementNo {1} が重複しています。" -f $WorksheetName,$siteNo)
            }
            $seen[$siteNo]=$true

            foreach ($columnInfo in $pathColumns) {
                if ($columnInfo.Path -eq "ManagementNo") { continue }
                $value=Get-ExcelOutputValue $data $cells $row $columnInfo.Column $columnInfo.Path $DisplayedTextPaths $WorksheetName
                if ($value -eq "") { continue }
                Add-Assignment $assignments $siteNo $columnInfo.Path $value $WorksheetName $row $columnInfo.Column
            }
        }
    } finally {
        Remove-ExcelComReference $cells
    }
    return $assignments
}

function Convert-VerticalJsonPath(
    [string]$Path,
    [Nullable[int]]$Order,
    [string]$OrderColumn,
    [string]$WorksheetName,
    [int]$ExcelRow
) {
    if ($Path -notmatch '\[\]') { return $Path }
    if ($null -eq $Order) {
        throw ("{0}シート {1}行目: 配列JSONPath {2} には、{3}の1以上の整数が必要です。" -f $WorksheetName,$ExcelRow,$Path,$OrderColumn)
    }
    $firstIndex=$Path.IndexOf('[]')
    $output=$Path.Remove($firstIndex,2).Insert($firstIndex,"[$([int]$Order-1)]")
    return $output.Replace('[]','[0]')
}

function Read-VerticalSheet(
    [object]$Worksheet,
    [string]$WorksheetName,
    [string]$OrderColumn,
    [string[]]$ExcludedPaths,
    [string[]]$DisplayedTextPaths
) {
    $data=Get-WorksheetData $Worksheet
    $layout=Get-WorksheetJsonLayout $data $WorksheetName
    $assignments=@{}
    $cells=$null
    try {
        $cells=$Worksheet.Cells
        $pathColumns=Get-JsonPathColumns $data $layout
        $orderColumns=@($pathColumns | Where-Object Path -eq $OrderColumn)
        if ($orderColumns.Count -gt 1) {
            throw ("{0}シートで{1}列が重複しています。" -f $WorksheetName,$OrderColumn)
        }
        if (@($pathColumns | Where-Object Path -match '\[\]').Count -gt 0 -and $orderColumns.Count -eq 0) {
            throw ("{0}シートに、配列の順序を表す{1}列がありません。" -f $WorksheetName,$OrderColumn)
        }

        for ($row=$layout.DataStartRow; $row -le $layout.LastRow; $row++) {
            $siteNo=([string](Get-WorksheetDataValue $data $row $layout.ManagementColumn)).Trim()
            if ($siteNo -notmatch '^[1-9]\d*$') { continue }

            $order=$null
            if ($orderColumns.Count -eq 1) {
                $orderText=([string](Get-WorksheetDataValue $data $row $orderColumns[0].Column)).Trim()
                if ($orderText -match '^[1-9]\d*$') { $order=[Nullable[int]]([int]$orderText) }
            }

            foreach ($columnInfo in $pathColumns) {
                $path=$columnInfo.Path
                if ($path -eq $OrderColumn -or $path -in $ExcludedPaths) { continue }
                $value=Get-ExcelOutputValue $data $cells $row $columnInfo.Column $path $DisplayedTextPaths $WorksheetName
                if ($value -eq "") { continue }
                $outputPath=Convert-VerticalJsonPath $path $order $OrderColumn $WorksheetName $row
                Add-Assignment $assignments $siteNo $outputPath $value $WorksheetName $row $columnInfo.Column
            }
        }
    } finally {
        Remove-ExcelComReference $cells
    }
    return $assignments
}

function Get-KeyValueSheetLayout(
    [object]$Data,
    [string]$WorksheetName,
    [string]$JsonPathColumnName,
    [string]$ValueColumnName
) {
    $searchLastRow=[Math]::Min($Data.LastRow,$Data.FirstRow+9)
    for ($row=$Data.FirstRow; $row -le $searchLastRow; $row++) {
        $jsonPathColumn=0
        $valueColumn=0
        for ($column=$Data.FirstColumn; $column -le $Data.LastColumn; $column++) {
            $heading=([string](Get-WorksheetDataValue $Data $row $column)).Trim()
            if ($heading -eq $JsonPathColumnName) { $jsonPathColumn=$column }
            if ($heading -eq $ValueColumnName) { $valueColumn=$column }
        }
        if ($jsonPathColumn -gt 0 -and $valueColumn -gt 0) {
            return [pscustomobject]@{
                DataStartRow=$row+1
                JsonPathColumn=$jsonPathColumn
                ValueColumn=$valueColumn
                LastRow=$Data.LastRow
            }
        }
    }
    throw ("{0}シートに'{1}'列と'{2}'列が同じ見出し行にありません。" -f $WorksheetName,$JsonPathColumnName,$ValueColumnName)
}

function Read-KeyValueSheet(
    [object]$Worksheet,
    [string]$WorksheetName,
    [string]$JsonPathColumnName,
    [string]$ValueColumnName,
    [string[]]$DisplayedTextPaths
) {
    $data=Get-WorksheetData $Worksheet
    $layout=Get-KeyValueSheetLayout $data $WorksheetName $JsonPathColumnName $ValueColumnName
    $assignments=New-Object System.Collections.ArrayList
    $cells=$null
    try {
        $cells=$Worksheet.Cells
        for ($row=$layout.DataStartRow; $row -le $layout.LastRow; $row++) {
            $path=([string](Get-WorksheetDataValue $data $row $layout.JsonPathColumn)).Trim()
            if ($path -eq "") { continue }
            $value=Get-ExcelOutputValue $data $cells $row $layout.ValueColumn $path $DisplayedTextPaths $WorksheetName
            if ($value -eq "") { continue }
            [void]$assignments.Add([pscustomobject]@{
                Path=$path; Value=$value; SheetName=$WorksheetName
                ExcelRow=$row; ExcelColumn=$layout.ValueColumn
            })
        }
    } finally {
        Remove-ExcelComReference $cells
    }
    return $assignments
}

Export-ModuleMember -Function @(
    "Remove-ExcelComReference",
    "Read-HorizontalSheet",
    "Read-VerticalSheet",
    "Read-KeyValueSheet"
)

