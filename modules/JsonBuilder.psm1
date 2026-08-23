<#
    Assignmentを拠点別のJSONオブジェクトへ組み立てる内部モジュール。
    Excel COMには依存しない。
#>

function Get-JsonPathValue {
    param(
        [System.Collections.IDictionary]$Root,
        [string]$Path
    )

    $parts = [regex]::Matches($Path, '(?<name>[^.\[\]]+)(?:\[(?<index>\d+)\])?')
    if ($parts.Count -eq 0) { throw "不正な条件JSONPathです: $Path" }
    $current = $Root

    foreach ($part in $parts) {
        $name = $part.Groups["name"].Value
        if ($current -isnot [System.Collections.IDictionary] -or
            -not $current.Contains($name)) {
            return $null
        }

        $current = $current[$name]
        if ($part.Groups["index"].Success) {
            $index = [int]$part.Groups["index"].Value
            if ($current -isnot [System.Collections.IList] -or
                $index -ge $current.Count) {
                return $null
            }
            $current = $current[$index]
        }
    }

    return $current
}

function Test-ApplyWhenAny {
    param(
        [System.Collections.IDictionary]$Body,
        [object[]]$Conditions
    )

    if ($null -eq $Conditions -or $Conditions.Count -eq 0) {
        return $true
    }

    foreach ($condition in $Conditions) {
        $path = ([string]$condition.Path).Trim()
        if ($path -eq "") {
            throw "ApplyWhenAnyのPathが空です。"
        }

        $actualValue = Get-JsonPathValue -Root $Body -Path $path
        if ([string]$actualValue -eq [string]$condition.Equals) {
            return $true
        }
    }

    return $false
}

function Convert-KeyValueAssignmentsToSites {
    param(
        [object[]]$Assignments,
        [hashtable]$BodiesBySite,
        [object[]]$ApplyWhenAny,
        [string]$WorksheetName
    )

    if ($BodiesBySite.Count -eq 0) {
        throw "${WorksheetName}シートより前に、ManagementNoを持つHorizontalまたはVerticalシートを指定してください。"
    }

    $assignmentsBySite = @{}
    foreach ($siteNo in $BodiesBySite.Keys) {
        if (-not (Test-ApplyWhenAny `
            -Body $BodiesBySite[$siteNo] `
            -Conditions $ApplyWhenAny)) {
            continue
        }

        $assignmentsBySite[$siteNo] = New-Object System.Collections.ArrayList
        foreach ($assignment in $Assignments) {
            [void]$assignmentsBySite[$siteNo].Add($assignment)
        }
    }

    return $assignmentsBySite
}

function Set-JsonPath {
    param(
        [System.Collections.IDictionary]$Root,
        [string]$Path,
        [object]$Value
    )

    $parts = [regex]::Matches($Path, '(?<name>[^.\[\]]+)(?:\[(?<index>\d+)\])?')
    if ($parts.Count -eq 0) { throw "不正なJSONPathです: $Path" }
    $current = $Root

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $name = $parts[$i].Groups["name"].Value
        $hasIndex = $parts[$i].Groups["index"].Success
        $isLast = ($i -eq $parts.Count - 1)

        if ($hasIndex) {
            $index = [int]$parts[$i].Groups["index"].Value

            if ($current.Contains($name) -and
                $current[$name] -isnot [System.Collections.IList]) {
                throw "$Path は既存の値またはオブジェクトと配列構造が衝突します。"
            }
            if (-not $current.Contains($name)) {
                $current[$name] = New-Object System.Collections.ArrayList
            }
            while ($current[$name].Count -le $index) {
                [void]$current[$name].Add($null)
            }

            if ($isLast) {
                $current[$name][$index] = $Value
            }
            else {
                if ($null -eq $current[$name][$index]) {
                    $current[$name][$index] = [ordered]@{}
                }
                elseif ($current[$name][$index] -isnot [System.Collections.IDictionary]) {
                    throw "$Path は既存の配列要素とオブジェクト構造が衝突します。"
                }
                $current = $current[$name][$index]
            }
        }
        elseif ($isLast) {
            $current[$name] = $Value
        }
        else {
            if ($current.Contains($name) -and
                $current[$name] -isnot [System.Collections.IDictionary]) {
                throw "$Path は既存の値または配列とオブジェクト構造が衝突します。"
            }
            if (-not $current.Contains($name)) {
                $current[$name] = [ordered]@{}
            }
            $current = $current[$name]
        }
    }
}

function Add-SheetAssignments {
    param(
        [hashtable]$BodiesBySite,
        [hashtable]$OwnersBySite,
        [hashtable]$SheetAssignments
    )

    foreach ($siteNo in $SheetAssignments.Keys) {
        if (-not $BodiesBySite.ContainsKey($siteNo)) {
            $BodiesBySite[$siteNo] = [ordered]@{}
            $OwnersBySite[$siteNo] = @{}
        }

        foreach ($assignment in $SheetAssignments[$siteNo]) {
            $owners = $OwnersBySite[$siteNo]

            if ($owners.ContainsKey($assignment.Path)) {
                $previous = $owners[$assignment.Path]

                # 縦型シートでは、拠点共通値を各行へ同じ値で記載することがある。
                # 同じシート・同じ値の反復だけは、上書きせず同一指定として無視する。
                if ($previous.SheetName -eq $assignment.SheetName -and
                    $previous.Value -eq $assignment.Value) {
                    continue
                }

                throw ("ManagementNo {0}: JSONPath '{1}' が重複しています。" +
                    " 先: {2}シート {3}行目 / 後: {4}シート {5}行目") -f
                    $siteNo,
                    $assignment.Path,
                    $previous.SheetName,
                    $previous.ExcelRow,
                    $assignment.SheetName,
                    $assignment.ExcelRow
            }

            try {
                Set-JsonPath `
                    -Root $BodiesBySite[$siteNo] `
                    -Path $assignment.Path `
                    -Value $assignment.Value
            }
            catch {
                throw ("ManagementNo {0} / {1}シート {2}行目 / JSONPath '{3}': {4}" -f
                    $siteNo,
                    $assignment.SheetName,
                    $assignment.ExcelRow,
                    $assignment.Path,
                    $_.Exception.Message)
            }

            $owners[$assignment.Path] = $assignment
        }
    }
}

function Assert-NoArrayHoles {
    param(
        [object]$Value,
        [string]$Path = '$'
    )

    if ($Value -is [System.Collections.IList]) {
        for ($index = 0; $index -lt $Value.Count; $index++) {
            if ($null -eq $Value[$index]) {
                throw "JSON配列 $Path の[$index]が欠番です。内部管理番号を1から連番にしてください。"
            }
            Assert-NoArrayHoles $Value[$index] "${Path}[$index]"
        }
        return
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            Assert-NoArrayHoles $Value[$key] "${Path}.$key"
        }
    }
}

function ConvertTo-IndentedJson {
    param([object]$Value)

    $rawJson = ConvertTo-Json -InputObject $Value -Depth 100
    $indent = 0
    return ($rawJson -split '\r?\n' | ForEach-Object {
        $line = $_.Trim()
        if ($line -match '^[}\]]') { $indent-- }
        $result = (' ' * ($indent * 4)) + $line
        if ($line -match '[{\[]$') { $indent++ }
        $result
    }) -join [Environment]::NewLine
}

Export-ModuleMember -Function @(
    "Convert-KeyValueAssignmentsToSites",
    "Add-SheetAssignments",
    "Assert-NoArrayHoles",
    "ConvertTo-IndentedJson"
)

