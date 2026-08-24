<#
    拠点選択文字列をManagementNoの一覧へ変換する。
    空欄またはALLは全拠点を表すため、空の配列を返す。
#>

function ConvertFrom-SiteSelection {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Selection = ""
    )

    $text = $Selection.Trim()
    if ($text -eq "" -or $text -ieq "ALL") {
        return @()
    }

    $numbers = New-Object System.Collections.ArrayList
    $seen = @{}
    $tokens = @($text -split '[,、，]')

    foreach ($rawToken in $tokens) {
        $token = $rawToken.Trim()
        if ($token -eq "") {
            throw "対象拠点の指定に空の要素があります: '$Selection'"
        }

        $start = 0
        $end = 0
        if ($token -match '^([1-9]\d*)\s*[-－ー〜～]\s*([1-9]\d*)$') {
            if (-not [int]::TryParse($Matches[1], [ref]$start) -or
                -not [int]::TryParse($Matches[2], [ref]$end)) {
                throw "対象拠点の番号が大きすぎます: '$token'"
            }
            if ($start -gt $end) {
                throw "対象拠点の範囲は開始番号を終了番号以下にしてください: '$token'"
            }
            if (($end - $start + 1) -gt 10000) {
                throw "対象拠点の1範囲は10000件以下にしてください: '$token'"
            }

            for ($number = $start; $number -le $end; $number++) {
                $key = [string]$number
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true
                [void]$numbers.Add($number)
            }
            continue
        }

        $single = 0
        if (-not [int]::TryParse($token, [ref]$single) -or $single -lt 1) {
            throw "対象拠点は '1' または '1-10, 50-51' の形式で指定してください: '$token'"
        }
        $singleKey = [string]$single
        if (-not $seen.ContainsKey($singleKey)) {
            $seen[$singleKey] = $true
            [void]$numbers.Add($single)
        }
    }

    return @($numbers | Sort-Object)
}

Export-ModuleMember -Function ConvertFrom-SiteSelection
