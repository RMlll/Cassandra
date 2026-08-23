<# Windows Forms版 RINK JSONツール。 #>
[CmdletBinding()]
param([switch]$SmokeTest)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$workflowPath = Join-Path $PSScriptRoot "Invoke-JsonWorkflow.ps1"
$defaultExcel = Join-Path $PSScriptRoot "param_v7.0.xlsx"
$defaultOutput = Join-Path $PSScriptRoot "output"
$defaultReport = Join-Path $PSScriptRoot "validation-report.html"
$script:activeJob = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = "Cassandra JSONツール - WinForms"
$form.StartPosition = "CenterScreen"
$form.MinimumSize = [Drawing.Size]::new(760, 560)
$form.Size = [Drawing.Size]::new(900, 650)
$form.Font = [Drawing.Font]::new("Yu Gothic UI", 9)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Cassandra JSON 生成・検証"
$title.Font = [Drawing.Font]::new("Yu Gothic UI", 16, [Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = [Drawing.Point]::new(20, 18)
$form.Controls.Add($title)

function Add-PathRow([string]$labelText, [int]$top, [string]$initialText) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = [Drawing.Point]::new(22, ($top + 5))
    $label.Size = [Drawing.Size]::new(90, 24)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $initialText
    $textBox.Location = [Drawing.Point]::new(115, $top)
    $textBox.Size = [Drawing.Size]::new(650, 26)
    $textBox.Anchor = "Top,Left,Right"
    $form.Controls.Add($textBox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "参照..."
    $button.Location = [Drawing.Point]::new(775, ($top - 1))
    $button.Size = [Drawing.Size]::new(90, 28)
    $button.Anchor = "Top,Right"
    $form.Controls.Add($button)

    return [pscustomobject]@{ TextBox=$textBox; Button=$button }
}

$excelRow = Add-PathRow "対象Excel" 68 $defaultExcel
$outputRow = Add-PathRow "JSON出力先" 108 $defaultOutput

$excelRow.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "対象Excelを選択"
    $dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*"
    $dialog.FileName = $excelRow.TextBox.Text
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $excelRow.TextBox.Text = $dialog.FileName
    }
    $dialog.Dispose()
})

$outputRow.Button.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "JSON出力先を選択"
    $dialog.SelectedPath = $outputRow.TextBox.Text
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $outputRow.TextBox.Text = $dialog.SelectedPath
    }
    $dialog.Dispose()
})

$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "JSON生成"
$generateButton.Location = [Drawing.Point]::new(22, 158)
$generateButton.Size = [Drawing.Size]::new(125, 38)
$form.Controls.Add($generateButton)

$validateButton = New-Object System.Windows.Forms.Button
$validateButton.Text = "JSON検証"
$validateButton.Location = [Drawing.Point]::new(157, 158)
$validateButton.Size = [Drawing.Size]::new(125, 38)
$form.Controls.Add($validateButton)

$allButton = New-Object System.Windows.Forms.Button
$allButton.Text = "生成 → 検証"
$allButton.Location = [Drawing.Point]::new(292, 158)
$allButton.Size = [Drawing.Size]::new(145, 38)
$form.Controls.Add($allButton)

$reportButton = New-Object System.Windows.Forms.Button
$reportButton.Text = "HTMLレポートを開く"
$reportButton.Location = [Drawing.Point]::new(447, 158)
$reportButton.Size = [Drawing.Size]::new(175, 38)
$form.Controls.Add($reportButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "待機中"
$statusLabel.Location = [Drawing.Point]::new(640, 166)
$statusLabel.Size = [Drawing.Size]::new(225, 28)
$statusLabel.TextAlign = "MiddleRight"
$statusLabel.Anchor = "Top,Right"
$form.Controls.Add($statusLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = "Both"
$logBox.WordWrap = $false
$logBox.BackColor = [Drawing.Color]::FromArgb(28, 31, 36)
$logBox.ForeColor = [Drawing.Color]::Gainsboro
$logBox.Font = [Drawing.Font]::new("Consolas", 10)
$logBox.Location = [Drawing.Point]::new(22, 215)
$logBox.Size = [Drawing.Size]::new(843, 365)
$logBox.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($logBox)

function Add-Log([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    $logBox.AppendText($text.TrimEnd() + [Environment]::NewLine)
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

function Set-Busy([bool]$busy) {
    $generateButton.Enabled = -not $busy
    $validateButton.Enabled = -not $busy
    $allButton.Enabled = -not $busy
    $excelRow.Button.Enabled = -not $busy
    $outputRow.Button.Enabled = -not $busy
    $statusLabel.Text = if ($busy) { "実行中..." } else { "待機中" }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    if ($null -eq $script:activeJob) { return }
    foreach ($item in @(Receive-Job -Job $script:activeJob)) {
        Add-Log ([string]$item)
    }
    if ($script:activeJob.State -in @("Completed", "Failed", "Stopped")) {
        $state = $script:activeJob.State
        if ($state -eq "Completed") {
            Add-Log "[$(Get-Date -Format 'HH:mm:ss')] 正常終了"
            $statusLabel.Text = "正常終了"
        } else {
            Add-Log "[$(Get-Date -Format 'HH:mm:ss')] $state : $($script:activeJob.ChildJobs[0].JobStateInfo.Reason.Message)"
            $statusLabel.Text = "エラー"
        }
        Remove-Job -Job $script:activeJob -Force
        $script:activeJob = $null
        $timer.Stop()
        Set-Busy $false
    }
})

function Start-Workflow([string]$mode) {
    if ($null -ne $script:activeJob) { return }
    if ([string]::IsNullOrWhiteSpace($excelRow.TextBox.Text)) {
        [Windows.Forms.MessageBox]::Show("対象Excelを指定してください。", "入力確認", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($outputRow.TextBox.Text)) {
        [Windows.Forms.MessageBox]::Show("JSON出力先を指定してください。", "入力確認", "OK", "Warning") | Out-Null
        return
    }

    Add-Log "[$(Get-Date -Format 'HH:mm:ss')] $mode を開始"
    Set-Busy $true
    $script:activeJob = Start-Job -ScriptBlock {
        param($workflow, $selectedMode, $excel, $output, $report)
        $ErrorActionPreference = "Stop"
        & $workflow -Mode $selectedMode -ExcelPath $excel -OutputDir $output -ReportPath $report *>&1
    } -ArgumentList $workflowPath, $mode, $excelRow.TextBox.Text, $outputRow.TextBox.Text, $defaultReport
    $timer.Start()
}

$generateButton.Add_Click({ Start-Workflow "Generate" })
$validateButton.Add_Click({ Start-Workflow "Validate" })
$allButton.Add_Click({ Start-Workflow "All" })
$reportButton.Add_Click({
    if (Test-Path -LiteralPath $defaultReport) {
        Start-Process -FilePath $defaultReport
    } else {
        [Windows.Forms.MessageBox]::Show("検証レポートがまだありません。", "確認", "OK", "Information") | Out-Null
    }
})

$form.Add_FormClosing({
    if ($null -ne $script:activeJob) {
        Stop-Job -Job $script:activeJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:activeJob -Force -ErrorAction SilentlyContinue
    }
})

if ($SmokeTest) {
    Write-Host "WinForms UI initialization OK"
    $timer.Dispose()
    $form.Dispose()
    return
}

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()
