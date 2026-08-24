[CmdletBinding()]
param(
    [switch]$SmokeTest,
    [string]$PreviewPath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$workflowPath = Join-Path $PSScriptRoot "Invoke-JsonWorkflow.ps1"
$defaultReport = Join-Path $PSScriptRoot "validation-report.html"
$script:activeJob = $null
$script:isClosing = $false

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Cassandra JSON Studio" Width="1080" Height="760"
        MinWidth="920" MinHeight="650" WindowStartupLocation="CenterScreen"
        Background="#EEF2F7" FontFamily="Yu Gothic UI">
  <Window.Resources>
    <Style x:Key="RoundedButton" TargetType="Button">
      <Setter Property="Height" Value="42"/>
      <Setter Property="Padding" Value="18,0"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.86"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.72"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource RoundedButton}">
      <Setter Property="Background" Value="#E9EFF8"/>
      <Setter Property="Foreground" Value="#25344A"/>
    </Style>
    <Style x:Key="BrowseButton" TargetType="Button" BasedOn="{StaticResource RoundedButton}">
      <Setter Property="Height" Value="38"/>
      <Setter Property="Background" Value="#E7ECF3"/>
      <Setter Property="Foreground" Value="#334155"/>
      <Setter Property="Padding" Value="14,0"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="250"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <Border Grid.Column="0" Background="#0F172A">
      <Grid Margin="26,30,24,26">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="52"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal">
          <Border Width="46" Height="46" CornerRadius="12" Background="#38BDF8">
            <TextBlock Text="CJ" Foreground="#082F49" FontSize="17" FontWeight="Bold"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel Margin="12,2,0,0">
            <TextBlock Text="Cassandra" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
            <TextBlock Text="JSON STUDIO" Foreground="#64748B" FontSize="11" FontWeight="Bold"/>
          </StackPanel>
        </StackPanel>

        <StackPanel Grid.Row="2">
          <TextBlock Text="WORKFLOW" Foreground="#64748B" FontSize="11" FontWeight="Bold" Margin="4,0,0,14"/>
          <Border Background="#1E293B" CornerRadius="8" Padding="14,12" Margin="0,0,0,9">
            <StackPanel Orientation="Horizontal">
              <Border Width="26" Height="26" CornerRadius="13" Background="#0EA5E9">
                <TextBlock Text="1" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Margin="11,0,0,0">
                <TextBlock Text="Excelを選択" Foreground="#E2E8F0" FontWeight="SemiBold"/>
                <TextBlock Text="入力元を指定" Foreground="#64748B" FontSize="11"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <Border Background="#1E293B" CornerRadius="8" Padding="14,12" Margin="0,0,0,9">
            <StackPanel Orientation="Horizontal">
              <Border Width="26" Height="26" CornerRadius="13" Background="#8B5CF6">
                <TextBlock Text="2" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Margin="11,0,0,0">
                <TextBlock Text="JSONを生成" Foreground="#E2E8F0" FontWeight="SemiBold"/>
                <TextBlock Text="Value2 / Text変換" Foreground="#64748B" FontSize="11"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <Border Background="#1E293B" CornerRadius="8" Padding="14,12">
            <StackPanel Orientation="Horizontal">
              <Border Width="26" Height="26" CornerRadius="13" Background="#10B981">
                <TextBlock Text="3" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Margin="11,0,0,0">
                <TextBlock Text="API検証" Foreground="#E2E8F0" FontWeight="SemiBold"/>
                <TextBlock Text="HTMLレポート" Foreground="#64748B" FontSize="11"/>
              </StackPanel>
            </StackPanel>
          </Border>
        </StackPanel>

        <Border Grid.Row="4" BorderBrush="#263349" BorderThickness="1" CornerRadius="8" Padding="13">
          <StackPanel>
            <TextBlock Text="LOCAL WORKSPACE" Foreground="#64748B" FontSize="10" FontWeight="Bold"/>
            <TextBlock Text="外部ライブラリ不使用" Foreground="#94A3B8" FontSize="11" Margin="0,5,0,0"/>
            <TextBlock Text="Microsoft .NET / PowerShell" Foreground="#94A3B8" FontSize="11"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <Grid Grid.Column="1" Margin="34,28,34,28">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="22"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="18"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="18"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="JSON Workspace" Foreground="#0F172A" FontSize="28" FontWeight="SemiBold"/>
          <TextBlock Text="ExcelデータをAPI送信用JSONへ変換・検証" Foreground="#64748B" FontSize="13" Margin="0,5,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#E2E8F0" CornerRadius="16" Padding="16,7" VerticalAlignment="Center">
          <StackPanel Orientation="Horizontal">
            <Ellipse Width="8" Height="8" Fill="#22C55E" Margin="0,0,8,0"/>
            <TextBlock x:Name="StatusText" Text="待機中" Foreground="#334155" FontWeight="SemiBold"/>
          </StackPanel>
        </Border>
      </Grid>

      <Border Grid.Row="2" Background="White" CornerRadius="12" Padding="22"
              BorderBrush="#DCE3EC" BorderThickness="1">
        <Border.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="2" Opacity="0.08"/></Border.Effect>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/><RowDefinition Height="18"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="18"/><RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="132"/><ColumnDefinition Width="*"/><ColumnDefinition Width="90"/>
          </Grid.ColumnDefinitions>

          <StackPanel Grid.Row="0" Grid.Column="0" VerticalAlignment="Center">
            <TextBlock Text="INPUT" Foreground="#0EA5E9" FontSize="10" FontWeight="Bold"/>
            <TextBlock Text="対象Excel" Foreground="#1E293B" FontWeight="SemiBold" Margin="0,3,0,0"/>
          </StackPanel>
          <Border Grid.Row="0" Grid.Column="1" Background="#F5F7FA" CornerRadius="8" BorderBrush="#DCE3EC" BorderThickness="1">
            <TextBox x:Name="ExcelPathBox" Height="40" Padding="12,6" BorderThickness="0" Background="Transparent" VerticalContentAlignment="Center"/>
          </Border>
          <Button x:Name="ExcelBrowseButton" Grid.Row="0" Grid.Column="2" Content="参照..." Margin="10,1,0,1" Style="{StaticResource BrowseButton}"/>

          <StackPanel Grid.Row="2" Grid.Column="0" VerticalAlignment="Center">
            <TextBlock Text="TARGET" Foreground="#10B981" FontSize="10" FontWeight="Bold"/>
            <TextBlock Text="対象拠点" Foreground="#1E293B" FontWeight="SemiBold" Margin="0,3,0,0"/>
            <TextBlock Text="空欄=全拠点" Foreground="#64748B" FontSize="10" Margin="0,2,0,0"/>
            <TextBlock Text="例: 1-10, 50-51" Foreground="#64748B" FontSize="10"/>
          </StackPanel>
          <Border Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Background="#F5F7FA" CornerRadius="8" BorderBrush="#DCE3EC" BorderThickness="1">
            <TextBox x:Name="SiteSelectionBox" Height="40" Padding="12,6" BorderThickness="0" Background="Transparent"
                     VerticalContentAlignment="Center" ToolTip="空欄は全拠点。例: 1 または 1-10, 50-51"/>
          </Border>

          <StackPanel Grid.Row="4" Grid.Column="0" VerticalAlignment="Center">
            <TextBlock Text="OUTPUT" Foreground="#8B5CF6" FontSize="10" FontWeight="Bold"/>
            <TextBlock Text="JSON出力先" Foreground="#1E293B" FontWeight="SemiBold" Margin="0,3,0,0"/>
          </StackPanel>
          <Border Grid.Row="4" Grid.Column="1" Background="#F5F7FA" CornerRadius="8" BorderBrush="#DCE3EC" BorderThickness="1">
            <TextBox x:Name="OutputPathBox" Height="40" Padding="12,6" BorderThickness="0" Background="Transparent" VerticalContentAlignment="Center"/>
          </Border>
          <Button x:Name="OutputBrowseButton" Grid.Row="4" Grid.Column="2" Content="参照..." Margin="10,1,0,1" Style="{StaticResource BrowseButton}"/>
        </Grid>
      </Border>

      <Grid Grid.Row="4">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="190"/><ColumnDefinition Width="12"/>
          <ColumnDefinition Width="135"/><ColumnDefinition Width="10"/>
          <ColumnDefinition Width="135"/><ColumnDefinition Width="10"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="AllButton" Grid.Column="0" Height="48" Content="▶  生成して検証"
                Background="#2563EB" Foreground="White" FontSize="14" Style="{StaticResource RoundedButton}"/>
        <Button x:Name="GenerateButton" Grid.Column="2" Height="48" Content="JSON生成" Style="{StaticResource SecondaryButton}"/>
        <Button x:Name="ValidateButton" Grid.Column="4" Height="48" Content="JSON検証" Style="{StaticResource SecondaryButton}"/>
        <Button x:Name="ReportButton" Grid.Column="6" Height="48" Content="レポートを開く ↗" HorizontalAlignment="Right"
                Background="#FFFFFF" Foreground="#334155" BorderBrush="#CBD5E1" BorderThickness="1" Style="{StaticResource RoundedButton}"/>
      </Grid>

      <Border Grid.Row="6" Background="#111827" CornerRadius="12" BorderBrush="#263244" BorderThickness="1">
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="46"/><RowDefinition Height="3"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <Grid Grid.Row="0" Margin="16,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Ellipse Width="9" Height="9" Fill="#FB7185" Margin="0,0,7,0"/>
              <Ellipse Width="9" Height="9" Fill="#FBBF24" Margin="0,0,7,0"/>
              <Ellipse Width="9" Height="9" Fill="#34D399" Margin="0,0,12,0"/>
              <TextBlock Text="PROCESS LOG" Foreground="#94A3B8" FontSize="11" FontWeight="Bold"/>
            </StackPanel>
            <TextBlock Grid.Column="1" Text="PowerShell 5.1" Foreground="#475569" FontSize="11" VerticalAlignment="Center"/>
          </Grid>
          <ProgressBar x:Name="BusyProgress" Grid.Row="1" IsIndeterminate="True" Visibility="Collapsed"
                       Background="#1F2937" Foreground="#38BDF8" BorderThickness="0"/>
          <TextBox x:Name="LogBox" Grid.Row="2" IsReadOnly="True" AcceptsReturn="True" Padding="16,12"
                   VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                   TextWrapping="NoWrap" Background="Transparent" Foreground="#D6E1EE"
                   BorderThickness="0" FontFamily="Consolas" FontSize="13"/>
        </Grid>
      </Border>
    </Grid>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$excelPathBox = $window.FindName("ExcelPathBox")
$siteSelectionBox = $window.FindName("SiteSelectionBox")
$outputPathBox = $window.FindName("OutputPathBox")
$excelBrowseButton = $window.FindName("ExcelBrowseButton")
$outputBrowseButton = $window.FindName("OutputBrowseButton")
$generateButton = $window.FindName("GenerateButton")
$validateButton = $window.FindName("ValidateButton")
$allButton = $window.FindName("AllButton")
$reportButton = $window.FindName("ReportButton")
$statusText = $window.FindName("StatusText")
$busyProgress = $window.FindName("BusyProgress")
$logBox = $window.FindName("LogBox")

$excelPathBox.Text = Join-Path $PSScriptRoot "param_v7.0.xlsx"
$outputPathBox.Text = Join-Path $PSScriptRoot "output"

function Add-Log([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    $logBox.AppendText($text.TrimEnd() + [Environment]::NewLine)
    $logBox.ScrollToEnd()
}

function Set-Busy([bool]$busy) {
    $generateButton.IsEnabled = -not $busy
    $validateButton.IsEnabled = -not $busy
    $allButton.IsEnabled = -not $busy
    $excelBrowseButton.IsEnabled = -not $busy
    $outputBrowseButton.IsEnabled = -not $busy
    $siteSelectionBox.IsEnabled = -not $busy
    $statusText.Text = if ($busy) { "実行中..." } else { "待機中" }
    $busyProgress.Visibility = if ($busy) { "Visible" } else { "Collapsed" }
}

$excelBrowseButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = "対象Excelを選択"
    $dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*"
    $dialog.FileName = $excelPathBox.Text
    if ($dialog.ShowDialog($window)) { $excelPathBox.Text = $dialog.FileName }
})

$outputBrowseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "JSON出力先を選択"
    $dialog.SelectedPath = $outputPathBox.Text
    if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $outputPathBox.Text = $dialog.SelectedPath
    }
    $dialog.Dispose()
})

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    if ($script:isClosing -or $null -eq $script:activeJob) { return }

    $job = $script:activeJob
    foreach ($item in @(Receive-Job -Job $job)) {
        Add-Log ([string]$item)
    }
    if ($job.State -in @("Completed", "Failed", "Stopped")) {
        $state = $job.State
        if ($state -eq "Completed") {
            Add-Log "[$(Get-Date -Format 'HH:mm:ss')] 正常終了"
            $statusText.Text = "正常終了"
        } else {
            Add-Log "[$(Get-Date -Format 'HH:mm:ss')] $state : $($job.ChildJobs[0].JobStateInfo.Reason.Message)"
            $statusText.Text = "エラー"
        }
        $script:activeJob = $null
        $timer.Stop()
        Remove-Job -Job $job -Force
        Set-Busy $false
    }
})

function Stop-ActiveWorkflowJob {
    $job = $script:activeJob
    $script:activeJob = $null
    if ($null -eq $job) { return }

    try {
        if ($job.State -notin @("Completed", "Failed", "Stopped")) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
        }
    } catch {
        # ウィンドウ終了時の後処理エラーで親PowerShellを終了させない。
    }

    try {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    } catch {
        # ジョブが同時に終了・削除された場合も、そのまま画面を閉じる。
    }
}

function Start-Workflow([string]$mode) {
    if ($null -ne $script:activeJob) { return }
    if ([string]::IsNullOrWhiteSpace($excelPathBox.Text)) {
        [Windows.MessageBox]::Show($window, "対象Excelを指定してください。", "入力確認", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($outputPathBox.Text)) {
        [Windows.MessageBox]::Show($window, "JSON出力先を指定してください。", "入力確認", "OK", "Warning") | Out-Null
        return
    }

    Add-Log "[$(Get-Date -Format 'HH:mm:ss')] $mode を開始"
    Set-Busy $true
    $siteSelection = $siteSelectionBox.Text
    $selectionLabel = if ([string]::IsNullOrWhiteSpace($siteSelection)) { "全拠点" } else { $siteSelection.Trim() }
    Add-Log "[$(Get-Date -Format 'HH:mm:ss')] 対象拠点: $selectionLabel"
    $script:activeJob = Start-Job -ScriptBlock {
        param($workflow, $selectedMode, $excel, $output, $report, $sites)
        $ErrorActionPreference = "Stop"
        & $workflow -Mode $selectedMode -ExcelPath $excel -OutputDir $output -ReportPath $report -SiteSelection $sites *>&1
    } -ArgumentList $workflowPath, $mode, $excelPathBox.Text, $outputPathBox.Text, $defaultReport, $siteSelection
    $timer.Start()
}

$generateButton.Add_Click({ Start-Workflow "Generate" })
$validateButton.Add_Click({ Start-Workflow "Validate" })
$allButton.Add_Click({ Start-Workflow "All" })
$reportButton.Add_Click({
    if (Test-Path -LiteralPath $defaultReport) {
        Start-Process -FilePath $defaultReport
    } else {
        [Windows.MessageBox]::Show($window, "検証レポートがまだありません。", "確認", "OK", "Information") | Out-Null
    }
})

$window.Add_Closing({
    $script:isClosing = $true
    $timer.Stop()
})

if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) {
    $previewFullPath = [IO.Path]::GetFullPath($PreviewPath)
    $window.Show()
    $window.UpdateLayout()
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new(
        [int]$window.ActualWidth,
        [int]$window.ActualHeight,
        96,
        96,
        [Windows.Media.PixelFormats]::Pbgra32
    )
    $bitmap.Render($window)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create($previewFullPath)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    $window.Close()
    Write-Host "WPF preview: $previewFullPath"
    return
}

if ($SmokeTest) {
    $autoCloseTimer = New-Object Windows.Threading.DispatcherTimer
    $autoCloseTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $autoCloseTimer.Add_Tick({
        $autoCloseTimer.Stop()
        $window.Close()
    })
    $autoCloseTimer.Start()
}

try {
    [void]$window.ShowDialog()
} finally {
    $script:isClosing = $true
    $timer.Stop()
    Stop-ActiveWorkflowJob
}

if ($SmokeTest) {
    Write-Host "WPF UI initialization OK"
}
