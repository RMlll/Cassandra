<#
    Cassandra JSON Studioのアニメーション付きランチャー。
    起動シーケンス完了後、既存のStart-Wpf.ps1をそのまま呼び出す。
#>
[CmdletBinding()]
param(
    [switch]$SmokeTest,
    [string]$PreviewPath = ""
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$mainWpfPath = Join-Path $PSScriptRoot "Start-Wpf.ps1"
if (-not (Test-Path -LiteralPath $mainWpfPath -PathType Leaf)) {
    throw "既存のWPF起動ファイルがありません: $mainWpfPath"
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Cassandra JSON Studio - Launch Sequence"
        Width="960" Height="560" WindowStartupLocation="CenterScreen"
        WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" Opacity="0" FontFamily="Yu Gothic UI">
  <Window.Resources>
    <LinearGradientBrush x:Key="ShellGradient" StartPoint="0,0" EndPoint="1,1">
      <GradientStop Color="#07111F" Offset="0"/>
      <GradientStop Color="#10182E" Offset="0.48"/>
      <GradientStop Color="#24123C" Offset="1"/>
    </LinearGradientBrush>
    <LinearGradientBrush x:Key="AccentGradient" StartPoint="0,0" EndPoint="1,0">
      <GradientStop Color="#38BDF8" Offset="0"/>
      <GradientStop Color="#2DD4BF" Offset="0.46"/>
      <GradientStop Color="#A78BFA" Offset="1"/>
    </LinearGradientBrush>
    <Style x:Key="SkipButtonStyle" TargetType="Button">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="Background" Value="#17233A"/>
      <Setter Property="BorderBrush" Value="#334155"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="15,7"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="15" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Chrome" Property="Background" Value="#21324F"/>
                <Setter Property="Foreground" Value="White"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Storyboard x:Key="AmbientMotion" RepeatBehavior="Forever">
      <DoubleAnimation Storyboard.TargetName="OuterRingRotate"
                       Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:8"/>
      <DoubleAnimation Storyboard.TargetName="InnerRingRotate"
                       Storyboard.TargetProperty="Angle" From="360" To="0" Duration="0:0:5"/>
      <DoubleAnimation Storyboard.TargetName="ScanTranslate"
                       Storyboard.TargetProperty="X" From="-760" To="1120" Duration="0:0:3.2"/>
      <DoubleAnimation Storyboard.TargetName="GlowOneTranslate"
                       Storyboard.TargetProperty="X" From="-20" To="45" Duration="0:0:4.5"
                       AutoReverse="True"/>
      <DoubleAnimation Storyboard.TargetName="GlowTwoTranslate"
                       Storyboard.TargetProperty="Y" From="20" To="-45" Duration="0:0:5.5"
                       AutoReverse="True"/>
      <DoubleAnimation Storyboard.TargetName="CoreGlow" Storyboard.TargetProperty="Opacity"
                       From="0.35" To="0.9" Duration="0:0:1.3" AutoReverse="True"/>
      <DoubleAnimation Storyboard.TargetName="SignalDotOne" Storyboard.TargetProperty="Opacity"
                       From="0.25" To="1" Duration="0:0:0.7" AutoReverse="True"/>
      <DoubleAnimation Storyboard.TargetName="SignalDotTwo" Storyboard.TargetProperty="Opacity"
                       From="1" To="0.25" BeginTime="0:0:0.2" Duration="0:0:0.7" AutoReverse="True"/>
      <DoubleAnimation Storyboard.TargetName="SignalDotThree" Storyboard.TargetProperty="Opacity"
                       From="0.25" To="1" BeginTime="0:0:0.4" Duration="0:0:0.7" AutoReverse="True"/>
    </Storyboard>
  </Window.Resources>

  <Border CornerRadius="24" BorderBrush="#42577A" BorderThickness="1" Background="{StaticResource ShellGradient}">
    <Border.Effect>
      <DropShadowEffect Color="#020617" BlurRadius="42" ShadowDepth="14" Opacity="0.72"/>
    </Border.Effect>
    <Grid ClipToBounds="True">
      <Canvas IsHitTestVisible="False">
        <Ellipse Width="480" Height="480" Canvas.Left="-180" Canvas.Top="-230" Fill="#173B82" Opacity="0.24">
          <Ellipse.Effect><BlurEffect Radius="90"/></Ellipse.Effect>
          <Ellipse.RenderTransform><TranslateTransform x:Name="GlowOneTranslate"/></Ellipse.RenderTransform>
        </Ellipse>
        <Ellipse Width="500" Height="500" Canvas.Left="680" Canvas.Top="240" Fill="#6D28D9" Opacity="0.23">
          <Ellipse.Effect><BlurEffect Radius="100"/></Ellipse.Effect>
          <Ellipse.RenderTransform><TranslateTransform x:Name="GlowTwoTranslate"/></Ellipse.RenderTransform>
        </Ellipse>
        <Rectangle Width="140" Height="760" Canvas.Left="0" Canvas.Top="-100" Opacity="0.10">
          <Rectangle.Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
              <GradientStop Color="Transparent" Offset="0"/>
              <GradientStop Color="#7DD3FC" Offset="0.5"/>
              <GradientStop Color="Transparent" Offset="1"/>
            </LinearGradientBrush>
          </Rectangle.Fill>
          <Rectangle.RenderTransform>
            <TransformGroup>
              <SkewTransform AngleX="-18"/>
              <TranslateTransform x:Name="ScanTranslate"/>
            </TransformGroup>
          </Rectangle.RenderTransform>
        </Rectangle>
      </Canvas>

      <Grid Margin="54,42,54,40">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel Orientation="Horizontal">
            <Ellipse Width="7" Height="7" Fill="#38BDF8" Margin="0,0,9,0"/>
            <TextBlock Text="CASSANDRA / LOCAL JSON WORKSPACE" Foreground="#7DD3FC"
                       FontFamily="Consolas" FontSize="12" FontWeight="Bold"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Ellipse x:Name="SignalDotOne" Width="6" Height="6" Fill="#2DD4BF" Margin="0,0,6,0"/>
            <Ellipse x:Name="SignalDotTwo" Width="6" Height="6" Fill="#38BDF8" Margin="0,0,6,0"/>
            <Ellipse x:Name="SignalDotThree" Width="6" Height="6" Fill="#A78BFA" Margin="0,0,10,0"/>
            <TextBlock Text="SECURE BOOT" Foreground="#64748B" FontFamily="Consolas" FontSize="11"/>
          </StackPanel>
        </Grid>

        <Grid Grid.Row="1">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="350"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Grid Grid.Column="0" VerticalAlignment="Center" HorizontalAlignment="Center" Width="290" Height="290">
            <Ellipse Width="260" Height="260" Stroke="#1E3A5F" StrokeThickness="1"/>
            <Ellipse Width="236" Height="236" Stroke="#38BDF8" StrokeThickness="2"
                     StrokeDashArray="2,11" Opacity="0.9" RenderTransformOrigin="0.5,0.5">
              <Ellipse.RenderTransform><RotateTransform x:Name="OuterRingRotate"/></Ellipse.RenderTransform>
            </Ellipse>
            <Ellipse Width="196" Height="196" Stroke="#A78BFA" StrokeThickness="1.5"
                     StrokeDashArray="18,5,3,5" Opacity="0.72" RenderTransformOrigin="0.5,0.5">
              <Ellipse.RenderTransform><RotateTransform x:Name="InnerRingRotate"/></Ellipse.RenderTransform>
            </Ellipse>
            <Ellipse Width="154" Height="154" Fill="#0B1830" Stroke="#315074" StrokeThickness="1"/>
            <Ellipse x:Name="CoreGlow" Width="130" Height="130" Opacity="0.5">
              <Ellipse.Fill>
                <RadialGradientBrush>
                  <GradientStop Color="#6938BDF8" Offset="0"/>
                  <GradientStop Color="#00101B30" Offset="1"/>
                </RadialGradientBrush>
              </Ellipse.Fill>
            </Ellipse>
            <Border Width="104" Height="104" CornerRadius="30" Background="{StaticResource AccentGradient}">
              <Border.Effect><DropShadowEffect Color="#38BDF8" BlurRadius="25" Opacity="0.55"/></Border.Effect>
              <Grid>
                <TextBlock Text="CJ" Foreground="#07111F" FontSize="31" FontWeight="Bold"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Grid>
            </Border>
          </Grid>

          <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="38,0,0,0">
            <TextBlock Text="Cassandra" Foreground="White" FontSize="46" FontWeight="SemiBold"/>
            <TextBlock Text="JSON STUDIO" Foreground="#A5B4FC" FontSize="16" FontWeight="Bold"
                       Margin="3,1,0,24"/>
            <Border Background="#101C31" BorderBrush="#263B5B" BorderThickness="1" CornerRadius="12" Padding="18,14">
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="80"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <TextBlock x:Name="PhaseText" Text="BOOT / 01" Foreground="#2DD4BF"
                           FontFamily="Consolas" FontSize="11" FontWeight="Bold"/>
                <StackPanel Grid.Column="1">
                  <TextBlock x:Name="StatusText" Text="ローカルランタイムを初期化" Foreground="#E2E8F0"
                             FontSize="14" FontWeight="SemiBold"/>
                  <TextBlock x:Name="DetailText" Text="Microsoft .NET / PowerShell 5.1" Foreground="#64748B"
                             FontFamily="Consolas" FontSize="11" Margin="0,4,0,0"/>
                </StackPanel>
              </Grid>
            </Border>
          </StackPanel>
        </Grid>

        <Grid Grid.Row="2">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="12"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock Text="LAUNCH SEQUENCE" Foreground="#64748B" FontFamily="Consolas" FontSize="10"/>
            <TextBlock x:Name="PercentText" Grid.Column="1" Text="08%" Foreground="#CBD5E1"
                       FontFamily="Consolas" FontSize="11" FontWeight="Bold"/>
          </Grid>
          <Border Grid.Row="1" Height="4" Background="#1E293B" CornerRadius="2">
            <Border x:Name="ProgressFill" Width="0" HorizontalAlignment="Left" CornerRadius="2"
                    Background="{StaticResource AccentGradient}"/>
          </Border>
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock Text="外部ライブラリ不使用  •  OFFLINE READY  •  ENCRYPTED WORKSPACE"
                       Foreground="#475569" FontFamily="Consolas" FontSize="10"/>
            <Button x:Name="SkipButton" Grid.Column="1" Content="SKIP  ↗" FontFamily="Consolas" FontSize="10"
                    Style="{StaticResource SkipButtonStyle}" Margin="0,12,0,0"/>
          </Grid>
        </Grid>
      </Grid>
    </Grid>
  </Border>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$statusText = $window.FindName("StatusText")
$detailText = $window.FindName("DetailText")
$phaseText = $window.FindName("PhaseText")
$percentText = $window.FindName("PercentText")
$progressFill = $window.FindName("ProgressFill")
$skipButton = $window.FindName("SkipButton")
$ambientMotion = $window.Resources["AmbientMotion"]

$requiredControls = @{
    StatusText=$statusText; DetailText=$detailText; PhaseText=$phaseText
    PercentText=$percentText; ProgressFill=$progressFill; SkipButton=$skipButton
    AmbientMotion=$ambientMotion
}
foreach ($controlName in $requiredControls.Keys) {
    if ($null -eq $requiredControls[$controlName]) {
        throw "起動画面の要素を取得できません: $controlName"
    }
}

$script:stepIndex = 0
$script:isCompleting = $false
$script:continueToStudio = $true
$script:statusTimer = $null
$script:finishTimer = $null
$script:smokeTimer = $null
$script:previewTimer = $null

$steps = @(
    @{ Phase="BOOT / 01"; Status="ローカルランタイムを初期化"; Detail="Microsoft .NET / PowerShell 5.1"; Percent=8 },
    @{ Phase="CORE / 02"; Status="Excel COMブリッジを準備"; Detail="Value2 reader / displayed text adapter"; Percent=27 },
    @{ Phase="RULE / 03"; Status="APIルールと値変換を読込み"; Detail="OpenAPI schema / enum mappings"; Percent=49 },
    @{ Phase="JSON / 04"; Status="JSONワークスペースを構成"; Detail="Horizontal / Vertical / KeyValue"; Percent=72 },
    @{ Phase="UI / 05"; Status="インターフェースを起動"; Detail="Cassandra JSON Studio is ready"; Percent=92 },
    @{ Phase="READY"; Status="起動準備が完了しました"; Detail="Welcome back."; Percent=100 }
)

function Set-LaunchStep([hashtable]$step) {
    $phaseText.Text = $step.Phase
    $statusText.Text = $step.Status
    $detailText.Text = $step.Detail
    $percentText.Text = ("{0:D2}%" -f [int]$step.Percent)

    $animation = New-Object Windows.Media.Animation.DoubleAnimation
    $animation.To = 7.98 * [int]$step.Percent
    $animation.Duration = [Windows.Duration]::new([TimeSpan]::FromMilliseconds(420))
    $animation.EasingFunction = New-Object Windows.Media.Animation.CubicEase
    $animation.EasingFunction.EasingMode = "EaseOut"
    $progressFill.BeginAnimation([Windows.FrameworkElement]::WidthProperty, $animation)
}

function Complete-Splash {
    if ($script:isCompleting) { return }
    $script:isCompleting = $true
    if ($null -ne $script:statusTimer) { $script:statusTimer.Stop() }
    if ($null -ne $script:finishTimer) { $script:finishTimer.Stop() }

    $fade = New-Object Windows.Media.Animation.DoubleAnimation
    $fade.To = 0
    $fade.Duration = [Windows.Duration]::new([TimeSpan]::FromMilliseconds(360))
    $fade.EasingFunction = New-Object Windows.Media.Animation.QuadraticEase
    $fade.Add_Completed({ $window.Close() })
    $window.BeginAnimation([Windows.Window]::OpacityProperty, $fade)
}

$skipButton.Add_Click({ Complete-Splash })
$window.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Key -eq "Escape") {
        $script:continueToStudio = $false
        $window.Close()
    }
    elseif ($eventArgs.Key -in @("Enter", "Space")) {
        Complete-Splash
    }
})

$window.Add_Loaded({
    $fadeIn = New-Object Windows.Media.Animation.DoubleAnimation
    $fadeIn.From = 0
    $fadeIn.To = 1
    $fadeIn.Duration = [Windows.Duration]::new([TimeSpan]::FromMilliseconds(520))
    $window.BeginAnimation([Windows.Window]::OpacityProperty, $fadeIn)
    $ambientMotion.Begin($window, $true)
    Set-LaunchStep $steps[0]

    if ($SmokeTest) {
        $script:continueToStudio = $false
        $script:smokeTimer = New-Object Windows.Threading.DispatcherTimer
        $script:smokeTimer.Interval = [TimeSpan]::FromMilliseconds(180)
        $script:smokeTimer.Add_Tick({
            $script:smokeTimer.Stop()
            $window.Close()
        })
        $script:smokeTimer.Start()
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) {
        $script:continueToStudio = $false
        $script:previewTimer = New-Object Windows.Threading.DispatcherTimer
        $script:previewTimer.Interval = [TimeSpan]::FromMilliseconds(1250)
        $script:previewTimer.Add_Tick({
            $script:previewTimer.Stop()
            Set-LaunchStep $steps[3]
            $window.UpdateLayout()
            $fullPreviewPath = [IO.Path]::GetFullPath($PreviewPath)
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
            $stream = [IO.File]::Create($fullPreviewPath)
            try { $encoder.Save($stream) } finally { $stream.Dispose() }
            $window.Close()
            Write-Host "Animated WPF preview: $fullPreviewPath"
        })
        $script:previewTimer.Start()
        return
    }

    $script:statusTimer = New-Object Windows.Threading.DispatcherTimer
    $script:statusTimer.Interval = [TimeSpan]::FromMilliseconds(560)
    $script:statusTimer.Add_Tick({
        $script:stepIndex++
        if ($script:stepIndex -ge $steps.Count) {
            $script:statusTimer.Stop()
            return
        }

        Set-LaunchStep $steps[$script:stepIndex]
        if ($script:stepIndex -eq ($steps.Count - 1)) {
            $script:statusTimer.Stop()
            $script:finishTimer = New-Object Windows.Threading.DispatcherTimer
            $script:finishTimer.Interval = [TimeSpan]::FromMilliseconds(650)
            $script:finishTimer.Add_Tick({
                $script:finishTimer.Stop()
                Complete-Splash
            })
            $script:finishTimer.Start()
        }
    })
    $script:statusTimer.Start()
})

try {
    [void]$window.ShowDialog()
}
finally {
    if ($null -ne $script:statusTimer) { $script:statusTimer.Stop() }
    if ($null -ne $script:finishTimer) { $script:finishTimer.Stop() }
}

if ($SmokeTest) {
    Write-Host "Animated WPF UI initialization OK"
    return
}

if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) { return }

if ($script:continueToStudio) {
    & $mainWpfPath
}
