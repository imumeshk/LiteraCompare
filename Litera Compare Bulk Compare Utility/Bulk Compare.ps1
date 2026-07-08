param(
    [string]$ExePath,
    [string]$Original,
    [string]$Modified,
    [string]$Output,
    [string]$Prefix = "redline_",
    [string]$Format = ".pdf",
    [string]$Style,
    [ValidateSet("Auto","Word","PowerPoint","Excel","PDF")] [string]$ComparisonEngine = "Auto",
    [ValidateSet("Single","Bulk","Exact")] [string]$Mode = "Single",
    [switch]$UseErrorsDialog,
    [switch]$ShowVisible,
    [switch]$RedlinePagesOnly,
    [switch]$TrackChanges,
    [switch]$RunConsole
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Ensure PowerShell is running in STA mode (required for WPF and Dialogs)
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "Please run this script in a PowerShell console started with the '-STA' switch for the GUI to function correctly."
}

# Global object to store Advanced Settings
$global:AdvSettings = @{
    Silent = $false; AutoStart = $false; AdvancedMode = $false;
    ChangeRep = ""; AutoOrg = ""; AutoMod = "";
    ExMulti = $false; ExBatch = $false; ExAspose = $false;
    Client = ""; Prop = ""
}

# Define the WPF UI via XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Litera Bulk Compare Tool" Height="730" Width="700" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="80"/>
        </Grid.ColumnDefinitions>

        <!-- Row 0: EXE Path -->
        <Label Grid.Row="0" Grid.Column="0" Content="Litera Auto EXE Path:" VerticalAlignment="Center"/>
        <Grid Grid.Row="0" Grid.Column="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="85"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtExe" Grid.Column="0" Margin="5" Text="C:\Program Files (x86)\litera\compare\lcp_auto.exe" VerticalContentAlignment="Center"/>
            <Button x:Name="btnTestExe" Grid.Column="1" Content="Test EXE" Margin="0,5,5,5" Background="#607D8B" Foreground="White" FontWeight="Bold"/>
        </Grid>
        <Button x:Name="btnExe" Grid.Row="0" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 1: Comparison Engine -->
        <Label Grid.Row="1" Grid.Column="0" Content="Comparison Engine:" VerticalAlignment="Center"/>
        <ComboBox x:Name="cmbEngine" Grid.Row="1" Grid.Column="1" Margin="5" SelectedIndex="0" VerticalContentAlignment="Center">
            <ComboBoxItem Content="Auto (lcp_auto.exe)" />
            <ComboBoxItem Content="Word (lcp_main.exe)" />
            <ComboBoxItem Content="PowerPoint (lcp_ppt.exe)" />
            <ComboBoxItem Content="Excel (lcx_main.exe)" />
            <ComboBoxItem Content="PDF (lcp_pdfcmp.exe)" />
        </ComboBox>

        <!-- Row 2: Mode Selection -->
        <Label Grid.Row="2" Grid.Column="0" Content="Comparison Mode:" VerticalAlignment="Center"/>
        <StackPanel Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5">
            <RadioButton x:Name="rbSingle" Content="One Original vs Folder of Modified" IsChecked="True" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <RadioButton x:Name="rbBulk" Content="Folder vs Folder (Sequential match)" VerticalAlignment="Center"/>
            <RadioButton x:Name="rbExact" Content="Folder vs Folder (Exact match)" Margin="15,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Row 3: Original Path -->
        <Label Grid.Row="3" Grid.Column="0" Content="Original (File/Folder):" VerticalAlignment="Center"/>
        <TextBox x:Name="txtOriginal" Grid.Row="3" Grid.Column="1" Margin="5" VerticalContentAlignment="Center"/>
        <Button x:Name="btnOriginal" Grid.Row="3" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 4: Modified Folder -->
        <Label Grid.Row="4" Grid.Column="0" Content="Modified Folder:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtModified" Grid.Row="4" Grid.Column="1" Margin="5" VerticalContentAlignment="Center"/>
        <Button x:Name="btnModified" Grid.Row="4" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 5: Output Folder -->
        <Label Grid.Row="5" Grid.Column="0" Content="Output Folder:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtOutput" Grid.Row="5" Grid.Column="1" Margin="5" VerticalContentAlignment="Center"/>
        <Button x:Name="btnOutput" Grid.Row="5" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 6: Prefix & Format -->
        <Label Grid.Row="6" Grid.Column="0" Content="Output Settings:" VerticalAlignment="Center"/>
        <StackPanel Grid.Row="6" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5">
            <Label Content="Prefix:" VerticalAlignment="Center" Padding="0,5,5,5"/>
            <TextBox x:Name="txtPrefix" Width="100" Text="redline_" VerticalContentAlignment="Center" Margin="0,0,15,0"/>
            <Label Content="Output Format:" VerticalAlignment="Center" Padding="0,5,5,5"/>
            <ComboBox x:Name="cmbFormat" Width="100" SelectedIndex="0" VerticalContentAlignment="Center">
                <ComboBoxItem Content=".pdf"/>
                <ComboBoxItem Content=".docx"/>
                <ComboBoxItem Content=".doc"/>
                <ComboBoxItem Content=".rtf"/>
                <ComboBoxItem Content=".pptm"/>
                <ComboBoxItem Content=".xlsm"/>
            </ComboBox>
        </StackPanel>

        <!-- Row 7: Style Path -->
        <Label Grid.Row="7" Grid.Column="0" Content="Comparison Style:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtStyle" Grid.Row="7" Grid.Column="1" Margin="5" VerticalContentAlignment="Center" ToolTip="Leave empty to use default style"/>
        <Button x:Name="btnStyle" Grid.Row="7" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 8: Options -->
        <Label Grid.Row="8" Grid.Column="0" Content="Options:" VerticalAlignment="Top"/>
        <StackPanel Grid.Row="8" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5" VerticalAlignment="Top">
            <CheckBox x:Name="chkErrors" Content="Show error dialog" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkVisible" Content="Visible window" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkRedlineOnly" Content="Redline pages only" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkTrackChanges" Content="Track changes output" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Row 9: Profile Save/Load -->
        <StackPanel Grid.Row="9" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5,15,5,0" HorizontalAlignment="Right">
            <Button x:Name="btnLoadProfile" Content="Load Profile" Width="100" Height="30" Margin="0,0,10,0" Background="#607D8B" Foreground="White" FontWeight="Bold"/>
            <Button x:Name="btnSaveProfile" Content="Save Profile" Width="100" Height="30" Background="#607D8B" Foreground="White" FontWeight="Bold"/>
        </StackPanel>

        <!-- Row 10: Run, Advanced & Email Buttons -->
        <Grid Grid.Row="10" Grid.Column="0" Grid.ColumnSpan="3" Margin="5,15,5,5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="150"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnRun" Grid.Column="0" Content="START BATCH COMPARISON" Height="40" Margin="0,0,5,0" Background="#4CAF50" Foreground="White" FontWeight="Bold" FontSize="14"/>
            <Button x:Name="btnAdvanced" Grid.Column="1" Content="Advanced Options" Height="40" Margin="5,0,5,0" Background="#FF9800" Foreground="White" FontWeight="Bold" FontSize="14"/>
            <Button x:Name="btnEmail" Grid.Column="2" Content="Email Results" Height="40" Margin="5,0,0,0" Background="#2196F3" Foreground="White" FontWeight="Bold" FontSize="14"/>
        </Grid>

        <!-- Row 11: Progress -->
        <Grid Grid.Row="11" Grid.Column="0" Grid.ColumnSpan="3" Margin="5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <ProgressBar x:Name="pbCompare" Grid.Column="0" Height="20" Minimum="0" Maximum="100" Value="0" Margin="0,0,10,0"/>
            <TextBlock x:Name="txtProgress" Grid.Column="1" Text="Ready" VerticalAlignment="Center" Width="130" TextAlignment="Right"/>
        </Grid>

        <!-- Row 12: Logs -->
        <TextBox x:Name="txtLog" Grid.Row="12" Grid.Column="0" Grid.ColumnSpan="3" Margin="5" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas"/>
    </Grid>
</Window>
"@

# Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map controls to PowerShell variables
$txtExe      = $window.FindName("txtExe")
$btnTestExe  = $window.FindName("btnTestExe")
$btnExe      = $window.FindName("btnExe")
$rbSingle    = $window.FindName("rbSingle")
$rbBulk      = $window.FindName("rbBulk")
$rbExact     = $window.FindName("rbExact")
$txtOriginal = $window.FindName("txtOriginal")
$btnOriginal = $window.FindName("btnOriginal")
$txtModified = $window.FindName("txtModified")
$btnModified = $window.FindName("btnModified")
$txtOutput   = $window.FindName("txtOutput")
$btnOutput   = $window.FindName("btnOutput")
$txtPrefix   = $window.FindName("txtPrefix")
$cmbFormat   = $window.FindName("cmbFormat")
$txtStyle    = $window.FindName("txtStyle")
$btnStyle    = $window.FindName("btnStyle")
$cmbEngine   = $window.FindName("cmbEngine")
$chkErrors   = $window.FindName("chkErrors")
$chkVisible  = $window.FindName("chkVisible")
$chkRedlineOnly = $window.FindName("chkRedlineOnly")
$chkTrackChanges = $window.FindName("chkTrackChanges")
$btnRun      = $window.FindName("btnRun")
$btnAdvanced = $window.FindName("btnAdvanced")
$btnEmail    = $window.FindName("btnEmail")
$btnLoadProfile = $window.FindName("btnLoadProfile")
$btnSaveProfile = $window.FindName("btnSaveProfile")
$pbCompare   = $window.FindName("pbCompare")
$txtProgress = $window.FindName("txtProgress")
$txtLog      = $window.FindName("txtLog")

# Initialize a log file in a temporary location by default
$global:LogFilePath = Join-Path -Path $env:TEMP -ChildPath ("BulkCompare_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
Set-Content -Path $LogFilePath -Value "Litera Bulk Compare Log Started at $(Get-Date)`r`n"

# Helper function for logging
function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$timestamp] $message"
    Add-Content -Path $LogFilePath -Value $line
    if ($txtLog -ne $null) {
        $txtLog.AppendText("$line`r`n")
        $txtLog.ScrollToEnd()
    } else {
        Write-Host $line
    }
}

function Update-Progress {
    param([int]$Current, [int]$Total, [string]$Message)
    if ($pbCompare -ne $null -and $txtProgress -ne $null) {
        $pbCompare.Maximum = if ($Total -gt 0) { $Total } else { 1 }
        $pbCompare.Value = $Current
        $txtProgress.Text = $Message
        # Flush WPF UI thread
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        $pct = if ($Total -gt 0) { ($Current / $Total) * 100 } else { 0 }
        Write-Progress -Activity "Bulk Compare" -Status $Message -PercentComplete $pct
    }
}

# Helper function to interpret Litera exit codes
function Get-LiteraExitMessage([int]$code) {
    switch ($code) {
        0 { return "0: Files compared successfully" }
        1 { return "1: Command line error" }
        2 { return "2: Comparison error" }
        3 { return "3: Error loading original file" }
        4 { return "4: Error loading modified file" }
        5 { return "5: Different content types of files" }
        6 { return "6: File format is not supported" }
        7 { return "7: Result comparison file can't be created" }
        8 { return "8: Litera Compare application is not installed" }
        default { return "$($code): Unknown exit code" }
    }
}

function Get-DocumentCategory {
    param([string]$extension)
    $ext = $extension.TrimStart('.').ToLowerInvariant()
    if (@('doc','docx','rtf','pdf','wpd','htm','html','txt') -contains $ext) {
        return 'Document'
    } elseif (@('xls','xlsx','xlsm','xlsb') -contains $ext) {
        return 'Spreadsheet'
    } elseif (@('ppt','pps','pptx','pptm','ppsx','ppsm') -contains $ext) {
        return 'Presentation'
    } elseif (@('png','bmp','jpg','jpeg') -contains $ext) {
        return 'Image'
    } else {
        return $null
    }
}

function Get-EngineSupportedInputExtensions {
    param([string]$Engine)

    switch ($Engine) {
        'Auto' {
            return @(
                '.doc', '.docx', '.rtf', '.pdf', '.wpd', '.htm', '.html', '.txt',
                '.xls', '.xlsx', '.xlsm', '.xlsb',
                '.ppt', '.pps', '.pptx', '.pptm', '.ppsx', '.ppsm',
                '.png', '.bmp', '.jpg', '.jpeg'
            )
        }
        'Word' {
            return @('.doc', '.docx', '.txt', '.pdf')
        }
        'PowerPoint' {
            return @('.ppt', '.pps', '.pptx', '.pptm', '.ppsx', '.ppsm')
        }
        'Excel' {
            return @('.xls', '.xlsx', '.xlsm', '.xlsb')
        }
        'PDF' {
            return @('.pdf')
        }
        default {
            return @()
        }
    }
}

function Get-EffectiveEngineFromExePath {
    param(
        [string]$ExePath,
        [string]$FallbackEngine = 'Auto'
    )

    $exeName = [System.IO.Path]::GetFileName($ExePath).ToLowerInvariant()
    switch ($exeName) {
        'lcp_auto.exe' { return 'Auto' }
        'lcp_main.exe' { return 'Word' }
        'lcp_ppt.exe' { return 'PowerPoint' }
        'lcx_main.exe' { return 'Excel' }
        'lcp_pdfcmp.exe' { return 'PDF' }
        default { return $FallbackEngine }
    }
}

function Get-FileDialogFilterForEngine {
    param([string]$Engine)

    switch ($Engine) {
        'Word' {
            return "Word / Text / PDF|*.doc;*.docx;*.txt;*.pdf|All Files|*.*"
        }
        'PowerPoint' {
            return "PowerPoint Files|*.ppt;*.pps;*.pptx;*.pptm;*.ppsx;*.ppsm|All Files|*.*"
        }
        'Excel' {
            return "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb|All Files|*.*"
        }
        'PDF' {
            return "PDF Files|*.pdf|All Files|*.*"
        }
        default {
            return "Supported Compare Files|*.doc;*.docx;*.rtf;*.pdf;*.wpd;*.htm;*.html;*.txt;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ppt;*.pps;*.pptx;*.pptm;*.ppsx;*.ppsm;*.png;*.bmp;*.jpg;*.jpeg|All Files|*.*"
        }
    }
}

function Get-FilteredFilesForEngine {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Engine
    )

    $allowed = Get-EngineSupportedInputExtensions -Engine $Engine
    return @(Get-ChildItem -Path $Path -File | Where-Object {
        $_.Name -notmatch '^\~\$' -and $allowed -contains $_.Extension.ToLowerInvariant()
    })
}

function Set-EngineSelectionFromExePath {
    param([string]$ExePath)

    $resolvedEngine = Get-EffectiveEngineFromExePath -ExePath $ExePath -FallbackEngine (Get-SelectedEngine)
    $index = switch ($resolvedEngine) {
        'Auto' { 0 }
        'Word' { 1 }
        'PowerPoint' { 2 }
        'Excel' { 3 }
        'PDF' { 4 }
        default { $null }
    }

    if ($null -ne $index -and $cmbEngine.SelectedIndex -ne $index) {
        $cmbEngine.SelectedIndex = $index
        Write-Log "Updated comparison engine to '$resolvedEngine' based on selected executable."
    }
}

function Get-EngineParameterSupport {
    param([string]$Engine)

    switch ($Engine) {
        'Auto' {
            return @{
                SupportsOutputFile = $true
                SupportsStyle = $true
                SupportsUseErrors = $true
                SupportsVisible = $true
                SupportsRedlineOnly = $false
                SupportsTrackChanges = $false
                SupportsWordPptAdvanced = $false
                SupportsExcelAdvanced = $false
                SupportedOutputLabel = "Word docs: .docx/.doc/.rtf; Excel: .xlsm/.xlsx/.xlsb/.xls; PowerPoint: .pptm/.pptx/.ppt; Images: .png/.bmp/.jpg/.jpeg"
            }
        }
        'Word' {
            return @{
                SupportsOutputFile = $true
                SupportsStyle = $true
                SupportsUseErrors = $false
                SupportsVisible = $true
                SupportsRedlineOnly = $true
                SupportsTrackChanges = $true
                SupportsWordPptAdvanced = $true
                SupportsExcelAdvanced = $false
                SupportedOutputLabel = ".docx/.doc/.rtf/.pdf"
            }
        }
        'PowerPoint' {
            return @{
                SupportsOutputFile = $true
                SupportsStyle = $true
                SupportsUseErrors = $false
                SupportsVisible = $true
                SupportsRedlineOnly = $true
                SupportsTrackChanges = $false
                SupportsWordPptAdvanced = $true
                SupportsExcelAdvanced = $false
                SupportedOutputLabel = ".pptm/.pptx/.ppt"
            }
        }
        'Excel' {
            return @{
                SupportsOutputFile = $true
                SupportsStyle = $true
                SupportsUseErrors = $false
                SupportsVisible = $false
                SupportsRedlineOnly = $false
                SupportsTrackChanges = $false
                SupportsWordPptAdvanced = $false
                SupportsExcelAdvanced = $true
                SupportedOutputLabel = ".xlsm/.xlsx/.xlsb/.xls"
            }
        }
        'PDF' {
            return @{
                SupportsOutputFile = $false
                SupportsStyle = $false
                SupportsUseErrors = $false
                SupportsVisible = $false
                SupportsRedlineOnly = $false
                SupportsTrackChanges = $false
                SupportsWordPptAdvanced = $false
                SupportsExcelAdvanced = $false
                SupportedOutputLabel = "No explicit output file parameter"
            }
        }
        default {
            return @{
                SupportsOutputFile = $true
                SupportsStyle = $false
                SupportsUseErrors = $false
                SupportsVisible = $false
                SupportsRedlineOnly = $false
                SupportsTrackChanges = $false
                SupportsWordPptAdvanced = $false
                SupportsExcelAdvanced = $false
                SupportedOutputLabel = ""
            }
        }
    }
}

function Update-UiForEngine {
    param([string]$Engine)

    $support = Get-EngineParameterSupport -Engine $Engine

    $chkErrors.IsEnabled = $support.SupportsUseErrors
    if (-not $support.SupportsUseErrors) { $chkErrors.IsChecked = $false }

    $chkVisible.IsEnabled = $support.SupportsVisible
    if (-not $support.SupportsVisible) { $chkVisible.IsChecked = $false }

    $chkRedlineOnly.IsEnabled = $support.SupportsRedlineOnly
    if (-not $support.SupportsRedlineOnly) { $chkRedlineOnly.IsChecked = $false }

    $chkTrackChanges.IsEnabled = $support.SupportsTrackChanges
    if (-not $support.SupportsTrackChanges) { $chkTrackChanges.IsChecked = $false }

    $txtPrefix.IsEnabled = $support.SupportsOutputFile
    $cmbFormat.IsEnabled = $support.SupportsOutputFile
    $txtStyle.IsEnabled = $support.SupportsStyle
    $btnStyle.IsEnabled = $support.SupportsStyle
}

function Get-AvailableOutputFormats {
    param(
        [Parameter(Mandatory)] [string]$Engine,
        [string]$OriginalPath
    )

    if ($Engine -eq 'Auto') {
        if (-not [string]::IsNullOrWhiteSpace($OriginalPath) -and (Test-Path -Path $OriginalPath -PathType Leaf)) {
            $category = Get-DocumentCategory([System.IO.Path]::GetExtension($OriginalPath))
            if ($category) {
                return Get-AllowedOutputFormats -engine 'Auto' -category $category
            }
        }

        return @('.docx', '.doc', '.rtf', '.xlsm', '.xlsx', '.xlsb', '.xls', '.pptm', '.pptx', '.ppt', '.png', '.bmp', '.jpg', '.jpeg')
    }

    return Get-AllowedOutputFormats -engine $Engine -category $null
}

function Refresh-OutputFormatOptions {
    param(
        [string]$PreferredFormat
    )

    $engine = Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine)
    $formats = Get-AvailableOutputFormats -Engine $engine -OriginalPath $txtOriginal.Text
    $currentValue = if ($PreferredFormat) { $PreferredFormat } elseif ($cmbFormat.SelectedItem) { [string]$cmbFormat.SelectedItem.Content } else { $null }

    $cmbFormat.Items.Clear()
    foreach ($format in $formats) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $format
        [void]$cmbFormat.Items.Add($item)
    }

    if ($cmbFormat.Items.Count -eq 0) {
        $cmbFormat.Text = ""
        return
    }

    $selectedItem = $cmbFormat.Items | Where-Object { $_.Content -eq $currentValue } | Select-Object -First 1
    if (-not $selectedItem) {
        $selectedItem = $cmbFormat.Items[0]
    }

    $cmbFormat.SelectedItem = $selectedItem
}

function Get-AllowedOutputFormats {
    param(
        [string]$engine,
        [string]$category
    )

    if ($engine -eq 'Auto') {
        switch ($category) {
            'Document' { return @('.docx', '.doc', '.rtf') }
            'Spreadsheet' { return @('.xlsm', '.xlsx', '.xlsb', '.xls') }
            'Presentation' { return @('.pptm', '.pptx', '.ppt') }
            'Image' { return @('.png', '.bmp', '.jpg', '.jpeg') }
            default { return @() }
        }
    } elseif ($engine -eq 'Word') {
        return @('.docx', '.doc', '.rtf', '.pdf')
    } elseif ($engine -eq 'PowerPoint') {
        return @('.pptm', '.pptx', '.ppt')
    } elseif ($engine -eq 'Excel') {
        return @('.xlsm', '.xlsx', '.xlsb', '.xls')
    } elseif ($engine -eq 'PDF') {
        return @()
    } else {
        return @()
    }
}

function Validate-ComparisonPair {
    param(
        [string]$OriginalFile,
        [string]$ModifiedFile,
        [string]$Engine
    )

    $origExt = [System.IO.Path]::GetExtension($OriginalFile).ToLowerInvariant()
    $modExt = [System.IO.Path]::GetExtension($ModifiedFile).ToLowerInvariant()
    $allowedInputs = Get-EngineSupportedInputExtensions -Engine $Engine
    if (-not ($allowedInputs -contains $origExt) -or -not ($allowedInputs -contains $modExt)) {
        return @{ valid = $false; message = "Unsupported file extensions for the $Engine engine: '$origExt' or '$modExt'." }
    }

    $origCat = Get-DocumentCategory($origExt)
    $modCat = Get-DocumentCategory($modExt)

    if ($Engine -eq 'Auto') {
        if ($origCat -eq 'Document' -and $modCat -eq 'Document') {
            return @{ valid = $true }
        }
        if ($origCat -ne $modCat) {
            return @{ valid = $false; message = "Cross-type comparisons like $origCat to $modCat are not supported." }
        }
        return @{ valid = $true }
    }

    switch ($Engine) {
        'Word' {
            if ($allowedInputs -contains $origExt -and $allowedInputs -contains $modExt) { return @{ valid = $true } }
            return @{ valid = $false; message = "Word engine supports only .doc, .docx, .txt, and .pdf comparisons." }
        }
        'PowerPoint' {
            if ($origCat -eq 'Presentation' -and $modCat -eq 'Presentation') { return @{ valid = $true } }
            return @{ valid = $false; message = "PowerPoint engine supports only presentation comparisons." }
        }
        'Excel' {
            if ($origCat -eq 'Spreadsheet' -and $modCat -eq 'Spreadsheet') { return @{ valid = $true } }
            return @{ valid = $false; message = "Excel engine supports only spreadsheet comparisons." }
        }
        'PDF' {
            if ($origExt -eq '.pdf' -and $modExt -eq '.pdf') { return @{ valid = $true } }
            return @{ valid = $false; message = "PDF engine supports only PDF to PDF comparisons." }
        }
        default {
            return @{ valid = $false; message = "Unknown engine selected." }
        }
    }
}

function Get-SelectedEngine {
    switch ($cmbEngine.SelectedIndex) {
        0 { return 'Auto' }
        1 { return 'Word' }
        2 { return 'PowerPoint' }
        3 { return 'Excel' }
        4 { return 'PDF' }
        default { return 'Auto' }
    }
}

function Build-LiteraArguments {
    param(
        [Parameter(Mandatory)] [string]$Engine,
        [Parameter(Mandatory)] [string]$OriginalFile,
        [Parameter(Mandatory)] [string]$ModifiedFile,
        [string]$OutputFile,
        [string]$StyleFile,
        [bool]$UseErrors,
        [bool]$ShowVisible,
        [bool]$RedlineOnly,
        [bool]$TrackChanges,
        [hashtable]$AdvOptions
    )

    $argsList = @()
    $support = Get-EngineParameterSupport -Engine $Engine
    switch ($Engine) {
        'Auto' {
            $argsList += @('-o', $OriginalFile, '-m', $ModifiedFile)
            if ($support.SupportsOutputFile -and $OutputFile) { $argsList += @('-r', $OutputFile) }
            if ($support.SupportsStyle -and $StyleFile) { $argsList += @('-s', $StyleFile) }
            if ($support.SupportsUseErrors -and $UseErrors) { $argsList += '-e' }
            if ($support.SupportsVisible -and $ShowVisible) { $argsList += '-v' }
        }
        'Word' {
            $argsList += @('-org', $OriginalFile, '-mod', $ModifiedFile)
            if ($support.SupportsOutputFile -and $OutputFile) { $argsList += @('-auto', $OutputFile) }
            if ($support.SupportsStyle -and $StyleFile) { $argsList += @('-style', $StyleFile) }
            if ($support.SupportsRedlineOnly -and $RedlineOnly) {
                $redlinePages = [System.IO.Path]::ChangeExtension($OutputFile, '.pdf')
                $argsList += @('-autoredp', $redlinePages)
            }
            if ($support.SupportsTrackChanges -and $TrackChanges -and $OutputFile) {
                $argsList += @('-autotrch', $OutputFile)
            }
            
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Silent) { $argsList += '-silent' }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.AdvancedMode) { $argsList += '-advanced' }
            if ($support.SupportsVisible -and $ShowVisible) {
                $argsList += '-visible'
                if ($support.SupportsWordPptAdvanced -and $AdvOptions.AutoStart) { $argsList += '-cmp' }
            }
            if ($support.SupportsWordPptAdvanced -and $OutputFile -and ($AdvOptions.ChangeRep -or $AdvOptions.AutoOrg -or $AdvOptions.AutoMod)) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
                if ($AdvOptions.ChangeRep) { $argsList += @('-autorep', (Join-Path $AdvOptions.ChangeRep "$baseName`_report.pdf")) }
                if ($AdvOptions.AutoOrg) {
                    $orgExt = [System.IO.Path]::GetExtension($OriginalFile)
                    $argsList += @('-autoorg', (Join-Path $AdvOptions.AutoOrg "$baseName`_orig$orgExt"))
                }
                if ($AdvOptions.AutoMod) {
                    $modExt = [System.IO.Path]::GetExtension($ModifiedFile)
                    $argsList += @('-automod', (Join-Path $AdvOptions.AutoMod "$baseName`_mod$modExt"))
                }
            }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Client) { $argsList += @('-client', $AdvOptions.Client) }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Prop) { $argsList += @('-prop', $AdvOptions.Prop) }
        }
        'PowerPoint' {
            $argsList += @('-org', $OriginalFile, '-mod', $ModifiedFile)
            if ($support.SupportsOutputFile -and $OutputFile) { $argsList += @('-auto', $OutputFile) }
            if ($support.SupportsStyle -and $StyleFile) { $argsList += @('-style', $StyleFile) }
            if ($support.SupportsRedlineOnly -and $RedlineOnly) {
                $redlinePages = [System.IO.Path]::ChangeExtension($OutputFile, '.pdf')
                $argsList += @('-autoredp', $redlinePages)
            }
            if ($support.SupportsTrackChanges -and $TrackChanges -and $OutputFile) {
                $argsList += @('-autotrch', $OutputFile)
            }
            
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Silent) { $argsList += '-silent' }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.AdvancedMode) { $argsList += '-advanced' }
            if ($support.SupportsVisible -and $ShowVisible) {
                $argsList += '-visible'
                if ($support.SupportsWordPptAdvanced -and $AdvOptions.AutoStart) { $argsList += '-cmp' }
            }
            if ($support.SupportsWordPptAdvanced -and $OutputFile -and ($AdvOptions.ChangeRep -or $AdvOptions.AutoOrg -or $AdvOptions.AutoMod)) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
                if ($AdvOptions.ChangeRep) { $argsList += @('-autorep', (Join-Path $AdvOptions.ChangeRep "$baseName`_report.pdf")) }
                if ($AdvOptions.AutoOrg) {
                    $orgExt = [System.IO.Path]::GetExtension($OriginalFile)
                    $argsList += @('-autoorg', (Join-Path $AdvOptions.AutoOrg "$baseName`_orig$orgExt"))
                }
                if ($AdvOptions.AutoMod) {
                    $modExt = [System.IO.Path]::GetExtension($ModifiedFile)
                    $argsList += @('-automod', (Join-Path $AdvOptions.AutoMod "$baseName`_mod$modExt"))
                }
            }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Client) { $argsList += @('-client', $AdvOptions.Client) }
            if ($support.SupportsWordPptAdvanced -and $AdvOptions.Prop) { $argsList += @('-prop', $AdvOptions.Prop) }
        }
        'Excel' {
            if ($AdvOptions.ExBatch) {
                $argsList += @('-b', $OriginalFile, $ModifiedFile)
                if ($support.SupportsOutputFile -and $OutputFile) { $argsList += $OutputFile }
            } else {
                $argsList += @('-s', '-lorg', $OriginalFile, '-lmod', $ModifiedFile)
                if ($support.SupportsOutputFile -and $OutputFile) { $argsList += @('-lres', $OutputFile) }
                if ($support.SupportsStyle -and $StyleFile) { $argsList += @('-style', $StyleFile) }
                if ($support.SupportsExcelAdvanced -and $AdvOptions.ExAspose) { $argsList += '-c' }
                if ($support.SupportsExcelAdvanced -and $AdvOptions.ExMulti) { $argsList += '-m' }
            }
            if ($RedlineOnly -or $TrackChanges) {
                Write-Log "WARNING: Excel engine ignores Redline pages only / Track changes options."
            }
        }
        'PDF' {
            $argsList += @('-org', $OriginalFile, '-mod', $ModifiedFile)
            if ($OutputFile) {
                Write-Log "NOTE: PDF engine uses only original and modified file paths; output file path is ignored by lcp_pdfcmp.exe."
            }
            if ($StyleFile -or $RedlineOnly -or $TrackChanges) {
                Write-Log "WARNING: PDF engine ignores style, redline pages only, and track changes options."
            }
        }
        default {
            $argsList += @('-o', $OriginalFile, '-m', $ModifiedFile)
            if ($OutputFile) { $argsList += @('-r', $OutputFile) }
        }
    }
    return $argsList
}

function Invoke-LiteraProcess {
    param(
        [Parameter(Mandatory)] [string]$ExePath,
        [Parameter(Mandatory)] [string]$Engine,
        [Parameter(Mandatory)] [string]$OriginalFile,
        [Parameter(Mandatory)] [string]$ModifiedFile,
        [string]$OutputFile,
        [string]$StyleFile,
        [bool]$UseErrors,
        [bool]$ShowVisible,
        [bool]$RedlineOnly,
        [bool]$TrackChanges,
        [hashtable]$AdvOptions
    )

    $argsList = Build-LiteraArguments -Engine $Engine -OriginalFile $OriginalFile -ModifiedFile $ModifiedFile -OutputFile $OutputFile -StyleFile $StyleFile -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
    Write-Log "Executing: $ExePath $($argsList -join ' ')"

    try {
        $process = Start-Process -FilePath $ExePath -ArgumentList $argsList -Wait -PassThru -NoNewWindow -ErrorAction Stop
        $exitMessage = Get-LiteraExitMessage $process.ExitCode
        Write-Log "Result: $exitMessage"
        return $process.ExitCode
    } catch {
        Write-Log "ERROR: Failed to start process. $($_.Exception.Message)"
        return -1
    }
}

function Invoke-BulkCompare {
    param(
        [Parameter(Mandatory)] [string]$ExePath,
        [Parameter(Mandatory)] [string]$OrigPath,
        [Parameter(Mandatory)] [string]$ModPath,
        [Parameter(Mandatory)] [string]$OutDir,
        [Parameter(Mandatory)] [string]$Prefix,
        [Parameter(Mandatory)] [string]$Format,
        [string]$Style,
        [Parameter(Mandatory)] [ValidateSet("Auto","Word","PowerPoint","Excel","PDF")] [string]$Engine,
        [Parameter(Mandatory)] [ValidateSet("Single","Bulk","Exact")] [string]$Mode,
        [bool]$UseErrors,
        [bool]$ShowVisible,
        [bool]$RedlineOnly,
        [bool]$TrackChanges,
        [hashtable]$AdvOptions
    )

    if ([string]::IsNullOrWhiteSpace($ExePath) -or -not (Test-Path -Path $ExePath)) { Write-Log "ERROR: Litera EXE path is empty or not found."; return }
    if ([string]::IsNullOrWhiteSpace($OrigPath) -or -not (Test-Path -Path $OrigPath)) { Write-Log "ERROR: Original path is empty or not found."; return }
    if ([string]::IsNullOrWhiteSpace($ModPath) -or -not (Test-Path -Path $ModPath)) { Write-Log "ERROR: Modified path is empty or not found."; return }
    if ([string]::IsNullOrWhiteSpace($OutDir)) { Write-Log "ERROR: Output Folder path is empty."; return }

    if (-not (Test-Path -Path $OutDir)) {
        Write-Log "Creating Output Directory: $OutDir"
        New-Item -ItemType Directory -Path $OutDir | Out-Null
    }
    
    if ($AdvOptions) {
        foreach ($dirProp in @('ChangeRep', 'AutoOrg', 'AutoMod')) {
            $dir = $AdvOptions.$dirProp
            if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -Path $dir)) {
                Write-Log "Creating Advanced Output Directory: $dir"
                New-Item -ItemType Directory -Path $dir | Out-Null
            }
        }
    }

    $outputExtension = [System.IO.Path]::GetExtension($Format).ToLowerInvariant()

    $engineSupport = Get-EngineParameterSupport -Engine $Engine

    if ($Engine -eq 'PDF') {
        if ($outputExtension) {
            Write-Log "WARNING: PDF engine does not support explicit output format; your chosen output format will be ignored for PDF compare." 
        }
    } elseif ($Engine -ne 'Auto') {
        $supportedOutputs = Get-AllowedOutputFormats -engine $Engine -category $null
        if ($outputExtension -and -not ($supportedOutputs -contains $outputExtension)) {
            Write-Log "ERROR: Output format '$outputExtension' is not supported for the selected engine '$Engine'."
            Write-Log "Supported formats for $Engine are: $($supportedOutputs -join ', ')"
            return
        }
    }

    $global:LogFilePath = Join-Path -Path $OutDir -ChildPath ("BulkCompare_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Set-Content -Path $LogFilePath -Value "Litera Bulk Compare Log Started at $(Get-Date)`r`n"

    Write-Log "Selected executable: $ExePath"
    Write-Log "Comparison engine: $Engine"
    Write-Log "Comparison mode: $Mode"
    Write-Log "Original path: $OrigPath"
    Write-Log "Modified path: $ModPath"
    Write-Log "Output folder: $OutDir"
    Write-Log "Output format: $Format"
    Write-Log "Comparison style: $([string]::IsNullOrWhiteSpace($Style) ? 'Default' : $Style)"
    Write-Log "Use error dialog: $UseErrors"
    Write-Log "Visible window: $ShowVisible"
    Write-Log "Redline pages only: $RedlineOnly"
    Write-Log "Track changes output: $TrackChanges"

    $successCount = 0
    $failCount = 0
    $resolvedEngine = Get-EffectiveEngineFromExePath -ExePath $ExePath -FallbackEngine $Engine

    if ($resolvedEngine -ne $Engine) {
        Write-Log "Comparison engine '$Engine' overridden by executable selection. Using '$resolvedEngine' for $([System.IO.Path]::GetFileName($ExePath))."
        $Engine = $resolvedEngine
        $engineSupport = Get-EngineParameterSupport -Engine $Engine
    }

    if ($Engine -eq 'Auto' -and $outputExtension) {
        $selectedCategory = Get-DocumentCategory([System.IO.Path]::GetExtension($OrigPath))
        $supportedOutputs = Get-AllowedOutputFormats -engine 'Auto' -category $selectedCategory
        if ($selectedCategory -and -not ($supportedOutputs -contains $outputExtension)) {
            Write-Log "ERROR: Output format '$outputExtension' is not supported by lcp_auto.exe for $selectedCategory comparisons."
            Write-Log "Supported Auto output formats for $selectedCategory are: $($supportedOutputs -join ', ')"
            return
        }
    }

    switch ($Mode) {
        'Single' {
            Write-Log "Starting Mode: Single Original vs Multiple Modified..."
            $origName = [System.IO.Path]::GetFileNameWithoutExtension($OrigPath)
            $modFiles = Get-FilteredFilesForEngine -Path $ModPath -Engine $Engine
            $totalFiles = $modFiles.Count
            $currentIdx = 0
            foreach ($mFile in $modFiles) {
                $currentIdx++
                Update-Progress -Current $currentIdx -Total $totalFiles -Message "Comparing file $currentIdx of $totalFiles"
                $outputFile = Join-Path -Path $OutDir -ChildPath "$Prefix$origName_$($mFile.BaseName)$Format"
                $validation = Validate-ComparisonPair -OriginalFile $OrigPath -ModifiedFile $mFile.FullName -Engine $Engine
                if (-not $validation.valid) {
                    Write-Log "SKIPPED: $($validation.message) for $OrigPath -> $($mFile.FullName)"
                    $failCount++
                    continue
                }

                Write-Log "Comparing full paths: $OrigPath -> $($mFile.FullName)"
                $exitCode = Invoke-LiteraProcess -ExePath $ExePath -Engine $Engine -OriginalFile $OrigPath -ModifiedFile $mFile.FullName -OutputFile $outputFile -StyleFile $Style -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
                if ($exitCode -eq 0) { $successCount++ } else { $failCount++ }
            }
        }
        'Bulk' {
            Write-Log "Starting Mode: Folder vs Folder (Sequential match)..."
            $origFiles = @(Get-FilteredFilesForEngine -Path $OrigPath -Engine $Engine | Sort-Object Name)
            $modFiles = @(Get-FilteredFilesForEngine -Path $ModPath -Engine $Engine | Sort-Object Name)
            $matchCount = [math]::Min($origFiles.Count, $modFiles.Count)
            if ($matchCount -eq 0) { Write-Log "Warning: One of the folders contains no files to compare." }
            $currentIdx = 0
            for ($i = 0; $i -lt $matchCount; $i++) {
                $currentIdx++
                Update-Progress -Current $currentIdx -Total $matchCount -Message "Comparing file $currentIdx of $matchCount"
                $oFile = $origFiles[$i]
                $mFile = $modFiles[$i]
                $outputFile = Join-Path -Path $OutDir -ChildPath "$Prefix$($oFile.BaseName)_$($mFile.BaseName)$Format"
                $validation = Validate-ComparisonPair -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -Engine $Engine
                if (-not $validation.valid) {
                    Write-Log "SKIPPED: $($validation.message) for $($oFile.FullName) -> $($mFile.FullName)"
                    $failCount++
                    continue
                }

                Write-Log "Comparing full paths: $($oFile.FullName) -> $($mFile.FullName)"
                $exitCode = Invoke-LiteraProcess -ExePath $ExePath -Engine $Engine -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -OutputFile $outputFile -StyleFile $Style -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
                if ($exitCode -eq 0) { $successCount++ } else { $failCount++ }
            }
            if ($origFiles.Count -gt $modFiles.Count) {
                for ($i = $matchCount; $i -lt $origFiles.Count; $i++) {
                    Write-Log "Skipped Original: No corresponding modified file found for '$($origFiles[$i].Name)'"
                }
            } elseif ($modFiles.Count -gt $origFiles.Count) {
                for ($i = $matchCount; $i -lt $modFiles.Count; $i++) {
                    Write-Log "Skipped Modified: No corresponding original file found for '$($modFiles[$i].Name)'"
                }
            }
        }
        'Exact' {
            Write-Log "Starting Mode: Folder vs Folder (Exact match)..."
            $origFiles = Get-FilteredFilesForEngine -Path $OrigPath -Engine $Engine
            $modFiles = Get-FilteredFilesForEngine -Path $ModPath -Engine $Engine
            $totalFiles = $origFiles.Count
            $currentIdx = 0
            foreach ($oFile in $origFiles) {
                $currentIdx++
                Update-Progress -Current $currentIdx -Total $totalFiles -Message "Comparing file $currentIdx of $totalFiles"
                $mFile = $modFiles | Where-Object { $_.BaseName -eq $oFile.BaseName } | Select-Object -First 1
                if ($mFile) {
                    $outputFile = Join-Path -Path $OutDir -ChildPath "$Prefix$($oFile.BaseName)_$($mFile.BaseName)$Format"
                    $validation = Validate-ComparisonPair -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -Engine $Engine
                    if (-not $validation.valid) {
                        Write-Log "SKIPPED: $($validation.message) for $($oFile.FullName) -> $($mFile.FullName)"
                        $failCount++
                        continue
                    }

                    Write-Log "Comparing full paths: $($oFile.FullName) -> $($mFile.FullName)"
                    $exitCode = Invoke-LiteraProcess -ExePath $ExePath -Engine $Engine -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -OutputFile $outputFile -StyleFile $Style -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
                    if ($exitCode -eq 0) { $successCount++ } else { $failCount++ }
                } else {
                    Write-Log "Skipped: No matching modified file found for '$($oFile.Name)'"
                }
            }
        }
    }

    Write-Log "BATCH COMPLETE. Success: $successCount, Failed: $failCount"
    Write-Log "Final log saved to: $LogFilePath"
    Update-Progress -Current 0 -Total 100 -Message "Ready"
}

# --- Event Handlers ---

$cmbEngine.Add_SelectionChanged({
    $selectedItem = $cmbEngine.SelectedItem
    if ($selectedItem -is [System.Windows.Controls.ComboBoxItem]) {
        $engineName = Get-SelectedEngine
        Update-UiForEngine -Engine $engineName

        # Auto-update the EXE path based on engine selection
        $currentExePath = $txtExe.Text
        if (-not [string]::IsNullOrWhiteSpace($currentExePath)) {
            $baseDir = [System.IO.Path]::GetDirectoryName($currentExePath)
            $newExeName = switch -Regex ($engineName) {
                "Word" { "lcp_main.exe" }
                "PowerPoint" { "lcp_ppt.exe" }
                "Excel" { "lcx_main.exe" }
                "PDF" { "lcp_pdfcmp.exe" }
                default { "lcp_auto.exe" }
            }
            $newExePath = Join-Path -Path $baseDir -ChildPath $newExeName
            $txtExe.Text = $newExePath
            Write-Log "Automatically updated EXE path to: $newExePath"
        }

        Refresh-OutputFormatOptions
    }
})

$cmbFormat.Add_SelectionChanged({
    $selectedFormatItem = $cmbFormat.SelectedItem
    if ($selectedFormatItem) {
        $selectedFormat = $selectedFormatItem.Content
        $activeExeEngine = Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine)
        # If output is any document type, auto-switch to the robust Word engine.
        if ($activeExeEngine -ne 'Auto' -and @('.pdf', '.docx', '.doc', '.rtf') -contains $selectedFormat) {
            foreach ($item in $cmbEngine.Items) {
                if ($item.Content -match 'Word') {
                    if ($cmbEngine.SelectedItem -ne $item) {
                        $cmbEngine.SelectedItem = $item
                        Write-Log "Automatically switched to Word engine for document output format '$selectedFormat'."
                    }
                    break
                }
            }
        }

        if ($activeExeEngine -eq 'Auto') {
            $currentOriginalPath = $txtOriginal.Text
            if (-not [string]::IsNullOrWhiteSpace($currentOriginalPath) -and (Test-Path -Path $currentOriginalPath -PathType Leaf)) {
                $category = Get-DocumentCategory([System.IO.Path]::GetExtension($currentOriginalPath))
                $supportedFormats = Get-AllowedOutputFormats -engine 'Auto' -category $category
                if ($category -and -not ($supportedFormats -contains $selectedFormat)) {
                    Write-Log "Selected output format '$selectedFormat' is not supported by lcp_auto.exe for $category comparisons. Supported formats: $($supportedFormats -join ', ')"
                }
            }
        }
    }
})

$chkTrackChanges.Add_Click({
    if ($chkTrackChanges.IsChecked) {
        # Track changes is ON. Restrict formats to Word documents.
        $validFormats = @(".docx", ".doc")
        $currentSelection = $cmbFormat.SelectedItem.Content
        
        # If current selection is invalid, switch to .docx as the default
        if ($validFormats -notcontains $currentSelection) {
            foreach ($item in $cmbFormat.Items) {
                if ($item.Content -eq ".docx") {
                    $cmbFormat.SelectedItem = $item
                    break
                }
            }
        }

        # Disable invalid items
        foreach ($item in $cmbFormat.Items) {
            $item.IsEnabled = ($validFormats -contains $item.Content)
        }
    } else {
        # Track changes is OFF. Re-enable all formats.
        foreach ($item in $cmbFormat.Items) {
            $item.IsEnabled = $true
        }
    }
})

$btnExe.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Executable Files|*.exe|All Files|*.*"
    if ($dialog.ShowDialog() -eq 'OK') {
        $txtExe.Text = $dialog.FileName
        Write-Log "Selected EXE: $($dialog.FileName)"
        Set-EngineSelectionFromExePath -ExePath $dialog.FileName
        Refresh-OutputFormatOptions
    }
})

$btnTestExe.Add_Click({
    $exePath = $txtExe.Text
    if ([string]::IsNullOrWhiteSpace($exePath) -or -not (Test-Path -Path $exePath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show("Executable not found at the specified path.`nPlease check the path and try again.", "Test Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    try {
        Write-Log "Testing executable: $exePath"
        # Launch process in hidden window. 
        $proc = Start-Process -FilePath $exePath -ArgumentList "-h" -PassThru -WindowStyle Hidden -ErrorAction Stop
        
        # Allow up to 3 seconds for it to exit (e.g. if it instantly prints a help command)
        $exited = $proc.WaitForExit(3000)
        if (-not $exited) {
            $proc.Kill()
            Write-Log "Test EXE: Process launched successfully but did not exit within 3 seconds. Terminated test instance."
        } else {
            Write-Log "Test EXE: Process launched and exited with code $($proc.ExitCode)."
        }
        
        [System.Windows.MessageBox]::Show("Successfully connected to the executable!`n`nPath: $exePath", "Test Successful", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        Write-Log "Test EXE ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to launch the executable.`n`nError: $($_.Exception.Message)", "Test Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnOriginal.Add_Click({
    $activeEngine = Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine)
    if ($rbSingle.IsChecked) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = Get-FileDialogFilterForEngine -Engine $activeEngine
        $dialog.Title = "Select Original Document"
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtOriginal.Text = $dialog.FileName
            Write-Log "Selected Original Document: $($dialog.FileName)"
            Refresh-OutputFormatOptions
        }
    } else {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select Folder containing Original Documents"
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtOriginal.Text = $dialog.SelectedPath
            Write-Log "Selected Original Folder: $($dialog.SelectedPath)"
            Refresh-OutputFormatOptions
        }
    }
})

$btnModified.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select Folder containing Modified Documents"
    if ($dialog.ShowDialog() -eq 'OK') { $txtModified.Text = $dialog.SelectedPath; Write-Log "Selected Modified Folder: $($dialog.SelectedPath)" }
})

$btnOutput.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select Output Folder for Redlines"
    if ($dialog.ShowDialog() -eq 'OK') { $txtOutput.Text = $dialog.SelectedPath; Write-Log "Selected Output Folder: $($dialog.SelectedPath)" }
})

$btnStyle.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Comparison Styles (*.tpx;*.tpz;*.tpp)|*.tpx;*.tpz;*.tpp|All Files|*.*"
    if ($dialog.ShowDialog() -eq 'OK') { $txtStyle.Text = $dialog.FileName; Write-Log "Selected Comparison Style: $($dialog.FileName)" }
})

$btnAdvanced.Add_Click({
    [xml]$advXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Advanced Litera Options" Height="550" Width="500" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Word &amp; PowerPoint Options" FontWeight="Bold" Margin="0,0,0,5" />
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <CheckBox x:Name="advChkSilent" Content="Silent Mode (-silent)" Margin="0,0,15,0" />
            <CheckBox x:Name="advChkAutoStart" Content="Auto-Start Compare (-cmp)" Margin="0,0,15,0" ToolTip="Requires 'Visible window' checked on main screen" />
            <CheckBox x:Name="advChkAdvanced" Content="Advanced Mode (-advanced)" />
        </StackPanel>

        <TextBlock Grid.Row="2" Text="Additional Output Directories" FontWeight="Bold" Margin="0,0,0,5" />
        <Grid Grid.Row="3" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="40"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Label Grid.Row="0" Grid.Column="0" Content="Change Reports:" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtChangeRep" Grid.Row="0" Grid.Column="1" Margin="2" />
            <Button x:Name="advBtnChangeRep" Grid.Row="0" Grid.Column="2" Content="..." Margin="2" />

            <Label Grid.Row="1" Grid.Column="0" Content="Original Copies:" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtAutoOrg" Grid.Row="1" Grid.Column="1" Margin="2" />
            <Button x:Name="advBtnAutoOrg" Grid.Row="1" Grid.Column="2" Content="..." Margin="2" />

            <Label Grid.Row="2" Grid.Column="0" Content="Modified Copies:" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtAutoMod" Grid.Row="2" Grid.Column="1" Margin="2" />
            <Button x:Name="advBtnAutoMod" Grid.Row="2" Grid.Column="2" Content="..." Margin="2" />
        </Grid>

        <TextBlock Grid.Row="4" Text="Excel Options" FontWeight="Bold" Margin="0,0,0,5" />
        <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,10">
            <CheckBox x:Name="advChkExMulti" Content="Multiple Instances (-m)" Margin="0,0,15,0" />
            <CheckBox x:Name="advChkExBatch" Content="Native Batch Mode (-b)" Margin="0,0,15,0" />
            <CheckBox x:Name="advChkExAspose" Content="Aspose Fallback (-c)" />
        </StackPanel>

        <TextBlock Grid.Row="6" Text="Application Properties" FontWeight="Bold" Margin="0,0,0,5" />
        <Grid Grid.Row="7" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="100"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Label Grid.Row="0" Grid.Column="0" Content="Client Name:" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtClient" Grid.Row="0" Grid.Column="1" Margin="2" />
            
            <Label Grid.Row="1" Grid.Column="0" Content="Properties:" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtProp" Grid.Row="1" Grid.Column="1" Margin="2" />
        </Grid>

        <StackPanel Grid.Row="8" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="advBtnSave" Content="Save Options" Width="100" Height="30" Background="#4CAF50" Foreground="White" FontWeight="Bold" Margin="0,0,10,0"/>
            <Button x:Name="advBtnCancel" Content="Cancel" Width="80" Height="30" Background="#f44336" Foreground="White" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $advReader = (New-Object System.Xml.XmlNodeReader $advXaml)
    $advWindow = [Windows.Markup.XamlReader]::Load($advReader)

    # Mapping controls
    $advChkSilent = $advWindow.FindName("advChkSilent"); $advChkSilent.IsChecked = $global:AdvSettings.Silent
    $advChkAutoStart = $advWindow.FindName("advChkAutoStart"); $advChkAutoStart.IsChecked = $global:AdvSettings.AutoStart
    $advChkAdvanced = $advWindow.FindName("advChkAdvanced"); $advChkAdvanced.IsChecked = $global:AdvSettings.AdvancedMode
    $advTxtChangeRep = $advWindow.FindName("advTxtChangeRep"); $advTxtChangeRep.Text = $global:AdvSettings.ChangeRep
    $advTxtAutoOrg = $advWindow.FindName("advTxtAutoOrg"); $advTxtAutoOrg.Text = $global:AdvSettings.AutoOrg
    $advTxtAutoMod = $advWindow.FindName("advTxtAutoMod"); $advTxtAutoMod.Text = $global:AdvSettings.AutoMod
    $advChkExMulti = $advWindow.FindName("advChkExMulti"); $advChkExMulti.IsChecked = $global:AdvSettings.ExMulti
    $advChkExBatch = $advWindow.FindName("advChkExBatch"); $advChkExBatch.IsChecked = $global:AdvSettings.ExBatch
    $advChkExAspose = $advWindow.FindName("advChkExAspose"); $advChkExAspose.IsChecked = $global:AdvSettings.ExAspose
    $advTxtClient = $advWindow.FindName("advTxtClient"); $advTxtClient.Text = $global:AdvSettings.Client
    $advTxtProp = $advWindow.FindName("advTxtProp"); $advTxtProp.Text = $global:AdvSettings.Prop
    
    # Directory buttons
    foreach ($pair in @(@("advBtnChangeRep","advTxtChangeRep"),@("advBtnAutoOrg","advTxtAutoOrg"),@("advBtnAutoMod","advTxtAutoMod"))) {
        $btn = $advWindow.FindName($pair[0]); $txt = $advWindow.FindName($pair[1])
        $btn.Add_Click({
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($fbd.ShowDialog() -eq 'OK') { $this.Tag.Text = $fbd.SelectedPath }
        }.GetNewClosure())
        $btn.Tag = $txt
    }

    $advWindow.FindName("advBtnCancel").Add_Click({ $advWindow.Close() })
    $advWindow.FindName("advBtnSave").Add_Click({
        $global:AdvSettings.Silent = [bool]$advChkSilent.IsChecked
        $global:AdvSettings.AutoStart = [bool]$advChkAutoStart.IsChecked
        $global:AdvSettings.AdvancedMode = [bool]$advChkAdvanced.IsChecked
        $global:AdvSettings.ChangeRep = $advTxtChangeRep.Text
        $global:AdvSettings.AutoOrg = $advTxtAutoOrg.Text
        $global:AdvSettings.AutoMod = $advTxtAutoMod.Text
        $global:AdvSettings.ExMulti = [bool]$advChkExMulti.IsChecked
        $global:AdvSettings.ExBatch = [bool]$advChkExBatch.IsChecked
        $global:AdvSettings.ExAspose = [bool]$advChkExAspose.IsChecked
        $global:AdvSettings.Client = $advTxtClient.Text
        $global:AdvSettings.Prop = $advTxtProp.Text
        $advWindow.Close()
    })
    $advWindow.ShowDialog() | Out-Null
})

$btnSaveProfile.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "JSON Profile (*.json)|*.json|All Files|*.*"
    if ($dialog.ShowDialog() -eq 'OK') {
        $profile = [PSCustomObject]@{
            ExePath = $txtExe.Text
            EngineIndex = $cmbEngine.SelectedIndex
            Mode = if ($rbSingle.IsChecked) { 'Single' } elseif ($rbBulk.IsChecked) { 'Bulk' } else { 'Exact' }
            Original = $txtOriginal.Text
            Modified = $txtModified.Text
            Output = $txtOutput.Text
            Prefix = $txtPrefix.Text
            FormatIndex = $cmbFormat.SelectedIndex
            Style = $txtStyle.Text
            Errors = [bool]$chkErrors.IsChecked
            Visible = [bool]$chkVisible.IsChecked
            Redline = [bool]$chkRedlineOnly.IsChecked
            TrackChanges = [bool]$chkTrackChanges.IsChecked
            AdvSettings = $global:AdvSettings
        }
        $profile | ConvertTo-Json -Depth 5 | Set-Content -Path $dialog.FileName
        Write-Log "Profile saved to $($dialog.FileName)"
    }
})

$btnLoadProfile.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "JSON Profile (*.json)|*.json|All Files|*.*"
    if ($dialog.ShowDialog() -eq 'OK') {
        try {
            $profile = Get-Content -Path $dialog.FileName -Raw | ConvertFrom-Json
            if ($profile) {
                $txtExe.Text = $profile.ExePath
                $cmbEngine.SelectedIndex = $profile.EngineIndex
                Set-EngineSelectionFromExePath -ExePath $profile.ExePath
                $rbSingle.IsChecked = ($profile.Mode -eq 'Single')
                $rbBulk.IsChecked = ($profile.Mode -eq 'Bulk')
                $rbExact.IsChecked = ($profile.Mode -eq 'Exact')
                $txtOriginal.Text = $profile.Original
                $txtModified.Text = $profile.Modified
                $txtOutput.Text = $profile.Output
                $txtPrefix.Text = $profile.Prefix
                Refresh-OutputFormatOptions
                if ($profile.PSObject.Properties.Name -contains 'FormatIndex' -and $profile.FormatIndex -lt $cmbFormat.Items.Count) {
                    $cmbFormat.SelectedIndex = $profile.FormatIndex
                }
                $txtStyle.Text = $profile.Style
                $chkErrors.IsChecked = $profile.Errors
                $chkVisible.IsChecked = $profile.Visible
                $chkRedlineOnly.IsChecked = $profile.Redline
                $chkTrackChanges.IsChecked = $profile.TrackChanges
                
                if ($profile.AdvSettings) {
                    $global:AdvSettings.Silent = [bool]$profile.AdvSettings.Silent
                    $global:AdvSettings.AutoStart = [bool]$profile.AdvSettings.AutoStart
                    $global:AdvSettings.AdvancedMode = [bool]$profile.AdvSettings.AdvancedMode
                    $global:AdvSettings.ChangeRep = [string]$profile.AdvSettings.ChangeRep
                    $global:AdvSettings.AutoOrg = [string]$profile.AdvSettings.AutoOrg
                    $global:AdvSettings.AutoMod = [string]$profile.AdvSettings.AutoMod
                    $global:AdvSettings.ExMulti = [bool]$profile.AdvSettings.ExMulti
                    $global:AdvSettings.ExBatch = [bool]$profile.AdvSettings.ExBatch
                    $global:AdvSettings.ExAspose = [bool]$profile.AdvSettings.ExAspose
                    $global:AdvSettings.Client = [string]$profile.AdvSettings.Client
                    $global:AdvSettings.Prop = [string]$profile.AdvSettings.Prop
                }
                Write-Log "Profile loaded from $($dialog.FileName)"
            }
        } catch {
            Write-Log "ERROR: Failed to load profile. $($_.Exception.Message)"
        }
    }
})

$btnRun.Add_Click({
    $exePath   = $txtExe.Text
    $origPath  = $txtOriginal.Text
    $modPath   = $txtModified.Text
    $outDir    = $txtOutput.Text
    $prefix    = $txtPrefix.Text
    $format    = $cmbFormat.Text
    $style     = $txtStyle.Text
    $engine    = Get-SelectedEngine
    $mode      = if ($rbSingle.IsChecked) { 'Single' } elseif ($rbBulk.IsChecked) { 'Bulk' } else { 'Exact' }
    $useErrors = $chkErrors.IsChecked
    $visible   = $chkVisible.IsChecked
    $redline   = $chkRedlineOnly.IsChecked
    $track     = $chkTrackChanges.IsChecked

    $btnRun.IsEnabled = $false
    Invoke-BulkCompare -ExePath $exePath -OrigPath $origPath -ModPath $modPath -OutDir $outDir -Prefix $prefix -Format $format -Style $style -Engine $engine -Mode $mode -UseErrors $useErrors -ShowVisible $visible -RedlineOnly $redline -TrackChanges $track -AdvOptions $global:AdvSettings
    $btnRun.IsEnabled = $true
})

$btnEmail.Add_Click({
    $outDir = $txtOutput.Text
    if ([string]::IsNullOrWhiteSpace($outDir) -or -not (Test-Path -Path $outDir)) {
        [System.Windows.MessageBox]::Show("Please select a valid Output Folder that exists.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    $files = Get-ChildItem -Path $outDir -File
    if ($files.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No files found in the Output Folder.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # Define Email Dialog XAML
    [xml]$emailXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email Results" Height="500" Width="450" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="To:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtTo" Grid.Row="0" Grid.Column="1" Margin="5" VerticalContentAlignment="Center"/>

        <Label Grid.Row="1" Grid.Column="0" Content="CC:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtCc" Grid.Row="1" Grid.Column="1" Margin="5" VerticalContentAlignment="Center"/>

        <Label Grid.Row="2" Grid.Column="0" Content="Subject:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtSubject" Grid.Row="2" Grid.Column="1" Margin="5" VerticalContentAlignment="Center" Text="Litera Compare Results"/>

        <Label Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Content="Select Files to Attach:" FontWeight="Bold" Margin="0,10,0,5"/>
        
        <Border Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" BorderBrush="Gray" BorderThickness="1" Background="White" Margin="5">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="pnlFiles" Margin="5"/>
            </ScrollViewer>
        </Border>

        <StackPanel Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5">
            <CheckBox x:Name="chkSelectAll" Content="Select All" VerticalAlignment="Center" Margin="0,0,20,0"/>
            <CheckBox x:Name="chkZip" Content="Compress selected files to ZIP" VerticalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnCreateEmail" Content="Create Draft" Width="100" Height="30" Margin="0,0,10,0" Background="#4CAF50" Foreground="White" FontWeight="Bold"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="30" Background="#f44336" Foreground="White" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $emailReader = (New-Object System.Xml.XmlNodeReader $emailXaml)
    $emailWindow = [Windows.Markup.XamlReader]::Load($emailReader)

    $txtTo = $emailWindow.FindName("txtTo")
    $txtCc = $emailWindow.FindName("txtCc")
    $txtSubject = $emailWindow.FindName("txtSubject")
    $pnlFiles = $emailWindow.FindName("pnlFiles")
    $chkSelectAll = $emailWindow.FindName("chkSelectAll")
    $chkZip = $emailWindow.FindName("chkZip")
    $btnCreateEmail = $emailWindow.FindName("btnCreateEmail")
    $btnCancel = $emailWindow.FindName("btnCancel")

    # Populate files
    foreach ($file in $files) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $file.Name
        $cb.Tag = $file.FullName
        $cb.Margin = "0,2,0,2"
        $pnlFiles.Children.Add($cb) | Out-Null
    }

    $chkSelectAll.Add_Click({
        $isChecked = $chkSelectAll.IsChecked
        foreach ($cb in $pnlFiles.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox]) {
                $cb.IsChecked = $isChecked
            }
        }
    })

    $btnCancel.Add_Click({
        $emailWindow.Close()
    })

    $btnCreateEmail.Add_Click({
        $selectedFiles = @()
        foreach ($cb in $pnlFiles.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.IsChecked) {
                $selectedFiles += $cb.Tag
            }
        }
        
        if ($selectedFiles.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one file to attach.", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        try {
            $outlook = New-Object -ComObject Outlook.Application
            $mail = $outlook.CreateItem(0)
            $mail.To = $txtTo.Text
            $mail.CC = $txtCc.Text
            $mail.Subject = $txtSubject.Text

            if ($chkZip.IsChecked) {
                $zipPath = Join-Path -Path $env:TEMP -ChildPath ("LiteraResults_{0}.zip" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                Compress-Archive -Path $selectedFiles -DestinationPath $zipPath -Force
                $mail.Attachments.Add($zipPath) | Out-Null
            } else {
                foreach ($file in $selectedFiles) {
                    $mail.Attachments.Add($file) | Out-Null
                }
            }

            $mail.Display()
            $emailWindow.Close()
        } catch {
            [System.Windows.MessageBox]::Show("Failed to create email. Ensure Outlook is installed and accessible.`nError: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })

    $emailWindow.ShowDialog() | Out-Null
})

if ($RunConsole -or ($ExePath -and $Original -and $Modified -and $Output)) {
    Invoke-BulkCompare -ExePath $ExePath -OrigPath $Original -ModPath $Modified -OutDir $Output -Prefix $Prefix -Format $Format -Style $Style -Engine $ComparisonEngine -Mode $Mode -UseErrors $UseErrorsDialog -ShowVisible $ShowVisible -RedlineOnly $RedlinePagesOnly -TrackChanges $TrackChanges -AdvOptions $global:AdvSettings
    return
}

# Show the GUI Window
Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
$window.ShowDialog() | Out-Null
