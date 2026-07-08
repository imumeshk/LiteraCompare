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
    [switch]$CategoryMatch,
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
    Client = ""; Prop = ""; TimeoutSeconds = 300
}

# Define the WPF UI via XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Litera Bulk Compare Tool" Height="760" Width="700" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
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
        <StackPanel Grid.Row="7" Grid.Column="1" Margin="5">
            <TextBox x:Name="txtStyle" VerticalContentAlignment="Center" ToolTip="Leave empty to use default style"/>
            <TextBlock x:Name="txtStyleNotice"
                       Margin="0,4,0,0"
                       Foreground="#7A5C00"
                       FontSize="11"
                       TextWrapping="Wrap"
                       Visibility="Collapsed"
                       Text="Match by document category uses Litera's default comparison styles. Custom style selection is disabled."/>
        </StackPanel>
        <Button x:Name="btnStyle" Grid.Row="7" Grid.Column="2" Content="Browse" Margin="5"/>

        <!-- Row 8: Options -->
        <Label Grid.Row="8" Grid.Column="0" Content="Options:" VerticalAlignment="Top"/>
        <StackPanel Grid.Row="8" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5" VerticalAlignment="Top">
            <CheckBox x:Name="chkErrors" Content="Show error dialog" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkVisible" Content="Visible window" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkRedlineOnly" Content="Redline pages only" Margin="0,0,15,0" VerticalAlignment="Center"/>
            <CheckBox x:Name="chkTrackChanges" Content="Track changes output" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Row 8b: Category Match Toggle -->
        <Label Grid.Row="9" Grid.Column="0" Content="" VerticalAlignment="Top"/>
        <StackPanel Grid.Row="9" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5,0,5,5" VerticalAlignment="Top">
            <CheckBox x:Name="chkCategoryMatch" VerticalAlignment="Center"
                      Content="Match by document category (Folder vs Folder, Auto engine)"
                      ToolTip="When enabled, Folder vs Folder modes ignore file extensions and pair files by comparison category (Document/Spreadsheet/Presentation/Image) in folder order, e.g. a .docx can pair with a .doc. Unmatched files in a category are skipped and logged. Requires the Auto (lcp_auto.exe) engine."/>
        </StackPanel>

        <!-- Row 10: Profile Save/Load -->
        <StackPanel Grid.Row="10" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="5,15,5,0" HorizontalAlignment="Right">
            <Button x:Name="btnLoadProfile" Content="Load Profile" Width="100" Height="30" Margin="0,0,10,0" Background="#607D8B" Foreground="White" FontWeight="Bold"/>
            <Button x:Name="btnSaveProfile" Content="Save Profile" Width="100" Height="30" Background="#607D8B" Foreground="White" FontWeight="Bold"/>
        </StackPanel>

        <!-- Row 11: Run, Advanced & Email Buttons -->
        <Grid Grid.Row="11" Grid.Column="0" Grid.ColumnSpan="3" Margin="5,15,5,5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="150"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="btnRun" Grid.Column="0" Content="START BATCH COMPARISON" Height="40" Margin="0,0,5,0" Background="#4CAF50" Foreground="White" FontWeight="Bold" FontSize="14"/>
            <Button x:Name="btnAdvanced" Grid.Column="1" Content="Advanced Options" Height="40" Margin="5,0,5,0" Background="#FF9800" Foreground="White" FontWeight="Bold" FontSize="14"/>
            <Button x:Name="btnEmail" Grid.Column="2" Content="Email Results" Height="40" Margin="5,0,0,0" Background="#2196F3" Foreground="White" FontWeight="Bold" FontSize="14"/>
        </Grid>

        <!-- Row 12: Progress -->
        <Grid Grid.Row="12" Grid.Column="0" Grid.ColumnSpan="3" Margin="5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <ProgressBar x:Name="pbCompare" Grid.Column="0" Height="20" Minimum="0" Maximum="100" Value="0" Margin="0,0,10,0"/>
            <TextBlock x:Name="txtProgress" Grid.Column="1" Text="Ready" VerticalAlignment="Center" Width="130" TextAlignment="Right"/>
        </Grid>

        <!-- Row 13: Logs -->
        <TextBox x:Name="txtLog" Grid.Row="13" Grid.Column="0" Grid.ColumnSpan="3" Margin="5" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas"/>
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
$txtStyleNotice = $window.FindName("txtStyleNotice")
$btnStyle    = $window.FindName("btnStyle")
$cmbEngine   = $window.FindName("cmbEngine")
$chkErrors   = $window.FindName("chkErrors")
$chkVisible  = $window.FindName("chkVisible")
$chkRedlineOnly = $window.FindName("chkRedlineOnly")
$chkTrackChanges = $window.FindName("chkTrackChanges")
$chkCategoryMatch = $window.FindName("chkCategoryMatch")
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
    param(
        [string]$message,
        [string]$UiMessage,
        [switch]$HideFromUi
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$timestamp] $message"
    Add-Content -Path $LogFilePath -Value $line
    if ($txtLog -ne $null) {
        if (-not $HideFromUi) {
            $displayLine = if ($PSBoundParameters.ContainsKey('UiMessage') -and $null -ne $UiMessage) {
                "[$timestamp] $UiMessage"
            } else {
                $line
            }
            $txtLog.AppendText("$displayLine`r`n")
            $txtLog.ScrollToEnd()
        }
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
function Show-MailtoFallback {
    <#
        .SYNOPSIS
        Opens a mailto: draft (works with New Outlook, classic Outlook, or whatever
        the default mail handler is) and reveals prepared attachments in Explorer,
        since mailto: links cannot carry attachments - that's a limitation of the
        protocol itself, not something any client or this script can work around.
    #>
    param(
        [string]$To,
        [string]$Cc,
        [string]$Subject,
        [string[]]$Attachments
    )

    $bodyNote = if ($Attachments.Count -gt 0) {
        "Attachments are prepared and ready - this draft was opened via your system's default mail handler, which cannot auto-attach files. An Explorer window showing the file(s) below will open after you click OK; drag them into this draft.`n`n" +
        (($Attachments | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "`n")
    } else { "" }

    $encodedSubject = [Uri]::EscapeDataString($Subject)
    $encodedBody = [Uri]::EscapeDataString($bodyNote)
    $mailtoUri = "mailto:$To"
    $queryParts = @()
    if ($Cc) { $queryParts += "cc=$([Uri]::EscapeDataString($Cc))" }
    if ($Subject) { $queryParts += "subject=$encodedSubject" }
    if ($bodyNote) { $queryParts += "body=$encodedBody" }
    if ($queryParts.Count -gt 0) { $mailtoUri += "?" + ($queryParts -join "&") }

    try {
        Start-Process $mailtoUri
    } catch {
        [System.Windows.MessageBox]::Show("Could not open a mail draft automatically. Your default mail app may not be configured.`nError: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    if ($Attachments.Count -gt 0) {
        [System.Windows.MessageBox]::Show("A mail draft was opened using your system's default mail app, since classic Outlook automation wasn't available on this machine (this is expected if New Outlook is active, since it doesn't support the COM automation classic Outlook used - it can also happen if Outlook isn't installed).`n`nClick OK to open the folder containing the attachment(s) to drag into the draft.", "Attach Files Manually", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        $folders = $Attachments | ForEach-Object { Split-Path -Path $_ -Parent } | Select-Object -Unique
        foreach ($folder in $folders) {
            Start-Process explorer.exe -ArgumentList $folder
        }
    }
}

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
        -1 { return "-1: Failed to start process" }
        -2 { return "-2: Timed out and was forcibly terminated" }
        default { return "$($code): Unknown exit code" }
    }
}

function Get-DocumentCategory {
    param([string]$extension)
    $ext = $extension.TrimStart('.').ToLowerInvariant()
    if (@('doc','docx','docm','dotm','rtf','pdf','wpd','htm','html','txt') -contains $ext) {
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
                '.doc', '.docx', '.docm', '.dotm', '.rtf', '.pdf', '.wpd', '.htm', '.html', '.txt',
                '.xls', '.xlsx', '.xlsm', '.xlsb',
                '.ppt', '.pps', '.pptx', '.pptm', '.ppsx', '.ppsm',
                '.png', '.bmp', '.jpg', '.jpeg'
            )
        }
        'Word' {
            return @('.doc', '.docx', '.docm', '.dotm', '.rtf', '.txt', '.pdf')
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
            return "Word / Text / PDF|*.doc;*.docx;*.docm;*.dotm;*.rtf;*.txt;*.pdf|All Files|*.*"
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
            return "Supported Compare Files|*.doc;*.docx;*.docm;*.dotm;*.rtf;*.pdf;*.wpd;*.htm;*.html;*.txt;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ppt;*.pps;*.pptx;*.pptm;*.ppsx;*.ppsm;*.png;*.bmp;*.jpg;*.jpeg|All Files|*.*"
        }
    }
}

function Test-IsWordAdvancedEligiblePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -Path $Path -PathType Leaf)) {
        return $false
    }

    $eligibleExtensions = @('.doc', '.docx', '.docm', '.dotm', '.txt', '.rtf', '.pdf')
    return $eligibleExtensions -contains ([System.IO.Path]::GetExtension($Path).ToLowerInvariant())
}

function Test-IsWordAdvancedEligiblePair {
    param(
        [string]$OriginalFile,
        [string]$ModifiedFile
    )

    $eligibleExtensions = @('.doc', '.docx', '.docm', '.dotm', '.txt', '.rtf', '.pdf')
    $origExt = [System.IO.Path]::GetExtension($OriginalFile).ToLowerInvariant()
    $modExt = [System.IO.Path]::GetExtension($ModifiedFile).ToLowerInvariant()
    return ($eligibleExtensions -contains $origExt) -and ($eligibleExtensions -contains $modExt)
}

function Get-CurrentMode {
    if ($rbSingle.IsChecked) { return 'Single' }
    if ($rbBulk.IsChecked) { return 'Bulk' }
    return 'Exact'
}

function Get-OptionSupportForContext {
    param(
        [string]$Engine,
        [string]$Mode,
        [string]$OriginalPath
    )

    $support = Get-EngineParameterSupport -Engine $Engine
    $canUseWordAdvancedInAuto = ($Engine -eq 'Auto' -and $Mode -eq 'Single' -and (Test-IsWordAdvancedEligiblePath -Path $OriginalPath))

    return @{
        SupportsUseErrors = $support.SupportsUseErrors
        SupportsVisible = $support.SupportsVisible
        SupportsRedlineOnly = ($support.SupportsRedlineOnly -or $canUseWordAdvancedInAuto)
        SupportsTrackChanges = ($support.SupportsTrackChanges -or $canUseWordAdvancedInAuto)
        SupportsOutputFile = $support.SupportsOutputFile
        SupportsStyle = $support.SupportsStyle
        SupportsCategoryMatch = ($Engine -eq 'Auto')
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

function Get-CategoryMatchedPairs {
    <#
        .SYNOPSIS
        Pairs files from two folders by comparison category (Document/Spreadsheet/
        Presentation/Image) rather than by filename or identical extension.

        .DESCRIPTION
        Files in each folder are grouped into categories using Get-DocumentCategory,
        preserving each folder's original directory-listing order within a category.
        Within each category, files are paired index-by-index (1st Document in
        Folder1 with 1st Document in Folder2, 2nd with 2nd, and so on). Any leftover
        files in a category that has no remaining counterpart on the other side are
        reported as unmatched/skipped rather than paired or dropped silently.

        Files whose extension does not resolve to a known category (per
        Get-DocumentCategory) are excluded entirely and reported as skipped, since
        there is no category to match them against.
    #>
    param(
        [Parameter(Mandatory)] [System.IO.FileInfo[]]$OriginalFiles,
        [Parameter(Mandatory)] [System.IO.FileInfo[]]$ModifiedFiles
    )

    $origByCategory = @{}
    foreach ($f in $OriginalFiles) {
        $cat = Get-DocumentCategory($f.Extension)
        if ($null -eq $cat) { continue }
        if (-not $origByCategory.ContainsKey($cat)) { $origByCategory[$cat] = New-Object System.Collections.ArrayList }
        [void]$origByCategory[$cat].Add($f)
    }

    $modByCategory = @{}
    foreach ($f in $ModifiedFiles) {
        $cat = Get-DocumentCategory($f.Extension)
        if ($null -eq $cat) { continue }
        if (-not $modByCategory.ContainsKey($cat)) { $modByCategory[$cat] = New-Object System.Collections.ArrayList }
        [void]$modByCategory[$cat].Add($f)
    }

    $pairs = New-Object System.Collections.ArrayList
    $unmatchedOriginal = New-Object System.Collections.ArrayList
    $unmatchedModified = New-Object System.Collections.ArrayList

    $allCategories = @(@($origByCategory.Keys) + @($modByCategory.Keys) | Select-Object -Unique)
    foreach ($cat in $allCategories) {
        $origList = if ($origByCategory.ContainsKey($cat)) { $origByCategory[$cat] } else { @() }
        $modList = if ($modByCategory.ContainsKey($cat)) { $modByCategory[$cat] } else { @() }
        $matchCount = [math]::Min($origList.Count, $modList.Count)

        for ($i = 0; $i -lt $matchCount; $i++) {
            [void]$pairs.Add(@{ Original = $origList[$i]; Modified = $modList[$i]; Category = $cat })
        }
        if ($origList.Count -gt $matchCount) {
            for ($i = $matchCount; $i -lt $origList.Count; $i++) { [void]$unmatchedOriginal.Add($origList[$i]) }
        }
        if ($modList.Count -gt $matchCount) {
            for ($i = $matchCount; $i -lt $modList.Count; $i++) { [void]$unmatchedModified.Add($modList[$i]) }
        }
    }

    return @{
        Pairs = $pairs
        UnmatchedOriginal = $unmatchedOriginal
        UnmatchedModified = $unmatchedModified
    }
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

    $support = Get-OptionSupportForContext -Engine $Engine -Mode (Get-CurrentMode) -OriginalPath $txtOriginal.Text

    $chkErrors.IsEnabled = $support.SupportsUseErrors
    if (-not $support.SupportsUseErrors) { $chkErrors.IsChecked = $false }

    $chkVisible.IsEnabled = $support.SupportsVisible
    if (-not $support.SupportsVisible) { $chkVisible.IsChecked = $false }

    $chkRedlineOnly.IsEnabled = $support.SupportsRedlineOnly
    if (-not $support.SupportsRedlineOnly) { $chkRedlineOnly.IsChecked = $false }

    $chkTrackChanges.IsEnabled = $support.SupportsTrackChanges
    if (-not $support.SupportsTrackChanges) { $chkTrackChanges.IsChecked = $false }

    $chkCategoryMatch.IsEnabled = $support.SupportsCategoryMatch
    if ($Engine -ne 'Auto') { $chkCategoryMatch.IsChecked = $false }

    $txtPrefix.IsEnabled = $support.SupportsOutputFile
    $cmbFormat.IsEnabled = $support.SupportsOutputFile
    Update-StyleControls -SupportsStyle $support.SupportsStyle -CategoryMatch ([bool]$chkCategoryMatch.IsChecked)
    Sync-OutputOptions
}

function Update-StyleControls {
    param(
        [bool]$SupportsStyle,
        [bool]$CategoryMatch
    )

    $useDefaultStyle = $SupportsStyle -and $CategoryMatch
    $txtStyle.IsEnabled = $SupportsStyle -and -not $useDefaultStyle
    $btnStyle.IsEnabled = $SupportsStyle -and -not $useDefaultStyle
    $txtStyleNotice.Visibility = if ($useDefaultStyle) { 'Visible' } else { 'Collapsed' }
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

        return @('.docx', '.doc', '.rtf', '.pdf', '.wpd', '.htm', '.html', '.xlsm', '.xlsx', '.xlsb', '.xls', '.pptm', '.pptx', '.ppt', '.ppsm', '.ppsx', '.pps', '.png', '.bmp', '.jpg', '.jpeg')
    }

    return Get-AllowedOutputFormats -engine $Engine -category $null
}

function Refresh-OutputFormatOptions {
    param(
        [string]$PreferredFormat
    )

    $engine = Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine)
    $formats = Get-AvailableOutputFormats -Engine $engine -OriginalPath $txtOriginal.Text
    if ($chkRedlineOnly.IsChecked) {
        $formats = @('.pdf')
    } elseif ($chkTrackChanges.IsChecked) {
        $formats = @('.docx', '.doc')
    }
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

function Sync-OutputOptions {
    if ($chkRedlineOnly.IsChecked) {
        if ($chkTrackChanges.IsChecked) {
            $chkTrackChanges.IsChecked = $false
            Write-Log "Track changes output was cleared because Redline pages only requires PDF output."
        }
        Refresh-OutputFormatOptions -PreferredFormat '.pdf'
        return
    }

    if ($chkTrackChanges.IsChecked) {
        Refresh-OutputFormatOptions -PreferredFormat $cmbFormat.Text
        return
    }

    Refresh-OutputFormatOptions -PreferredFormat $cmbFormat.Text
}

function Get-AllowedOutputFormats {
    param(
        [string]$engine,
        [string]$category
    )

    if ($engine -eq 'Auto') {
        switch ($category) {
            # Per the Litera Compare Command Line spec, lcp_auto.exe supports the
            # same file types for input and output within each comparison category.
            'Document' { return @('.docx', '.doc', '.rtf', '.pdf', '.wpd', '.htm', '.html') }
            'Spreadsheet' { return @('.xlsm', '.xlsx', '.xlsb', '.xls') }
            'Presentation' { return @('.pptm', '.pptx', '.ppt', '.ppsm', '.ppsx', '.pps') }
            'Image' { return @('.png', '.bmp', '.jpg', '.jpeg') }
            default { return @() }
        }
    } elseif ($engine -eq 'Word') {
        return @('.docx', '.doc', '.rtf', '.pdf')
    } elseif ($engine -eq 'PowerPoint') {
        return @('.pptm', '.pptx', '.ppt', '.ppsm', '.ppsx', '.pps')
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
            return @{ valid = $false; message = "Word engine supports only .doc, .docx, .docm, .dotm, .rtf, .txt, and .pdf comparisons." }
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
            $usingSpecialWordOutput = ($RedlineOnly -or $TrackChanges)
            if ($support.SupportsOutputFile -and $OutputFile -and -not $usingSpecialWordOutput) { $argsList += @('-auto', $OutputFile) }
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
            $usingSpecialPptOutput = ($RedlineOnly -or $TrackChanges)
            if ($support.SupportsOutputFile -and $OutputFile -and -not $usingSpecialPptOutput) { $argsList += @('-auto', $OutputFile) }
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

function Stop-ProcessTree {
    <#
        .SYNOPSIS
        Kills a process and any descendant processes it spawned, by walking
        Win32_Process parent/child relationships via CIM. Used when a Litera
        comparison times out, since lcp_auto.exe and friends can launch helper
        processes (e.g. Office conversion hosts) that would otherwise be left
        running after the parent is killed.
    #>
    param([Parameter(Mandatory)] [int]$ProcessId)

    try {
        $children = Get-CimInstance -ClassName Win32_Process -Filter "ParentProcessId = $ProcessId" -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            Stop-ProcessTree -ProcessId $child.ProcessId
        }
    } catch {
        # CIM may be unavailable in some environments; proceed to kill the parent regardless.
    }

    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    } catch {
        # Process may have already exited on its own between the timeout and the kill attempt.
    }
}

function Resolve-ExecutionPlan {
    param(
        [Parameter(Mandatory)] [string]$ExePath,
        [Parameter(Mandatory)] [string]$Engine,
        [Parameter(Mandatory)] [string]$OriginalFile,
        [Parameter(Mandatory)] [string]$ModifiedFile,
        [bool]$RedlineOnly,
        [bool]$TrackChanges
    )

    $resolvedExePath = $ExePath
    $resolvedEngine = $Engine

    if ($Engine -eq 'Auto' -and ($RedlineOnly -or $TrackChanges) -and (Test-IsWordAdvancedEligiblePair -OriginalFile $OriginalFile -ModifiedFile $ModifiedFile)) {
        $currentExeName = [System.IO.Path]::GetFileName($ExePath).ToLowerInvariant()
        if ($currentExeName -eq 'lcp_auto.exe') {
            $wordExePath = Join-Path -Path ([System.IO.Path]::GetDirectoryName($ExePath)) -ChildPath 'lcp_main.exe'
            if (Test-Path -Path $wordExePath -PathType Leaf) {
                $resolvedExePath = $wordExePath
                $resolvedEngine = 'Word'
            } else {
                Write-Log "WARNING: Redline pages only / Track changes requested with lcp_auto.exe, but companion lcp_main.exe was not found at '$wordExePath'. Proceeding without those Word-only options." -HideFromUi
            }
        }
    }

    return @{
        ExePath = $resolvedExePath
        Engine = $resolvedEngine
        CanUseWordAdvanced = ($resolvedEngine -eq 'Word' -and (Test-IsWordAdvancedEligiblePair -OriginalFile $OriginalFile -ModifiedFile $ModifiedFile))
    }
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

    $executionPlan = Resolve-ExecutionPlan -ExePath $ExePath -Engine $Engine -OriginalFile $OriginalFile -ModifiedFile $ModifiedFile -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges
    $effectiveExePath = $executionPlan.ExePath
    $effectiveEngine = $executionPlan.Engine
    $effectiveRedlineOnly = $RedlineOnly
    $effectiveTrackChanges = $TrackChanges

    if ($Engine -eq 'Auto' -and $effectiveEngine -ne 'Word') {
        if ($RedlineOnly) { $effectiveRedlineOnly = $false }
        if ($TrackChanges) { $effectiveTrackChanges = $false }
    }

    if (($RedlineOnly -or $TrackChanges) -and $Engine -eq 'Auto' -and $effectiveEngine -eq 'Word') {
        Write-Log "INFO: Using lcp_main.exe for this Word-style comparison so Redline pages only / Track changes can be applied." -HideFromUi
    }

    $argsList = Build-LiteraArguments -Engine $effectiveEngine -OriginalFile $OriginalFile -ModifiedFile $ModifiedFile -OutputFile $OutputFile -StyleFile $StyleFile -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $effectiveRedlineOnly -TrackChanges $effectiveTrackChanges -AdvOptions $AdvOptions

    # Start-Process -ArgumentList does not auto-quote elements containing spaces;
    # without quoting, a path like "Sample Docs\Test Org.docx" is split into multiple
    # tokens and the EXE receives a truncated/garbled path (often surfacing as exit
    # code 6 "File format is not supported"). Quote any token that needs it before
    # invocation, while keeping the unquoted array for clean logging.
    $quotedArgsList = $argsList | ForEach-Object {
        if ($_ -match '[\s]' -and $_ -notmatch '^".*"$') { '"{0}"' -f $_ } else { $_ }
    }

    $uiSummary = "Path: $OriginalFile vs Path: $ModifiedFile"
    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $uiSummary += "`r`nOutput: Path: $OutputFile"
    }
    Write-Log "Executing: $effectiveExePath $($argsList -join ' ')" -UiMessage $uiSummary

    $timeoutSeconds = if ($AdvOptions -and $AdvOptions.ContainsKey('TimeoutSeconds')) { [int]$AdvOptions.TimeoutSeconds } else { 300 }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $effectiveExePath
        $psi.Arguments = ($quotedArgsList -join ' ')
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()

        if ($timeoutSeconds -le 0) {
            # 0 means "no timeout" - wait indefinitely, matching original behavior.
            $proc.WaitForExit()
            $finished = $true
        } else {
            $finished = $proc.WaitForExit($timeoutSeconds * 1000)
        }

        if (-not $finished) {
            Write-Log "TIMEOUT: Comparison exceeded $timeoutSeconds second(s) and appears stuck. Terminating process (PID $($proc.Id)) and any child processes so the batch can continue."
            Stop-ProcessTree -ProcessId $proc.Id
            # Give the OS a brief moment to finish tearing down the tree before moving on.
            Start-Sleep -Milliseconds 500
            Write-Log "Result: Timed out (process forcibly terminated)"
            return -2
        }

        $exitMessage = Get-LiteraExitMessage $proc.ExitCode
        Write-Log "Result: $exitMessage"
        return $proc.ExitCode
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
        [bool]$CategoryMatch,
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

    $successCount = 0
    $failCount = 0
    $resolvedEngine = Get-EffectiveEngineFromExePath -ExePath $ExePath -FallbackEngine $Engine

    if ($resolvedEngine -ne $Engine) {
        Write-Log "Comparison engine '$Engine' overridden by executable selection. Using '$resolvedEngine' for $([System.IO.Path]::GetFileName($ExePath))."
        $Engine = $resolvedEngine
        $engineSupport = Get-EngineParameterSupport -Engine $Engine
    }

    if ($CategoryMatch -and $Engine -ne 'Auto') {
        Write-Log "WARNING: Match by document category is only supported with the Auto (lcp_auto.exe) engine, since only lcp_auto.exe can compare mismatched-but-compatible formats (e.g. .doc vs .docx) directly. Falling back to normal sequential matching for engine '$Engine'."
        $CategoryMatch = $false
    }

    if ($CategoryMatch -and -not [string]::IsNullOrWhiteSpace($Style)) {
        Write-Log "NOTE: Match by document category ignores custom comparison styles. Proceeding with Litera's default styles."
        $Style = $null
    }

    if ($CategoryMatch -and $Mode -eq 'Single') {
        Write-Log "NOTE: Match by document category does not apply to 'Single Original vs Multiple Modified' mode; it only changes folder-vs-folder matching."
    }

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
    Write-Log "Match by document category: $CategoryMatch"
    $effectiveTimeout = if ($AdvOptions -and $AdvOptions.ContainsKey('TimeoutSeconds')) { [int]$AdvOptions.TimeoutSeconds } else { 300 }
    Write-Log "Per-comparison timeout: $(if ($effectiveTimeout -le 0) { 'disabled (wait indefinitely)' } else { "$effectiveTimeout second(s)" })"

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
            if ($CategoryMatch) {
                Write-Log "Starting Mode: Folder vs Folder (Match by document category)..."
                $origFiles = @(Get-FilteredFilesForEngine -Path $OrigPath -Engine $Engine)
                $modFiles = @(Get-FilteredFilesForEngine -Path $ModPath -Engine $Engine)
                $matchResult = Get-CategoryMatchedPairs -OriginalFiles $origFiles -ModifiedFiles $modFiles
                $totalPairs = $matchResult.Pairs.Count
                if ($totalPairs -eq 0) { Write-Log "Warning: No files with a matching category were found between the two folders." }
                $currentIdx = 0
                foreach ($pair in $matchResult.Pairs) {
                    $currentIdx++
                    $oFile = $pair.Original
                    $mFile = $pair.Modified
                    Update-Progress -Current $currentIdx -Total $totalPairs -Message "Comparing file $currentIdx of $totalPairs"
                    $outputFile = Join-Path -Path $OutDir -ChildPath "$Prefix$($oFile.BaseName)_$($mFile.BaseName)$Format"
                    $validation = Validate-ComparisonPair -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -Engine $Engine
                    if (-not $validation.valid) {
                        Write-Log "SKIPPED: $($validation.message) for $($oFile.FullName) -> $($mFile.FullName)"
                        $failCount++
                        continue
                    }

                    Write-Log "Comparing full paths ($($pair.Category)): $($oFile.FullName) -> $($mFile.FullName)"
                    $exitCode = Invoke-LiteraProcess -ExePath $ExePath -Engine $Engine -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -OutputFile $outputFile -StyleFile $Style -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
                    if ($exitCode -eq 0) { $successCount++ } else { $failCount++ }
                }
                foreach ($oFile in $matchResult.UnmatchedOriginal) {
                    Write-Log "Skipped Original: No matching category file found for '$($oFile.Name)' (category: $(Get-DocumentCategory($oFile.Extension)))"
                }
                foreach ($mFile in $matchResult.UnmatchedModified) {
                    Write-Log "Skipped Modified: No matching category file found for '$($mFile.Name)' (category: $(Get-DocumentCategory($mFile.Extension)))"
                }
            } else {
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
        }
        'Exact' {
            if ($CategoryMatch) {
                Write-Log "Starting Mode: Folder vs Folder (Match by document category; filename-exact-match overridden by category toggle)..."
                $origFiles = @(Get-FilteredFilesForEngine -Path $OrigPath -Engine $Engine)
                $modFiles = @(Get-FilteredFilesForEngine -Path $ModPath -Engine $Engine)
                $matchResult = Get-CategoryMatchedPairs -OriginalFiles $origFiles -ModifiedFiles $modFiles
                $totalPairs = $matchResult.Pairs.Count
                if ($totalPairs -eq 0) { Write-Log "Warning: No files with a matching category were found between the two folders." }
                $currentIdx = 0
                foreach ($pair in $matchResult.Pairs) {
                    $currentIdx++
                    $oFile = $pair.Original
                    $mFile = $pair.Modified
                    Update-Progress -Current $currentIdx -Total $totalPairs -Message "Comparing file $currentIdx of $totalPairs"
                    $outputFile = Join-Path -Path $OutDir -ChildPath "$Prefix$($oFile.BaseName)_$($mFile.BaseName)$Format"
                    $validation = Validate-ComparisonPair -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -Engine $Engine
                    if (-not $validation.valid) {
                        Write-Log "SKIPPED: $($validation.message) for $($oFile.FullName) -> $($mFile.FullName)"
                        $failCount++
                        continue
                    }

                    Write-Log "Comparing full paths ($($pair.Category)): $($oFile.FullName) -> $($mFile.FullName)"
                    $exitCode = Invoke-LiteraProcess -ExePath $ExePath -Engine $Engine -OriginalFile $oFile.FullName -ModifiedFile $mFile.FullName -OutputFile $outputFile -StyleFile $Style -UseErrors $UseErrors -ShowVisible $ShowVisible -RedlineOnly $RedlineOnly -TrackChanges $TrackChanges -AdvOptions $AdvOptions
                    if ($exitCode -eq 0) { $successCount++ } else { $failCount++ }
                }
                foreach ($oFile in $matchResult.UnmatchedOriginal) {
                    Write-Log "Skipped Original: No matching category file found for '$($oFile.Name)' (category: $(Get-DocumentCategory($oFile.Extension)))"
                }
                foreach ($mFile in $matchResult.UnmatchedModified) {
                    Write-Log "Skipped Modified: No matching category file found for '$($mFile.Name)' (category: $(Get-DocumentCategory($mFile.Extension)))"
                }
            } else {
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

        Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine $engineName)
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

$chkRedlineOnly.Add_Click({
    if ($chkRedlineOnly.IsChecked) {
        if ($chkTrackChanges.IsChecked) {
            $chkTrackChanges.IsChecked = $false
            Write-Log "Track changes output was cleared because Redline pages only requires PDF output."
        }
        Refresh-OutputFormatOptions -PreferredFormat '.pdf'
    } else {
        Sync-OutputOptions
    }
})

$chkCategoryMatch.Add_Checked({
    Update-StyleControls -SupportsStyle (Get-EngineParameterSupport -Engine (Get-SelectedEngine)).SupportsStyle -CategoryMatch $true
    Write-Log "Match by document category enabled. Comparison will use Litera's default styles."
})

$chkCategoryMatch.Add_Unchecked({
    Update-StyleControls -SupportsStyle (Get-EngineParameterSupport -Engine (Get-SelectedEngine)).SupportsStyle -CategoryMatch $false
})

$chkTrackChanges.Add_Click({
    if ($chkTrackChanges.IsChecked) {
        if ($chkRedlineOnly.IsChecked) {
            $chkRedlineOnly.IsChecked = $false
            Write-Log "Redline pages only was cleared because Track changes output requires Word output."
        }
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
        Sync-OutputOptions
    }
})

$btnExe.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Executable Files|*.exe|All Files|*.*"
    if ($dialog.ShowDialog() -eq 'OK') {
        $txtExe.Text = $dialog.FileName
        Write-Log "Selected EXE: $($dialog.FileName)"
        Set-EngineSelectionFromExePath -ExePath $dialog.FileName
        Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
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
            Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
        }
    } else {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select Folder containing Original Documents"
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtOriginal.Text = $dialog.SelectedPath
            Write-Log "Selected Original Folder: $($dialog.SelectedPath)"
            Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
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

$txtExe.Add_TextChanged({
    Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
})

$txtOriginal.Add_TextChanged({
    Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
})

$rbSingle.Add_Checked({
    Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
})

$rbBulk.Add_Checked({
    Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
})

$rbExact.Add_Checked({
    Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
})

$btnAdvanced.Add_Click({
    [xml]$advXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Advanced Litera Options" Height="610" Width="500" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
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

        <TextBlock Grid.Row="8" Text="Reliability" FontWeight="Bold" Margin="0,0,0,5" />
        <Grid Grid.Row="9" Margin="0,0,0,10" VerticalAlignment="Top">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="80"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Label Grid.Column="0" Content="Per-comparison timeout (seconds):" VerticalAlignment="Center"/>
            <TextBox x:Name="advTxtTimeout" Grid.Column="1" Margin="2" VerticalContentAlignment="Center"
                     ToolTip="If a single comparison runs longer than this, the process is forcibly terminated and logged as a timeout so the batch can continue with the next file. Set to 0 to disable (wait indefinitely, original behavior)."/>
        </Grid>

        <StackPanel Grid.Row="10" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
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
    $advTxtTimeout = $advWindow.FindName("advTxtTimeout"); $advTxtTimeout.Text = [string]$global:AdvSettings.TimeoutSeconds
    
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

        $parsedTimeout = 0
        if ([int]::TryParse($advTxtTimeout.Text, [ref]$parsedTimeout) -and $parsedTimeout -ge 0) {
            $global:AdvSettings.TimeoutSeconds = $parsedTimeout
        } else {
            [System.Windows.MessageBox]::Show("Timeout must be a whole number of seconds (0 or greater). Keeping previous value of $($global:AdvSettings.TimeoutSeconds) seconds.", "Invalid Timeout", "OK", "Warning") | Out-Null
        }

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
            CategoryMatch = [bool]$chkCategoryMatch.IsChecked
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
                if ($profile.PSObject.Properties.Name -contains 'CategoryMatch') {
                    $chkCategoryMatch.IsChecked = [bool]$profile.CategoryMatch
                } else {
                    $chkCategoryMatch.IsChecked = $false
                }
                Update-UiForEngine -Engine (Get-SelectedEngine)
                
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
                    if ($profile.AdvSettings.PSObject.Properties.Name -contains 'TimeoutSeconds') {
                        $global:AdvSettings.TimeoutSeconds = [int]$profile.AdvSettings.TimeoutSeconds
                    }
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
    $catMatch  = [bool]$chkCategoryMatch.IsChecked

    $btnRun.IsEnabled = $false
    Invoke-BulkCompare -ExePath $exePath -OrigPath $origPath -ModPath $modPath -OutDir $outDir -Prefix $prefix -Format $format -Style $style -Engine $engine -Mode $mode -UseErrors $useErrors -ShowVisible $visible -RedlineOnly $redline -TrackChanges $track -CategoryMatch $catMatch -AdvOptions $global:AdvSettings
    $btnRun.IsEnabled = $true
})

$btnEmail.Add_Click({
    $outDir = $txtOutput.Text
    $origDir = $txtOriginal.Text
    $modDir = $txtModified.Text

    if ([string]::IsNullOrWhiteSpace($outDir) -or -not (Test-Path -Path $outDir)) {
        [System.Windows.MessageBox]::Show("Please select a valid Output Folder that exists.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    $outputFiles = @(Get-ChildItem -Path $outDir -File -ErrorAction SilentlyContinue)
    $originalFiles = if (-not [string]::IsNullOrWhiteSpace($origDir) -and (Test-Path -Path $origDir -PathType Container)) { @(Get-ChildItem -Path $origDir -File -ErrorAction SilentlyContinue) } else { @() }
    $modifiedFiles = if (-not [string]::IsNullOrWhiteSpace($modDir) -and (Test-Path -Path $modDir -PathType Container)) { @(Get-ChildItem -Path $modDir -File -ErrorAction SilentlyContinue) } else { @() }
    # Original/Modified can also be single files (Single mode), not just folders.
    if (-not [string]::IsNullOrWhiteSpace($origDir) -and (Test-Path -Path $origDir -PathType Leaf)) { $originalFiles = @(Get-Item -Path $origDir) }
    if (-not [string]::IsNullOrWhiteSpace($modDir) -and (Test-Path -Path $modDir -PathType Leaf)) { $modifiedFiles = @(Get-Item -Path $modDir) }

    if ($outputFiles.Count -eq 0 -and $originalFiles.Count -eq 0 -and $modifiedFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No files found in the Output, Original, or Modified locations.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # Define Email Dialog XAML
    [xml]$emailXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email Results" Height="640" Width="480" WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
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

        <TabControl Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Margin="5">
            <TabItem Header="Redline Output">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" BorderBrush="Gray" BorderThickness="1" Background="White">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="pnlOutputFiles" Margin="5"/>
                        </ScrollViewer>
                    </Border>
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,8,0,0">
                        <CheckBox x:Name="chkSelectAllOutput" Content="Select All" VerticalAlignment="Center" Margin="0,0,20,0"/>
                        <CheckBox x:Name="chkZipOutput" Content="Compress selected to ZIP" VerticalAlignment="Center"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            <TabItem Header="Original Document(s)">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" BorderBrush="Gray" BorderThickness="1" Background="White">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="pnlOriginalFiles" Margin="5"/>
                        </ScrollViewer>
                    </Border>
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,8,0,0">
                        <CheckBox x:Name="chkSelectAllOriginal" Content="Select All" VerticalAlignment="Center" Margin="0,0,20,0"/>
                        <CheckBox x:Name="chkZipOriginal" Content="Compress selected to ZIP" VerticalAlignment="Center"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            <TabItem Header="Modified Document(s)">
                <Grid Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" BorderBrush="Gray" BorderThickness="1" Background="White">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="pnlModifiedFiles" Margin="5"/>
                        </ScrollViewer>
                    </Border>
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,8,0,0">
                        <CheckBox x:Name="chkSelectAllModified" Content="Select All" VerticalAlignment="Center" Margin="0,0,20,0"/>
                        <CheckBox x:Name="chkZipModified" Content="Compress selected to ZIP" VerticalAlignment="Center"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>

        <StackPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
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
    $pnlOutputFiles = $emailWindow.FindName("pnlOutputFiles")
    $pnlOriginalFiles = $emailWindow.FindName("pnlOriginalFiles")
    $pnlModifiedFiles = $emailWindow.FindName("pnlModifiedFiles")
    $chkSelectAllOutput = $emailWindow.FindName("chkSelectAllOutput")
    $chkSelectAllOriginal = $emailWindow.FindName("chkSelectAllOriginal")
    $chkSelectAllModified = $emailWindow.FindName("chkSelectAllModified")
    $chkZipOutput = $emailWindow.FindName("chkZipOutput")
    $chkZipOriginal = $emailWindow.FindName("chkZipOriginal")
    $chkZipModified = $emailWindow.FindName("chkZipModified")
    $btnCreateEmail = $emailWindow.FindName("btnCreateEmail")
    $btnCancel = $emailWindow.FindName("btnCancel")

    # Populate each tab's file list and disable Select All / Zip when there's nothing to show.
    function Populate-EmailFileList {
        param($Panel, $FileList, $SelectAllCheckbox, $ZipCheckbox)

        if ($FileList.Count -eq 0) {
            $msg = New-Object System.Windows.Controls.TextBlock
            $msg.Text = "No files found in this location."
            $msg.Foreground = "Gray"
            $msg.Margin = "2"
            $Panel.Children.Add($msg) | Out-Null
            $SelectAllCheckbox.IsEnabled = $false
            $ZipCheckbox.IsEnabled = $false
            return
        }

        foreach ($file in $FileList) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $file.Name
            $cb.Tag = $file.FullName
            $cb.Margin = "0,2,0,2"
            $Panel.Children.Add($cb) | Out-Null
        }
    }

    Populate-EmailFileList -Panel $pnlOutputFiles -FileList $outputFiles -SelectAllCheckbox $chkSelectAllOutput -ZipCheckbox $chkZipOutput
    Populate-EmailFileList -Panel $pnlOriginalFiles -FileList $originalFiles -SelectAllCheckbox $chkSelectAllOriginal -ZipCheckbox $chkZipOriginal
    Populate-EmailFileList -Panel $pnlModifiedFiles -FileList $modifiedFiles -SelectAllCheckbox $chkSelectAllModified -ZipCheckbox $chkZipModified

    foreach ($pair in @(
        @($chkSelectAllOutput, $pnlOutputFiles),
        @($chkSelectAllOriginal, $pnlOriginalFiles),
        @($chkSelectAllModified, $pnlModifiedFiles)
    )) {
        $selectAllCb = $pair[0]; $panel = $pair[1]
        $selectAllCb.Add_Click({
            $isChecked = $this.IsChecked
            foreach ($cb in $this.Tag.Children) {
                if ($cb -is [System.Windows.Controls.CheckBox]) { $cb.IsChecked = $isChecked }
            }
        }.GetNewClosure())
        $selectAllCb.Tag = $panel
    }

    $btnCancel.Add_Click({
        $emailWindow.Close()
    })

    $btnCreateEmail.Add_Click({
        function Get-CheckedFiles($Panel) {
            $result = @()
            foreach ($cb in $Panel.Children) {
                if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.IsChecked) { $result += $cb.Tag }
            }
            return $result
        }

        $selectedOutput = Get-CheckedFiles $pnlOutputFiles
        $selectedOriginal = Get-CheckedFiles $pnlOriginalFiles
        $selectedModified = Get-CheckedFiles $pnlModifiedFiles

        if ($selectedOutput.Count -eq 0 -and $selectedOriginal.Count -eq 0 -and $selectedModified.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Please select at least one file to attach.", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        # Build the final attachment list once (zipping groups as requested), independent of
        # which email path (classic Outlook COM vs. mailto fallback) ends up being used below.
        # Each group (Redline/Original/Modified) becomes either one ZIP attachment (if that
        # group's "Compress" box is checked) or its individual files.
        $attachments = @()
        foreach ($group in @(
            @{ Files = $selectedOutput; Zip = [bool]$chkZipOutput.IsChecked; Label = "RedlineOutput" },
            @{ Files = $selectedOriginal; Zip = [bool]$chkZipOriginal.IsChecked; Label = "OriginalDocuments" },
            @{ Files = $selectedModified; Zip = [bool]$chkZipModified.IsChecked; Label = "ModifiedDocuments" }
        )) {
            if ($group.Files.Count -eq 0) { continue }

            if ($group.Zip) {
                $zipPath = Join-Path -Path $env:TEMP -ChildPath ("Litera_{0}_{1}.zip" -f $group.Label, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                try {
                    Compress-Archive -Path $group.Files -DestinationPath $zipPath -Force
                    $attachments += $zipPath
                } catch {
                    [System.Windows.MessageBox]::Show("Failed to create ZIP for $($group.Label): $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
            } else {
                $attachments += $group.Files
            }
        }

        $usedClassicOutlook = $false
        try {
            $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
            $mail = $outlook.CreateItem(0)
            $mail.To = $txtTo.Text
            $mail.CC = $txtCc.Text
            $mail.Subject = $txtSubject.Text

            foreach ($file in $attachments) {
                $mail.Attachments.Add($file) | Out-Null
            }

            $mail.Display()
            $usedClassicOutlook = $true
        } catch {
            Write-Log "Classic Outlook COM automation unavailable or failed ($($_.Exception.Message)). This is expected if New Outlook is the active client, since it does not support COM automation. Falling back to a mailto draft."
        }

        if ($usedClassicOutlook) {
            $emailWindow.Close()
            return
        }

        # Fallback path: New Outlook (and most other modern mail apps/webmail set as the
        # default handler) supports the mailto: protocol but mailto: links cannot carry
        # attachments - that limitation is part of the protocol itself, not this script.
        # So we open a draft via mailto: for To/CC/Subject, then reveal the prepared
        # attachment(s) in Explorer so the user can drag them into the draft in one step.
        Show-MailtoFallback -To $txtTo.Text -Cc $txtCc.Text -Subject $txtSubject.Text -Attachments $attachments
        $emailWindow.Close()
    })

    $emailWindow.ShowDialog() | Out-Null
})

if ($RunConsole -or ($ExePath -and $Original -and $Modified -and $Output)) {
    Invoke-BulkCompare -ExePath $ExePath -OrigPath $Original -ModPath $Modified -OutDir $Output -Prefix $Prefix -Format $Format -Style $Style -Engine $ComparisonEngine -Mode $Mode -UseErrors $UseErrorsDialog -ShowVisible $ShowVisible -RedlineOnly $RedlinePagesOnly -TrackChanges $TrackChanges -CategoryMatch $CategoryMatch -AdvOptions $global:AdvSettings
    return
}

# Show the GUI Window
Update-UiForEngine -Engine (Get-EffectiveEngineFromExePath -ExePath $txtExe.Text -FallbackEngine (Get-SelectedEngine))
$window.ShowDialog() | Out-Null
