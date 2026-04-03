# Cult Connector - Windows installer (graphical)
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$packageRoot = if ($PSScriptRoot) { [System.IO.Path]::GetFullPath($PSScriptRoot) } else { $null }
if ([string]::IsNullOrWhiteSpace($packageRoot)) {
    [void][System.Windows.Forms.MessageBox]::Show(
        'Installer could not determine its folder. Double-click Install.vbs next to the SupportFiles folder (do not run Install.ps1 directly).',
        'Cult Connector',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
Set-Location -LiteralPath $packageRoot

$srcItem = Get-ChildItem -LiteralPath $packageRoot -Filter 'Cult Connector*.jsxbin' -File -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $srcItem) {
    [void][System.Windows.Forms.MessageBox]::Show(
        'The Cult Connector panel file is missing from the SupportFiles folder. Copy the whole package again (Install.vbs and the SupportFiles folder together).',
        'Cult Connector',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
$sourcePath = $srcItem.FullName
$panelMenuName = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)

function Ensure-ScriptUIPanelsFolder {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) { return $true }
    try {
        New-Item -ItemType Directory -LiteralPath $Path -Force -ErrorAction Stop | Out-Null
        return (Test-Path -LiteralPath $Path)
    }
    catch {
        return $false
    }
}

function Get-ScriptUIPanelDirs {
    $dirs = New-Object System.Collections.Generic.List[string]
    $aeRoaming = Join-Path $env:APPDATA 'Adobe\After Effects'
    if (Test-Path -LiteralPath $aeRoaming) {
        Get-ChildItem -LiteralPath $aeRoaming -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Join-Path $_.FullName 'Scripts\ScriptUI Panels'
            [void](Ensure-ScriptUIPanelsFolder -Path $p)
            if (Test-Path -LiteralPath $p) { [void]$dirs.Add($p) }
        }
    }
    foreach ($rootspec in @('ProgramFiles', 'ProgramFiles(x86)')) {
        $base = [Environment]::GetEnvironmentVariable($rootspec)
        if (-not $base) { continue }
        $adobe = Join-Path $base 'Adobe'
        if (-not (Test-Path -LiteralPath $adobe)) { continue }
        Get-ChildItem -LiteralPath $adobe -Directory -Filter 'Adobe After Effects*' -ErrorAction SilentlyContinue |
            ForEach-Object {
                $p = Join-Path $_.FullName 'Support Files\Scripts\ScriptUI Panels'
                [void](Ensure-ScriptUIPanelsFolder -Path $p)
                if (Test-Path -LiteralPath $p) { [void]$dirs.Add($p) }
            }
    }
    $dirs | Select-Object -Unique
}

function Remove-CultConnectorJsxbinInDir {
    param([string]$DirectoryPath)
    if (-not (Test-Path -LiteralPath $DirectoryPath)) { return }
    Get-ChildItem -LiteralPath $DirectoryPath -Filter 'Cult Connector*.jsxbin' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Cult Connector  -  Installer'
$form.ClientSize = New-Object System.Drawing.Size(544, 668)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(244, 245, 250)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = [System.Windows.Forms.DockStyle]::Fill
$header.BackColor = [System.Drawing.Color]::FromArgb(22, 25, 40)
$header.Padding = New-Object System.Windows.Forms.Padding(28, 20, 28, 18)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Cult Connector'
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font('Segoe UI', 17, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(0, 4)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'After Effects  +  Crowdin'
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(168, 178, 205)
$subtitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(0, 44)

$header.Controls.AddRange(@($title, $subtitle))

# Main scroll area
$body = New-Object System.Windows.Forms.Panel
$body.Dock = [System.Windows.Forms.DockStyle]::Fill
$body.Padding = New-Object System.Windows.Forms.Padding(26, 18, 26, 8)
$body.BackColor = $form.BackColor
$body.AutoScroll = $true

$panelPre = New-Object System.Windows.Forms.Panel
$panelPre.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelPre.BackColor = [System.Drawing.Color]::Transparent
$panelPre.AutoScroll = $true

$card = New-Object System.Windows.Forms.FlowLayoutPanel
$card.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$card.WrapContents = $false
$card.AutoSize = $true
$card.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$card.Width = 468
$card.BackColor = [System.Drawing.Color]::White
$card.Padding = New-Object System.Windows.Forms.Padding(22, 20, 22, 20)
$card.BorderStyle = [System.Windows.Forms.BorderStyle]::None

$lblCardTitle = New-Object System.Windows.Forms.Label
$lblCardTitle.Text = 'Before you start'
$lblCardTitle.Font = New-Object System.Drawing.Font('Segoe UI', 11.5, [System.Drawing.FontStyle]::Bold)
$lblCardTitle.ForeColor = [System.Drawing.Color]::FromArgb(32, 34, 45)
$lblCardTitle.Width = 424
$lblCardTitle.AutoSize = $true
$lblCardTitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

$lblStepsBody = New-Object System.Windows.Forms.Label
$lblStepsBody.Text = @'
1. Save your After Effects project and fully quit After Effects (check the system tray for a background icon).

2. Click Install Cult Connector below. The installer removes any previous Cult Connector .jsxbin files in your ScriptUI Panels folders, then copies the new panel.

3. In After Effects (after you restart): choose Edit > Preferences > Scripting & Expressions. Enable Allow Scripts to Write Files and Access Network. Click OK. If you just turned this on, quit and reopen After Effects once.

4. When this installer says it is done: restart After Effects (or start it). Open the panel from the Window menu and look for the Cult Connector entry.
'@
$lblStepsBody.Font = New-Object System.Drawing.Font('Segoe UI', 9.35)
$lblStepsBody.ForeColor = [System.Drawing.Color]::FromArgb(68, 72, 88)
$lblStepsBody.Width = 424
$lblStepsBody.MaximumSize = New-Object System.Drawing.Size(424, 0)
$lblStepsBody.AutoSize = $true
$lblStepsBody.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 14)

$lblTip = New-Object System.Windows.Forms.Label
$lblTip.Text = 'If a folder under Program Files cannot be written, close this window, then right-click Install.vbs (the file next to SupportFiles) and choose Run as administrator. Your user Scripts folder is enough for most setups.'
$lblTip.Font = New-Object System.Drawing.Font('Segoe UI', 8.75, [System.Drawing.FontStyle]::Italic)
$lblTip.ForeColor = [System.Drawing.Color]::FromArgb(115, 120, 135)
$lblTip.Width = 424
$lblTip.MaximumSize = New-Object System.Drawing.Size(424, 0)
$lblTip.AutoSize = $true
$lblTip.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)

$card.Controls.AddRange(@($lblCardTitle, $lblStepsBody, $lblTip))
$panelPre.Controls.Add($card)

$panelPost = New-Object System.Windows.Forms.Panel
$panelPost.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelPost.Visible = $false
$panelPost.BackColor = [System.Drawing.Color]::Transparent
$panelPost.AutoScroll = $true

$successFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$successFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$successFlow.WrapContents = $false
$successFlow.AutoSize = $true
$successFlow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$successFlow.Width = 476
$successFlow.BackColor = [System.Drawing.Color]::FromArgb(232, 248, 237)
$successFlow.Padding = New-Object System.Windows.Forms.Padding(22, 20, 22, 22)

$lblOk = New-Object System.Windows.Forms.Label
$lblOk.Text = 'Installation complete'
$lblOk.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$lblOk.ForeColor = [System.Drawing.Color]::FromArgb(22, 95, 58)
$lblOk.AutoSize = $true
$lblOk.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

$successSteps = New-Object System.Windows.Forms.Label
$successSteps.Text = @"

Restart After Effects before using the panel.

Next steps:

1. Quit After Effects completely if it is still running, then open it again.

2. One-time in AE: Edit > Preferences > Scripting & Expressions > enable Allow Scripts to Write Files and Access Network. Restart After Effects if you changed this setting.

3. Open the Window menu and choose: $panelMenuName

4. Sign in from the panel when you are ready to connect to Crowdin.

You may close this installer when you have read the steps above.
"@
$successSteps.Font = New-Object System.Drawing.Font('Segoe UI', 9.35)
$successSteps.ForeColor = [System.Drawing.Color]::FromArgb(38, 88, 60)
$successSteps.Width = 424
$successSteps.MaximumSize = New-Object System.Drawing.Size(424, 0)
$successSteps.AutoSize = $true

$btnDone = New-Object System.Windows.Forms.Button
$btnDone.Text = 'Close installer'
$btnDone.Width = 160
$btnDone.Height = 38
$btnDone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDone.BackColor = [System.Drawing.Color]::FromArgb(34, 120, 75)
$btnDone.ForeColor = [System.Drawing.Color]::White
$btnDone.FlatAppearance.BorderSize = 0
$btnDone.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$btnDone.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnDone.Margin = New-Object System.Windows.Forms.Padding(0, 16, 0, 0)

$successFlow.Controls.AddRange(@($lblOk, $successSteps, $btnDone))
$panelPost.Controls.Add($successFlow)

$btnDone.Add_Click({ $form.Close() })

$body.Controls.Add($panelPost)
$body.Controls.Add($panelPre)

# Bottom bar
$bottomBar = New-Object System.Windows.Forms.Panel
$bottomBar.Dock = [System.Windows.Forms.DockStyle]::Fill
$bottomBar.Padding = New-Object System.Windows.Forms.Padding(26, 6, 26, 18)
$bottomBar.BackColor = $form.BackColor

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Width = 476
$progress.Height = 6
$progress.Location = New-Object System.Drawing.Point(26, 8)
$progress.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progress.Visible = $false

$statusLbl = New-Object System.Windows.Forms.Label
$statusLbl.Text = ''
$statusLbl.ForeColor = [System.Drawing.Color]::FromArgb(85, 90, 105)
$statusLbl.AutoSize = $true
$statusLbl.MaximumSize = New-Object System.Drawing.Size(476, 0)
$statusLbl.Location = New-Object System.Drawing.Point(26, 22)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = 'Install Cult Connector'
$btnInstall.Width = 216
$btnInstall.Height = 42
$btnInstall.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(86, 82, 212)
$btnInstall.ForeColor = [System.Drawing.Color]::White
$btnInstall.FlatAppearance.BorderSize = 0
$btnInstall.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnInstall.Location = New-Object System.Drawing.Point(26, 52)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Width = 108
$btnClose.Height = 42
$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClose.BackColor = [System.Drawing.Color]::FromArgb(232, 234, 242)
$btnClose.ForeColor = [System.Drawing.Color]::FromArgb(48, 50, 62)
$btnClose.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 204, 216)
$btnClose.Location = New-Object System.Drawing.Point(254, 52)

$bottomBar.Controls.AddRange(@($progress, $statusLbl, $btnInstall, $btnClose))

# Table layout: header / body / footer
$table = New-Object System.Windows.Forms.TableLayoutPanel
$table.Dock = [System.Windows.Forms.DockStyle]::Fill
$table.ColumnCount = 1
$table.RowCount = 3
$table.BackColor = $form.BackColor
[void]$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 88)))
[void]$table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 122)))
$table.Controls.Add($header, 0, 0)
$table.Controls.Add($body, 0, 1)
$table.Controls.Add($bottomBar, 0, 2)

$form.Controls.Add($table)

$btnClose.Add_Click({ $form.Close() })

$btnInstall.Add_Click({
    $targets = @(Get-ScriptUIPanelDirs)
    if ($targets.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "After Effects was not detected in the usual locations.`r`n`r`nInstall and launch After Effects once, then run this installer again.",
            'Cult Connector',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnInstall.Enabled = $false
    $btnClose.Enabled = $false
    $progress.Visible = $true
    $progress.Minimum = 0
    $progress.Maximum = [Math]::Max(1, $targets.Count)
    $progress.Value = 0

    $ok = 0
    $installState = @{ OptionalWriteFail = $false }

    foreach ($dest in $targets) {
        $statusLbl.Text = 'Installing...'
        [void]$form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        Remove-CultConnectorJsxbinInDir -DirectoryPath $dest
        try {
            Copy-Item -LiteralPath $sourcePath -Destination $dest -Force -ErrorAction Stop
            $ok++
        }
        catch {
            $installState['OptionalWriteFail'] = $true
        }
        $progress.Value = [Math]::Min($progress.Maximum, $progress.Value + 1)
        [System.Windows.Forms.Application]::DoEvents()
    }

    $progress.Visible = $false
    $btnClose.Enabled = $true

    if ($ok -eq 0) {
        $btnInstall.Enabled = $true
        $statusLbl.Text = 'Installation failed.'
        [void][System.Windows.Forms.MessageBox]::Show(
            "Could not copy the panel to any location.`r`n`r`nTry Run as administrator on Install.vbs, or copy the .jsxbin from this SupportFiles folder yourself.",
            'Cult Connector',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if ($installState['OptionalWriteFail']) {
        $statusLbl.Text = 'Done. If the panel is missing later, run Install.vbs as administrator.'
    }
    else {
        $statusLbl.Text = 'Done.'
    }

    $table.SuspendLayout()
    $bottomBar.Visible = $false
    $table.RowStyles[2] = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 0)
    $table.ResumeLayout($true)
    $panelPre.Visible = $false
    $panelPost.Visible = $true
    [void]$form.Refresh()
})

[void]$form.ShowDialog()
exit 0
