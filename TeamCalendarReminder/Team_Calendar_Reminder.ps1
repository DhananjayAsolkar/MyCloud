<#
    Team_Calendar_Reminder.ps1
    ---------------------------------------------------------------------
    A simple Windows PowerShell GUI tool for team task reminders and a
    calendar-style checklist. Data is stored in tasks.json and change_log.csv
    next to this script, so if this whole folder lives inside a OneDrive or
    SharePoint synced folder, every team member sees the same task list.

    Run this file with: right-click -> "Run with PowerShell", or double
    click "Run_Team_Calendar_Reminder.bat" (included next to this script).

    No admin rights and no external modules/packages are required - this
    script only uses features already built into Windows (.NET WinForms).
    ---------------------------------------------------------------------
#>

param(
    # Optional: force a specific shared data folder. If not given, the
    # folder that contains this script is used (recommended: put this
    # whole folder inside your OneDrive/SharePoint synced location).
    [string]$DataPath
)

# -----------------------------------------------------------------------
# 0. Make sure we are running in STA mode (required by WinForms).
#    If someone runs this in a non-STA host, relaunch ourselves correctly.
# -----------------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    if ($PSCommandPath) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`""
        )
        exit
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# -----------------------------------------------------------------------
# 1. Work out where the shared data files live.
# -----------------------------------------------------------------------
if (-not $DataPath) { $DataPath = $PSScriptRoot }
if (-not $DataPath) { $DataPath = (Get-Location).Path }

$script:TasksFile      = Join-Path $DataPath 'tasks.json'
$script:ChangeLogFile  = Join-Path $DataPath 'change_log.csv'
$script:DayNotesFile   = Join-Path $DataPath 'day_notes.json'
$script:Tasks          = [System.Collections.ArrayList]::new()
$script:DayNotes       = @{}   # yyyy-MM-dd -> free-text note for that diary page
$script:LastLoadedWriteTime = $null
$script:LastLoadedDayNotesWriteTime = $null
$script:CurrentPageDate = (Get-Date).Date
$script:CurrentFilterName = 'All'   # All / Open / In Progress / Done / Assigned To Me

# Reminder pop-ups must never be shared between users, so this state is
# kept only in this process' memory (not written to tasks.json).
$script:SnoozeOverrides   = @{}   # TaskId -> DateTime the reminder should re-fire
$script:ShownReminderKeys = @{}   # "TaskId|yyyyMMddHHmm" already shown this session
$script:OpenReminderPopupCount = 0   # so multiple pop-ups at once don't stack on top of each other

# Suppresses the checklist's own change events while it is being rebuilt in
# code (so a refresh doesn't look like the user clicked a checkbox), and
# tracks the two view-mode toggles.
$script:SuppressDayChecklistEvents = $false
$script:IsCompactMode = $false
$script:IsLocked = $false

# =========================================================================
# 2. Helper functions - user/config handling
# =========================================================================

function Get-UserConfigPath {
    # Per-user settings are stored locally on each PC (NOT in the shared
    # folder) so that two people don't overwrite each other's name.
    return (Join-Path $env:APPDATA 'TeamCalendarReminder\user.json')
}

function Set-CurrentUser {
    $default = if ($script:CurrentUser) { $script:CurrentUser } else { $env:USERNAME }
    $name = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Enter your name.`nThis is shown as 'Created By' and used for the 'Assigned To Me' filter.",
        'Team Calendar Reminder - Your Name',
        $default
    )
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $default }
    $cfgPath = Get-UserConfigPath
    $cfgDir = Split-Path $cfgPath -Parent
    if (-not (Test-Path $cfgDir)) { New-Item -Path $cfgDir -ItemType Directory -Force | Out-Null }
    (@{ UserName = $name } | ConvertTo-Json) | Set-Content -Path $cfgPath -Encoding UTF8
    $script:CurrentUser = $name
    return $name
}

function Get-CurrentUser {
    $cfgPath = Get-UserConfigPath
    if (Test-Path $cfgPath) {
        try {
            $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
            if ($cfg.UserName) { return $cfg.UserName }
        } catch { }
    }
    return Set-CurrentUser
}

# =========================================================================
# 3. Helper functions - task data (load / save / JSON safety)
# =========================================================================

function New-TaskObject {
    param($Title, $Description, $TaskDate, $DueTime, $ReminderTime, $Priority, $AssignedTo, $Status, $CreatedBy)
    $now = Get-Date
    return [PSCustomObject]@{
        Id               = [guid]::NewGuid().ToString()
        Title            = $Title
        Description      = $Description
        TaskDate         = $TaskDate          # the calendar "page" this task currently shows on
        OriginalDate     = $TaskDate          # never changes - the date it was first created for
        DueTime          = $DueTime
        ReminderTime     = $ReminderTime
        Priority         = $Priority
        AssignedTo       = $AssignedTo
        Status           = $Status
        CreatedBy        = $CreatedBy
        CreatedDate      = $now.ToString('yyyy-MM-dd HH:mm:ss')
        CompletedDate    = if ($Status -eq 'Done') { $now.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        CarriedForward   = $false
        Note             = ''
        LastModifiedBy   = $CreatedBy
        LastModifiedDate = $now.ToString('yyyy-MM-dd HH:mm:ss')
    }
}

function ConvertTo-SafeJsonArray {
    # ConvertTo-Json can collapse a 1-item array into a plain object when
    # piped. Passing the array in with -InputObject avoids that, and this
    # extra check guarantees we always write a valid "[ ... ]" JSON array.
    param($InputArray)
    $arr = @($InputArray)
    if ($arr.Count -eq 0) { return '[]' }
    $json = ConvertTo-Json -InputObject $arr -Depth 6
    if ($json.TrimStart()[0] -ne '[') { $json = "[$json]" }
    return $json
}

function Load-Tasks {
    if (-not (Test-Path $script:TasksFile)) {
        $script:Tasks = [System.Collections.ArrayList]::new()
        return
    }
    try {
        $raw = Get-Content -Path $script:TasksFile -Raw -ErrorAction Stop
        $script:Tasks = [System.Collections.ArrayList]::new()
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $loaded = @($raw | ConvertFrom-Json)
            foreach ($t in $loaded) { [void]$script:Tasks.Add($t) }
        }
        $script:LastLoadedWriteTime = (Get-Item $script:TasksFile).LastWriteTime
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not read tasks.json. It may be corrupted or still syncing.`n`n$($_.Exception.Message)",
            'Load Error', 'OK', 'Warning') | Out-Null
        $script:Tasks = [System.Collections.ArrayList]::new()
    }
}

function Save-Tasks {
    for ($i = 0; $i -lt 5; $i++) {
        try {
            if (Test-Path $script:TasksFile) {
                Copy-Item -Path $script:TasksFile -Destination "$($script:TasksFile).bak" -Force -ErrorAction SilentlyContinue
            }
            $json = ConvertTo-SafeJsonArray -InputArray $script:Tasks
            $tmpFile = "$($script:TasksFile).tmp"
            Set-Content -Path $tmpFile -Value $json -Encoding UTF8 -ErrorAction Stop
            Move-Item -Path $tmpFile -Destination $script:TasksFile -Force -ErrorAction Stop
            $script:LastLoadedWriteTime = (Get-Item $script:TasksFile).LastWriteTime
            return $true
        } catch {
            Start-Sleep -Milliseconds 400
        }
    }
    [System.Windows.Forms.MessageBox]::Show(
        "Could not save tasks.json after several attempts.`nCheck that the shared folder is reachable and not locked by another program.",
        'Save Error', 'OK', 'Error') | Out-Null
    return $false
}

function Write-ChangeLog {
    param($Action, $TaskId, $TaskTitle, $Details)
    $entry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        User      = $script:CurrentUser
        Action    = $Action
        TaskId    = $TaskId
        TaskTitle = $TaskTitle
        Details   = $Details
    }
    try {
        $entry | Export-Csv -Path $script:ChangeLogFile -NoTypeInformation -Encoding UTF8 -Append
    } catch {
        Start-Sleep -Milliseconds 500
        try { $entry | Export-Csv -Path $script:ChangeLogFile -NoTypeInformation -Encoding UTF8 -Append } catch { }
    }
}

function Load-DayNotes {
    # One free-text note per diary date, shared the same way tasks.json is.
    $script:DayNotes = @{}
    if (-not (Test-Path $script:DayNotesFile)) { return }
    try {
        $raw = Get-Content -Path $script:DayNotesFile -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $obj = $raw | ConvertFrom-Json
            foreach ($prop in $obj.PSObject.Properties) { $script:DayNotes[$prop.Name] = $prop.Value }
        }
        $script:LastLoadedDayNotesWriteTime = (Get-Item $script:DayNotesFile).LastWriteTime
    } catch { }
}

function Save-DayNotes {
    for ($i = 0; $i -lt 5; $i++) {
        try {
            $json = $script:DayNotes | ConvertTo-Json -Depth 3
            $tmpFile = "$($script:DayNotesFile).tmp"
            Set-Content -Path $tmpFile -Value $json -Encoding UTF8 -ErrorAction Stop
            Move-Item -Path $tmpFile -Destination $script:DayNotesFile -Force -ErrorAction Stop
            return $true
        } catch {
            Start-Sleep -Milliseconds 400
        }
    }
    return $false
}

# =========================================================================
# 4. Helper functions - carry forward rule
# =========================================================================

function Invoke-CarryForward {
    # Any task that is NOT Done and whose page-date is before today gets
    # moved onto today's page. Completed tasks are never touched, so they
    # stay on the day they were finished. This only ever edits the existing
    # task record (never creates a copy), so nothing is duplicated.
    param([switch]$Silent)

    $today = (Get-Date).Date
    $todayStr = $today.ToString('yyyy-MM-dd')
    $count = 0

    foreach ($t in $script:Tasks) {
        if ($t.Status -eq 'Done') { continue }
        try { $taskDate = [datetime]::ParseExact($t.TaskDate, 'yyyy-MM-dd', $null) } catch { continue }
        if ($taskDate -lt $today) {
            $oldDate = $t.TaskDate
            $t.TaskDate = $todayStr
            $t.CarriedForward = $true
            $t.Note = "Carried forward from previous date ($oldDate)."
            $t.LastModifiedBy = 'System (Auto Carry Forward)'
            $t.LastModifiedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Write-ChangeLog -Action 'Task Carried Forward' -TaskId $t.Id -TaskTitle $t.Title -Details "From $oldDate to $todayStr"
            $count++
        }
    }

    if ($count -gt 0) { [void](Save-Tasks) }
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show("Carried forward $count task(s) to today.", 'Carry Forward Complete') | Out-Null
    }
    return $count
}

# =========================================================================
# 5. Small date helpers
# =========================================================================

function Get-WeekStart {
    param([datetime]$Date)
    $diff = ([int]$Date.DayOfWeek + 6) % 7   # Monday = start of week
    return $Date.Date.AddDays(-$diff)
}

function Parse-TaskDate {
    param($DateString)
    try { return [datetime]::ParseExact($DateString, 'yyyy-MM-dd', $null) } catch { return $null }
}

# =========================================================================
# 6. GUI - main window shell (clean floating diary page)
# =========================================================================

$script:MainForm = New-Object System.Windows.Forms.Form
$script:MainForm.Text = 'Team Calendar & Reminder'
$script:MainForm.Size = New-Object System.Drawing.Size(560, 680)
$script:MainForm.MinimumSize = New-Object System.Drawing.Size(420, 420)
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$script:MainForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$script:ToolTip = New-Object System.Windows.Forms.ToolTip

# ---- Slim top bar: day navigation, date, and a small right-hand cluster ---
$pnlTopBar = New-Object System.Windows.Forms.Panel
$pnlTopBar.Dock = 'Top'
$pnlTopBar.Height = 104
$pnlTopBar.BackColor = [System.Drawing.Color]::WhiteSmoke

$btnPrevDay = New-Object System.Windows.Forms.Button
$btnPrevDay.Text = '< Prev'
$btnPrevDay.Location = New-Object System.Drawing.Point(10, 8)
$btnPrevDay.Size = New-Object System.Drawing.Size(70, 28)

$btnToday = New-Object System.Windows.Forms.Button
$btnToday.Text = 'Today'
$btnToday.Location = New-Object System.Drawing.Point(88, 8)
$btnToday.Size = New-Object System.Drawing.Size(64, 28)

$btnNextDay = New-Object System.Windows.Forms.Button
$btnNextDay.Text = 'Next >'
$btnNextDay.Location = New-Object System.Drawing.Point(160, 8)
$btnNextDay.Size = New-Object System.Drawing.Size(70, 28)

$script:lblCurrentDate = New-Object System.Windows.Forms.Label
$script:lblCurrentDate.Location = New-Object System.Drawing.Point(10, 42)
$script:lblCurrentDate.Size = New-Object System.Drawing.Size(520, 30)
$script:lblCurrentDate.Anchor = 'Top,Left,Right'
$script:lblCurrentDate.TextAlign = 'MiddleCenter'
$script:lblCurrentDate.ForeColor = [System.Drawing.Color]::Black
$script:lblCurrentDate.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)

$btnAddOnDay = New-Object System.Windows.Forms.Button
$btnAddOnDay.Text = '+ Add Task'
$btnAddOnDay.Location = New-Object System.Drawing.Point(10, 78)
$btnAddOnDay.Size = New-Object System.Drawing.Size(110, 26)

$btnCompactToggle = New-Object System.Windows.Forms.Button
$btnCompactToggle.Text = 'Compact View'
$btnCompactToggle.Location = New-Object System.Drawing.Point(221, 78)
$btnCompactToggle.Size = New-Object System.Drawing.Size(95, 26)
$btnCompactToggle.Anchor = 'Top,Right'

$trackOpacity = New-Object System.Windows.Forms.TrackBar
$trackOpacity.Location = New-Object System.Drawing.Point(324, 74)
$trackOpacity.Size = New-Object System.Drawing.Size(90, 30)
$trackOpacity.Anchor = 'Top,Right'
$trackOpacity.Minimum = 40
$trackOpacity.Maximum = 100
$trackOpacity.Value = 100
$trackOpacity.TickFrequency = 20
$trackOpacity.TickStyle = 'None'

$btnLock = New-Object System.Windows.Forms.Button
$btnLock.Text = 'Lock'
$btnLock.Location = New-Object System.Drawing.Point(422, 78)
$btnLock.Size = New-Object System.Drawing.Size(45, 26)
$btnLock.Anchor = 'Top,Right'

$btnMore = New-Object System.Windows.Forms.Button
$btnMore.Text = 'More'
$btnMore.Location = New-Object System.Drawing.Point(475, 78)
$btnMore.Size = New-Object System.Drawing.Size(55, 26)
$btnMore.Anchor = 'Top,Right'

$script:ToolTip.SetToolTip($btnCompactToggle, 'Switch between compact and normal view')
$script:ToolTip.SetToolTip($trackOpacity, 'Window opacity - drag to make the window more see-through')
$script:ToolTip.SetToolTip($btnLock, 'Lock this window on top of other windows')
$script:ToolTip.SetToolTip($btnMore, 'More options: edit/delete task, filters, week view, future tasks, change log, export, change user')

$pnlTopBar.Controls.AddRange(@(
    $btnPrevDay, $btnToday, $btnNextDay, $script:lblCurrentDate,
    $btnAddOnDay, $btnCompactToggle, $trackOpacity, $btnLock, $btnMore
))

# ---- Slim bottom status bar ---------------------------------------------
$pnlStatusBar = New-Object System.Windows.Forms.Panel
$pnlStatusBar.Dock = 'Bottom'
$pnlStatusBar.Height = 24
$pnlStatusBar.BackColor = [System.Drawing.Color]::WhiteSmoke
$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Dock = 'Fill'
$script:lblStatus.TextAlign = 'MiddleLeft'
$script:lblStatus.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
$script:lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$pnlStatusBar.Controls.Add($script:lblStatus)

# ---- The diary page itself: a white card floating on the grey background --
$pnlBackground = New-Object System.Windows.Forms.Panel
$pnlBackground.Dock = 'Fill'
$pnlBackground.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$pnlBackground.Padding = New-Object System.Windows.Forms.Padding(12)

$pnlDiaryPage = New-Object System.Windows.Forms.Panel
$pnlDiaryPage.Dock = 'Fill'
$pnlDiaryPage.BackColor = [System.Drawing.Color]::White
$pnlDiaryPage.Padding = New-Object System.Windows.Forms.Padding(18)
$pnlDiaryPage.Add_Paint({
    param($sender, $e)
    $rect = New-Object System.Drawing.Rectangle(0, 0, $sender.Width - 1, $sender.Height - 1)
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gainsboro, 1)
    $e.Graphics.DrawRectangle($pen, $rect)
    $pen.Dispose()
})

$pnlDayNotes = New-Object System.Windows.Forms.Panel
$pnlDayNotes.Dock = 'Bottom'
$pnlDayNotes.Height = 110

$lblDayNotes = New-Object System.Windows.Forms.Label
$lblDayNotes.Text = 'Notes for this day'
$lblDayNotes.Dock = 'Top'
$lblDayNotes.Height = 22
$lblDayNotes.ForeColor = [System.Drawing.Color]::DimGray

$script:txtDayNotes = New-Object System.Windows.Forms.TextBox
$script:txtDayNotes.Dock = 'Fill'
$script:txtDayNotes.Multiline = $true
$script:txtDayNotes.ScrollBars = 'Vertical'
$script:txtDayNotes.BorderStyle = 'FixedSingle'
$script:txtDayNotes.ForeColor = [System.Drawing.Color]::Black
$script:txtDayNotes.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$pnlDayNotes.Controls.Add($lblDayNotes)
$pnlDayNotes.Controls.Add($script:txtDayNotes)

$lblChecklistHeader = New-Object System.Windows.Forms.Label
$lblChecklistHeader.Text = 'Tasks'
$lblChecklistHeader.Dock = 'Top'
$lblChecklistHeader.Height = 26
$lblChecklistHeader.ForeColor = [System.Drawing.Color]::DimGray
$lblChecklistHeader.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# The checklist: a headerless, borderless, gridline-free ListView with
# native checkboxes. Clicking a checkbox checks/unchecks (and selects) a
# task, clicking its text just selects it, and double-clicking opens the
# full Edit Task form. Columns give clean alignment without owner-draw.
$script:lvDayChecklist = New-Object System.Windows.Forms.ListView
$script:lvDayChecklist.Dock = 'Fill'
$script:lvDayChecklist.View = 'Details'
$script:lvDayChecklist.CheckBoxes = $true
$script:lvDayChecklist.FullRowSelect = $true
$script:lvDayChecklist.GridLines = $false
$script:lvDayChecklist.HeaderStyle = 'None'
$script:lvDayChecklist.MultiSelect = $false
$script:lvDayChecklist.BorderStyle = 'None'
$script:lvDayChecklist.BackColor = [System.Drawing.Color]::White
$script:lvDayChecklist.Font = New-Object System.Drawing.Font('Segoe UI', 11.5)
[void]$script:lvDayChecklist.Columns.Add('Task', 260)
[void]$script:lvDayChecklist.Columns.Add('Due', 65)
[void]$script:lvDayChecklist.Columns.Add('Priority', 65)
[void]$script:lvDayChecklist.Columns.Add('Assigned', 90)
$script:lvDayChecklist.Add_Resize({
    if ($script:lvDayChecklist.Columns.Count -lt 4) { return }
    $fixedWidth = $script:lvDayChecklist.Columns[1].Width + $script:lvDayChecklist.Columns[2].Width + $script:lvDayChecklist.Columns[3].Width
    $newWidth = $script:lvDayChecklist.ClientSize.Width - $fixedWidth
    if ($newWidth -gt 60) { $script:lvDayChecklist.Columns[0].Width = $newWidth }
})

$pnlDiaryPage.Controls.Add($pnlDayNotes)
$pnlDiaryPage.Controls.Add($lblChecklistHeader)
$pnlDiaryPage.Controls.Add($script:lvDayChecklist)

$pnlBackground.Controls.Add($pnlDiaryPage)

$script:MainForm.Controls.Add($pnlTopBar)
$script:MainForm.Controls.Add($pnlStatusBar)
$script:MainForm.Controls.Add($pnlBackground)

# =========================================================================
# 7. View refresh logic
# =========================================================================

function Add-StandardColumns {
    # Used by the reference pop-up windows (This Week / Future Tasks), which
    # keep the older multi-column grid look since they are secondary,
    # read-only lookups rather than the main diary page.
    param([System.Windows.Forms.ListView]$ListView)
    $ListView.View = 'Details'
    $ListView.FullRowSelect = $true
    $ListView.GridLines = $true
    $ListView.MultiSelect = $false
    [void]$ListView.Columns.Add('Date', 90)
    [void]$ListView.Columns.Add('Title', 170)
    [void]$ListView.Columns.Add('Priority', 65)
    [void]$ListView.Columns.Add('Due Time', 65)
    [void]$ListView.Columns.Add('Reminder', 65)
    [void]$ListView.Columns.Add('Assigned To', 110)
    [void]$ListView.Columns.Add('Status', 85)
    [void]$ListView.Columns.Add('Created By', 100)
    [void]$ListView.Columns.Add('Note', 230)
}

function Update-TaskListView {
    param([System.Windows.Forms.ListView]$ListView, [array]$Tasks)
    $ListView.BeginUpdate()
    $ListView.Items.Clear()
    $sorted = $Tasks | Sort-Object TaskDate, DueTime
    foreach ($t in $sorted) {
        $due = if ($t.DueTime) { $t.DueTime } else { '-' }
        $rem = if ($t.ReminderTime) { $t.ReminderTime } else { '-' }
        $item = New-Object System.Windows.Forms.ListViewItem($t.TaskDate)
        [void]$item.SubItems.Add($t.Title)
        [void]$item.SubItems.Add($t.Priority)
        [void]$item.SubItems.Add($due)
        [void]$item.SubItems.Add($rem)
        [void]$item.SubItems.Add($t.AssignedTo)
        [void]$item.SubItems.Add($t.Status)
        [void]$item.SubItems.Add($t.CreatedBy)
        [void]$item.SubItems.Add($t.Note)
        $item.Tag = $t.Id
        if ($t.Status -eq 'Done') {
            $item.ForeColor = [System.Drawing.Color]::Gray
        } elseif ($t.Priority -eq 'High') {
            $item.Font = New-Object System.Drawing.Font($ListView.Font, [System.Drawing.FontStyle]::Bold)
        }
        [void]$ListView.Items.Add($item)
    }
    $ListView.EndUpdate()
}

function Apply-StatusFilter {
    param($Tasks)
    switch ($script:CurrentFilterName) {
        'Open'            { return @($Tasks | Where-Object { $_.Status -eq 'Open' }) }
        'In Progress'     { return @($Tasks | Where-Object { $_.Status -eq 'In Progress' }) }
        'Done'            { return @($Tasks | Where-Object { $_.Status -eq 'Done' }) }
        'Assigned To Me'  { return @($Tasks | Where-Object { $_.AssignedTo -and ($_.AssignedTo.Trim().ToLower() -eq $script:CurrentUser.Trim().ToLower()) }) }
        default           { return @($Tasks) }
    }
}

function Save-CurrentDayNote {
    # Persists whatever is currently typed in the Notes box against the
    # diary page that's on screen right now. Call this before navigating
    # to a different day so a note is never silently lost.
    $dateKey = $script:CurrentPageDate.ToString('yyyy-MM-dd')
    $text = $script:txtDayNotes.Text
    if ([string]::IsNullOrWhiteSpace($text)) {
        if ($script:DayNotes.ContainsKey($dateKey)) {
            $script:DayNotes.Remove($dateKey)
            [void](Save-DayNotes)
        }
    } elseif ($script:DayNotes[$dateKey] -ne $text) {
        $script:DayNotes[$dateKey] = $text
        [void](Save-DayNotes)
    }
}

function Update-DayChecklist {
    # Renders the diary page's task list as a clean checklist: a checkbox,
    # the title, the due time, a small priority label, and (only when it
    # isn't you) who it's assigned to. Runs on every refresh, including
    # background ones triggered by unrelated actions elsewhere in the app,
    # so it must not disturb a task the user has selected or a note they
    # are mid-typing.
    $dateKey = $script:CurrentPageDate.ToString('yyyy-MM-dd')
    $dayTasks = @($script:Tasks | Where-Object { $_.TaskDate -eq $dateKey })
    $dayTasks = @(Apply-StatusFilter -Tasks $dayTasks | Sort-Object DueTime, Title)

    $previousSelectedId = if ($script:lvDayChecklist.SelectedItems.Count -gt 0) { $script:lvDayChecklist.SelectedItems[0].Tag } else { $null }

    $script:SuppressDayChecklistEvents = $true
    $script:lvDayChecklist.BeginUpdate()
    $script:lvDayChecklist.Items.Clear()
    foreach ($t in $dayTasks) {
        $isDone = ($t.Status -eq 'Done')
        $item = New-Object System.Windows.Forms.ListViewItem($t.Title)
        $item.Tag = $t.Id
        $item.Checked = $isDone
        $item.UseItemStyleForSubItems = $false

        $due = if ($t.DueTime) { $t.DueTime } else { '-' }
        [void]$item.SubItems.Add($due)
        [void]$item.SubItems.Add($t.Priority)
        $assignedText = if ($t.AssignedTo -and $t.AssignedTo.Trim().ToLower() -ne $script:CurrentUser.Trim().ToLower()) { $t.AssignedTo } else { '' }
        [void]$item.SubItems.Add($assignedText)

        $baseColor = if ($isDone) { [System.Drawing.Color]::Gray } else { [System.Drawing.Color]::Black }
        $item.SubItems[0].ForeColor = $baseColor
        $item.SubItems[1].ForeColor = $baseColor
        $item.SubItems[3].ForeColor = $baseColor
        $item.SubItems[2].ForeColor = if ($isDone) {
            [System.Drawing.Color]::Gray
        } else {
            switch ($t.Priority) {
                'High'  { [System.Drawing.Color]::Firebrick }
                'Low'   { [System.Drawing.Color]::Gray }
                default { [System.Drawing.Color]::Black }
            }
        }
        if (-not $isDone -and $t.Priority -eq 'High') {
            $item.Font = New-Object System.Drawing.Font($script:lvDayChecklist.Font, [System.Drawing.FontStyle]::Bold)
        }

        [void]$script:lvDayChecklist.Items.Add($item)
        if ($previousSelectedId -and $t.Id -eq $previousSelectedId) { $item.Selected = $true }
    }
    $script:lvDayChecklist.EndUpdate()
    $script:SuppressDayChecklistEvents = $false

    # Load this page's saved note - but never while the user has the box
    # focused (mid-typing), so an unrelated refresh can't erase their draft.
    if (-not $script:txtDayNotes.Focused) {
        $script:txtDayNotes.Text = if ($script:DayNotes.ContainsKey($dateKey)) { $script:DayNotes[$dateKey] } else { '' }
    }
}

function Refresh-AllViews {
    $script:lblCurrentDate.Text = $script:CurrentPageDate.ToString('dddd, dd MMMM yyyy')
    Update-DayChecklist
    $script:lblStatus.Text = "$($script:Tasks.Count) task(s) total  |  Logged in as $script:CurrentUser"
}

function Get-SelectedTask {
    if ($script:lvDayChecklist.SelectedItems.Count -eq 0) { return $null }
    $id = $script:lvDayChecklist.SelectedItems[0].Tag
    return ($script:Tasks | Where-Object { $_.Id -eq $id } | Select-Object -First 1)
}

# =========================================================================
# 8. Add / Edit Task dialog
# =========================================================================

function Show-TaskEditor {
    param(
        [PSCustomObject]$ExistingTask = $null,
        [datetime]$DefaultDate = (Get-Date).Date
    )
    $isEdit = $null -ne $ExistingTask

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($isEdit) { 'Edit Task' } else { 'Add Task' }
    $dlg.Size = New-Object System.Drawing.Size(440, 560)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.TopMost = $true

    $labelX = 15; $ctrlX = 130; $ctrlW = 275; $y = 15; $rowH = 34

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Task Title:*'
    $lblTitle.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblTitle.Size = New-Object System.Drawing.Size(110, 20)
    $txtTitle = New-Object System.Windows.Forms.TextBox
    $txtTitle.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $txtTitle.Size = New-Object System.Drawing.Size($ctrlW, 22)
    if ($isEdit) { $txtTitle.Text = $ExistingTask.Title }
    $dlg.Controls.AddRange(@($lblTitle, $txtTitle))
    $y += $rowH

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Description /`nComment:"
    $lblDesc.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblDesc.Size = New-Object System.Drawing.Size(110, 40)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $txtDesc.Size = New-Object System.Drawing.Size($ctrlW, 60)
    $txtDesc.Multiline = $true
    if ($isEdit) { $txtDesc.Text = $ExistingTask.Description }
    $dlg.Controls.AddRange(@($lblDesc, $txtDesc))
    $y += 70

    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = 'Task Date:*'
    $lblDate.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblDate.Size = New-Object System.Drawing.Size(110, 20)
    $dtpDate = New-Object System.Windows.Forms.DateTimePicker
    $dtpDate.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $dtpDate.Size = New-Object System.Drawing.Size($ctrlW, 22)
    $dtpDate.Format = 'Short'
    $dtpDate.Value = if ($isEdit) {
        # Falls back to today if TaskDate is somehow missing/unparseable,
        # so a damaged record can't crash the Edit dialog.
        $parsedExistingDate = Parse-TaskDate $ExistingTask.TaskDate
        if ($parsedExistingDate) { $parsedExistingDate } else { (Get-Date).Date }
    } else { $DefaultDate }
    $dlg.Controls.AddRange(@($lblDate, $dtpDate))
    $y += $rowH

    $lblDue = New-Object System.Windows.Forms.Label
    $lblDue.Text = 'Due Time:'
    $lblDue.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblDue.Size = New-Object System.Drawing.Size(110, 20)
    $dtpDue = New-Object System.Windows.Forms.DateTimePicker
    $dtpDue.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $dtpDue.Size = New-Object System.Drawing.Size(100, 22)
    $dtpDue.Format = 'Custom'
    $dtpDue.CustomFormat = 'HH:mm'
    $dtpDue.ShowUpDown = $true
    if ($isEdit -and $ExistingTask.DueTime) {
        try { $dtpDue.Value = [datetime]::ParseExact($ExistingTask.DueTime, 'HH:mm', $null) } catch { }
    }
    $dlg.Controls.AddRange(@($lblDue, $dtpDue))
    $y += $rowH

    $chkReminder = New-Object System.Windows.Forms.CheckBox
    $chkReminder.Text = 'Set Reminder:'
    $chkReminder.Location = New-Object System.Drawing.Point($labelX, $y)
    $chkReminder.Size = New-Object System.Drawing.Size(115, 22)
    $dtpReminder = New-Object System.Windows.Forms.DateTimePicker
    $dtpReminder.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $dtpReminder.Size = New-Object System.Drawing.Size(100, 22)
    $dtpReminder.Format = 'Custom'
    $dtpReminder.CustomFormat = 'HH:mm'
    $dtpReminder.ShowUpDown = $true
    $dtpReminder.Enabled = $false
    if ($isEdit -and $ExistingTask.ReminderTime) {
        $chkReminder.Checked = $true
        $dtpReminder.Enabled = $true
        try { $dtpReminder.Value = [datetime]::ParseExact($ExistingTask.ReminderTime, 'HH:mm', $null) } catch { }
    }
    $chkReminder.Add_CheckedChanged({ $dtpReminder.Enabled = $chkReminder.Checked })
    $dlg.Controls.AddRange(@($chkReminder, $dtpReminder))
    $y += $rowH

    $lblPriority = New-Object System.Windows.Forms.Label
    $lblPriority.Text = 'Priority:'
    $lblPriority.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblPriority.Size = New-Object System.Drawing.Size(110, 20)
    $cmbPriority = New-Object System.Windows.Forms.ComboBox
    $cmbPriority.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $cmbPriority.Size = New-Object System.Drawing.Size(150, 22)
    $cmbPriority.DropDownStyle = 'DropDownList'
    [void]$cmbPriority.Items.AddRange(@('Low', 'Medium', 'High'))
    $cmbPriority.SelectedItem = if ($isEdit) { $ExistingTask.Priority } else { 'Medium' }
    $dlg.Controls.AddRange(@($lblPriority, $cmbPriority))
    $y += $rowH

    $lblAssigned = New-Object System.Windows.Forms.Label
    $lblAssigned.Text = 'Assigned Person:'
    $lblAssigned.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblAssigned.Size = New-Object System.Drawing.Size(110, 20)
    $txtAssigned = New-Object System.Windows.Forms.TextBox
    $txtAssigned.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $txtAssigned.Size = New-Object System.Drawing.Size($ctrlW, 22)
    $txtAssigned.Text = if ($isEdit) { $ExistingTask.AssignedTo } else { $script:CurrentUser }
    $dlg.Controls.AddRange(@($lblAssigned, $txtAssigned))
    $y += $rowH

    $lblStatus2 = New-Object System.Windows.Forms.Label
    $lblStatus2.Text = 'Status:'
    $lblStatus2.Location = New-Object System.Drawing.Point($labelX, $y)
    $lblStatus2.Size = New-Object System.Drawing.Size(110, 20)
    $cmbStatus = New-Object System.Windows.Forms.ComboBox
    $cmbStatus.Location = New-Object System.Drawing.Point($ctrlX, $y - 2)
    $cmbStatus.Size = New-Object System.Drawing.Size(150, 22)
    $cmbStatus.DropDownStyle = 'DropDownList'
    [void]$cmbStatus.Items.AddRange(@('Open', 'In Progress', 'Done'))
    $cmbStatus.SelectedItem = if ($isEdit) { $ExistingTask.Status } else { 'Open' }
    $dlg.Controls.AddRange(@($lblStatus2, $cmbStatus))
    $y += $rowH

    if ($isEdit) {
        $lblMeta = New-Object System.Windows.Forms.Label
        $lblMeta.Location = New-Object System.Drawing.Point($labelX, $y)
        $lblMeta.Size = New-Object System.Drawing.Size(395, 40)
        $lblMeta.ForeColor = [System.Drawing.Color]::DimGray
        $lblMeta.Text = "Created by $($ExistingTask.CreatedBy) on $($ExistingTask.CreatedDate)"
        $dlg.Controls.Add($lblMeta)
        $y += 44
    }

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = 'OK'
    $btnOK.Location = New-Object System.Drawing.Point(150, $y)
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.DialogResult = 'OK'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(250, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = 'Cancel'

    $dlg.Controls.AddRange(@($btnOK, $btnCancel))
    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel
    $dlg.ClientSize = New-Object System.Drawing.Size(420, $y + 50)

    $btnOK.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtTitle.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Task Title is required.', 'Missing Information', 'OK', 'Warning') | Out-Null
            $dlg.DialogResult = 'None'
        }
    })

    $result = $dlg.ShowDialog($script:MainForm)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    return @{
        Title        = $txtTitle.Text.Trim()
        Description  = $txtDesc.Text.Trim()
        TaskDate     = $dtpDate.Value.ToString('yyyy-MM-dd')
        DueTime      = $dtpDue.Value.ToString('HH:mm')
        ReminderTime = if ($chkReminder.Checked) { $dtpReminder.Value.ToString('HH:mm') } else { '' }
        Priority     = $cmbPriority.SelectedItem
        AssignedTo   = $txtAssigned.Text.Trim()
        Status       = $cmbStatus.SelectedItem
    }
}

# =========================================================================
# 9. Reminder pop-up
# =========================================================================

function Show-ReminderPopup {
    param($Task)

    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = 'Task Reminder'
    $popup.Size = New-Object System.Drawing.Size(360, 240)
    $popup.FormBorderStyle = 'FixedToolWindow'
    $popup.TopMost = $true
    $popup.StartPosition = 'Manual'
    $workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

    # If several reminders fire around the same time, stack the pop-ups
    # upward instead of drawing them all in the exact same spot.
    $stackIndex = $script:OpenReminderPopupCount
    $script:OpenReminderPopupCount++
    $yPos = $workArea.Height - $popup.Height - 20 - ($stackIndex * ($popup.Height + 10))
    if ($yPos -lt 0) { $yPos = 20 }
    $popup.Location = New-Object System.Drawing.Point(($workArea.Width - $popup.Width - 20), $yPos)
    $popup.Add_FormClosed({ $script:OpenReminderPopupCount-- })

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(12, 12)
    $lblInfo.Size = New-Object System.Drawing.Size(330, 130)
    $lblInfo.Text = "Task: $($Task.Title)`r`nDue Time: $($Task.DueTime)`r`nAssigned To: $($Task.AssignedTo)`r`nPriority: $($Task.Priority)`r`n`r`nComment: $($Task.Description)"
    $popup.Controls.Add($lblInfo)

    $btnSnooze5 = New-Object System.Windows.Forms.Button
    $btnSnooze5.Text = 'Snooze 5m'
    $btnSnooze5.Location = New-Object System.Drawing.Point(12, 155)
    $btnSnooze5.Size = New-Object System.Drawing.Size(78, 28)

    $btnSnooze15 = New-Object System.Windows.Forms.Button
    $btnSnooze15.Text = 'Snooze 15m'
    $btnSnooze15.Location = New-Object System.Drawing.Point(96, 155)
    $btnSnooze15.Size = New-Object System.Drawing.Size(78, 28)

    $btnSnooze30 = New-Object System.Windows.Forms.Button
    $btnSnooze30.Text = 'Snooze 30m'
    $btnSnooze30.Location = New-Object System.Drawing.Point(180, 155)
    $btnSnooze30.Size = New-Object System.Drawing.Size(78, 28)

    $btnDismiss = New-Object System.Windows.Forms.Button
    $btnDismiss.Text = 'Dismiss'
    $btnDismiss.Location = New-Object System.Drawing.Point(264, 155)
    $btnDismiss.Size = New-Object System.Drawing.Size(78, 28)

    $btnSnooze5.Add_Click({ $script:SnoozeOverrides[$Task.Id] = (Get-Date).AddMinutes(5); $popup.Close() })
    $btnSnooze15.Add_Click({ $script:SnoozeOverrides[$Task.Id] = (Get-Date).AddMinutes(15); $popup.Close() })
    $btnSnooze30.Add_Click({ $script:SnoozeOverrides[$Task.Id] = (Get-Date).AddMinutes(30); $popup.Close() })
    $btnDismiss.Add_Click({ $popup.Close() })

    $popup.Controls.AddRange(@($btnSnooze5, $btnSnooze15, $btnSnooze30, $btnDismiss))
    $popup.Show()
}

# =========================================================================
# 10. Reference pop-up windows (This Week / Future Tasks / Change Log)
#     Kept out of the main diary page entirely - reached only via "More".
# =========================================================================

function Show-ReferenceListPopup {
    param([string]$Title, [array]$Tasks, [string]$RangeLabel = '')

    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = $Title
    $popup.Size = New-Object System.Drawing.Size(760, 480)
    $popup.MinimumSize = New-Object System.Drawing.Size(500, 300)
    $popup.StartPosition = 'CenterParent'

    if ($RangeLabel) {
        $lblRange = New-Object System.Windows.Forms.Label
        $lblRange.Text = $RangeLabel
        $lblRange.Dock = 'Top'
        $lblRange.Height = 30
        $lblRange.TextAlign = 'MiddleLeft'
        $lblRange.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
        $lblRange.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $popup.Controls.Add($lblRange)
    }

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    Add-StandardColumns -ListView $lv
    Update-TaskListView -ListView $lv -Tasks $Tasks
    $popup.Controls.Add($lv)

    [void]$popup.ShowDialog($script:MainForm)
}

function Show-WeekPopup {
    $weekStart = Get-WeekStart -Date (Get-Date).Date
    $weekEnd = $weekStart.AddDays(6)
    $weekTasks = @($script:Tasks | Where-Object {
        $d = Parse-TaskDate $_.TaskDate
        $d -and $d -ge $weekStart -and $d -le $weekEnd
    })
    $rangeLabel = "Week: $($weekStart.ToString('dd MMM yyyy')) - $($weekEnd.ToString('dd MMM yyyy'))"
    Show-ReferenceListPopup -Title 'This Week' -Tasks (Apply-StatusFilter -Tasks $weekTasks) -RangeLabel $rangeLabel
}

function Show-FuturePopup {
    $futureTasks = @($script:Tasks | Where-Object {
        $d = Parse-TaskDate $_.TaskDate
        $d -and $d -gt (Get-Date).Date
    })
    Show-ReferenceListPopup -Title 'Future Tasks' -Tasks (Apply-StatusFilter -Tasks $futureTasks)
}

function Show-ChangeLogPopup {
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = 'Change Log'
    $popup.Size = New-Object System.Drawing.Size(820, 480)
    $popup.MinimumSize = New-Object System.Drawing.Size(500, 300)
    $popup.StartPosition = 'CenterParent'

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    [void]$lv.Columns.Add('Timestamp', 140)
    [void]$lv.Columns.Add('User', 110)
    [void]$lv.Columns.Add('Action', 150)
    [void]$lv.Columns.Add('Task Title', 180)
    [void]$lv.Columns.Add('Details', 260)

    if (Test-Path $script:ChangeLogFile) {
        try {
            $entries = @(Import-Csv -Path $script:ChangeLogFile) | Sort-Object { [datetime]$_.Timestamp } -Descending
            foreach ($e in $entries) {
                $item = New-Object System.Windows.Forms.ListViewItem($e.Timestamp)
                [void]$item.SubItems.Add($e.User)
                [void]$item.SubItems.Add($e.Action)
                [void]$item.SubItems.Add($e.TaskTitle)
                [void]$item.SubItems.Add($e.Details)
                [void]$lv.Items.Add($item)
            }
        } catch { }
    }
    $popup.Controls.Add($lv)
    [void]$popup.ShowDialog($script:MainForm)
}

# =========================================================================
# 11. Actions shared between the diary page and the "More" menu
# =========================================================================

function Invoke-AddTaskFlow {
    $result = Show-TaskEditor -DefaultDate $script:CurrentPageDate
    if ($result) {
        $newTask = New-TaskObject -Title $result.Title -Description $result.Description -TaskDate $result.TaskDate `
            -DueTime $result.DueTime -ReminderTime $result.ReminderTime -Priority $result.Priority `
            -AssignedTo $result.AssignedTo -Status $result.Status -CreatedBy $script:CurrentUser
        [void]$script:Tasks.Add($newTask)
        Write-ChangeLog -Action 'Task Created' -TaskId $newTask.Id -TaskTitle $newTask.Title -Details "Date: $($newTask.TaskDate), Assigned: $($newTask.AssignedTo)"
        [void](Save-Tasks)
        Refresh-AllViews
    }
}

function Invoke-EditSelectedTask {
    $task = Get-SelectedTask
    if (-not $task) { [System.Windows.Forms.MessageBox]::Show('Please select a task first.', 'No Task Selected') | Out-Null; return }
    $result = Show-TaskEditor -ExistingTask $task
    if ($result) {
        $dateChanged = ($task.TaskDate -ne $result.TaskDate)
        $task.Title        = $result.Title
        $task.Description  = $result.Description
        $task.TaskDate      = $result.TaskDate
        $task.DueTime       = $result.DueTime
        $task.ReminderTime  = $result.ReminderTime
        $task.Priority      = $result.Priority
        $task.AssignedTo    = $result.AssignedTo
        $task.Status        = $result.Status
        $task.LastModifiedBy = $script:CurrentUser
        $task.LastModifiedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        if ($dateChanged) {
            # A manual reschedule is different from an automatic carry-forward.
            $task.CarriedForward = $false
            $task.Note = ''
        }
        if ($result.Status -eq 'Done' -and -not $task.CompletedDate) {
            $task.CompletedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        } elseif ($result.Status -ne 'Done') {
            $task.CompletedDate = ''
        }
        Write-ChangeLog -Action 'Task Edited' -TaskId $task.Id -TaskTitle $task.Title -Details "Status: $($task.Status), Date: $($task.TaskDate)"
        [void](Save-Tasks)
        Refresh-AllViews
    }
}

function Invoke-MarkSelectedTaskDone {
    $task = Get-SelectedTask
    if (-not $task) { [System.Windows.Forms.MessageBox]::Show('Please select a task first.', 'No Task Selected') | Out-Null; return }
    if ($task.Status -eq 'Done') { [System.Windows.Forms.MessageBox]::Show('That task is already marked Done.', 'Already Done') | Out-Null; return }
    $task.Status = 'Done'
    $task.CompletedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $task.LastModifiedBy = $script:CurrentUser
    $task.LastModifiedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-ChangeLog -Action 'Task Completed' -TaskId $task.Id -TaskTitle $task.Title -Details "Completed on $($task.TaskDate)"
    [void](Save-Tasks)
    Refresh-AllViews
}

function Invoke-DeleteSelectedTask {
    $task = Get-SelectedTask
    if (-not $task) { [System.Windows.Forms.MessageBox]::Show('Please select a task first.', 'No Task Selected') | Out-Null; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show("Delete task '$($task.Title)'?", 'Confirm Delete', 'YesNo', 'Question')
    if ($confirm -eq 'Yes') {
        $keep = [System.Collections.ArrayList]::new()
        foreach ($t in $script:Tasks) { if ($t.Id -ne $task.Id) { [void]$keep.Add($t) } }
        $script:Tasks = $keep
        Write-ChangeLog -Action 'Task Deleted' -TaskId $task.Id -TaskTitle $task.Title -Details "Deleted by $script:CurrentUser"
        [void](Save-Tasks)
        Refresh-AllViews
    }
}

function Invoke-RefreshNow {
    Load-Tasks
    if (-not $script:txtDayNotes.Focused) { Load-DayNotes }
    [void](Invoke-CarryForward -Silent)
    Refresh-AllViews
    $script:lblStatus.Text = "Refreshed at $((Get-Date).ToString('HH:mm:ss'))"
}

function Invoke-CarryForwardNow {
    [void](Invoke-CarryForward)
    Refresh-AllViews
}

function Invoke-ExportTasksToCsv {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.InitialDirectory = $DataPath
    $sfd.FileName = "TasksExport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    $sfd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:Tasks | Select-Object Title, Description, TaskDate, DueTime, ReminderTime, Priority, AssignedTo, Status, CreatedBy, CreatedDate, CompletedDate, CarriedForward, Note |
                Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exported $($script:Tasks.Count) task(s) to:`n$($sfd.FileName)", 'Export Complete') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", 'Export Error', 'OK', 'Error') | Out-Null
        }
    }
}

function Invoke-ChangeUserFlow {
    Set-CurrentUser | Out-Null
    Refresh-AllViews
}

# =========================================================================
# 12. The "More" menu - everything that isn't on the main diary page
# =========================================================================

$script:MoreMenu = New-Object System.Windows.Forms.ContextMenuStrip

$miEdit = New-Object System.Windows.Forms.ToolStripMenuItem 'Edit Task'
$miMarkDone = New-Object System.Windows.Forms.ToolStripMenuItem 'Mark as Done'
$miDelete = New-Object System.Windows.Forms.ToolStripMenuItem 'Delete Task'
$miEdit.Add_Click({ Invoke-EditSelectedTask })
$miMarkDone.Add_Click({ Invoke-MarkSelectedTaskDone })
$miDelete.Add_Click({ Invoke-DeleteSelectedTask })

$miFilter = New-Object System.Windows.Forms.ToolStripMenuItem 'Filter'
$script:FilterMenuItems = @{}
foreach ($filterName in @('All', 'Open', 'In Progress', 'Done', 'Assigned To Me')) {
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem $filterName
    $mi.CheckOnClick = $false
    $mi.Checked = ($filterName -eq 'All')
    $mi.Add_Click({
        param($sender, $e)
        foreach ($kv in $script:FilterMenuItems.GetEnumerator()) { $kv.Value.Checked = $false }
        $sender.Checked = $true
        $script:CurrentFilterName = $sender.Text
        Refresh-AllViews
    })
    [void]$miFilter.DropDownItems.Add($mi)
    $script:FilterMenuItems[$filterName] = $mi
}

$miWeek = New-Object System.Windows.Forms.ToolStripMenuItem 'This Week...'
$miFuture = New-Object System.Windows.Forms.ToolStripMenuItem 'Future Tasks...'
$miChangeLog = New-Object System.Windows.Forms.ToolStripMenuItem 'Change Log...'
$miWeek.Add_Click({ Show-WeekPopup })
$miFuture.Add_Click({ Show-FuturePopup })
$miChangeLog.Add_Click({ Show-ChangeLogPopup })

$miRefresh = New-Object System.Windows.Forms.ToolStripMenuItem 'Refresh Now'
$miCarryForward = New-Object System.Windows.Forms.ToolStripMenuItem 'Carry Forward Open Tasks Now'
$miExport = New-Object System.Windows.Forms.ToolStripMenuItem 'Export to CSV...'
$miRefresh.Add_Click({ Invoke-RefreshNow })
$miCarryForward.Add_Click({ Invoke-CarryForwardNow })
$miExport.Add_Click({ Invoke-ExportTasksToCsv })

$miChangeUser = New-Object System.Windows.Forms.ToolStripMenuItem 'Change User...'
$miChangeUser.Add_Click({ Invoke-ChangeUserFlow })

[void]$script:MoreMenu.Items.Add($miEdit)
[void]$script:MoreMenu.Items.Add($miMarkDone)
[void]$script:MoreMenu.Items.Add($miDelete)
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$script:MoreMenu.Items.Add($miFilter)
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$script:MoreMenu.Items.Add($miWeek)
[void]$script:MoreMenu.Items.Add($miFuture)
[void]$script:MoreMenu.Items.Add($miChangeLog)
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$script:MoreMenu.Items.Add($miRefresh)
[void]$script:MoreMenu.Items.Add($miCarryForward)
[void]$script:MoreMenu.Items.Add($miExport)
[void]$script:MoreMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$script:MoreMenu.Items.Add($miChangeUser)

# =========================================================================
# 13. Event wiring
# =========================================================================

$btnAddOnDay.Add_Click({ Invoke-AddTaskFlow })

$btnPrevDay.Add_Click({
    Save-CurrentDayNote
    $script:CurrentPageDate = $script:CurrentPageDate.AddDays(-1)
    Refresh-AllViews
})
$btnNextDay.Add_Click({
    Save-CurrentDayNote
    $script:CurrentPageDate = $script:CurrentPageDate.AddDays(1)
    Refresh-AllViews
})
$btnToday.Add_Click({
    Save-CurrentDayNote
    $script:CurrentPageDate = (Get-Date).Date
    Refresh-AllViews
})

$script:txtDayNotes.Add_Leave({ Save-CurrentDayNote })

# Click the checkbox = check/uncheck (mark Done / reopen), and it also
# selects the row. Click the task's text = select only. Double-click = open
# the full Edit Task form.
$script:lvDayChecklist.Add_ItemChecked({
    param($sender, $e)
    if ($script:SuppressDayChecklistEvents) { return }
    $taskId = $e.Item.Tag
    $task = $script:Tasks | Where-Object { $_.Id -eq $taskId } | Select-Object -First 1
    if (-not $task) { return }

    if ($e.Item.Checked) {
        $task.Status = 'Done'
        $task.CompletedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Write-ChangeLog -Action 'Task Completed' -TaskId $task.Id -TaskTitle $task.Title -Details "Completed on $($task.TaskDate) (checklist)"
    } else {
        $task.Status = 'Open'
        $task.CompletedDate = ''
        Write-ChangeLog -Action 'Task Reopened' -TaskId $task.Id -TaskTitle $task.Title -Details "Reopened on $($task.TaskDate) (checklist)"
    }
    $task.LastModifiedBy = $script:CurrentUser
    $task.LastModifiedDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [void](Save-Tasks)

    # Rebuilding the checklist from inside its own ItemChecked event is not
    # safe (the click hasn't finished being processed yet), so the refresh
    # is queued to run right after this event handler returns.
    $script:MainForm.BeginInvoke([Action]{ Refresh-AllViews }) | Out-Null
})

$script:lvDayChecklist.Add_DoubleClick({ Invoke-EditSelectedTask })

# ---- Compact / Normal view toggle ---------------------------------------
$btnCompactToggle.Add_Click({
    $script:IsCompactMode = -not $script:IsCompactMode
    $pnlDayNotes.Visible = -not $script:IsCompactMode
    $lblChecklistHeader.Visible = -not $script:IsCompactMode
    $btnCompactToggle.Text = if ($script:IsCompactMode) { 'Normal View' } else { 'Compact View' }
})

# ---- Lock (always on top) toggle ----------------------------------------
$btnLock.Add_Click({
    $script:IsLocked = -not $script:IsLocked
    $script:MainForm.TopMost = $script:IsLocked
    $btnLock.Text = if ($script:IsLocked) { 'Unlock' } else { 'Lock' }
    $btnLock.BackColor = if ($script:IsLocked) { [System.Drawing.Color]::FromArgb(224, 236, 255) } else { [System.Drawing.SystemColors]::Control }
})

# ---- Opacity slider ------------------------------------------------------
$trackOpacity.Add_Scroll({
    $script:MainForm.Opacity = $trackOpacity.Value / 100.0
    $script:ToolTip.SetToolTip($trackOpacity, "Opacity: $($trackOpacity.Value)%")
})

# ---- "More" menu ---------------------------------------------------------
$btnMore.Add_Click({
    $script:MoreMenu.Show($btnMore, (New-Object System.Drawing.Point(0, $btnMore.Height)))
})

# =========================================================================
# 14. Timers - reminders + shared-file auto refresh
# =========================================================================

$script:ReminderTimer = New-Object System.Windows.Forms.Timer
$script:ReminderTimer.Interval = 20000   # check every 20 seconds
$script:ReminderTimer.Add_Tick({
    $todayStr = (Get-Date).ToString('yyyy-MM-dd')
    $now = Get-Date
    foreach ($t in $script:Tasks) {
        if ($t.Status -eq 'Done' -or $t.TaskDate -ne $todayStr -or -not $t.ReminderTime) { continue }
        try { $baseDT = [datetime]::ParseExact("$($t.TaskDate) $($t.ReminderTime)", 'yyyy-MM-dd HH:mm', $null) } catch { continue }
        $effectiveDT = $baseDT
        if ($script:SnoozeOverrides.ContainsKey($t.Id)) { $effectiveDT = $script:SnoozeOverrides[$t.Id] }
        $key = "$($t.Id)|$($effectiveDT.ToString('yyyyMMddHHmm'))"
        if ($now -ge $effectiveDT -and -not $script:ShownReminderKeys.ContainsKey($key)) {
            $script:ShownReminderKeys[$key] = $true
            Show-ReminderPopup -Task $t
        }
    }
})
$script:ReminderTimer.Start()

$script:AutoRefreshTimer = New-Object System.Windows.Forms.Timer
$script:AutoRefreshTimer.Interval = 90000   # reload the shared file every 90 seconds
$script:AutoRefreshTimer.Add_Tick({
    try {
        $didReload = $false
        if (Test-Path $script:TasksFile) {
            $lastWrite = (Get-Item $script:TasksFile).LastWriteTime
            if ($lastWrite -ne $script:LastLoadedWriteTime) {
                Load-Tasks
                [void](Invoke-CarryForward -Silent)
                $didReload = $true
            }
        }
        if ((-not $script:txtDayNotes.Focused) -and (Test-Path $script:DayNotesFile)) {
            $lastNotesWrite = (Get-Item $script:DayNotesFile).LastWriteTime
            if ($lastNotesWrite -ne $script:LastLoadedDayNotesWriteTime) {
                Load-DayNotes
                $didReload = $true
            }
        }
        if ($didReload) {
            Refresh-AllViews
            $script:lblStatus.Text = "Auto-refreshed at $((Get-Date).ToString('HH:mm:ss'))"
        }
    } catch { }
})
$script:AutoRefreshTimer.Start()

# =========================================================================
# 15. Startup sequence
# =========================================================================

$script:CurrentUser = Get-CurrentUser
Load-Tasks
Load-DayNotes
[void](Invoke-CarryForward -Silent)
Refresh-AllViews
$script:lblStatus.Text = "Loaded at $((Get-Date).ToString('HH:mm:ss')) from $DataPath"

[System.Windows.Forms.Application]::Run($script:MainForm)
