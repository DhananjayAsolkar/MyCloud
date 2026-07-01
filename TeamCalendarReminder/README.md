# Team Calendar & Reminder

A simple Windows desktop tool for team task reminders and a calendar-style
checklist. It runs from a single PowerShell script, needs no admin rights
and no extra installs, and stores its data in a shared OneDrive/SharePoint
synced folder so the whole team sees the same task list.

## Files in this folder

| File                              | Purpose                                                   |
|------------------------------------|------------------------------------------------------------|
| `Team_Calendar_Reminder.ps1`       | The main application (PowerShell + WinForms GUI).          |
| `Run_Team_Calendar_Reminder.bat`   | Double-click launcher (no need to touch PowerShell settings). |
| `tasks.json`                       | Shared task data. Sample tasks are included to start with. |
| `change_log.csv`                   | Shared audit trail of every change made to any task.        |
| `README.md`                        | This file.                                                  |

## 1. Where to place the files

Copy the **entire `TeamCalendarReminder` folder** (all files together) into
a location that is synced by OneDrive or SharePoint and shared with your
team, for example:

```
OneDrive - YourCompany\Team Documents\TeamCalendarReminder\
```

Every team member should point to that same synced folder (either because
OneDrive syncs it down to their own PC automatically, or because they open
it from a shared/mapped drive). The script automatically uses the folder
it is run from to find `tasks.json` and `change_log.csv`, so as long as
everyone runs the copy that lives inside the synced folder, everyone reads
and writes the same data.

You do **not** need to install anything or ask IT for admin rights - the
tool only uses features already built into Windows.

## 2. How to run the script

**Easiest way:** double-click `Run_Team_Calendar_Reminder.bat`.

**Alternative:** right-click `Team_Calendar_Reminder.ps1` and choose
**"Run with PowerShell"**.

The first time you run it, a small box will ask for **your name**. This is
stored only on your own PC (in `%APPDATA%\TeamCalendarReminder\user.json`),
never in the shared folder, so it won't overwrite a teammate's name. You
can change it later with the **"Change User"** button.

### Everyday use

- **Add Task** - opens a form to create a new task (title, description,
  date, due time, reminder time, priority, assigned person, status).
- **Edit Task** / **Mark as Done** / **Delete Task** - act on whichever
  task is selected in the currently open tab.
- **Refresh** - reloads the shared file immediately (it also auto-refreshes
  every 90 seconds on its own).
- **Carry Forward Open Tasks** - manually run the carry-forward check (see
  below); it also runs automatically every time the app starts or refreshes.
- **Export to CSV** - saves all tasks to a `.csv` file you can open in Excel.
- **Filter** dropdown - narrows any tab to Open / In Progress / Done /
  Assigned To Me.
- **Always on Top** - keeps the window visible above other programs.
- Tabs: **Day View** (use Prev/Next/Today or the calendar to jump to any
  date), **This Week**, **Future Tasks**, and **Change Log**.

## 3. How the shared OneDrive/SharePoint folder works

`tasks.json` and `change_log.csv` are plain text files that live in the
synced folder. When one person adds or edits a task, the app saves the
file, OneDrive/SharePoint syncs it, and everyone else's copy picks up the
change the next time they hit **Refresh** or within about 90 seconds
automatically.

This is a simple "last save wins" model, good enough for a small office
team making occasional changes. It is **not** built for many people
editing the exact same task at the exact same second - if that happens,
the most recent save is the one that is kept. Keep edits quick and use
Refresh often if several people work on the list at once.

Every save also keeps a `tasks.json.bak` backup copy in the same folder in
case a file ever gets corrupted (see Common Errors below).

## 4. How incomplete tasks move to tomorrow

Every time the app starts, refreshes (manually or automatically), or you
click **Carry Forward Open Tasks**, it checks every task that is **not**
marked Done:

- If its date is in the past, the task's date is moved forward to **today**
  and a note is added: *"Carried forward from previous date (yyyy-mm-dd)."*
- The task's original `CreatedDate` and `Assigned To` never change - only
  the date it appears on ("page") moves.
- The **same task record** is updated in place - nothing is duplicated.
- **Completed (Done) tasks are never moved.** They stay on the date they
  were finished, exactly as required.
- If you manually change a task's date yourself using **Edit Task**, that
  counts as an intentional reschedule (not a carry-forward), so the
  automatic note is cleared.

## 5. Reminders

If a task has a Reminder Time set, once that time passes on the task's
date, a small pop-up appears in the corner of the screen showing the task
name, due time, assigned person and comment. You can **Snooze** it for 5,
15 or 30 minutes, or **Dismiss** it. Reminders are personal to your own PC
- they are not written back to the shared file, so snoozing or dismissing
a reminder on your machine does not affect what your teammates see.

## 6. Common errors and fixes

**"Running scripts is disabled on this system"**
Use `Run_Team_Calendar_Reminder.bat` instead of running the `.ps1` file
directly - it starts PowerShell with a temporary, this-window-only bypass
and does not change any system-wide settings. You can also open PowerShell
yourself and run:
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`
before launching the script.

**Windows says the file is "blocked" because it came from the internet**
Right-click `Team_Calendar_Reminder.ps1` -> Properties -> tick **Unblock**
-> OK. Or run `Unblock-File .\Team_Calendar_Reminder.ps1` in PowerShell.

**"Could not read/save tasks.json. It may be corrupted or still syncing"**
Wait a few seconds for OneDrive to finish syncing (check the sync icon in
the taskbar) and click **Refresh**. If `tasks.json` ever becomes genuinely
corrupted, rename `tasks.json.bak` to `tasks.json` to restore the last
good save.

**Reminder pop-ups are not appearing**
Make sure the task's Reminder Time field is filled in, the task's date is
today, and the app window/process is still open. Also check your PC's
clock and time zone are correct.

**Two people edited the list at the same time and a change seems "lost"**
This tool uses a simple last-save-wins approach (see section 3). Click
**Refresh** often, and try to keep edits short.

**The window looks tiny or oddly placed**
Resize or move it as normal - it remembers nothing about window position
between runs, so it always opens centered at a default size.

## MVP feature checklist

- [x] Add task
- [x] Show Today / Tomorrow (via Day View + Next Day) / This Week
- [x] Mark done
- [x] Carry forward incomplete tasks (automatic + manual button)
- [x] Save/load from shared JSON file
- [x] Popup reminder with snooze
