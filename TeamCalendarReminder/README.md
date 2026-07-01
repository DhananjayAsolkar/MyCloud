# Team Calendar & Reminder

A simple Windows desktop tool for team task reminders and a calendar-style
checklist. It runs from a single PowerShell script, needs no admin rights
and no extra installs, and stores its data in a shared OneDrive/SharePoint
synced folder so the whole team sees the same task list.

## Quick Start (for everyone - no computer skills needed)

Follow these steps in order. They only need to be done once per person.

1. **Get the folder.** Ask whoever set this up for the location of the
   `TeamCalendarReminder` folder inside your team's OneDrive or SharePoint.
   It might already be on your computer under a path like
   `OneDrive - YourCompany\Team Documents\TeamCalendarReminder`.
2. **Open the folder** in File Explorer.
3. **Double-click `Run_Team_Calendar_Reminder.bat`.**
   - A black window will flash open briefly - that is normal, leave it be.
   - The Team Calendar & Reminder window will then open.
4. **The first time only**, a small box pops up asking for your name.
   Type your first name (or full name) and click OK. You only need to do
   this once - the tool will remember you after that.
5. You're in! The window shows today's tasks on a clean white page. See
   "Your diary page" below for what everything does.
6. **To close the tool**, just close the window like any other program
   (click the X in the top-right corner). It is safe to close any time -
   nothing is lost, since every change is already saved automatically.

**Tip:** Create a shortcut to `Run_Team_Calendar_Reminder.bat` on your
Desktop so you don't have to hunt for the folder every day. Right-click the
`.bat` file -> **Send to** -> **Desktop (create shortcut)**.

### Your diary page

The window is designed to feel like a small floating sticky note/diary
page sitting on your desktop, not a big software dashboard. It always shows
**one day at a time** on a plain white page:

- The date is shown clearly at the top, e.g. **Tuesday, 01 July 2026**.
- Underneath it is a plain checklist of that day's tasks, one line each:
  `[ ] Prepare weekly status report    16:00   High`
  `[x] Call supplier                   11:00   Medium`
- Below the checklist is a **Notes** box for anything you want to jot down
  about that specific day (it saves automatically as you move to another
  day or click elsewhere).

**Using the checklist with your mouse:**

| Click | Result |
|---|---|
| Click directly on the checkbox | Checks/unchecks the task - marks it Done or reopens it. |
| Click anywhere else on the row | Just selects that task (so Mark as Done/Edit/Delete from the More menu know which one you mean). |
| Double-click the row | Opens the full Edit Task form for that task. |

A task's Priority shows as a small coloured word next to its due time
(red for High, plain black for Medium, grey for Low), and Done tasks turn
grey with a ticked checkbox.

### Getting around and the small controls

| Control | What it does |
|---|---|
| **< Prev** / **Today** / **Next >** | Step through the diary one day at a time, or jump straight back to today. |
| **+ Add Task** | Add a new task directly onto the day you're currently looking at. |
| **Compact View / Normal View** | Compact hides the Notes box for a smaller, glance-only view; Normal shows everything. |
| **Opacity slider** | Drag left to make the window see-through (handy for keeping it visible over other work), right for fully solid. It won't go below 40% so the text stays readable. |
| **Lock / Unlock** | Lock keeps the window on top of every other window, like pinning a sticky note to your screen. |
| **More** | Everything else - see below. |

The window can be resized and moved anywhere on screen, just like any
other window.

### The "More" button

To keep the main page uncluttered, everything you don't need every minute
lives behind the **More** button in the top-right corner:

- **Edit Task** / **Mark as Done** / **Delete Task** - act on whichever
  task you selected on the checklist (same as clicking its checkbox for
  Mark as Done, or double-clicking it for Edit).
- **Filter** - show only Open / In Progress / Done / tasks assigned to you.
  This applies to the diary page and to This Week / Future Tasks below.
- **This Week...** / **Future Tasks...** - opens a small read-only window
  listing everything due this week, or everything still to come.
- **Change Log...** - opens a read-only window showing a history of every
  change anyone has made (who did what and when) - useful if something
  looks wrong and you want to know why.
- **Refresh Now** - grabs the latest list immediately in case a teammate
  just added something (it also happens automatically every 90 seconds).
- **Carry Forward Open Tasks Now** - runs the carry-forward check
  immediately (it also runs automatically every time the app starts).
- **Export to CSV...** - saves the whole list as a file you can open in
  Excel.
- **Change User...** - change the name you're logged in as.

### A 2-minute test to make sure everything works

1. Click **+ Add Task**, type a title like "Test task", leave everything
   else as-is, and click OK. It should appear as a new unchecked line at
   the bottom of today's checklist.
2. Click directly on its checkbox - it should become checked and turn grey.
   Click it again to reopen it.
3. Click on its row once (not the checkbox) - it should highlight/select.
   Double-click it - the Edit Task form should open; change the title
   slightly and click OK.
4. Click **More -> This Week...** - the task should show up there too
   (same week). Close that window.
5. Click **More -> Change Log...** - you should see rows for "Task
   Created", "Task Completed", "Task Reopened" and "Task Edited" for your
   test task. Close that window.
6. Click the **Lock** button (title should change to "Unlock" and the
   window should stay on top of other windows), click it again to unlock.
   Try the **Opacity** slider and the **Compact View** / **Normal View**
   button too.
7. Back on the diary page, select the test task and use **More -> Delete
   Task**, confirm Yes - it disappears from the checklist (it stays in the
   Change Log as a record).

If all seven steps behave as described, the tool is working correctly on
your computer.

## Files in this folder

| File                              | Purpose                                                   |
|------------------------------------|------------------------------------------------------------|
| `Team_Calendar_Reminder.ps1`       | The main application (PowerShell + WinForms GUI).          |
| `Run_Team_Calendar_Reminder.bat`   | Double-click launcher (no need to touch PowerShell settings). |
| `tasks.json`                       | Shared task data. Sample tasks are included to start with. |
| `day_notes.json`                   | Shared per-day notes typed into the diary page's Notes box. |
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

See **Quick Start** at the top of this file for the simple version.
Technical details for anyone curious:

- The `.bat` file runs `Team_Calendar_Reminder.ps1` with a temporary,
  this-window-only PowerShell execution-policy bypass, so it works even on
  PCs where scripts are blocked by default - without changing any
  system-wide setting.
- Alternative way to launch it: right-click `Team_Calendar_Reminder.ps1`
  and choose **"Run with PowerShell"**.
- Your name is stored only on your own PC, in
  `%APPDATA%\TeamCalendarReminder\user.json`, never in the shared folder,
  so it won't overwrite a teammate's name. Change it any time via
  **More -> Change User...**.

## 3. How the shared OneDrive/SharePoint folder works

`tasks.json` and `change_log.csv` are plain text files that live in the
synced folder. When one person adds or edits a task, the app saves the
file, OneDrive/SharePoint syncs it, and everyone else's copy picks up the
change the next time they use **More -> Refresh Now** or within about 90
seconds automatically.

This is a simple "last save wins" model, good enough for a small office
team making occasional changes. It is **not** built for many people
editing the exact same task at the exact same second - if that happens,
the most recent save is the one that is kept. Keep edits quick and use
Refresh often if several people work on the list at once.

Every save also keeps a `tasks.json.bak` backup copy in the same folder in
case a file ever gets corrupted (see Common Errors below).

## 4. How incomplete tasks move to tomorrow

Every time the app starts, refreshes (manually or automatically), or you
use **More -> Carry Forward Open Tasks Now**, it checks every task that is
**not** marked Done:

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
This tool uses a simple last-save-wins approach (see section 3). Use
**More -> Refresh Now** often, and try to keep edits short.

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
