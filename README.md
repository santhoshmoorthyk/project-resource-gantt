# Project Resource Scheduling

An Excel/VBA Gantt scheduler for task planning, team resource allocation and submission tracking.

> **Note:** All names and data in this repository are sample/dummy data for demonstration purposes.

## Overview

This workbook plans work across a team on a day-by-day calendar. You build the schedule visually by picking tasks from dropdowns and dragging bars across dates. Macros then turn that calendar into a task list with durations, change history and submission dates.

You can:

- Assign tasks to team members and see who is working on what, each day
- Mark holidays and leave directly on the calendar
- Track task durations, start and end dates, and date changes
- Track submission dates for each item
- Filter the calendar to follow a single task across the whole team

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Sheet2 – Gantt Calendar** | The main working sheet. Dates run across the top, team members run down the left, and tasks appear as coloured bars |
| **Sheet1 – Task List** | Auto-generated list of every bar: group, item, task, assignee, duration, from/to dates, change dates and submission dates |
| **Sheet3 – Item Lookup** | Level 1 and Level 2 dropdown values, plus submission date and hover message for each item |
| **Sheet4 – Task Types** | Level 3 (task type) and Level 4 (coverage) dropdown values |
| **Sheet12 – History Snapshot** | Previous dates saved so that date changes can be detected and recorded |
| **Sheet6 – Reference List** | Extra reference list |
| **DropdownHelper** | Hidden helper sheet used by the dropdown system |

## How the Gantt Calendar (Sheet2) Works

### Layout

- **Row 1** holds the dates (one column per day).
- **Column A** holds the team member names.
- Each person has **two rows**, so one person can be given two tasks on the same day.
- Every cell in the grid from B2 onwards is an input cell.

### Adding a task: 4-level dropdown

Click an empty cell and a dropdown opens automatically. Each choice opens the next dropdown:

1. **Level 1** – group
2. **Level 2** – item within that group
3. **Level 3** – task type
4. **Level 4** – coverage (for example full or partial)

Each dropdown has a **`<< Back`** option to go up one level. When Level 4 is chosen, the cell fills with the combined text:

```
Group | Item | Task | Coverage
```

and it is coloured automatically. Every task type has its own colour.

### Leave and Holiday

Choose **Leave** or **Holiday** at Level 1. It is a single choice with no further levels. Leave shows as a red cell and Holiday as an orange cell.

### Making a bar longer

Select the finished cell and **drag sideways** across the dates the task covers. The cells merge into one continuous bar.

Rules enforced by the macro:

- Only **left-right** dragging is allowed. Up-down dragging is rejected and undone.
- A bar can be at most **30 cells** long. Longer drags are undone with a message.
- Dragging over empty cells clears them.
- Dragging a Leave or Holiday cell fills each cell separately, so each one stays editable.

### Editing or removing a bar

- Click a bar to change it. The Level 4 dropdown opens so you can adjust the coverage.
- Clear the cell to remove the bar. The cell goes back to a Level 1 dropdown with no colour.

### Filter by task (double-click)

Double-click any bar to show only the rows where the same item and task appear. Other rows are hidden or faded, so you can see who is working on it and when. Double-click the same bar again, or an empty cell, to clear the filter.

### Hover notes

After the macros run, each bar gets a hover note with a message defined for that item in the lookup sheet.

## Macros

| Macro | What it does |
|---|---|
| `RunAll` | Runs the full update below in one click |
| `SyncSheet1FromSheet2` | Reads every bar from the calendar and writes it to the Task List. Date changes are detected against the history snapshot and recorded as 1st, 2nd, 3rd change and so on |
| `PopulateSubmissionDates` | Fills submission dates from the lookup sheet |
| `UpdateTimeTaken` | Calculates the time taken for each task from its bar dates, taking the person's leave into account |
| `AddBarMessageTags` | Adds hover notes to the calendar bars |
| `HideWeekendColumns` | Hides Saturday and Sunday columns on the calendar |
| `UnhideAllColumns` | Brings the hidden columns back |
| `RefreshDropdownData` / `ManualRefresh` | Rebuilds the dropdowns after you edit the lookup sheets |

### Upcoming tasks panel

When the workbook opens, a small floating panel appears at the top right of the screen. It lists the upcoming tasks sorted by date, with today's date shown. Double-click an item in the panel to filter by it.

## Typical Workflow

1. Add or edit your groups, items and task types in the lookup sheets, then run `RefreshDropdownData`.
2. In **Sheet2**, pick tasks from the dropdowns and drag the bars across the dates.
3. Mark leave and holidays for each person.
4. Run `RunAll` to update the Task List, submission dates, durations and hover notes.
5. Use `HideWeekendColumns` if you only want working days visible.
6. When dates change, move the bars on the calendar and run `RunAll` again. The Task List records the change.

## Getting Started

1. Download `gantt_process_scheduling.xlsm`.
2. Open it in Microsoft Excel (desktop version).
3. Click **Enable Content** when prompted so the macros can run.
4. Open **Sheet2** and start scheduling.

Screenshots

Add a screenshot of the Gantt calendar here:
<img width="1285" height="610" alt="image" src="https://github.com/user-attachments/assets/01fa578f-e01c-412f-b624-95f92d5c3c94" />
<img width="1606" height="798" alt="image" src="https://github.com/user-attachments/assets/b5eb7cc2-bb99-41f3-bad7-04ddbc4608aa" />
<img width="1543" height="833" alt="image" src="https://github.com/user-attachments/assets/b96f85aa-9efb-499d-a046-825ea53fa2d6" />
<img width="481" height="811" alt="image" src="https://github.com/user-attachments/assets/8b9abfc2-b6a1-491b-9c6d-7cdce272582d" />
<img width="576" height="802" alt="image" src="https://github.com/user-attachments/assets/4a976a16-321f-40e9-a1e1-871088783e32" />

## Requirements

- Microsoft Excel 2016 or later (desktop)
- Macros enabled

## Tech Used

- Microsoft Excel
- VBA (worksheet events, dropdown validation, UserForm)

## Author

**santhoshmoorthyk**
