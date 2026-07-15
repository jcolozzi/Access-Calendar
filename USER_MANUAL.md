# Access Calendar — User Manual

This manual covers the day-to-day features of Access Calendar: views, calendars, appointments, recurrence, reminders, Outlook integration, and appearance settings.

## Contents

1. [Overview of the Calendar Screen](#1-overview-of-the-calendar-screen)
2. [Navigating Dates](#2-navigating-dates)
3. [Views](#3-views)
4. [Calendars](#4-calendars)
5. [Calendar Groups](#5-calendar-groups)
6. [Creating and Editing Appointments](#6-creating-and-editing-appointments)
7. [Recurring Appointments](#7-recurring-appointments)
8. [Reminders](#8-reminders)
9. [Moving, Resizing, and Previewing Appointments](#9-moving-resizing-and-previewing-appointments)
10. [Sending Appointments to Outlook](#10-sending-appointments-to-outlook)
11. [Appearance Settings](#11-appearance-settings)
12. [Real-Time Updates (Multiple Users)](#12-real-time-updates-multiple-users)

---

## 1. Overview of the Calendar Screen

The screen is made up of four areas:

- **Header** — app title, month/year navigator, view tabs (Month / Week / Day), a **Today** button, and a status indicator (top right) that reads **● Connected** when idle, **Saving…** while a change is being written to the database, or **Error** if something went wrong.
- **Sidebar** (left) — a mini-month picker, your list of calendars and calendar groups (see Sections 4–5), and a **Settings** button (see Section 11). Click the arrow (**‹**) at the top of the sidebar to collapse it for more screen space; click again to expand.
- **Calendar tab bar** — appears only when you have 2 or more calendars open at once; lets you switch which calendar is "primary" (details in Section 4).
- **Calendar grid** — the main Month, Week, or Day view where appointments are displayed.

---

## 2. Navigating Dates

- **Prev / Next arrows** (‹ ›) next to the date label move back/forward one month, one week, or one day, depending on the current view.
- **Today** button jumps back to the current date.
- **Click the date label** in the header to open a Month/Year picker — choose a month and year and click **Go** to jump straight there.
- **Mini calendar** (sidebar) — click any date to jump the main view to that day. Use its own ‹ › arrows to browse months without changing your main view's date.
- **Week view day headers** — in Week view, clicking a day's date at the top of a column switches to Day view for that date.

---

## 3. Views

| View | Description |
|---|---|
| **Month** | Shows the full month in a grid. Each day shows a short list of that day's appointments. |
| **Week** | Shows a 7-day (or 5-day) view with hourly time slots. Toggle between **7 days** and **5 days** (work week) using the buttons that appear in Week view. |
| **Day** | Shows a single day with hourly time slots, useful when a day has many overlapping appointments. |

Switch views using the **Month / Week / Day** tabs in the header.

---

## 4. Calendars

Access Calendar supports multiple calendars (e.g., "Work", "Personal", "Team"), each with its own color.

### Creating a calendar

1. Click the **+** button next to a group header in the sidebar (or the ungrouped section) to add a new calendar. Groups are explained in Section 5.
2. Enter a **Name** and pick a **Calendar Color** from the color options.
3. Optionally check **Enable work hours shading** and set a **Work Start** / **Work End** time — this highlights that calendar's working hours (e.g., 9am–5pm) with a shaded band in Week/Day view, making it easy to spot business hours at a glance.
4. Click **Create**.

### Editing a calendar

Click the pencil (✎) icon on a calendar in the sidebar to reopen the same dialog and change its name, color, or work-hours shading.

### Deleting a calendar

Click the ✕ icon on a calendar in the sidebar. You'll be asked to confirm — this removes the calendar and all of its appointments.

### Showing / hiding calendars

Each calendar has a checkbox in the sidebar:

- **Checked** — calendar's appointments are shown ("open").
- **Unchecked** — calendar's appointments are hidden.
- You can have up to **5 calendars open** at once — the checkbox simply won't respond if you try to check a 6th.
- At least **1 calendar must remain open** — the checkbox for the last open calendar won't uncheck.

### Primary calendar

The **primary** calendar is where new appointments are created by default (via the **Calendar** dropdown in the New Appointment window described in Section 6, which defaults to the primary). When only one calendar is open, it's automatically primary.

When 2+ calendars are open, a **calendar tab bar** appears below the header:

- Click a tab's name to make that calendar primary (highlighted).
- Click the **×** on a tab to hide that calendar (same as unchecking it in the sidebar).

---

## 5. Calendar Groups

Groups let you organize related calendars together in the sidebar (e.g., "Work" group containing "Meetings" and "Projects" calendars).

- **Create a group**: click the **+** button in the sidebar header (next to "Calendars").
- **Rename a group**: click the pencil (✎) icon on the group header.
- **Delete a group**: click the ✕ icon on the group header. Calendars inside the group are **not** deleted — they become ungrouped.
- **Move a calendar into a group**: drag the calendar item and drop it onto the target group.
- **Collapse/expand a group**: click the group name/arrow to show or hide its calendar list.

---

## 6. Creating and Editing Appointments

Click any date (Month view) or time slot (Week/Day view) to open the **New Appointment** window. Click an existing appointment to **edit** it.

Fields available:

| Field | Notes |
|---|---|
| **Calendar** | Which open calendar the appointment belongs to. |
| **Title** | Appointment name. If left blank, it's saved as "No title." |
| **Date** / **End Date** | Leave End Date the same as Date for a single-day appointment; set it to a later date to create a multi-day appointment. |
| **All Day** | Toggle on to hide start/end time fields and mark the appointment as all-day. |
| **Start Time / End Time** | Only shown when All Day is off. |
| **Notes** | Free-text notes/description. |
| **Location** | Optional location text. |
| **Color** | Overrides the calendar's default color for this appointment. |
| **Reminder** | Choose how long before the appointment a reminder popup should appear, or **None**. See Section 8 for how reminders work. |

Click **Save** to store the appointment, or **Cancel** to discard changes.

When editing an existing appointment, a **Delete** button is also available. Deleting a plain (non-recurring) appointment happens immediately with no confirmation prompt, so double-check before clicking it. Deleting a recurring appointment asks you to choose a scope — see Section 7.

There's no Undo command in Access Calendar. If you delete, move, or resize something by mistake, you'll need to manually re-create or fix it.

---

## 7. Recurring Appointments

Click **Make Recurring** in the appointment window to reveal the recurrence panel and set up a repeating pattern.

### Pattern types

- **Daily** — every *N* day(s).
- **Weekly** — every *N* week(s), on one or more selected days of the week.
- **Monthly** — either:
  - a fixed day of the month (e.g., "the 15th of every 2 months"), or
  - a relative weekday (e.g., "the second Friday of every month").
- **Yearly** — every *N* year(s), on the appointment's original date.

### Range of Recurrence

- **Start** — the date the pattern begins (defaults to the appointment's date).
- **End by** — a specific end date.
- **End after** — a specific number of occurrences.
- **No end date** — repeats indefinitely.

### Editing or deleting a recurring appointment

Opening any occurrence and clicking **Save** or **Delete** prompts you to choose the scope:

- **This occurrence** — only the single date you opened is changed/removed; the rest of the series is unaffected (editing creates a one-off exception and books a separate standalone appointment on the new date/time; deleting excludes just that date from the series).
- **All occurrences** — the entire series' pattern is updated, or the whole series is deleted.

---

## 8. Reminders

Set a reminder when creating or editing an appointment using the **Reminder** dropdown (options range from **0 minutes** up to **2 days** before, or **None**).

When a reminder's time arrives, Access Calendar shows a popup notification with the appointment title. For a recurring appointment, the reminder fires separately before *each* upcoming occurrence, not just the first.

> **Note:** Reminders are only checked while the Access Calendar form is open on your screen. If you close the calendar form (or close Access), reminders will not pop up — reopen the calendar to resume receiving them.

---

## 9. Moving, Resizing, and Previewing Appointments

- **Move**: drag an appointment to a new day (Month view) or a new day/time (Week/Day view) to reschedule it.
- **Resize**: in Week/Day view, drag the top or bottom edge of an appointment to change its start or end time (the date stays the same).
- **Reassign to another calendar**: dragging an appointment onto a different calendar's column/area (when multiple calendars are open) moves it to that calendar.
- **Preview**: hover the mouse over any appointment on the grid to see a quick-preview popup with its calendar, date/time, location, notes, and a recurring icon if it repeats — without opening the full edit window.

---

## 10. Sending Appointments to Outlook

From the appointment window, two export actions are available:

- **Add to Outlook** — creates the appointment directly in your default Outlook calendar (as a single appointment, or as a recurring Outlook appointment if the item has a recurrence pattern).
- **Export .ics** — builds a standard `.ics` calendar file (a universal calendar-file format that other apps, including Outlook, can open to add the appointment automatically) and opens a new Outlook email with the file already attached, ready to send to others.

A small pop-up confirmation message appears in the bottom corner after either action to show success or report an error.

---

## 11. Appearance Settings

Click **⚙ Settings** in the sidebar to open the Settings panel:

- **Appearance** — toggle between Light Mode and Dark Mode using the theme button.
- **Main Color** — the primary accent color used throughout the UI (buttons, highlights). Set via the color picker or by typing a hex value.
- **Accent Hover Color** — the color shown on hover for accent-colored elements.
- **Reset** — restores the default colors.
- **Save** / **Cancel** — apply or discard your changes.

Theme and color preferences are remembered the next time you open the calendar **on the same computer**. They are stored locally on your machine, not shared with other users and not synced if you switch computers.

---

## 12. Real-Time Updates (Multiple Users)

If other users have the calendar open at the same time, changes anyone makes (new/edited/deleted appointments, calendars, or groups) automatically appear in your view within a few seconds — no manual refresh is needed.
