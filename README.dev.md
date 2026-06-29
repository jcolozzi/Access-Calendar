# Access Calendar

Technical and developer guide. Public overview: [README.public.md](README.public.md).

A Microsoft Access calendar app that combines:

- Access/DAO for data storage
- VBA class modules for business logic
- HTML/CSS/JavaScript for the calendar UI
- Edge WebView2 (`WebBrowser0`) as the bridge between VBA and the browser UI

The app supports multi-calendar scheduling, recurring events, drag-and-drop editing, calendar grouping, reminders, theme preferences, an optional event location, adding events to Outlook, and exporting events as `.ics` files for email.

## Highlights

- Month, week, 5-day week, and day views
- Multiple visible calendars at once (with a primary calendar)
- Calendar groups (add, rename, delete, move calendars between groups)
- Recurrence support:
  - Daily, weekly, monthly, yearly
  - Edit one occurrence or entire series
  - Occurrence exceptions and rescheduling
- Drag/drop and resize interactions for timed and all-day events
- Reminder popups driven by a timer loop in VBA
- Light/dark theme persistence and customizable accent colors
- Optional location field on events
- Add events to the user's Outlook calendar (single or recurring) via late-bound COM
- Export events to `.ics` and open a pre-filled Outlook email with the file attached
- Standalone browser mode for UI experimentation (without Access persistence)

## Architecture

The runtime follows a three-layer pattern:

1. JavaScript UI (`calendar.html`) renders state and queues commands.
2. VBA classes process commands and enforce data rules.
3. Access tables persist calendars, groups, and appointments.

Data flow (simplified):

1. `Form_Load` wires all classes and navigates WebView2 to the calendar HTML.
2. `Form_Timer` (300 ms) checks reminders and polls JS for one pending command.
3. `clsCommandProcessor` dispatches the command to repositories/managers.
4. `clsJSBridge.SendDataToCalendar` sends full JSON state back to JS.
5. `window.loadData(...)` rehydrates UI state and re-renders.
6. `frmObserver` (hidden, 3 s timer) polls `tblChangeLog` for changes by other users; `clsPubSubBroker` raises events that trigger `SendDataToCalendar` in other open sessions.

## Repository Layout

```text
Access-Calendar/
  Calendar.accdb
  calendar.html
  ARCHITECTURE.md
  css/
    calendar.css
  js/
    calendar.js
    calendar.ics.js
    vendor/
      ics.js
  bas/
    mod_ChangeLog.bas
  cls/
    clsPubSubBroker.txt
    frmObserver.txt
    clsAppointmentRepo.txt
    clsCalendarExporter.txt
    clsCalendarGroupRepo.txt
    clsCalendarRepo.txt
    clsCommandProcessor.txt
    clsDateHelper.txt
    clsJSBridge.txt
    clsJSONHelper.txt
    clsOutlookService.txt
    clsRecurrenceEngine.txt
    clsReminderManager.txt
    clsThemeManager.txt
    clsUserDetails.txt
    vba_form_code_behind.txt
```

## Core VBA Components

- `mod_ChangeLog`: `LogChange` / `PurgeChangeLog` — audit trail for pub/sub sync
- `clsPubSubBroker`: event dispatcher (AppointmentsChanged, CalendarsChanged, PollingError)
- `frmObserver`: hidden polling form — reads tblChangeLog every 3 s, broadcasts changes to active sessions
- `clsJSONHelper`: JSON escaping and lightweight extraction helpers
- `clsDateHelper`: ISO/date/time conversion and date math helpers
- `clsRecurrenceEngine`: expands recurring masters into virtual occurrences
- `clsAppointmentRepo`: appointment CRUD and recurrence operations
- `clsCalendarRepo`: calendar CRUD + active/open calendar tracking
- `clsCalendarGroupRepo`: group CRUD and calendar-group assignment
- `clsThemeManager`: theme preference synchronization
- `clsReminderManager`: reminder checks and popup notifications
- `clsJSBridge`: WebView2 communication and full-state JSON push
- `clsCommandProcessor`: command router from JS action to VBA handler (also routes Outlook/ICS commands to the exporter and surfaces their result)
- `clsCalendarExporter`: orchestrates Outlook add and `.ics` export (parses the command JSON, decodes the base64 `.ics` payload, writes it with `ADODB.Stream`, attaches it to a mail)
- `clsOutlookService`: late-bound Outlook COM wrapper (single/recurring appointments, mail with attachment); no Outlook reference required

## Database Model

Primary tables used by the app:

- `tblAppointments`
  - Includes recurrence fields (`RecurType`, interval, end rules, exceptions)
  - Includes reminder fields (`ReminderMinutes`, `ReminderFired`)
  - Includes an optional `Location` text field (255)
  - Uses soft-delete (`IsDeleted`)
- `tblCalendars`
  - Calendar metadata, color, optional work hours, optional `GroupID`
  - Uses soft-delete (`IsDeleted`)
- `tblCalendarGroups`
  - Group names and soft-delete flag
- `tblChangeLog`
  - ChangeID (AutoNumber PK), ChangeType (Text), RecordID (Long), Action (Text), ChangedBy (Text), ChangedOn (DateTime)
  - Purged automatically after 7 days via `PurgeChangeLog`

## Command Contract (JS -> VBA)

Actions currently handled by `clsCommandProcessor`:

- Theme and context: `setTheme`, `setActiveCalendar`, `setOpenCalendars`
- Calendar: `addCalendar`, `editCalendar`, `deleteCalendar`
- Groups: `addGroup`, `renameGroup`, `deleteGroup`, `moveCalendarToGroup`
- Appointments: `save`, `move`, `resize`, `rescheduleOccurrence`, `delete`, `deleteOccurrence`, `deleteSeries`
- Outlook / ICS export: `addToOutlook`, `exportIcs` (both return `NORELOAD`; the result text is shown as a toast via `clsCommandProcessor.UserMessage` / `UserMessageKind`)

## Quick Start

1. Open `Calendar.accdb` in Microsoft Access on Windows.
2. Open the calendar form (startup form in the database, or your calendar host form).
3. Ensure the Edge browser control (`WebBrowser0`) can navigate to the local HTML file.
4. Use the calendar UI to create/edit events and calendars.

## Developer Setup Notes

If you are rebuilding the form/classes from this folder export:

1. Import each file in `cls/` as a VBA class module (remove `.txt` extension as needed).
2. Paste `cls/vba_form_code_behind.txt` into your Access form module.
3. Ensure your form contains an Edge browser control named `WebBrowser0`.
4. Keep `calendar.html` in the same folder as `Calendar.accdb` (or adjust navigation path).
5. Import `bas/mod_ChangeLog.bas` as a standard module.
6. Create the `frmObserver` hidden form and import its code-behind from `cls/frmObserver.txt`.
7. Create the `tblChangeLog` table (ChangeID AutoNumber PK, ChangeType Text, RecordID Long, Action Text, ChangedBy Text, ChangedOn DateTime).
8. Import `cls/clsOutlookService.txt` and `cls/clsCalendarExporter.txt` as class modules. Outlook is accessed late-bound, so no `Microsoft Outlook xx.x Object Library` reference is required.
9. Ensure `tblAppointments` has a `Location` text field (255) for event locations.
10. Confirm the front-end files are present alongside `calendar.html`: `css/calendar.css`, `js/calendar.js`, `js/calendar.ics.js`, and `js/vendor/ics.js`.

## Important Filename Note

The form shell in `cls/vba_form_code_behind.txt` currently declares:

```vba
Private Const CALENDAR_NAME As String = "calendar.html"
```

This repository currently contains `calendar.html`. Use one of these approaches:

- change `CALENDAR_NAME` to `"calendar.html"`, or
- rename/copy the HTML file to `calendarv27.html`.

Without this alignment, `DocumentComplete` checks and initial navigation can miss the expected page name.

## Standalone UI Mode

When `calendar.html` is opened outside Access, it runs in standalone mode:

- interactions still render and behave in the browser
- data is not persisted to Access tables
- useful for fast UI experimentation and CSS/JS iteration

## Reference Documentation

- `ARCHITECTURE.md` contains a deeper walkthrough of lifecycle, class responsibilities, and integration flow.
