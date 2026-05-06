# Access Calendar Architecture & Integration Guide

**Overview**: A multi-calendar event management system built with VBA (server-side data & business logic), HTML/CSS/JavaScript (client-side UI), and WebView2 (browser control in Access).

---

## 🏗️ High-Level Architecture

### Three-Tier Communication Pattern

```
┌─────────────────────────────────────────────────────────────┐
│                   HTML/CSS/JavaScript (WebView2)            │
│          (UI Rendering, User Interactions, State)           │
└── ↕ ────────────────────────────────────────────────────────┘
     VBA Bridge (clsJSBridge)
     - Polls for JS commands via getPendingCommand()
     - Pushes full state JSON via SendDataToCalendar()
     - Executes JS via ExecuteJavascript()
┌─────────────────────────────────────────────────────────────┐
│                  VBA Class Modules (Data & Logic)           │
│           Access Form + 10 OOP Class Modules                │
└─────────────────────────────────────────────────────────────┘
     DAO Database Access (SQL)
┌─────────────────────────────────────────────────────────────┐
│      Access Database (tblAppointments, tblCalendars, etc.)  │
└─────────────────────────────────────────────────────────────┘
```

---

## 🔄 Request/Response Cycle

### 1. **Form Load** (`Form_Load`)

```vba
Private Sub Form_Load()
    DoCmd.Maximize
    
    ' Phase 1: Create utility classes (pure functions)
    Dim jh As New clsJSONHelper
    Dim dh As New clsDateHelper
    
    ' Phase 2: Create recurrence engine (converts recurring master → virtual occurrences)
    Dim re As New clsRecurrenceEngine
    re.Init jh, dh
    
    ' Phase 3: Create data-access repos
    Dim ar As New clsAppointmentRepo:  ar.Init jh, dh
    Dim cr As New clsCalendarRepo:     cr.Init jh, dh
    Dim gr As New clsCalendarGroupRepo: gr.Init jh
    
    ' Phase 4: Create managers
    Dim th As New clsThemeManager:     th.Init jh
    Set m_Reminders = New clsReminderManager
    m_Reminders.Init re
    
    ' Phase 5: Create JS bridge (wires WebView2 control)
    Set m_Bridge = New clsJSBridge
    m_Bridge.Init Me.WebBrowser0, jh, re, cr, th
    
    ' Phase 6: Create command dispatcher
    Set m_CmdProc = New clsCommandProcessor
    m_CmdProc.Init jh, ar, cr, gr, th
    
    ' Navigate WebView2 to HTML file
    Me.WebBrowser0.Navigate "https://msaccess/" & CurrentProject.Path & "\calendar.html"
    Me.TimerInterval = 300  ' Poll every 300ms
End Sub
```

**Key Points:**

- **Dependency Injection**: Classes receive their dependencies via `Init()`, enabling loose coupling
- **Singleton Pattern**: Form-level variables (`m_Bridge`, `m_CmdProc`, `m_Reminders`) hold the same instances throughout the session
- **Timer Polling**: 300ms interval triggers both reminder checks and JS command polling

---

### 2. **HTML Loading** (`WebBrowser0_DocumentComplete`)

When the HTML finishes loading:

```vba
Private Sub WebBrowser0_DocumentComplete(URL As Variant)
    If InStr(1, URL, CALENDAR_NAME, vbTextCompare) > 0 Then
        On Error GoTo DCErr
        m_Bridge.SendDataToCalendar  ' Push data immediately
        Exit Sub
    End If
DCErr:
    ' Error handling...
End Sub
```

**In VBA (clsJSBridge.SendDataToCalendar):**

- Reads calendars, calendar groups, and today's appointments from the database
- Expands recurring appointments into virtual occurrences
- Builds a comprehensive JSON payload
- Calls `ExecuteJavascript("loadData(...)")` to populate the JS state

```vba
Public Sub SendDataToCalendar(Optional restoreCalId As String = "")
    ' Build JSON with calendars, groups, and all visible appointments
    Dim json As String
    json = BuildFullJSON(restoreCalId)
    
    ' Execute JavaScript to populate the page state
    ExecuteJavascript "loadData(" & json & ");"
End Sub
```

**In JavaScript (calendar.html):**

```javascript
window.loadData = function(json) {
    const d = JSON.parse(json);
    calendars      = d.calendars;        // List of calendar objects
    calendarGroups = d.calendarGroups;   // List of calendar groups
    events         = d.appointments;     // All appointment JSON arrays
    
    // Sync theme and open-calendar preferences
    primaryCalendarId = d.primaryCalendarId;
    openCalendarIds   = d.openCalendarIds;
    
    renderSidebar();  // Rebuild left panel
    render();         // Rebuild calendar view
    setStatus('idle');
};
```

---

### 3. **User Interaction → Command Queue**

When a user creates/edits/deletes an appointment or changes settings:

1. **JavaScript collects form data and sends to VBA:**

```javascript
function saveAppointment(apptData) {
    const cmd = {
        action: 'saveAppointment',
        id: apptData.id || '',
        calendarId: apptData.calendarId,
        title: apptData.title,
        date: apptData.date,        // ISO date
        start: apptData.start,      // HH:mm
        end: apptData.end,
        color: apptData.color,
        allDay: apptData.allDay,
        recurType: apptData.recurType,
        recurInterval: apptData.recurInterval,
        // ... other recurrence fields
    };
    queueCommand(cmd);  // Store in pendingCommand global variable
}
```

2. **VBA polls for commands:**

```vba
Private Sub Form_Timer()
    If m_Processing Then Exit Sub
    m_Processing = True
    On Error GoTo EH
    
    ' Check reminders
    On Error Resume Next
    m_Reminders.CheckAll
    On Error GoTo EH
    
    ' Poll for pending JS command
    Dim cmd As String
    cmd = m_Bridge.GetPendingCommand()
    If cmd = "" Then GoTo ExitHere
    
    ' Dispatch command and get return code
    Dim restoreCalId As String
    restoreCalId = m_CmdProc.ProcessCommand(cmd)
    
    ' If data changed, reload from database
    If restoreCalId <> "NORELOAD" Then
        DoEvents
        m_Bridge.SendDataToCalendar restoreCalId
    End If
    
ExitHere:
    m_Processing = False
    Exit Sub
EH:
    ' Error handling...
End Sub
```

---

### 4. **Command Processing** (clsCommandProcessor)

```vba
Public Function ProcessCommand(ByVal jsonStr As String) As String
    Dim action As String
    action = m_JSON.ExtractValue(jsonStr, "action")
    
    Select Case action
        Case "setTheme"
            m_Theme.SetTheme jsonStr
            ProcessCommand = "NORELOAD"  ' No DB change, don't reload
            
        Case "saveAppointment"
            Dim calId As String
            calId = m_Appts.Save(jsonStr)  ' Insert or update
            ProcessCommand = calId         ' Reload data for this calendar
            
        Case "deleteAppointment"
            m_Appts.Delete jsonStr
            ProcessCommand = m_Cals.ActiveCalendarID
            
        Case "moveAppointment"
            m_Appts.Move jsonStr
            ProcessCommand = m_Cals.ActiveCalendarID
            
        ' ... more cases for calendars, groups, etc.
    End Select
End Function
```

**Return Codes:**

- `"NORELOAD"`: No database change; don't refresh UI
- Calendar ID string: Database changed; reload and restore view to this calendar
- Empty: Calendar was deleted; reload and show first remaining calendar

---

## 📚 VBA Class Modules Overview

### **clsJSONHelper** — Utility for JSON

- **`Escape(value)`**: Escapes quotes, newlines, backslashes for safe JSON embedding
- **`ExtractValue(jsonStr, key)`**: Lightweight key lookup (no external JSON library)
- **`BuildOccJSON(...)`**: Constructs one recurring-occurrence JSON event object

**Example:**

```vba
Dim jh As New clsJSONHelper
Dim title As String: title = jh.Escape("John's ""Birthday"" Party")
' Result: "John's \"Birthday\" Party"

Dim value As String: value = jh.ExtractValue("{""name"":""Alice"",""age"":30}", "name")
' Result: "Alice"
```

---

### **clsDateHelper** — Date Math & Conversions

- **`ISOToDate(iso)`**: Converts "2026-03-11" → `3/11/2026`
- **`HMToTime(hm)`**: Converts "14:30" → `2:30:00 PM`
- **`GetNthWeekdayOfMonth(yr, mn, weekOrd, targetDOW)`**: E.g., "last Friday of March 2026"

**Example:**

```vba
Dim dh As New clsDateHelper
Dim d As Date: d = dh.ISOToDate("2026-03-15")
Dim t As Date: t = dh.HMToTime("14:30")
Dim lastFriday As Date: lastFriday = dh.GetNthWeekdayOfMonth(2026, 3, -1, 5)
```

---

### **clsRecurrenceEngine** — Expands Recurring Appointments

**Core Concept**: Master record in database vs. virtual occurrences

- **Input**: One master appointment record with recurrence fields
  - `RecurType`: "daily", "weekly", "monthly", "yearly"
  - `RecurInterval`: How many units between recurrences
  - `RecurDaysOfWeek`: For weekly: "0,2,4" = Sun, Tue, Thu
  - `RecurMonthlyMode`: "day" (15th of month) or "weekday" (2nd Friday)
  - `RecurEndType`: "count" (N occurrences) or "date" (end on this date)
  - `RecurExceptions`: JSON array of excluded dates

- **Output**: Comma-separated JSON objects for each occurrence in a window [winStart, winEnd]

**Example:**

```vba
Dim re As New clsRecurrenceEngine
re.Init jh, dh

Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("SELECT * FROM tblAppointments WHERE AppointmentID = 123")

Dim json As String
json = re.BuildOccurrences(rs, Date(), Date() + 30)
' json = {"id":123,"date":"2026-03-01",...},{"id":123,"date":"2026-03-08",...}
```

**Recurrence Patterns (simplified):**

| Pattern | Example |
| --------- | --------- |
| **Daily** | Every 2 days |
| **Weekly** | Every 1 week on Mon, Wed, Fri |
| **Monthly (day)** | The 15th of every month |
| **Monthly (weekday)** | The 2nd Friday of every month |
| **Yearly** | Every year on 3/15 |

---

### **clsAppointmentRepo** — Appointment Data Access

- **`Save(jsonStr)`**: Insert or update appointment (detects recurrence fields)
- **`Delete(jsonStr)`**: Soft-delete (sets `IsDeleted = True`)
- **`Move(jsonStr)`**: Change appointment date/time
- **`Resize(jsonStr)`**: Change start/end times

**Example:**

```vba
Dim ar As New clsAppointmentRepo
ar.Init jh, dh

Dim json As String: json = "{""id"":""123"",""calendarId"":""1"",""title"":""Meeting"",...}"
Dim calId As String: calId = ar.Save(json)
' Returns the calendarId, used to reload the UI
```

---

### **clsCalendarRepo** — Calendar Data Access

- **`Add(jsonStr)`**: Create new calendar
- **`Edit(jsonStr)`**: Rename calendar or change properties
- **`Delete(jsonStr)`**: Soft-delete calendar
- **`SetActiveCalendar(jsonStr)`**: Change which calendar is primary
- **`SetOpenCalendars(jsonStr)`**: Sync the checkboxes (multiple calendars visible)

**Properties:**

- `ActiveCalendarID`: The primary calendar (used when creating new appointments)
- `OpenCalendarIds`: Comma-separated list of visible calendars

```vba
Dim cr As New clsCalendarRepo
cr.Init jh, dh
Dim newId As String: newId = cr.Add("{""name"":""Work Events""}")
cr.ActiveCalendarID = newId
```

---

### **clsCalendarGroupRepo** — Calendar Groups

- **`Add(jsonStr)`**: Create new group
- **`Rename(jsonStr)`**: Rename group
- **`Delete(jsonStr)`**: Soft-delete group
- **`MoveCalendar(jsonStr)`**: Move a calendar into/out of a group

```vba
Dim gr As New clsCalendarGroupRepo
gr.Init jh
Dim newGroupId As String: newGroupId = gr.Add("{""name"":""Personal""}")
```

---

### **clsThemeManager** — Dark/Light Mode

- **`SetTheme(jsonStr)`**: Store dark mode preference
- **`DarkMode`** property: Read/write boolean
- **`SyncFromBrowser(wbCtl)`**: Read theme from browser's localStorage

```vba
Dim th As New clsThemeManager
th.Init jh
th.SetTheme("{""dark"":true}")
If th.DarkMode Then
    ' Apply dark theme to printed reports, etc.
End If
```

---

### **clsReminderManager** — Popup Notifications

**Two types:**

1. **Non-recurring**: Tracked via `ReminderFired` field (persistent)
2. **Recurring**: Tracked via in-memory `Collection` (session-scoped)

- **`CheckAll()`**: Called every Form_Timer tick
- Calculates if reminder time passed, fires popup notification
- Popup shows title and reminder minutes before start

```vba
Dim rm As New clsReminderManager
rm.Init re  ' Pass recurrence engine for expanding recurring reminders

' In Form_Timer:
rm.CheckAll()  ' Polls DB every 300ms; fires popups when time arrives
```

---

### **clsJSBridge** — WebView2 Communication

**Bi-directional Bridge:**

**VBA → JavaScript:**

- `ExecuteJavascript(jsCode)`: Run arbitrary JS
- `SendDataToCalendar(restoreCalId)`: Pump full state JSON
- `RetrieveJavascriptValue(jsExpr)`: Get JS value synchronously

**JavaScript → VBA:**

- `queueCommand(obj)`: Store command in global `pendingCommand`
- VBA calls `GetPendingCommand()` every 300ms to retrieve it

**Example:**

```vba
Public Sub ExecuteJS(ByVal code As String)
    m_WB.ExecuteJavascript code
End Sub

Public Function GetPendingCommand() As String
    Dim raw As String
    raw = Nz(m_WB.RetrieveJavascriptValue("getPendingCommand()"), "")
    If raw = "" Or raw = "null" Then
        GetPendingCommand = ""
    Else
        ' Unescape and return
        GetPendingCommand = m_JSON.Unescape(raw)
    End If
End Function

Public Sub SendDataToCalendar(Optional restoreCalId As String = "")
    Dim json As String
    json = BuildFullJSON(restoreCalId)
    ExecuteJS "loadData(" & json & ");"
End Sub
```

---

### **clsCommandProcessor** — Request Router

Dispatches commands from JS to the appropriate repo/manager:

```vba
Public Function ProcessCommand(ByVal jsonStr As String) As String
    Dim action As String
    action = m_JSON.ExtractValue(jsonStr, "action")
    
    Select Case action
        Case "setTheme"
            m_Theme.SetTheme jsonStr
            ProcessCommand = "NORELOAD"
            
        Case "addCalendar"
            Dim newCalId As String
            newCalId = m_Cals.Add(jsonStr)
            ProcessCommand = newCalId
            
        Case "saveAppointment"
            m_Appts.Save jsonStr
            ProcessCommand = m_Cals.ActiveCalendarID
            
        Case "deleteCalendar"
            Dim firstId As String
            firstId = m_Cals.Delete(jsonStr)
            If firstId = "" Then
                ProcessCommand = "NORELOAD"
            Else
                ProcessCommand = firstId
            End If
            
        ' ... more cases
    End Select
End Function
```

---

## 🎨 HTML/CSS/JavaScript Structure

### **calendar.html Overview**

**Main Sections:**

1. **Header** — Month/year nav, view tabs (month/week/day), theme toggle
2. **Sidebar** — Mini-calendar, calendar list (checkboxes), groups
3. **Calendar View** — Month/week/day grid with appointments
4. **Modals** — Create/edit appointment, calendar settings

### **CSS Custom Properties (Theme System)**

Defined in `:root` and `[data-theme="dark"]`:

```css
:root {
  --bg: #f0f2f5;
  --surface: #ffffff;
  --text: #1a1a2e;
  --accent: #4f46e5;
  --shadow: rgba(0,0,0,.08);
}
[data-theme="dark"] {
  --bg: #0f1117;
  --surface: #1e2130;
  --text: #e4e6f0;
  --accent: #7c74f5;
}
```

All colors use CSS custom properties for instant theme switching.

### **JavaScript Global State**

```javascript
let calendars = [];              // Array of {id, name, color, groupId}
let calendarGroups = [];         // Array of {id, name}
let events = [];                 // Array of all appointment occurrences
let openCalendarIds = [];        // IDs of visible calendars
let primaryCalendarId = null;    // ID of "active" calendar for new appointments
let selectedCalId = null;        // Currently selected calendar in tabs
let darkMode = false;            // Theme preference
let pendingCommand = null;       // Queued command for VBA
```

### **Key Rendering Functions**

| Function | Purpose |
| ---------- | --------- |
| `loadData(json)` | Called by VBA; loads calendars and appointments |
| `render()` | Rebuilds calendar grid (month/week/day) |
| `renderSidebar()` | Rebuilds left panel with calendars and groups |
| `renderCalTabs()` | Renders tab bar (when 2+ calendars open) |
| `calEvents()` | Filters events to only open calendars |
| `queueCommand(cmd)` | Stores command in `pendingCommand` for VBA to poll |

### **Multi-Calendar Features**

**Checkboxes in Sidebar:**

```javascript
function toggleCalendar(id) {
  id = String(id);
  const idx = openCalendarIds.indexOf(id);
  if (idx >= 0) {
    if (openCalendarIds.length <= 1) return; // Can't close last one
    openCalendarIds.splice(idx, 1);
  } else {
    if (openCalendarIds.length >= 5) return; // Max 5 overlays
    openCalendarIds.push(id);
  }
  renderSidebar();
  render();
  notifyOpenCalendars();  // Tell VBA to save preference
}
```

**Calendar Tabs (when 2+ open):**
When multiple calendars are open, a colored tab bar appears below the header. Clicking a tab sets that calendar as "primary" (events created default to it).

```javascript
function setPrimaryCalendar(id) {
  id = String(id);
  if (!openCalendarIds.includes(id)) return;
  primaryCalendarId = id;
  renderSidebar();
  renderCalTabs();  // Update highlighted tab
}
```

---

## 📋 Data Flow Examples

### **Example 1: Creating an Appointment**

**User Action:** Click 2026-03-15 in month view, fill form, click Save

1. **JS collects data:**

   ```javascript
   const apptData = {
       action: 'saveAppointment',
       calendarId: '1',
       title: 'Team Meeting',
       date: '2026-03-15',
       start: '14:00',
       end: '15:00',
       color: '#4f46e5',
       allDay: false,
       recurType: 'weekly',
       recurInterval: 1,
       recurDaysOfWeek: '1,3,5',  // Mon, Wed, Fri
       recurEndType: 'date',
       recurEndDate: '2026-06-30'
   };
   queueCommand(apptData);
   ```

2. **VBA polls Form_Timer:**
   - Calls `m_Bridge.GetPendingCommand()`
   - Returns the JSON command

3. **clsCommandProcessor dispatches:**
   - Action = "saveAppointment"
   - Calls `m_Appts.Save(jsonStr)`

4. **clsAppointmentRepo.Save:**
   - Parses JSON (title, date, time, recurrence fields)
   - Checks if AppointmentID exists (update) or is new (insert)
   - Writes to `tblAppointments`
   - Returns CalendarID

5. **Form_Timer reloads:**
   - Calls `m_Bridge.SendDataToCalendar(calendarId)`
   - VBA rebuilds full JSON with expanded recurrences
   - Calls `loadData(json)` in JavaScript
   - JS re-renders the calendar with new appointment

---

### **Example 2: Toggle Dark Mode**

**User Action:** Click theme icon in header

1. **JS calls:**

   ```javascript
   const newDarkMode = !darkMode;
   darkMode = newDarkMode;
   applyTheme(darkMode, true);  // Apply CSS + save preference
   
   queueCommand({
       action: 'setTheme',
       dark: newDarkMode ? 'true' : 'false'
   });
   ```

2. **VBA receives command:**
   - `m_CmdProc.ProcessCommand(cmd)`
   - Action = "setTheme"
   - Calls `m_Theme.SetTheme(jsonStr)`

3. **clsThemeManager:**
   - Parses `"dark"` value
   - Updates `m_DarkMode` flag
   - Returns "NORELOAD" (no DB change)

4. **Form_Timer does NOT reload**
   - Only JS side changes (CSS is instant in browser)

---

### **Example 3: Moving Event Between Calendars**

**User Action:** Drag appointment from Calendar A (open) to Calendar B

1. **JS calculates new state:**

   ```javascript
   const cmd = {
       action: 'moveAppointment',
       id: appointmentId,
       calendarId: targetCalendarId,
       newDate: newDateISO
   };
   queueCommand(cmd);
   ```

2. **VBA processes:**
   - `m_CmdProc.ProcessCommand(cmd)`
   - Calls `m_Appts.Move(jsonStr)`

3. **clsAppointmentRepo.Move:**
   - Finds appointment by ID
   - Updates `CalendarID` and `StartApptDate`
   - Writes to DB

4. **Form_Timer reloads:**
   - Calls `SendDataToCalendar(targetCalendarId)`
   - VBA rebuilds appointments for both old and new calendars
   - JS re-renders calendar

---

## 🔑 Key Design Patterns

### **1. MVC-like Separation**

| Layer | Responsibility |
| ------- | ----------------- |
| **Models** | TblAppointments, TblCalendars (DB tables) |
| **Views** | calendar.html (HTML/CSS rendering) |
| **Controllers** | VBA classes dispatch commands → repos |
| **Bridge** | clsJSBridge handles VBA ↔ JS communication |

### **2. Dependency Injection**

All classes receive dependencies via `Init()`:

```vba
Dim jh As New clsJSONHelper
Dim dh As New clsDateHelper
Dim ar As New clsAppointmentRepo
ar.Init jh, dh  ' Inject utilities
```

**Benefits:**

- Testability (can inject mocks)
- Loose coupling (classes don't create dependencies)
- Reusability (same utility shared across multiple classes)

### **3. Thin Form Shell**

Form code-behind is minimal; all logic lives in classes:

```vba
' Form only owns:
' - Initialization (Form_Load)
' - Timer loop (Form_Timer)
' - Error handling
' - Singleton instance variables
```

### **4. Command Pattern**

JS queues actions as JSON command objects; VBA dispatches via `clsCommandProcessor`:

```vba
{
    "action": "saveAppointment",
    "id": "123",
    "title": "Meeting",
    ...
}
```

**Benefits:**

- Asynchronous (polling, no blocking waits)
- Undo-friendly (can log commands)
- Testable (can inject mock commands)

---

## 🐛 Common Gotchas

### **1. Timezone & Date Handling**

- JS uses ISO dates (`"2026-03-15"`) as strings
- VBA internally uses `Date` type (datetime)
- `clsDateHelper` converts both directions

**Rule:** Always use `ISOToDate()` when reading from JS; use `Format(d, "yyyy-mm-dd")` when writing to JS

### **2. Recurring Appointments**

- Single master record in DB
- VBA expands into **virtual occurrences** (never persisted)
- Each occurrence shares `AppointmentID` but has calculated start date

**Gotcha:** Don't try to save `RecurType="none"` for a single occurrence of a recurring appointment; that's already the master record.

### **3. Multi-Calendar Context**

- JS tracks open calendars (`openCalendarIds`) and primary calendar (`primaryCalendarId`)
- VBA also tracks these in `clsCalendarRepo`
- **Must stay in sync** or UI shows wrong calendar

**Best Practice:** Always reload full state after calendar changes:

```vba
ProcessCommand = m_Cals.ActiveCalendarID  ' Force full reload
```

### **4. Reminder Firing**

- Non-recurring: Persistent flag (`ReminderFired`) survives sessions
- Recurring: In-memory tracking (cleared when form closes)

**Gotcha:** If you set `ReminderFired = True` manually in DB, the reminder won't fire next time the form opens (even for the same occurrence).

---

## 📊 Database Tables (Implied)

Based on code, the database contains:

| Table | Key Fields |
| ------- | ----------- |
| **tblAppointments** | AppointmentID (PK), CalendarID (FK), Title, StartApptDate, StartTime, EndTime, AllDay, Color, Notes, RecurType, RecurInterval, RecurDaysOfWeek, RecurMonthlyMode, RecurMonthDay, RecurMonthWeek, RecurMonthDOW, RecurEndType, RecurEndDate, RecurCount, RecurExceptions, ReminderMinutes, ReminderFired, IsDeleted |
| **tblCalendars** | CalendarID (PK), GroupID (FK), CalendarName, Color, IsDeleted |
| **tblCalendarGroups** | GroupID (PK), GroupName, IsDeleted |

All tables use soft-delete (`IsDeleted = True`) rather than hard deletion.

---

## 🚀 Extending the System

### **Add a New Command**

1. **Add JS handler** in calendar.html (e.g., `saveSettings(data)` → `queueCommand({action: 'saveSettings', ...})`)
2. **Add VBA case** in `clsCommandProcessor.ProcessCommand()`
3. **Add repo method** (e.g., `clsSettingsRepo.Save(jsonStr)`)
4. **Return reload code** ("NORELOAD" or CalendarID)

### **Add a New Field to Appointments**

1. **Add DB column** to `tblAppointments`
2. **Update form modal** in HTML to include input
3. **Update `clsAppointmentRepo.Save()`** to extract and persist the field
4. **Update `clsRecurrenceEngine.BuildOccJSON()`** to include in event JSON
5. **Update JS rendering** to display the field

---

## ✅ Summary

| Component | Role |
| ----------- | ------ |
| **calendar.html** | UI, event listeners, rendering, command queueing |
| **CSS** | Theme system, layouts, responsive design |
| **JavaScript** | Application state, user interactions, validation |
| **clsJSONHelper** | JSON escaping, lightweight parsing |
| **clsDateHelper** | Date arithmetic, ISO ↔ VBA conversions |
| **clsRecurrenceEngine** | Expands masters into virtual occurrences |
| **clsAppointmentRepo** | CRUD for appointments |
| **clsCalendarRepo** | CRUD for calendars + active/open tracking |
| **clsCalendarGroupRepo** | CRUD for calendar groups |
| **clsThemeManager** | Dark/light mode preference |
| **clsReminderManager** | Checks reminders every timer tick |
| **clsJSBridge** | VBA ↔ JS communication (polling, callbacks) |
| **clsCommandProcessor** | Router: dispatches actions to repos |
| **Form (VBA)** | Initialization, timer loop, singleton instances |

**The key insight:** JS handles UI and user input; VBA handles data integrity and business logic. They communicate via JSON command queue + full state sync.
